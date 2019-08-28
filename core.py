import win32com.client as client
import csv
import sqlite3
import uuid
import os
import json
import datetime
import exceptions
import pandas

outlook = client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# generate config file
def generate():
    if "config.json" in os.listdir():
        raise exceptions.ExistingConfig

    config = {
        "index": {},
        "settings": {
            "limit": "10d"
        },
        "categories": [],
        "gui": {
            "category_colors": {}
        }
    }
    for account in client.Dispatch("Outlook.Application").Session.Accounts:
        config["index"][str(account)] = str(uuid.uuid4()).replace("-", "_")

    with open("config.json", "w") as f:
        json.dump(config, f)

    print("Successfully generated configuration file")

def strpdelta(ds):
    mag = int(ds[:-1])
    unit = ds[-1]
    if unit == "d":
        delta = datetime.timedelta(days=mag)
        return delta
    elif unit == "h":
        delta = datetime.timedelta(hours=mag)
        return delta
    else:
        raise exceptions.InvalidLimit(str(ds))

class Session:
    def __init__(self):
        # verify config file exists
        if "config.json" not in os.listdir():
            raise exceptions.NoConfigFile
        # verify limit is valid
        with open("config.json", "r") as f:
            self.config = json.load(f)
        strpdelta(self.config["settings"]["limit"])
        # load outlook accounts
        self.accounts = []
        for account in client.Dispatch("Outlook.Application").Session.Accounts:
            self.accounts.append(account)
        # establish sqlite3 connection
        self.conn = sqlite3.connect("core.db")
        self.c = self.conn.cursor()
        # update table structure
        self._updatedb()

    def __repr__(self):
        msg = "outlook-scraper session:\n"
        msg += "-accounts:\n"

        for account in self.config:
            query = "SELECT count(*) FROM '{0}'".format(self.config["index"][account])
            self.c.execute(query)
            table_count = self.c.fetchone()[0]
            msg += "\t{0}: {1} email(s) stored\n".format(account, table_count)

        return msg[:-1]

    def _inboxlist(self, inbox, floor):
        # restrict inbox items to floor timestamp
        dtf = "%Y-%m-%d %I:%M %p"
        floor_str = datetime.datetime.strftime(floor, dtf)
        filter = "[SentOn] > '{0}'".format(floor_str)
        inbox = inbox.Restrict(filter)
        # generate inbox items in reverse
        for i in range(0, inbox.Count):
            try:
                sent_on = str(inbox[i].SentOn).split("+")[0]
                ts = datetime.datetime.strptime(sent_on, "%Y-%m-%d %H:%M:%S")
                if ts > floor:
                    yield inbox[i]
                else:
                    continue
            except:
                continue

    def update(self, gui=False):
        if gui:
            extension = {}

        def buildrow(obj):
            row = []

            attr_list = (
                "EntryID",
                "SentOn",
                "SenderEmailAddress",
                "SenderName",
                "Subject",
                "Body"
            )
            for attr in attr_list:
                try:
                    entry = str(getattr(obj, attr))
                    if attr == "SentOn":
                        entry = entry.split("+")[0]
                except:
                    entry = "ERROR"

                row.append(entry)

            return tuple(row)

        # initialize update queue
        queue = {}
        # for each account:
        count = 0
        for account in self.accounts:
            acc_name = str(account)
            print("Scanning account: {0}".format(acc_name))
            # load table id from table index
            table_id = self.config["index"][acc_name]
            queue[table_id] = []
            # locate inbox folder
            inbox = outlook.Folders(account.DeliveryStore.DisplayName)
            inbox = inbox.Folders("Inbox").Items
            # find last saved timestamp from database
            last_ts = self._lastts(table_id)
            limit = strpdelta(self.config["settings"]["limit"])
            curr_time = datetime.datetime.now()
            limit = curr_time - limit
            if last_ts is None:
                floor = limit
            else:
                floor = max(last_ts, limit)
            # parse each email in inbox, add to list
            i = 0
            for obj in self._inboxlist(inbox, floor):
                queue[table_id].append(buildrow(obj))
                i += 1

            if i == 0:
                print("...account is up to date")
                if gui:
                    extension[acc_name] = None
            else:
                print("...found ({0}) new email(s)".format(i))
                if gui:
                    extension[acc_name] = pandas.DataFrame(
                        queue[table_id],
                        columns=[
                            "entry_id",
                            "sent_at",
                            "email",
                            "name",
                            "subject",
                            "content"
                        ]
                    )

        self._write(queue)
        if gui:
            return extension

    def loadtable(self, account, mirror=None):
        table_id = self.config["index"][account]
        query = "SELECT * FROM '{0}'"
        if mirror is None:
            query = query.format(table_id)
            data = pandas.read_sql(query, con=self.conn)
        elif mirror is "archive":
            mirror_id = "archive_" + table_id
            query = "SELECT * FROM '{0}'".format(mirror_id)
            data = pandas.read_sql(query, con=self.conn)
        elif mirror is "submissions":
            mirror_id = "submissions_" + table_id
            query = "SELECT * FROM '{0}'".format(mirror_id)
            data = pandas.read_sql(query, con=self.conn)
        else:
            raise exceptions.InvalidMirrorTable

        return data

    def load(self, account):
        table_id = self.config["index"][account]
        query = "SELECT * FROM '{0}'".format(table_id)
        data = pandas.read_sql(query, con=self.conn)

        return data

    def load_archive(self, account, category=None):
        table_id = self.config["index"][account]
        archive_id = "archive_" + table_id
        if category is None:
            query = "SELECT * FROM '{0}'".format(archive_id)
        else:
            query = "SELECT * FROM '{0}' WHERE category='{1}'".format(archive_id, category)

        data = pandas.read_sql(query, con=self.conn)

        return data


    def debug(self):
        for account in self.accounts:
            acc_name = str(account)
            inbox = outlook.Folders(account.DeliveryStore.DisplayName)
            inbox = inbox.Folders("Inbox")
            attr_list = (
                "EntryID",
                "SentOn",
                "SenderEmailAddress",
                "SenderName",
                "Subject",
                "Body"
            )
            for item in inbox.Items:
                for attr in attr_list:
                    try:
                        getattr(item, attr)
                    except:
                        print("Failed on object {0}".format(str(item)))
                        print("account: {0}\n".format(acc_name))
                        print("...attribute: {0}\n".format(attr))
                        return item

    def makemaster(self, path="master.csv"):
        with open(path, "w", encoding="utf-8") as f:
            writer = csv.writer(f, lineterminator='\n')
            i = 0
            for table_name in self.config["index"]:
                query = "SELECT * FROM '{0}'".format(self.config["index"][table_name])
                self.c.execute(query)
                for crow in self.c.fetchall():
                    row = [table_name] + [str(x) for x in list(crow)]
                    writer.writerow(row)
                    i += 1

        print("Successfully created master file: parsed ({0}) email(s)".format(i))

    def _updatedb(self):
        def genquery(table_id, archive=False):
            query = """CREATE TABLE '{0}'
                (entry_id text unique,
                sent_at text,
                email text,
                name text,
                subject text,
                content text)""".format(table_id)

            if archive:
                query = query[:-1] + ",\n\t" + "category text)"

            return query
        # for each email account
        i = 0
        for account in self.accounts:
            # if account does not exist in table index:
            if str(account) not in self.config["index"]:
                table_id = str(uuid.uuid4()).replace("-", "_")
                # create sqlite3 table
                self.c.execute(genquery(table_id))
                self.c.execute(genquery("archive_" + table_id, archive=True))
                i += 1
                # update table index
                self.config["index"][str(account)] = table_id
                with open("config.json", "w") as f:
                    json.dump(self.config, f)
            else: # assure table exists in database
                table_id = self.config["index"][str(account)]
                query = "SELECT name FROM sqlite_master WHERE type='table' AND name='{0}'"
                query = query.format(table_id)
                self.c.execute(query)
                if type(self.c.fetchone()) is tuple:
                    continue
                else: # create table if not exists
                    self.c.execute(genquery(table_id))
                # assure archive table exists
                archive_id = "archive_{0}".format(table_id)
                query = "SELECT name FROM sqlite_master WHERE type='table' AND name='{0}'".format(archive_id)
                self.c.execute(query)
                # create archives table if not exist
                if type(self.c.fetchone()) is not tuple:
                    self.c.execute(genquery(archive_id, archive=True))

        self.conn.commit()

        if i == 0:
            print("Table structure up to date")
        else:
            print("Added ({0}) new index entry(s)".format(i))

    def _lastts(self, table_id):
        # check if table is empty
        query = "SELECT count(*) FROM '{0}'".format(table_id)
        self.c.execute(query)
        if int(self.c.fetchone()[0]) == 0:
            return None
        # if not empty, return last entry ID
        query = "SELECT sent_at FROM '{0}'".format(table_id)
        self.c.execute(query)
        ts_list = [str(x[0]) for x in self.c.fetchall()]
        x, y = 0, len(ts_list)
        while x < y:
            if ts_list[x] == "ERROR":
                ts_list.pop(x)
                y -= 1
                continue

            x += 1

        convert_ts = lambda ts: datetime.datetime.strptime(
            ts.split("+")[0],
            "%Y-%m-%d %H:%M:%S"
        )
        ts_list = [convert_ts(ts) for ts in ts_list]

        return max(ts_list)

    def archive_one(self, account, entry_id, category=None):
        table_id = self.config["index"][account]
        query = "SELECT * FROM '{0}' WHERE entry_id='{1}'".format(table_id, entry_id)
        row = self.c.execute(query)
        row = self.c.fetchone() + (category,)
        if type(row) is tuple:
            query = "INSERT INTO 'archive_{0}' VALUES (?, ?, ?, ?, ?, ?, ?)".format(table_id)
            try:
                self.c.execute(query, row)
                self.conn.commit()
            except sqlite3.IntegrityError:
                raise exceptions.DuplicateEntryID
        else:
            raise exceptions.InvalidEntryID

    def archive_many(self, account, entry_ids, category=None, ignore_duplicates=False):
        for entry_id in entry_ids:
            try:
                self.archive_one(account, entry_id, category)
            except exceptions.DuplicateEntryID:
                if ignore_duplicates:
                    continue
                else:
                    raise
            except:
                raise

    def delete_one(self, account, entry_id, archive=False):
        table_id = self.config["index"][account]
        if archive is False:
            query_check = "SELECT count(*) FROM '{0}' WHERE entry_id='{1}'".format(table_id, entry_id)
            query_del = "DELETE FROM '{0}' WHERE entry_id='{1}'".format(table_id, entry_id)
        elif archive is True:
            query_check = "SELECT count(*) FROM 'archive_{0}' WHERE entry_id='{1}'".format(table_id, entry_id)
            query_del = "DELETE FROM 'archive_{0}' WHERE entry_id='{1}'".format(table_id, entry_id)
        else:
            raise ValueError("archive takes (True/False)")

        row_count = self.c.execute(query_check)
        row_count = self.c.fetchone()[0]
        if row_count == 1:
            self.c.execute(query_del)
            self.conn.commit()
        else:
            raise exceptions.InvalidEntryID

    def delete_many(self, account, entry_ids, archive=False):
        for entry_id in entry_ids:
            self.delete_one(account, entry_id, archive)

    def _write(self, data):
        # for each email account inbox:
        row_counts = []
        for table_id in data:
            # insert parsed email data
            query = "INSERT INTO '{0}' VALUES (?, ?, ?, ?, ?, ?)".format(table_id)
            self.c.executemany(query, data[table_id])
            row_counts.append(len(data[table_id]))

        self.conn.commit()
        total = 0
        for c in row_counts:
            total += c

        print("...successfully saved to database")
