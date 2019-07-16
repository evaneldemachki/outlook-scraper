import win32com.client as client
import csv
import sqlite3
import uuid
import os
import json
import datetime
import exceptions

outlook = client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# generate config file
def generate():
    if "config.json" in os.listdir():
        raise exceptions.ExistingConfig

    config = {
        "index": {},
        "settings": {
            "limit": "10d"
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
            config = json.load(f)
        strpdelta(config["settings"]["limit"])
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

        with open("config.json", "r") as f:
            config = json.load(f)
        for account in config:
            query = "SELECT count(*) FROM '{0}'".format(config["index"][account])
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

    def update(self, mode=None):
        if mode == "debug":
            with open("debug.txt", "a", encoding="utf-8") as f:
                pass

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
                if mode == "debug":
                    with open("debug.txt", "a", encoding="utf-8") as f:
                        f.write("..." + attr + ":\n")
                try:
                    entry = str(getattr(obj, attr))
                    if attr == "SentOn":
                        entry = entry.split("+")[0]
                except:
                    entry = "ERROR"
                if mode == "debug":
                    with open("debug.txt", "a", encoding="utf-8") as f:
                        f.write("......" + entry + ":\n")

                row.append(entry)

            return tuple(row)

        # load config file
        with open("config.json", "r") as f:
            config = json.load(f)
        # initialize update queue
        queue = {}
        # for each account:
        count = 0
        for account in self.accounts:
            acc_name = str(account)
            print("Scanning account: {0}".format(acc_name))
            # log to debug if in debug mode
            if mode == "debug":
                with open("debug.txt", "a", encoding="utf-8") as f:
                    f.write("account: " + acc_name + ":\n")
            # load table id from table index
            table_id = config["index"][acc_name]
            queue[table_id] = []
            # locate inbox folder
            inbox = outlook.Folders(account.DeliveryStore.DisplayName)
            inbox = inbox.Folders("Inbox").Items
            # find last saved timestamp from database
            last_ts = self._lastts(table_id)
            limit = strpdelta(config["settings"]["limit"])
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
            else:
                print("...found ({0}) new email(s)".format(i))
                self._write(queue)

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
        # load table index
        with open("config.json", "r") as f:
            config = json.load(f)
        with open(path, "w", encoding="utf-8") as f:
            writer = csv.writer(f, lineterminator='\n')
            i = 0
            for table_name in config["index"]:
                query = "SELECT * FROM '{0}'".format(config["index"][table_name])
                self.c.execute(query)
                for crow in self.c.fetchall():
                    row = [table_name] + [str(x) for x in list(crow)]
                    writer.writerow(row)
                    i += 1

        print("Successfully created master file: parsed ({0}) email(s)".format(i))

    def _updatedb(self):
        def genquery(table_id):
            query = """CREATE TABLE '{0}'
                (entry_id text,
                sent_at text,
                email text,
                name text,
                subject text,
                content text)""".format(table_id)

            return query
        # load configuration file
        with open("config.json", "r") as f:
            config = json.load(f)
        # for each email account
        i = 0
        for account in self.accounts:
            # if account does not exist in table index:
            if str(account) not in config["index"]:
                table_id = str(uuid.uuid4()).replace("-", "_")
                # create sqlite3 table
                self.c.execute(genquery(table_id))
                i += 1
                # update table index
                config["index"][str(account)] = table_id
                with open("config.json", "w") as f:
                    json.dump(config, f)
            else: # assure table exists in database
                table_id = config["index"][str(account)]
                query = "SELECT name FROM sqlite_master WHERE type='table' AND name='{0}'"
                query = query.format(table_id)
                self.c.execute(query)
                if type(self.c.fetchone()) is tuple:
                    continue
                else: # create table if not exists
                    self.c.execute(genquery(table_id))

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

        ts_list = [datetime.datetime.strptime(
            ts.split("+")[0], "%Y-%m-%d %H:%M:%S") for ts in ts_list]

        return max(ts_list)

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
