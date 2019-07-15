import win32com.client as client
import csv
import sqlite3
import uuid
import os
import json
import datetime

if "index.json" not in os.listdir():
    with open("index.json", "w") as f:
        json.dump({}, f)
        
outlook = client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
class Session:
    def __init__(self):
        self.accounts = []
        for account in client.Dispatch("Outlook.Application").Session.Accounts:
            self.accounts.append(account)
        
        self.master = "master.csv"
        self.conn = sqlite3.connect("core.db")
        self.c = self.conn.cursor()
        
        self._updatedb()
    
    def __repr__(self):
        msg = "outlook-scraper session:\n"
        msg += "-accounts:\n"
        
        with open("index.json", "r") as f:
            table_index = json.load(f)
        for account in table_index:
            query = "SELECT count(*) FROM '{0}'".format(table_index[account])
            self.c.execute(query)
            table_count = self.c.fetchone()[0]
            msg += "\t{0}: {1} email(s) stored\n".format(account, table_count)
        
        return msg[:-1]
        
        
            
    def update(self, mode=None):
    
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
                
                with open("debug.txt", "a", encoding="utf-8") as f:
                    f.write("......" + entry + ":\n")
                    
                row.append(entry)
           
            return tuple(row)
            
        # load table index
        with open("index.json", "r") as f:
            table_index = json.load(f)
        data = {}
        # for each account:
        for account in self.accounts:
            with open("debug.txt", "a", encoding="utf-8") as f:
                f.write("account: " + str(account) + ":\n")
            table_id = table_index[str(account)]
            data[table_id] = []
            # locate inbox folder
            inbox = outlook.Folders(account.DeliveryStore.DisplayName)
            inbox = inbox.Folders("Inbox")
            # parse each email object, add to list
            last_ts = self._lastts(table_id)
            i = 0
            for obj in reversed(list(inbox.Items)):
                try:
                    sent_on = str(obj.SentOn)
                except:
                    continue
                    
                ts = datetime.datetime.strptime(sent_on.split("+")[0], "%Y-%m-%d %H:%M:%S")
                if last_ts is None:
                    data[table_id].append(buildrow(obj))
                    i += 1
                    continue
                elif ts > last_ts:
                    data[table_id].append(buildrow(obj))
                    i += 1
                    continue
        
        self._write(data)
                
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
                        
                        
    def makemaster(self):
        # load table index
        with open("index.json", "r") as f:
            table_index = json.load(f)
        with open(self.master, "w", encoding="utf-8") as f:
            writer = csv.writer(f, lineterminator='\n')
            data = []
            for table_name in table_index:
                query = "SELECT * FROM '{0}'".format(table_index[table_name])
                self.c.execute(query)
                for crow in self.c.fetchall():
                    row = [table_name] + [str(x) for x in list(crow)]
                    data.append(row)
            i = 0
            for row in data:
                writer.writerow(row)
                i += 1
        
        print("Successfully created master file: parsed ({0}) new email(s)".format(i))
        
    def _updatedb(self):
        # load table index
        with open("index.json", "r") as f:
            table_index = json.load(f)
        # for each email account
        for account in self.accounts:
            # if account does not exist in table index:
            i = 0
            if str(account) not in table_index:
                table_id = str(uuid.uuid4()).replace("-","_")
                # create sqlite3 table
                query = """CREATE TABLE '{0}'
                    (entry_id text,
                    sent_at text,
                    email text,
                    name text,
                    subject text,
                    content text)""".format(table_id)
                self.c.execute(query)
                i += 1
                # update table index
                table_index[str(account)] = table_id
                with open("index.json", "w") as f:
                    json.dump(table_index, f)
                    
        
        self.conn.commit()
        if i == 0:
            print("Table structure up to date")
        else:
            print("Added ({0}) new table(s)".format(i))
    
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
                
        ts_list = [datetime.datetime.strptime(ts.split("+")[0], "%Y-%m-%d %H:%M:%S") for ts in ts_list]
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
            
        if c == 0:
            print("All data is up to date")
        else:
            print("Successfully wrote ({0}) row(s) to database".format(total))
        
    
        
    
    
                
        