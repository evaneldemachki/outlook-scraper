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
            
    def update(self):
    
        def buildrow(obj):
            row = []
            row.append(str(obj.EntryID))
            row.append(str(obj.SentOn))
            row.append(str(obj.SenderEmailAddress))
            row.append(str(obj.SenderName))
            row.append(str(obj.Subject))
            row.append(str(obj.Body))
            return tuple(row)
            
        # load table index
        with open("index.json", "r") as f:
            table_index = json.load(f)
        data = {}
        # for each account:
        for account in self.accounts:
            table_id = table_index[str(account)]
            data[table_id] = []
            # locate inbox folder
            inbox = outlook.Folders(account.DeliveryStore.DisplayName)
            inbox = inbox.Folders("Inbox")
            # parse each email object, add to list
            last_ts = self._lastts(table_id)
            i = 0
            for obj in reversed(list(inbox.Items)):
                ts = datetime.datetime.strptime(str(obj.SentOn).split("+")[0], "%Y-%m-%d %H:%M:%S")
                if last_ts is None:
                    data[table_id].append(buildrow(obj))
                    i += 1
                    continue
                elif ts > last_ts:
                    data[table_id].append(buildrow(obj))
                    i += 1
                    continue
        
        self._write(data)
                
        
    def makemaster(self):
        # load table index
        with open("index.json", "r") as f:
            table_index = json.load(f)
        with open(self.master, "w") as f:
            writer = csv.writer(f, lineterminator='\n')
            data = []
            for table_name in table_index:
                query = "SELECT * FROM '{0}'".format(table_index[table_name])
                self.c.execute(query)
                for crow in self.c.fetchall():
                    row = [table_name] + [str(x) for x in list(crow)][:-1]
                    data.append(row)
            i = 0
            for row in data:
                writer.writerow(row)
                i += 1
        
        print("Update successful: parsed ({0}) new email(s)".format(i))
        
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
        print("Successfully wrote ({0}) row(s) to database".format(total))
        
    
    
                
        