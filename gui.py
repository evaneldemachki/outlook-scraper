import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets
import core
import json
import pandas

def bridge():
    #sess.update()
    with open("config.json", "r") as f:
        config = json.load(f)

    index = {}
    for account in config["index"]:
        entry = {
            "object": None,
            "inbox": {}
        }
        inbox = entry["inbox"]
        data = sess.loadtable(account)
        for row in data.itertuples():
            inbox[row[1]] = {
                "attributes": [*row[2:-1]],
                "content": row[-1]
            }

        index[account] = entry

    return index

class FlipTable(QtWidgets.QTableWidget):
    def __init__(self, parent, columns):
        super().__init__(parent)

        _translate = QtCore.QCoreApplication.translate

        self.columns = columns
        self.pages = {}
        self.active = None

        self.setColumnCount(len(columns))
        self.setRowCount(0)

        item = QtWidgets.QTableWidgetItem()
        item.setText(_translate("MainWindow", "0"))
        self.setVerticalHeaderItem(0, item)
        for col, n in zip(self.columns, range(len(self.columns))):
            item = QtWidgets.QTableWidgetItem()
            item.setText(_translate("MainWindow", self.columns[col]))
            self.setHorizontalHeaderItem(n, item)

    def addpage(self, page):
        self.pages[page] = []
        if len(self.pages) == 1:
            self.flipto(page)

    def flipto(self, page):
        if self.active is not None and len(self.pages[page]) != 0:
            active_page = self.pages[self.active]
            for r in range(len(active_page)):
                for c in range(len(self.columns)):
                    active_page[r][c] = self.takeItem(r, c)

        if len(self.pages[page]) != 0:
            self.setRowCount(len(self.pages[page]))
            for r in range(len(self.pages[page])):
                for c in range(len(self.columns)):
                    item = self.pages[page][r][c]
                    self.setItem(r, c, item)

        self.active = page

    def refresh(self, extended):
        active_page = self.pages[self.active]
        for r in range(len(active_page) - extended):
            for c in range(len(self.columns)):
                active_page[r][c] = self.takeItem(r, c)

        self.setRowCount(len(active_page))
        for r in range(len(active_page)):
            for c in range(len(self.columns)):
                item = self.pages[self.active][r][c]
                self.setItem(r, c, item)

    def addrows(self, page, rows):
        _translate = QtCore.QCoreApplication.translate
        extended = 0
        for row in rows:
            item_row = []
            for cell in row:
                item = QtWidgets.QTableWidgetItem()
                item.setText(_translate("MainWindow", cell))
                item_row.append(item)

            self.pages[page].append(item_row)
            extended += 1

        if self.active == page:
            self.refresh(extended)

class Ui_MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self.cache = {
            "index": bridge(),
            "active": None
        }

        self.setupUi()
        self.show()

    def setupUi(self):
        _translate = QtCore.QCoreApplication.translate
        # set main window parameters
        self.setObjectName("MainWindow")
        self.resize(1000, 650)
        # create central window widget
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        # create inbox selection sidebar
        self.sidebar = QtWidgets.QFrame(self.centralwidget)
        self.sidebar.setGeometry(QtCore.QRect(10, 50, 161, 521))
        self.sidebar.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.sidebar.setFrameShadow(QtWidgets.QFrame.Raised)
        self.sidebar.setObjectName("sidebar")
        # dynamically generate inbox selector buttons
        top_pos = 10
        for account in self.cache["index"]:
            pushButton = QtWidgets.QPushButton(self.sidebar)
            pushButton.setGeometry(QtCore.QRect(10, top_pos, 141, 31))
            pushButton.setObjectName(account)
            pushButton.clicked.connect(self.flip)
            self.cache["index"][account]["object"] = pushButton
            top_pos += 40
        # create selector sidebar label
        self.sidebar_label = QtWidgets.QLabel(self.centralwidget)
        self.sidebar_label.setGeometry(QtCore.QRect(10, 0, 171, 41))
        self.sidebar_label.setAlignment(QtCore.Qt.AlignCenter)
        self.sidebar_label.setObjectName("sidebar_label")

        # create inbox list widget and label
        table_columns = {
            "sent_at": "Timestamp",
            "email": "Sender Email",
            "name": "Sender Name",
            "subject": "Subject"
        }
        self.tableWidget = FlipTable(self.centralwidget, table_columns)
        self.tableWidget.setGeometry(QtCore.QRect(180, 50, 621, 211))
        self.tableWidget.setObjectName("email_attributes")
        for account in self.cache["index"]:
            self.tableWidget.addpage(account)
            rows = []
            for key, data in self.cache["index"][account]["inbox"].items():
                rows.append(data["attributes"])

            self.tableWidget.addrows(account, rows)
        # bind showcontent() to table indexes
        self.tableWidget.clicked.connect(self.showcontent)
        # set active element to FlipTable.active
        self.cache["active"] = self.tableWidget.active
        # create table label
        self.item_attr_label = QtWidgets.QLabel(self.centralwidget)
        self.item_attr_label.setGeometry(QtCore.QRect(180, 10, 621, 31))
        self.item_attr_label.setAlignment(QtCore.Qt.AlignCenter)
        self.item_attr_label.setObjectName("email_attr_label")
        # create email content viewer
        self.item_content = QtWidgets.QTextBrowser(self.centralwidget)
        self.item_content.setGeometry(QtCore.QRect(180, 310, 631, 261))
        self.item_content.setObjectName("email_content")
        # create email content label
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(180, 270, 621, 31))
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("content_label")
        # create control panel
        self.control = QtWidgets.QFrame(self.centralwidget)
        self.control.setGeometry(QtCore.QRect(820, 50, 161, 521))
        self.control.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.control.setFrameShadow(QtWidgets.QFrame.Raised)
        self.control.setObjectName("control_panel")
        # create control panel label
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(810, 10, 171, 41))
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("control_label")
        # create control panel buttons
        self.updateButton = QtWidgets.QPushButton(self.control)
        self.updateButton.setGeometry(QtCore.QRect(10, 10, 141, 31))
        self.updateButton.setObjectName("pushButton_2")
        self.updateButton.clicked.connect(self.update)
        self.masterButton = QtWidgets.QPushButton(self.control)
        self.masterButton.setGeometry(QtCore.QRect(10, 50, 141, 31))
        self.masterButton.setObjectName("pushButton_3")
        self.masterButton.clicked.connect(self.makemaster)

        self.setCentralWidget(self.centralwidget)
        QtCore.QMetaObject.connectSlotsByName(self)
        self.retranslateUi()

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "Outlook Parser"))
        __sortingEnabled = self.tableWidget.isSortingEnabled()
        self.tableWidget.setSortingEnabled(False)

        for key in self.cache["index"]:
            self.cache["index"][key]["object"].setText(_translate("MainWindow", key))

        self.updateButton.setText(_translate("MainWindow", "Update"))
        self.masterButton.setText(_translate("MainWindow", "Make Master"))

        #for account in self.cache["index"]:
            #for email in self.cache["index"][account]["inbox"]:
                #item = self.cache["index"][account]["inbox"][email]["object"]
                #text = self.cache["index"][account]["inbox"][email]["text"]
                #item.setText(_translate("MainWindow", text))

        self.sidebar_label.setText(_translate("MainWindow", "Inbox Selector"))

    def flip(self):
        sender_name = self.sender().objectName()
        self.tableWidget.flipto(sender_name)
        self.cache["active"] = sender_name
        self.clearcontent()

    def showcontent(self, table_click):
        _translate = QtCore.QCoreApplication.translate

        row = table_click.row()
        account = self.cache["active"]
        key = list(self.cache["index"][account]["inbox"])[row]
        content = self.cache["index"][account]["inbox"][key]["content"]
        self.item_content.setHtml(_translate("MainWindow", content))

    def clearcontent(self):
        _translate = QtCore.QCoreApplication.translate

        self.item_content.setHtml(_translate("MainWindow", ""))

    def update(self):
        extension = sess.update(gui=True)
        for account in extension:
            if extension[account] is not None:
                entries = {}
                for row in extension[account]:
                    entries[row[0]] = {
                        "attributes": [*row[1:-1]],
                        "content": row[-1]
                    }

                    self.cache["index"][account]["inbox"][row[0]] = entries[row[0]]

                rows = [[*entries[entry]["attributes"]] for entry in entries]
                self.tableWidget.addrows(account, rows)

    def makemaster(self):
        sess.makemaster()

sess = core.Session()

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ui = Ui_MainWindow()
    #_logcache()
    sys.exit(app.exec_())

ui = None
app = None
def run():
    import sys
    global ui
    global app
    app = QtWidgets.QApplication(sys.argv)
    #MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    #_logcache()
    sys.exit(app.exec_())


def _logcache():

    def strfobj(entry):
        if entry is not None:
            try:
                return entry.text()
            except:
                return "ERROR"
        else:
            return None

    def parsecache(obj):
        temp_cache = {}

        focus = obj
        clone = temp_cache
        for key in focus:
            item = focus[key]
            mirror = clone[key] = {}
            mirror["object"] = strfobj(item["object"])
            for key_2 in list(item)[1:]:
                if type(item[key_2]) is dict:
                    mirror[key_2] = parsecache(item[key_2])
                else:
                    mirror[key_2] = item[key_2]

        return clone

    obj = ui.cache["index"]
    temp_cache = {"index": parsecache(obj)}

    with open("log.json", "w") as f:
        json.dump(temp_cache, f)
