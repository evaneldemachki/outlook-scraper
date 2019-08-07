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
            "inbox": {
                "main": {},
                "archive": {},
                "submissions": {}
            }
        }

        main_data = sess.loadtable(account)
        arch_data = sess.loadtable(account, "archive")
        subm_data = sess.loadtable(account, "submissions")

        main_content = main_data.pop("content")
        arch_content = arch_data.pop("content")
        subm_content = subm_data.pop("content")

        main_ids = main_data.pop("entry_id")

        arch_ids = arch_data.pop("entry_id")
        arch_index = []
        for id in arch_ids:
            if id in main_ids:
                arch_index.append(True)
            else:
                arch_index.append(False)
        arch_index = pandas.Series(arch_index)

        subm_ids = subm_data.pop("entry_id")
        subm_index = []
        for id in subm_ids:
            if id in main_ids:
                subm_index.append(True)
            else:
                subm_index.append(False)
        subm_index = pandas.Series(subm_index)
        # rename columns to numbered
        data.columns = [0,1,2,3]
        entry["inbox"]["main"] = {
            "object": None,
            "attributes": main_data,
            "content": main_content,
            "ids": main_ids,
            "archive": arch_index,
            "submissions": subm_index
        }
        entry["inbox"]["archive"] = {
            "object": None,
            "attributes": arch_data,
            "content": main_content,
            "ids": arch_ids
        }
        entry["inbox"]["submissions"] = {
            "object": None,
            "attributes": subm_data,
            "content": main_content,
            "ids": subm_ids
        }

        index[account] = entry

    return index

class InboxModel(QtCore.QAbstractTableModel):
    def __init__(self, data, entry_ids, archive, submissions, parent=None):
        super().__init__(parent)

        self.headings = ["Timestamp", "Sender Address", "Sender Name", "Subject"]
        self.dataframe = data
        self.entry_ids = entry_ids
        self.archive = archive
        self.submissions = submissions

    def rowCount(self, parent):
        return self.dataframe.shape[0]

    def columnCount(self, parent):
        return self.dataframe.shape[1]

    def data(self, index, role):
        if role == QtCore.Qt.DisplayRole:
            return self.dataframe[index.column()][index.row()]

        return QtCore.QVariant()

    def headerData(self, section, orientation, role):
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        return self.headings[section]

    def addRows(self, rows, entry_ids, archive, submissions):
        position = self.dataframe.shape[0] - 1
        end = position + rows.shape[0] - 1
        # begin model insertion
        self.beginInsertRows(QtCore.QModelIndex(), position, end)
        # append rows to dataframe
        self.dataframe = self.dataframe.append(rows, ignore_index=True)
        # append entry_ids and mirrors
        self.entry_ids = self.entry_ids.append(entry_ids, ignore_index=True)
        self.archive = self.archive.append(archive, ignore_index=True)
        self.submissions = self.submissions.append(submissions, ignore_index=True)
        # end model insertion
        self.endInsertRows()
        # shade rows
        for i in range(0, len(self.entry_ids)):
            if self.archive[i] is True and self.submissions[i] is False:
                for j in range(0, 4):
                    self.setData(model.index(i, j), QtCore.QBrush(QtCore.Qt.red), QtCore.Qt.BackgroundRole)
            if self.archive[i] is True and self.submissions[i] is True:
                for j in range(0, 4):
                    self.setData(model.index(i, j), QtCore.QBrush(QtCore.Qt.blue), QtCore.Qt.BackgroundRole)
            if self.archive[i] is False and self.submissions[i] is True:
                for j in range(0, 4):
                    self.setData(model.index(i, j), QtCore.QBrush(QtCore.Qt.green), QtCore.Qt.BackgroundRole)

    def addToMirror(self, rows, mirror):
        if mirror == "archive":
            for row in rows:
                self.archive[row] = True
        if mirror == "submissions":
            for row in rows:
                self.submissions[row] = True
        # shade rows
        for i in range(0, len(self.entry_ids)):
            if self.archive[i] is True and self.submissions[i] is False:
                for j in range(0, 4):
                    self.setData(model.index(i, j), QtCore.QBrush(QtCore.Qt.red), QtCore.Qt.BackgroundRole)
            if self.archive[i] is True and self.submissions[i] is True:
                for j in range(0, 4):
                    self.setData(model.index(i, j), QtCore.QBrush(QtCore.Qt.blue), QtCore.Qt.BackgroundRole)
            if self.archive[i] is False and self.submissions[i] is True:
                for j in range(0, 4):
                    self.setData(model.index(i, j), QtCore.QBrush(QtCore.Qt.green), QtCore.Qt.BackgroundRole)



class MirrorModel(QtCore.QAbstractTableModel):
    def __init__(self, data, entry_ids, parent=None):
        super().__init__(parent)

        self.headings = ["Timestamp", "Sender Address", "Sender Name", "Subject"]
        self.dataframe = data
        self.entry_ids = entry_ids

    def rowCount(self, parent):
        return self.dataframe.shape[0]

    def columnCount(self, parent):
        return self.dataframe.shape[1]

    def data(self, index, role):
        if role == QtCore.Qt.DisplayRole:
            return self.dataframe[index.column()][index.row()]

        return QtCore.QVariant()

    def headerData(self, section, orientation, role):
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        return self.headings[section]

    def addRows(self, rows, entry_ids):
        position = self.dataframe.shape[0] - 1
        end = position + rows.shape[0] - 1
        # begin model insertion
        self.beginInsertRows(QtCore.QModelIndex(), position, end)
        # append rows to dataframe
        self.dataframe = self.dataframe.append(rows, ignore_index=True)
        # add entry ids
        self.entry_ids = self.entry_ids.append(entry_ids, ignore_index=True)
        # end model insertion
        self.endInsertRows()


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
        self.resize(1440, 750)
        # create central window widget
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        # create inbox selection sidebar
        self.sidebar = QtWidgets.QFrame(self.centralwidget)
        self.sidebar.setGeometry(QtCore.QRect(5, 50, 180, 300))
        self.sidebar.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.sidebar.setFrameShadow(QtWidgets.QFrame.Raised)
        self.sidebar.setObjectName("sidebar")
        # dynamically generate inbox selector buttons
        top_pos = 0
        for account in self.cache["index"]:
            pushButton = QtWidgets.QPushButton(self.sidebar)
            pushButton.setGeometry(QtCore.QRect(5, top_pos, 170, 30))
            pushButton.setObjectName(account)
            pushButton.clicked.connect(self.flip)
            self.cache["index"][account]["object"] = pushButton
            top_pos += 40
        # create selector sidebar label
        self.sidebar_label = QtWidgets.QLabel(self.centralwidget)
        self.sidebar_label.setGeometry(QtCore.QRect(10, 0, 180, 50))
        self.sidebar_label.setAlignment(QtCore.Qt.AlignCenter)
        self.sidebar_label.setObjectName("sidebar_label")
        # set live table view
        self.liveView = QtWidgets.QTableView(self.centralwidget)
        self.liveView.setGeometry(QtCore.QRect(200, 50, 600, 300))
        self.liveView.setObjectName("email_attributes")
        self.liveView.setSelectionMode(QtWidgets.QAbstractItemView.ContiguousSelection)
        self.liveView.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.liveView.clicked.connect(self.showcontent)
        vh = self.liveView.verticalHeader()
        vh.setVisible(False)
        hh = self.liveView.horizontalHeader()
        hh.setStretchLastSection(True)
        hh.setVisible(True)
        # set archive table view
        #self.archiveView = QtWidgets.QTableView(self.centralwidget)
        #self.archiveView.setGeometry(QtCore.QRect(820, 50, 600, 300))
        #self.archiveView.setObjectName("archive_attributes")
        #self.archiveView.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        #self.archiveView.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        #vh = self.archiveView.verticalHeader()
        #vh.setVisible(False)
        #hh = self.archiveView.horizontalHeader()
        #hh.setStretchLastSection(True)
        #hh.setVisible(True)
        # build inbox models
        for account in self.cache["index"]:
            model = InboxModel(
                self.cache["index"][account]["inbox"]["attributes"],
                self.cache["index"][account]["inbox"]["ids"],
                self.cache["index"][account]["inbox"]["archive"],
                self.cache["index"][account]["inbox"]["submissions"]
            )
            self.cache["index"][account]["inbox"]["object"] = model
        # set active element to first account
        self.cache["active"] = list(self.cache["index"])[0]
        # set view model to active model
        self.liveView.setModel(self.cache["index"][self.cache["active"]]["inbox"]["object"])
        # create live table label
        self.attributesLabel = QtWidgets.QLabel(self.centralwidget)
        self.attributesLabel.setGeometry(QtCore.QRect(180, 0, 200, 50))
        self.attributesLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.attributesLabel.setObjectName("attributes_label")
        # create live table control panel
        self.liveControl = QtWidgets.QFrame(self.centralwidget)
        self.liveControl.setGeometry(QtCore.QRect(380, 0, 400, 50))
        self.liveControl.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.liveControl.setFrameShadow(QtWidgets.QFrame.Raised)
        self.liveControl.setObjectName("attributes_control")
        # create live table control buttons
        self.deleteButton = QtWidgets.QPushButton(self.liveControl)
        self.deleteButton.setGeometry(QtCore.QRect(10, 10, 100, 30))
        self.deleteButton.setObjectName("delete_items")
        #self.deleteButton.clicked.connect(self.removeitems)
        self.archiveButton = QtWidgets.QPushButton(self.liveControl)
        self.archiveButton.setGeometry(QtCore.QRect(170, 10, 100, 30))
        self.archiveButton.setObjectName("archive_items")
        #self.archiveButton.clicked.connect(self.archiveitems)
        self.submissionsButton = QtWidgets.QPushButton(self.liveControl)
        self.submissionsButton.setGeometry(QtCore.QRect(170, 10, 100, 30))
        self.submissionsButton.setObjectName("archive_items")
        #self.submissionsButton.clicked.connect(self.submititems)
        # create email content viewer
        self.contentWidget = QtWidgets.QTextBrowser(self.centralwidget)
        self.contentWidget.setGeometry(QtCore.QRect(200, 400, 600, 200))
        self.contentWidget.setObjectName("email_content")
        # create email content label
        self.content_label = QtWidgets.QLabel(self.centralwidget)
        self.content_label.setGeometry(QtCore.QRect(180, 350, 600, 50))
        self.content_label.setAlignment(QtCore.Qt.AlignCenter)
        self.content_label.setObjectName("content_label")
        # create control panel
        self.control = QtWidgets.QFrame(self.centralwidget)
        self.control.setGeometry(QtCore.QRect(5, 400, 180, 250))
        self.control.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.control.setFrameShadow(QtWidgets.QFrame.Raised)
        self.control.setObjectName("control_panel")
        # create control panel label
        self.control_label = QtWidgets.QLabel(self.centralwidget)
        self.control_label.setGeometry(QtCore.QRect(5, 350, 180, 50))
        self.control_label.setAlignment(QtCore.Qt.AlignCenter)
        self.control_label.setObjectName("control_label")
        # create control panel buttons
        self.updateButton = QtWidgets.QPushButton(self.control)
        self.updateButton.setGeometry(QtCore.QRect(5, 0, 170, 30))
        self.updateButton.setObjectName("pushButton_2")
        self.updateButton.clicked.connect(self.update)
        self.masterButton = QtWidgets.QPushButton(self.control)
        self.masterButton.setGeometry(QtCore.QRect(5, 40, 170, 30))
        self.masterButton.setObjectName("pushButton_3")
        self.masterButton.clicked.connect(self.makemaster)

        self.setCentralWidget(self.centralwidget)
        QtCore.QMetaObject.connectSlotsByName(self)
        self.retranslateUi()

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "Outlook Parser"))
        #__sortingEnabled = self.tableModel.isSortingEnabled()
        #self.tableModel.setSortingEnabled(False)

        for key in self.cache["index"]:
            self.cache["index"][key]["object"].setText(_translate("MainWindow", key))

        self.updateButton.setText(_translate("MainWindow", "Update"))
        self.masterButton.setText(_translate("MainWindow", "Make Master"))
        self.deleteButton.setText(_translate("MainWindow", "Delete Email(s)"))
        self.archiveButton.setText(_translate("MainWindow", "Archive Email(s)"))

        #for account in self.cache["index"]:
            #for email in self.cache["index"][account]["inbox"]:
                #item = self.cache["index"][account]["inbox"][email]["object"]
                #text = self.cache["index"][account]["inbox"][email]["text"]
                #item.setText(_translate("MainWindow", text))

        self.sidebar_label.setText(_translate("MainWindow", "Inbox Selector"))
        self.content_label.setText(_translate("MainWindow", "Email Content"))
        self.attributesLabel.setText(_translate("MainWindow", "Email Attributes"))
        self.control_label.setText(_translate("MainWindow", "Control Panel"))

    def flip(self):
        sender_name = self.sender().objectName()
        model = self.cache["index"][sender_name]["inbox"]["object"]
        self.liveView.setModel(model)
        self.cache["active"] = sender_name
        self.clearcontent()

    def showcontent(self, table_click):
        _translate = QtCore.QCoreApplication.translate

        row = table_click.rows()
        print(row)

        #account = self.cache["active"]
        #content = self.cache["index"][account]["inbox"]["main"]["content"][row]
        #self.contentWidget.setHtml(_translate("MainWindow", content))

    def clearcontent(self):
        _translate = QtCore.QCoreApplication.translate

        self.contentWidget.setHtml(_translate("MainWindow", ""))

    def update(self):
        extension = sess.update(gui=True)
        for account in extension:
            if extension[account] is not None:
                attributes, content = [], []
                for row in extension[account]:
                    attributes.append([*row[1:-1]])
                    content.append(row[-1])

                attributes = pandas.DataFrame(attributes)
                content = pandas.Series(content)

                self.cache["index"][account]["inbox"]["attributes"] = self.cache["index"][account]["inbox"]["attributes"].append(attributes, ignore_index=True)
                self.cache["index"][account]["inbox"]["content"] = self.cache["index"][account]["inbox"]["content"].append(content, ignore_index=True)
                self.cache["index"][account]["inbox"]["object"].addRows(attributes)

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
