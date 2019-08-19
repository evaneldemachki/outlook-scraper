import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets
import core
import json
import pandas
import stylesheet

def bridge():
    config = sess.config
    #sess.update()

    index = {}
    for account in config["index"]:
        entry = {
            "object": None,
            "inbox": {
                "live": {},
                "archive": {}
            }
        }

        live_data = sess.load(account)
        arch_data = sess.load_archive(account)
        # generate archive index for live data model
        arch_index = []
        for id in live_data['entry_id']:
            match = arch_data.loc[arch_data['entry_id'] == id]['category']
            if match.shape[0] > 0:
                arch_index.append(match.iloc[0])
            else:
                arch_index.append(False)
        # add archive_index column to live data
        arch_index = pandas.Series(arch_index)
        live_data['archive_index'] = arch_index

        # rename columns to numbered
        #live_data.columns = [0,1,2,3]
        entry["inbox"]["live"] = {
            "object": None,
            "data": live_data
        }
        entry["inbox"]["archive"] = {
            "object": None,
            "data": arch_data
        }

        index[account] = entry

    return index

class InboxModel(QtCore.QAbstractTableModel):
    def __init__(self, data, parent=None):
        super().__init__(parent)

        self.headings = live_headings
        self.origin = data.copy()
        self.dataframe = self.origin[
            [
                'sent_at',
                'email',
                'name',
                'subject'
            ]
        ]

    def rowCount(self, parent):
        return self.dataframe.shape[0]

    def columnCount(self, parent):
        return self.dataframe.shape[1]

    def data(self, index, role):
        if role == QtCore.Qt.DisplayRole:
            return self.dataframe.iloc[index.row()][index.column()]
        elif role == QtCore.Qt.BackgroundRole:
            if self.origin.loc[index.row(), "archive_index"] != False:
                return QtGui.QBrush(QtCore.Qt.lightGray)
            else:
                return QtCore.QVariant()
        else:
            return QtCore.QVariant()

    def headerData(self, section, orientation, role):
        if role == QtCore.Qt.DisplayRole:
            return self.headings[section]
        else:
            return QtCore.QVariant()

    def addRows(self, data):
        position = self.dataframe.shape[0] - 1
        end = data.shape[0] - 1
        # begin model insertion
        self.beginInsertRows(QtCore.QModelIndex(), position, end)
        # redefine origin + dataframe
        self.origin = data.copy()
        self.dataframe = self.origin[
            [
                'sent_at',
                'email',
                'name',
                'subject'
            ]
        ]
        # end model insertion
        self.endInsertRows()

    def removeRows(self, data):
        position = data.shape[0]
        end = self.dataframe.shape[0] - 1
        self.beginRemoveRows(QtCore.QModelIndex(), position, end)
        self.origin = data.copy()
        self.dataframe = self.origin[
            [
                'sent_at',
                'email',
                'name',
                'subject'
            ]
        ]
        self.endRemoveRows()

    def addToArchive(self, data, rows, category=None):
        self.origin = data.copy()
        self.dataframe = self.origin[
            [
                'sent_at',
                'email',
                'name',
                'subject'
            ]
        ]

class ArchiveModel(QtCore.QAbstractTableModel):
    def __init__(self, data, parent=None):
        super().__init__(parent)

        self.headings = arch_headings
        self.origin = data.copy()
        self.dataframe = self.origin[
            [
                'sent_at',
                'email',
                'name',
                'subject',
                'category'
            ]
        ]

    def rowCount(self, parent):
        return self.dataframe.shape[0]

    def columnCount(self, parent):
        return self.dataframe.shape[1]

    def data(self, index, role):
        if role == QtCore.Qt.DisplayRole:
            return self.dataframe.iloc[index.row()][index.column()]

        return QtCore.QVariant()

    def headerData(self, section, orientation, role):
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        return self.headings[section]

    def addRows(self, data):
        position = self.dataframe.shape[0]
        end = data.shape[0] - 1
        # begin model insertion
        self.beginInsertRows(QtCore.QModelIndex(), position, end)
        # update origin and dataframe
        self.origin = data.copy()
        self.dataframe = self.origin[
            [
                'sent_at',
                'email',
                'name',
                'subject',
                'category'
            ]
        ]
        # end model insertion
        self.endInsertRows()

    def removeRows(self, data):
        position = data.shape[0]
        end = self.dataframe.shape[0] - 1
        self.beginRemoveRows(QtCore.QModelIndex(), position, end)
        self.origin = data.copy()
        self.dataframe = self.origin[
            [
                'sent_at',
                'email',
                'name',
                'subject',
                'category'
            ]
        ]
        self.endRemoveRows()

class FilterModel(QtCore.QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.filter_value = ''

    def setFilterValue(self, val):
        self.filter_value = val.lower()

    def filterAcceptsRow(self, sourceRow, sourceParent):
        index = self.sourceModel().index(
            sourceRow,
            self.filterKeyColumn(),
            sourceParent
        )
        data = self.sourceModel().data(index, QtCore.Qt.DisplayRole)

        return self.filter_value in data

class Ui_MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self.cache = {
            "index": bridge()
        }

        self.active = None
        self.view = None
        self.proxy_active = False

        self.setupUi()
        self.show()

    def setupUi(self):
        _translate = QtCore.QCoreApplication.translate
        # set main window parameters
        self.setObjectName("MainWindow")
        self.resize(1440, 850)
        # create central window widget
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        # create inbox selection sidebar
        self.sidebar = QtWidgets.QFrame(self.centralwidget)
        self.sidebar.setGeometry(QtCore.QRect(5, 50, 180, 200))
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
        self.mainView = QtWidgets.QTableView(self.centralwidget)
        self.mainView.setGeometry(QtCore.QRect(200, 50, 800, 400))
        self.mainView.setObjectName("email_attributes")
        self.mainView.setSelectionMode(QtWidgets.QAbstractItemView.ContiguousSelection)
        self.mainView.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.mainView.setStyleSheet(stylesheet.tableStyle)
        self.mainView.clicked.connect(self.showcontent)
        vh = self.mainView.verticalHeader()
        vh.setVisible(False)
        hh = self.mainView.horizontalHeader()
        hh.setStretchLastSection(True)
        hh.setVisible(True)
        # create table models
        for account in self.cache["index"]:
            model = InboxModel(
                self.cache["index"][account]["inbox"]["live"]["data"]
            )
            self.cache["index"][account]["inbox"]["live"]["object"] = model
            archive_model = ArchiveModel(
                self.cache["index"][account]["inbox"]["archive"]["data"]
            )
            self.cache["index"][account]["inbox"]["archive"]["object"] = archive_model
        # create proxy model
        self.filterModel = FilterModel()
        # set active element to first account
        self.active = list(self.cache["index"])[0]
        # set mode to live
        self.mode = "live"
        # set view model to active model
        self.mainView.setModel(self.cache["index"][self.active]["inbox"]["live"]["object"])
        # create live table label
        self.attributesLabel = QtWidgets.QLabel(self.centralwidget)
        self.attributesLabel.setGeometry(QtCore.QRect(200, 0, 120, 50))
        self.attributesLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.attributesLabel.setObjectName("attributes_label")
        # create live table control panel
        self.liveControl = QtWidgets.QFrame(self.centralwidget)
        self.liveControl.setGeometry(QtCore.QRect(320, 0, 280, 50))
        self.liveControl.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.liveControl.setFrameShadow(QtWidgets.QFrame.Raised)
        self.liveControl.setObjectName("attributes_control")
        # create live table control buttons
        self.deleteButton = QtWidgets.QPushButton(self.liveControl)
        self.deleteButton.setGeometry(QtCore.QRect(5, 10, 120, 30))
        self.deleteButton.setObjectName("delete_items")
        self.deleteButton.clicked.connect(self.deleterows)
        self.archiveButton = QtWidgets.QPushButton(self.liveControl)
        self.archiveButton.setGeometry(QtCore.QRect(130, 10, 120, 30))
        self.archiveButton.setObjectName("archive_items")
        self.archiveButton.clicked.connect(self.archiverows)
        # create email content viewer
        self.contentWidget = QtWidgets.QTextBrowser(self.centralwidget)
        self.contentWidget.setGeometry(QtCore.QRect(200, 500, 800, 300))
        self.contentWidget.setObjectName("email_content")
        # create email content label
        self.content_label = QtWidgets.QLabel(self.centralwidget)
        self.content_label.setGeometry(QtCore.QRect(200, 450, 800, 50))
        self.content_label.setAlignment(QtCore.Qt.AlignCenter)
        self.content_label.setObjectName("content_label")
        # create control panel
        self.control = QtWidgets.QFrame(self.centralwidget)
        self.control.setGeometry(QtCore.QRect(5, 300, 180, 250))
        self.control.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.control.setFrameShadow(QtWidgets.QFrame.Raised)
        self.control.setObjectName("control_panel")
        # create control panel label
        self.control_label = QtWidgets.QLabel(self.centralwidget)
        self.control_label.setGeometry(QtCore.QRect(5, 250, 180, 50))
        self.control_label.setAlignment(QtCore.Qt.AlignCenter)
        self.control_label.setObjectName("control_label")
        # create control panel buttons
        self.updateButton = QtWidgets.QPushButton(self.control)
        self.updateButton.setGeometry(QtCore.QRect(5, 0, 170, 30))
        self.updateButton.setObjectName("update_button")
        self.updateButton.clicked.connect(self.update)
        self.masterButton = QtWidgets.QPushButton(self.control)
        self.masterButton.setGeometry(QtCore.QRect(5, 40, 170, 30))
        self.masterButton.setObjectName("master_button")
        self.masterButton.clicked.connect(self.makemaster)
        self.modeButton = QtWidgets.QPushButton(self.control)
        self.modeButton.setGeometry(QtCore.QRect(5, 80, 170, 30))
        self.modeButton.setObjectName("mode_button")
        self.modeButton.clicked.connect(self.switchmode)

        self.filterPanel = QtWidgets.QFrame(self.centralwidget)
        self.filterPanel.setGeometry(QtCore.QRect(600, 0, 400, 50))
        self.dropColumns = QtWidgets.QComboBox(self.filterPanel)
        self.dropColumns.setGeometry(QtCore.QRect(5, 10, 170, 30))
        self.dropColumns.currentIndexChanged[str].connect(self.changecolumn)
        for column in live_headings:
            self.dropColumns.addItem(column)
        self.searchValue = QtWidgets.QLineEdit(self.filterPanel)
        self.searchValue.setGeometry(QtCore.QRect(180, 10, 215, 30))
        self.searchValue.textChanged[str].connect(self.filter)

        self.setCentralWidget(self.centralwidget)
        QtCore.QMetaObject.connectSlotsByName(self)
        self.retranslateUi()

    def changecolumn(self, text):
        if self.proxy_active:
            col = self.translate_column(text)
            self.mainView.model().setFilterKeyColumn(col)

    def translate_column(self, column):
        if self.mode == "live":
            headings = live_headings
        else:
            headings = arch_headings

        return headings.index(column)

    def setfilter(self, model, column, text):
        col = self.translate_column(column)
        self.filterModel.setSourceModel(model)
        self.filterModel.setFilterKeyColumn(col)
        self.filterModel.setFilterValue(text)
        self.mainView.setModel(self.filterModel)
        self.filterModel.setFilterFixedString("")
        self.clearcontent()

    def filter(self, text):
        if not self.proxy_active:
            model = self.mainView.model()
            column = self.dropColumns.currentText()
            self.setfilter(model, column, text)
            #self.filter

            self.proxy_active = True
        else:
            if text == '':
                self.filterModel.setFilterValue('')
                model = self.cache["index"][self.active]["inbox"][self.mode]["object"]
                self.mainView.setModel(model)

                self.proxy_active = False
                return

            column = self.dropColumns.currentText()
            self.filterModel.setFilterValue(text)
            self.filterModel.setFilterFixedString("")


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
        self.modeButton.setText(_translate("MainWindow", "Archive Mode"))

        self.sidebar_label.setText(_translate("MainWindow", "Inbox Selector"))
        self.content_label.setText(_translate("MainWindow", "Email Content"))
        self.attributesLabel.setText(_translate("MainWindow", "Email Attributes"))
        self.control_label.setText(_translate("MainWindow", "Control Panel"))

    def flip(self):
        account = self.sender().objectName()
        if account != self.active:
            model = self.cache["index"][account]["inbox"][self.mode]["object"]
            if not self.proxy_active:
                self.mainView.setModel(model)
            else:
                column = self.dropColumns.currentText()
                text = self.searchValue.text()
                self.setfilter(model, column, text)

            self.active = account
            self.clearcontent()

    def switchmode(self):
        _translate = QtCore.QCoreApplication.translate
        if self.mode == "live":
            mode = "archive"
            button_text = "Live View"
        else:
            mode = "live"
            button_text = "Archive View"

        model = self.cache["index"][self.active]["inbox"][mode]["object"]
        if not self.proxy_active:
            self.mainView.setModel(model)
        else:
            column = self.dropColumns.currentText()
            text = self.searchValue.text()
            self.setfilter(model, column, text)

        self.mode = mode
        self.modeButton.setText(_translate("MainWindow", button_text))
        self.clearcontent()


    def selectedids(self):
        if not self.proxy_active:
            model = self.mainView.model()
            # retrieve selected rows
            rows = self.mainView.selectionModel().selectedRows()
            model_rows = sorted([row.row() for row in rows])
            # convert selected rows to entry_ids
            entry_ids = [model.origin['entry_id'][n] for n in model_rows]
        else:
            model = self.cache["index"][self.active]["inbox"][self.mode]["object"]
            index_list = self.mainView.selectionModel().selectedRows()
            proxy = self.mainView.model()
            model_rows = list()
            for index in index_list:
                model_rows.append(proxy.mapToSource(index).row())

            entry_ids = [model.origin['entry_id'][n] for n in model_rows]

        return entry_ids, model_rows

    def deleterows(self):
        if self.mode == "live":
            archive = False
        if self.mode == "archive":
            archive = True

        account = self.active
        model = self.cache["index"][self.active]["inbox"][self.mode]["object"]
        # get selected entry ids
        entry_ids, model_rows = self.selectedids()
        if len(entry_ids) > 0:
            # retrieve cache data
            cache = self.cache["index"][account]["inbox"][self.mode]
            # use entry_ids to obtain cache row #'s
            df = cache['data']
            cache_rows = [int(df[df['entry_id'] == id].index[0]) for id in entry_ids]
            # execute delete function in session
            sess.delete_many(account, entry_ids, archive=archive)
            # remove rows from cache
            cache["data"].drop(cache_rows, axis=0, inplace=True)
            cache["data"].reset_index(drop=True, inplace=True)
            # remove rows from model
            model.removeRows(cache["data"])
            if self.proxy_active:
                self.filterModel.setFilterFixedString("")

    def archiverows(self):
        if self.mode == "live":
            account = self.active
            model = self.cache["index"][self.active]["inbox"][self.mode]["object"]
            # get selected entry ids
            entry_ids, model_rows = self.selectedids()
            # retrieve cache data
            live_cache = self.cache["index"][account]["inbox"]["live"]
            arch_cache = self.cache["index"][account]["inbox"]["archive"]
            # use entry_ids to obtain live cache row #'s
            df = live_cache['data']
            cache_rows = [int(df[df['entry_id'] == id].index[0]) for id in entry_ids]
            # remove already-archived entry_ids from entry_ids & row_nums
            for i in range(len(entry_ids) - 1, -1, -1):
                id = entry_ids[i]
                if id in list(arch_cache['data']['entry_id']):
                    entry_ids.pop(i)
                    cache_rows.pop(i)
                    model_rows.pop(i)

            if len(entry_ids) > 0:
                # execute archive function in session
                sess.archive_many(account, entry_ids)
                # REPLACE NONE WITH CATEGORY
                # update live cache's archive index
                for n in cache_rows:
                    live_cache["data"].loc[n, "archive_index"] = None
                # add archived rows to archive cache
                #   assume archive_index column = category column
                arch_ext = pandas.DataFrame([live_cache["data"].loc[n].copy() for n in cache_rows])
                # rename archive_index column to category
                arch_ext.rename(columns={"archive_index": "category"}, inplace=True)
                arch_cache["data"] = arch_cache["data"].append(arch_ext, ignore_index=True)
                # update live model archive index
                model.addToArchive(live_cache["data"], model_rows)
                # update archive model
                archive_model = self.cache["index"][account]["inbox"]["archive"]["object"]
                archive_model.addRows(arch_cache["data"])

            self.mainView.clearSelection()

    def showcontent(self, table_click):
        _translate = QtCore.QCoreApplication.translate

        if not self.proxy_active:
            row = table_click.row()
            model = self.mainView.model()
            content = model.origin["content"][row]
        else:
            model = self.cache["index"][self.active]["inbox"][self.mode]["object"]
            row = self.mainView.model().mapToSource(table_click).row()
            content = model.origin["content"][row]

        self.contentWidget.setHtml(_translate("MainWindow", content))

    def clearcontent(self):
        _translate = QtCore.QCoreApplication.translate
        self.contentWidget.setHtml(_translate("MainWindow", ""))

    def update(self):
        extension = sess.update(gui=True)
        for account in extension:
            print(account)
            test_model = self.cache["index"][account]["inbox"]["live"]["object"]
            print("...[PRE] model row count: {0}".format(test_model.rowCount(None)))
            if extension[account] is not None:
                live_ext = extension[account]
                #for row in extension[account]:
                    #live_ext.append(row)


                #live_ext = pandas.DataFrame(live_ext)
                #print(live_ext)
                live_ext["archive_index"] = [False for i in range(live_ext.shape[0])]
                live_cache = self.cache["index"][account]["inbox"]["live"]
                live_cache["data"] = live_cache["data"].append(live_ext, ignore_index=True)

                live_model = self.cache["index"][account]["inbox"]["live"]["object"]
                live_model.addRows(live_cache["data"])
                print("...[POST]: model row count: {0}".format(live_model.rowCount(None)))

    def makemaster(self):
        sess.makemaster()

sess = core.Session()
live_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject"]
arch_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject", "Category"]

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
