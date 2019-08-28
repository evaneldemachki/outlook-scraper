import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets
from models.inbox import InboxModel
from models.archive import ArchiveModel
from models.filter import FilterModel
from widgets.category_select import CategorySelector
from widgets.archive_window import ArchiveWindow
import bridge
import pandas
import stylesheet
import sys
from exceptions import UiRuntimeError
import random
from functools import partial

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self.cache = bridge.init()
        self.active = None
        self.mode = None

        self.setupUi()
        self.show()

    def setupUi(self):
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
            pushButton.clicked.connect(self.switchaccount)
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
        model = self.getmodel(mode="live")
        self.setview(model)
        # initalize proxy model in main view
        self.mainView.setModel(self.filterModel)
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
        self.archiveButton.clicked.connect(self.archive_window)
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
        self.modeButton.clicked.connect(self.togglemode)

        self.filterPanel = QtWidgets.QFrame(self.centralwidget)
        self.filterPanel.setGeometry(QtCore.QRect(600, 0, 400, 50))
        self.dropColumns = QtWidgets.QComboBox(self.filterPanel)
        self.dropColumns.setGeometry(QtCore.QRect(5, 10, 170, 30))
        self.dropColumns.currentIndexChanged[str].connect(self.set_filter_column)
        for column in live_headings:
            self.dropColumns.addItem(column)
        self.searchValue = QtWidgets.QLineEdit(self.filterPanel)
        self.searchValue.setGeometry(QtCore.QRect(180, 10, 215, 30))
        self.searchValue.textChanged[str].connect(self.set_filter_text)
        # create category filter dict
        self.category_filters = {}

        rect = (1000, 50, 440, 400)
        self.categorySelector = CategorySelector(self.cache["categories"], rect, self.centralwidget)
        mp_func = lambda obj_name, null: self.selectcategory(obj_name)
        for category in self.categorySelector.index:
            self.category_filters[category] = False
            obj = self.categorySelector.index[category]["object"]
            obj.setCursor(QtCore.Qt.PointingHandCursor)
            obj.mousePressEvent = partial(mp_func, obj.objectName())
            #obj.setStyleSheet("QFrame::item:hover{background-color:#999966;}")

        self.setCentralWidget(self.centralwidget)
        QtCore.QMetaObject.connectSlotsByName(self)
        self.retranslateUi()

    def retranslateUi(self):
        self.setWindowTitle(_translate("MainWindow", "Outlook Parser"))

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

    def selectcategory(self, category):
        obj = self.categorySelector.index[category]["object"]

        if not self.category_filters[category]:
            style = "QFrame#{0}".format(category) + " {background-color:lightgray;}"
            obj.setStyleSheet(style)
            self.category_filters[category] = True
            # set filter model category filters
            filter_categories = []
            for cat, val in self.category_filters.items():
                if val:
                    filter_categories.append(cat)

            self.filterModel.setFilterCategories(filter_categories)
            self.filterModel.setFilterFixedString("")
            self.clearcontent()
        else:
            style = "QFrame#{0}".format(category) + " {}"
            obj.setStyleSheet(style)
            self.category_filters[category] = False
            # set filter model category filters
            filter_categories = []
            for cat, val in self.category_filters.items():
                if val:
                    filter_categories.append(cat)

            self.filterModel.setFilterCategories(filter_categories)
            self.filterModel.setFilterFixedString("")
            self.clearcontent()

    def archive_window(self):
        if self.mode != 'live':
            return
        # get selected entry ids
        entry_ids, model_rows = self.selectedids()
        if entry_ids is None:
            return

        popup = ArchiveWindow(self.categorySelector, self)
        popup.setModal(True)
        popup.exec_()
        category = popup.submitclose()
        # call archive function
        self.archiverows(entry_ids, model_rows, category)

    def set_filter_column(self): # update column filter key index
        # get column value from column dropdown
        column = self.dropColumns.currentText()
        # convert column index
        if self.mode == "live":
            headings = live_headings
        else:
            headings = arch_headings

        column_index = headings.index(column)
        # set filter key column to column_index
        self.filterModel.setFilterKeyColumn(column_index)
        # refresh proxy model in view
        self.filterModel.setFilterFixedString("")

    def set_filter_text(self, text): # update text filter key string
        # set filter key string to search bar text
        self.filterModel.setFilterValue(text)
        # refresh proxy model in view
        self.filterModel.setFilterFixedString("")

    def getmodel(self, account=None, mode=None): # return model from cache
        if account is None:
            account = self.active
        if mode is None:
            mode = self.mode

        return self.cache["index"][account]["inbox"][mode]["object"]

    def setview(self, model): # set view to model
        # set filter model source
        self.filterModel.setSourceModel(model)
        # refresh filter model in view
        self.filterModel.setFilterFixedString('')

    def switchaccount(self): # switch account view
        account = self.sender().objectName()
        if account != self.active:
            model = self.getmodel(account=account)
            self.setview(model)
            self.active = account
            self.clearcontent()

    def togglemode(self): # toggle between live/archive view mode
        if self.mode == "live":
            mode = "archive"
            button_text = "Live View"
            archive_flag = True
        else:
            mode = "live"
            button_text = "Archive View"
            archive_flag = False

        model = self.getmodel(mode=mode)
        self.setview(model)
        self.mode = mode
        self.modeButton.setText(_translate("MainWindow", button_text))
        # set proxy model archive flag
        self.filterModel.isArchive(archive_flag)
        # refresh proxy model in view
        self.filterModel.setFilterFixedString("")

        self.clearcontent()

    def selectedids(self): # get view selections' entry IDs & row #'s
        model = self.getmodel()
        proxy_model = self.mainView.model()
        # get selected rows in active proxy model
        proxy_rows = self.mainView.selectionModel().selectedRows()
        # if no rows selected:
        if len(proxy_rows) == 0: # return None
            return None, None
        # translate proxy row #'s to source model #'s
        source_rows = list()
        for index in proxy_rows:
            source_rows.append(proxy_model.mapToSource(index))

        source_rows = sorted([row.row() for row in source_rows])
        # convert selected rows to entry_ids
        entry_ids = [model.origin['entry_id'][n] for n in source_rows]

        return entry_ids, source_rows

    def deleterows(self): # delete rows
        if self.mode == "live":
            archive = False
        if self.mode == "archive":
            archive = True

        account = self.active
        model = self.getmodel()
        # get selected entry ids
        entry_ids, source_rows = self.selectedids()
        if len(entry_ids) > 0:
            # retrieve cache data
            cache = self.cache["index"][account]["inbox"][self.mode]
            # use entry_ids to obtain cache row #'s
            df = cache['data']
            cache_rows = [int(df[df['entry_id'] == id].index[0]) for id in entry_ids]
            # execute delete function in session
            bridge.sess().delete_many(account, entry_ids, archive=archive)
            # remove rows from cache
            cache["data"].drop(cache_rows, axis=0, inplace=True)
            cache["data"].reset_index(drop=True, inplace=True)
            # remove rows from model
            model.removeRows(cache["data"])
            # refresh proxy model in view
            self.filterModel.setFilterFixedString("")

    # add live rows to archive
    def archiverows(self, entry_ids, source_rows, category):
        if self.mode != "live": # if not in archive mode
            return # do not continue

        account = self.active
        model = self.getmodel()
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
                source_rows.pop(i)

        if len(entry_ids) > 0:
            # execute archive function in session
            bridge.sess().archive_many(account, entry_ids, category)
            # update live cache's archive index
            for n in cache_rows:
                live_cache["data"].loc[n, "archive_index"] = category
            # add archived rows to archive cache
            #   assume archive_index column = category column
            arch_ext = pandas.DataFrame([live_cache["data"].loc[n].copy() for n in cache_rows])
            # rename archive_index column to category
            arch_ext.rename(columns={"archive_index": "category"}, inplace=True)
            arch_cache["data"] = arch_cache["data"].append(arch_ext, ignore_index=True)
            # update live model archive index
            model.addToArchive(live_cache["data"], source_rows)
            # update archive model
            archive_model = self.cache["index"][account]["inbox"]["archive"]["object"]
            archive_model.addRows(arch_cache["data"])
            # refresh proxy model in view
            self.filterModel.setFilterFixedString("")
            # clear view selections
            self.mainView.clearSelection()

    def showcontent(self, table_click):
        model, proxy_model = self.getmodel(), self.mainView.model()
        row = proxy_model.mapToSource(table_click).row()
        content = model.origin["content"][row]

        self.contentWidget.setHtml(_translate("MainWindow", content))

    def clearcontent(self):
        self.contentWidget.setHtml(_translate("MainWindow", ""))

    #def clear_category_filters
    def update(self):
        extension = bridge.sess().update(gui=True)
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
                # refresh proxy model in view
                self.filterModel.setFilterFixedString("")

    def makemaster(self):
        bridge.sess().makemaster()

_translate = QtCore.QCoreApplication.translate
live_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject"]
arch_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject", "Category"]

def run():
    app = QtWidgets.QApplication(sys.argv)
    ui = MainWindow()
    sys.exit(app.exec_())

if __name__ == "__main__":
    run()
