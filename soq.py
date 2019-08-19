from PyQt5 import QtCore, QtGui, QtWidgets
import pandas
import sys

class TableModel(QtCore.QAbstractTableModel):
    def __init__(self, data, parent=None):
        super().__init__(parent)

        self.headings = ["name", "address", "number"]
        self.dataset = data

    def rowCount(self, parent):
        return self.dataset.shape[0]

    def columnCount(self, parent):
        return self.dataset.shape[0]

    def data(self, index, role):
        if role == QtCore.Qt.DisplayRole:
            return self.dataset.iloc[index.row()][index.column()]

        return QtCore.QVariant()

    def headerData(self, section, orientation, role):
        if role == QtCore.Qt.DisplayRole and orientation == QtCore.Qt.Horizontal:
            return self.headings[section]

        return QtCore.QVariant()

class ProxyTableModel(QtCore.QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.filter_str = ''

    def setFilterStr(self, val):
        self.filter_str = val.lower()

    def filterAcceptsRow(self, sourceRow, sourceParent):
        index = self.sourceModel().index(sourceRow, 0, sourceParent)
        data = self.sourceModel().data(index, QtCore.Qt.DisplayRole)

        return self.filter_str in data.lower()

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, data):
        super().__init__()
        self.setupUi(data)
        self.show()

    def setupUi(self, data):
        self.setObjectName("MainWindow")
        self.resize(500, 500)
        # create central display widget
        self.centralWidget = QtWidgets.QWidget(self)
        self.centralWidget.setObjectName("central_widget")
        # create table view
        self.tableView = QtWidgets.QTableView(self.centralWidget)
        self.tableView.setGeometry(QtCore.QRect(0, 50, 500, 450))
        self.tableView.setObjectName("email_attributes")
        # create models & set proxy model to view
        self.mainModel = TableModel(data)
        self.proxyModel = ProxyTableModel()
        self.proxyModel.setSourceModel(self.mainModel)
        self.tableView.setModel(self.proxyModel)
        # create column filter search bar
        self.searchBar = QtWidgets.QLineEdit(self.centralWidget)
        self.searchBar.setGeometry(QtCore.QRect(0, 0, 500, 50))
        self.searchBar.textChanged[str].connect(self.filternames)
        # set central widget & connect slots
        self.setCentralWidget(self.centralWidget)
        QtCore.QMetaObject.connectSlotsByName(self)

    def filternames(self, text):
        # set proxy model filter string
        self.proxyModel.setFilterStr(text)
        # set filterFixedString() to '' to refresh view
        self.proxyModel.setFilterFixedString('')

# underlying dataset for table model
df = {
    "name": ["John Smith", "Larry David", "George Washington"],
    "address": ["10 Forest Dr", "15 East St", "12 Miami Ln"],
    "phone_number": ["000-111-2222", "222-000-1111", "111-222-0000"]
}
data = pandas.DataFrame(df)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    ui = MainWindow(data)
    sys.exit(app.exec_())
