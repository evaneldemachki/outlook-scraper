import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets

live_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject"]
arch_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject", "Category"]

class FilterModel(QtCore.QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.filter_value = ''
        self.filter_categories = []
        self.is_archive = False

    def setFilterValue(self, val):
        self.filter_value = val.lower()
        print("<FILTER MODEL> filter_categories={0}".format(self.filter_categories))
        print("<FILTER MODEL> filterKeyColumn={0}".format(self.filterKeyColumn()))
        print("<FILTER MODEL> filter_value={0}".format(self.filter_value))
    
    def setFilterCategories(self, categories):
        self.filter_categories = categories
        if len(self.filter_categories) == 0:
            self.has_categories = False
        else:
            self.has_categories = True

        print("<FILTER MODEL> filter_categories={0}".format(categories))
        print("<FILTER MODEL> filterKeyColumn={0}".format(self.filterKeyColumn()))
        print("<FILTER MODEL> filter_value={0}".format(self.filter_value))

    def isArchive(self, val):
        print("<FILTER MODEL> is_archive={0}".format(val))
        self.is_archive = val

    def filterAcceptsRow(self, sourceRow, sourceParent):
        index = self.sourceModel().index(
            sourceRow,
            self.filterKeyColumn(),
            sourceParent
        )
        val_data = self.sourceModel().data(index, QtCore.Qt.DisplayRole)
        if not self.is_archive:
            return self.filter_value in val_data.lower()
        elif self.has_categories:
            index = self.sourceModel().index(
                sourceRow,
                arch_headings.index("Category"),
                sourceParent
            )
            cat_data = self.sourceModel().data(index, QtCore.Qt.DisplayRole)

            return self.filter_value in val_data.lower() and cat_data in self.filter_categories
        else:
            return self.filter_value in val_data.lower()