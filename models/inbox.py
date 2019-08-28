import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets

live_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject"]
arch_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject", "Category"]

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