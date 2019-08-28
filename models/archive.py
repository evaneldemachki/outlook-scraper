import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets

live_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject"]
arch_headings = ["Timestamp", "Sender Address", "Sender Name", "Subject", "Category"]

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