import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets

class CategorySelector(QtWidgets.QFrame):
    def __init__(self, colors, rect, parent):
        super().__init__(parent)
        self.setGeometry(QtCore.QRect(*rect))

        _translate = QtCore.QCoreApplication.translate
        self.index = {}
        # generate category fields
        top_pos = 0
        w, l = rect[2], rect[3]
        for category, color in colors.items():

            self.index[category] = {}
            # create category field container
            category_field = QtWidgets.QFrame(self)
            category_field.setGeometry(QtCore.QRect(0, top_pos, w, 50))
            category_field.setObjectName(category)
            # create category color display box
            category_color = QtWidgets.QFrame(category_field)
            category_color.setGeometry(QtCore.QRect(5, 5, 40, 40))
            category_color.setObjectName("color")
            # set display box color
            category_color.setAutoFillBackground(True)
            bg = category_color.palette()
            bg.setColor(category_color.backgroundRole(), QtGui.QColor(color))
            category_color.setPalette(bg)
            # create category label
            category_label = QtWidgets.QLabel(category_field)
            category_label.setGeometry(QtCore.QRect(50, 0, w-50, 50))
            category_label.setText(_translate("MainWindow", category.capitalize()))
            category_label.setObjectName("label")
            self.index[category]["object"] = category_field
            self.index[category]["color"] = color

            top_pos += 50

    def addcategory(self, category, color):
        pass

    def dropcategory(self, category):
        pass

    def recolor(self, category, color):
        pass
