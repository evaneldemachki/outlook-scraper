import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets
from functools import partial

_translate = QtCore.QCoreApplication.translate

class CategoryBox(QtWidgets.QFrame):
    def __init__(self, name, color, parent):
        super().__init__(parent)
        self.setObjectName(name)
        self.addLabel(name)
        self.initStyle(color)
        self.color = color
    
    def addLabel(self, name):
        self.label = QtWidgets.QLabel(self)
        self.label.setText(_translate("ArchiveWindow", name))
        self.label.setGeometry(QtCore.QRect(0,0,80,80))
        self.label.setAlignment(QtCore.Qt.AlignCenter)
    
    def initStyle(self, color):
        #qcolor = QtGui.QColor(color)
        self.setStyleSheet(
            """
            [Active=true] {{
                background-color: {0}; 
                border-style: outset;
                border-width: 2px;
                border-color: black;
            }}
            [Active=false] {{
                background-color: {0}; 
                border-style: outset;
                border-width: 0;
                border-color: black;
            }}
            """.format(color)
        )
        self.setProperty('Active', False)
        #self.setAutoFillBackground(True)
        #bg = self.palette()
        #bg.setColor(self.backgroundRole(), qcolor)
        #self.setPalette(bg)
    
class ArchiveWindow(QtWidgets.QDialog):
    def __init__(self, categoryselector, parent):
        super().__init__(parent)

        self.setWindowTitle("Select Category")
        self.value = None
        
        self.categories = categoryselector.index
        self.setupUi()
    
    def setupUi(self):
        # set window parameters
        self.setObjectName("ArchiveWindow")
        self.resize(500, 500)
        # create category selection
        self.objects = {}
        dim = [10, 10, 80, 80]
        setValue = lambda val, _: self.setValue(val)
        for category, entry in self.categories.items():
            obj = CategoryBox(category, entry["color"], self)
            obj.setGeometry(QtCore.QRect(*dim))
            obj.mousePressEvent = partial(setValue, category)

            self.objects[category] = obj
            dim[0] += 90
        
        self.addButtons()
        QtCore.QMetaObject.connectSlotsByName(self)
    
    def addButtons(self):
        self.cancel = QtWidgets.QPushButton(self)
        self.cancel.setText(_translate("ArchiveWindow", "Cancel"))
        self.cancel.setGeometry(QtCore.QRect(100, 450, 100, 50))
        self.submit = QtWidgets.QPushButton(self)
        self.submit.setText(_translate("ArchiveWindow", "Submit"))
        self.submit.setGeometry(QtCore.QRect(300, 450, 100, 50))
        self.submit.clicked.connect(self.close)
    
    def setValue(self, value):
        if self.value != value:
            if self.value != None:
                # remove current object border
                obj = self.objects[self.value]
                obj.setProperty('Active', False)
                obj.style().unpolish(obj)
                obj.style().polish(obj)
            # activate new object border
            obj = self.objects[value]
            obj.setProperty('Active', True)
            obj.style().unpolish(obj)
            obj.style().polish(obj)
            # set current value to new value
            self.value = value
        elif self.value == value:
            # remove current object border
            obj = self.objects[self.value]
            obj.setProperty('Active', False)
            obj.style().unpolish(obj)
            obj.style().polish(obj)

            self.value = None

    def submitclose(self):
        return self.value

