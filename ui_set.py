from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from Excel_Action import Excel_action


class Ui_Form(object):

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(451, 225)
        self.tabWidget = QtWidgets.QTabWidget(Form)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 451, 191))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.progressBar = QtWidgets.QProgressBar(self.tab)
        self.progressBar.setGeometry(QtCore.QRect(170, 140, 271, 20))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.label_2 = QtWidgets.QLabel(self.tab)
        self.label_2.setGeometry(QtCore.QRect(40, 60, 54, 21))
        self.label_2.setObjectName("label_2")
        self.tickEdit = QtWidgets.QLineEdit(self.tab)
        self.tickEdit.setGeometry(QtCore.QRect(100, 60, 71, 20))
        self.tickEdit.setObjectName("tickEdit")
        self.label_3 = QtWidgets.QLabel(self.tab)
        self.label_3.setGeometry(QtCore.QRect(180, 60, 81, 21))
        self.label_3.setObjectName("label_3")
        self.newFileNameEdit = QtWidgets.QLineEdit(self.tab)
        self.newFileNameEdit.setGeometry(QtCore.QRect(270, 60, 121, 20))
        self.newFileNameEdit.setObjectName("newFileNameEdit")
        self.ExecuteBtn = QtWidgets.QPushButton(self.tab)
        self.ExecuteBtn.setGeometry(QtCore.QRect(314, 100, 81, 23))
        self.ExecuteBtn.setObjectName("ExecuteBtn")
        self.showLabel = QtWidgets.QLabel(self.tab)
        self.showLabel.setGeometry(QtCore.QRect(40, 140, 121, 21))
        self.showLabel.setObjectName("showLabel")
        self.widget = QtWidgets.QWidget(self.tab)
        self.widget.setGeometry(QtCore.QRect(40, 20, 351, 33))
        self.widget.setObjectName("widget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.widget)
        self.horizontalLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setSpacing(6)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.widget)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.pathEdit = QtWidgets.QLineEdit(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(12)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pathEdit.sizePolicy().hasHeightForWidth())
        self.pathEdit.setSizePolicy(sizePolicy)
        self.pathEdit.setObjectName("pathEdit")
        self.horizontalLayout.addWidget(self.pathEdit)
        self.openFileBtn = QtWidgets.QPushButton(self.widget)
        self.openFileBtn.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(2)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.openFileBtn.sizePolicy().hasHeightForWidth())
        self.openFileBtn.setSizePolicy(sizePolicy)
        self.openFileBtn.setMinimumSize(QtCore.QSize(21, 23))
        self.openFileBtn.setObjectName("openFileBtn")
        self.horizontalLayout.addWidget(self.openFileBtn)
        self.horizontalLayout.setStretch(1, 12)
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.tabWidget.addTab(self.tab_2, "")

        self.retranslateUi(Form)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(Form)
        self.tickEdit.setText("500")

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "工具箱"))
        self.label_2.setText(_translate("Form", "人数差值："))
        self.label_3.setText(_translate("Form", "生成后文件名："))
        self.newFileNameEdit.setText(_translate("Form", "TheNewExcel"))
        self.ExecuteBtn.setText(_translate("Form", "生成"))
        self.showLabel.setText(_translate("Form", "Thanks for use it!"))
        self.label.setText(_translate("Form", "文件选中："))
        self.openFileBtn.setText(_translate("Form", "..."))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("Form", "青年大学习中奖人筛选"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("Form", "Building..."))

    def actionList(self):
        def openfilebtn_action():
            openfilename = QFileDialog.getOpenFileUrl()
            global FilePathLine
            global FilePathNameLine
            FilePathLine = openfilename[0].path().replace(openfilename[0].fileName(), "")
            FilePathLine = FilePathLine.replace("/", "", 1)
            FilePathNameLine = openfilename[0].path().replace("/", "", 1)
            self.pathEdit.setText(FilePathNameLine)
            # print(("action!"))

        def execute_action():

            if (self.pathEdit.text() == "" or self.tickEdit.text() == "" or self.newFileNameEdit.text() == ""):
                QMessageBox.about(None, "Warning", "The edit place is null!")
            else:
                The_main_path = self.pathEdit.text()
                memberTick = self.tickEdit.text()
                Excel_action(FilePathLine,FilePathNameLine,int(self.tickEdit.text()),self.newFileNameEdit.text())
                QMessageBox.about(None, "Warning", "action!" + " " + The_main_path)

        self.openFileBtn.clicked.connect(openfilebtn_action)
        self.ExecuteBtn.clicked.connect(execute_action)
