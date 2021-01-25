import sys

from PyQt5.QtWidgets import QApplication, QMainWindow

import ui_set

if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = ui_set.Ui_Form()
    ui.setupUi(MainWindow)
    MainWindow.show()
    ui.actionList()
    sys.exit(app.exec_())
