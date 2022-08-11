from PyQt5 import QtWidgets, QtGui
import sys
import main_ui
import Picture_rc


class ISC_program:
    def __init__(self):
        app = QtWidgets.QApplication(sys.argv)
        app.setWindowIcon(QtGui.QIcon('Pukontu.ico'))
        ui = main_ui.MainUI()
        ui.show()
        app.setStyle("Windows")
        sys.exit(app.exec_())

if __name__ == '__main__':
    ISC_program()
