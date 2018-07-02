from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QProgressBar, QLabel, QLineEdit
from PyQt5.uic import loadUi


class MyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        loadUi('user_interface.ui', self)
        self.lineEditOpenPassword.setEchoMode(QLineEdit.Password)
        self.lineEditEditPassword.setEchoMode(QLineEdit.Password)
        progressBar = QProgressBar(self)
        progressBar.setRange(0, 100)
        progressBar.setValue(25)
        progressText = QLabel('就绪')
        self.statusBar().addWidget(progressText)
        self.statusBar().addPermanentWidget(progressBar)
        self.pushButtonSelectFiles.clicked.connect(self.selectFiles)

    def selectFiles(self):

        files, ok1 = QFileDialog.getOpenFileNames(self,
                                                  "多文件选择",
                                                  "C:/",
                                                  "Excel Files (*.xls *.xlsx)")
        print(files, ok1)


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    myshow = MyWindow()
    myshow.show()
    sys.exit(app.exec_())