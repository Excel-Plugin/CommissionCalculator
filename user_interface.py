from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QProgressBar, QLabel, QLineEdit
from PyQt5.uic import loadUi


class MyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
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
        self.pushButtonQuit.clicked.connect(self.quit)

    def selectFiles(self):

        files, ok1 = QFileDialog.getOpenFileNames(self,
                                                  "多文件选择",
                                                  "C:/",
                                                  "Excel Files (*.xls *.xlsx)")
        print(files, ok1)

    def quit(self):
        # 未来此处可能还有其他清理工作
        exit()


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    myshow = MyWindow()
    myshow.show()
    sys.exit(app.exec_())