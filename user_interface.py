from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QProgressBar, QLabel, QLineEdit, QMessageBox
from PyQt5.uic import loadUi
from worker_thread import WorkerThread
import os


class MyWindow(QMainWindow):
    progressSignal = QtCore.pyqtSignal(int, name='Progress')

    def __init__(self):
        super(MyWindow, self).__init__()
        self.progressSignal.connect(self.updateProgressSlot)
        self.__workerThread = WorkerThread(self.progressSignal)
        self.initUI()

    def initUI(self):
        loadUi('user_interface.ui', self)
        self.lineEditOpenPassword.setEchoMode(QLineEdit.Password)
        self.lineEditEditPassword.setEchoMode(QLineEdit.Password)
        self.progressBar = QProgressBar(self)
        self.progressBar.setRange(0, 100)
        self.progressBar.setValue(0)
        self.progressText = QLabel('就绪')
        self.statusBar().addWidget(self.progressText)
        self.statusBar().addPermanentWidget(self.progressBar)
        self.pushButtonSelectFiles.clicked.connect(self.selectFiles)
        self.pushButtonStart.clicked.connect(self.startWork)
        self.pushButtonQuit.clicked.connect(self.quit)

    def selectFiles(self):
        self.progressText.setText("选择文件")
        files, ok = QFileDialog.getOpenFileNames(self, "文件选择", os.getcwd(), "Excel Files (*.xls *.xlsx)")
        self.__workerThread.setFiles(files)

    def startWork(self):
        # 保证线程相关参数已经就绪
        if not self.__workerThread.isReady():
            QMessageBox.warning(self, "文件错误", "您尚未选择文件！")
            return

        # 避免同一个线程start多次
        if self.__workerThread.isRunning():
            QMessageBox.warning(self, "正在运行", "处理正在进行中，请等待当前处理结束")
            return

        self.__workerThread.start()
        self.progressText.setText("正在处理...")

    def updateProgressSlot(self, progress):
        self.progressBar.setValue(progress)
        if progress >= 100:
            self.progressText.setText("处理完成")

    # 关闭前的清理工作
    def quit(self):
        self.__workerThread.exit()  # 退出子线程
        exit()


if __name__ == "__main__":

    import sys
    app = QApplication(sys.argv)
    myshow = MyWindow()
    myshow.show()
    sys.exit(app.exec_())