from PyQt5.QtCore import QThread


class WorkerThread(QThread):

    def __init__(self, signal):
        super(WorkerThread, self).__init__()
        self.__signal = signal
        self.__files = []

    def setFiles(self, files):
        self.__files = files

    # 线程是否已就绪
    def isReady(self):
        return len(self.__files) > 0  # files不能为空

    def run(self):
        self.__work()

    # 调用这个函数来更新UI上的进度至progress（progress取值范围应为0-100）
    def __updateProgress(self, progress):
        self.__signal.emit(progress)

    # TODO:具体功能在这里实现，下面的代码为示例
    def __work(self):
        for i in range(0, 101):
            self.msleep(100)
            self.__updateProgress(i)