from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

import time
import traceback, sys

class WorkerSignals(QObject):
    '''
    Defines the signals available from a running worker thread.
    Supported signals are:
        finished
            No data
        error
            `tuple` (exctype, value, traceback.format_exc() )
        result
            `object` data returned from processing, anything
        progress
            `int` indicating % progress
    '''

    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)
    progress = pyqtSignal(int)

class Worker(QRunnable):

    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()

        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

        kwargs['progress_callback'] = self.signals.progress

    @pyqtSlot()
    def run(self):

        try:
            result = self.fn(*self.args, **self.kwargs)
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result) # Return the result of the processing
        finally:
            self.signals.finished.emit() # Done


class MainWindow(QMainWindow):

    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        
        layout = QVBoxLayout()

        self.bar = QProgressBar()
        self.button = QPushButton("START IT UP")
        self.button.pressed.connect(self.execute)

        layout.addWidget(self.bar)
        layout.addWidget(self.button)

        w = QWidget()
        w.setLayout(layout)
        self.setCentralWidget(w)

        self.show()
        self.threadpool = QThreadPool()
        print("Multithreading with maximum %d threads" % self.threadpool.maxThreadCount())

    # ==============================================================
    #    Thread functions
    # ==============================================================
    def demo_task(self, progress_callback):
        self.button.setEnabled(0)
        for i in range(1,101):
            progress_callback.emit(i)
            time.sleep(0.050)
        return "Done"

    def demo_progress(self, n):
        self.bar.setValue(n)

    def demo_result(self,s):
        print(s)

    def demo_finished(self):
        self.button.setEnabled(True)
    # ==============================================================

    def execute(self):
        worker = Worker(self.demo_task)
        worker.signals.progress.connect(self.demo_progress)
        worker.signals.result.connect(self.demo_result)
        worker.signals.finished.connect(self.demo_finished)
        
        # Execute
        self.threadpool.start(worker)


app = QApplication([])
window = MainWindow()
app.exec_()