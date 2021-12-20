from os import write
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.uic.uiparser import QtCore
import pyqtgraph as pg

import sys
from qt_material import apply_stylesheet
import traceback, sys
import serial
import random
from time import sleep
from collections import deque


if sys.platform == "linux" or sys.platform == "linux2":
    pass

elif sys.platform == "win32":
    import ctypes
    myappid = u'prl.microk'  # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

elif sys.platform == "darwin":
    pass

class WorkerSignals(QObject):
    error = pyqtSignal(tuple)
    progress = pyqtSignal(int)
    result = pyqtSignal(object)
    finished = pyqtSignal()
    
class Worker(QRunnable):

    def __init__(self, fn, signals_flag, *args, **kwargs):
        # signals_flag = [progress, result, finished]

        super(Worker, self).__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs    
        self.signals = WorkerSignals()
        self.signals_flag = signals_flag

        if self.signals_flag[0]:
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
            if self.signals_flag[1]:
                self.signals.result.emit(result) # Return the result of the processing
        finally:
            if self.signals_flag[2]:
                self.signals.finished.emit() # Done

# class serialThread(QThread):
#     progress_output = QtCore.Signal(int)

#     def __init__(self):
#         QThread.__init__(self)

#     def __del__(self):
#         print("Thread terminating...")

#     def run(self):
#         pass

class Ui(QMainWindow):
    
    def __init__(self, *args, **kwargs):
        super(Ui, self).__init__(*args, **kwargs)

        uic.loadUi('main_ui.ui', self)
        self.setWindowTitle("MickroK Temperature Reader")

        self.threadpool = QThreadPool()
        print("Multithreading with maximum %d threads" % self.threadpool.maxThreadCount())
        
        self.win = pg.GraphicsWindow()
        self.plot = self.win.addPlot(title='', labels={'bottom': 'time','left': "temperature"})
        self.curve = self.plot.plot()
        self.data = deque(maxlen=20)
        
        self.show()

        self.read_serial_flag = True

        self.val_lbls = {
            1   :   self.lbl_com1_val,
            2   :   self.lbl_com2_val,
            3   :   self.lbl_com3_val,
            10  :   self.lbl_com4_val,
            11  :   self.lbl_com5_val,
            12  :   self.lbl_com6_val,
            13  :   self.lbl_com7_val,
            14  :   self.lbl_com8_val,
            15  :   self.lbl_com9_val,
            16  :   self.lbl_com10_val
        }
        
        self.chbx_grp = {
            1   :   self.chbx_com1,
            2   :   self.chbx_com2,
            3   :   self.chbx_com3,
            10  :   self.chbx_com4,
            11  :   self.chbx_com5,
            12  :   self.chbx_com6,
            13  :   self.chbx_com7,
            14  :   self.chbx_com8,
            15  :   self.chbx_com9,
            16  :   self.chbx_com10,
        }

        self.btn_stop.setEnabled(False)
        self.chbx_com3.setChecked(True)

        self.btn_start.pressed.connect(self.start)
        self.btn_stop.pressed.connect(self.stop)
        self.btn_clear.pressed.connect(self.clear)

    def closeEvent(self,event):
        self.read_serial_flag = False
    
    # ==============================================================
    #    Thread functions
    # ==============================================================
    def read_serial_task(self, readport, writeport):
        
        x_ = 0
        while True:
            self.data.append({"x":x_, "y":random.randint(1,5)})
            x = [item['x'] for item in self.data]
            y = [item['y'] for item in self.data]
            self.curve.setData(x,y)
            sleep(0.1)
            x_ += 1

        # try:
        #     ser = serial.Serial(readport, 9600)
        #     ser_write = serial.Serial(writeport, 9600)

        #     channels_to_read = []

        #     for key, value in self.chbx_grp.items():
        #         if value.isChecked():
        #             channels_to_read.append(key)

        #     print("channels_to_read: ", channels_to_read)

        # except:
        #     print("Invalid port")
        #     msg = QMessageBox()
        #     msg.setWindowTitle("Error")
        #     msg.setText("No Device Found!")
        #     msg.setIcon(QMessageBox.Critical)
        #     msg.exec_()
        #     self.btn_start.setEnabled(True)
        #     self.btn_stop.setEnabled(False)
        #     self.btn_clear.setEnabled(True)
        #     return None

        # print("Listening on COM PORT: ",readport)
        # channel_number = 1

        # while self.read_serial_flag:
        #     for channel_number in self.chbx_grp.keys():
                
        #         if not self.read_serial_flag:
        #             break

        #         if channel_number in channels_to_read:
        #             try:
        #                 cmd = 'READ{}?\r\n'.format(channel_number)
        #                 print("\n",cmd.strip())
                        
        #                 ser.write(cmd.encode())
        #                 ser.flush()
        #                 line = ser.read_until(b'\r').decode().strip()

        #                 if self.read_serial_flag:
        #                     print("Channel number ",channel_number," : ",line)

        #                     if "not enabled" in line:
        #                         self.val_lbls[channel_number].setText("Not Enabled!") 
                            
        #                     else:
        #                         val = round(float(line.split(",")[0]),4)
        #                         val = str(val)+"\n"
                                
        #                         if channel_number == 3:
        #                             print("sending: ",val) 
        #                             ser_write.write(val.encode())
                                
        #                         self.val_lbls[channel_number].setText(line)            
                    
        #             except Exception as e:
        #                 print("exception: ", e)
            

    # For this task other functions are not required
    # ==============================================================
    
    def start(self):
        self.read_serial_flag = True
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_clear.setEnabled(False)
        read_port = self.cmb_readport.currentText()
        write_port = self.cmb_writeport.currentText()

        # Starting Threads
        worker = Worker(self.read_serial_task,[False,False,False],read_port, write_port)
        print("starting thread")
        self.threadpool.start(worker)

    
    def stop(self):
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.btn_clear.setEnabled(True)
        self.read_serial_flag = False
        print("terminating thread")
    
    def clear(self):
        for lbl in self.val_lbls:
            lbl.setText("-")

app = QApplication(sys.argv)
app.setWindowIcon(QIcon("resources/icons/prl.png"))
apply_stylesheet(app, theme='dark_blue.xml')
window = Ui()
app.exec_()
