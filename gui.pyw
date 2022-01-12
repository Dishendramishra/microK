from PySide2 import QtGui
from PySide2 import QtCore
from PySide2.QtWidgets import *
from PySide2.QtCore import *
from PySide2.QtGui import *

from qtpy import uic
import pyqtgraph as pg
from qt_material import apply_stylesheet

import serial
from openpyxl import Workbook

import sys
from time import sleep
from collections import deque
from datetime import datetime


if sys.platform == "linux" or sys.platform == "linux2":
    pass

elif sys.platform == "win32":
    import ctypes
    myappid = u'prl.microk'  # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

elif sys.platform == "darwin":
    pass

class serialThread(QThread):
    progress_output_signal     = QtCore.Signal(list)
    graph_datapoint_signal     = QtCore.Signal(list)
    rw_channels_signal         = QtCore.Signal(bool)
    finished_signal            = QtCore.Signal(bool)

    def __init__(self):
        QThread.__init__(self)

        self.read_channels  = None
        self.write_channels = None
        self.ser_read       = None
        self.ser_write      = None
        self.serial_flag    = False

    def __del__(self):
        print("Thread terminating...")

    def kill(self):
        try:
            self.ser_read.close()
            self.ser_write.close()
        except serial.serialutil.PortNotOpenError:
            print("Already Closed!")
        self.serial_flag  = False

    def initiate(self, readport, writeport):
        try:
            self.ser_read  = serial.Serial(readport, 9600)
            self.ser_write = serial.Serial(writeport, 9600)

        except:
            print("Invalid port!")
            msg = QMessageBox()
            msg.setWindowTitle("Error")
            msg.setText("No Device Found!")
            msg.setIcon(QMessageBox.Critical)
            msg.exec_()
            self.finished.emit(True)
            return None

        self.serial_flag = True

        # don't call run() directly, instead call start()
        # start() will autonatically run()
        self.start()

    def run(self):

        x_ = 1

        while self.serial_flag:
            self.rw_channels_signal.emit(True)
            if not self.read_channels:
                continue

            print("write :",self.write_channels)
            print("read  :",self.read_channels)

            
            for channel_number in self.read_channels:
                try:
                    cmd = 'READ{}?\r\n'.format(channel_number)
                    self.ser_read.write(cmd.encode())
                    self.ser_read.flush()

                    output = self.ser_read.read_until(b'\r').decode().strip()
                    print("output: ",output)
                    
                    if "not enabled" in output:
                        self.progress_output_signal.emit([channel_number, "Not Enabled!"])

                    else:
                        self.progress_output_signal.emit([channel_number, output])
                        
                        val = round(float(output.split(",")[0]),4)
                        numb_val =  val
                        val = str(val)+"\n"
                        print("numb_val: ",numb_val)
                        
                        self.graph_datapoint_signal.emit([ x_ , numb_val])

                        if channel_number in self.write_channels:
                            self.ser_write.write(val.encode())
                            self.ser_write.flush()

                except Exception as e:
                            print(e)
                x_ += 1

class Ui(QMainWindow):
    
    def __init__(self, *args, **kwargs):
        super(Ui, self).__init__(*args, **kwargs)

        uic.loadUi('main_ui.ui', self)
        self.setWindowTitle("MickroK Temperature Reader")

        # ==========================================================
        #                       Graphing Window 
        # ==========================================================
        self.win = pg.GraphicsWindow()
        self.plot = self.win.addPlot(title='', labels={'bottom': 'time','left': "temperature"})
        self.curve = self.plot.plot()
        self.data = deque(maxlen=20)
        # ==========================================================


        # ==========================================================
        #                       Global Vars
        # ==========================================================
        self.val_lbls = {
            1   :   self.lbl_ch1_val,
            2   :   self.lbl_ch2_val,
            3   :   self.lbl_ch3_val,
            10  :   self.lbl_ch4_val,
            11  :   self.lbl_ch5_val,
            12  :   self.lbl_ch6_val,
            13  :   self.lbl_ch7_val,
            14  :   self.lbl_ch8_val,
            15  :   self.lbl_ch9_val,
            16  :   self.lbl_ch10_val
        }
        
        self.read_chbx_grp = {
            1   :   self.chbx_read1,
            2   :   self.chbx_read2,
            3   :   self.chbx_read3,
            10  :   self.chbx_read4,
            11  :   self.chbx_read5,
            12  :   self.chbx_read6,
            13  :   self.chbx_read7,
            14  :   self.chbx_read8,
            15  :   self.chbx_read9,
            16  :   self.chbx_read10,
        }

        self.write_chbx_grp = {
            1   :   self.chbx_write1,
            2   :   self.chbx_write2,
            3   :   self.chbx_write3,
            10  :   self.chbx_write4,
            11  :   self.chbx_write5,
            12  :   self.chbx_write6,
            13  :   self.chbx_write7,
            14  :   self.chbx_write8,
            15  :   self.chbx_write9,
            16  :   self.chbx_write10,
        }

        # ==========================================================
        #                 Thread Signals & Slots
        # ==========================================================
        self.serial_thread = serialThread()
        self.serial_thread.finished_signal.connect(self.stop)
        self.serial_thread.rw_channels_signal.connect(self.get_rw_channels)
        self.serial_thread.progress_output_signal.connect(self.update_channel_val)
        self.serial_thread.graph_datapoint_signal.connect(self.update_graph)
        
        # ==========================================================

        self.btn_stop.setEnabled(False)


        # ==========================================================
        #                        GUI Slots
        # ==========================================================
        self.btn_start.pressed.connect(self.start)
        self.btn_stop.pressed.connect(self.stop)
        self.btn_clear.pressed.connect(self.clear)
        # ==========================================================

        self.show()

    def closeEvent(self,event):
        self.serial_thread.kill()
        while self.serial_thread.isRunning():
            sleep(0.1)

        self.win.close()

    def get_rw_channels(self):

        channels_to_read = []
        channels_to_write = []

        for key, value in self.read_chbx_grp.items():
                if value.isChecked():
                    channels_to_read.append(key)


        for key, value in self.write_chbx_grp.items():
                if value.isChecked():
                    channels_to_write.append(key)

        self.serial_thread.read_channels  = channels_to_read
        self.serial_thread.write_channels = channels_to_write

    def update_channel_val(self, data_list):
        self.val_lbls[data_list[0]].setText(data_list[1])


    def update_graph(self, data_point):
        print("update_graph(): ", data_point)
        val_x, val_y = data_point
        self.data.append({"x" : val_x, "y" : val_y})
        x = [item['x'] for item in self.data]
        y = [item['y'] for item in self.data]
        self.curve.setData(x,y)
        self.sheet['A'+str(data_point[0])].value = val_x
        self.sheet['B'+str(data_point[0])].value = val_y

    def start(self):
        self.data.clear()
        self.curve.clear()
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_clear.setEnabled(False)
        read_port = self.cmb_readport.currentText()
        write_port = self.cmb_writeport.currentText()

        self.workbook = Workbook()
        self.sheet = self.workbook.active

        self.serial_thread.initiate(read_port, write_port)

    
    def stop(self):
        filename = datetime.now().strftime("%Y-%m-%d-%f")
        self.workbook.save(filename="{}.xlsx".format(filename))
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.btn_clear.setEnabled(True)
        self.serial_thread.kill()
    
    def clear(self):
        for ch_num in self.val_lbls:
            self.val_lbls[ch_num].setText("-")

app = QApplication(sys.argv)
app.setWindowIcon(QIcon("resources/icons/prl.png"))
apply_stylesheet(app, theme='dark_blue.xml')
window = Ui()
app.exec_()
