import serial
import serial.tools.list_ports
from time import sleep
from qt_material import apply_stylesheet


# =====================================================
#           serial communication
# =====================================================
# comment line 31-33 and uncomment line 30 for identifying your hw id
ports = serial.tools.list_ports.comports()
target_port = None
for port, desc, hwid in sorted(ports):
    # print("{}: {} [{}]".format(port, desc, hwid))
    if "PID=1A86:7523" in hwid:
        # print(port)
        target_port = port
# =====================================================

if not target_port:
    print("No COM Port Found! :(")
    quit()

ser = serial.Serial(target_port, 9600)

while True:

    try:
        for chn in range(1,11):
            cmd = 'READ{}?\r\n'.format(chn)
            print("Channel {}: ".format(chn),end="\t")
            ser.write(cmd.encode())     # cmd to read data
            ser.flush()
            # sleep(0.1)
            line = ser.read_until(b'\r').decode()
            print(line)

    except:
        print("exception!")
        quit()
