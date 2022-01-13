import serial
from time import sleep
import random

ser = serial.Serial("COM8",9600)

while True:
    output = ser.read_until(b'\r').decode().strip()
    print(output)
    ser.write(b"okay")
    ser.write(b"\r")
    sleep(1)