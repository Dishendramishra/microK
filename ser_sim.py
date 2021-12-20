import serial
from time import sleep

ser = serial.Serial("COM10",9600)

dialouges = ["Ab sab kuch upar wale ki haath mein hain",
            "Kya issi din ke liye tujhe pal pos ke paida kiya tha",
            "Main tujhe ek phooti kaudi nahin doonga",
            "Ek baar mujhe maa kehkar pukaro beta",
            "Hawaldaar, giraftaar kar lo isse",
            "Mein tumhare bache ki maa baaney wali hoon",
            "Hum kisi ko muh dikhane layek nahin rahe",
            "Maine injection laga diya hai.. Kuch hi der mein hosh aajayega",
            "Operation karna hoga.. Dus hazaar rupaiye lagenge.",
            "Ghar mein do do jawaan betiyaan hain"]

while True:
    for i in dialouges:
        ser.write(i.encode())
        ser.write(b"\r\n")
        sleep(1)