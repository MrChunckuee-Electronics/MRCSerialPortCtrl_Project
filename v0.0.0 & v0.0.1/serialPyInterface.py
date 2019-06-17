import serial
import time

serial_object = serial.Serial('COM4', 9600)
serial_object.flushInput()
print('Connected to: ' + serial_object.portstr)
time.sleep(1)

while True:
        getSerialValue = serial_object.readline()
        outputSerialValue = getSerialValue.decode('utf-8')
        outputSerialValue = outpuSerialValue[:-2] # Removes last two characters (\r\n)
        print ('Serial Data: %s' %outputSerialValue)
