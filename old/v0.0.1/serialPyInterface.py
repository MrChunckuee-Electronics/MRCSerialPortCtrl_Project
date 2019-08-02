"""
 *******************************************************************************
 *
 *         SerialPyInterface
 *
 *******************************************************************************
 * FileName:        PY_002.py
 * Complier:        EPython 3.7.1
 * Author:          Pedro Sánchez (MrChunckuee)
 * Blog:            http://mrchunckuee.blogspot.com/
 * Email:           mrchunckuee.psr@gmail.com
 * Description:     Python based GUI for the serial port
 *******************************************************************************
 * Rev.         Date            Comment
 *   v0.0.0     22/02/2019      - Creación del firmware
 *   v0.0.1     14/03/2019      - Se modifico para leer datos en una sola linea
 *                              - Se paso de byte a ASCII o UTF-8 y se elimino \r\n
 *******************************************************************************
"""

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
