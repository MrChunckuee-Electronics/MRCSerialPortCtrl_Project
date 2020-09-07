"""
 *******************************************************************************
 *
 *    Puerto serie - Recepción de datos
 *
 *******************************************************************************
 * FileName:        PY_001.py
 * Complier:        EPython 3.7.1
 * Author:          Pedro Sánchez (MrChunckuee)
 * Blog:            http://mrchunckuee.blogspot.com/
 * Email:           mrchunckuee.psr@gmail.com
 * Description:     Ejemplo de recepcion de datos atravez de un puerto COM
 *                  usando la libreria pySerial.
 *******************************************************************************
 * Rev.         Date            Comment
 *   v0.0.0     22/02/2019      Creación del firmware
 *******************************************************************************
"""

import serial

#Usamos puerto COM4 a 9600 baudios
serial_object = serial.Serial("COM4", 9600)
print("Connected to: " + serial_object.portstr)

value = 0

while True:
    value = serial_object.read(5);
    print (value)

