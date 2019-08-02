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
 *   v0.0.2     31/07/2019      - Se realizo la GUI para conectarse a un puerto
 *******************************************************************************
"""

import time
import threading
import serial
import tkinter
from tkinter import *
from tkinter import ttk

serial_data = ''
filter_data = ''
update_period = 5
serial_object = None
gui = Tk()
gui.title("SerialPyInterface")

def connect():
  """
  La función inicia la conexión del dispositivo UART con los datos establecidos en
  la GUI:
      - BaudRate
      - Numero de puerto
      - Linux o Windows
  """

  global serial_object

  version_ = button_var.get()  
  port = port_entry.get()
  baud = baud_entry.get()

  try:
    if version_ == 2:
      print("Esta usando Linux")
      try:
        serial_object = serial.Serial('/dev/tty' + str(port), baud)
        print('Se ha conectado a ' + serial_object.portstr)
      except:
          print ("Cant Open Specified Port")
    elif version_ == 1:
      print("Esta usando Windows")
      serial_object = serial.Serial('COM' + str(port), baud)
      print('Se ha conectado a ' + serial_object.portstr)

  except ValueError:
    print ("Enter Baud and Port")
    return

  serial_object.flushInput()
  time.sleep(1)
  
  t1 = threading.Thread(target = get_data)
  t1.daemon = True
  t1.start()

def get_data():
  """
  Esta función sirve para recopilar datos del UART y almacenar los datos filtrados
  en una variable global.
  """

  global serial_object
  global filter_data

  while(1):
    try:
      serial_data = serial_object.readline()
      temp_data = serial_data.decode('utf-8')
      #filter_data = temp_data[:-2] # Removes last two characters (\r\n)
      filter_data = temp_data.strip('\n').strip('\r')
      print ('Serial Data: %s' %filter_data)
    except TypeError:
      pass

    
if __name__ == "__main__":
  """
  The main loop consists of all the GUI objects and its placement.
  The main loop handles all the widget placements.
  """

  #Frames
  frame_2 = Frame(height = 150, width = 480, bd = 3, relief = 'groove').place(x = 7, y = 5) 

  #Labels
  baud   = Label(text = "Baud").place(x = 100, y = 53)
  port   = Label(text = "Port").place(x = 200, y = 53)

  #Entry
  baud_entry = Entry(width = 7)
  baud_entry.place(x = 100, y = 70)

  port_entry = Entry(width = 7)
  port_entry.place(x = 200, y = 70)

  #Radio button
  button_var = IntVar()
  radio_1 = Radiobutton(text = "Windows", variable = button_var, value = 1).place(x = 10, y = 20)
  radio_2 = Radiobutton(text = "Linux", variable = button_var, value = 2).place(x = 110, y = 20)

  #Button
  connect = Button(text = "Connect", command = connect).place(x = 15, y = 65)
  
  #mainloop
  gui.geometry('500x200')
  gui.mainloop()
  
