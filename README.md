# MRCSerialPortCtrl Project
A Desktop Application developed in Python for advanced serial port communication and monitoring. This tool is designed to interface with microcontrollers (PIC, Arduino, ESP32) and provide a robust environment for data visualization and logging.

## Features
* **Real-Time Monitoring:** High-speed data reception and transmission via Serial/UART.
* **Graphical User Interface (GUI):** Intuitive controls for port selection, baud rate configuration, and data flow.
* **Data Export:** Integrated functionality to save monitored data into **CSV** and **Excel** formats for post-processing.
* **Auto-Reconnect:** Smart handling of port disconnections and available hardware scanning.
* **Custom Commands:** Ability to send predefined command strings for hardware testing.

##  Requirements & Installation
This project requires Python 3.x and the following libraries:

1. Clone the repository:
   git clone https://github.com/MrChunckuee/MRCSerialPortCtrl_Project.git

2. Install dependencies:
   pip install pyserial pandas openpyxl

3. Run the application:
   python MRCSerialPortCtrl.py

##  Usage & GUI Overview
The application is structured to be plug-and-play:
1. **Connect:** Select the COM port and baud rate.
2. **Interact:** Use the terminal area to view incoming data or send commands.
3. **Log:** Use the export checkboxes to start recording data directly to an Excel spreadsheet.

## Documentation & Tutorial
For a detailed walkthrough on how to integrate the export logic with Python and how to handle serial interrupts effectively, visit my blog: https://mrchunckuee.blogspot.com/2019/06/MRCSerialPortCtrlProject.html
