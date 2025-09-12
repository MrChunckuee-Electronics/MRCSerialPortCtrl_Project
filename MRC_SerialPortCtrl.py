import time
import threading
import serial
import serial.tools.list_ports
import tkinter as tk
from tkinter import ttk, messagebox

class SerialPyInterface(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Variables de instancia
        self.serial_object = None
        self.stop_event = threading.Event()
        self.last_data_time = 0
        self.buffer = ''

        self.title("MCR SerialPortCtrl - Desconectado")
        self.geometry('700x420')
        self.create_widgets()

    def update_title(self, status):
        """Actualiza el título de la ventana con el estado de la conexión."""
        if status == "Conectado":
            self.title("MRC SerialPortCtrl - Conectado")
        else:
            self.title("MRC SerialPortCtrl - Desconectado")

    def create_widgets(self):
        # Frame de configuración
        frame_config = ttk.LabelFrame(self, text="Configuración de Conexión", padding="10")
        frame_config.pack(padx=10, pady=10, fill="x")

        # Controles de puerto y Baud Rate
        ttk.Label(frame_config, text="Puerto:").grid(row=0, column=0, padx=5, pady=5)
        self.port_combobox = ttk.Combobox(frame_config, width=10)
        self.port_combobox.grid(row=0, column=1, padx=5, pady=5)
        self.list_ports()

        ttk.Label(frame_config, text="Baud Rate:").grid(row=0, column=2, padx=5, pady=5)
        self.baud_combobox = ttk.Combobox(frame_config, width=10)
        self.baud_combobox['values'] = [
            "300", "600", "1200", "2400", "4800", "9600", 
            "14400", "19200", "28800", "38400", "57600", "115200"
        ]
        self.baud_combobox.set("115200")
        self.baud_combobox.grid(row=0, column=3, padx=5, pady=5)
        
        # Botón único de conexión/desconexión
        self.connect_button = ttk.Button(frame_config, text="Conectar", command=self.toggle_connection)
        self.connect_button.grid(row=0, column=4, padx=10, pady=5)

        # Frame para enviar datos
        frame_send = ttk.LabelFrame(self, text="Enviar Datos", padding="10")
        frame_send.pack(padx=10, pady=10, fill="x")

        ttk.Label(frame_send, text="Mensaje:").pack(side="left", padx=5)
        self.send_entry = ttk.Entry(frame_send, width=50, state=tk.DISABLED)
        self.send_entry.pack(side="left", padx=5, fill="x", expand=True)

        self.send_button = ttk.Button(frame_send, text="Enviar", command=self.send_data, state=tk.DISABLED)
        self.send_button.pack(side="left", padx=5)

        # Frame de visualización de datos
        frame_data = ttk.LabelFrame(self, text="Visualización de Datos", padding="10")
        frame_data.pack(padx=10, pady=10, fill="both", expand=True)

        frame_widgets = ttk.Frame(frame_data)
        frame_widgets.pack(fill="both", expand=True)

        # Widget para datos ASCII
        frame_ascii = ttk.LabelFrame(frame_widgets, text="Datos Recibidos (ASCII)", padding=5)
        frame_ascii.pack(side="left", fill="both", expand=True, padx=5)
        self.serial_text_ascii = tk.Text(frame_ascii, wrap='word', state='disabled', height=10, width=35)
        self.serial_text_ascii.pack(fill="both", expand=True)
        scrollbar_ascii = ttk.Scrollbar(frame_ascii, command=self.serial_text_ascii.yview)
        scrollbar_ascii.pack(side="right", fill="y")
        self.serial_text_ascii.config(yscrollcommand=scrollbar_ascii.set)

        # Widget para datos Hexadecimales
        frame_hex = ttk.LabelFrame(frame_widgets, text="Datos Recibidos (Hexadecimal)", padding=5)
        frame_hex.pack(side="left", fill="both", expand=True, padx=5)
        self.serial_text_hex = tk.Text(frame_hex, wrap='word', state='disabled', height=10, width=35)
        self.serial_text_hex.pack(fill="both", expand=True)
        scrollbar_hex = ttk.Scrollbar(frame_hex, command=self.serial_text_hex.yview)
        scrollbar_hex.pack(side="right", fill="y")
        self.serial_text_hex.config(yscrollcommand=scrollbar_hex.set)

    def list_ports(self):
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.port_combobox['values'] = port_list
        if port_list:
            self.port_combobox.set(port_list[0])

    def toggle_connection(self):
        if self.serial_object and self.serial_object.is_open:
            self.disconnect()
        else:
            self.connect()

    def connect(self):
        port = self.port_combobox.get()
        baud_rate = self.baud_combobox.get()

        if not port or not baud_rate:
            messagebox.showerror("Error de Conexión", "Por favor, seleccione un puerto y un Baud Rate.")
            return

        try:
            self.serial_object = serial.Serial(port, int(baud_rate), rtscts=False, dsrdtr=False)
            messagebox.showinfo("Conectado", f"Conectado a {self.serial_object.portstr}")
            
            if not self.serial_object.is_open:
                raise serial.SerialException("No se pudo abrir el puerto serial.")
            
            self.update_title("Conectado")
            self.connect_button.config(text="Desconectar")
            
            self.send_entry.config(state=tk.NORMAL)
            self.send_button.config(state=tk.NORMAL)

            self.last_data_time = time.time()
            self.stop_event.clear()
            self.thread = threading.Thread(target=self.get_data)
            self.thread.daemon = True
            self.thread.start()
        except serial.SerialException as e:
            messagebox.showerror("Error de Conexión", f"No se pudo abrir el puerto. Verifique la conexión. Error: {e}")
            self.update_title("Desconectado")
        except ValueError:
            messagebox.showerror("Error", "Baud Rate o Puerto inválido.")

    def send_data(self):
        if self.serial_object and self.serial_object.is_open:
            message = self.send_entry.get()
            if message:
                try:
                    self.serial_object.write((message + '\n').encode('utf-8'))
                    self.send_entry.delete(0, 'end')
                except serial.SerialException as e:
                    messagebox.showerror("Error de Envío", f"No se pudo enviar el mensaje. Error: {e}")

    def get_data(self):
        while not self.stop_event.is_set():
            if not self.serial_object or not self.serial_object.is_open:
                self.after(0, self.disconnect)
                self.after(0, lambda: messagebox.showwarning("Conexión Perdida", "El dispositivo se ha desconectado."))
                break
            
            if self.serial_object.in_waiting > 0:
                try:
                    data = self.serial_object.read(self.serial_object.in_waiting).decode('utf-8')
                    self.buffer += data
                    
                    while '\n' in self.buffer:
                        line, self.buffer = self.buffer.split('\n', 1)
                        line = line.strip('\r').strip()
                        
                        self.after(0, self.update_text_widgets, line)

                except (serial.SerialException, UnicodeDecodeError):
                    self.after(0, self.disconnect)
                    self.after(0, lambda: messagebox.showwarning("Conexión Perdida", "El dispositivo se ha desconectado."))
                    break
            
            time.sleep(0.01)

    def update_text_widgets(self, data):
        """Añade los datos a ambos widgets de texto con una marca de tiempo."""
        timestamp = time.strftime("[%H:%M:%S]", time.localtime())

        # Visualización ASCII
        self.serial_text_ascii.config(state='normal')
        self.serial_text_ascii.insert('end', f"{timestamp} {data}\n")
        self.serial_text_ascii.see('end')
        self.serial_text_ascii.config(state='disabled')
        
        # Visualización Hexadecimal
        hex_data = ' '.join([f'{ord(c):02X}' for c in data])
        self.serial_text_hex.config(state='normal')
        self.serial_text_hex.insert('end', f"{timestamp} {hex_data}\n")
        self.serial_text_hex.see('end')
        self.serial_text_hex.config(state='disabled')
    
    def disconnect(self):
        """Cierra la conexión serial y detiene el hilo."""
        if self.serial_object and self.serial_object.is_open:
            self.stop_event.set()
            self.serial_object.close()
            
            self.update_title("Desconectado")
            self.connect_button.config(text="Conectar")
            
            self.send_entry.config(state=tk.DISABLED)
            self.send_button.config(state=tk.DISABLED)

            self.serial_text_ascii.config(state='normal')
            self.serial_text_ascii.delete('1.0', 'end')
            self.serial_text_ascii.config(state='disabled')
            self.serial_text_hex.config(state='normal')
            self.serial_text_hex.delete('1.0', 'end')
            self.serial_text_hex.config(state='disabled')
            
            messagebox.showinfo("Desconectado", "La conexión ha sido cerrada.")

if __name__ == "__main__":
    app = SerialPyInterface()
    app.mainloop()
