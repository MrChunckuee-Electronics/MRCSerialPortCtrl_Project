import time
import threading
import serial
import serial.tools.list_ports
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re 
import csv # Necesario para escribir en formato CSV

# Nota: El uso de 'pandas' se eliminó en esta versión, pero la importación 
# del sistema aún podría tenerla si no se elimina manualmente del archivo original.
# Para asegurar que sea liviano, se usa solo 'csv'.

class SerialPyInterface(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Variables de instancia
        self.serial_object = None
        self.stop_event = threading.Event()
        self.buffer = ''
        
        # --- Variables para CSV ---
        self.csv_file = None # Objeto de archivo para escritura
        self.csv_writer = None # Objeto para manejar la escritura CSV
        self.csv_file_path = None # Ruta del archivo seleccionado
        # -------------------------
        
        # Variables para la exportación y el CSV
        self.data_column_name = tk.StringVar(value="DatoRecibidoCompleto") 
        self.csv_column_names = tk.StringVar(value="BitsIR, VoltageIR,TempAmbiente,TempObjeto") # Nombres de las sub-columnas

        self.title("MCR SerialPortCtrl - Desconectado")
        self.geometry('750x500')
        self.create_widgets()

    def update_title(self, status):
        """Actualiza el título de la ventana con el estado de la conexión."""
        if status == "Conectado":
            self.title("MRC SerialPortCtrl - Conectado")
        else:
            self.title("MRC SerialPortCtrl - Desconectado")

    def list_ports(self):
        """Lista los puertos seriales disponibles y actualiza el ComboBox."""
        # CORRECCIÓN: Este método está correctamente definido dentro de la clase.
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.port_combobox['values'] = port_list
        if port_list:
            self.port_combobox.set(port_list[0])

    def create_widgets(self):
        # Frame de configuración
        frame_config = ttk.LabelFrame(self, text="Configuración de Conexión", padding="10")
        frame_config.pack(padx=10, pady=10, fill="x")

        # Controles de puerto y Baud Rate
        ttk.Label(frame_config, text="Puerto:").grid(row=0, column=0, padx=5, pady=5)
        self.port_combobox = ttk.Combobox(frame_config, width=10)
        self.port_combobox.grid(row=0, column=1, padx=5, pady=5)
        self.list_ports() # Llamada correcta usando self.

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
        frame_send.pack(padx=10, pady=5, fill="x")

        ttk.Label(frame_send, text="Mensaje:").pack(side="left", padx=5)
        self.send_entry = ttk.Entry(frame_send, width=50, state=tk.DISABLED)
        self.send_entry.pack(side="left", padx=5, fill="x", expand=True)

        self.send_button = ttk.Button(frame_send, text="Enviar", command=self.send_data, state=tk.DISABLED)
        self.send_button.pack(side="left", padx=5)

        # --- FRAME DE EXPORTACIÓN CSV ---
        frame_export = ttk.LabelFrame(self, text="Configuración de Archivo CSV", padding="10")
        frame_export.pack(padx=10, pady=5, fill="x")
        
        # Fila 1: Nombre de la columna completa
        ttk.Label(frame_export, text="Columna Completa:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(frame_export, width=25, textvariable=self.data_column_name).grid(row=0, column=1, padx=5, pady=5, sticky='w')

        # Fila 2: Nombres de las columnas CSV
        ttk.Label(frame_export, text="Nombres Sub-Columnas (CSV):").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(frame_export, width=40, textvariable=self.csv_column_names).grid(row=1, column=1, padx=5, pady=5, sticky='w')

        # Botón para seleccionar ubicación del archivo
        export_button = ttk.Button(
            frame_export, 
            text="Seleccionar Ubicación CSV", 
            command=self.select_csv_location
        )
        export_button.grid(row=0, column=2, rowspan=2, padx=15, pady=5, sticky='ns')
        # -----------------------------------

        # Frame de visualización de datos
        frame_data = ttk.LabelFrame(self, text="Visualización de Datos", padding="10")
        frame_data.pack(padx=10, pady=10, fill="both", expand=True)

        frame_widgets = ttk.Frame(frame_data)
        frame_widgets.pack(fill="both", expand=True)

        frame_ascii = ttk.LabelFrame(frame_widgets, text="Datos Recibidos (ASCII)", padding=5)
        frame_ascii.pack(side="left", fill="both", expand=True, padx=5)
        self.serial_text_ascii = tk.Text(frame_ascii, wrap='word', state='disabled', height=10, width=35)
        self.serial_text_ascii.pack(fill="both", expand=True)
        scrollbar_ascii = ttk.Scrollbar(frame_ascii, command=self.serial_text_ascii.yview)
        scrollbar_ascii.pack(side="right", fill="y")
        self.serial_text_ascii.config(yscrollcommand=scrollbar_ascii.set)

        frame_hex = ttk.LabelFrame(frame_widgets, text="Datos Recibidos (Hexadecimal)", padding=5)
        frame_hex.pack(side="left", fill="both", expand=True, padx=5)
        self.serial_text_hex = tk.Text(frame_hex, wrap='word', state='disabled', height=10, width=35)
        self.serial_text_hex.pack(fill="both", expand=True)
        scrollbar_hex = ttk.Scrollbar(frame_hex, command=self.serial_text_hex.yview)
        scrollbar_hex.pack(side="right", fill="y")
        self.serial_text_hex.config(yscrollcommand=scrollbar_hex.set)
        
    def toggle_connection(self):
        """Alterna entre conectar y desconectar."""
        if self.serial_object and self.serial_object.is_open:
            self.disconnect()
        else:
            self.connect()

    def select_csv_location(self):
        """Pide al usuario la ubicación y el nombre para el archivo CSV."""
        self.csv_file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Archivos CSV", "*.csv")],
            initialfile="datos_serial_" + time.strftime("%Y%m%d_%H%M") + ".csv"
        )
        if self.csv_file_path:
            messagebox.showinfo("Ubicación Seleccionada", f"Los datos se guardarán en:\n{self.csv_file_path}")
        else:
            messagebox.showwarning("Advertencia", "No se seleccionó ninguna ubicación. La conexión no guardará datos.")


    def connect(self):
        """Establece la conexión serial y abre el archivo CSV si se seleccionó la ruta."""
        port = self.port_combobox.get()
        baud_rate = self.baud_combobox.get()

        if not port or not baud_rate:
            messagebox.showerror("Error de Conexión", "Por favor, seleccione un puerto y un Baud Rate.")
            return
        
        # 1. Verificar y abrir el archivo CSV antes de conectar
        if self.csv_file_path:
            try:
                # Abrir el archivo en modo append ('a') si ya existe, sino lo crea.
                # 'newline=' es crucial para evitar líneas en blanco extra en Windows.
                self.csv_file = open(self.csv_file_path, mode='a', newline='', encoding='utf-8')
                self.csv_writer = csv.writer(self.csv_file)
                
                # Escribir encabezado solo si el archivo está vacío (posición 0)
                if self.csv_file.tell() == 0:
                    col_names_main = [self.data_column_name.get()]
                    # Obtener los nombres de sub-columnas del Entry
                    col_names_csv = [name.strip() for name in self.csv_column_names.get().split(',') if name.strip()]
                    header = ["Timestamp_Raw", "Timestamp"] + col_names_csv + col_names_main 
                    self.csv_writer.writerow(header)
                
            except Exception as e:
                messagebox.showerror("Error de Archivo", f"No se pudo abrir el archivo CSV para escritura: {e}")
                self.csv_file = None
                self.csv_writer = None
                return # Detener la conexión si falla la apertura del archivo

        # 2. Conexión Serial
        try:
            self.serial_object = serial.Serial(port, int(baud_rate), rtscts=False, dsrdtr=False)
            
            if not self.serial_object.is_open:
                raise serial.SerialException("No se pudo abrir el puerto serial.")
            
            messagebox.showinfo("Conectado", f"Conectado a {self.serial_object.portstr}")
            self.update_title("Conectado")
            self.connect_button.config(text="Desconectar")
            
            self.send_entry.config(state=tk.NORMAL)
            self.send_button.config(state=tk.NORMAL)

            self.stop_event.clear()
            self.thread = threading.Thread(target=self.get_data)
            self.thread.daemon = True
            self.thread.start()
        except serial.SerialException as e:
            messagebox.showerror("Error de Conexión", f"No se pudo abrir el puerto. Error: {e}")
            self.update_title("Desconectado")
        except ValueError:
            messagebox.showerror("Error", "Baud Rate o Puerto inválido.")

    def send_data(self):
        """Envía el mensaje escrito al puerto serial."""
        if self.serial_object and self.serial_object.is_open:
            message = self.send_entry.get()
            if message:
                try:
                    self.serial_object.write((message + '\n').encode('utf-8')) 
                    self.send_entry.delete(0, 'end')
                except serial.SerialException as e:
                    messagebox.showerror("Error de Envío", f"No se pudo enviar el mensaje. Error: {e}")

    def get_data(self):
        """Hilo de lectura: lee los datos del puerto serial."""
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
                    self.after(0, lambda: messagebox.showwarning("Conexión Perdida", "Error de lectura/decodificación."))
                    break
            
            time.sleep(0.01) 

    def update_text_widgets(self, data):
        """Añade los datos a los widgets de texto y los escribe en el archivo CSV."""
        timestamp_str = time.strftime("[%H:%M:%S]", time.localtime())
        timestamp_raw = time.time()
        
        # --- Lógica de Separación y Extracción Numérica (CORREGIDA) ---
        split_data_raw = data.split(',')
        numeric_values = []
        
        # Expresión regular para encontrar números (enteros o decimales)
        number_pattern = r"[-+]?\d*\.?\d+"
        
        for item in split_data_raw:
            match = re.search(number_pattern, item)
            
            if match:
                try:
                    # Se almacena el número encontrado como float
                    numeric_values.append(float(match.group(0))) 
                except ValueError:
                    # Si falla, se almacena una cadena vacía para CSV
                    numeric_values.append("") 
            else:
                # Si no se encuentra un número en el segmento, se almacena una cadena vacía
                numeric_values.append("") 
        
        # 1. Escritura Directa al CSV
        if self.csv_writer:
            # Crea la fila: [Timestamp_Raw, Timestamp] + [Subcolumnas Numéricas] + [Dato Completo]
            row_data = [
                timestamp_raw,
                timestamp_str.strip('[]')
            ] + numeric_values + [data]
            
            try:
                self.csv_writer.writerow(row_data)
                self.csv_file.flush() # Forzar la escritura al disco
            except Exception as e:
                # Mostrar un error menos intrusivo en la consola si la escritura falla
                print(f"Error escribiendo en CSV: {e}") 

        # 2. Visualización GUI
        self.serial_text_ascii.config(state='normal')
        self.serial_text_ascii.insert('end', f"{timestamp_str} {data}\n")
        self.serial_text_ascii.see('end')
        self.serial_text_ascii.config(state='disabled')
        
        hex_data = ' '.join([f'{ord(c):02X}' for c in data])
        self.serial_text_hex.config(state='normal')
        self.serial_text_hex.insert('end', f"{timestamp_str} {hex_data}\n")
        self.serial_text_hex.see('end')
        self.serial_text_hex.config(state='disabled')


    def disconnect(self):
        """Cierra la conexión serial, detiene el hilo y cierra el archivo CSV."""
        if self.serial_object and self.serial_object.is_open:
            self.stop_event.set()
            self.serial_object.close()
            
            self.update_title("Desconectado")
            self.connect_button.config(text="Conectar")
            
            self.send_entry.config(state=tk.DISABLED)
            self.send_button.config(state=tk.DISABLED)

            # --- Limpiar y cerrar CSV ---
            if self.csv_file:
                self.csv_file.close()
                self.csv_file = None
                self.csv_writer = None
            # ----------------------------

            # Limpiar widgets de visualización
            self.serial_text_ascii.config(state='normal'); self.serial_text_ascii.delete('1.0', 'end'); self.serial_text_ascii.config(state='disabled')
            self.serial_text_hex.config(state='normal'); self.serial_text_hex.delete('1.0', 'end'); self.serial_text_hex.config(state='disabled')
            
            messagebox.showinfo("Desconectado", "La conexión y la escritura de datos han sido cerradas.")
    
if __name__ == "__main__":
    app = SerialPyInterface()
    app.mainloop()
