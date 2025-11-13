import time
import threading
import serial
import serial.tools.list_ports
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd 
import re # Importar la librería de expresiones regulares

class SerialPyInterface(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Variables de instancia
        self.serial_object = None
        self.stop_event = threading.Event()
        self.last_data_time = 0
        self.buffer = ''
        
        # Nueva variable para almacenar el historial de datos
        self.data_history = [] 
        
        # Variables para la exportación y el CSV
        self.data_column_name = tk.StringVar(value="DatoRecibidoCompleto") 
        self.csv_column_names = tk.StringVar(value="IR,TempAmbiente,TempObjeto") # Nombres de las sub-columnas

        self.title("MCR SerialPortCtrl - Desconectado")
        self.geometry('750x500')
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
        frame_send.pack(padx=10, pady=5, fill="x")

        ttk.Label(frame_send, text="Mensaje:").pack(side="left", padx=5)
        self.send_entry = ttk.Entry(frame_send, width=50, state=tk.DISABLED)
        self.send_entry.pack(side="left", padx=5, fill="x", expand=True)

        self.send_button = ttk.Button(frame_send, text="Enviar", command=self.send_data, state=tk.DISABLED)
        self.send_button.pack(side="left", padx=5)

        # --- FRAME DE EXPORTACIÓN ---
        frame_export = ttk.LabelFrame(self, text="Exportar Datos a Excel (Separación por Coma)", padding="10")
        frame_export.pack(padx=10, pady=5, fill="x")
        
        # Fila 1: Nombre de la columna completa
        ttk.Label(frame_export, text="Columna Completa:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(frame_export, width=25, textvariable=self.data_column_name).grid(row=0, column=1, padx=5, pady=5, sticky='w')

        # Fila 2: Nombres de las columnas CSV
        ttk.Label(frame_export, text="Nombres Sub-Columnas (CSV):").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(frame_export, width=40, textvariable=self.csv_column_names).grid(row=1, column=1, padx=5, pady=5, sticky='w')

        # Botón de exportación
        export_button = ttk.Button(
            frame_export, 
            text="Exportar a Excel (.xlsx)", 
            command=self.export_to_excel
        )
        export_button.grid(row=0, column=2, rowspan=2, padx=15, pady=5, sticky='ns')
        # -----------------------------------

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
        """Lista los puertos seriales disponibles y actualiza el ComboBox."""
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.port_combobox['values'] = port_list
        if port_list:
            self.port_combobox.set(port_list[0])

    def toggle_connection(self):
        """Alterna entre conectar y desconectar."""
        if self.serial_object and self.serial_object.is_open:
            self.disconnect()
        else:
            self.connect()

    def connect(self):
        """Establece la conexión serial e inicia el hilo de lectura."""
        port = self.port_combobox.get()
        baud_rate = self.baud_combobox.get()

        if not port or not baud_rate:
            messagebox.showerror("Error de Conexión", "Por favor, seleccione un puerto y un Baud Rate.")
            return

        try:
            self.serial_object = serial.Serial(port, int(baud_rate), rtscts=False, dsrdtr=False)
            
            if not self.serial_object.is_open:
                raise serial.SerialException("No se pudo abrir el puerto serial.")
            
            messagebox.showinfo("Conectado", f"Conectado a {self.serial_object.portstr}")
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
        """Añade los datos a los widgets de texto con marca de tiempo y los almacena, aplicando el filtro numérico."""
        timestamp_str = time.strftime("[%H:%M:%S]", time.localtime())
        timestamp_raw = time.time()
        
        # --- Lógica de Separación y Filtrado Numérico CORREGIDA ---
        split_data_raw = data.split(',')
        numeric_values = []
        
        # Expresión regular para encontrar números (enteros o decimales)
        # Busca un patrón que coincida con números enteros o con punto decimal
        # Ejemplo: " IR: 486" -> ['486']
        # Ejemplo: " TempA: 24.17" -> ['24.17']
        number_pattern = r"[-+]?\d*\.?\d+"
        
        for item in split_data_raw:
            # 1. Buscar el número usando Regex
            match = re.search(number_pattern, item)
            
            if match:
                value_str = match.group(0)
                try:
                    # 2. Convertir a float
                    numeric_values.append(float(value_str)) 
                except ValueError:
                    # Esto solo ocurriría si el regex encuentra algo raro que float() no acepta, 
                    # pero es una buena práctica de seguridad.
                    numeric_values.append(None) 
            else:
                # 3. Si no se encuentra ningún número en el segmento (ej: solo dice 'ERROR')
                numeric_values.append(None) 
        
        # Crear la entrada de historial
        history_entry = {
            'Timestamp_Raw': timestamp_raw,
            'Timestamp': timestamp_str.strip('[]'),
            self.data_column_name.get(): data, # Guardamos el dato completo original
            'Data_Split_Numeric': numeric_values # Guardamos solo los valores numéricos (o None)
        }
        self.data_history.append(history_entry)
        # ------------------------------------------------

        # Visualización ASCII
        self.serial_text_ascii.config(state='normal')
        self.serial_text_ascii.insert('end', f"{timestamp_str} {data}\n")
        self.serial_text_ascii.see('end')
        self.serial_text_ascii.config(state='disabled')
        
        # Visualización Hexadecimal
        hex_data = ' '.join([f'{ord(c):02X}' for c in data])
        self.serial_text_hex.config(state='normal')
        self.serial_text_hex.insert('end', f"{timestamp_str} {hex_data}\n")
        self.serial_text_hex.see('end')
        self.serial_text_hex.config(state='disabled')
    
    def disconnect(self):
        """Cierra la conexión serial, detiene el hilo y limpia la UI."""
        if self.serial_object and self.serial_object.is_open:
            self.stop_event.set()
            self.serial_object.close()
            
            self.update_title("Desconectado")
            self.connect_button.config(text="Conectar")
            
            self.send_entry.config(state=tk.DISABLED)
            self.send_button.config(state=tk.DISABLED)

            # Limpiar historial de datos y widgets de visualización
            self.data_history = []
            self.serial_text_ascii.config(state='normal'); self.serial_text_ascii.delete('1.0', 'end'); self.serial_text_ascii.config(state='disabled')
            self.serial_text_hex.config(state='normal'); self.serial_text_hex.delete('1.0', 'end'); self.serial_text_hex.config(state='disabled')
            
            messagebox.showinfo("Desconectado", "La conexión ha sido cerrada.")

    def export_to_excel(self):
        """Exporta el historial de datos, expandiendo los campos numéricos separados por coma."""
        if not self.data_history:
            messagebox.showwarning("Exportar", "No hay datos recibidos para exportar.")
            return

        # Abrir diálogo para guardar el archivo
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            initialfile="datos_serial_" + time.strftime("%Y%m%d_%H%M") + ".xlsx"
        )

        if not file_path:
            return

        try:
            # 1. Crear el DataFrame inicial
            df = pd.DataFrame(self.data_history)
            
            # 2. Obtener los nombres de las columnas para los datos separados
            col_names_str = self.csv_column_names.get().strip()
            col_names = [name.strip() for name in col_names_str.split(',') if name.strip()]

            # 3. Expandir la columna 'Data_Split_Numeric' en múltiples columnas
            df_expanded = df['Data_Split_Numeric'].apply(pd.Series)
            
            # 4. Asignar los nombres de columna definidos por el usuario
            num_expanded_cols = len(df_expanded.columns)
            
            if col_names and len(col_names) >= num_expanded_cols:
                 # Asigna los nombres proporcionados
                 df_expanded.columns = col_names[:num_expanded_cols]
            elif num_expanded_cols > 0:
                 # Si no hay suficientes nombres, usa los índices por defecto (0, 1, 2, ...)
                 df_expanded = df_expanded.rename(columns=lambda x: f"Valor{x+1}")
            
            # 5. Concatenar el DataFrame original con el expandido, eliminando columnas temporales
            df_final = pd.concat([df.drop(columns=['Timestamp_Raw', 'Data_Split_Numeric']), df_expanded], axis=1)

            # 6. Exportar a Excel
            df_final.to_excel(file_path, index=False)
            
            messagebox.showinfo("Exportación Exitosa", f"Datos exportados correctamente a:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error de Exportación", f"Ocurrió un error al guardar el archivo: {e}")

if __name__ == "__main__":
    app = SerialPyInterface()
    app.mainloop()
