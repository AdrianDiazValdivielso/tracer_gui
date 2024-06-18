from tkinter import *
import os
import csv
from tkinter import messagebox
from PIL import Image, ImageTk
import serial.tools.list_ports
import sys
#import pandas as pd
import matplotlib.pyplot as plt
#import openpyxl
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog


# ============main============================

class Tracer_App:
    def __init__(self, root):
        self.root = root
        self.root.geometry("1024x768+0+0") # +0+0 establece la posición de la ventana en la pantalla
        self.root.title("Trazador curvas IV")
        self.root.attributes('-fullscreen', True)
        defaultbg = root.cget('bg')
        bg_color = "#badc57" #bg_color
        root.configure(bg="#caec67")
        # title = Label(self.root, text="Trazador de curvas I-V", font=('times new roman', 30, 'bold'), pady=2, bd=12,
        # bg="white", fg="Black", relief=GROOVE)
        # title.pack(fill=X)

        # ======================Conexión serial========================
        self.serial_connection=serial.Serial()
        self.connected = False
        # ==================Mediciones temperatura=====================
        self.Pt100_1 = 0
        self.T_front = StringVar()

        self.Pt100_2 = 0
        self.T_rear = StringVar()
        # ==================Mediciones irradiancia=====================
        self.I_NES1_F = ""
        self.I_front = StringVar()

        self.I_NES1_T = ""
        self.I_rear = StringVar()
        # =====================Otras variables========================
        self.voltmeter1_config = StringVar()
        self.capacitance = StringVar()
        self.max_voltage = StringVar()
        self.T45 = BooleanVar()
        self.T70 = BooleanVar()
        # ===================Mediciones curva I-V======================
        self.voltages = []
        self.currents = []
        self.Isc = StringVar()  # Corriente de cortocircuito
        self.Uoc = StringVar()  # Voltaje de circuito abierto
        self.Ipmax = StringVar()  # Corriente del punto de potencia máxima
        self.Upmax = StringVar()  # Voltaje del punto de potencia máxima
        self.Isc_0 = StringVar()  # Corriente de cortocircuito para condiciones normalizadas
        self.Uoc_0 = StringVar()  # Voltaje de circuito abierto para condiciones normalizadas
        self.Ipmax_0 = StringVar()  # Corriente del punto de potencia máxima para condiciones normalizadas
        self.Upmax_0 = StringVar()  # Voltaje del punto de potencia máxima para condiciones normalizadas
        self.Ppk = StringVar()  # Potencia máxima para condiciones normalizadas

        # ===================Otras variables (parte 2)======================
        self.desired_voltage = StringVar()
        self.busy = False

        # ===================Comportamiento al cerrar la aplicación======================
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # =============Conexión======================
        F0 = LabelFrame(self.root, text="Conexión", font=('times new roman', 16, 'bold'), bd=10, fg="Black",
                        bg=bg_color)
        F0.place(x=0, y=0, width=185, height=165)
        cnct_btn = Button(F0, width=10, text="Conectar", command=self.connect, bg=bg_color,
                           font=('times new roman', 15, 'bold'),
                           bd=4)
        cnct_btn.grid(row=0, column=0, padx=15, pady=(15,2))
        dcnct_btn = Button(F0, width=10, text="Desconectar", command=self.disconnect, bg=bg_color,
                           font=('times new roman', 15, 'bold'),
                           bd=4)
        dcnct_btn.grid(row=1, column=0, pady=(0,15), padx=5)

        # =============Medir Temperatura======================
        F1 = LabelFrame(self.root, text="Temperatura", font=('times new roman', 14, 'bold'), bd=10, fg="Black",
                        bg=bg_color)
        F1.place(x=190, y=0, width=375, height=165)

        tmsr_btn = Button(F1, text="Medir\nTemperatura", command=self.temperature_measurement, bg=bg_color,
                          font=('times new roman', 16, 'bold'), fg="black", bd=4, height=3)
        tmsr_btn.grid(row=0, column=1, padx=15, pady=15, sticky='WSN', rowspan=2)

        tfrnt_lbl = Label(F1, width=6, text="Frontal:", font=('times new roman', 14, 'bold'), bg=bg_color)
        tfrnt_lbl.grid(row=0, column=2, pady=5, padx=(0,2))
        trear_lbl = Label(F1, width=6, text="Trasera:", font=('times new roman', 14, 'bold'), bg=bg_color)
        trear_lbl.grid(row=1, column=2, pady=5, padx=(0,2))

        t1_lbl = Label(F1, width=8, textvariable=self.T_front, font='arial 15', bd=3, relief=GROOVE)
        t1_lbl.grid(row=0, column=3, pady=5, padx=0)
        t2_lbl = Label(F1, width=8, textvariable=self.T_rear, font='arial 15', bd=3, relief=GROOVE)
        t2_lbl.grid(row=1, column=3, pady=5, padx=0)

        # =============Medir Irradiancia======================
        F2 = LabelFrame(self.root, text="Irradiancia", font=('times new roman', 14, 'bold'), bd=10, fg="Black",
                        bg=bg_color)
        F2.place(x=565, y=0, width=380, height=165)
        imsr_btn = Button(F2, text="Medir\nIrradiancia", command=self.irradiance_measurement, bg=bg_color,
                          font=('times new roman', 16, 'bold'),
                          fg="black", bd=4, height=3)
        imsr_btn.grid(row=0, column=1, padx=15, pady=15, sticky='WSN', rowspan=2)

        ifrnt_lbl = Label(F2, width=6, text="Frontal:", font=('times new roman', 14, 'bold'), bg=bg_color)
        ifrnt_lbl.grid(row=0, column=2, pady=5, padx=(0,2))
        irear_lbl = Label(F2, width=6, text="Trasera:", font=('times new roman', 14, 'bold'), bg=bg_color)
        irear_lbl.grid(row=1, column=2, pady=5, padx=(0,2))

        i1_lbl = Label(F2, width=10, textvariable=self.I_front, font='arial 15', bd=3, relief=GROOVE)
        i1_lbl.grid(row=0, column=3, pady=5, padx=0)
        i2_lbl = Label(F2, width=10, textvariable=self.I_rear, font='arial 15', bd=3, relief=GROOVE)
        i2_lbl.grid(row=1, column=3, pady=5, padx=0)

        # =============Alerta de sobrecalentamiento======================
        F3 = LabelFrame(self.root, text="", font=('times new roman', 15, 'bold'), bd=10, fg="Black",
                        bg=bg_color)
        F3.place(x=950, y=0, width=74, height=165)
        t45_lbl = Label(F3, width=3, text="T45", font=('times new roman', 15, 'bold'), bg=bg_color)
        t45_lbl.grid(row=0, column=0, pady=(0,0), padx=5, sticky="W")
        t70_lbl = Label(F3, width=3, text="T70", font=('times new roman', 15, 'bold'), bg=bg_color)
        t70_lbl.grid(row=2, column=0, pady=(20,0), padx=5, sticky="W")
        self.t45_canvas = Canvas(F3, bg=bg_color, width=40, height=30, highlightthickness=0)
        self.t45_canvas.grid(row=1, column=0, pady=0, padx=10, sticky="W")
        self.t45 = self.t45_canvas.create_oval(2, 2, 28, 28, outline="black", fill="black")
        self.t70_canvas = Canvas(F3, bg=bg_color, width=40, height=30, highlightthickness=0)
        self.t70_canvas.grid(row=3, column=0, pady=0, padx=10, sticky="W")
        self.t70 = self.t70_canvas.create_oval(2, 2, 28, 28, outline="black", fill="black")

        # ===================Configuración actual======================
        F4 = LabelFrame(self.root, text="Configuración actual", font=('times new roman', 15, 'bold'), bd=10,
                        fg="Black",
                        bg=bg_color)
        F4.place(x=0, y=170, width=285, height=140)

        cphn_lbl = Label(F4, text="Voltaje máximo:", bg=bg_color, font=('times new roman', 14, 'bold'),
                         fg="black", bd=4)
        cphn_lbl.grid(row=1, column=0, padx=(5,5), pady=(15,0), sticky='W')
        cphn_lbl = Label(F4, width=8, textvariable=self.max_voltage, font='arial 15', bd=3, relief=GROOVE)
        cphn_lbl.grid(row=1, column=1, padx=0, pady=(15,0), sticky='W')

        cphn_lbl = Label(F4, text="Capacitancia\nactual:", bg=bg_color, font=('times new roman', 14, 'bold'),
                         fg="black", bd=4)
        cphn_lbl.grid(row=2, column=0, padx=(5, 5), pady=0, sticky='W')
        cphn_lbl = Label(F4, width=8, textvariable=self.capacitance, font='arial 15', bd=3, relief=GROOVE)
        cphn_lbl.grid(row=2, column=1, padx=0, pady=0, sticky='W')


        # ===================Actualizar configuración======================
        F5 = LabelFrame(self.root, text="Actualizar configuración", font=('times new roman', 15, 'bold'), bd=10,
                        fg="Black", bg=bg_color)
        F5.place(x=0, y=315, width=285, height=220)

        cphn_lbl = Button(F5, text="Medir voltaje", command=self.get_voltmeter1_config, bg=bg_color,
                          font=('times new roman', 14, 'bold'), fg="black", bd=4)
        cphn_lbl.grid(row=0, column=0, padx=10, pady=15, sticky='W')
        cphn_lbl = Label(F5, width=8, textvariable=self.voltmeter1_config, font='arial 15', bd=3, relief=GROOVE)
        cphn_lbl.grid(row=0, column=1, padx=0, pady=15, sticky='W')

        cphn_lbl = Label(F5, text="Configuración\nvoltaje máximo:", bg=bg_color, font=('times new roman', 14, 'bold'),
                         fg="black", bd=4)
        cphn_lbl.grid(row=1, column=0, padx=10, pady=0, sticky='W')
        cap_cfg_menu = OptionMenu(F5, self.desired_voltage, "Auto", "1500 V", "900 V", "450 V")
        cap_cfg_menu.config(width=6, font='arial 12', bd=0, relief=GROOVE)
        cap_cfg_menu.grid(row=1, column=1, padx=0, pady=(0,15), sticky='W')

        cphn_lbl = Button(F5, text="Aplicar cambios",
                          command=lambda: self.update_config(self.desired_voltage.get(), user=True), bg=bg_color,
                          font=('times new roman', 14, 'bold'), fg="black", bd=4)
        cphn_lbl.grid(row=2, column=0, padx=(10,0), pady=5, sticky='WE',columnspan=2)

        # =====================Curva I-V=========================

        F6 = LabelFrame(self.root, text="Curva I-V", font=('times new roman', 15, 'bold'), bd=10, fg="Black",
                        bg=bg_color)
        F6.place(x=290, y=170, width=734, height=445)

        trace_button = Button(F6, text="Trazar curva I-V", command=self.trace_curve, bg=bg_color,
                              font=('times new roman', 14, 'bold'), fg="black", bd=4)
        trace_button.grid(row=0, column=0, padx=(10, 15), pady=0, sticky='WE', columnspan=2)

        # Isc
        Isc_lbl = Label(F6, text="I", font=('times new roman', 16, 'bold'), bg=bg_color, fg="black")
        Isc_lbl.grid(row=1, column=0, padx=(10, 0), pady=(10,0), sticky='W')
        Isc_sub_lbl = Label(F6, text="sc", font=('times new roman', 10), bg=bg_color, fg="black")
        Isc_sub_lbl.grid(row=1, column=0, padx=(20, 0), pady=(8, 0), sticky='sw')
        Isc_txt = Label(F6, width=7, textvariable=self.Isc, font=('times new roman', 16, 'bold'), bd=3,
                        relief=GROOVE)
        Isc_txt.grid(row=1, column=1, padx=0, pady=(10,0), sticky='W')

        # Uoc
        Uoc_lbl = Label(F6, text="U", font=('times new roman', 16, 'bold'), bg=bg_color, fg="black")
        Uoc_lbl.grid(row=2, column=0, padx=(10, 0), pady=0, sticky='W')
        Uoc_sub_lbl = Label(F6, text="oc", font=('times new roman', 10), bg=bg_color, fg="black")
        Uoc_sub_lbl.grid(row=2, column=0, padx=(26, 0), pady=(8, 0), sticky='sw')
        Uoc_txt = Label(F6, width=7, textvariable=self.Uoc, font=('times new roman', 16, 'bold'), bd=3,
                        relief=GROOVE)
        Uoc_txt.grid(row=2, column=1, padx=0, pady=0, sticky='W')

        # Ipmax
        Ipmax_lbl = Label(F6, text="I", font=('times new roman', 16, 'bold'), bg=bg_color, fg="black")
        Ipmax_lbl.grid(row=3, column=0, padx=(10, 0), pady=0, sticky='W')
        Ipmax_sub_lbl = Label(F6, text="p, max", font=('times new roman', 10), bg=bg_color, fg="black")
        Ipmax_sub_lbl.grid(row=3, column=0, padx=(20, 0), pady=(8, 0), sticky='sw')
        Ipmax_txt = Label(F6, width=7, textvariable=self.Ipmax, font=('times new roman', 16, 'bold'), bd=3,
                          relief="groove")
        Ipmax_txt.grid(row=3, column=1, padx=0, pady=0, sticky='W')

        # Upmax
        Upmax_lbl = Label(F6, text="U", font=('times new roman', 16, 'bold'), bg=bg_color, fg="black")
        Upmax_lbl.grid(row=4, column=0, padx=(10, 0), pady=0, sticky='W')
        Upmax_sub_lbl = Label(F6, text="p, max", font=('times new roman', 10), bg=bg_color, fg="black")
        Upmax_sub_lbl.grid(row=4, column=0, padx=(26, 0), pady=(8, 0), sticky='sw')
        Upmax_txt = Label(F6, width=7, textvariable=self.Upmax, font=('times new roman', 16, 'bold'), bd=3,
                          relief=GROOVE)
        Upmax_txt.grid(row=4, column=1, padx=0, pady=0, sticky='W')

        # Isc0
        Isc0_lbl = Label(F6, text="I", font=('times new roman', 16, 'bold'), bg=bg_color, fg="black")
        Isc0_lbl.grid(row=5, column=0, padx=(10, 0), pady=0, sticky='W')
        Isc0_sub_lbl = Label(F6, text="sc", font=('times new roman', 10), bg=bg_color, fg="black")
        Isc0_sub_lbl.grid(row=5, column=0, padx=(20, 0), pady=(8, 0), sticky='sw')
        Isc0_sup_lbl = Label(F6, text="0", font=('times new roman', 8), bg=bg_color, fg="black")  # Superscript label
        Isc0_sup_lbl.grid(row=5, column=0, padx=(20, 0), pady=(0, 8), sticky='nw')
        Isc0_txt = Label(F6, width=7, textvariable=self.Isc_0, font=('times new roman', 16, 'bold'), bd=3,
                         relief="groove")
        Isc0_txt.grid(row=5, column=1, padx=0, pady=0, sticky='W')

        # Uoc0
        Uoc0_lbl = Label(F6, text="U", font=('times new roman', 16, 'bold'), bg=bg_color, fg="black")
        Uoc0_lbl.grid(row=6, column=0, padx=(10, 0), pady=0, sticky='W')
        Uoc0_sub_lbl = Label(F6, text="oc", font=('times new roman', 10), bg=bg_color, fg="black")
        Uoc0_sub_lbl.grid(row=6, column=0, padx=(26, 0), pady=(8, 0), sticky='sw')
        Uoc0_sup_lbl = Label(F6, text="0", font=('times new roman', 8), bg=bg_color, fg="black")  # Superscript label
        Uoc0_sup_lbl.grid(row=6, column=0, padx=(26, 0), pady=(0, 8), sticky='nw')
        Uoc0_txt = Label(F6, width=7, textvariable=self.Uoc_0, font=('times new roman', 16, 'bold'), bd=3,
                         relief="groove")
        Uoc0_txt.grid(row=6, column=1, padx=0, pady=0, sticky='W')

        # Ipmax0
        Ipmax0_lbl = Label(F6, text="I", font=('times new roman', 16, 'bold'), bg=bg_color, fg="black")
        Ipmax0_lbl.grid(row=7, column=0, padx=(10, 0), pady=0, sticky='W')
        Ipmax0_sub_lbl = Label(F6, text="p, max", font=('times new roman', 10), bg=bg_color, fg="black")
        Ipmax0_sub_lbl.grid(row=7, column=0, padx=(20, 0), pady=(8, 0), sticky='sw')
        Ipmax0_sup_lbl = Label(F6, text="0", font=('times new roman', 8), bg=bg_color, fg="black")  # Superscript label
        Ipmax0_sup_lbl.grid(row=7, column=0, padx=(20, 0), pady=(0, 8), sticky='nw')
        Ipmax0_txt = Label(F6, width=7, textvariable=self.Ipmax_0, font=('times new roman', 16, 'bold'), bd=3,
                           relief="groove")
        Ipmax0_txt.grid(row=7, column=1, padx=0, pady=0, sticky='W')

        # Upmax0
        Upmax0_lbl = Label(F6, text="U", font=('times new roman', 16, 'bold'), bg=bg_color, fg="black")
        Upmax0_lbl.grid(row=8, column=0, padx=(10, 0), pady=0, sticky='W')
        Upmax0_sub_lbl = Label(F6, text="p, max", font=('times new roman', 10), bg=bg_color, fg="black")
        Upmax0_sub_lbl.grid(row=8, column=0, padx=(26, 0), pady=(8, 0), sticky='sw')
        Upmax0_sup_lbl = Label(F6, text="0", font=('times new roman', 8), bg=bg_color, fg="black")  # Superscript label
        Upmax0_sup_lbl.grid(row=8, column=0, padx=(26, 0), pady=(0, 8), sticky='nw')
        Upmax0_txt = Label(F6, width=7, textvariable=self.Upmax_0, font=('times new roman', 16, 'bold'), bd=3,
                           relief="groove")
        Upmax0_txt.grid(row=8, column=1, padx=0, pady=0, sticky='W')

        # Ppk
        Ppk_lbl = Label(F6, text="P", font=('times new roman', 16, 'bold'), bg=bg_color, fg="black")
        Ppk_lbl.grid(row=9, column=0, padx=(10, 0), pady=0, sticky='W')
        Ppk_sub_lbl = Label(F6, text="pk", font=('times new roman', 10), bg=bg_color, fg="black")
        Ppk_sub_lbl.grid(row=9, column=0, padx=(27, 0), pady=(8, 0), sticky='sw')
        Ppk_txt = Label(F6, width=7, textvariable=self.Ppk, font=('times new roman', 16, 'bold'), bd=3,
                        relief="groove")
        Ppk_txt.grid(row=9, column=1, padx=0, pady=0, sticky='W')

        save_btn = Button(F6, text="Guardar trazado", command=self.save_measurements, bg=bg_color,
                          font=('times new roman', 14, 'bold'), fg="black", bd=4)
        save_btn.grid(row=10, column=0, padx=10, pady=(10,0), sticky='WE', columnspan=2)

        # Gráfico de la curva
        fig, self.ax = plt.subplots(figsize=(5, 4))
        self.ax.set_xlabel('Voltaje (V)')
        self.ax.set_ylabel('Corriente (A)')
        self.ax.set_title('Curva I-V')
        self.ax.grid(False)
        self.canvas = FigureCanvasTkAgg(fig, master=F6)
        self.canvas.get_tk_widget().grid(row=0, column=2, rowspan=15, columnspan=15, padx=10, pady=0)

        # =======================Consola=============

        F7 = LabelFrame(self.root, text="Consola", font=('times new roman', 15, 'bold'), bd=10, fg="Black",
                        bg=bg_color)
        F7.place(x=0, y=535, width=290, height=233)
        self.console = Text(F7, wrap="word", state="disabled", font=('Courier', 8), height=13, width=38)
        self.console.grid(row=0,column=0, pady=5)

        # Redirect console output to the text widget
        import sys
        sys.stdout = self

        # # =====================Logotipos=========================
        F8 = Label(self.root, bg=bg_color)

        F8.columnconfigure(0, weight=1)
        F8.columnconfigure(1, weight=1)
        F8.columnconfigure(2, weight=1)

        F8.place(x=290, y=610, width=734, height=158)#(x=0, y=535, width=285, height=243)
        LogoUPV = Image.open(self.resource_path("LogoUPV_trimmed.jpg"))
        LogoUPV = LogoUPV.resize((170, 158))
        LogoUPV = ImageTk.PhotoImage(LogoUPV)
        image_label = Label(F8, image=LogoUPV) # Create a label to display the image
        image_label.image = LogoUPV  # Keep a reference to the image to prevent garbage collection
        image_label.grid(row=0,column=0)

        LogoTiM = Image.open(self.resource_path("LogoTIM.jpg"))
        LogoTiM = LogoTiM.resize((325, 158))
        LogoTiM = ImageTk.PhotoImage(LogoTiM)
        image_label = Label(F8, image=LogoTiM)
        image_label.image = LogoTiM  # Keep a reference to the image to prevent garbage collection
        image_label.grid(row=0,column=1)

        LogoEnertis = Image.open(self.resource_path("LogoEnertisApplus_trimmed.jpg"))
        LogoEnertis = LogoEnertis.resize((239, 158))
        LogoEnertis = ImageTk.PhotoImage(LogoEnertis)
        image_label = Label(F8, image=LogoEnertis)
        image_label.image = LogoEnertis  # Keep a reference to the image to prevent garbage collection
        image_label.grid(row=0,column=2)
    # ==========================Funciones internas==================================
    def flush(self):
        pass

    # Está función se ejecuta al cerrar la ventana
    def on_close(self):
        print("Cerrando aplicación...")
        root.destroy()
        sys.exit()

    def write(self, text):
        self.console.configure(state="normal")
        self.console.insert(END, text)
        self.console.configure(state="disabled")
        self.console.see(END)  # Scroll to the bottom

    def connect(self):
        ports = serial.tools.list_ports.comports()
        usb_serial_ports = [port for port, desc, hwid in ports if "USB Serial Port" in desc]
        if (len(usb_serial_ports) == 1) and self.connected==False:
            print(f">>>Conectando automaticamente a {usb_serial_ports[0]}.")
            self.serial_connection = serial.Serial(usb_serial_ports[0], baudrate=9600, timeout=0.1)

            self.serial_connection.write("hi  ".encode())
            # Leer respuesta de 16 bits
            response = self.serial_connection.read(2)
            if response == b'hi':
                self.connected = True
                print(">>>Conexión establecida con éxito.")
                self.update_config(new_maxvoltage="Auto", user=False)
            else:
                self.connected = False
                self.serial_connection.close()
                print(">>>No se ha recibido respuesta del trazador.")
        elif self.connected:
            print(">>>La conexión ya ha sido establecida")
        else:
            print(">>>Puerto serial no detectado.")
            self.connected = False

    def disconnect(self):
        # Verificamos si hay una conexión activa
        if self.serial_connection is not None and self.serial_connection.is_open:
            # Cerramos la conexión serial
            self.serial_connection.close()
            self.connected = False
            self.T45.set(False)
            self.T70.set(False)
            self.t45_canvas.itemconfig(self.t45, fill="black")
            self.t70_canvas.itemconfig(self.t70, fill="black")
            print(">>>Conexión con el trazador cerrada correctamente.")
        else:
            print(">>>No hay conexión activa para cerrar.")

    def temperature_measurement(self):
        if self.connected:
            # Enviar instrucción para leer sensor de temperatura Pt100_1 (frontal)
            self.serial_connection.write("t1  ".encode())

            # Leer respuesta de 16 bits
            response = self.serial_connection.read(2)
            if response:
                self.Pt100_1=round(int.from_bytes(response, byteorder="big") / 32 - 256, 1)
                if self.Pt100_1<150 and self.Pt100_1>-50:
                    self.T_front.set(str(self.Pt100_1) + "ºC")
                    print(">>>Temperatura frontal:", self.T_front.get())
                    #self.home_frame_label_1.configure(text=self.pt100_1)
                else:
                    self.T_front.set("")
                    print("Error Tf=",str(self.Pt100_1))
                    print(">>>Sensor de temperatura frontal no conectado")
            else:
                print(">>>Ha habido un problema al medir la temperatura. Inténtelo de nuevo.")

            # Enviar la instrucción para leer sensor de temperatura Pt100_2 (trasero)
            self.serial_connection.write("t7  ".encode())

            # Leer respuesta de 16 bits
            response = self.serial_connection.read(2)
            if response:
                self.Pt100_2 = round(int.from_bytes(response, byteorder="big") / 32 - 256, 1)
                if self.Pt100_2 < 150 and self.Pt100_2 > -50:
                    self.T_rear.set(str(self.Pt100_2) + "ºC")
                    print(">>>Temperatura trasera:", self.T_rear.get())
                else:
                    self.T_rear.set("")
                    print(">>>Sensor de temperatura trasero no conectado")
            else:
                print(">>>Ha habido un problema al medir la temperatura. Inténtelo de nuevo.")
        else:
            print(">>>Trazador I-V no conectado")

    def irradiance_measurement(self):
        if self.connected:
            # Enviar instrucción para leer sensor de irradiancia I_NES1_F (frontal)
            self.serial_connection.write("l1  ".encode())

            # Leer respuesta de 16 bits
            response = self.serial_connection.read(2)
            if response:
                self.I_NES1_F=round(int.from_bytes(response, byteorder="big")/4095*3.3/11/0.0951*1000, 1)
                if self.I_NES1_F<1850:
                    self.I_front.set(str(self.I_NES1_F) + " W/m\u00B2")
                    print(">>>Irradiancia frontal:", self.I_front.get())
                else:
                    self.I_front.set("")
                    print(">>>Sensor de irradiancia frontal no conectado")
            else:
                print(">>>Ha habido un problema al medir la irradiancia. Inténtelo de nuevo.")

            # Enviar la instrucción para leer sensor de irradiancia I_NES1_T (trasero)
            self.serial_connection.write("l2  ".encode())

            # Leer respuesta de 16 bits
            response = self.serial_connection.read(2)
            if response:
                self.I_NES1_T = round(int.from_bytes(response, byteorder="big")/4095*3.3/11/0.0951*1000, 1)
                if self.I_NES1_T < 1900:
                    self.I_rear.set(str(self.I_NES1_T) + " W/m\u00B2")
                    print(">>>Irradiancia trasera:", self.I_rear.get())
                else:
                    self.I_rear.set("")
                    print(">>>Sensor de irradiancia trasero no conectado")
            else:
                print(">>>Ha habido un problema al medir la temperatura. Inténtelo de nuevo.")
        else:
            print(">>>Trazador I-V no conectado")

    def get_voltmeter1_config(self):
        if self.connected:
            # Enviar instrucción para leer voltaje
            self.serial_connection.write("v1  ".encode())

            # Leer respuesta de 16 bits
            vm1_config = self.serial_connection.read(2)
            if vm1_config:
                self.voltmeter1_config.set("<"+str(int.from_bytes(vm1_config, byteorder="big"))+"V")
                print(">>>Lectura de voltaje:", self.voltmeter1_config.get())
            else:
                print(">>>Ha habido un problema al medir el voltaje. Inténtelo de nuevo.")
        else:
            print(">>>Trazador I-V no conectado")

    def update_config(self, new_maxvoltage, user):
        if self.connected:
            if new_maxvoltage=="":
                messagebox.showinfo("Advertencia", "No se ha seleccionado ninguna configuración de voltaje máximo.\n"
                                                   "Por favor, seleccione una configuración válida en el menú desplegable.")
                print(">>>Configuración para voltaje máximo no especificada.")
                return
            if user:    #Si el usuario quiere cambiar la configuración, primero se le pregunta para prevenir equivocaciones.
                yn = messagebox.askyesno("¡Alerta!", "¿Estás seguro de qué la configuración es correcta?\nLas tensiones a medir nunca deben superar el voltaje máximo seleccionado para el trazador.")
                if yn > 0:
                    self.max_voltage.set(new_maxvoltage)
                    if self.max_voltage.get() == "Auto":
                        self.serial_connection.write("c0  ".encode())
                        self.capacitance.set("2.5 mF")
                    elif self.max_voltage.get() == "1500 V":
                        self.serial_connection.write("c1  ".encode())
                        self.capacitance.set("2.5 mF")
                    elif self.max_voltage.get() == "900 V":
                        self.serial_connection.write("c2  ".encode())
                        self.capacitance.set("10 mF")
                    elif self.max_voltage.get() == "450 V":
                        self.serial_connection.write("c3  ".encode())
                        self.capacitance.set("40 mF")
                    else:
                        print(">>>No se ha podido actualizar la configuración.")
                    self.serial_connection.read(2)
                    print(">>>Capacitancia configurada a:", self.max_voltage.get())
                    messagebox.showinfo("Actualizado", "Configuración actualizada con éxito.\n"+
                                        "Voltaje máximo del trazador configurado a " + self.max_voltage.get()+"\n"+
                                        "Capacitancia del trazador configurada a " + self.capacitance.get())
                else:
                    return
            else:   #La GUI puede cambiar la configuración del voltaje máximo automáticamente sin preguntar al usuario
                self.max_voltage.set(new_maxvoltage)
                if self.max_voltage.get() == "Auto":
                    self.serial_connection.write("c0  ".encode())
                elif self.max_voltage.get() == "1500 V":
                    self.serial_connection.write("c1  ".encode())
                elif self.max_voltage.get() == "900 V":
                    self.serial_connection.write("c2  ".encode())
                elif self.max_voltage.get() == "450 V":
                    self.serial_connection.write("c3  ".encode())
                self.serial_connection.read(2)
                print(">>>Voltaje máximo del trazador configurado a:", self.max_voltage.get())
        else:
            print(">>>Trazador I-V no conectado")

    def trace_curve(self):
        if self.connected:
            # Enviar instrucción para realizar el trazado de curva I-V
            self.serial_connection.write("iv  ".encode())

            # Leer respuesta de 16 bits
            vm1_config = int.from_bytes(self.serial_connection.read(2), byteorder="big")
            if vm1_config:
                self.voltages = []
                self.currents = []

                # Borrar medición anterior
                self.ax.clear()
                self.voltmeter1_config.set("<" + str(vm1_config) + "V")
                print(">>>Lectura de voltaje:", self.voltmeter1_config.get())
                #print(f">>>Escala del voltímetro: 0V-{vm1_config}V")
                for i in range(101):
                    self.voltages.append(int.from_bytes(self.serial_connection.read(2), byteorder="big")/2**16*vm1_config)
                    self.currents.append(int.from_bytes(self.serial_connection.read(2), byteorder="big")/2**16*20)
                    #print(f'i={i}: voltages={self.voltages[i]}')
                self.ax.plot(self.voltages, self.currents)
                self.canvas.draw()
                self.Isc.set(str(round(self.currents[0],3))+" A")    # Corriente de cortocircuito
                self.Uoc.set(str(round(self.voltages[-1],3))+" V")  # Voltaje de circuito abierto
                self.Ipmax.set("0.995 A")  # Corriente del punto de potencia máxima
                self.Upmax.set("66.070 V")  # Voltaje del punto de potencia máxima
                self.Isc_0.set("1.424 A")  # Corriente de cortocircuito para condiciones normalizadas
                self.Uoc_0.set("88.215 V")  # Voltaje de circuito abierto para condiciones normalizadas
                self.Ipmax_0.set("1.292 A")  # Corriente del punto de potencia máxima para condiciones normalizadas
                self.Upmax_0.set("70.107 V")  # Voltaje del punto de potencia máxima para condiciones normalizadas
                self.Ppk.set("90.604 W")
            else:
                print(">>>Ha habido un problema en el trazado de curva I-V. Inténtelo de nuevo.")
        else:
            print(">>>Trazador I-V no conectado")

    def save_measurements(self):
        if len(self.voltages):
            data = [
                ["T_front", self.T_front.get()],
                ["T_rear", self.T_rear.get()],
                ["Irr_front", self.I_NES1_F],
                ["Irr_rear", self.I_NES1_T],
                ["Isc", self.Isc.get()],
                ["Uoc", self.Uoc.get()],
                ["Ipmax", self.Ipmax.get()],
                ["Upmax", self.Upmax.get()],
                ["Isc 0", self.Isc_0.get()],
                ["Uoc 0", self.Uoc_0.get()],
                ["Ipmax0", self.Ipmax_0.get()],
                ["Upmax0", self.Upmax_0.get()],
                ["Ppk", self.Ppk.get()],
                ["",""],
                ["U in V", "I in A"]
            ]
            for v, i in zip(self.voltages, self.currents):
                data.append([v, i])

            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Save CSV File"
            )

            # Write the data to the selected file
            with open(file_path, 'w', newline='') as file:
                writer = csv.writer(file, delimiter=';')
                writer.writerows(data)

            with open('solar_data.csv', 'w', newline='') as file:
                writer = csv.writer(file, delimiter=';')
                writer.writerows(data)
        else:
            print(">>>No hay datos para guardar")

    def check_overheat(self):
        self.root.after(600, self.check_overheat)
        if self.connected:
            self.serial_connection.write("ov  ".encode())
            # Leer respuesta de 8 bits
            overheat_indicators = int.from_bytes(self.serial_connection.read(1), byteorder="big")
            if overheat_indicators==1 or overheat_indicators==3:
                # Encender piloto de emergencia
                current_color = self.t45_canvas.itemcget(self.t45, "fill")
                if current_color == "black":
                    self.t45_canvas.itemconfig(self.t45, fill="red")
                elif current_color == "red":
                    self.t45_canvas.itemconfig(self.t45, fill="black")
                if self.T45.get()==False:
                    print(">>>La temperatura del trazador ha superado los 45ºC. Se iniciarán los ventiladores.")
                self.T45.set(True)
            else:
                self.T45.set(False)
                self.t45_canvas.itemconfig(self.t45, fill="black")
            if overheat_indicators>1 and self.T45.get():
                # Encender piloto de emergencia
                current_color = self.t70_canvas.itemcget(self.t70, "fill")
                if current_color == "black":
                    self.t70_canvas.itemconfig(self.t70, fill="red")
                elif current_color == "red":
                    self.t70_canvas.itemconfig(self.t70, fill="black")
                if self.T70.get() == False:
                    print(">>>La temperatura del trazador ha superado los 70ºC. Se procederá al apagado automático.")
                self.T70.set(True)
            else:
                self.T70.set(False)
                self.t70_canvas.itemconfig(self.t70, fill="black")

    # Función empleada durante la conversión del programa a .exe
    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

root = Tk()
obj = Tracer_App(root)
obj.check_overheat()
root.mainloop()
