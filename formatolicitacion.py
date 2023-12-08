import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import re
class InterfazGrafica:
    def __init__(self, master):
        self.master = master
        master.title("Formato Licitaciones ChileCompra")

        # Fondo de la ventana
        master.configure(bg='#E3F2FD')

        # Establecer el estilo
        self.estilo = ttk.Style()
        self.estilo.configure('TButton', font=('calibri', 10, 'bold'), borderwidth='4', background='#2196F3', foreground='#FFFFFF')

        # Etiqueta de descripción
        self.label_descripcion = tk.Label(master, text="Programa que realiza el formato de los archivos de Chile compra.\n Favor de seleccionar las siguientes rutas: ", font=('calibri', 12), bg='#E3F2FD')
        self.label_descripcion.pack(pady=10)

        # Botón para seleccionar la ruta del archivo
        self.btn_ruta_archivo = tk.Button(master, text="Seleccionar Ruta del Archivo", command=self.seleccionar_ruta_archivo, width=30, bg='#1976D2', fg='#FFFFFF')
        self.btn_ruta_archivo.pack(pady=10)

        # Botón para seleccionar la ruta y nombre del archivo
        self.btn_ruta_nombre_archivo = tk.Button(master, text="Seleccionar Ruta y Nombre del Archivo", command=self.seleccionar_ruta_nombre_archivo, width=30, bg='#1565C0', fg='#FFFFFF')
        self.btn_ruta_nombre_archivo.pack(pady=10)

        # Botón para correr
        self.btn_correr = tk.Button(master, text="Correr", command=self.correr_accion, width=30, bg='#0D47A1', fg='#FFFFFF')
        self.btn_correr.pack(pady=10)

        self.ruta_archivo = ''
        self.ruta_nombre_archivo = ''

    def seleccionar_ruta_archivo(self):
        self.ruta_archivo = filedialog.askopenfilename()
        print("Ruta del Archivo Seleccionado:", self.ruta_archivo)

    def seleccionar_ruta_nombre_archivo(self):
        self.ruta_nombre_archivo = filedialog.asksaveasfilename()
        print("Ruta y Nombre del Archivo Seleccionado:", self.ruta_nombre_archivo)

    def correr_accion(self):
        if self.ruta_archivo != '' and self.ruta_nombre_archivo != '':
            print("Acción en ejecución")
            # Simula una tarea que lleva tiempo
            df = pd.read_csv(self.ruta_archivo,encoding='latin-1',sep=';')
            dfin = self.formato_archivo(df)
            if dfin.empty:
                self.mostrar_mensaje("No se pudo procesar el archivo")
            else:
                dfin.to_excel(self.ruta_nombre_archivo+'.xlsx')
                self.mostrar_mensaje("La ejecución ha terminado")
        else:
            self.mostrar_mensaje("Por favor selecciona la ruta del archivo y la ruta y nombre del archivo")

    def mostrar_mensaje(self, mensaje):
        messagebox.showinfo("Mensaje", mensaje)
    def formato_archivo(self,df):
        try:
            #+--------------------------------------------------------------------------------------------------------------
            #--------------------SELECCIONAR COLUMNAS-----------------------------------------------------------------------
            # Creating the boolean mask
            booleanMask = df.columns.isin(['Nro Licitaci?n P?blica', 'Id Convenio Marco', 'Convenio Marco', 'CodigoOC', 'NombreOC', 'Fecha Env?o OC', 'EstadoOC', 'Proviene de Gran Compra', 'IDProductoCM', 'Tipo de Producto', "Marca", 'Modelo', 'Cantidad', 'Rut Unidad de Compra', 'Unidad de Compra', 'Raz?n Social Comprador', 'Sector', 'Rut Proveedor', 'Nombre Proveedor Sucursal'])	
            # saving the selected columns 
            selectedCols = df.columns[booleanMask]
            # selecting the desired columns
            df2 = df[selectedCols]
            #--------------------------------------------------------------------------------------------------------------
            #--------------------FILTRAR POR ID----------------------------------------------------------------------------
            ID = 5802324
            DFIDFILTRADO = df2.loc[df2['Id Convenio Marco'] == ID]

            DFIDFILTRADO.head()
            #--------------------------------------------------------------------------------------------------------------
            # se agregan Marca de Producto,	Procesador,	Marca Procesador,	Familia Procesador,	Memoria RAM, Almacenamiento
            # 'Marca de Producto', 'Procesador', 'Marca Procesador', 'Familia Procesador', 'Memoria RAM', 'Almacenamiento'

            DFIDFILTRADO['Marca de Producto'] = df2['Marca']

            # Buscar el tipo de procesador en el enunciado
            enunciado = "M75Q WINDOWS 11 PRO AMD RYZEN 7 PRO 5750GE 16 GB RAM 512 SSD"
            patron_procesador_intel = re.compile(r'INTEL\s[^\d]+(?:\sCORE\s?[^\d]+)?\d+-\d+')
            patron_procesador_amd1 = re.compile(r'AMD[^\d]+RYZEN\s?\d+-?\d*[A-Z]+')
            patron_procesador_amd2 = re.compile(r'AMD\s+RYZEN\s+\d+\s+PRO\s+\d+[A-Z]+')
            patron_procesador_amd3 = re.compile(r'AMD\s+RYZEN\s+\d+\s?-?\d*[A-Z]+')
            def buscar_procesador(enunciado):
                if 'INTEL' in enunciado:
                    return (patron_procesador_intel.search(enunciado)).group()
                else:
                    if patron_procesador_amd1.search(enunciado):
                        return (patron_procesador_amd1.search(enunciado)).group()
                    else:
                        if patron_procesador_amd2.search(enunciado):
                            return (patron_procesador_amd2.search(enunciado)).group()
                        else:
                            return (patron_procesador_amd3.search(enunciado)).group()

            DFIDFILTRADO['Procesador'] = [buscar_procesador(enunciado) for enunciado in DFIDFILTRADO['Modelo']]
            DFIDFILTRADO['Marca Procesador'] = ['INTEL' if 'INTEL' in procesador else 'AMD' for procesador in DFIDFILTRADO['Procesador']]
            patron = re.compile(r'(RYZEN\s+\d+|CORE\s+I\d+)')
            DFIDFILTRADO['Familia Procesador'] = [patron.search(procesador).group() for procesador in DFIDFILTRADO['Procesador']]
            DFIDFILTRADO['Memoria RAM'] = [re.search(r'\d+\s?GB', enunciado).group() for enunciado in DFIDFILTRADO['Modelo']]
            DFIDFILTRADO['Almacenamiento'] = [re.search(r'\b(\d+)\s*SSD\b', enunciado).group() for enunciado in DFIDFILTRADO['Modelo']]

            #--------------------------------------------------------------------------------------------------------------
            #--------------------Reorganizar columnas------------------------------------------------------------------------

            dfin = DFIDFILTRADO[['Nro Licitaci?n P?blica', 'Id Convenio Marco', 'Convenio Marco', 'CodigoOC', 'NombreOC', 'Fecha Env?o OC', 'EstadoOC', 'Proviene de Gran Compra', 'IDProductoCM', 'Tipo de Producto', 'Marca de Producto', 'Procesador', 'Marca Procesador', 'Familia Procesador', 'Memoria RAM', 'Almacenamiento', 'Modelo', 'Cantidad', 'Rut Unidad de Compra', 'Unidad de Compra', 'Raz?n Social Comprador', 'Sector', 'Rut Proveedor', 'Nombre Proveedor Sucursal']]
            dfin = dfin.rename(columns={'Nro Licitaci?n P?blica': 'Nro Licitación Pública','Fecha Env?o OC':'Fecha Envío OC','Raz?n Social Comprador':'Razón Social Comprador'})
            dfin.head()
        except:
            dfin = pd.DataFrame()
        return dfin

if __name__ == "__main__":
    root = tk.Tk()
    interfaz = InterfazGrafica(root)
    root.mainloop()