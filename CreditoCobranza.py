from datetime import datetime
from tkinter import messagebox, ttk, filedialog
from PIL import Image, ImageTk
import tkinter as tk
import os, re 
from NormalizarAntiguedad import normalizar_antiguedad
from LeerBaseDatos import cargar_comentarios, cargar_info_cliente
from ExportarExcel import exportar_excel


# Crea aplicación
class newApp:
    
    #==========================================================================
    #                      GENERAR INTERFAZ DE APLICACIÓN                      
    #==========================================================================
    
    def __init__(wd, rt):
        wd.rt = rt

        # Ruta de los archivos AGL_001 y Plantilla Excel
        wd.file_path_Antiguedad  = None
        wd.file_path_BaseDatos = None

        # Icono, titulo y tamaño de la ventana
        wd.rt.iconphoto(True, ImageTk.PhotoImage(Image.open("icon_R.png")))
        wd.rt.title("Generar de Antigüedad de Saldos")
        wd.rt.geometry("500x150")
        wd.rt.resizable(0, 0)

        # Estilo ttk
        wd.style = ttk.Style()
        wd.style.theme_use("xpnative")	
        
        # Grid de seleccion de archivos
        wd.rt.grid_columnconfigure(0, weight=1)
        wd.rt.grid_columnconfigure(1, weight=1)

        # Crear secciones de archivos
        wd.create_file_section(0, 0, "Antigüedad de Saldos")
        wd.create_file_section(0, 1, "Base de Datos")

        # Boton para generar registro
        wd.button_generate = ttk.Button(rt, text="Generar Antigüedad", command=wd.generate, style="Main.TButton")
        wd.button_generate.grid(row=4, column=0, columnspan=2, pady=20)

    # Función para crear sección para subir archivos
    def create_file_section(self, row, col, label_text):
        
        # Creacion del grid para la seleccion de archivos
        grid_file_section = ttk.Frame(self.rt)
        grid_file_section.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")

        # Label de tipo de seleccion de archivo
        label_arch_type = ttk.Label(grid_file_section, text=label_text, font=("Helvetica", 10, "bold"), style="Subtitle.TLabel")
        label_arch_type.pack()

        # Tipos de archivos que recibe la interfaz (ambos aceptan Excel)
        filetypes = [("Archivos de Excel", "*.xlsx")]

        # Boton para seleccionar archivo
        button_arch_select = ttk.Button(grid_file_section, text="Seleccionar Archivo", command=lambda: self.arch_select(row, col, filetypes), width=30, style="Main.TButton")
        button_arch_select.pack(expand=True, fill=tk.BOTH)

        # Grid para nombre de archivo seleccionado
        frame_arch = ttk.Frame(self.rt)
        frame_arch.grid(row=row+1, column=col, padx=5, pady=5, sticky="nsew")

        # Nombre de archivo seleccionado
        label_arch_name = ttk.Label(frame_arch, text="", wraplength=250, style="Subtitle.TLabel")
        label_arch_name.grid(row=0, column=0, padx=0, pady=0, sticky="w")

        # Boton de eliminar archivo
        delete_icon = ImageTk.PhotoImage(Image.open("icon_D.png").resize((20, 20)))
        button_arch_delete = ttk.Button(frame_arch, image=delete_icon, command=lambda row=row, col=col: self.arch_delete(row, col), style="Delete.TButton")
        button_arch_delete.grid(row=0, column=1, padx=5, pady=0, sticky="w")
        button_arch_delete.grid_remove()

        # Muestra archivo cargado
        self.rt.widget_refs = getattr(self.rt, 'widget_refs', {})
        self.rt.widget_refs[(row, col)] = {'label': label_arch_name, 'button': button_arch_delete, 'delete_icon': delete_icon}
            
    # Función para seleccionar archivos
    def arch_select(wd, row, col, fltp):
        
        # Elementos que van a mostrar el archivo seleccionado
        fl = filedialog.askopenfilename(filetypes=fltp)
        label  = wd.rt.widget_refs[(row, col)]['label']
        button = wd.rt.widget_refs[(row, col)]['button']

        # Seleccion de archivo entre texto y excel
        if fl:
            label.config(text=os.path.basename(fl))
            button.grid()
            if col == 0:
                wd.file_path_Antiguedad = fl
            elif col == 1:
                wd.file_path_BaseDatos = fl
        else:
            label.config(text="")
            button.grid_remove()

    # Función para verificar extensiones de archivos necesarios
    def check_required_files(self):
        if not self.file_path_Antiguedad:
            self.show_error_message("Suba la Antigüedad de Saldos para generar la antigüedad.")
            return False
        if not self.file_path_BaseDatos:
            self.show_error_message("Suba la Base de Datos para generar la antigüedad")
            return False
        return True
    
    # Función para eliminar archivos seleccionados
    def arch_delete(wd, row, col):
        widgets = wd.rt.widget_refs[(row, col)]
        widgets['label'].config(text="")
        widgets['button'].grid_remove()
    
    # Función para generar ventana de error
    def show_error_message(wd, error_message):
        messagebox.showerror("Error al generar la antigüedad", error_message)

    #==============================================================================
    #                     GENERAR ANTIGÜEDAD DE SALDOS TRABAJADA                      
    #==============================================================================

    def generate(wd):
        try:
            # Verificar si los archivos requeridos están cargados
            if not wd.check_required_files(): return

            # Agregar archivos requeridos al proceso
            path_antiguedad = wd.file_path_Antiguedad
            path_datos      = wd.file_path_BaseDatos

            def generar_registros(path_antiguedad, path_datos):
                facturas = normalizar_antiguedad(path_antiguedad)
                comentarios = cargar_comentarios(path_datos)
                info_clientes = cargar_info_cliente(path_datos)

                registros = []
                
                DEPTO_POR_PREFIJO = {
                    'BG': 'Brigar',
                    'NL': 'Abacer',
                    'CIG': 'Marlboro',
                    'HO': 'Holanda',
                    'PMX': 'Piso',
                }

                for f in facturas:
                    # Extraer código de cliente
                    match = re.search(r'CLIENTE:\s*(\d+)', str(f['cliente_raw']))
                    codigo_cliente = match.group(1) if match else ''

                    # Datos del cliente
                    info = info_clientes.get(codigo_cliente, {})
                    dep_info = info.get(DEPTO_POR_PREFIJO.get(f['prefijo']), {})
            
                    registro = {
                        'factura': f['factura'],
                        'departamento': DEPTO_POR_PREFIJO.get(f['prefijo'], ''),
                        'id_cliente': codigo_cliente,
                        'cliente': re.sub(r'^CLIENTE:\s*\d+\s*/\s*', '', f['cliente_raw']).strip().lower().title(),
                        'negocio': info.get('negocio', ''),
                        'fecha': f['fecha'].strftime('%d/%m/%Y'),
                        'antiguedad': (datetime.now().date() - f['fecha'].date()).days,
                        'total': f['total'],
                        'tipo_pago': info.get('tipo_pago', ''),
                        'ruta': dep_info.get('ruta', ''),
                        'dia': dep_info.get('dia', ''),
                        'fisico': '',
                        'comentarios': comentarios.get(f['factura'], '')
                    }
                    
                    registros.append(registro)
                    
                return registros
            
            registros = generar_registros(path_antiguedad, path_datos)

            salida = os.path.join(
                os.path.expanduser("~"), "Desktop", 
                f"Antigüedad de Saldos al {datetime.now().strftime('%Y-%m-%d')}.xlsx"
            )

            exportar_excel(
                registros, path_datos, salida
            )

        except Exception as e:
            error_message = "Hubo un error al generar el archivo. Verifique que los archivos seleccionados sean los correctos."
            print(f"Error: {e}")
            wd.show_error_message(error_message)

if __name__ == "__main__":
    rt  = tk.Tk()
    app = newApp(rt)
    rt.mainloop()