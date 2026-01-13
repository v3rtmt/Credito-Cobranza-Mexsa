from datetime import datetime
import os
import re
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from PIL import Image, ImageTk

from NormalizarAntiguedad import normalizar_antiguedad
from LeerBaseDatos import cargar_comentarios, cargar_info_cliente
from ExportarExcel import exportar_excel, generar_archivos_por_ruta


# ============================================================================
#                           APLICACIÓN PRINCIPAL
# ============================================================================
class newApp:

    # =========================================================================
    #                   INICIALIZACIÓN DE LA INTERFAZ
    # =========================================================================
    def __init__(wd, rt):
        wd.rt = rt

        # =========================
        # Rutas de archivos
        # =========================
        wd.file_path_Antiguedad = None
        wd.file_path_BaseDatos = None

        # =========================
        # Configuración de ventana
        # =========================
        wd.rt.iconphoto(True, ImageTk.PhotoImage(Image.open("icon_R.png")))
        wd.rt.title("Generar Antigüedad de Saldos")
        wd.rt.geometry("500x150")
        wd.rt.resizable(0, 0)

        # =========================
        # Estilo ttk
        # =========================
        wd.style = ttk.Style()
        wd.style.theme_use("xpnative")

        # =========================
        # Grid principal
        # =========================
        wd.rt.grid_columnconfigure(0, weight=1)
        wd.rt.grid_columnconfigure(1, weight=1)

        # =========================
        # Secciones de archivos
        # =========================
        wd.create_file_section(0, 0, "Antigüedad de Saldos")
        wd.create_file_section(0, 1, "Base de Datos")

        # =========================
        # Botón principal
        # =========================
        wd.button_generate = ttk.Button(
            rt,
            text="Generar Antigüedad",
            command=wd.generate,
            style="Main.TButton"
        )
        wd.button_generate.grid(row=4, column=0, columnspan=2, pady=20)

    # =========================================================================
    #                   CREACIÓN DE SECCIÓN DE ARCHIVO
    # =========================================================================
    def create_file_section(self, row, col, label_text):

        # Contenedor
        frame = ttk.Frame(self.rt)
        frame.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")

        # Título
        ttk.Label(
            frame,
            text=label_text,
            font=("Helvetica", 10, "bold"),
            style="Subtitle.TLabel"
        ).pack()

        # Tipos permitidos
        filetypes = [("Archivos de Excel", "*.xlsx")]

        # Botón seleccionar
        ttk.Button(
            frame,
            text="Seleccionar Archivo",
            command=lambda: self.arch_select(row, col, filetypes),
            width=30,
            style="Main.TButton"
        ).pack(expand=True, fill=tk.BOTH)

        # Mostrar archivo seleccionado
        frame_arch = ttk.Frame(self.rt)
        frame_arch.grid(row=row + 1, column=col, padx=5, pady=5, sticky="nsew")

        label_arch = ttk.Label(frame_arch, text="", wraplength=250)
        label_arch.grid(row=0, column=0, sticky="w")

        delete_icon = ImageTk.PhotoImage(
            Image.open("icon_D.png").resize((20, 20))
        )

        button_delete = ttk.Button(
            frame_arch,
            image=delete_icon,
            command=lambda: self.arch_delete(row, col),
            style="Delete.TButton"
        )
        button_delete.grid(row=0, column=1, padx=5)
        button_delete.grid_remove()

        # Guardar referencias
        self.rt.widget_refs = getattr(self.rt, 'widget_refs', {})
        self.rt.widget_refs[(row, col)] = {
            'label': label_arch,
            'button': button_delete,
            'delete_icon': delete_icon
        }

    # =========================================================================
    #                   SELECCIÓN DE ARCHIVO
    # =========================================================================
    def arch_select(wd, row, col, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        widgets = wd.rt.widget_refs[(row, col)]

        if not path:
            widgets['label'].config(text="")
            widgets['button'].grid_remove()
            return

        widgets['label'].config(text=os.path.basename(path))
        widgets['button'].grid()

        if col == 0:
            wd.file_path_Antiguedad = path
        else:
            wd.file_path_BaseDatos = path

    # =========================================================================
    #                   VALIDACIÓN DE ARCHIVOS
    # =========================================================================
    def check_required_files(self):
        if not self.file_path_Antiguedad:
            self.show_error_message("Suba la Antigüedad de Saldos.")
            return False
        if not self.file_path_BaseDatos:
            self.show_error_message("Suba la Base de Datos.")
            return False
        return True

    # =========================================================================
    #                   ELIMINAR ARCHIVO SELECCIONADO
    # =========================================================================
    def arch_delete(wd, row, col):
        widgets = wd.rt.widget_refs[(row, col)]
        widgets['label'].config(text="")
        widgets['button'].grid_remove()

    # =========================================================================
    #                   MENSAJE DE ERROR
    # =========================================================================
    def show_error_message(wd, msg):
        messagebox.showerror("Error", msg)

    # =========================================================================
    #                   PROCESO PRINCIPAL
    # =========================================================================
    def generate(wd):
        try:
            if not wd.check_required_files():
                return

            path_antiguedad = wd.file_path_Antiguedad
            path_datos = wd.file_path_BaseDatos

            # =========================
            # Generación de registros
            # =========================
            def generar_registros(path_antiguedad, path_datos):
                facturas = normalizar_antiguedad(path_antiguedad)
                comentarios = cargar_comentarios(path_datos)
                clientes = cargar_info_cliente(path_datos)

                DEPTO_POR_PREFIJO = {
                    'BG': 'Brigar',
                    'NL': 'Abacer',
                    'CIG': 'Marlboro',
                    'HO': 'Holanda',
                    'PMX': 'Piso',
                }

                registros = []

                for f in facturas:
                    match = re.search(r'CLIENTE:\s*(\d+)', str(f['cliente_raw']))
                    codigo = match.group(1) if match else ''

                    info_cliente = clientes.get(codigo, {})
                    departamento = DEPTO_POR_PREFIJO.get(f['prefijo'], '')
                    dep_info = info_cliente.get(departamento, {})

                    info_com = comentarios.get(f['factura'], {})
                    ruta = info_com.get('ruta') or dep_info.get('ruta', '')

                    registros.append({
                        'factura': f['factura'],
                        'departamento': departamento,
                        'id_cliente': codigo,
                        'cliente': re.sub(
                            r'^CLIENTE:\s*\d+\s*/\s*',
                            '',
                            f['cliente_raw']
                        ).strip().title(),
                        'negocio': info_cliente.get('negocio', ''),
                        'fecha': f['fecha'].strftime('%d/%m/%Y'),
                        'antiguedad': (datetime.now().date() - f['fecha'].date()).days,
                        'total': f['total'],
                        'tipo_pago': info_cliente.get('tipo_pago', ''),
                        'ruta': ruta,
                        'dia': dep_info.get('dia', ''),
                        'fisico': '',
                        'comentarios': info_com.get('comentario', '')
                    })

                return registros

            registros = generar_registros(path_antiguedad, path_datos)

            # =========================
            # Rutas de salida
            # =========================
            fecha_str = datetime.now().strftime('%d-%m-%Y')
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")

            carpeta = os.path.join(desktop, f"Antigüedad al {fecha_str}")
            os.makedirs(carpeta, exist_ok=True)

            archivo_general = os.path.join(
                carpeta,
                f"Antigüedad de Saldos al {fecha_str}.xlsx"
            )

            # =========================
            # Exportación
            # =========================
            exportar_excel(registros, path_datos, archivo_general)
            generar_archivos_por_ruta(archivo_general, carpeta, fecha_str)

        except Exception as e:
            print(f"Error: {e}")
            wd.show_error_message(
                "Hubo un error al generar la antigüedad. Verifique los archivos."
            )


# ============================================================================
#                               EJECUCIÓN
# ============================================================================
if __name__ == "__main__":
    rt = tk.Tk()
    app = newApp(rt)
    rt.mainloop()
