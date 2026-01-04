from openpyxl import load_workbook

def cargar_comentarios(path_datos):
    wb = load_workbook(path_datos, data_only=True)
    hoja = wb['Comentarios']

    comentarios = {}

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        factura = fila[0]
        comentario = fila[3] if len(fila) > 3 else None

        if factura:
            comentarios[str(factura).strip()] = comentario or ''

    return comentarios

def cargar_info_cliente(path_datos):
    wb = load_workbook(path_datos, data_only=True)
    hoja = wb['Datos']

    info = {}

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        codigo = fila[0]
        if not codigo:
            continue

        info[str(codigo).strip()] = {
            'negocio': fila[2] or '',
            'tipo_pago': fila[3] or '',

            'Abacer': {
                'ruta': fila[5] or '',
                'dia': fila[6] or ''
            },
            'Brigar': {
                'ruta': fila[7] or '',
                'dia': fila[8] or ''
            },
            'Marlboro': {
                'ruta': fila[9] or '',
                'dia': fila[10] or ''
            },
            'Holanda': {
                'ruta': fila[11] or '',
                'dia': fila[12] or ''
            },
            'Piso': {
                'ruta': fila[13] or '',
                'dia': fila[14] or ''
            }
        }

    return info