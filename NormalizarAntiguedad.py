from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from openpyxl.styles import numbers
from datetime import datetime
import re

def normalizar_antiguedad(path_antiguedad):
    wb = load_workbook(path_antiguedad)
    hoja = wb.active

    for rango in list(hoja.merged_cells):
        min_col, min_row, max_col, max_row = range_boundaries(str(rango))
        if min_row <= 9:
            hoja.unmerge_cells(
                start_row=min_row,
                start_column=min_col,
                end_row=max_row,
                end_column=max_col
            )

    hoja.delete_rows(1, 11)

    def unmerge_hasta_col(col_limite):
        a_desunir = []
        for r in hoja.merged_cells.ranges:
            if r.min_col <= col_limite:
                a_desunir.append(r)

        for r in a_desunir:
            try:
                hoja.unmerge_cells(str(r))
            except:
                pass

    unmerge_hasta_col(3)
    unmerge_hasta_col(5)

    hoja.delete_cols(5)
    hoja.delete_cols(3)
    hoja.delete_cols(1)

    for fila in range(1, hoja.max_row + 1):
        if not hoja[f'B{fila}'].value:
            hoja[f'B{fila}'].value = hoja[f'B{fila-1}'].value

    for fila in range(hoja.max_row, 0, -1):
        fecha = hoja[f'D{fila}'].value
        if not re.match(r'^\d{2}/\d{2}/\d{4}$', str(fecha)):
            hoja.delete_rows(fila)

    hoja.insert_cols(5)

    for fila in range(1, hoja.max_row + 1):
        if hoja[f'F{fila}'].value in (None, ''):
            break

        total = 0
        for col in range(6, 14):
            celda = hoja.cell(row=fila, column=col)
            if isinstance(celda.value, (int, float)):
                total += celda.value

        hoja[f'E{fila}'].value = total
        hoja[f'E{fila}'].number_format = numbers.FORMAT_NUMBER

    hoja.delete_cols(6, 12)

    facturas = []

    for fila in range(1, hoja.max_row + 1):
        factura_raw = hoja[f'A{fila}'].value

        if not factura_raw:
            continue

        if not (str(factura_raw).startswith('FACI') or str(factura_raw).startswith('FACE')):
            continue

        facturas.append({
            'factura_raw': factura_raw,
            'factura': str(factura_raw).split('/')[-1],
            'prefijo': str(factura_raw).split('/')[1],
            'cliente_raw': hoja[f'B{fila}'].value,
            'fecha': datetime.strptime(hoja[f'C{fila}'].value, '%d/%m/%Y'),
            'total': hoja[f'E{fila}'].value
        })

    facturas.sort(key=lambda x: x['fecha'], reverse=True)

    return facturas