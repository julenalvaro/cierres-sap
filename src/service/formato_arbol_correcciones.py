# PATH: src/service/formato_arbol_correcciones.py

from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink

def agregar_enlace_arbol(ws, modelos):
    # Encontrar la columna que contiene el encabezado "Modelo"
    modelo_col = None
    for cell in ws[1]:
        if cell.value == "Modelo":
            modelo_col = cell.column_letter
            break

    if modelo_col is None:
        raise ValueError("No se encontró una columna con el encabezado 'Modelo'")

    # Agregar enlaces a las celdas en la columna "Modelo"
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cell = row[ws[modelo_col + '1'].col_idx - 1]  # Obtener la celda en la columna "Modelo" de la fila actual
        if cell.value in modelos:
            modelo = cell.value
            cell.hyperlink = Hyperlink(ref=f'{modelo_col}{cell.row}', target=f'#\'{modelo}\'!A1', tooltip=f'Ir a {modelo}')
            cell.font = Font(color='0000FF', underline='single')