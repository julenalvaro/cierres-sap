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

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

def format_arbol_correcciones(arbol_ws):
    # Formatear como tabla
    tab = Table(displayName="ArbolCorrecciones", ref=arbol_ws.dimensions)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                            showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    arbol_ws.add_table(tab)

    # Formato condicional para la columna "margen_ordenes"
    umb = 0
    col_margen_ordenes = None
    for cell in arbol_ws[1]:  # Buscar columna "margen_ordenes" en la primera fila
        if cell.value == "margen_ordenes":
            col_margen_ordenes = cell.column_letter
            break

    if col_margen_ordenes:
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

        arbol_ws.conditional_formatting.add(f"{col_margen_ordenes}2:{col_margen_ordenes}{arbol_ws.max_row}",
                                            CellIsRule(operator="lessThan", formula=[str(umb)], fill=red_fill))
        arbol_ws.conditional_formatting.add(f"{col_margen_ordenes}2:{col_margen_ordenes}{arbol_ws.max_row}",
                                            CellIsRule(operator="greaterThan", formula=[str(umb)], fill=yellow_fill))
        arbol_ws.conditional_formatting.add(f"{col_margen_ordenes}2:{col_margen_ordenes}{arbol_ws.max_row}",
                                            CellIsRule(operator="equal", formula=[str(umb)], fill=green_fill))
