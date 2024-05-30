# PATH: src/service/formato_arbol_correcciones.py

from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.comments import Comment

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
    print("Iniciando formateo del árbol de correcciones...")

    # Filtrar las filas que tienen valores en blanco o cadena vacía en "margen_ordenes"
    col_margen_ordenes = None
    for cell in arbol_ws[1]:  # Buscar columna "margen_ordenes" en la primera fila
        if cell.value == "margen_ordenes":
            col_margen_ordenes = cell.column_letter
            break
    
    if col_margen_ordenes:
        print(f"Columna 'margen_ordenes' encontrada en {col_margen_ordenes}")
    else:
        print("Columna 'margen_ordenes' no encontrada")

    # Eliminación de filas con valores vacíos en "margen_ordenes"
    rows_to_delete = []
    if col_margen_ordenes:
        for row in arbol_ws.iter_rows(min_row=2, max_row=arbol_ws.max_row):
            cell_value = row[arbol_ws[col_margen_ordenes + '1'].column - 1].value
            if cell_value is None or cell_value == "":
                rows_to_delete.append(row[0].row)
        
        print(f"Filas a eliminar: {rows_to_delete}")
        for row in reversed(rows_to_delete):
            arbol_ws.delete_rows(row)

    # Ajustar y añadir tabla
    if arbol_ws.max_row > 1:
        table_range = f"A1:{arbol_ws.dimensions.split(':')[1]}"
        tab = Table(displayName="ArbolCorrecciones", ref=table_range)
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style
        arbol_ws.add_table(tab)
        print(f"Tabla añadida en el rango {table_range}")

    # Ajuste de ancho de columnas
    for col in arbol_ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
        arbol_ws.column_dimensions[column].width = max_length + 2
    print("Anchos de columna ajustados.")

    # Añadir comentarios a los encabezados y configurar alineaciones
    center_alignment = Alignment(horizontal="center", vertical="center")
    left_alignment = Alignment(horizontal="left", vertical="center")
    for cell in arbol_ws[1]:
        cell.comment = Comment(text="Encabezado de columna", author="System")
        cell.alignment = center_alignment

    for row in arbol_ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column_letter in ['A', 'B']:  # Asumiendo que A y B son 'pos_estructura' y 'Nivel explosión'
                cell.alignment = left_alignment
            else:
                cell.alignment = center_alignment
    print("Comentarios y alineaciones configuradas.")

    # Formato condicional
    if col_margen_ordenes and arbol_ws.max_row > 1:
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        arbol_ws.conditional_formatting.add(f"{col_margen_ordenes}2:{col_margen_ordenes}{arbol_ws.max_row}",
                                            CellIsRule(operator="lessThan", formula=["0"], fill=red_fill))
        arbol_ws.conditional_formatting.add(f"{col_margen_ordenes}2:{col_margen_ordenes}{arbol_ws.max_row}",
                                            CellIsRule(operator="greaterThan", formula=["0"], fill=yellow_fill))
        arbol_ws.conditional_formatting.add(f"{col_margen_ordenes}2:{col_margen_ordenes}{arbol_ws.max_row}",
                                            CellIsRule(operator="equal", formula=["0"], fill=green_fill))

    # Congelar paneles para que los encabezados siempre sean visibles
    arbol_ws.freeze_panes = arbol_ws['A2']