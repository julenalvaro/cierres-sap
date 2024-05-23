# PATH: src/app/generar_excel_crosstabs_acabados.py

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font
import datetime

from src.service.generar_crosstab_modelo_materiales import (cargar_datos,
                                                            preparar_datos_coois,
                                                            generar_crosstab_modelo_materiales)

def cargar_umb(bom, modelo):
    """ Devuelve un diccionario de las cantidades UMB por componente para el modelo dado. """
    bom_filtrado = bom[bom['Modelo'] == modelo]
    return bom_filtrado.set_index('Nº componentes')['Ctd.componente (UMB)'].to_dict()

def agregar_cantidad_bom_header(ws, bom, modelo):
    """ Ajusta los encabezados de material y añade una fila para Cantidad_BOM. """
    umb_dict = cargar_umb(bom, modelo)
    ws.cell(row=1, column=1, value="Material").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=1, column=1).fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    ws.cell(row=2, column=1, value="Cantidad_BOM").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=2, column=1).fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    for col in range(2, ws.max_column + 1):
        material = ws.cell(row=1, column=col).value.split(' (')[0]  # Eliminar cantidades anteriores si están presentes
        cantidad_bom = umb_dict.get(material, 'N/A')
        ws.cell(row=1, column=col, value=material)
        ws.cell(row=2, column=col, value=cantidad_bom)
        ws.cell(row=1, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=1, column=col).fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        ws.cell(row=2, column=col).font = Font(bold=True)

def format_crosstabs(ws, bom, modelo):
    umb_dict = cargar_umb(bom, modelo)
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=2, max_col=ws.max_column):
        for cell in row:
            cell.value = 0 if cell.value is None else cell.value
            material = ws.cell(row=1, column=cell.column).value
            umb = umb_dict.get(material, 0)
            if cell.value < umb:
                cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            elif cell.value > umb:
                cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            elif cell.value == umb:
                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

def generar_excel_crosstabs_acabados(archivo, sheet_bom, sheet_coois):
    bom, coois = cargar_datos(archivo, sheet_bom, sheet_coois)
    coois = preparar_datos_coois(coois)

    wb = Workbook()
    wb.remove(wb.active)  # Elimina la pestaña predeterminada
    unique_modelos = bom['Modelo'].unique()

    for modelo in unique_modelos:
        ws = wb.create_sheet(title=modelo)
        crosstab = generar_crosstab_modelo_materiales(bom, coois, modelo)
        for r in dataframe_to_rows(crosstab, index=True, header=True):
            ws.append(r)
        agregar_cantidad_bom_header(ws, bom, modelo)
        format_crosstabs(ws, bom, modelo)

    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f'crosstabs_materiales_acabados_{timestamp}.xlsx'
    wb.save(filename)
    return filename