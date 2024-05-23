# PATH: src/app/generar_excel_crosstabs_acabados.py

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime

from src.service.generar_crosstab_modelo_materiales import (cargar_datos,
                                                            preparar_datos_coois,
                                                            generar_crosstab_modelo_materiales)


def generar_excel_crosstabs_acabados(archivo, sheet_bom, sheet_coois):
    """
    Proceso completo para generar el archivo Excel con crosstabs por cada modelo.
    """
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
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f'crosstabs_materiales_acabados_{timestamp}.xlsx'
    wb.save(filename)
    return filename

