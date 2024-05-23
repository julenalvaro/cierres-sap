import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime

def cargar_datos(archivo, sheet_bom, sheet_coois):
    """
    Carga los datos desde las pestañas especificadas del archivo Excel.
    """
    bom = pd.read_excel(archivo, sheet_name=sheet_bom)
    coois = pd.read_excel(archivo, sheet_name=sheet_coois)
    return bom, coois

def preparar_datos_coois(coois):
    """
    Añade una nueva columna 'mod_ud' que combina 'Model (Effectivity)' con 'Z Unit' formateado.
    """
    coois['mod_ud'] = coois['Model (Effectivity)'].astype(str) + coois['Z Unit'].apply(lambda x: f"{x:03d}")
    return coois


def generar_crosstab_modelo_materiales(bom, coois, modelo):
    print(f"Procesando modelo: {modelo}")
    bom_filtrado = bom[bom['Modelo'] == modelo]
    coois_filtrado = coois[coois['Model (Effectivity)'] == modelo]

    # # Imprimir valores únicos para verificar discrepancias (debug)
    # print("Valores únicos en BOM 'Nº componentes':", bom_filtrado['Nº componentes'].unique())
    # print("Valores únicos en COOIS 'Material':", coois_filtrado['Material'].unique())

    # Verificar coincidencia
    coois_filtrado = coois_filtrado[coois_filtrado['Material'].isin(bom_filtrado['Nº componentes'])]

    if coois_filtrado.empty:
        print("No hay coincidencias entre COOIS 'Material' y BOM 'Nº componentes'.")
    else:
        print(f"Coincidencias encontradas: {len(coois_filtrado)} registros")

    crosstab = pd.crosstab(coois_filtrado['mod_ud'], coois_filtrado['Material'],
                           values=coois_filtrado['Order quantity (GMEIN)'], aggfunc='sum').fillna(0)
    return crosstab


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

# Configuración de los parámetros del archivo
archivo_excel = r"C:\Users\18287\OneDrive - Construcciones y Auxiliar deFerrocarriles SA (CAF)\projects\25-mudanza-ordenes-baan\archivos\descargas\subset\datos descargados - 22_05_2024 - subset.xlsx"
sheet_bom = 'todo'
sheet_coois = 'descarga_coois'

# Ejecución del proceso para generar el archivo Excel
archivo_generado = generar_excel_crosstabs_acabados(archivo_excel, sheet_bom, sheet_coois)
print(f'Archivo generado: {archivo_generado}')
