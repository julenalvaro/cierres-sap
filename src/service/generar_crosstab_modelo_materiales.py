# PATH: src/service/generar_crosstab_modelo_materiales.py

import pandas as pd

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
    # print(f"Procesando modelo: {modelo}")
    bom_filtrado = bom[bom['Modelo'] == modelo]
    coois_filtrado = coois[coois['Model (Effectivity)'] == modelo]

    # # Imprimir valores únicos para verificar discrepancias (debug)
    # print("Valores únicos en BOM 'Nº componentes':", bom_filtrado['Nº componentes'].unique())
    # print("Valores únicos en COOIS 'Material':", coois_filtrado['Material'].unique())

    # Verificar coincidencia
    coois_filtrado = coois_filtrado[coois_filtrado['Material'].isin(bom_filtrado['Nº componentes'])]

    # if coois_filtrado.empty (debug):
    #     print("No hay coincidencias entre COOIS 'Material' y BOM 'Nº componentes'.")
    # # else:
    # #     print(f"Coincidencias encontradas: {len(coois_filtrado)} registros")

    crosstab = pd.crosstab(coois_filtrado['mod_ud'], coois_filtrado['Material'],
                           values=coois_filtrado['Order quantity (GMEIN)'], aggfunc='sum').fillna(0)
    

    return crosstab