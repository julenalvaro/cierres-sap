# PATH: src/service/generar_crosstab_modelo_materiales.py

import pandas as pd

def cargar_datos(archivo, sheet_bom_ea, sheet_bom_eb, sheet_coois, sheet_stocks, sheet_fabricacion_real_ea, sheet_fabricacion_real_eb):
    bom_ea = pd.read_excel(archivo, sheet_name=sheet_bom_ea)
    bom_eb = pd.read_excel(archivo, sheet_name=sheet_bom_eb)
    coois = pd.read_excel(archivo, sheet_name=sheet_coois)
    stocks = pd.read_excel(archivo, sheet_name=sheet_stocks)
    fabricacion_real_ea = pd.read_excel(archivo, sheet_name=sheet_fabricacion_real_ea)
    fabricacion_real_eb = pd.read_excel(archivo, sheet_name=sheet_fabricacion_real_eb)
    return bom_ea, bom_eb, coois, stocks, fabricacion_real_ea, fabricacion_real_eb

def preparar_datos_coois(coois):
    # Asegura que la columna 'Model (Effectivity)' se maneje como string
    coois['Model (Effectivity)'] = coois['Model (Effectivity)'].astype(str)
    coois['Material'] = coois['Material'].astype(str)
    
    # Agrega la columna 'mod_ud' combinando 'Model (Effectivity)' y 'Z Unit'
    coois['mod_ud'] = coois['Model (Effectivity)'] + coois['Z Unit'].apply(lambda x: f"{x:03d}")
    coois['mod-mat'] = coois['Model (Effectivity)'] + '-' + coois['Material']
    
    # Filtra los subconjuntos EA y EB, manejando valores NA correctamente
    coois_ea = coois[coois['Model (Effectivity)'].str.contains('EA', na=False)]
    coois_eb = coois[coois['Model (Effectivity)'].str.contains('EB', na=False)]
    
    return coois_ea, coois_eb

def preparar_datos_stocks(stocks):
    # Cambiar el tipo de las columnas
    stocks = stocks.astype({"Material": str, "Unrestricted Stock": int, 
                            "Description of Storage Location": str, "WBS Element": str})

    # Añadir la columna "Proyecto"
    stocks['Proyecto'] = stocks['WBS Element'].apply(lambda x: x.split('-')[1] if '-' in x else '')

    # Añadir la columna "proj-mat"
    stocks['proj-mat'] = stocks['Proyecto'] + '-' + stocks['Material']
    
    return stocks

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