# PATH: src/service/generar_crosstab_modelo_materiales.py

import pandas as pd
import numpy as np

def cargar_datos(archivo, sheet_bom_ea, sheet_bom_eb, sheet_coois, sheet_stocks, sheet_fabricacion_real_ea, sheet_fabricacion_real_eb):
    bom_ea = pd.read_excel(archivo, sheet_name=sheet_bom_ea)
    bom_eb = pd.read_excel(archivo, sheet_name=sheet_bom_eb)
    coois = pd.read_excel(archivo, sheet_name=sheet_coois)
    stocks = pd.read_excel(archivo, sheet_name=sheet_stocks)
    fabricacion_real_ea = pd.read_excel(archivo, sheet_name=sheet_fabricacion_real_ea)
    fabricacion_real_eb = pd.read_excel(archivo, sheet_name=sheet_fabricacion_real_eb)
    return bom_ea, bom_eb, coois, stocks, fabricacion_real_ea, fabricacion_real_eb

def transformar_coois(coois):

    # Cambio de tipo de las columnas en df_python
    convert_dict = {
        'Order': 'Int64',
        'Project Number': 'str',
        'Material': 'str',
        'Material description': 'str',
        'Model (Effectivity)': 'str',
        'Z Unit': 'Int64',
        'System Status': 'str',
        'WBS Element': 'str',
        'BOM Version': 'Int64',
        'Production Version': 'Int64',
        'Routing Version': 'Int64',
        'Release date (actual)': 'datetime64[ns]',
        'Order quantity (GMEIN)': 'Int64'
    }

    coois = coois.astype(convert_dict)

    # Añadir una columna personalizada en coois
    coois['mod-mat'] = coois['Model (Effectivity)'] + '-' + coois['Material']
    
    # Agrega la columna 'mod_ud' combinando 'Model (Effectivity)' y 'Z Unit'
    coois['mod_ud'] = coois['Model (Effectivity)'] + coois['Z Unit'].apply(lambda x: f"{x:03d}")

    # Filtra los subconjuntos EA y EB, manejando valores NA correctamente
    coois_ea = coois[coois['Model (Effectivity)'].str.contains('EA', na=False)]
    coois_eb = coois[coois['Model (Effectivity)'].str.contains('EB', na=False)]
    
    return coois_ea, coois_eb

def transformar_stocks(df):
    # Cambiar el tipo de las columnas especificadas
    df = df.astype({
        'Material': 'str',
        'Unrestricted Stock': 'Int64',
        'Description of Storage Location': 'str',
        'WBS Element': 'str'
    })
    
    # Añadir la columna Proyecto
    df['Proyecto'] = df['WBS Element'].apply(lambda x: x.split('-')[1] if '-' in x else '')
    
    # Añadir la columna personalizada proj-mat
    df['proj-mat'] = df['Proyecto'] + '-' + df['Material']
    
    return df

def transformar_fabricacion_real(df):
    # Cambiar el tipo de las columnas especificadas
    df = df.astype({
        'Proyecto_sap': 'Int64',
        'Vértice': 'str',
        'modelo': 'str',
        'Fase': 'str',
        'Mod-Fas': 'str',
        'Tramos fabricados': 'str',
        'Tramos no fabricados': 'str',
        'cant_fabricados': 'Int64',
        'cant_no_fab': 'Int64',
        'Unidades fabricadas': 'str',
        'Unidades no fabricadas': 'str'
    })
    
    # Convertir 'nan' texto a np.nan
    df['modelo'] = df['modelo'].replace('nan', np.nan)
    
    # Eliminar filas con valores nulos en la columna 'modelo'
    df = df.dropna(subset=['modelo'])
    
    return df

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