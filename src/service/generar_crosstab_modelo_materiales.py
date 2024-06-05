# PATH: src/service/generar_crosstab_modelo_materiales.py

import pandas as pd
import numpy as np

def cargar_datos_maestros(archivo, sheet_bom_ea, sheet_bom_eb, sheet_coois, sheet_stocks, sheet_fabricacion_real_ea, sheet_fabricacion_real_eb):
    bom_ea = pd.read_excel(archivo, sheet_name=sheet_bom_ea)
    bom_eb = pd.read_excel(archivo, sheet_name=sheet_bom_eb)
    fabricacion_real_ea = pd.read_excel(archivo, sheet_name=sheet_fabricacion_real_ea)
    fabricacion_real_eb = pd.read_excel(archivo, sheet_name=sheet_fabricacion_real_eb)
    return bom_ea, bom_eb,fabricacion_real_ea, fabricacion_real_eb


def transformar_coois(coois):
    
    # PARTE 1: TRADUCCIÓN DE COLUMNAS POR SI SE DESCARGAN EN CASTELLANO
    
    # Diccionario de traducción de columnas del español al inglés
    traducciones = {
        'Orden': 'Order',
        'Número de proyecto': 'Project Number',
        'Número material': 'Material',
        'Texto breve material': 'Material description',
        'Modelo': 'Model (Effectivity)',
        'Unidad': 'Z Unit',
        'Estado de sistema': 'System Status',
        'Elemento PEP': 'WBS Element',
        'Versión de lista de materiales': 'BOM Version',
        'Versión fabricación': 'Production Version',
        'Versión hoja ruta': 'Routing Version',
        'Fecha liberac.real': 'Release date (actual)',
        'Cantidad orden (GMEIN)': 'Order quantity (GMEIN)'
    }
    
    # Lista de columnas originales
    columnas_originales = coois.columns.tolist()
    
    # Traducir las columnas si están en español
    columnas_traducidas = [traducciones.get(col, col) for col in columnas_originales]
    
    # Asignar las nuevas columnas traducidas al DataFrame
    coois.columns = columnas_traducidas

    # PARTE 2: TRANSFORMACIONES

    # Eliminar las filas donde "System Status" contiene "TECO" o "CTEC"
    coois = coois[~coois['System Status'].str.contains('TECO|CTEC', na=False)]

    # Añadir una columna personalizada en coois
    coois['mod-mat'] = coois['Model (Effectivity)'] + '-' + coois['Material']
    
    # Agrega la columna 'mod_ud' combinando 'Model (Effectivity)' y 'Z Unit'
    coois['mod_ud'] = coois['Model (Effectivity)'] + coois['Z Unit'].apply(lambda x: f"{x:03d}")

    # Filtra los subconjuntos EA y EB, manejando valores NA correctamente
    coois_ea = coois[coois['Model (Effectivity)'].str.contains('EA', na=False)]
    coois_eb = coois[coois['Model (Effectivity)'].str.contains('EB', na=False)]
    
    return coois_ea, coois_eb


def transformar_stocks(df):
    
    # PARTE 1: TRADUCCIÓN DE COLUMNAS POR SI SE DESCARGAN EN CASTELLANO

    # Diccionario de traducción de columnas del español al inglés
    traducciones = {
        'Material': 'Material',
        'Stock de libre utilización': 'Unrestricted Stock',
        'Descripción de almacén': 'Description of Storage Location',
        'Elemento PEP': 'WBS Element'
    }
    
    # Lista de columnas originales
    columnas_originales = df.columns.tolist()
    
    # Traducir las columnas si están en español
    columnas_traducidas = [traducciones.get(col, col) for col in columnas_originales]
    
    # Asignar las nuevas columnas traducidas al DataFrame
    df.columns = columnas_traducidas

    # PARTE 2: TRANSFORMACIONES
    
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

