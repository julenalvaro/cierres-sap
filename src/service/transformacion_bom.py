# PATH: src/service/transformacion_bom.py

import pandas as pd
import numpy as np
from datetime import datetime
import os
from openpyxl import load_workbook

def transformacion_bom(file_path):
    # Extraer el nombre del archivo sin la ruta y sin la extensión
    base_name = os.path.basename(file_path)
    nombre_modelo = base_name.split('_')[1].split('.')[0]
    nombre_componente = base_name.split('_')[0]

    # Cargar el archivo Excel
    df = pd.read_excel(file_path)

    # Crear y añadir la fila con Entrada en tabla = 0
    new_row = {
        'Versión fabricación': 1,
        'Nivel explosión': "0",
        'Nº componentes': nombre_componente,
        'Texto breve-objeto': np.nan,
        'Grupo de artículos': 'FES*P',
        'Almacén producción': np.nan,
        'Ctd.componente (UMB)': 1,
        'Nivel': 0,
        'Ruta (predecesor)': 0,
        'Entrada en tabla': 0
    }

    # Crear la columna 'Modelo'
    df.insert(0, 'Modelo', nombre_modelo)

    df = pd.concat([pd.DataFrame([new_row]), df], ignore_index=True)

    # Inicializar la columna 'pos_estructura' como tipo object
    df['pos_estructura'] = np.nan
    df['pos_estructura'] = df['pos_estructura'].astype('object')
    
    last_seen = {}

    for index, row in df.iterrows():
        nivel = int(row['Nivel'])
        if nivel not in last_seen:
            last_seen[nivel] = 'A'
        else:
            last_seen[nivel] = chr(ord(last_seen[nivel]) + 1)
        
        estructura = ''
        for i in range(nivel + 1):
            if i in last_seen:
                if estructura:
                    estructura += '-'
                estructura += str(i) + last_seen[i]
        df.at[index, 'pos_estructura'] = estructura
    
    # Filtrar las filas que no tienen 'Versión fabricación' como null
    df = df[df['Versión fabricación'].notna()]

    # # Guardar el resultado en un nuevo archivo
    # result_dir = os.path.join(os.path.dirname(file_path), 'results')
    # if not os.path.exists(result_dir):
    #     os.makedirs(result_dir)

    # timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    # output_file = f"{base_name}_mod_{timestamp}.xlsx"
    # output_path = os.path.join(result_dir, output_file)
    # df.to_excel(output_path, index=False)

    return df

# Ejemplo de cómo llamar a la función
# df_result = transformacion_bom('dir_base_entorno_pruebas/descarga listados BOM/ML5-FESAP_RO1010.EB.A.xlsx')

