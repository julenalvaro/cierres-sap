import os
import pandas as pd
from datetime import datetime
from src.service.transformacion_bom import transformacion_bom
from src.config.config import obtener_configuracion

def generar_bom_modificado():
    config = obtener_configuracion()
    dir_bom = os.path.join(config.DIR_BASE, "descarga listados BOM")
    dir_resultados = os.path.join(dir_bom, "results")  # Directorio para guardar resultados
    
    # Crear el directorio de resultados si no existe
    if not os.path.exists(dir_resultados):
        os.makedirs(dir_resultados)
    
    df_concat = pd.DataFrame()  # DataFrame vacío para concatenar los resultados
    archivos_procesados = 0
    archivos_intentados = 0
    errores = 0
    
    # Listar todos los archivos .xlsx en el directorio especificado
    for archivo in os.listdir(dir_bom):
        if archivo.endswith('.xlsx') and not archivo.startswith('~$'):  # Ignorar archivos temporales de Excel
            archivos_intentados += 1
            path_completo = os.path.join(dir_bom, archivo)
            try:
                print(f"Intentando procesar el archivo: {archivo}")
                df_temp = transformacion_bom(path_completo)
                df_concat = pd.concat([df_concat, df_temp], ignore_index=True)
                archivos_procesados += 1
                print(f"Archivo procesado con éxito: {archivo}")
            except Exception as e:
                errores += 1
                print(f"Error al procesar el archivo {archivo}: {e}")
    
    # Guardar el DataFrame concatenado en un nuevo archivo Excel
    if not df_concat.empty:
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        output_file = f"bom_modificado_{timestamp}.xlsx"
        output_path = os.path.join(dir_resultados, output_file)  # Usar el directorio de resultados
        df_concat.to_excel(output_path, index=False)
        print(f"Proceso completado. Archivos intentados: {archivos_intentados}, procesados correctamente: {archivos_procesados}, errores: {errores}.")
        print(f"Resultado combinado guardado en: {output_path}")
    else:
        print("No se procesaron archivos o todos los archivos tuvieron errores.")
        print(f"Proceso completado. Archivos intentados: {archivos_intentados}, procesados correctamente: {archivos_procesados}, errores: {errores}.")
