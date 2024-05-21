import os
import pandas as pd
from datetime import datetime
from src.service.transformacion_bom import transformacion_bom
from src.config.config import obtener_configuracion

def generar_bom_modificado():
    config = obtener_configuracion()
    dir_bom = os.path.join(config.DIR_BASE, "descarga listados BOM-acabados")
    dir_resultados = os.path.join(dir_bom, "results")
    
    if not os.path.exists(dir_resultados):
        os.makedirs(dir_resultados)
    
    df_concat = pd.DataFrame()
    archivos_procesados = 0
    archivos_intentados = 0
    errores = 0
    indice = 1
    
    for archivo in os.listdir(dir_bom):
        if archivo.endswith('.xlsx') and not archivo.startswith('~$'):
            archivos_intentados += 1
            path_completo = os.path.join(dir_bom, archivo)
            try:
                print(f"Intentando procesar el archivo: {archivo}")
                df_temp = transformacion_bom(path_completo)
                df_temp = df_temp.sort_values(by="Entrada en tabla")
                
                # Asignar índices crecientes a df_temp
                df_temp.insert(0, 'index', range(indice, indice + len(df_temp)))
                indice += len(df_temp)
                
                df_concat = pd.concat([df_concat, df_temp], ignore_index=True)
                archivos_procesados += 1
                print(f"Archivo procesado con éxito: {archivo}")
            except Exception as e:
                errores += 1
                print(f"Error al procesar el archivo {archivo}: {e}")
    
    if not df_concat.empty:
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        output_file = f"bom_modificado_{timestamp}.xlsx"
        output_path = os.path.join(dir_resultados, output_file)

        try:
            # Crear un objeto ExcelWriter con la ruta de salida
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Guardar el df completo en la primera pestaña
                df_concat.to_excel(writer, index=False, sheet_name='con_caldereria')

                # Filtrar y guardar en la segunda pestaña
                filtro1 = ~df_concat['Grupo de artículos'].astype(str).str.contains('Z00170000|Z00440000', na=False) & \
                          ~df_concat['Grupo de artículos'].astype(str).str.startswith('31')
                df_sin_caldereria = df_concat[filtro1]
                df_sin_caldereria.to_excel(writer, index=False, sheet_name='sin_caldereria')

                # Filtrar y guardar en la tercera pestaña
                filtro2 = ~df_concat['Grupo de artículos'].astype(str).str.contains('Z00450000|Z00180000|Z00590000', na=False)
                df_sin_cald_sin_sub = df_concat[filtro1 & filtro2]
                df_sin_cald_sin_sub.to_excel(writer, index=False, sheet_name='sin_cald_sin_sub')

            # Verificar que el archivo se ha creado correctamente
            if os.path.exists(output_path):
                print(f"Proceso completado. Archivos intentados: {archivos_intentados}, procesados correctamente: {archivos_procesados}, errores: {errores}.")
                print(f"Resultado combinado guardado en: {output_path}")
            else:
                print("Error: El archivo Excel no se generó correctamente.")
        except Exception as e:
            print(f"Error al intentar guardar el archivo Excel: {e}")
    else:
        print("No se procesaron archivos o todos los archivos tuvieron errores.")
        print(f"Proceso completado. Archivos intentados: {archivos_intentados}, procesados correctamente: {archivos_procesados}, errores: {errores}.")

