import os
import traceback
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from src.config.config import obtener_configuracion
from src.service.generar_crosstab_modelo_materiales import cargar_datos, transformar_coois, transformar_stocks, transformar_fabricacion_real, generar_crosstab_modelo_materiales
from src.service.formato_crosstab import format_crosstabs, agregar_cantidad_bom_header, formato_indice, agregar_enlace_indice, agregar_enlace_indice_hoja, guardar_excel
from src.service.transformar_bom_a_arbol_correcciones import transformar_bom_a_arbol_correcciones

def generar_excel_crosstabs_completo(archivo, sheet_bom_ea, sheet_bom_eb, sheet_coois, sheet_stocks, sheet_fabricacion_real_ea, sheet_fabricacion_real_eb):
    config = obtener_configuracion()
    dir_crosstabs = os.path.join(config.DIR_BASE, "informes_crosstab")

    try:
        print('Cargando datos...')
        bom_ea, bom_eb, coois, stocks, fabricacion_real_ea, fabricacion_real_eb = cargar_datos(
            archivo, 
            sheet_bom_ea, 
            sheet_bom_eb, 
            sheet_coois, 
            sheet_stocks,
            sheet_fabricacion_real_ea,
            sheet_fabricacion_real_eb
        )

        print('Preparando datos...')

        # transformaciones primarias tablas, traducidas del código M 
        
        coois_ea, coois_eb = transformar_coois(coois)
        stocks = transformar_stocks(stocks)
        fabricacion_real_eb = transformar_fabricacion_real(fabricacion_real_eb)
        fabricacion_real_ea = transformar_fabricacion_real(fabricacion_real_ea)

        results = []

        for bom, coois, fabricacion_real, subset_name in [(bom_ea, coois_ea, fabricacion_real_ea, 'EA'), (bom_eb, coois_eb, fabricacion_real_eb, 'EB')]:
            try:
                print(f'Generando archivo Excel para {subset_name}...')
                wb = Workbook()
                wb.remove(wb.active)  # Elimina la pestaña predeterminada

                index_sheet = wb.create_sheet(title="Índice", index=1)
                index_sheet.append(["Modelo"])
                unique_modelos = sorted(bom['Modelo'].unique())

                # Generar crosstabs modelos
                for i, modelo in enumerate(unique_modelos, start=2):
                    ws = wb.create_sheet(title=modelo, index=i + 2)
                    crosstab = generar_crosstab_modelo_materiales(bom, coois, modelo)
                    for row in dataframe_to_rows(crosstab, index=True, header=True):
                        ws.append(row)
                    agregar_cantidad_bom_header(ws, bom, modelo)
                    format_crosstabs(ws, bom, modelo)

                    # Agregar enlace al índice y darle formato
                    agregar_enlace_indice(index_sheet, modelo, i + 2)
                    agregar_enlace_indice_hoja(ws)

                # sección del árbol de correcciones

                print('Transformando árbol...')

                arbol_correcciones = transformar_bom_a_arbol_correcciones(bom, coois, fabricacion_real, stocks)

                # Reemplazar NA con un placeholder antes de escribir en Excel
                arbol_correcciones = arbol_correcciones.astype('object')  # Convertir todas las columnas a tipo object
                arbol_correcciones.fillna('', inplace=True)

                # Aquí debes considerar cómo y dónde deseas guardar el DataFrame `arbol_correcciones`
                arbol_ws = wb.create_sheet(title=f'arbol_correcciones_{subset_name}', index=0)
                for row in dataframe_to_rows(arbol_correcciones, index=False, header=True):
                    arbol_ws.append(row)

                formato_indice(index_sheet)

                # Reordenar hojas: árbol de correcciones primero, luego índice, luego crosstabs
                wb._sheets.sort(key=lambda sheet: sheet.title not in [f'arbol_correcciones_{subset_name}', "Índice"])
                
                filename = guardar_excel(wb, dir_crosstabs, subset_name)
                results.append(filename)
                print(f'Archivo generado para {subset_name}: {filename}')
            except Exception as e:
                print(f'Ocurrió un error generando el archivo para {subset_name}: {e}')
                traceback.print_exc()  # Imprimir el rastreo completo del error
                results.append(None)

        return results
    except Exception as e:
        print(f'Ocurrió un error en la función principal: {e}')
        traceback.print_exc()  # Imprimir el rastreo completo del error
        return None, None
