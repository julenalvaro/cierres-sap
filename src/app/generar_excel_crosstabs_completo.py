import os
import traceback
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink

from src.config.config import obtener_configuracion
from src.service.generar_crosstab_modelo_materiales import cargar_datos, transformar_coois, transformar_stocks, transformar_fabricacion_real, generar_crosstab_modelo_materiales
from src.service.formato_crosstab import format_crosstabs, agregar_cantidad_bom_header, formato_indice, agregar_enlace_indice, agregar_enlace_indice_hoja, guardar_excel
from src.service.transformar_bom_a_arbol_correcciones import transformar_bom_a_arbol_correcciones

def agregar_enlace_arbol(ws, modelos):
    # Encontrar la columna que contiene el encabezado "Modelo"
    modelo_col = None
    for cell in ws[1]:
        if cell.value == "Modelo":
            modelo_col = cell.column_letter
            break

    if modelo_col is None:
        raise ValueError("No se encontró una columna con el encabezado 'Modelo'")

    # Agregar enlaces a las celdas en la columna "Modelo"
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cell = row[ws[modelo_col + '1'].col_idx - 1]  # Obtener la celda en la columna "Modelo" de la fila actual
        if cell.value in modelos:
            modelo = cell.value
            cell.hyperlink = Hyperlink(ref=f'{modelo_col}{cell.row}', target=f'#\'{modelo}\'!A1', tooltip=f'Ir a {modelo}')
            cell.font = Font(color='0000FF', underline='single')

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

                # Agregar enlaces en la hoja del árbol de correcciones
                agregar_enlace_arbol(arbol_ws, unique_modelos)

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
