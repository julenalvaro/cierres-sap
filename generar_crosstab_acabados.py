# PATH: generar_crosstab_acabados.py

from src.app.generar_excel_crosstabs_acabados import generar_excel_crosstabs_acabados

if __name__ == "__main__":
    # Configuración de los parámetros del archivo
    archivo_excel = r"C:\Users\18287\OneDrive - Construcciones y Auxiliar deFerrocarriles SA (CAF)\projects\25-mudanza-ordenes-baan\archivos\descargas\subset\datos descargados - 22_05_2024 - subset.xlsx"
    sheet_bom = 'todo'
    sheet_coois = 'descarga_coois'

    # Ejecución del proceso para generar el archivo Excel
    archivo_generado = generar_excel_crosstabs_acabados(archivo_excel, sheet_bom, sheet_coois)
    print(f'Archivo generado: {archivo_generado}')