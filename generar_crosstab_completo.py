# PATH: generar_crosstab_completo.py

from src.app.generar_excel_crosstabs_completo import generar_excel_crosstabs_completo

if __name__ == "__main__":
    archivo_excel = r"C:\Users\18287\OneDrive - Construcciones y Auxiliar deFerrocarriles SA (CAF)\projects\25-mudanza-ordenes-baan\archivos\descargas\descarga_producción\datos descargados - 05_06_2024 - complet - Copy.xlsx"
    args = {
        "archivo": archivo_excel,
        "sheet_bom_ea": 'bom-ea',
        "sheet_bom_eb": 'bom-eb',
        "sheet_coois": 'descarga_coois',
        "sheet_stocks": 'stocks',
        "sheet_fabricacion_real_ea": 'fabricacion_real_ea',
        "sheet_fabricacion_real_eb": 'fabricacion_real_eb'
    }

    archivo_generado_ea, archivo_generado_eb = generar_excel_crosstabs_completo(**args)


