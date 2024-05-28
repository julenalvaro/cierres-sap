# PATH: generar_crosstab_acabados.py

from src.app.generar_excel_crosstabs_completo import generar_excel_crosstabs_completo

if __name__ == "__main__":
    # prod
    # archivo_excel = r"C:\Users\18287\OneDrive - Construcciones y Auxiliar deFerrocarriles SA (CAF)\projects\25-mudanza-ordenes-baan\archivos\descargas\subset\datos descargados - 24_05_2024 - complet.xlsx"
    # dev
    
    archivo_excel = r"C:\Users\18287\OneDrive - Construcciones y Auxiliar deFerrocarriles SA (CAF)\projects\25-mudanza-ordenes-baan\archivos\descargas\subset\datos descargados - 24_05_2024 - complet - subset.xlsx"
    # sheet_fabricacion_eb = 'fabricacion_real_eb'
    # sheet_fabricacion_ea = 'fabricacion_real_ea'

    args = {
        "archivo": archivo_excel,
        "sheet_bom_ea": 'bom-ea',
        "sheet_bom_eb": 'bom-eb',
        "sheet_coois": 'descarga_coois'
    }

    archivo_generado_ea, archivo_generado_eb = generar_excel_crosstabs_completo(**args)
    print(f'Archivo generado para EA: {archivo_generado_ea}')
    print(f'Archivo generado para EB: {archivo_generado_eb}')
