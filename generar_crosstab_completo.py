# PATH: generar_crosstab_completo.py

from src.app.generar_excel_crosstabs_completo import generar_excel_crosstabs_completo

if __name__ == "__main__":
    archivo_excel_master_data = r"C:\Users\18287\OneDrive - Construcciones y Auxiliar deFerrocarriles SA (CAF)\projects\25-mudanza-ordenes-baan\archivos\descargas\subset\datos descargados - 24_05_2024 - complet.xlsx"
    archivo_stocks = 
    archivo_coois = 
    args = {
        "archivo_stocks": archivo_stocks,
        "archivo_coois": archivo_coois,
        "archivo_master_data": archivo_excel_master_data,
        "sheet_bom_ea": 'bom-ea',
        "sheet_bom_eb": 'bom-eb',
        "sheet_fabricacion_real_ea": 'fabricacion_real_ea',
        "sheet_fabricacion_real_eb": 'fabricacion_real_eb'
    }
    
    archivo_generado_ea, archivo_generado_eb = generar_excel_crosstabs_completo(**args)