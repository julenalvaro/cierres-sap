# PATH: generar_crosstab_completo.py

from src.app.generar_excel_crosstabs_completo import generar_excel_crosstabs_completo

if __name__ == "__main__":
    archivo_excel_master_data = r"prod_files\data\master_data.xlsx"
    archivo_stocks = r"C:\Users\18287\OneDrive - Construcciones y Auxiliar deFerrocarriles SA (CAF)\projects\25-mudanza-ordenes-baan\archivos\descargas\descarga_producción\comparativa_idiomas\Materiales_ES.xlsx"
    archivo_coois = r"C:\Users\18287\OneDrive - Construcciones y Auxiliar deFerrocarriles SA (CAF)\projects\25-mudanza-ordenes-baan\archivos\descargas\descarga_producción\comparativa_idiomas\coois_es.xlsx"
    
    args = {
        "archivo_stocks": archivo_stocks,
        "archivo_coois": archivo_coois,
        "archivo_maestros": archivo_excel_master_data
    }
    
    archivo_generado_ea, archivo_generado_eb = generar_excel_crosstabs_completo(**args)