# PATH: generar_crosstab_completo.py

from src.app.generar_excel_crosstabs_completo import generar_excel_crosstabs_completo

if __name__ == "__main__":
    archivo_excel_master_data = r"prod_files\data\master_data.xlsx"
    archivo_stocks = r"prod_files\data\stocks_EB_2024_06_05.xlsx"
    archivo_coois = r"prod_files\data\coois_eb_ejemplo.xlsx"
    
    args = {
        "archivo_stocks": archivo_stocks,
        "archivo_coois": archivo_coois,
        "archivo_master_data": archivo_excel_master_data
    }
    
    archivo_generado_ea, archivo_generado_eb = generar_excel_crosstabs_completo(**args)