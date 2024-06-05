# PATH: generar_crosstab_completo.py

from src.app.generar_excel_crosstabs_completo import generar_excel_crosstabs_completo

if __name__ == "__main__":
    archivo_excel_master_data = r"prod_files\data\master_data.xlsx"
    archivo_stocks = r"C:\Users\18287\Downloads\Materials (5).xlsx"
    archivo_coois = r"C:\Users\18287\Downloads\EXPORT - 2024-06-05T194251.068.xlsx"
    args = {
        "archivo_stocks": archivo_stocks,
        "archivo_coois": archivo_coois,
        "archivo_maestros": archivo_excel_master_data
    }
    
    archivo_generado_ea, archivo_generado_eb = generar_excel_crosstabs_completo(**args)