# PATH: src/service/transformar_bom_a_arbol_correcciones.py

import pandas as pd
import numpy as np
import traceback

def transformar_bom_a_arbol_correcciones(bom, coois, stocks, fabricacion_real, subset_name):
    try:
        # Renombrar columnas y cambiar tipos de datos
        bom.rename(columns={"Texto breve-objeto": "Texto breve"}, inplace=True)
        bom = bom.astype({"index": int, "Versión fabricación": int, "Nivel explosión": str, 
                          "Nº componentes": str, "Texto breve": str, "Grupo de artículos": str, 
                          "Almacén producción": str, "Ctd.componente (UMB)": int, 
                          "Nivel": int, "Ruta (predecesor)": int, "Entrada en tabla": int, 
                          "Modelo": str, "pos_estructura": str})

        # Crear la columna "mod-mat"
        bom['mod-mat'] = bom['Modelo'] + '-' + bom['Nº componentes']

        # Lista de valores a excluir
        valores_a_excluir = {"25172023", "25173928", "25173930", "31122300", "31231500", 
                              "31341100", "31341800", "31400000", "39121407", "39121600", 
                              "39121721", "Z00350000", "Z00170000"}

        # Filtrar la tabla excluyendo los valores de la lista
        bom = bom[~bom['Grupo de artículos'].isin(valores_a_excluir)]
        bom = bom[(bom['Versión fabricación'] != 0) &
                  (~bom['Grupo de artículos'].isin({"Z00360000", "Z00380000", "Z00440000", 
                                                      "Z00460000", "Z00500000", "Z00550000", 
                                                      "Z00570000", "Z00590000", "Z00600000"}))]

        # Añadir columna "fase"
        conditions = [
            (bom['Nº componentes'].str.contains("Z1")),
            (bom['Nº componentes'].str.contains("Z2")),
            (bom['Texto breve'].notna() & bom['Texto breve'].str.lower().str.contains("chorr")),
            (bom['Grupo de artículos'] == "Z00210000")
        ]
        choices = ["Z1", "Z2", "PINTURA", "SUBCONJUNTOS"]
        bom['fase'] = np.select(conditions, choices, default=None)

        # Merge con descarga_coois
        bom = pd.merge(bom, coois, on="mod-mat", how="inner")
        bom['ordenes_GMEIN_disp'] = bom.groupby('mod-mat')['Order quantity (GMEIN)'].transform('sum')

        # Merge con fabricacion_real
        bom['mod-fase'] = bom['Modelo'] + '-' + bom['fase']
        bom = pd.merge(bom, fabricacion_real, left_on="mod-fase", right_on="Mod-Fas", how="left")
        bom.rename(columns={"cant_no_fab": "ordenes_necesarias"}, inplace=True)

        # Reordenar y limpiar las columnas para el resultado final
        bom['unidades_necesarias'] = bom['ordenes_necesarias']
        bom['ordenes_necesarias'] = bom['unidades_necesarias'] * bom['Ctd.componente (UMB)']
        bom['margen_ordenes'] = bom['ordenes_GMEIN_disp'] - bom['ordenes_necesarias']
        bom['proj-mat'] = bom['Modelo'].str.slice(2, 6) + '-' + bom['Nº componentes']

        # Merge con stocks
        bom = pd.merge(bom, stocks, on="proj-mat", how="left")
        bom['stock'] = bom.groupby('proj-mat')['Unrestricted Stock'].transform('count')

        # Reordenar y limpiar las columnas para el resultado final
        column_order = ["index", "Versión fabricación", "Nivel explosión", "Nº componentes", 
                        "Texto breve", "Grupo de artículos", "Almacén producción", "Nivel", 
                        "Ruta (predecesor)", "Entrada en tabla", "Modelo", "pos_estructura", 
                        "mod-mat", "fase", "mod-fase", "proj-mat", "Ctd.componente (UMB)", 
                        "unidades_necesarias", "ordenes_necesarias", "ordenes_GMEIN_disp", 
                        "margen_ordenes", "stock"]

        # Verificar si todas las columnas están presentes antes de reordenar
        missing_columns = [col for col in column_order if col not in bom.columns]
        if missing_columns:
            raise KeyError(f"Faltan las siguientes columnas: {missing_columns}")

        return bom[column_order]
    except Exception as e:
        print(f"Error en transformar_bom_a_arbol_correcciones: {e}")
        traceback.print_exc()
        return pd.DataFrame()