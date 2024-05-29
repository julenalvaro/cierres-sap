# PATH: src/service/transformar_bom_a_arbol_correcciones.py

import pandas as pd
import numpy as np
import traceback

def transformar_bom_a_arbol_correcciones(bom, coois, fabr_real, stocks):
    # Verificar y renombrar columnas
    if "Texto breve-objeto" in bom.columns:
        bom.rename(columns={"Texto breve-objeto": "Texto breve"}, inplace=True)

    # Cambiar tipos de datos si las columnas existen
    cols_to_convert = {
        "index": "int64",
        "Versión fabricación": "int64",
        "Nivel explosión": "str",
        "Nº componentes": "str",
        "Texto breve": "str",
        "Grupo de artículos": "str",
        "Almacén producción": "str",
        "Ctd.componente (UMB)": "int64",
        "Nivel": "int64",
        "Ruta (predecesor)": "int64",
        "Entrada en tabla": "int64",
        "Modelo": "str",
        "pos_estructura": "str"
    }

    for col, dtype in cols_to_convert.items():
        if col in bom.columns:
            bom[col] = bom[col].astype(dtype)

    # Añadir columna personalizada
    bom["mod-mat"] = bom["Modelo"] + "-" + bom["Nº componentes"]

    # Filtrar valores excluidos
    valores_a_excluir = [
        "25172023", "25173928", "25173930", "31122300", "31231500", "31341100", "31341800", "31400000",
        "39121407", "39121600", "39121721", "Z00350000", "Z00170000"
    ]
    bom = bom[
        ~bom["Grupo de artículos"].isin(valores_a_excluir) &
        (bom["Versión fabricación"] != 0) &
        (~bom["Grupo de artículos"].isin([
            "Z00360000", "Z00380000", "Z00440000", "Z00460000", "Z00500000", "Z00550000", "Z00570000", "Z00590000", "Z00600000"
        ]))
    ]

    # Añadir columna "fase"
    bom.loc[:, "fase"] = bom.apply(
    lambda row: "Z1" if "Z1" in row["Nº componentes"] else
                "Z2" if "Z2" in row["Nº componentes"] else
                "PINTURA" if pd.notnull(row["Texto breve"]) and "chorr" in row["Texto breve"].lower() else
                "SUBCONJUNTOS" if row["Grupo de artículos"] == "Z00210000" else
                None, axis=1)

    # Sumarizar coois antes del merge
    coois_aggregated = coois.groupby('mod-mat').agg({'Order quantity (GMEIN)': 'sum'}).reset_index()
    coois_aggregated.rename(columns={'Order quantity (GMEIN)': 'ordenes_GMEIN_disp'}, inplace=True)

    # Unir con coois sumarizado
    bom = bom.merge(coois_aggregated, on="mod-mat", how="inner")

    # Añadir columna personalizada "mod-fase"
    bom["mod-fase"] = bom["Modelo"] + "-" + bom["fase"]

    # Unir con fabricacion_real_acabados
    bom = bom.merge(fabr_real, left_on="mod-fase", right_on="Mod-Fas", how="left")
    bom.rename(columns={"cant_no_fab": "ordenes_necesarias"}, inplace=True)

    # Ordenar y calcular columnas adicionales
    bom.sort_values(by=["index"], inplace=True)
    bom["unidades_necesarias"] = bom["ordenes_necesarias"] * bom["Ctd.componente (UMB)"]
    bom["margen_ordenes"] = bom["ordenes_GMEIN_disp"] - bom["unidades_necesarias"]
    bom["proj-mat"] = bom["Modelo"].str[2:6] + "-" + bom["Nº componentes"]

    # Sumarizar stocks antes del merge
    stocks_aggregated = stocks.groupby('proj-mat').agg({'Unrestricted Stock': 'sum'}).reset_index()
    stocks_aggregated.rename(columns={'Unrestricted Stock': 'stock'}, inplace=True)

    # Unir con stocks sumarizado
    bom = bom.merge(stocks_aggregated, on="proj-mat", how="left")

    # Reordenar columnas
    bom = bom[[
        "index", "Versión fabricación", "Nivel explosión", "Nº componentes", "Texto breve", "Grupo de artículos",
        "Almacén producción", "Nivel", "Ruta (predecesor)", "Entrada en tabla", "Modelo", "pos_estructura", "mod-mat",
        "fase", "mod-fase", "proj-mat", "Ctd.componente (UMB)", "unidades_necesarias", "ordenes_necesarias",
        "ordenes_GMEIN_disp", "margen_ordenes", "stock"
    ]]

    return bom