# ============================================================
# APP WIZZ — GABARIT pour un nouveau partenaire
# ============================================================
# uvicorn app_wizz:app --port 8020
#
# Ce fichier est un MODÈLE à copier-coller pour brancher
# n'importe quel nouveau partenaire :
#   1. Copier ce fichier -> app_<partenaire>.py
#   2. Changer PARTENAIRE = "..." (doit exister dans
#      config.PARTENAIRES)
#   3. Remplir MAPPING_COLONNES avec les vrais noms de colonnes
#      du fichier Excel de ce partenaire
#   4. Changer le port au lancement (uvicorn ... --port XXXX)
# Rien d'autre à modifier : la suite du fichier est identique
# pour tout le monde.
# ============================================================

import traceback
from typing import List

import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from config import DB_PATH, PARTENAIRES, make_sqlite_engine
from common.excel_common import (
    lire_fichiers_excel,
    appliquer_mapping,
    valider_colonnes_standard,
    compiler_dataframes,
    separer_w2b_b2w,
    sauvegarder_excel_w2b_b2w,
)
from common.http_export import respond_sheets
from common.sqlite_io import lire_table_json

PARTENAIRE = "WIZZ"
TABLES = PARTENAIRES[PARTENAIRE]["tables"]

# --------------------------------------------------------
# TODO : à remplir avec les VRAIS noms de colonnes du fichier
# Excel de ce partenaire -> noms standard (config.COLONNES_STANDARD_EXCEL)
# --------------------------------------------------------
MAPPING_COLONNES = {
    # "NOM_BRUT_COLONNE_WIZZ": "DATE TRANSACTION",
    # "NOM_BRUT_COLONNE_WIZZ": "TYPE TRANSACTION",
    # "NOM_BRUT_COLONNE_WIZZ": "CODE TRANSACTION OPERATEUR",
    # "NOM_BRUT_COLONNE_WIZZ": "NUMERO COMPTE",
    # "NOM_BRUT_COLONNE_WIZZ": "MONTANT",
}

print(f"[app_wizz.py] Base SQLite utilisée : {DB_PATH}")

engine = make_sqlite_engine()

app = FastAPI(title=f"Excel Upload API — {PARTENAIRE.title()}")


@app.post("/process-excel")
async def process_excel(
    files: List[UploadFile] = File(...),
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):

    try:
        dfs = lire_fichiers_excel(files)
        dfs = appliquer_mapping(dfs, MAPPING_COLONNES)

        for df in dfs:
            valider_colonnes_standard(df, PARTENAIRE)

        df_final = compiler_dataframes(dfs)
        compile_w2b, compile_b2w = separer_w2b_b2w(df_final)

        sauvegarder_excel_w2b_b2w(df_final, TABLES, engine, "app_wizz")

        return respond_sheets(
            {
                "Compilation": df_final,
                "Compilation_W2B": compile_w2b,
                "Compilation_B2W": compile_b2w,
            },
            filename=f"Compilation_{PARTENAIRE}.xlsx",
            format=format,
        )

    except HTTPException:
        raise
    except Exception as e:
        return {"status": "error", "message": str(e), "trace": traceback.format_exc()}


@app.get("/db/compilation")
def get_compilation(limit: int = Query(None), offset: int = Query(0)):
    return lire_table_json(engine, TABLES["excel"], limit=limit, offset=offset)


@app.get("/db/w2b")
def get_w2b(limit: int = Query(None), offset: int = Query(0)):
    return lire_table_json(engine, TABLES["excel_w2b"], limit=limit, offset=offset)


@app.get("/db/b2w")
def get_b2w(limit: int = Query(None), offset: int = Query(0)):
    return lire_table_json(engine, TABLES["excel_b2w"], limit=limit, offset=offset)


@app.get("/health")
def health():
    return {"status": "ok", "partenaire": PARTENAIRE, "db_path": DB_PATH}