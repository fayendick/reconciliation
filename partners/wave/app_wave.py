# ============================================================
# APP WAVE — service Excel dédié au partenaire Wave
# ============================================================
# uvicorn app_wave:app --port 8000
#
# Les fichiers Excel Wave-Partenaire ont DÉJÀ les colonnes du
# schéma standard (DATE TRANSACTION, TYPE TRANSACTION,
# CODE TRANSACTION OPERATEUR, NUMERO COMPTE, MONTANT) donc pas
# de mapping à faire ici.
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

PARTENAIRE = "WAVE"
TABLES = PARTENAIRES[PARTENAIRE]["tables"]

# --------------------------------------------------------
# Mapping de colonnes brutes Wave -> schéma standard.
# Vide car les colonnes Wave sont déjà standard.
# --------------------------------------------------------
MAPPING_COLONNES = {}

print(f"[app_wave.py] Base SQLite utilisée : {DB_PATH}")

engine = make_sqlite_engine()

app = FastAPI(title="Excel Upload API — Wave")


@app.post("/process-excel")
async def process_excel(
    files: List[UploadFile] = File(...),
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):

    try:
        # 1. LECTURE + MAPPING (identité ici) + VALIDATION
        dfs = lire_fichiers_excel(files)
        dfs = appliquer_mapping(dfs, MAPPING_COLONNES)

        for df in dfs:
            valider_colonnes_standard(df, PARTENAIRE)

        # 2. COMPILATION
        df_final = compiler_dataframes(dfs)

        # 3. SPLIT W2B / B2W (mémoire, pour l'export éventuel)
        compile_w2b, compile_b2w = separer_w2b_b2w(df_final)

        # 4. SAUVEGARDE SQLITE : 1 table + 2 vues (pas 3 to_sql)
        sauvegarder_excel_w2b_b2w(df_final, TABLES, engine, "app_wave")

        # 5. EXPORT (Excel si demandé, sinon JSON léger pour /charger)
        return respond_sheets(
            {
                "Compilation": df_final,
                "Compilation_W2B": compile_w2b,
                "Compilation_B2W": compile_b2w,
            },
            filename="Compilation_WAVE.xlsx",
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