from fastapi import HTTPException, Query

import pandas as pd
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    register_db_routes,
    save_split_w2b_b2w,
    split_by_type_contains,
)

### uvicorn wizz_flex_api:app --port 8021
### GABARIT à copier-coller pour tout nouveau partenaire côté
### extraction Oracle. TODO : remplacer le SQL par la vraie
### requête Oracle de ce partenaire.

PARTENAIRE = "WIZZ"
rt = bootstrap_flex_service(PARTENAIRE, title="Wizz Interco Flex API", log_prefix="wizz_flex_api")
app = rt.app
TABLES = rt.tables


# ============================================================
# EXTRACTION ORACLE — TODO : remplacer par le SQL réel Wizz
# ------------------------------------------------------------
# Colonnes minimales attendues en sortie (après upper()) :
#   CODE_TRANSACTION_OPERATEUR, NUMERO_COMPTE, TYPE_TRANSACTION
#   (doit contenir "W2B"/"B2W"), DATE_VALEUR,
#   MOUVEMENT_DEBIT / MOUVEMENT_CREDIT
# ============================================================

def get_wizz_inter_flex(date_debut: str, date_fin: str) -> pd.DataFrame:

    sql = text("""
        -- TODO : requête Oracle spécifique Wizz
        SELECT
            s.DT_CRE_ENREG,
            s.DT_MAJ_ENREG,
            s.TXNREFNO,
            s.CODE_TRANSACTION_OPERATEUR,
            s.CODE_TRANSACTION,
            s.REF_TRANSACTION_BQ,
            s.STATUT,
            s.NUMERO_COMPTE,
            s.FRAIS,
            s.TYPE_TRANSACTION
        FROM OMB_SN.INTERCOTRANSACTION_WIZZ s
        WHERE s.TXNREFNO <> '0'
          AND s.TRANSDATE >= TO_DATE(:m_start,'DD/MM/YYYY')
          AND s.TRANSDATE <  TO_DATE(:m_end,'DD/MM/YYYY') + 1
    """)

    return oracle_query(rt, sql, {"m_start": date_debut, "m_end": date_fin})


@app.get("/wizz-flex")
def export_wizz_inter_flex(
    date_debut: str,
    date_fin: str,
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):
    df = get_wizz_inter_flex(date_debut, date_fin)

    if "TYPE_TRANSACTION" not in df.columns:
        raise HTTPException(status_code=500, detail=f"TYPE_TRANSACTION absente. Colonnes: {list(df.columns)}")

    df["TYPE_TRANSACTION"] = df["TYPE_TRANSACTION"].fillna("").astype(str).str.upper()
    w2b, b2w = split_by_type_contains(df)
    return save_split_w2b_b2w(
        rt,
        df,
        sheets={"WIZZ_FLEX": df, "WIZZ_FLEX_W2B": w2b, "WIZZ_FLEX_B2W": b2w},
        filename="WIZZ_FLEX.xlsx",
        format=format,
        mode="contains",
    )


register_db_routes(rt, [
    ("/db/wizz", TABLES["flex"]),
    ("/db/wizz-w2b", TABLES["flex_w2b"]),
    ("/db/wizz-b2w", TABLES["flex_b2w"]),
])
