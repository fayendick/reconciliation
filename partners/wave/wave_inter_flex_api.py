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

### uvicorn wave_inter_flex_api:app --port 8001
### Service SPÉCIFIQUE à Wave (SQL Oracle ci-dessous).

PARTENAIRE = "WAVE"
rt = bootstrap_flex_service(PARTENAIRE, title="Wave Interco Flex API", log_prefix="wave_inter_flex_api")
app = rt.app
TABLES = rt.tables



# ============================================================
# EXTRACTION ORACLE (SQL SPÉCIFIQUE WAVE — inchangé)
# ============================================================

def get_wave_inter_flex(date_debut: str, date_fin: str) -> pd.DataFrame:

    sql = text("""
    WITH Journal AS (

        SELECT
            NVL(c.PARENT_GL, s.DR_GL) PARENT_GL,
            NVL(c.GL_DESC, s.AC_DESC) DESCRIPTION,
            a.AC_NO ACCOUNT_NO,
            a.AC_BRANCH CODE_AGENCE,
            a.EXTERNAL_REF_NO,
            b.BRANCH_NAME LIBELLE_AGENCE,

            DECODE(a.DRCR_IND,'D',a.LCY_AMOUNT,0) MOUVEMENT_DEBIT,
            DECODE(a.DRCR_IND,'C',a.LCY_AMOUNT,0) MOUVEMENT_CREDIT,

            TO_CHAR(a.TRN_DT,'DD/MM/YYYY HH24:MI:SS') DATE_SAISIE,
            a.SAVE_TIMESTAMP AS DATE_VALEUR,

            a.USER_ID UTIL_SAISI,
            a.AUTH_ID UTIL_VALID,
            a.BATCH_NO,
            a.GRP_REF_NO NO_DOSSIER,
            a.TRN_REF_NO,
            a.TRN_CODE,
            t.TRN_DESC,
            a.RELATED_ACCOUNT,
            a.MODULE,

            NVL(
                od.ADDL_TEXT,
                NVL(xx.ADDL_TEXT,d.DESCRIPTION)
            ) DESCRIPTION_BATCH

        FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES a

        LEFT JOIN CFSFCUBS145.GLTM_GLMASTER c
            ON c.GL_CODE = a.AC_NO

        LEFT JOIN CFSFCUBS145.STTM_CUST_ACCOUNT s
            ON s.CUST_AC_NO = a.AC_NO

        LEFT JOIN CFSFCUBS145.STTM_TRN_CODE t
            ON t.TRN_CODE = a.TRN_CODE

        LEFT JOIN CFSFCUBS145.STTM_BRANCH b
            ON b.BRANCH_CODE = a.AC_BRANCH

        LEFT JOIN CFSFCUBS145.DETBS_JRNL_TXN_DETAIL xx
            ON a.TRN_REF_NO = xx.REFERENCE_NO
            AND a.EVENT_SR_NO = xx.SERIAL_NO

        LEFT JOIN CFSFCUBS145.DETB_BATCH_MASTER d
            ON a.BATCH_NO = d.BATCH_NO
            AND a.AC_BRANCH = d.BRANCH_CODE

        LEFT JOIN CFSFCUBS145.DETB_UPLOAD_DETAIL od
            ON od.BATCH_NO = a.BATCH_NO
            AND od.ACCOUNT_BRANCH = a.AC_BRANCH
            AND od.CURR_NO = a.CURR_NO

        WHERE a.SAVE_TIMESTAMP >= TO_DATE(:m_start,'DD/MM/YYYY')
          AND a.SAVE_TIMESTAMP <  TO_DATE(:m_end,'DD/MM/YYYY') + 1
    ),

    COFIMOBILPLUS AS (

        SELECT
            NVL(TXNREFNO,0) TXNREFNO,
            DT_CRE_ENREG,
            DT_MAJ_ENREG,
            NUMERO_COMPTE,
            FRAIS,
            MESSAGE,
            MONTANT,
            STATUT,
            TRANSDATE,
            CODE_TRANSACTION_OPERATEUR,
            CODE_TRANSACTION,
            REF_TRANSACTION_BQ,
            TYPE_TRANSACTION
        FROM OMB_SN.INTERCOTRANSACTION
        WHERE TXNREFNO <> '0'
    )

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
        s.TYPE_TRANSACTION,

        j.PARENT_GL,
        j.DESCRIPTION,
        j.ACCOUNT_NO,
        j.CODE_AGENCE,
        j.EXTERNAL_REF_NO,
        j.LIBELLE_AGENCE,
        j.MOUVEMENT_DEBIT,
        j.MOUVEMENT_CREDIT,
        j.DATE_SAISIE,
        j.DATE_VALEUR,
        j.UTIL_SAISI,
        j.UTIL_VALID,
        j.BATCH_NO,
        j.NO_DOSSIER,
        j.TRN_REF_NO,
        j.TRN_CODE,
        j.TRN_DESC,
        j.RELATED_ACCOUNT,
        j.MODULE,
        j.DESCRIPTION_BATCH

    FROM Journal j
    JOIN COFIMOBILPLUS s
        ON s.TXNREFNO = j.EXTERNAL_REF_NO

    WHERE (j.ACCOUNT_NO LIKE '251%' OR j.ACCOUNT_NO LIKE '253%')

    ORDER BY j.DATE_VALEUR DESC
    """)

    return oracle_query(rt, sql, {"m_start": date_debut, "m_end": date_fin})


# ============================================================
# API PRINCIPALE + STOCKAGE SQLITE (tables WAVE dédiées)
# ============================================================

@app.get("/wave-inter-flex")
def export_wave_inter_flex(
    date_debut: str,
    date_fin: str,
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):
    df = get_wave_inter_flex(date_debut, date_fin)

    if "TYPE_TRANSACTION" not in df.columns:
        raise HTTPException(
            status_code=500,
            detail=f"TYPE_TRANSACTION absente. Colonnes: {list(df.columns)}"
        )

    df["TYPE_TRANSACTION"] = df["TYPE_TRANSACTION"].fillna("").astype(str).str.upper()
    w2b, b2w = split_by_type_contains(df)
    return save_split_w2b_b2w(
        rt,
        df,
        sheets={"WAVE_FLEX": df, "WAVE_FLEX_W2B": w2b, "WAVE_FLEX_B2W": b2w},
        filename="WAVE_FLEX.xlsx",
        format=format,
        mode="contains",
    )


register_db_routes(rt, [
    ("/db/wave", TABLES["flex"]),
    ("/db/wave-w2b", TABLES["flex_w2b"]),
    ("/db/wave-b2w", TABLES["flex_b2w"]),
])