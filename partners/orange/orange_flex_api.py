from fastapi import HTTPException, Query

import pandas as pd
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    register_db_routes,
    save_split_w2b_b2w,
    split_by_dc,
)

### uvicorn orange_flex_api:app --port 8011
### Squelette du service Orange, calqué sur wave_inter_flex_api.py.
### -> Le rapprochement se fait ici sur : CODE_AGENCE, DEBIT, CREDIT
###    et TYPE_TRANSACTION, ce dernier correspondant directement au
###    sens de l'écriture (DRCR_IND) : "D" = Débit, "C" = Crédit.
### -> Boilerplate via flex_common.

PARTENAIRE = "ORANGE_AGENCE"
rt = bootstrap_flex_service(PARTENAIRE, title="Orange Interco Flex API", log_prefix="orange_flex_api")
app = rt.app
TABLES = rt.tables


# ============================================================
# EXTRACTION ORACLE
# ------------------------------------------------------------
# Colonnes de rapprochement attendues par reconciliation_engine.py :
#   - CODE_AGENCE
#   - DEBIT
#   - CREDIT
#   - TYPE_TRANSACTION  -> sens de l'écriture : "D" (débit) / "C" (crédit)
# ============================================================

def get_orange_inter_flex(date_debut: str, date_fin: str) -> pd.DataFrame:

    sql = text("""
        WITH Journal AS (
            SELECT
                TRN_REF_NO, AC_ENTRY_SR_NO, EVENT_SR_NO, EVENT, AC_BRANCH, AC_NO, AC_CCY, CATEGORY, DRCR_IND, TRN_CODE, FCY_AMOUNT, EXCH_RATE, LCY_AMOUNT, VALUE_DT AS TRN_DT, VALUE_DT, TXN_INIT_DATE, AMOUNT_TAG, RELATED_ACCOUNT, RELATED_CUSTOMER, RELATED_REFERENCE, MIS_HEAD, MIS_FLAG, INSTRUMENT_CODE, BANK_CODE, BALANCE_UPD, AUTH_STAT, MODULE, CUST_GL, DLY_HIST, FINANCIAL_CYCLE, PERIOD_CODE, BATCH_NO, USER_ID, CURR_NO, PRINT_STAT, AUTH_ID, GLMIS_VAL_UPD_FLAG, EXTERNAL_REF_NO, DONT_SHOWIN_STMT, IC_BAL_INCLUSION, AML_EXCEPTION, IB, GLMIS_UPDATE_FLAG, PRODUCT_ACCRUAL, ORIG_PNL_GL, STMT_DT, ENTRY_SEQ_NO, VIRTUAL_AC_NO, CLAIM_AMOUNT, GRP_REF_NO, SAVE_TIMESTAMP, AUTH_TIMESTAMP, PRODUCT_PROCESSOR, RELATED_AC_ENTRY_SR_NO, DONT_SHOWIN_STMT_FEE, ORG_SOURCE, ORG_SOURCE_REF, SOURCE_CODE
            FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES
            WHERE MODULE = 'DE'
            UNION
            SELECT
                TRN_REF_NO, AC_ENTRY_SR_NO, EVENT_SR_NO, EVENT, AC_BRANCH, AC_NO, AC_CCY, CATEGORY, DRCR_IND, TRN_CODE, FCY_AMOUNT, EXCH_RATE, LCY_AMOUNT, TRN_DT, VALUE_DT, TXN_INIT_DATE, AMOUNT_TAG, RELATED_ACCOUNT, RELATED_CUSTOMER, RELATED_REFERENCE, MIS_HEAD, MIS_FLAG, INSTRUMENT_CODE, BANK_CODE, BALANCE_UPD, AUTH_STAT, MODULE, CUST_GL, DLY_HIST, FINANCIAL_CYCLE, PERIOD_CODE, BATCH_NO, USER_ID, CURR_NO, PRINT_STAT, AUTH_ID, GLMIS_VAL_UPD_FLAG, EXTERNAL_REF_NO, DONT_SHOWIN_STMT, IC_BAL_INCLUSION, AML_EXCEPTION, IB, GLMIS_UPDATE_FLAG, PRODUCT_ACCRUAL, ORIG_PNL_GL, STMT_DT, ENTRY_SEQ_NO, VIRTUAL_AC_NO, CLAIM_AMOUNT, GRP_REF_NO, SAVE_TIMESTAMP, AUTH_TIMESTAMP, PRODUCT_PROCESSOR, RELATED_AC_ENTRY_SR_NO, DONT_SHOWIN_STMT_FEE, ORG_SOURCE, ORG_SOURCE_REF, SOURCE_CODE
            FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES
            WHERE MODULE <> 'DE'
        )
        select
               nvl(c.gl_code,s.DR_GL) "PARENT_GL"
               ,nvl(c.GL_DESC,s.ac_desc) "DESCRIPTION"
               ,a.ac_branch "CODE_AGENCE"
               ,b.branch_name "LIBELLE_AGENCE"
               ,decode(a.drcr_ind, 'D', a.lcy_amount, 0) "DEBIT"
               ,decode(a.drcr_ind, 'C', a.lcy_amount, 0) "CREDIT"
               ,a.drcr_ind "TYPE_TRANSACTION"
               ,to_char(a.trn_dt,'dd/mm/yyyy') "DATE_SAISIE"
               ,to_char(a.value_dt,'dd/mm/yyyy') "DATE_VALEUR"
               ,a.USER_ID "UTIL_SAISI"
               ,a.AC_NO "ACCOUNT_NO"
               ,a.AUTH_ID "UTIL_VALID"
               ,a.BATCH_NO
               ,a.trn_ref_no
               ,a.trn_code
               ,nvl((select u.ADDL_TEXT from CFSFCUBS145.detb_upload_detail u WHERE u.BATCH_NO=A.BATCH_NO AND U.ACCOUNT=A.AC_NO and U.VALUE_DATE=A.VALUE_DT and u.AMOUNT=A.lcy_amount AND A.CURR_NO=U.CURR_NO),t.TRN_DESC) "LIBELLE_OPER"
               ,nvl(od.ADDL_TEXT, nvl(xx.ADDL_TEXT, d.DESCRIPTION)) "DESCRIPTION_BATCH"
               ,a.related_customer "MATRICULE_CLIENT"
               ,s.ALT_AC_NO "ACCOUNT_NAFA"
        from Journal a
               LEFT JOIN CFSFCUBS145.gltm_glmaster c ON c.gl_code = a.AC_NO
               LEFT JOIN CFSFCUBS145.STTM_CUST_ACCOUNT s ON s.CUST_AC_NO = a.AC_NO
               left JOIN CFSFCUBS145.STTM_TRN_CODE t ON t.TRN_CODE = a.TRN_CODE
               left JOIN CFSFCUBS145.STTM_BRANCH b ON b.branch_code = a.ac_branch
               left join cfsfcubs145.DETBS_JRNL_TXN_DETAIL xx on a.TRN_REF_NO = xx.REFERENCE_NO and a.EVENT_SR_NO=xx.SERIAL_NO
               left join cfsfcubs145.detb_batch_master d on a.batch_no = d.batch_no and a.ac_branch = d.BRANCH_CODE
               left join cfsfcubs145.detb_upload_detail od on od.batch_no = a.batch_no and a.ac_branch = od.ACCOUNT_BRANCH and a.CURR_NO = od.CURR_NO
        WHERE
               A.TRN_REF_NO in (select TRN_REF_NO from CFSFCUBS145.ACVW_ALL_AC_ENTRIES
               where ac_no in ('102100000015',
        '372000000001',
        '372000000002',
        '372000000003',
        '372000000004',
        '372000000005',
        '372000000006',
        '372000000007',
        '372000000008',
        '372000000009',
        '372000000011',
        '372000000012',
        '372000000013',
        '372000000014',
        '372000000015',
        '372000000016',
        '372000000017',
        '372000000018',
        '372000000019',
        '372000000020',
        '372000000021',
        '372000000022',
        '372000000023',
        '372000000024',
        '372000000025',
        '372000000026'

               ))

             and a.TRN_DT  >= TO_DATE(:m_start,'DD/MM/YYYY')
             AND a.TRN_DT  <=  TO_DATE(:m_end,'DD/MM/YYYY')
             and nvl(c.GL_DESC,s.ac_desc) = 'COMPTE ATTENTE ORANGE MONEY'
             and a.BATCH_NO is null

    """)

    return oracle_query(rt, sql, {"m_start": date_debut, "m_end": date_fin})


# ============================================================
# API PRINCIPALE + STOCKAGE SQLITE (tables ORANGE dédiées)
# ============================================================

@app.get("/orange-flex")
def export_orange_inter_flex(
    date_debut: str,
    date_fin: str,
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):
    df = get_orange_inter_flex(date_debut, date_fin)

    required_cols = {"CODE_AGENCE", "DEBIT", "CREDIT", "TYPE_TRANSACTION"}
    missing = required_cols - set(df.columns)
    if missing:
        raise HTTPException(
            status_code=500,
            detail=f"Colonnes manquantes {missing}. Colonnes disponibles: {list(df.columns)}"
        )

    # TYPE_TRANSACTION : "D" = Débit, "C" = Crédit (DRCR_IND Oracle)
    df["TYPE_TRANSACTION"] = (
        df["TYPE_TRANSACTION"].fillna("").astype(str).str.strip().str.upper()
    )
    debit_df, credit_df = split_by_dc(df)
    return save_split_w2b_b2w(
        rt,
        df,
        sheets={
            "ORANGE_FLEX": df,
            "ORANGE_FLEX_DEBIT": debit_df,
            "ORANGE_FLEX_CREDIT": credit_df,
        },
        filename="ORANGE_FLEX.xlsx",
        format=format,
        mode="dc",
    )


register_db_routes(rt, [
    ("/db/orange", TABLES["flex"]),
    ("/db/orange-debit", TABLES["flex_w2b"]),
    ("/db/orange-credit", TABLES["flex_b2w"]),
])
