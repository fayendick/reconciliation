from fastapi import HTTPException, Query

import pandas as pd
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    register_db_routes,
    save_single_flex,
)

### uvicorn RIA_agence_flex:app --port 8033
### Service SPÉCIFIQUE au partenaire RIA_AGENCE (mode "agence").
### Ce service extrait directement les écritures du compte transit
### RIA ("CPTE TRANSIT ET TRANSITOIR RIA") par agence, sans
### jointure avec la table interco, et alimente le partenaire
### RIA_AGENCE (réconciliation agrégée par CODE_AGENCE).

PARTENAIRE = "RIA_AGENCE"
rt = bootstrap_flex_service(PARTENAIRE, title="RIA Agence Flex API", log_prefix="RIA_agence_flex")
app = rt.app
TABLES = rt.tables
TABLE_AGENCE_FLEX = TABLES["flex"]


# ============================================================
# EXTRACTION ORACLE (SQL SPÉCIFIQUE RIA — compte transit par agence)
# ============================================================

def get_ria_agence_flex(date_debut: str, date_fin: str) -> pd.DataFrame:

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
       ,a.ac_branch "CODE AGENCE"
       ,b.branch_name "LIBELLE AGENCE"
       ,decode(a.drcr_ind, 'D', a.lcy_amount, 0) Debit
       ,decode(a.drcr_ind, 'C', a.lcy_amount, 0) Credit
       ,to_char(a.trn_dt,'dd/mm/yyyy')"DATE_SAISIE"
       ,to_char(a.value_dt,'dd/mm/yyyy')"DATE_VALEUR"
       ,a.USER_ID "UTIL SAISI"
       ,a.AC_NO "ACCOUNT_NO"
       ,a.AUTH_ID "UTIL VALID"
       ,a.BATCH_NO
       ,a.trn_ref_no
       ,a.trn_code
       --t.TRN_DESC,
       ,nvl((select u.ADDL_TEXT from CFSFCUBS145.detb_upload_detail u WHERE u.BATCH_NO=A.BATCH_NO AND U.ACCOUNT=A.AC_NO and U.VALUE_DATE=A.VALUE_DT and u.AMOUNT=A.lcy_amount AND A.CURR_NO=U.CURR_NO),t.TRN_DESC)"LIBELLE_OPER"
       ,nvl(od.ADDL_TEXT, nvl(xx.ADDL_TEXT, d.DESCRIPTION)) "DESCRIPTION BATCH"
       ,a.related_customer "MATRICULE_CLIENT"
       ,s.ALT_AC_NO "ACCOUNT NAFA"
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
       where ac_no in ('372000000089',
'372000000090',
'372000000091',
'372000000092',
'372000000093',
'372000000094',
'372000000095',
'372000000096',
'372000000097',
'372000000099',
'372000000100',
'372000000101',
'372000000102',
'372000000103',
'372000000104',
'372000000105',
'372000000106',
'372000000107',
'372000000108',
'372000000109',
'372000000111',
'372000000112',
'372000000114',
'372000000115',
'372000000116'
))
       
       
    
    
    
  
    and a.TRN_DT BETWEEN TO_DATE(:m_start, 'DD/MM/YYYY') AND TO_DATE(:m_end, 'DD/MM/YYYY')
    
    
    and nvl(c.GL_DESC,s.ac_desc) = 'CPTE TRANSIT ET TRANSITOIR RIA'
    and a.BATCH_NO is null 
    
order by 
       a.TRN_DT desc
    
    

    
    
    """)

    df = oracle_query(rt, sql, {"m_start": date_debut, "m_end": date_fin})
    df.rename(columns={"CODE AGENCE": "CODE_AGENCE", "LIBELLE AGENCE": "LIBELLE_AGENCE"}, inplace=True)
    return df


# ============================================================
# API PRINCIPALE + STOCKAGE SQLITE (table RIA dédiée)
# ============================================================

@app.get("/ria-agence-flex")
def export_ria_agence_flex(
    date_debut: str,
    date_fin: str,
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):
    df = get_ria_agence_flex(date_debut, date_fin)
    return save_single_flex(
        rt,
        df,
        sheet_name="RIA_AGENCE_FLEX",
        filename="RIA_AGENCE_FLEX.xlsx",
        format=format,
    )


register_db_routes(rt, [
    ("/db/ria-agence", TABLE_AGENCE_FLEX),
])
