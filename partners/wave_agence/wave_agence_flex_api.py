

from fastapi import HTTPException, Query

import pandas as pd
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    register_db_routes,
    save_single_flex,
)

### uvicorn wave_agence_flex_api:app --port 8031
### Service SPÉCIFIQUE au partenaire WAVE_AGENCE (mode "agence",
### voir config.py). Contrairement à wave_inter_flex_api.py (qui
### rapproche le journal Oracle des transactions interco
### OMB_SN.INTERCOTRANSACTION pour le partenaire WAVE en mode
### "two_pointers"), ce service extrait directement les écritures
### du compte d'attente Wave ("COMPTE ATTENTE WAVE") par agence,
### sans jointure avec la table interco, et alimente le partenaire
### WAVE_AGENCE (réconciliation agrégée par CODE_AGENCE).

PARTENAIRE = "WAVE_AGENCE"
rt = bootstrap_flex_service(PARTENAIRE, title="Wave Agence Flex API", log_prefix="wave_agence_flex_api")
app = rt.app
TABLES = rt.tables
TABLE_AGENCE_FLEX = TABLES["flex"]


# ============================================================
# EXTRACTION ORACLE (SQL SPÉCIFIQUE WAVE — compte attente par agence)
# ============================================================

def get_wave_agence_flex(date_debut: str, date_fin: str) -> pd.DataFrame:

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
       where ac_no in ('102130000010',
'372000000060',
'372000000061',
'372000000062',
'372000000063',
'372000000064',
'372000000065',
'372000000066',
'372000000067',
'372000000068',
'372000000070',
'372000000071',
'372000000072',
'372000000073',
'372000000074',
'372000000075',
'372000000076',
'372000000077',
'372000000078',
'372000000079',
'372000000081',
'372000000082',
'372000000084',
'372000000086',
'372000000087',
'372000000088'
       
       
       )
       
       
       
       
       
       
       )
       
     
     and a.TRN_DT BETWEEN TO_DATE(:m_start, 'DD/MM/YYYY') AND TO_DATE(:m_end, 'DD/MM/YYYY')
     and nvl(c.GL_DESC,s.ac_desc) = 'COMPTE ATTENTE  WAVE'
     and a.BATCH_NO is null
     
order by 
       a.TRN_DT desc
       
 
 

 


  
            

    


    
    """)

    df = oracle_query(rt, sql, {"m_start": date_debut, "m_end": date_fin})
    df.rename(columns={"CODE AGENCE": "CODE_AGENCE", "LIBELLE AGENCE": "LIBELLE_AGENCE"}, inplace=True)
    return df


# ============================================================
# API PRINCIPALE + STOCKAGE SQLITE (table WAVE dédiée)
# ============================================================

@app.get("/wave-agence-flex")
def export_wave_agence_flex(
    date_debut: str,
    date_fin: str,
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):
    df = get_wave_agence_flex(date_debut, date_fin)

    nb_agences = 0
    if "CODE_AGENCE" in df.columns and len(df):
        nb_agences = int(df["CODE_AGENCE"].nunique(dropna=True))

    return save_single_flex(
        rt,
        df,
        sheet_name="WAVE_AGENCE_FLEX",
        filename="WAVE_AGENCE_FLEX.xlsx",
        format=format,
        headers={
            "X-Nb-Ecritures": str(len(df)),
            "X-Nb-Agences": str(nb_agences),
        },
    )


register_db_routes(rt, [
    ("/db/wave-agence", TABLE_AGENCE_FLEX),
])
