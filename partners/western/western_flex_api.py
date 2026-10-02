# ============================================================
# WESTERN FLEX API
# Service d'extraction Oracle dédié au partenaire WESTERN
# Réconciliation agrégée par agence
# ============================================================

from fastapi import Query

import pandas as pd
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    register_db_routes,
    save_single_flex,
)


# ============================================================
# CONFIGURATION
# ============================================================

PARTENAIRE = "WESTERN"

rt = bootstrap_flex_service(
    PARTENAIRE,
    title="Western Flex API",
    log_prefix="western_flex_api",
)

app = rt.app
TABLES = rt.tables
TABLE_WESTERN_FLEX = TABLES["flex"]


# ============================================================
# EXTRACTION ORACLE WESTERN
# ============================================================

def get_western_flex(
    date_debut: str,
    date_fin: str,
) -> pd.DataFrame:

    sql = text("""
        WITH Journal AS (

            SELECT
                TRN_REF_NO,
                AC_ENTRY_SR_NO,
                EVENT_SR_NO,
                EVENT,
                AC_BRANCH,
                AC_NO,
                AC_CCY,
                CATEGORY,
                DRCR_IND,
                TRN_CODE,
                FCY_AMOUNT,
                EXCH_RATE,
                LCY_AMOUNT,
                VALUE_DT AS TRN_DT,
                VALUE_DT,
                TXN_INIT_DATE,
                AMOUNT_TAG,
                RELATED_ACCOUNT,
                RELATED_CUSTOMER,
                RELATED_REFERENCE,
                MIS_HEAD,
                MIS_FLAG,
                INSTRUMENT_CODE,
                BANK_CODE,
                BALANCE_UPD,
                AUTH_STAT,
                MODULE,
                CUST_GL,
                DLY_HIST,
                FINANCIAL_CYCLE,
                PERIOD_CODE,
                BATCH_NO,
                USER_ID,
                CURR_NO,
                PRINT_STAT,
                AUTH_ID,
                GLMIS_VAL_UPD_FLAG,
                EXTERNAL_REF_NO,
                DONT_SHOWIN_STMT,
                IC_BAL_INCLUSION,
                AML_EXCEPTION,
                IB,
                GLMIS_UPDATE_FLAG,
                PRODUCT_ACCRUAL,
                ORIG_PNL_GL,
                STMT_DT,
                ENTRY_SEQ_NO,
                VIRTUAL_AC_NO,
                CLAIM_AMOUNT,
                GRP_REF_NO,
                SAVE_TIMESTAMP,
                AUTH_TIMESTAMP,
                PRODUCT_PROCESSOR,
                RELATED_AC_ENTRY_SR_NO,
                DONT_SHOWIN_STMT_FEE,
                ORG_SOURCE,
                ORG_SOURCE_REF,
                SOURCE_CODE

            FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES

            WHERE MODULE = 'DE'


            UNION


            SELECT
                TRN_REF_NO,
                AC_ENTRY_SR_NO,
                EVENT_SR_NO,
                EVENT,
                AC_BRANCH,
                AC_NO,
                AC_CCY,
                CATEGORY,
                DRCR_IND,
                TRN_CODE,
                FCY_AMOUNT,
                EXCH_RATE,
                LCY_AMOUNT,
                TRN_DT,
                VALUE_DT,
                TXN_INIT_DATE,
                AMOUNT_TAG,
                RELATED_ACCOUNT,
                RELATED_CUSTOMER,
                RELATED_REFERENCE,
                MIS_HEAD,
                MIS_FLAG,
                INSTRUMENT_CODE,
                BANK_CODE,
                BALANCE_UPD,
                AUTH_STAT,
                MODULE,
                CUST_GL,
                DLY_HIST,
                FINANCIAL_CYCLE,
                PERIOD_CODE,
                BATCH_NO,
                USER_ID,
                CURR_NO,
                PRINT_STAT,
                AUTH_ID,
                GLMIS_VAL_UPD_FLAG,
                EXTERNAL_REF_NO,
                DONT_SHOWIN_STMT,
                IC_BAL_INCLUSION,
                AML_EXCEPTION,
                IB,
                GLMIS_UPDATE_FLAG,
                PRODUCT_ACCRUAL,
                ORIG_PNL_GL,
                STMT_DT,
                ENTRY_SEQ_NO,
                VIRTUAL_AC_NO,
                CLAIM_AMOUNT,
                GRP_REF_NO,
                SAVE_TIMESTAMP,
                AUTH_TIMESTAMP,
                PRODUCT_PROCESSOR,
                RELATED_AC_ENTRY_SR_NO,
                DONT_SHOWIN_STMT_FEE,
                ORG_SOURCE,
                ORG_SOURCE_REF,
                SOURCE_CODE

            FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES

            WHERE MODULE <> 'DE'
        )


        SELECT

            NVL(c.gl_code, s.DR_GL) AS "PARENT_GL",

            NVL(c.GL_DESC, s.ac_desc) AS "DESCRIPTION",

            a.ac_branch AS "CODE AGENCE",

            b.branch_name AS "LIBELLE AGENCE",

            DECODE(
                a.drcr_ind,
                'D',
                a.lcy_amount,
                0
            ) AS DEBIT,

            DECODE(
                a.drcr_ind,
                'C',
                a.lcy_amount,
                0
            ) AS CREDIT,

            TO_CHAR(
                a.trn_dt,
                'dd/mm/yyyy'
            ) AS "DATE_SAISIE",

            TO_CHAR(
                a.value_dt,
                'dd/mm/yyyy'
            ) AS "DATE_VALEUR",

            a.USER_ID AS "UTIL SAISI",

            a.AC_NO AS "ACCOUNT_NO",

            a.AUTH_ID AS "UTIL VALID",

            a.BATCH_NO,

            a.trn_ref_no,

            a.trn_code,


            NVL(
                (
                    SELECT
                        u.ADDL_TEXT

                    FROM CFSFCUBS145.detb_upload_detail u

                    WHERE
                        u.BATCH_NO = A.BATCH_NO
                        AND U.ACCOUNT = A.AC_NO
                        AND U.VALUE_DATE = A.VALUE_DT
                        AND u.AMOUNT = A.lcy_amount
                        AND A.CURR_NO = U.CURR_NO
                ),
                t.TRN_DESC
            ) AS "LIBELLE_OPER",


            NVL(
                od.ADDL_TEXT,
                NVL(
                    xx.ADDL_TEXT,
                    d.DESCRIPTION
                )
            ) AS "DESCRIPTION BATCH",


            a.related_customer AS "MATRICULE_CLIENT",

            s.ALT_AC_NO AS "ACCOUNT NAFA"


        FROM Journal a


        LEFT JOIN CFSFCUBS145.gltm_glmaster c
            ON c.gl_code = a.AC_NO


        LEFT JOIN CFSFCUBS145.STTM_CUST_ACCOUNT s
            ON s.CUST_AC_NO = a.AC_NO


        LEFT JOIN CFSFCUBS145.STTM_TRN_CODE t
            ON t.TRN_CODE = a.TRN_CODE


        LEFT JOIN CFSFCUBS145.STTM_BRANCH b
            ON b.branch_code = a.ac_branch


        LEFT JOIN cfsfcubs145.DETBS_JRNL_TXN_DETAIL xx
            ON a.TRN_REF_NO = xx.REFERENCE_NO
            AND a.EVENT_SR_NO = xx.SERIAL_NO


        LEFT JOIN cfsfcubs145.detb_batch_master d
            ON a.batch_no = d.batch_no
            AND a.ac_branch = d.BRANCH_CODE


        LEFT JOIN cfsfcubs145.detb_upload_detail od
            ON od.batch_no = a.batch_no
            AND a.ac_branch = od.ACCOUNT_BRANCH
            AND a.CURR_NO = od.CURR_NO


        WHERE

            A.TRN_REF_NO IN (

                SELECT
                    TRN_REF_NO

                FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES

                WHERE AC_NO IN (

                    '371120000002',
                    '371120000003',
                    '371120000004',
                    '371120000005',
                    '371120000006',
                    '371120000007',
                    '371120000008',
                    '371120000009',
                    '371120000010',
                    '371120000012',
                    '371120000013',
                    '371120000014',
                    '371120000016',
                    '371120000017',
                    '371120000018',
                    '371120000019',
                    '371120000020',
                    '371120000021',
                    '371120000022',
                    '371120000024',
                    '371120000026',
                    '371120000027',
                    '371120000028',
                    '371120000029',
                    '371120000030'
                )
            )


            AND a.TRN_DT BETWEEN
                TO_DATE(:m_start, 'DD/MM/YYYY')
                AND TO_DATE(:m_end, 'DD/MM/YYYY')


        ORDER BY
            a.TRN_DT DESC
    """)


    df = oracle_query(
        rt,
        sql,
        {
            "m_start": date_debut,
            "m_end": date_fin,
        },
    )


    # ========================================================
    # NORMALISATION DES COLONNES
    # ========================================================

    df.rename(
        columns={
            "CODE AGENCE": "CODE_AGENCE",
            "LIBELLE AGENCE": "LIBELLE_AGENCE",
        },
        inplace=True,
    )


    # ========================================================
    # NORMALISATION DES MONTANTS
    # ========================================================

    if "DEBIT" in df.columns:
        df["DEBIT"] = pd.to_numeric(
            df["DEBIT"],
            errors="coerce",
        ).fillna(0)


    if "CREDIT" in df.columns:
        df["CREDIT"] = pd.to_numeric(
            df["CREDIT"],
            errors="coerce",
        ).fillna(0)


    # ========================================================
    # CODE AGENCE
    # ========================================================

    if "CODE_AGENCE" in df.columns:

        df["CODE_AGENCE"] = pd.to_numeric(
            df["CODE_AGENCE"],
            errors="coerce",
        ).astype("Int64")


    # ========================================================
    # SOLDE FLEX
    #
    # Même convention que le partenaire :
    #
    # DEBIT - CREDIT
    # ========================================================

    df["SOLDE_FLEX"] = (
        df.get("DEBIT", 0)
        - df.get("CREDIT", 0)
    )


    return df


# ============================================================
# API PRINCIPALE
# ============================================================

@app.get("/western-flex")
def export_western_flex(
    date_debut: str,
    date_fin: str,
    format: str = Query(
        "excel",
        description=(
            "excel (défaut) ou json "
            "(skip openpyxl, pour /charger)"
        ),
    ),
):

    df = get_western_flex(
        date_debut,
        date_fin,
    )


    nb_agences = 0

    if (
        "CODE_AGENCE" in df.columns
        and len(df)
    ):

        nb_agences = int(
            df["CODE_AGENCE"].nunique(
                dropna=True
            )
        )


    return save_single_flex(
        rt,
        df,
        sheet_name="WESTERN_FLEX",
        filename="WESTERN_FLEX.xlsx",
        format=format,
        headers={
            "X-Nb-Ecritures": str(len(df)),
            "X-Nb-Agences": str(nb_agences),
        },
    )


# ============================================================
# ROUTE SQLITE
# ============================================================

register_db_routes(
    rt,
    [
        (
            "/db/western",
            TABLE_WESTERN_FLEX,
        ),
    ],
)