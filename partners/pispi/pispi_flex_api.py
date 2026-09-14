# ============================================================
# PI/SPI — SERVICE FLEX ORACLE
# ============================================================

from fastapi import HTTPException, Query

import pandas as pd
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    register_db_routes,
)
from common.sqlite_io import ecrire_table, creer_vues_split
from common.http_export import respond_sheets


# ============================================================
# CONFIGURATION
# ============================================================

PARTENAIRE = "PISPI"

rt = bootstrap_flex_service(
    PARTENAIRE,
    title="PI/SPI Flex API",
    log_prefix="pispi_flex_api",
)

app = rt.app
TABLES = rt.tables


# ============================================================
# EXTRACTION ORACLE PI/SPI
# ============================================================

def get_pispi_flex(
    date_debut: str,
    date_fin: str,
    compte: str,
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

            /* ==================================================
               REFERENCE
               ================================================== */

            o.REFERENCETRANSACTION
                AS "REFERENCETRANSACTION",

            a.TRN_REF_NO
                AS "TRN_REF_NO",


            /* ==================================================
               NUMERO ERC
               ================================================== */

            a.AC_ENTRY_SR_NO
                AS "NO_ERC",


            /* ==================================================
               GL / DESCRIPTION
               ================================================== */

            NVL(
                c.GL_CODE,
                s.DR_GL
            ) AS "PARENT_GL",

            NVL(
                c.GL_DESC,
                s.AC_DESC
            ) AS "DESCRIPTION",


            /* ==================================================
               AGENCE
               ================================================== */

            a.AC_BRANCH
                AS "CODE AGENCE",

            b.BRANCH_NAME
                AS "LIBELLE AGENCE",


            /* ==================================================
               DEBIT
               ================================================== */

            DECODE(
                a.DRCR_IND,
                'D',
                a.LCY_AMOUNT,
                0
            ) AS "Debit",


            /* ==================================================
               CREDIT
               ================================================== */

            DECODE(
                a.DRCR_IND,
                'C',
                a.LCY_AMOUNT,
                0
            ) AS "Credit",


            /* ==================================================
               SENS
               ================================================== */

            a.DRCR_IND
                AS "SENS",


            /* ==================================================
               STATUT PI
               ================================================== */

            o.STATUSFINAL
                AS "STATUSFINAL",


            /* ==================================================
               DATES
               ================================================== */

            a.TRN_DT
                AS "DATE_SAISIE",

            a.VALUE_DT
                AS "DATE_VALEUR",


            /* ==================================================
               UTILISATEURS
               ================================================== */

            a.USER_ID
                AS "UTIL SAISI",

            a.AC_NO
                AS "ACCOUNT_NO",

            a.AUTH_ID
                AS "UTIL VALID",


            /* ==================================================
               INFORMATIONS TRANSACTION
               ================================================== */

            a.BATCH_NO,

            a.TRN_CODE,


            /* ==================================================
               LIBELLE OPERATION
               ================================================== */

            NVL(
                (
                    SELECT
                        u.ADDL_TEXT

                    FROM CFSFCUBS145.DETB_UPLOAD_DETAIL u

                    WHERE
                        u.BATCH_NO = a.BATCH_NO

                        AND u.ACCOUNT = a.AC_NO

                        AND u.VALUE_DATE = a.VALUE_DT

                        AND u.AMOUNT = a.LCY_AMOUNT

                        AND a.CURR_NO = u.CURR_NO
                ),
                t.TRN_DESC
            ) AS "LIBELLE_OPER",


            /* ==================================================
               DESCRIPTION BATCH
               ================================================== */

            NVL(
                od.ADDL_TEXT,
                NVL(
                    xx.ADDL_TEXT,
                    d.DESCRIPTION
                )
            ) AS "DESCRIPTION BATCH",


            /* ==================================================
               INFORMATIONS CLIENT
               ================================================== */

            a.RELATED_CUSTOMER
                AS "MATRICULE_CLIENT",

            a.RELATED_CUSTOMER
                AS "MATRICULE_CLIENT_ORACLE",


            /* ==================================================
               COMPTE CLIENT
               ================================================== */

            a.RELATED_ACCOUNT
                AS "COMPTE_CLIENT",

            a.RELATED_ACCOUNT
                AS "COMPTE_ASSOCIES",


            /* ==================================================
               COMPTE NAFA
               ================================================== */

            s.ALT_AC_NO
                AS "COMPTE_NAFA",

            s.ALT_AC_NO
                AS "ACCOUNT_NAFA",


            /* ==================================================
               NOM / PRENOM CLIENT
               ================================================== */

            NVL(
                c.GL_DESC,
                s.AC_DESC
            ) AS "NOM_PRENOM_CLIENT",


            /* ==================================================
               AUTRES INFORMATIONS
               ================================================== */

            a.EVENT
                AS "EVENEMENT",

            a.AMOUNT_TAG
                AS "ETIQUETTE",

            a.SAVE_TIMESTAMP

        FROM Journal a


        /* ======================================================
           PITRANSACTION
           ====================================================== */

        LEFT JOIN (

            SELECT
                NCPDEBITEUR,
                REFERENCETRANSACTION,
                STATUSFINAL

            FROM (

                SELECT
                    p.NCPDEBITEUR,
                    p.REFERENCETRANSACTION,
                    p.STATUSFINAL,

                    ROW_NUMBER() OVER (
                        PARTITION BY p.NCPDEBITEUR
                        ORDER BY p.DATETRANS DESC
                    ) AS rn

                FROM OMB_SN.PITRANSACTION p

                WHERE p.STATUSFINAL = 'SUCCESSFUL'
            )

            WHERE rn = 1

        ) o

            ON o.NCPDEBITEUR = a.AC_NO


        /* ======================================================
           GL MASTER
           ====================================================== */

        LEFT JOIN CFSFCUBS145.GLTM_GLMASTER c

            ON c.GL_CODE = a.AC_NO


        /* ======================================================
           COMPTE CLIENT
           ====================================================== */

        LEFT JOIN CFSFCUBS145.STTM_CUST_ACCOUNT s

            ON s.CUST_AC_NO = a.AC_NO


        /* ======================================================
           CODE TRANSACTION
           ====================================================== */

        LEFT JOIN CFSFCUBS145.STTM_TRN_CODE t

            ON t.TRN_CODE = a.TRN_CODE


        /* ======================================================
           AGENCE
           ====================================================== */

        LEFT JOIN CFSFCUBS145.STTM_BRANCH b

            ON b.BRANCH_CODE = a.AC_BRANCH


        /* ======================================================
           DETAIL JOURNAL
           ====================================================== */

        LEFT JOIN CFSFCUBS145.DETBS_JRNL_TXN_DETAIL xx

            ON a.TRN_REF_NO = xx.REFERENCE_NO

            AND a.EVENT_SR_NO = xx.SERIAL_NO


        /* ======================================================
           BATCH
           ====================================================== */

        LEFT JOIN CFSFCUBS145.DETB_BATCH_MASTER d

            ON a.BATCH_NO = d.BATCH_NO

            AND a.AC_BRANCH = d.BRANCH_CODE


        /* ======================================================
           UPLOAD DETAIL
           ====================================================== */

        LEFT JOIN CFSFCUBS145.DETB_UPLOAD_DETAIL od

            ON od.BATCH_NO = a.BATCH_NO

            AND a.AC_BRANCH = od.ACCOUNT_BRANCH

            AND a.CURR_NO = od.CURR_NO


        /* ======================================================
           FILTRE COMPTE SUSPENSE
           ====================================================== */

        WHERE A.TRN_REF_NO IN (

            SELECT
                TRN_REF_NO

            FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES

            WHERE AC_NO = :compte
        )


        /* ======================================================
           DATE DEBUT INCLUSE
           ====================================================== */

        AND a.SAVE_TIMESTAMP >= :date_debut


        /* ======================================================
           DATE FIN EXCLUE
           ====================================================== */

        AND a.SAVE_TIMESTAMP < :date_fin


        /* ======================================================
           UNIQUEMENT PI SUCCESSFUL
           ====================================================== */

        AND o.STATUSFINAL = 'SUCCESSFUL'


        ORDER BY
            a.VALUE_DT DESC
    """)

    # ========================================================
# FENÊTRE MÉTIER PI/SPI
# 09:00 du jour de début
# jusqu'à 09:00 du jour de fin
# ========================================================

    date_debut = (
        pd.Timestamp(date_debut)
        .normalize()
        + pd.Timedelta(hours=9)
    )

    date_fin = (
        pd.Timestamp(date_fin)
        .normalize()
        + pd.Timedelta(hours=9)
    )

    return oracle_query(
        rt,
        sql,
        {
            "compte": compte,
            "date_debut": date_debut.to_pydatetime(),
            "date_fin": date_fin.to_pydatetime(),
        },
    )

    if df is None or df.empty:
        return df

    # ========================================================
    # COLONNES STANDARD POUR LE MOTEUR DE RECONCILIATION
    # ========================================================

    df["CODE_TRANSACTION_OPERATEUR"] = df["TRN_REF_NO"]
    df["CODE_TRANSACTION"] = df["TRN_REF_NO"]
    df["NUMERO_COMPTE"] = df["ACCOUNT_NO"]
    df["DATE_VALEUR"] = df["SAVE_TIMESTAMP"]

    df["MOUVEMENT_DEBIT"] = pd.to_numeric(
        df["DEBIT"],
        errors="coerce",
    ).fillna(0)

    df["MOUVEMENT_CREDIT"] = pd.to_numeric(
        df["CREDIT"],
        errors="coerce",
    ).fillna(0)

    return df


# ============================================================
# API PRINCIPALE
# ============================================================

@app.get("/pispi-flex")
def export_pispi_flex(
    date_debut: str,
    date_fin: str,
    compte: str = Query(
        "114100000052",
        description="Compte suspense PI/SPI",
    ),
    format: str = Query(
        "excel",
        description="excel ou json",
    ),
):

    df = get_pispi_flex(
        date_debut,
        date_fin,
        compte,
    )

    if df is None or df.empty:
        raise HTTPException(
            status_code=404,
            detail="Aucune transaction PI/SPI trouvée.",
        )

    if "SENS" not in df.columns:
        raise HTTPException(
            status_code=500,
            detail=(
                "Colonne SENS absente. "
                f"Colonnes disponibles : {list(df.columns)}"
            ),
        )

    # ========================================================
    # NORMALISATION DU SENS
    # ========================================================

    df["SENS"] = (
        df["SENS"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ========================================================
    # PI/SPI :
    #
    # CREDIT → W2B
    # DEBIT  → B2W
    # ========================================================

    w2b = df[
        df["SENS"] == "C"
    ].copy()

    b2w = df[
        df["SENS"] == "D"
    ].copy()

    # ========================================================
    # TYPE TRANSACTION
    # ========================================================
    # Le moteur de réconciliation travaille avec :
    # CREDIT → W2B
    # DEBIT  → B2W

    df["TYPE_TRANSACTION"] = df["SENS"].map(
        {
            "C": "W2B",
            "D": "B2W",
        }
    )
    
    # ========================================================
    # MOUVEMENTS COMPTABLES
    # ========================================================
    # PI/SPI :
    # C = Crédit -> montant dans MOUVEMENT_CREDIT
    # D = Débit  -> montant dans MOUVEMENT_DEBIT

    df["MOUVEMENT_CREDIT"] = pd.to_numeric(
        df["CREDIT"],
        errors="coerce",
    ).fillna(0)

    df["MOUVEMENT_DEBIT"] = pd.to_numeric(
        df["DEBIT"],
        errors="coerce",
    ).fillna(0)

    w2b["TYPE_TRANSACTION"] = "W2B"
    b2w["TYPE_TRANSACTION"] = "B2W"

    # ========================================================
    # SAUVEGARDE
    # ========================================================

    from common.flex_common import save_split_w2b_b2w

    # Sauvegarde de la table complète
    ecrire_table(
        df,
        TABLES["flex"],
        rt.sqlite,
        log_prefix=rt.log_prefix,
    )

    # PI/SPI :
    # C = Crédit = W2B
    # D = Débit   = B2W
    creer_vues_split(
        rt.sqlite,
        TABLES["flex"],
        TABLES["flex_w2b"],
        TABLES["flex_b2w"],
        "UPPER(TRIM(CAST(TYPE_TRANSACTION AS TEXT))) = 'W2B'",
        "UPPER(TRIM(CAST(TYPE_TRANSACTION AS TEXT))) = 'B2W'",
        log_prefix=rt.log_prefix,
    )

    return respond_sheets(
        {
            "PISPI_FLEX": df,
            "PISPI_FLEX_W2B": w2b,
            "PISPI_FLEX_B2W": b2w,
        },
        filename="PISPI_FLEX.xlsx",
        format=format,
    )


# ============================================================
# ROUTES SQLITE
# ============================================================

register_db_routes(
    rt,
    [
        ("/db/pispi", TABLES["flex"]),
        ("/db/pispi-w2b", TABLES["flex_w2b"]),
        ("/db/pispi-b2w", TABLES["flex_b2w"]),
    ],
)
