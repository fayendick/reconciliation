from fastapi import HTTPException, Query

import pandas as pd
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    save_split_w2b_b2w,
    split_by_dc,
)


# ============================================================
# CONFIGURATION
# ============================================================

PARTENAIRE = "NIIRPAY"

rt = bootstrap_flex_service(
    PARTENAIRE,
    title="NiirPay Flex API",
    log_prefix="niirpay_flex_api",
)

app = rt.app
TABLES = rt.tables


# ============================================================
# EXTRACTION ORACLE NIIRPAY
# ============================================================
#
# Correspondances métier :
#
# NiirPay                         Oracle
# -------------------------------------------------------------
# niirPay Transaction ID    <->   GUTRANSACTIONID
# amount                    <->   AMOUNT
# fees                      <->   FRAIS
# Transaction Date          <->   TRANSACTIONDATE
# currency                  <->   CURRENCY
#
# Détermination du statut comptable :
#
# NiirPay.TXNREFNO
#       <-> ACVW_ALL_AC_ENTRIES.EXTERNAL_REF_NO
#
# avec :
#
# NiirPay.NUMEROCPTDEBITEUR
#       <-> ACVW_ALL_AC_ENTRIES.AC_NO
#
# ============================================================


def get_niirpay_flex(
    date_debut: str,
    date_fin: str,
) -> pd.DataFrame:

    sql = text(
        """
        WITH ecriture AS (
            SELECT
                A.AC_NO,
                A.TRN_REF_NO,
                A.EXTERNAL_REF_NO AS REF_NIIRPAY,
                A.LCY_AMOUNT,
                A.TRN_DT
            FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES A
            JOIN CFSFCUBS145.STTM_CUST_ACCOUNT CPT
                ON A.AC_NO = CPT.CUST_AC_NO
            JOIN CFSFCUBS145.STTM_ACCOUNT_CLASS CL
                ON CPT.ACCOUNT_CLASS = CL.ACCOUNT_CLASS
            WHERE CL.ACCOUNT_CODE IN ('251', '253')
        ),

        niirpay AS (
            SELECT
                NP.*,
                C.AC_DESC,
                C.CUST_AC_NO
            FROM OMB_SN.NIIRPAYTRANSACTION NP
            LEFT JOIN CFSFCUBS145.STTM_CUST_ACCOUNT C
                ON C.CUST_AC_NO = NP.NUMEROCPTDEBITEUR
        ),

        statut_transaction_niirpay AS (
            SELECT
                NI.GUTRANSACTIONID,
                E.LCY_AMOUNT,

                CASE
                    WHEN E.LCY_AMOUNT IS NOT NULL
                        THEN 'SUCCESS'
                    ELSE 'FAIL'
                END AS STATUT

            FROM niirpay NI

            LEFT JOIN ecriture E
                ON NI.TXNREFNO = E.REF_NIIRPAY
                AND E.AC_NO = NI.NUMEROCPTDEBITEUR
        )

        SELECT
            I.*,
            S.STATUT AS STATUT_TRANSACTION_NIIRPAY

        FROM niirpay I

        LEFT JOIN statut_transaction_niirpay S
            ON I.GUTRANSACTIONID = S.GUTRANSACTIONID

        WHERE I.DT_CRE_ENREG >= TO_DATE(:date_debut, 'DD/MM/YYYY')
          AND I.DT_CRE_ENREG < TO_DATE(:date_fin, 'DD/MM/YYYY')
        """
    )

    return oracle_query(
        rt,
        sql,
        {
            "date_debut": date_debut,
            "date_fin": date_fin,
        },
    )


# ============================================================
# PREPARATION DU DATAFRAME FLEX NIIRPAY
# ============================================================


def preparer_niirpay_flex(df: pd.DataFrame) -> pd.DataFrame:
    """
    Transforme les données Oracle NiirPay dans le schéma
    attendu par reconciliation_engine.py.

    Le moteur commun attend notamment :

        DATE_VALEUR
        CODE_TRANSACTION_OPERATEUR
        NUMERO_COMPTE
        MOUVEMENT_DEBIT
        MOUVEMENT_CREDIT
        TYPE_TRANSACTION

    IMPORTANT :

    Le Gateway actuel traite NiirPay dans le flux W2B.

    Le moteur W2B compare :
        Excel.MONTANT
        avec
        Flex.MOUVEMENT_CREDIT

    Nous plaçons donc le montant NiirPay dans
    MOUVEMENT_CREDIT pour rester compatible avec le
    flux NiirPay actuellement configuré.

    Cette adaptation est UNIQUEMENT dans le connecteur NiirPay.
    Aucun changement n'est fait dans le moteur commun.
    """

    df = df.copy()

    # --------------------------------------------------------
    # Montants
    # --------------------------------------------------------

    df["AMOUNT"] = pd.to_numeric(
        df["AMOUNT"],
        errors="coerce",
    )

    if "FRAIS" in df.columns:
        df["FRAIS"] = pd.to_numeric(
            df["FRAIS"],
            errors="coerce",
        )
    else:
        df["FRAIS"] = 0

    df["AMOUNT"] = df["AMOUNT"].fillna(0)
    df["FRAIS"] = df["FRAIS"].fillna(0)

    # --------------------------------------------------------
    # Dates
    # --------------------------------------------------------

    if "TRANSACTIONDATE" in df.columns:
        df["TRANSACTIONDATE"] = pd.to_datetime(
            df["TRANSACTIONDATE"],
            errors="coerce",
        )

    # DATE TRANSACTION : colonne utilisée pour l'audit
    df["DATE TRANSACTION"] = df["TRANSACTIONDATE"]

    # DATE_VALEUR : colonne utilisée par le moteur
    df["DATE_VALEUR"] = df["TRANSACTIONDATE"]

    # --------------------------------------------------------
    # Identifiant transaction
    # --------------------------------------------------------

    df["CODE TRANSACTION OPERATEUR"] = (
        df["GUTRANSACTIONID"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Le moteur utilise CODE_TRANSACTION_OPERATEUR
    # avant de le renommer en CODE_TRANSACTION.
    df["CODE_TRANSACTION_OPERATEUR"] = (
        df["CODE TRANSACTION OPERATEUR"]
    )

    # --------------------------------------------------------
    # Numéro compte
    # --------------------------------------------------------

    if "NUMEROCPTDEBITEUR" in df.columns:

        df["NUMERO COMPTE"] = (
            df["NUMEROCPTDEBITEUR"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    else:

        df["NUMERO COMPTE"] = ""

    # --------------------------------------------------------
    # Montant
    # --------------------------------------------------------

    df["MONTANT"] = df["AMOUNT"]

    # --------------------------------------------------------
    # Mouvements comptables attendus par le moteur
    # --------------------------------------------------------
    #
    # Pour W2B, reconciliation_engine.py utilise :
    #
    #     MOUVEMENT_CREDIT
    #
    # Le flux NiirPay est actuellement chargé dans :
    #
    #     COMPILATION_NIIRPAY_W2B
    #
    # On met donc AMOUNT dans MOUVEMENT_CREDIT.
    #
    # MOUVEMENT_DEBIT est conservé à 0.
    #
    # --------------------------------------------------------

    df["MOUVEMENT_CREDIT"] = df["AMOUNT"]

    df["MOUVEMENT_DEBIT"] = 0.0

    # --------------------------------------------------------
    # Type transaction
    # --------------------------------------------------------

    df["TYPE_TRANSACTION"] = "D"

    # --------------------------------------------------------
    # Montant avec frais
    # --------------------------------------------------------

    df["MONTANT_AVEC_FRAIS"] = (
        df["AMOUNT"]
        + df["FRAIS"]
    )

    # --------------------------------------------------------
    # Statut NiirPay
    # --------------------------------------------------------

    if "STATUT_TRANSACTION_NIIRPAY" in df.columns:

        df["STATUT_TRANSACTION_NIIRPAY"] = (
            df["STATUT_TRANSACTION_NIIRPAY"]
            .fillna("FAIL")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    else:

        df["STATUT_TRANSACTION_NIIRPAY"] = "FAIL"

    # --------------------------------------------------------
    # CURRENCY
    # --------------------------------------------------------

    if "CURRENCY" not in df.columns:
        df["CURRENCY"] = None

    # --------------------------------------------------------
    # Colonnes utiles pour audit
    # --------------------------------------------------------

    if "TXNREFNO" in df.columns:

        df["TXNREFNO"] = (
            df["TXNREFNO"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    return df


# ============================================================
# API PRINCIPALE
# ============================================================


@app.get("/niirpay-flex")
def export_niirpay_flex(
    date_debut: str,
    date_fin: str,
    format: str = Query(
        "excel",
        description="excel (défaut) ou json",
    ),
):

    # --------------------------------------------------------
    # Extraction Oracle
    # --------------------------------------------------------

    df = get_niirpay_flex(
        date_debut,
        date_fin,
    )

    print(
        f"[niirpay_flex_api] Extraction Oracle : "
        f"{len(df)} lignes"
    )

    # --------------------------------------------------------
    # Vérification des colonnes NiirPay
    # --------------------------------------------------------

    required_cols = {
        "GUTRANSACTIONID",
        "AMOUNT",
        "FRAIS",
        "TRANSACTIONDATE",
    }

    missing = required_cols - set(df.columns)

    if missing:

        raise HTTPException(
            status_code=500,
            detail=(
                f"Colonnes NiirPay manquantes : {missing}. "
                f"Colonnes disponibles : {list(df.columns)}"
            ),
        )

    # --------------------------------------------------------
    # Préparation standard pour le moteur
    # --------------------------------------------------------

    df = preparer_niirpay_flex(df)

    # --------------------------------------------------------
    # Séparation W2B / B2W
    # --------------------------------------------------------
    #
    # split_by_dc() regarde TYPE_TRANSACTION.
    #
    # NiirPay est actuellement D.
    #
    # On conserve donc la logique existante.
    #
    # --------------------------------------------------------

    debit_df, credit_df = split_by_dc(df)

    print(
        "--------------------------------------------------"
    )

    print(
        f"[niirpay_flex_api] Nombre total : {len(df)}"
    )

    print(
        f"[niirpay_flex_api] Nombre débit : "
        f"{len(debit_df)}"
    )

    print(
        f"[niirpay_flex_api] Nombre crédit : "
        f"{len(credit_df)}"
    )

    print(
        "--------------------------------------------------"
    )

    # --------------------------------------------------------
    # IMPORTANT
    # --------------------------------------------------------
    #
    # Le Gateway NiirPay attend :
    #
    #     flex_w2b
    #     flex_b2w
    #
    # Les tables sont créées par save_split_w2b_b2w()
    # à partir du DataFrame et des feuilles ci-dessous.
    #
    # --------------------------------------------------------

    return save_split_w2b_b2w(
        rt,
        df,
        sheets={
            "NIIRPAY_FLEX": df,
            "NIIRPAY_FLEX_DEBIT": debit_df,
            "NIIRPAY_FLEX_CREDIT": credit_df,
        },
        filename="NIIRPAY_FLEX.xlsx",
        format=format,
        mode="dc",
    )