from fastapi import HTTPException, Query

import pandas as pd
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    split_by_dc,
)

from common.sqlite_io import (
    ecrire_table,
    creer_vues_split,
)

from common.http_export import respond_sheets


# ============================================================
# CONFIGURATION
# ============================================================

PARTENAIRE = "GTP"

rt = bootstrap_flex_service(
    PARTENAIRE,
    title="GTP Flex API",
    log_prefix="gtp_flex_api",
)

app = rt.app
TABLES = rt.tables


# ============================================================
# EXTRACTION ORACLE GTP
# ============================================================
#
# GTP API
#
# Correspondances :
#
# Partner Reference
#       <-> numéro de carte extrait de DESC_OPS
#
# Partner Amount
#       <-> MONTANT
#
# Partner Date
#       <-> SAVE_TIMESTAMP
#
# Informations conservées pour audit :
#
# AC_ENTRY_SR_NO
# NUMCPT
# TXNREFNO
# DRCR_IND
# TRN_DT
# SAVE_TIMESTAMP
# DESC_OPS
# FRAIS
# STATUT
#
# ============================================================


def get_gtp_flex(
    date_debut: str,
    date_fin: str,
) -> pd.DataFrame:

    sql = text(
        """
        /* =======================
           COFICARTE RECHARGEMENT
           ======================= */

        SELECT
            a.AC_ENTRY_SR_NO,
            i.MONTANT AS MONTANT,
            a.AC_NO AS NUMCPT,
            i.TXNREFNO,
            a.DRCR_IND,
            a.TRN_DT,
            a.SAVE_TIMESTAMP,

            (
                SELECT
                    'RECHARMENT E-COFICARTE Numéro  ' || IDENTIFIANTCARTE
                FROM CFSFCUBS145.STTM_CUST_ACCOUNT
                WHERE CUST_AC_NO = i.COMPTE
            ) AS DESC_OPS,

            i.FRAIS AS FRAIS,

            CASE
                WHEN a.LCY_AMOUNT IS NOT NULL
                    THEN 'SUCCESS'
                ELSE 'FAIL'
            END AS STATUT

        FROM OMB_SN.GTPTRANSACTION i,
             CFSFCUBS145.ACVW_ALL_AC_ENTRIES a

        WHERE i.TXNREFNO = a.EXTERNAL_REF_NO
          AND a.AC_NO = i.COMPTE

          AND a.TRN_DT >= TO_DATE(
                :date_debut,
                'DD/MM/YYYY'
              )

          AND a.TRN_DT < TO_DATE(
                :date_fin,
                'DD/MM/YYYY'
              ) + 1
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
# EXTRACTION DU NUMÉRO DE CARTE
# ============================================================


def extraire_numero_carte(value) -> str:
    """
    Extrait le numéro de carte depuis DESC_OPS.

    Exemple :

        RECHARMENT E-COFICARTE Numéro 0022666431

    devient :

        0022666431
    """

    if pd.isna(value):
        return ""

    texte = str(value).strip()

    if not texte:
        return ""

    # On récupère le dernier groupe numérique.
    import re

    numeros = re.findall(r"\d+", texte)

    if not numeros:
        return ""

    return numeros[-1].strip()


# ============================================================
# PREPARATION DU DATAFRAME FLEX GTP
# ============================================================


def preparer_gtp_flex(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prépare les données GTP.

    IMPORTANT :

    Le moteur commun n'est pas modifié.

    On conserve également les colonnes natives GTP
    nécessaires au rapprochement spécifique :
        - CARD_NUMBER
        - SAVE_TIMESTAMP
        - MONTANT
        - TXNREFNO
        - AC_ENTRY_SR_NO
        - FRAIS
        - STATUT
    """

    df = df.copy()

    # --------------------------------------------------------
    # Vérification minimale
    # --------------------------------------------------------

    required_cols = {
        "AC_ENTRY_SR_NO",
        "MONTANT",
        "NUMCPT",
        "TXNREFNO",
        "DRCR_IND",
        "TRN_DT",
        "SAVE_TIMESTAMP",
        "DESC_OPS",
        "FRAIS",
        "STATUT",
    }

    missing = required_cols - set(df.columns)

    if missing:
        raise HTTPException(
            status_code=500,
            detail=(
                f"Colonnes GTP manquantes : {sorted(missing)}. "
                f"Colonnes disponibles : {list(df.columns)}"
            ),
        )

    # --------------------------------------------------------
    # Montant
    # --------------------------------------------------------

    df["MONTANT"] = pd.to_numeric(
        df["MONTANT"],
        errors="coerce",
    )

    df["MONTANT"] = df["MONTANT"].abs()

    # --------------------------------------------------------
    # Frais
    # --------------------------------------------------------

    df["FRAIS"] = pd.to_numeric(
        df["FRAIS"],
        errors="coerce",
    ).fillna(0)

    # --------------------------------------------------------
    # Dates
    # --------------------------------------------------------

    df["TRN_DT"] = pd.to_datetime(
    df["TRN_DT"],
    errors="coerce",
    )

    df["SAVE_TIMESTAMP"] = (
        pd.to_datetime(
            df["SAVE_TIMESTAMP"],
            errors="coerce",
        )
        - pd.Timedelta(hours=5)
    )

    # --------------------------------------------------------
    # Numéro de carte
    # --------------------------------------------------------

    df["CARD_NUMBER"] = (
        df["DESC_OPS"]
        .apply(extraire_numero_carte)
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Référence transaction
    # --------------------------------------------------------

    df["TXNREFNO"] = (
        df["TXNREFNO"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Compte
    # --------------------------------------------------------

    df["NUMCPT"] = (
        df["NUMCPT"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Statut
    # --------------------------------------------------------

    df["STATUT"] = (
        df["STATUT"]
        .fillna("FAIL")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # --------------------------------------------------------
    # Colonnes compatibles moteur commun
    # --------------------------------------------------------
    #
    # Elles sont conservées pour ne pas casser les endpoints
    # génériques existants.
    #
    # Le rapprochement GTP spécifique sera traité séparément.
    # --------------------------------------------------------

    df["DATE TRANSACTION"] = df["SAVE_TIMESTAMP"]

    df["DATE_VALEUR"] = df["SAVE_TIMESTAMP"]

    df["CODE TRANSACTION OPERATEUR"] = df["TXNREFNO"]

    df["NUMERO COMPTE"] = df["CARD_NUMBER"]

    df["MONTANT_AVEC_FRAIS"] = (
        df["MONTANT"] + df["FRAIS"]
    )

    # --------------------------------------------------------
    # Type transaction
    # --------------------------------------------------------

    df["TYPE_TRANSACTION"] = (
        df["DRCR_IND"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # --------------------------------------------------------
    # Mouvements
    # --------------------------------------------------------

    df["MOUVEMENT_DEBIT"] = 0.0
    df["MOUVEMENT_CREDIT"] = 0.0

    debit_mask = (
        df["TYPE_TRANSACTION"] == "D"
    )

    credit_mask = (
        df["TYPE_TRANSACTION"] == "C"
    )

    df.loc[
        debit_mask,
        "MOUVEMENT_DEBIT"
    ] = df.loc[
        debit_mask,
        "MONTANT"
    ]

    df.loc[
        credit_mask,
        "MOUVEMENT_CREDIT"
    ] = df.loc[
        credit_mask,
        "MONTANT"
    ]

    return df


# ============================================================
# SAUVEGARDE FLEX GTP
# ============================================================


def sauvegarder_gtp_flex(
    df: pd.DataFrame,
    debit_df: pd.DataFrame,
    credit_df: pd.DataFrame,
    *,
    format: str = "excel",
):
    """
    Sauvegarde le Flex GTP.

    Contrairement à certains partenaires existants, nous conservons
    ici les deux sens D/C sans modifier leur signification.
    """

    try:

        # ----------------------------------------------------
        # Table complète
        # ----------------------------------------------------

        ecrire_table(
            df,
            TABLES["flex"],
            rt.sqlite,
            log_prefix=rt.log_prefix,
        )

        # ----------------------------------------------------
        # Vues W2B / B2W
        # ----------------------------------------------------

        creer_vues_split(
            rt.sqlite,
            TABLES["flex"],
            TABLES["flex_w2b"],
            TABLES["flex_b2w"],
            'UPPER(TRIM(CAST("TYPE_TRANSACTION" AS TEXT))) = \'C\'',
            'UPPER(TRIM(CAST("TYPE_TRANSACTION" AS TEXT))) = \'D\'',
            log_prefix=rt.log_prefix,
        )

        # ----------------------------------------------------
        # Export
        # ----------------------------------------------------

        return respond_sheets(
            {
                "GTP_FLEX": df,
                "GTP_FLEX_DEBIT": debit_df,
                "GTP_FLEX_CREDIT": credit_df,
            },
            filename="GTP_FLEX.xlsx",
            format=format,
        )

    except HTTPException:
        raise

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Erreur sauvegarde/export Flex GTP : {e}",
        )


# ============================================================
# HEALTH CHECK
# ============================================================


@app.get("/health")
def health():
    return {
        "status": "ok",
        "partenaire": PARTENAIRE,
        "service": "gtp-flex",
    }


# ============================================================
# API PRINCIPALE
# ============================================================


@app.get("/gtp-flex")
def export_gtp_flex(
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

    df = get_gtp_flex(
        date_debut,
        date_fin,
    )

    print(
        f"[gtp_flex_api] Extraction Oracle : "
        f"{len(df)} lignes"
    )

    # --------------------------------------------------------
    # Préparation
    # --------------------------------------------------------

    df = preparer_gtp_flex(df)

    # --------------------------------------------------------
    # Séparation D / C
    # --------------------------------------------------------

    debit_df, credit_df = split_by_dc(df)

    print(
        "--------------------------------------------------"
    )

    print(
        f"[gtp_flex_api] Nombre total : {len(df)}"
    )

    print(
        f"[gtp_flex_api] Nombre débit : "
        f"{len(debit_df)}"
    )

    print(
        f"[gtp_flex_api] Nombre crédit : "
        f"{len(credit_df)}"
    )

    print(
        f"[gtp_flex_api] Cartes identifiées : "
        f"{(df['CARD_NUMBER'] != '').sum()}"
    )

    print(
        "--------------------------------------------------"
    )

    # --------------------------------------------------------
    # Sauvegarde
    # --------------------------------------------------------

    return sauvegarder_gtp_flex(
        df,
        debit_df,
        credit_df,
        format=format,
    )