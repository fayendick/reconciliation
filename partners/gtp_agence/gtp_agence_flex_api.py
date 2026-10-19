from fastapi import FastAPI, HTTPException

import pandas as pd
from sqlalchemy import text

from common.config import make_oracle_engine
from common.config import make_sqlite_engine
from common.sqlite_io import ecrire_table


app = FastAPI(
    title="GTP Agence FLEX",
    version="1.0.0",
)

oracle_engine = make_oracle_engine()
sqlite_engine = make_sqlite_engine()


TABLE_FLEX = "GTP_AGENCE_FLEX"
TABLE_FLEX_W2B = "GTP_AGENCE_FLEX_W2B"
TABLE_FLEX_B2W = "GTP_AGENCE_FLEX_B2W"


def preparer_gtp_agence_flex(
    df: pd.DataFrame,
) -> pd.DataFrame:

    df = df.copy()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.lower()
    )

    df["lcy_amount"] = pd.to_numeric(
        df["lcy_amount"],
        errors="coerce",
    )

    df["trn_dt"] = pd.to_datetime(
        df["trn_dt"],
        errors="coerce",
    )

    df["save_timestamp"] = pd.to_datetime(
        df["save_timestamp"],
        errors="coerce",
    )

    df["ac_no"] = (
        df["ac_no"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["module"] = (
        df["module"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df["drcr_ind"] = (
        df["drcr_ind"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df["trn_ref_no"] = (
        df["trn_ref_no"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # =========================================================
    # CORRECTION DU FUSEAU HORAIRE
    # =========================================================
    #
    # FLEXCUBE est décalé de +5 heures par rapport au fichier
    # GTP Agence.
    #
    # On corrige uniquement SAVE_TIMESTAMP.
    #
    # Exemple :
    #
    # FLEXCUBE : 13:03:06
    # GTP      : 08:03:06
    #
    # =========================================================

    df["save_timestamp"] = (
        df["save_timestamp"]
        - pd.Timedelta(hours=5)
    )

    # =========================================================
    # COLONNES STANDARD
    # =========================================================

    df["DATE TRANSACTION"] = df["save_timestamp"]

    df["DATE_VALEUR"] = df["save_timestamp"]

    df["CODE_TRANSACTION"] = df["trn_ref_no"]

    df["NUMERO COMPTE"] = df["ac_no"]

    df["MONTANT"] = df["lcy_amount"].abs()

    df["MONTANT_AVEC_FRAIS"] = df["MONTANT"]

    df["TYPE TRANSACTION"] = df["drcr_ind"]

    df["MOUVEMENT_DEBIT"] = 0.0
    df["MOUVEMENT_CREDIT"] = 0.0

    mask_debit = df["drcr_ind"] == "D"
    mask_credit = df["drcr_ind"] == "C"

    df.loc[
        mask_debit,
        "MOUVEMENT_DEBIT",
    ] = df.loc[
        mask_debit,
        "MONTANT",
    ]

    df.loc[
        mask_credit,
        "MOUVEMENT_CREDIT",
    ] = df.loc[
        mask_credit,
        "MONTANT",
    ]

    return df


def creer_vues_split(
    df: pd.DataFrame,
):
    # GTP Agence :
    # C = crédit / W2B
    # D = débit / B2W

    df_w2b = df[
        df["TYPE TRANSACTION"] == "C"
    ].copy()

    df_b2w = df[
        df["TYPE TRANSACTION"] == "D"
    ].copy()

    return df_w2b, df_b2w


def extraire_gtp_agence_flex(
    date_debut: str,
    date_fin: str,
) -> pd.DataFrame:

    query = text("""
        SELECT
            a.AC_NO,
            a.MODULE,
            a.DRCR_IND,
            a.LCY_AMOUNT,
            a.TRN_DT,
            a.SAVE_TIMESTAMP,
            a.TRN_REF_NO
        FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES a
        WHERE
            a.TRN_DT >= TO_DATE(
                :date_debut,
                'DD/MM/YYYY'
            )
            AND a.TRN_DT < TO_DATE(
                :date_fin,
                'DD/MM/YYYY'
            ) + 1
            AND a.AC_NO  = '331600000011'
    """)

    params = {
        "date_debut": date_debut,
        "date_fin": date_fin,
    }

    with oracle_engine.connect() as conn:
        df = pd.read_sql(
            query,
            conn,
            params=params,
        )

    return preparer_gtp_agence_flex(df)


def sauvegarder_gtp_agence_flex(
    df: pd.DataFrame,
):
    df_w2b, df_b2w = creer_vues_split(df)

    ecrire_table(
        df,
        TABLE_FLEX,
        sqlite_engine,
        log_prefix="gtp_agence_flex",
    )

    ecrire_table(
        df_w2b,
        TABLE_FLEX_W2B,
        sqlite_engine,
        log_prefix="gtp_agence_flex_w2b",
    )

    ecrire_table(
        df_b2w,
        TABLE_FLEX_B2W,
        sqlite_engine,
        log_prefix="gtp_agence_flex_b2w",
    )

    return df_w2b, df_b2w


@app.get("/health")
def health():
    return {
        "status": "ok",
        "partenaire": "GTP_AGENCE",
        "service": "gtp-agence-flex",
        "table": TABLE_FLEX,
    }


@app.get("/gtp-agence-flex")
def export_gtp_agence_flex(
    date_debut: str,
    date_fin: str,
    format: str = "json",
):
    try:
        df = extraire_gtp_agence_flex(
            date_debut,
            date_fin,
        )

        df_w2b, df_b2w = sauvegarder_gtp_agence_flex(
            df
        )

        return {
            "status": "ok",
            "partenaire": "GTP_AGENCE",
            "format": format,
            "date_debut": date_debut,
            "date_fin": date_fin,
            "total_rows": len(df),
            "w2b_rows": len(df_w2b),
            "b2w_rows": len(df_b2w),
            "tables": {
                "flex": TABLE_FLEX,
                "flex_w2b": TABLE_FLEX_W2B,
                "flex_b2w": TABLE_FLEX_B2W,
            },
        }

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=(
                "Erreur extraction GTP Agence FLEX : "
                f"{exc}"
            ),
        )