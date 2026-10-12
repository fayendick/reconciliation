import io

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile

from common.config import make_sqlite_engine
from common.sqlite_io import ecrire_table


app = FastAPI(
    title="GTP Agence",
    version="1.0.0",
)

engine = make_sqlite_engine()

TABLE_EXCEL = "COMPILATION_GTP_AGENCE"


def preparer_gtp_agence(
    contenu: bytes,
    filename: str,
) -> pd.DataFrame:

    try:
        df = pd.read_csv(
            io.BytesIO(contenu),
            sep=",",
            encoding="latin1",
        )
    except Exception as exc:
        raise ValueError(
            f"Impossible de lire le fichier {filename} : {exc}"
        )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    colonnes_obligatoires = [
        "Trans ID",
        "Date",
        "Amount",
        "Fees",
        "Description",
        "Reference",
        "Account Id",
        "Last Name",
    ]

    manquantes = [
        col
        for col in colonnes_obligatoires
        if col not in df.columns
    ]

    if manquantes:
        raise ValueError(
            "Colonnes manquantes dans le fichier GTP Agence : "
            + ", ".join(manquantes)
        )

    # ---------------------------------------------------------
    # Colonnes GTP Agence
    # ---------------------------------------------------------

    df["WP_TRANS_ID"] = (
        df["Trans ID"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["WP_DATE_HEURE"] = pd.to_datetime(
        df["Date"],
        format="%d-%m-%y %H:%M:%S",
        errors="coerce",
    )

    df["WP_MONTANT"] = pd.to_numeric(
        df["Amount"],
        errors="coerce",
    ).abs()

    df["WP_FRAIS"] = pd.to_numeric(
        df["Fees"],
        errors="coerce",
    ).fillna(0)

    df["WP_DESCRIPTION"] = (
        df["Description"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["WP_REFERENCE"] = (
        df["Reference"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["WP_ACCOUNT_ID"] = (
        df["Account Id"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["WP_LAST_NAME"] = (
        df["Last Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ---------------------------------------------------------
    # GTP AGENCE
    # ---------------------------------------------------------

    mask_gtp_agence = (
        df["WP_DESCRIPTION"]
        .str.upper()
        .str.contains(
            "RECHARGE DE LA CARTE",
            na=False,
        )
    )

    # IMPORTANT :
    # On ne supprime aucune ligne du fichier.
    # On garde uniquement les opérations GTP Agence
    # pour la table utilisée par ce partenaire.
    df = df.loc[
        mask_gtp_agence
    ].copy()

    df["STATUT_RECONCILIATION"] = "A RAPPROCHER"

    # ---------------------------------------------------------
    # Colonnes standard du moteur commun
    # ---------------------------------------------------------

    df["DATE TRANSACTION"] = df["WP_DATE_HEURE"]

    df["DATE_VALEUR"] = df["WP_DATE_HEURE"]

    # Les opérations GTP Agence correspondent au crédit
    # dans FLEXCUBE.
    df["TYPE TRANSACTION"] = "C"

    df["CODE_TRANSACTION"] = df["WP_TRANS_ID"]

    df["NUMERO COMPTE"] = df["WP_ACCOUNT_ID"]

    df["MONTANT"] = df["WP_MONTANT"]

    df["MONTANT_AVEC_FRAIS"] = (
        df["WP_MONTANT"] + df["WP_FRAIS"]
    )

    df["MOUVEMENT_DEBIT"] = 0.0

    df["MOUVEMENT_CREDIT"] = df["WP_MONTANT"]

    return df

def sauvegarder_gtp_agence(
    df: pd.DataFrame,
):
    TABLE_EXCEL_W2B = "COMPILATION_GTP_AGENCE_W2B"
    TABLE_EXCEL_B2W = "COMPILATION_GTP_AGENCE_B2W"

    # 1. Sauvegarder le fichier complet
    ecrire_table(
        df,
        TABLE_EXCEL,
        engine,
        log_prefix="gtp_agence_excel",
    )

    # 2. Normaliser le type de transaction
    type_transaction = (
        df["TYPE TRANSACTION"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # 3. W2B = Crédit (C)
    df_w2b = df[type_transaction == "C"].copy()

    # 4. B2W = Débit (D)
    df_b2w = df[type_transaction == "D"].copy()

    # 5. Sauvegarder W2B
    ecrire_table(
        df_w2b,
        TABLE_EXCEL_W2B,
        engine,
        log_prefix="gtp_agence_excel_w2b",
    )

    # 6. Sauvegarder B2W
    ecrire_table(
        df_b2w,
        TABLE_EXCEL_B2W,
        engine,
        log_prefix="gtp_agence_excel_b2w",
    )

    return df

@app.get("/health")
def health():
    return {
        "status": "ok",
        "partenaire": "GTP_AGENCE",
        "service": "gtp-agence-excel",
        "table": TABLE_EXCEL,
    }


@app.post("/process-excel")
async def process_excel(
    files: list[UploadFile] = File(...),
    format: str = "json",
):
    if not files:
        raise HTTPException(
            status_code=400,
            detail="Aucun fichier GTP Agence fourni.",
        )

    file = files[0]

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="Nom de fichier absent.",
        )

    try:
        contenu = await file.read()

        df = preparer_gtp_agence(
            contenu,
            file.filename,
        )

        sauvegarder_gtp_agence(df)

        return {
            "status": "ok",
            "partenaire": "GTP_AGENCE",
            "filename": file.filename,
            "total_rows": len(df),
            "date_min": (
                df["WP_DATE_HEURE"].min().isoformat()
                if not df.empty
                and df["WP_DATE_HEURE"].notna().any()
                else None
            ),
            "date_max": (
                df["WP_DATE_HEURE"].max().isoformat()
                if not df.empty
                and df["WP_DATE_HEURE"].notna().any()
                else None
            ),
            "columns": list(df.columns),
        }

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Erreur Upload GTP Agence : {exc}",
        )