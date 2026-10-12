from io import BytesIO

import pandas as pd

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import JSONResponse

from common.sqlite_io import ecrire_table
from common.http_export import respond_sheets
from common.config import get_partenaire, make_sqlite_engine


# ============================================================
# CONFIGURATION
# ============================================================

PARTENAIRE = "GTP"

CONFIG = get_partenaire(PARTENAIRE)

# Moteur SQLite commun à l'application
engine = make_sqlite_engine()

app = FastAPI(
    title="GTP Excel API",
    version="1.0.0",
)


# ============================================================
# TABLES SQLITE
# ============================================================

TABLE_EXCEL = CONFIG["tables"]["excel"]
TABLE_W2B = CONFIG["tables"]["excel_w2b"]
TABLE_B2W = CONFIG["tables"]["excel_b2w"]


# ============================================================
# LECTURE DU FICHIER GTP
# ============================================================

def lire_fichier_gtp(contenu: bytes) -> pd.DataFrame:
    """
    Lit le fichier CSV GTP.

    Structure attendue :

        Trans ID
        Date
        Amount
        Fees
        Running Bal
        Description
        Reference
        Account Id
        Last Name
    """

    try:
        df = pd.read_csv(
            BytesIO(contenu),
            dtype={
                "Trans ID": "string",
                "Reference": "string",
                "Account Id": "string",
                "Last Name": "string",
                "Description": "string",
            },
        )

    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail=f"Impossible de lire le fichier GTP : {e}",
        )

    # ---------------------------------------------------------
    # Nettoyage des noms de colonnes
    # ---------------------------------------------------------

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    # ---------------------------------------------------------
    # Suppression des colonnes complètement vides
    # ---------------------------------------------------------

    df = df.dropna(
        axis=1,
        how="all",
    )

    # ---------------------------------------------------------
    # Colonnes obligatoires
    # ---------------------------------------------------------

    colonnes_obligatoires = {
        "Trans ID",
        "Date",
        "Amount",
        "Fees",
        "Description",
        "Reference",
        "Account Id",
        "Last Name",
    }

    manquantes = colonnes_obligatoires - set(df.columns)

    if manquantes:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Colonnes GTP manquantes : "
                f"{sorted(manquantes)}. "
                f"Colonnes disponibles : "
                f"{list(df.columns)}"
            ),
        )

    return df


# ============================================================
# PREPARATION GTP
# ============================================================

def preparer_gtp(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prépare le fichier partenaire GTP pour
    la réconciliation.
    """

    df = df.copy()

    # ========================================================
    # REFERENCE / NUMERO DE CARTE
    # ========================================================

    df["WP_REFERENCE"] = (
        df["Reference"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # DATE + HEURE
    # ========================================================

    # Format observé :
    # 29-09-26 00:36:57

    df["WP_DATE_HEURE"] = pd.to_datetime(
        df["Date"],
        format="%d-%m-%y %H:%M:%S",
        errors="coerce",
    )

    # ========================================================
    # MONTANT
    # ========================================================

    # Le fichier partenaire contient des montants négatifs.
    # Pour le rapprochement :
    #
    #     abs(Amount)

    df["WP_MONTANT"] = (
        pd.to_numeric(
            df["Amount"],
            errors="coerce",
        )
        .abs()
    )

    # ========================================================
    # FRAIS
    # ========================================================

    df["WP_FRAIS"] = pd.to_numeric(
        df["Fees"],
        errors="coerce",
    ).fillna(0)

    # ========================================================
    # IDENTIFIANT TRANSACTION
    # ========================================================

    df["WP_TRANS_ID"] = (
        df["Trans ID"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # COMPTE PARTENAIRE
    # ========================================================

    df["WP_ACCOUNT_ID"] = (
        df["Account Id"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # INFORMATIONS ORIGINALES
    # ========================================================

    df["WP_DESCRIPTION"] = (
        df["Description"]
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

    # ========================================================
    # COLONNES STANDARD DU MOTEUR COMMUN
    # ========================================================

    df["DATE TRANSACTION"] = df["WP_DATE_HEURE"]

    # --------------------------------------------------------
    # GTP API = B2W
    # --------------------------------------------------------
    #
    # Le FLEX GTP correspondant est en DRCR_IND = "D".
    #
    # Le moteur commun compare donc :
    #
    #     GTP B2W <-> FLEX B2W
    #
    # avec :
    #
    #     NUMERO COMPTE
    #     +
    #     MONTANT
    #     +
    #     DATE/HEURE
    #

    df["TYPE TRANSACTION"] = "D"

    df["CODE TRANSACTION OPERATEUR"] = (
        df["WP_TRANS_ID"]
    )

    # --------------------------------------------------------
    # IMPORTANT :
    # Le numéro de carte est la clé métier GTP.
    #
    # WP_REFERENCE contient le numéro de carte.
    #
    # Le moteur commun utilisera donc :
    #
    #     NUMERO COMPTE = NUMERO DE CARTE
    # --------------------------------------------------------

    df["NUMERO COMPTE"] = df["WP_REFERENCE"]

    df["MONTANT"] = df["WP_MONTANT"]

    # ========================================================
    # STATUT DE PARTICIPATION A LA RECONCILIATION
    # ========================================================

    df["STATUT_RECONCILIATION"] = "A RAPPROCHER"

    # --------------------------------------------------------
    # GTP AGENCE
    # --------------------------------------------------------
    #
    # Les lignes "Recharge de la Carte" concernent
    # la réconciliation GTP AGENCE.
    #
    # Elles restent dans le fichier principal
    # mais ne doivent PAS participer à la
    # réconciliation GTP API.
    #

    mask_gtp_agence = (
        df["Description"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.contains(
            "RECHARGE DE LA CARTE",
            na=False,
        )
    )

    df.loc[
        mask_gtp_agence,
        "STATUT_RECONCILIATION",
    ] = "EXCLU GTP API"

    # --------------------------------------------------------
    # Compatibilité avec les traitements existants
    # --------------------------------------------------------

    df["WP_STATUT"] = df["STATUT_RECONCILIATION"]

    # ========================================================
    # CONTROLE DES DATES
    # ========================================================

    dates_invalides = (
        df["WP_DATE_HEURE"]
        .isna()
        .sum()
    )

    if dates_invalides:
        print(
            f"[gtp] Attention : "
            f"{dates_invalides} date(s) partenaire "
            f"non reconnue(s)"
        )

    # ========================================================
    # CONTROLE DES MONTANTS
    # ========================================================

    montants_invalides = (
        df["WP_MONTANT"]
        .isna()
        .sum()
    )

    if montants_invalides:
        print(
            f"[gtp] Attention : "
            f"{montants_invalides} montant(s) partenaire "
            f"non reconnu(s)"
        )

    # ========================================================
    # TRI CHRONOLOGIQUE
    # ========================================================

    df = (
        df.sort_values(
            by="WP_DATE_HEURE",
            ascending=True,
            na_position="last",
        )
        .reset_index(drop=True)
    )

    return df


# ============================================================
# SAUVEGARDE SQLITE
# ============================================================

def sauvegarder_gtp(
    df: pd.DataFrame,
    df_w2b: pd.DataFrame,
    df_b2w: pd.DataFrame,
):
    """
    Sauvegarde les données GTP dans SQLite :

    - table complète
    - W2B
    - B2W
    """

    # ---------------------------------------------------------
    # Table complète
    # ---------------------------------------------------------

    ecrire_table(
        df,
        TABLE_EXCEL,
        engine,
        log_prefix="gtp_excel",
    )

    # ---------------------------------------------------------
    # W2B
    # ---------------------------------------------------------

    ecrire_table(
        df_w2b,
        TABLE_W2B,
        engine,
        log_prefix="gtp_excel_w2b",
    )

    # ---------------------------------------------------------
    # B2W
    # ---------------------------------------------------------

    ecrire_table(
        df_b2w,
        TABLE_B2W,
        engine,
        log_prefix="gtp_excel_b2w",
    )

    return {
        "total_rows": len(df),
        "w2b_rows": len(df_w2b),
        "b2w_rows": len(df_b2w),
    }


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/health")
def health():

    return {
        "status": "ok",
        "partenaire": PARTENAIRE,
        "service": "gtp-excel",
        "table": TABLE_EXCEL,
    }


# ============================================================
# API PRINCIPALE
# ============================================================

@app.post("/process-excel")
async def process_excel(
    files: list[UploadFile] = File(...),
    format: str = "json",
):
    if not files:
        raise HTTPException(
            status_code=400,
            detail="Aucun fichier GTP fourni.",
        )

    file = files[0]
    """
    Reçoit le fichier CSV GTP partenaire.

    Exemple :

        POST /process-excel
    """

    # ========================================================
    # CONTROLE FICHIER
    # ========================================================

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="Aucun fichier fourni.",
        )

    contenu = await file.read()

    if not contenu:
        raise HTTPException(
            status_code=400,
            detail="Le fichier GTP est vide.",
        )

    print(
        f"[gtp_excel] Fichier reçu : "
        f"{file.filename}"
    )

    print(
        f"[gtp_excel] Taille : "
        f"{len(contenu)} octets"
    )

    # ========================================================
    # LECTURE
    # ========================================================

    df_brut = lire_fichier_gtp(contenu)

    print(
        f"[gtp_excel] Nombre de lignes brutes : "
        f"{len(df_brut)}"
    )

    # ========================================================
    # PREPARATION
    # ========================================================

    df = preparer_gtp(df_brut)

    print(
        f"[gtp_excel] Nombre de lignes préparées : "
        f"{len(df)}"
    )

    print(
        f"[gtp_excel] Cartes identifiées : "
        f"{(df['WP_REFERENCE'] != '').sum()}"
    )

    print(
        f"[gtp_excel] Montants valides : "
        f"{df['WP_MONTANT'].notna().sum()}"
    )

    # ========================================================
    # SEPARATION W2B / B2W
    # ========================================================
    #
    # GTP API utilise le flux B2W.
    #
    # Les lignes :
    #
    #     A RAPPROCHER
    #
    # sont envoyées dans GTP_B2W.
    #
    # Les lignes :
    #
    #     EXCLU GTP API
    #
    # restent dans GTP_EXCEL mais ne sont pas
    # envoyées dans GTP_B2W.
    #

    df_a_rapprocher = df[
        df["STATUT_RECONCILIATION"]
        == "A RAPPROCHER"
    ].copy()

    # GTP API ne possède actuellement
    # aucun flux W2B.

    df_w2b = df.iloc[0:0].copy()

    # Flux utilisé par la réconciliation GTP API

    df_b2w = df_a_rapprocher.copy()

    print(
        f"[gtp_excel] Répartition : "
        f"{len(df)} total / "
        f"{len(df_w2b)} W2B / "
        f"{len(df_b2w)} B2W / "
        f"{len(df) - len(df_b2w)} exclu"
    )

    # ========================================================
    # SAUVEGARDE SQLITE
    # ========================================================

    sauvegarde = sauvegarder_gtp(
        df,
        df_w2b,
        df_b2w,
    )

    print(
        f"[gtp_excel] SQLite : "
        f"{sauvegarde['total_rows']} total / "
        f"{sauvegarde['w2b_rows']} W2B / "
        f"{sauvegarde['b2w_rows']} B2W"
    )

    # ========================================================
    # EXPORT EXCEL
    # ========================================================

    if format.lower() == "excel":

        return respond_sheets(
            {
                "GTP_EXCEL": df,
                "GTP_W2B": df_w2b,
                "GTP_B2W": df_b2w,
            },
            filename="GTP_EXCEL.xlsx",
            format="excel",
        )

    # ========================================================
    # REPONSE JSON
    # ========================================================

    return JSONResponse(
        content={
            "status": "ok",
            "partenaire": PARTENAIRE,
            "filename": file.filename,
            "total_rows": len(df),
            "w2b_rows": len(df_w2b),
            "b2w_rows": len(df_b2w),
            "columns": df.columns.tolist(),
        }
    )