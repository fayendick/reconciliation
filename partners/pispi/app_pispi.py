# ============================================================
# PI/SPI — SERVICE EXCEL
# ============================================================

import io
import unicodedata
from typing import List

import pandas as pd

from fastapi import (
    FastAPI,
    UploadFile,
    File,
    HTTPException,
    Query,
)

from common.config import (
    DB_PATH,
    PARTENAIRES,
    make_sqlite_engine,
)

from common.excel_common import sauvegarder_sqlite
from common.sqlite_io import lire_table_json
from common.http_export import respond_sheets, wants_excel


# ============================================================
# CONFIGURATION
# ============================================================

PARTENAIRE = "PISPI"

if PARTENAIRE not in PARTENAIRES:
    raise RuntimeError(
        f"Le partenaire '{PARTENAIRE}' n'existe pas dans "
        f"common.config.PARTENAIRES."
    )

TABLES = PARTENAIRES[PARTENAIRE]["tables"]

engine = make_sqlite_engine()


# ============================================================
# APPLICATION FASTAPI
# ============================================================

app = FastAPI(
    title="PI/SPI Excel API",
    description=(
        "Service de chargement et normalisation "
        "des fichiers PI/SPI pour la réconciliation."
    ),
)


# ============================================================
# OUTILS
# ============================================================

def normaliser_texte(valeur):
    if valeur is None:
        return ""

    try:
        if pd.isna(valeur):
            return ""
    except Exception:
        pass

    texte = str(valeur).strip().upper()

    texte = unicodedata.normalize(
        "NFKD",
        texte,
    )

    texte = "".join(
        caractere
        for caractere in texte
        if not unicodedata.combining(caractere)
    )

    return " ".join(texte.split())


def trouver_colonne(df, candidats):

    if df is None or df.empty:
        return None

    mapping = {}

    for colonne in df.columns:

        mapping[
            normaliser_texte(colonne)
        ] = colonne

    for candidat in candidats:

        cle = normaliser_texte(candidat)

        if cle in mapping:
            return mapping[cle]

    return None


# ============================================================
# LECTURE FICHIER PI/SPI
# ============================================================

def lire_fichier_pispi(upload_file):

    contenu = upload_file.file.read()

    if not contenu:
        raise ValueError(
            "Le fichier PI/SPI est vide."
        )

    excel = pd.ExcelFile(
        io.BytesIO(contenu)
    )

    feuille_detail = None

    for feuille in excel.sheet_names:

        if normaliser_texte(feuille) == normaliser_texte(
            "Détail Compensation"
        ):
            feuille_detail = feuille
            break

    if feuille_detail is None:

        raise ValueError(
            "La feuille 'Détail Compensation' "
            "n'a pas été trouvée. "
            f"Feuilles disponibles : {excel.sheet_names}"
        )

    df = pd.read_excel(
        io.BytesIO(contenu),
        sheet_name=feuille_detail,
        header=1,
    )

    df.columns = [
        str(colonne).strip()
        for colonne in df.columns
    ]

    df = df.dropna(
        axis=1,
        how="all",
    )

    df = df.dropna(
        axis=0,
        how="all",
    )

    df = df.reset_index(
        drop=True
    )

    return df


# ============================================================
# TRAITEMENT PI/SPI
# ============================================================

def traiter_pispi(upload_file):

    df = lire_fichier_pispi(
        upload_file
    )

    # --------------------------------------------------------
    # RECHERCHE DES COLONNES
    # --------------------------------------------------------

    col_reference = trouver_colonne(
        df,
        [
            "Référence",
            "Reference",
        ],
    )

    col_sens = trouver_colonne(
        df,
        [
            "Sens compensation",
            "Sens",
            "Statut",
        ],
    )

    col_montant = trouver_colonne(
        df,
        [
            "Montant",
        ],
    )

    col_date = trouver_colonne(
        df,
        [
            "Date irrévocabilité",
            "Date irrevocabilite",
            "Date",
        ],
    )
    
    col_compte = trouver_colonne(
        df,
        [
            "Numero de compte du client",
            "Numéro de compte du client",
        ],
    )

    colonnes_manquantes = []

    if col_reference is None:
        colonnes_manquantes.append(
            "Référence"
        )

    if col_sens is None:
        colonnes_manquantes.append(
            "Sens compensation"
        )

    if col_montant is None:
        colonnes_manquantes.append(
            "Montant"
        )

    if col_date is None:
        colonnes_manquantes.append(
            "Date irrévocabilité"
        )
    if col_compte is None:
            colonnes_manquantes.append(
            "Numero de compte du client"
        )

    if colonnes_manquantes:

        raise ValueError(
            "Colonnes PI/SPI manquantes : "
            + ", ".join(
                colonnes_manquantes
            )
        )

    # --------------------------------------------------------
    # FORMAT STANDARD
    # --------------------------------------------------------

    df["REFERENCETRANSACTION"] = (
    df[col_reference]
    .astype(str)
    .str.strip()
    )

    # Éviter le conflit SQLite entre "Montant" et "montant"
    if col_montant == "Montant":
        df = df.rename(
            columns={
                col_montant: "Montant_Excel"
            }
        )
        col_montant = "Montant_Excel"

    df["MONTANT"] = pd.to_numeric(
        df[col_montant],
        errors="coerce",
    ).abs()

    df["date"] = pd.to_datetime(
        df[col_date],
        errors="coerce",
    )

    # --------------------------------------------------------
    # SENS
    #
    # REÇU   → CREDIT → W2B
    # ENVOYE → DEBIT  → B2W
    # --------------------------------------------------------

    df["SENS"] = (
        df[col_sens]
        .apply(normaliser_texte)
        .map(
            {
                "RECU": "CREDIT",
                "ENVOYE": "DEBIT",
                "CREDIT": "CREDIT",
                "DEBIT": "DEBIT",
            }
        )
    )

    # --------------------------------------------------------
    # TYPE TRANSACTION
    #
    # CREDIT → W2B
    # DEBIT  → B2W
    # --------------------------------------------------------

    df["TYPE_TRANSACTION"] = df[
        "SENS"
    ].map(
        {
            "CREDIT": "C",
            "DEBIT": "D",
        }
    )

    # --------------------------------------------------------
    # NETTOYAGE
    # --------------------------------------------------------

    df = df.dropna(
        subset=[
            "REFERENCETRANSACTION",
            "MONTANT",
            "date",
            "SENS",
        ]
    ).copy()

    df = df.sort_values(
        by="date"
    ).reset_index(
        drop=True
    )
     
    
    # ========================================================
    # Colonnes standard attendues par le moteur de réconciliation
    # ========================================================

    df["DATE TRANSACTION"] = df["date"]

    df["TYPE TRANSACTION"] = df["SENS"].map({
        "CREDIT": "W2B",
        "DEBIT": "B2W",
    })

    df["CODE TRANSACTION OPERATEUR"] = df["REFERENCETRANSACTION"]
    df["CODE_TRANSACTION"] = df["REFERENCETRANSACTION"]

    df["NUMERO COMPTE"] = df[col_compte]

    
     
    
    
    return df


# ============================================================
# SEPARATION W2B / B2W
# ============================================================

def separer_w2b_b2w(df):

    df = df.copy()

    # CREDIT → W2B
    df_w2b = df[
        df["SENS"] == "CREDIT"
    ].copy()

    # DEBIT → B2W
    df_b2w = df[
        df["SENS"] == "DEBIT"
    ].copy()

    return (
        df_w2b,
        df_b2w,
    )


# ============================================================
# SAUVEGARDE SQLITE
# ============================================================

def sauvegarder_pispi(df):

    df_w2b, df_b2w = separer_w2b_b2w(
        df
    )

    sauvegarder_sqlite(
        df,
        TABLES["excel"],
        engine,
    )

    sauvegarder_sqlite(
        df_w2b,
        TABLES["excel_w2b"],
        engine,
    )

    sauvegarder_sqlite(
        df_b2w,
        TABLES["excel_b2w"],
        engine,
    )

    return (
        df_w2b,
        df_b2w,
    )


# ============================================================
# ENDPOINT PRINCIPAL
# ============================================================

@app.post("/process-excel")
async def process_excel(
    files: List[UploadFile] = File(...),
    format: str = Query(
        "excel",
        description="excel ou json",
    ),
):

    if not files:

        raise HTTPException(
            status_code=400,
            detail="Aucun fichier PI/SPI fourni.",
        )

    resultats = []
    erreurs = []

    # --------------------------------------------------------
    # TRAITEMENT DES FICHIERS
    # --------------------------------------------------------

    for fichier in files:

        try:

            df = traiter_pispi(
                fichier
            )

            resultats.append(
                df
            )

        except Exception as erreur:

            erreurs.append(
                {
                    "fichier": fichier.filename,
                    "erreur": str(erreur),
                }
            )

    if not resultats:

        raise HTTPException(
            status_code=400,
            detail={
                "message": (
                    "Aucun fichier PI/SPI "
                    "n'a pu être traité."
                ),
                "erreurs": erreurs,
            },
        )

    # --------------------------------------------------------
    # CONSOLIDATION
    # --------------------------------------------------------

    df_final = pd.concat(
        resultats,
        ignore_index=True,
    )

    # --------------------------------------------------------
    # SAUVEGARDE
    # --------------------------------------------------------

    df_w2b, df_b2w = sauvegarder_pispi(
        df_final
    )

    # --------------------------------------------------------
    # RETOUR
    # --------------------------------------------------------

    sheets = {
        "PISPI": df_final,
        "PISPI_W2B": df_w2b,
        "PISPI_B2W": df_b2w,
    }

    if format.lower() == "json":

        return {
            "status": "ok",
            "format": "json",
            "filename": "PISPI.xlsx",
            "fichiers": [
                fichier.filename
                for fichier in files
            ],
            "nb_fichiers": len(
                resultats
            ),
            "nb_transactions": len(
                df_final
            ),
            "nb_w2b": len(
                df_w2b
            ),
            "nb_b2w": len(
                df_b2w
            ),
            "erreurs": erreurs,
        }

    return respond_sheets(
        sheets,
        filename="PISPI.xlsx",
        format=format,
    )


# ============================================================
# ROUTES SQLITE
# ============================================================

@app.get("/db/pispi")
def db_pispi():

    return lire_table_json(
        DB_PATH,
        TABLES["excel"],
    )


@app.get("/db/pispi-w2b")
def db_pispi_w2b():

    return lire_table_json(
        DB_PATH,
        TABLES["excel_w2b"],
    )


@app.get("/db/pispi-b2w")
def db_pispi_b2w():

    return lire_table_json(
        DB_PATH,
        TABLES["excel_b2w"],
    )