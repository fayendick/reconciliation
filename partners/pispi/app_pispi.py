from __future__ import annotations

import os
import re
import unicodedata
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import JSONResponse

from config import DB_PATH, PARTENAIRES, make_sqlite_engine
from common.excel_common import sauvegarder_sqlite
from common.sqlite_io import lire_table_json


# ============================================================
# CONFIGURATION
# ============================================================

PARTENAIRE = "pispi"

SHEET_PISPI = "Liste des transferts de fonds"

TABLE_COMPILATION = "COMPILATION_PISPI"
TABLE_W2B = "PISPI_W2B"
TABLE_B2W = "PISPI_B2W"

# Statuts Excel autorisés
STATUT_RECU = "RECU"
STATUT_ENVOYE = "ENVOYE"


# ============================================================
# APPLICATION FASTAPI
# ============================================================

app = FastAPI(
    title="PISPI Excel Service",
    description="Service de traitement des fichiers PI/SPI",
    version="1.0.0",
)


# ============================================================
# MOTEUR SQLITE COMMUN
# ============================================================

engine = make_sqlite_engine()


# ============================================================
# OUTILS
# ============================================================

def normaliser_texte(value: Any) -> str:
    """
    Normalise un texte :
    - suppression des accents
    - passage en majuscules
    - suppression des espaces superflus
    """
    if value is None or pd.isna(value):
        return ""

    value = str(value).strip()

    value = unicodedata.normalize("NFKD", value)
    value = "".join(
        c
        for c in value
        if not unicodedata.combining(c)
    )

    value = value.upper()

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def trouver_colonne(df: pd.DataFrame, nom_recherche: str) -> str | None:
    """
    Recherche une colonne en ignorant :
    - accents
    - majuscules/minuscules
    - espaces superflus
    """
    cible = normaliser_texte(nom_recherche)

    for colonne in df.columns:
        if normaliser_texte(colonne) == cible:
            return colonne

    return None


def normaliser_compte(value: Any) -> str:
    """
    Conserve le numéro de compte sous forme de texte.

    Gère notamment le cas où Excel transforme un compte
    numérique en valeur du type 123456.0.
    """
    if value is None or pd.isna(value):
        return ""

    # Cas numérique
    if isinstance(value, (int, float)):
        try:
            if float(value).is_integer():
                return str(int(value))
        except Exception:
            pass

    valeur = str(value).strip()

    # Cas Excel : "123456.0"
    if valeur.endswith(".0"):
        partie = valeur[:-2]

        if partie.isdigit():
            return partie

    return valeur


def convertir_montant(value: Any):
    """
    Convertit correctement les montants Excel en numérique.

    Gère notamment :
    - 1000
    - 1000.50
    - 1000,50
    - 1 000,50
    - 1.000,50
    """
    if value is None or pd.isna(value):
        return pd.NA

    if isinstance(value, (int, float)):
        return float(value)

    valeur = str(value).strip()

    if not valeur:
        return pd.NA

    valeur = valeur.replace("\u00A0", "")
    valeur = valeur.replace(" ", "")

    # Présence simultanée de . et ,
    if "," in valeur and "." in valeur:

        # Exemple européen : 1.234,56
        if valeur.rfind(",") > valeur.rfind("."):
            valeur = valeur.replace(".", "")
            valeur = valeur.replace(",", ".")

        # Exemple anglo-saxon : 1,234.56
        else:
            valeur = valeur.replace(",", "")

    elif "," in valeur:
        # Exemple : 1234,56
        valeur = valeur.replace(",", ".")

    return pd.to_numeric(valeur, errors="coerce")


# ============================================================
# LECTURE DU FICHIER EXCEL
# ============================================================

def lire_fichier_pispi(fichier: str) -> pd.DataFrame:
    """
    Lecture du nouvel onglet PI/SPI :

    Liste des transferts de fonds
    """

    if not os.path.exists(fichier):
        raise FileNotFoundError(
            f"Fichier introuvable : {fichier}"
        )

    print(
        f"[PISPI] Lecture du fichier : {fichier}"
    )

    try:
        df = pd.read_excel(
            fichier,
            sheet_name=SHEET_PISPI,
            header=0,
        )
    except ValueError as exc:
        raise ValueError(
            f"Onglet '{SHEET_PISPI}' introuvable dans le fichier."
        ) from exc

    print(
        f"[PISPI] Onglet '{SHEET_PISPI}' : "
        f"{len(df)} lignes"
    )

    print(
        f"[PISPI] Colonnes détectées : "
        f"{df.columns.tolist()}"
    )

    return df


# ============================================================
# TRAITEMENT PI/SPI
# ============================================================

def traiter_pispi(df: pd.DataFrame) -> pd.DataFrame:
    """
    Transforme le nouveau format PI/SPI vers le format
    exploitable par le rapprochement commun.

    Correspondances :

    Date
        -> date

    Référence
        -> REFERENCETRANSACTION

    REÇU / réception :
        Compte du payé
            -> NUMERO_COMPTE

    ENVOYÉ / envoi :
        Compte du payeur
            -> NUMERO_COMPTE

    Montant
        -> montant

    REÇU
        -> CREDIT
        -> W2B

    ENVOYÉ
        -> DEBIT
        -> B2W
    """

    if df is None or df.empty:
        return pd.DataFrame(
            columns=[
                "REFERENCETRANSACTION",
                "NUMERO_COMPTE",
                "CLIENT_PAYE",
                "montant",
                "MONTANT",
                "date",
                "DATE TRANSACTION",
                "SENS",
                "TYPE_TRANSACTION",
                "TYPE TRANSACTION",
            ]
        )

    df = df.copy()

    # --------------------------------------------------------
    # Colonnes obligatoires
    # --------------------------------------------------------

    colonnes_requises = [
        "Date",
        "Référence",
        "Compte du payeur",
        "Compte du payé",
        "Client payé",
        "Montant",
        "Statut",
    ]

    colonnes = {}

    for nom in colonnes_requises:
        colonne = trouver_colonne(df, nom)

        if colonne is None:
            raise ValueError(
                f"Colonne obligatoire absente : '{nom}'. "
                f"Colonnes disponibles : {df.columns.tolist()}"
            )

        colonnes[nom] = colonne

    print(
        "[PISPI] Colonnes utilisées : "
        f"Date='{colonnes['Date']}', "
        f"Référence='{colonnes['Référence']}', "
        f"Compte du payeur='{colonnes['Compte du payeur']}', "
        f"Compte du payé='{colonnes['Compte du payé']}', "
        f"Montant='{colonnes['Montant']}', "
        f"Statut='{colonnes['Statut']}'"
    )

    # --------------------------------------------------------
    # Normalisation du statut
    # --------------------------------------------------------

    df["_STATUT_NORMALISE"] = (
        df[colonnes["Statut"]]
        .apply(normaliser_texte)
    )

    nb_avant = len(df)

    # IMPORTANT :
    # On garde uniquement les opérations réussies
    # REÇU / ENVOYÉ
    df = df[
        df["_STATUT_NORMALISE"].isin(
            [STATUT_RECU, STATUT_ENVOYE]
        )
    ].copy()

    print(
        f"[PISPI] Filtre opérations : "
        f"{nb_avant} avant / {len(df)} après"
    )

    if df.empty:
        return pd.DataFrame(
            columns=[
                "REFERENCETRANSACTION",
                "NUMERO_COMPTE",
                "CLIENT_PAYE",
                "montant",
                "MONTANT",
                "date",
                "DATE TRANSACTION",
                "SENS",
                "TYPE_TRANSACTION",
                "TYPE TRANSACTION",
            ]
        )

    # --------------------------------------------------------
    # Référence
    # --------------------------------------------------------

    df["REFERENCETRANSACTION"] = (
        df[colonnes["Référence"]]
        .apply(
            lambda x: ""
            if pd.isna(x)
            else str(x).strip()
        )
    )

    # --------------------------------------------------------
    # COMPTE UTILISÉ POUR LE RAPPROCHEMENT
    #
    # REÇU / réception -> Compte du payé
    # ENVOYÉ / envoi    -> Compte du payeur
    #
    # Cette règle est spécifique à PI/SPI.
    # --------------------------------------------------------

    df["NUMERO_COMPTE"] = ""

    masque_recu = df["_STATUT_NORMALISE"] == STATUT_RECU
    masque_envoye = df["_STATUT_NORMALISE"] == STATUT_ENVOYE

    df.loc[masque_recu, "NUMERO_COMPTE"] = (
        df.loc[masque_recu, colonnes["Compte du payé"]]
        .apply(normaliser_compte)
    )

    df.loc[masque_envoye, "NUMERO_COMPTE"] = (
        df.loc[masque_envoye, colonnes["Compte du payeur"]]
        .apply(normaliser_compte)
    )

    # --------------------------------------------------------
    # CLIENT PAYÉ
    # --------------------------------------------------------

    df["CLIENT_PAYE"] = (
        df[colonnes["Client payé"]]
        .apply(normaliser_texte)
    )

    # --------------------------------------------------------
    # MONTANT
    # --------------------------------------------------------

    df["montant"] = (
        df[colonnes["Montant"]]
        .apply(convertir_montant)
    )

    # --------------------------------------------------------
    # DATE / HEURE
    # --------------------------------------------------------

    df["date"] = pd.to_datetime(
        df[colonnes["Date"]],
        errors="coerce",
        dayfirst=True,
    )

    # --------------------------------------------------------
    # SENS
    # --------------------------------------------------------

    df["SENS"] = df["_STATUT_NORMALISE"].map(
        {
            STATUT_RECU: "CREDIT",
            STATUT_ENVOYE: "DEBIT",
        }
    )

    # --------------------------------------------------------
    # TYPE DE TRANSACTION
    # --------------------------------------------------------

    df["TYPE_TRANSACTION"] = df["SENS"].map(
        {
            "CREDIT": "W2B",
            "DEBIT": "B2W",
        }
    )

    # Alias attendu par le moteur commun
    df["TYPE TRANSACTION"] = df["TYPE_TRANSACTION"]

    # --------------------------------------------------------
    # Nettoyage
    #
    # La référence n'est PAS utilisée comme clé
    # de rapprochement.
    #
    # Donc une référence vide ne doit pas supprimer
    # une transaction valide.
    # --------------------------------------------------------

    df["NUMERO_COMPTE"] = (
        df["NUMERO_COMPTE"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["REFERENCETRANSACTION"] = (
        df["REFERENCETRANSACTION"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["montant"] = pd.to_numeric(
        df["montant"],
        errors="coerce",
    )

    df["date"] = pd.to_datetime(
        df["date"],
        errors="coerce",
    )

    # Pour le rapprochement, ces trois informations sont
    # nécessaires :
    # - compte
    # - montant
    # - date
    df = df[
        (df["NUMERO_COMPTE"] != "")
        & df["montant"].notna()
        & df["date"].notna()
    ].copy()

    # --------------------------------------------------------
    # Colonnes finales
    #
    # Les colonnes "montant" et "date" sont conservées pour
    # l'exploitation PI/SPI.
    #
    # Les alias sont ajoutés pour respecter le format attendu
    # par le moteur commun de réconciliation.
    # --------------------------------------------------------

    df["MONTANT"] = df["montant"]
    df["MONTANT_COMPARAISON"] = df["montant"]
    df["DATE TRANSACTION"] = df["date"]

    colonnes_finales = [
        "REFERENCETRANSACTION",
        "NUMERO_COMPTE",
        "CLIENT_PAYE",
        "montant",
        "MONTANT",
        "MONTANT_COMPARAISON",
        "date",
        "DATE TRANSACTION",
        "SENS",
        "TYPE_TRANSACTION",
        "TYPE TRANSACTION",
    ]

    df = df[colonnes_finales].copy()

    # --------------------------------------------------------
    # Tri
    # --------------------------------------------------------

    df = df.sort_values(
        by="date",
        kind="stable",
    ).reset_index(drop=True)

    print(
        f"[PISPI] Transactions finales : {len(df)}"
    )

    print(
        "[PISPI] Répartition SENS :\n"
        f"{df['SENS'].value_counts(dropna=False).to_string()}"
    )

    print(
        "[PISPI] Répartition TYPE_TRANSACTION :\n"
        f"{df['TYPE_TRANSACTION'].value_counts(dropna=False).to_string()}"
    )

    return df


# ============================================================
# SEPARATION W2B / B2W
# ============================================================

def separer_w2b_b2w(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:

    if df is None or df.empty:
        return (
            df.copy(),
            df.copy(),
        )

    w2b = df[
        df["TYPE_TRANSACTION"] == "W2B"
    ].copy()

    b2w = df[
        df["TYPE_TRANSACTION"] == "B2W"
    ].copy()

    print(
        f"[PISPI] W2B / CREDIT : {len(w2b)}"
    )

    print(
        f"[PISPI] B2W / DEBIT : {len(b2w)}"
    )

    return w2b, b2w


# ============================================================
# SECURITE AVANT SQLITE
# ============================================================

def supprimer_doublons_colonnes_sqlite(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Sécurité supplémentaire contre les collisions de noms
    SQLite, par exemple :

        Date
        date

    ou :

        Montant
        montant
    """

    if df is None or df.empty:
        return df

    colonnes_vues = set()
    colonnes_a_garder = []

    for colonne in df.columns:

        cle = str(colonne).strip().lower()

        if cle in colonnes_vues:
            print(
                f"[PISPI] Colonne supprimée avant SQLite : "
                f"{colonne}"
            )
            continue

        colonnes_vues.add(cle)
        colonnes_a_garder.append(colonne)

    return df.loc[:, colonnes_a_garder].copy()


# ============================================================
# SAUVEGARDE SQLITE
# ============================================================

def sauvegarder_pispi(
    df: pd.DataFrame,
) -> dict[str, Any]:

    if df is None:
        df = pd.DataFrame()

    df = supprimer_doublons_colonnes_sqlite(df)

    w2b, b2w = separer_w2b_b2w(df)

    print(
        f"[PISPI] Sauvegarde {TABLE_COMPILATION} : "
        f"{len(df)} lignes"
    )

    sauvegarder_sqlite(
        df,
        TABLE_COMPILATION,
        engine,
        log_prefix="PISPI",
    )

    print(
        f"[PISPI] Sauvegarde {TABLE_W2B} : "
        f"{len(w2b)} lignes"
    )

    sauvegarder_sqlite(
        w2b,
        TABLE_W2B,
        engine,
        log_prefix="PISPI",
    )

    print(
        f"[PISPI] Sauvegarde {TABLE_B2W} : "
        f"{len(b2w)} lignes"
    )

    sauvegarder_sqlite(
        b2w,
        TABLE_B2W,
        engine,
        log_prefix="PISPI",
    )

    return {
        "table_compilation": TABLE_COMPILATION,
        "table_w2b": TABLE_W2B,
        "table_b2w": TABLE_B2W,
        "nb_total": len(df),
        "nb_w2b": len(w2b),
        "nb_b2w": len(b2w),
    }


# ============================================================
# ENDPOINT PRINCIPAL
# ============================================================

@app.post("/process-excel")
async def process_excel(
    fichier: UploadFile = File(...),
):
    """
    Endpoint appelé par le gateway.

    Exemple :
        /svc/pispi-excel/process-excel
    """

    nom_fichier = fichier.filename or "pispi.xlsx"

    if not nom_fichier.lower().endswith(
        (".xlsx", ".xls")
    ):
        raise HTTPException(
            status_code=400,
            detail="Le fichier doit être un fichier Excel (.xlsx ou .xls).",
        )

    chemin_temporaire = os.path.join(
        "data",
        f"_pispi_{nom_fichier}",
    )

    os.makedirs(
        "data",
        exist_ok=True,
    )

    try:

        # ----------------------------------------------------
        # Sauvegarde temporaire
        # ----------------------------------------------------

        contenu = await fichier.read()

        with open(
            chemin_temporaire,
            "wb",
        ) as f:
            f.write(contenu)

        print(
            f"[PISPI] Fichier reçu : {nom_fichier}"
        )

        # ----------------------------------------------------
        # Lecture
        # ----------------------------------------------------

        df_brut = lire_fichier_pispi(
            chemin_temporaire
        )

        # ----------------------------------------------------
        # Transformation
        # ----------------------------------------------------

        df_final = traiter_pispi(
            df_brut
        )

        # ----------------------------------------------------
        # Sauvegarde SQLite
        # ----------------------------------------------------

        sauvegarde = sauvegarder_pispi(
            df_final
        )

        # ----------------------------------------------------
        # Réponse
        # ----------------------------------------------------

        return JSONResponse(
            content={
                "success": True,
                "partenaire": PARTENAIRE,
                "fichier": nom_fichier,
                "onglet": SHEET_PISPI,
                "nb_lignes_brutes": len(df_brut),
                "nb_lignes_finales": len(df_final),
                "nb_w2b": sauvegarde["nb_w2b"],
                "nb_b2w": sauvegarde["nb_b2w"],
                "tables": {
                    "compilation": TABLE_COMPILATION,
                    "w2b": TABLE_W2B,
                    "b2w": TABLE_B2W,
                },
            }
        )

    except HTTPException:
        raise

    except Exception as exc:

        print(
            f"[PISPI] ERREUR : {exc}"
        )

        raise HTTPException(
            status_code=500,
            detail=str(exc),
        )

    finally:

        # ----------------------------------------------------
        # Suppression du fichier temporaire
        # ----------------------------------------------------

        try:
            if os.path.exists(
                chemin_temporaire
            ):
                os.remove(
                    chemin_temporaire
                )
        except Exception as exc:
            print(
                f"[PISPI] Impossible de supprimer "
                f"le fichier temporaire : {exc}"
            )


# ============================================================
# LECTURE SQLITE
# ============================================================

@app.get("/db/pispi")
def get_pispi():
    try:
        return lire_table_json(
            engine,
            TABLE_COMPILATION,
        )
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=str(exc),
        )


@app.get("/db/pispi-w2b")
def get_pispi_w2b():
    try:
        return lire_table_json(
            engine,
            TABLE_W2B,
        )
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=str(exc),
        )


@app.get("/db/pispi-b2w")
def get_pispi_b2w():
    try:
        return lire_table_json(
            engine,
            TABLE_B2W,
        )
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=str(exc),
        )


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/health")
def health():
    return {
        "status": "ok",
        "service": "pispi-excel",
        "partenaire": PARTENAIRE,
        "sheet": SHEET_PISPI,
        "database": DB_PATH,
    }