# ============================================================
# EXCEL COMMON — logique PARTAGÉE par tous les app_<partenaire>.py
# ------------------------------------------------------------
# Ce module ne connaît AUCUN partenaire en particulier. Chaque
# app_<partenaire>.py :
#   1. lit son propre fichier Excel (colonnes brutes propres au
#      partenaire),
#   2. renomme SES colonnes vers le schéma standard
#      (COLONNES_STANDARD_EXCEL dans config.py) - c'est la SEULE
#      partie spécifique à chaque partenaire,
#   3. appelle ensuite les fonctions communes ci-dessous, qui
#      sont identiques pour tout le monde.
#
# EXCEPTION : ORANGE_USSD (voir COLONNES_STANDARD_EXCEL_PAR_PARTENAIRE
# et les commentaires [CORRECTIF] ci-dessous). Ce partenaire fournit
# DEBIT et CREDIT séparément (au lieu d'un MONTANT unique) et
# DATE_HEURE (au lieu de DATE TRANSACTION). Tous les autres
# partenaires (Wave, Wizz, ...) gardent EXACTEMENT le même
# comportement qu'avant, colonne pour colonne.
# ============================================================

from typing import List, Optional

import pandas as pd
from fastapi import UploadFile, HTTPException
import io

from config import COLONNES_STANDARD_EXCEL


# ============================================================
# [CORRECTIF] SCHÉMA STANDARD SPÉCIFIQUE À ORANGE_USSD
# ------------------------------------------------------------
# ORANGE_USSD fournit désormais DATE_HEURE (Date+Heure déjà
# fusionnées par app_orange_ussd.py) et DEBIT/CREDIT séparés (au
# lieu de DATE TRANSACTION + MONTANT), pour permettre à
# reconciliation_engine.preparer_wave_partenaire() de choisir DEBIT
# ou CREDIT selon le sens (W2B/B2W), exactement comme il le fait déjà
# côté Flex Oracle (MOUVEMENT_DEBIT/MOUVEMENT_CREDIT).
#
# Ce dict ne contient QUE les partenaires qui dérogent au schéma
# standard global (COLONNES_STANDARD_EXCEL, dans config.py). Tout
# partenaire absent de ce dict (Wave, Wizz, ...) continue d'utiliser
# COLONNES_STANDARD_EXCEL tel quel, sans AUCUN changement de
# comportement.
# ============================================================
COLONNES_STANDARD_EXCEL_PAR_PARTENAIRE = {
    "ORANGE_USSD": [
        "DATE_HEURE",
        "TYPE TRANSACTION",
        "CODE TRANSACTION OPERATEUR",
        "NUMERO COMPTE",
        "DEBIT",
        "CREDIT",
    ],
}


# ============================================================
# 1. LECTURE BRUTE (identique pour tous : lit le fichier, upper+strip
#    les en-têtes, ajoute la colonne de traçabilité __SOURCE_FILE__)
# ============================================================

def lire_fichiers_excel(files: List[UploadFile], sheet_name=0) -> List[pd.DataFrame]:

    dfs = []

    print("=" * 80)
    print(f"[excel_common] Nombre de fichiers reçus : {len(files)}")

    for file in files:
        try:
            content = file.file.read()

            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name)
            df = df.dropna(how="all").reset_index(drop=True)

            df.columns = (
                df.columns.astype(str)
                .str.strip()
                .str.upper()
            )

            df["__SOURCE_FILE__"] = file.filename
            dfs.append(df)

        finally:
            file.file.close()

    return dfs


# ============================================================
# 2. RENOMMAGE VERS LE SCHÉMA STANDARD
# ------------------------------------------------------------
# `mapping` : dict "NOM BRUT MAJUSCULE" -> "NOM STANDARD"
# Pour un partenaire dont les colonnes sont déjà standard
# (comme Wave), passer un mapping vide {}.
# ============================================================

def appliquer_mapping(dfs: List[pd.DataFrame], mapping: dict) -> List[pd.DataFrame]:
    if not mapping:
        return dfs
    return [df.rename(columns=mapping) for df in dfs]


# ============================================================
# 3. VALIDATION DU SCHÉMA STANDARD (avant compilation)
# ------------------------------------------------------------
# [CORRECTIF] Utilise COLONNES_STANDARD_EXCEL_PAR_PARTENAIRE si le
# partenaire y figure (actuellement : ORANGE_USSD uniquement), sinon
# retombe sur COLONNES_STANDARD_EXCEL (config.py), comme avant.
# Aucun changement de comportement pour Wave/Wizz/autres.
# ============================================================

def valider_colonnes_standard(df: pd.DataFrame, partenaire: str):
    colonnes_attendues = COLONNES_STANDARD_EXCEL_PAR_PARTENAIRE.get(
        partenaire, COLONNES_STANDARD_EXCEL
    )
    manquantes = [c for c in colonnes_attendues if c not in df.columns]
    if manquantes:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Colonnes standard manquantes pour {partenaire} : {manquantes}. "
                f"Colonnes trouvées dans le fichier : {list(df.columns)}. "
                f"Vérifie le mapping de colonnes dans app_{partenaire.lower()}.py."
            )
        )


# ============================================================
# 4. NETTOYAGE (dates + montants) — identique pour tous
# ------------------------------------------------------------
# [CORRECTIF] Gère aussi DATE_HEURE (en plus de DATE TRANSACTION) et
# DEBIT/CREDIT (en plus de MONTANT) dans la liste de colonnes
# numériques par défaut. Ces ajouts sont sans effet pour les
# partenaires qui n'ont pas ces colonnes (Wave, Wizz, ...), puisque
# chaque conversion ne s'applique que "if col in df.columns".
# ============================================================

def nettoyer_dataframe(
    df: pd.DataFrame,
    colonnes_numeriques: Optional[List[str]] = None,
) -> pd.DataFrame:

    df = df.copy()

    if "DATE TRANSACTION" in df.columns:
        df["DATE TRANSACTION"] = pd.to_datetime(df["DATE TRANSACTION"], errors="coerce")

    if "DATE_HEURE" in df.columns:
        df["DATE_HEURE"] = pd.to_datetime(df["DATE_HEURE"], errors="coerce")

    colonnes_numeriques = colonnes_numeriques or [
        "MONTANT",
        "DEBIT",
        "CREDIT",
        "FRAIS OPERATEUR",
        "FRAIS BANQUE",
        "CIONS PERÇUES",
        "CIONS RETROCEDES",
        "CIONS NETS",
    ]

    for col in colonnes_numeriques:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


# ============================================================
# 5. COMPILATION (concatène tous les fichiers d'un partenaire)
# ============================================================

def compiler_dataframes(
    df_list: List[pd.DataFrame],
    colonnes_numeriques: Optional[List[str]] = None,
) -> pd.DataFrame:

    cleaned = [nettoyer_dataframe(df, colonnes_numeriques) for df in df_list]
    return pd.concat(cleaned, ignore_index=True)


# ============================================================
# 6. SPLIT W2B / B2W — identique pour tous
# ============================================================

def separer_w2b_b2w(df: pd.DataFrame):
    type_transaction = df["TYPE TRANSACTION"].astype(str).str.upper().str.strip()
    return (
        df[type_transaction == "W2B"].reset_index(drop=True),
        df[type_transaction == "B2W"].reset_index(drop=True),
    )


# ============================================================
# 7. SAUVEGARDE SQLITE — identique pour tous
# ============================================================

def sauvegarder_sqlite(df: pd.DataFrame, table_name: str, engine, log_prefix: str = "excel_common"):
    from common.sqlite_io import ecrire_table
    ecrire_table(df, table_name, engine, log_prefix=log_prefix)


def sauvegarder_excel_w2b_b2w(
    df_final: pd.DataFrame,
    tables: dict,
    engine,
    log_prefix: str = "excel_common",
):
    """Une écriture de la compilation + vues excel_w2b / excel_b2w."""
    from common.sqlite_io import ecrire_avec_vues_w2b_b2w
    return ecrire_avec_vues_w2b_b2w(
        df_final,
        tables["excel"],
        tables["excel_w2b"],
        tables["excel_b2w"],
        engine,
        type_col="TYPE TRANSACTION",
        mode="exact",
        log_prefix=log_prefix,
    )