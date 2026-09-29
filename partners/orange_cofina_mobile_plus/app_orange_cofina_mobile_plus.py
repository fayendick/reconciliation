import io
import re
import unicodedata
from typing import List

import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException, Query

from config import DB_PATH, PARTENAIRES, make_sqlite_engine
from common.excel_common import (
    valider_colonnes_standard,
    compiler_dataframes,
    separer_w2b_b2w,
    sauvegarder_excel_w2b_b2w,
)
from common.http_export import respond_sheets


PARTENAIRE = "ORANGE_COFINA_MOBILE_PLUS"
TABLES = PARTENAIRES[PARTENAIRE]["tables"]

print(
    f"[app_orange_cofina_mobile_plus] Base SQLite utilisée : {DB_PATH}"
)

engine = make_sqlite_engine()

app = FastAPI(
    title="Excel Upload API — Orange API / Cofina Mobile Plus"
)


# ============================================================
# OUTILS
# ============================================================

def _normaliser_texte(valeur) -> str:
    if valeur is None:
        return ""

    texte = str(valeur).strip().upper()
    texte = unicodedata.normalize("NFKD", texte)
    texte = "".join(
        c for c in texte
        if not unicodedata.combining(c)
    )

    return " ".join(texte.split())


def _normaliser_colonne(valeur) -> str:
    texte = _normaliser_texte(valeur)
    texte = re.sub(r"[^A-Z0-9]+", " ", texte)

    return " ".join(texte.split())


def _nettoyer_numero(valeur) -> str:
    if pd.isna(valeur):
        return ""

    texte = str(valeur).strip()

    if texte.endswith(".0") and texte[:-2].isdigit():
        texte = texte[:-2]

    return re.sub(r"\D", "", texte)


def _montant(valeur):
    if pd.isna(valeur) or str(valeur).strip() == "":
        return 0.0

    if isinstance(valeur, (int, float)):
        return float(valeur)

    texte = str(valeur).strip().replace("\u00a0", "")
    texte = texte.replace(" ", "")

    # Rapports Orange :
    # "10 000" / "1 000 000"
    if "," in texte and "." not in texte:
        texte = texte.replace(",", ".")
    else:
        texte = texte.replace(",", "")

    try:
        return float(texte)
    except ValueError:
        return 0.0


def _trouver_ligne_entete(brut: pd.DataFrame) -> int:
    mots = {
        "DATE",
        "HEURE",
        "REFERENCE",
        "SERVICE",
        "STATUT",
    }

    limite = min(len(brut), 80)

    for i in range(limite):
        valeurs = {
            _normaliser_colonne(v)
            for v in brut.iloc[i].tolist()
        }

        score = len(mots.intersection(valeurs))

        if score >= 4:
            return i

    raise HTTPException(
        status_code=400,
        detail=(
            "Impossible de trouver la ligne d'en-têtes "
            "du fichier Orange API / Cofina Mobile Plus."
        ),
    )


def _charger_feuille_excel(contenu: bytes) -> pd.DataFrame:
    fichier = io.BytesIO(contenu)

    xls = pd.ExcelFile(fichier)

    if not xls.sheet_names:
        raise HTTPException(
            status_code=400,
            detail="Le fichier Excel ne contient aucune feuille.",
        )

    # Le rapport Orange attendu est sur la première feuille utile.
    brut = pd.read_excel(
        fichier,
        sheet_name=xls.sheet_names[0],
        header=None,
    )

    if brut.empty:
        raise HTTPException(
            status_code=400,
            detail="Le fichier Orange API est vide.",
        )

    ligne_entete = _trouver_ligne_entete(brut)

    df = brut.iloc[ligne_entete + 1:].copy()

    noms = []
    compte_occurrences = {}

    for valeur in brut.iloc[ligne_entete].tolist():

        base = (
            str(valeur).strip()
            if pd.notna(valeur)
            else ""
        )

        normalise = _normaliser_colonne(base)

        compte_occurrences[normalise] = (
            compte_occurrences.get(normalise, 0) + 1
        )

        if normalise == "N DE COMPTE":
            # Le rapport Orange comporte deux N° de Compte :
            # Agent puis Correspondant.
            # On garde des noms explicites pour sélectionner
            # le correspondant.
            noms.append(
                f"N° de Compte {compte_occurrences[normalise]}"
            )
        else:
            noms.append(base)

    df.columns = noms

    df = df.dropna(how="all").copy()

    return df


def _colonne(
    df: pd.DataFrame,
    candidats: List[str],
    obligatoire=True,
):
    index = {
        _normaliser_colonne(c): c
        for c in df.columns
    }

    for candidat in candidats:

        cle = _normaliser_colonne(candidat)

        if cle in index:
            return index[cle]

    if obligatoire:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Colonne Orange API introuvable parmi {candidats}. "
                f"Colonnes disponibles : {list(df.columns)}"
            ),
        )

    return None


def nettoyer_et_mapper(contenu: bytes) -> pd.DataFrame:

    df_brut = _charger_feuille_excel(contenu)

    # ========================================================
    # IDENTIFICATION DES COLONNES DU FICHIER ORANGE
    # ========================================================

    c_date = _colonne(
        df_brut,
        ["Date"],
    )

    c_heure = _colonne(
        df_brut,
        ["Heure"],
    )

    c_reference = _colonne(
        df_brut,
        ["Référence", "Reference"],
    )

    c_service = _colonne(
        df_brut,
        ["Service"],
    )

    c_statut = _colonne(
        df_brut,
        ["Statut"],
    )

    c_debit = _colonne(
        df_brut,
        ["Débit", "Debit"],
        obligatoire=False,
    )

    c_credit = _colonne(
        df_brut,
        ["Crédit", "Credit"],
        obligatoire=False,
    )

       # ========================================================
    # TELEPHONE
    # ========================================================
    # Le fichier contient deux colonnes "N° de Compte".
    #
    # N° de Compte 1 = compte Orange Money de COFINA
    # N° de Compte 2 = numéro du correspondant
    #
    # Le rapprochement doit utiliser le numéro du correspondant.
    # ========================================================

    c_tel = None

    for c in df_brut.columns:
        if _normaliser_colonne(c) == "N DE COMPTE 2":
            c_tel = c
            break

    if c_tel is None:
        raise ValueError(
            "La colonne 'N° de Compte 2' est introuvable "
            "dans le fichier Orange Cofina Mobile Plus."
        )
    # ========================================================
    # FILTRE DES TRANSACTIONS REUSSIES
    # ========================================================

    statut = df_brut[c_statut].map(
        _normaliser_texte
    )

    succes = statut.str.contains(
        r"SUCCES|SUCCESS|REUSSI",
        regex=True,
        na=False,
    )

    df = df_brut.loc[succes].copy()

    if df.empty:

        return pd.DataFrame(
            columns=[
                "DATE TRANSACTION",
                "TYPE TRANSACTION",
                "CODE TRANSACTION OPERATEUR",
                "NUMERO COMPTE",
                "MONTANT",
            ]
        )

    # ========================================================
    # NORMALISATION DU SERVICE
    # ========================================================

    service = df[c_service].map(
        _normaliser_texte
    )

    # ========================================================
    # EXCLUSION O2C TRANSFER
    # ========================================================
    #
    # O2C Transfer = approvisionnement.
    #
    # Ce n'est pas une transaction W2B/B2W.
    #
    # Le rapport ne fournit pas de téléphone correspondant
    # pour cette opération.
    #
    # On l'exclut donc AVANT le contrôle des services.
    # ========================================================

    masque_o2c = service.eq(
        "O2C TRANSFER"
    )

    if masque_o2c.any():

        df = df.loc[
            ~masque_o2c
        ].copy()

        service = service.loc[
            ~masque_o2c
        ].copy()

    # ========================================================
    # NORMALISATION CASH IN / CASH OUT
    # ========================================================

    service = service.replace(
        {
            "CASHIN": "W2B",
            "CASH IN": "W2B",
            "CASH-IN": "W2B",
            "CASHOUT": "B2W",
            "CASH OUT": "B2W",
            "CASH-OUT": "B2W",
        }
    )

    # ========================================================
    # CONTROLE DES SERVICES
    # ========================================================

    inconnus = sorted(
        v
        for v in service.dropna().unique()
        if v not in {"W2B", "B2W"}
    )

    if inconnus:

        raise HTTPException(
            status_code=400,
            detail=(
                "Services de transactions Orange API "
                "non reconnus : "
                f"{inconnus}. "
                "Aucun rapprochement n'a été lancé."
            ),
        )

    # ========================================================
    # DATE
    # ========================================================

    date_part = pd.to_datetime(
        df[c_date],
        errors="coerce",
        dayfirst=True,
    )

    # ========================================================
    # HEURE
    # ========================================================

    heure_brute = df[c_heure]

    heure_part = pd.to_timedelta(
        heure_brute.astype(str).str.strip(),
        errors="coerce",
    )

    # Si Excel fournit l'heure comme datetime
    # (date 1899 + heure), on ne garde que l'heure.
    try:

        heure_dt = pd.to_datetime(
            heure_brute,
            errors="coerce",
        )

        masque_dt = heure_dt.notna()

        heure_part.loc[masque_dt] = (
            heure_dt.loc[masque_dt]
            - heure_dt.loc[masque_dt].dt.normalize()
        )

    except Exception:
        pass

    # ========================================================
    # NORMALISATION VERS LE SCHEMA STANDARD DU MOTEUR
    # ========================================================
    #
    # Fichier Orange Cofina :
    #
    #   DATE
    #   HEURE
    #   REFERENCE
    #   SERVICE
    #   STATUT
    #   N° DE COMPTE
    #   DEBIT
    #   CREDIT
    #
    # Schema attendu par le moteur :
    #
    #   DATE TRANSACTION
    #   TYPE TRANSACTION
    #   CODE TRANSACTION OPERATEUR
    #   NUMERO COMPTE
    #   MONTANT
    #
    # ========================================================

    resultat = pd.DataFrame(
        index=df.index
    )

    # Date + heure de la transaction
    resultat["DATE TRANSACTION"] = (
        date_part + heure_part
    )

    # Type de transaction
    #
    # Le service a déjà été normalisé :
    # CASH IN  -> W2B
    # CASH OUT -> B2W
    resultat["TYPE TRANSACTION"] = service

    # Référence opérateur.
    # Cette colonne reste informative.
    resultat[
        "CODE TRANSACTION OPERATEUR"
    ] = (
        df[c_reference]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Numéro de téléphone / compte correspondant
    resultat[
        "NUMERO COMPTE"
    ] = df[c_tel].map(
        _nettoyer_numero
    )

    # ========================================================
    # CALCUL DU MONTANT
    # ========================================================

    if c_debit:

        debit = df[c_debit].map(
            _montant
        )

    else:

        debit = pd.Series(
            0.0,
            index=df.index,
        )

    if c_credit:

        credit = df[c_credit].map(
            _montant
        )

    else:

        credit = pd.Series(
            0.0,
            index=df.index,
        )

    # Le moteur utilise une seule colonne MONTANT.
    #
    # Pour une transaction W2B :
    #     DEBIT
    #
    # Pour une transaction B2W :
    #     CREDIT
    #
    # On prend donc DEBIT lorsqu'il est renseigné,
    # sinon CREDIT.

    resultat["MONTANT"] = debit.where(
        debit.ne(0),
        credit,
    )

    # ========================================================
    # ELIMINATION DES LIGNES TECHNIQUEMENT INEXPLOITABLES
    # ========================================================

    resultat = resultat[
        resultat["DATE TRANSACTION"].notna()
        & resultat["NUMERO COMPTE"].ne("")
        & resultat["MONTANT"].notna()
    ].copy()

    return resultat.reset_index(
        drop=True
    )


def lire_et_mapper_fichiers(
    files: List[UploadFile],
) -> List[pd.DataFrame]:

    dfs = []

    for fichier in files:

        contenu = fichier.file.read()

        dfs.append(
            nettoyer_et_mapper(
                contenu
            )
        )

    return dfs


# ============================================================
# API
# ============================================================

@app.post("/process-excel")
async def process_excel(
    files: List[UploadFile] = File(...),
    format: str = Query(
        "excel",
        description="excel (défaut) ou json",
    ),
):

    if not files:

        raise HTTPException(
            status_code=400,
            detail="Aucun fichier Orange API fourni.",
        )

    try:

        # ====================================================
        # LECTURE ET NORMALISATION
        # ====================================================

        dfs = lire_et_mapper_fichiers(
            files
        )

        df_final = compiler_dataframes(
            dfs
        )

        # ====================================================
        # CONTROLE DU RESULTAT
        # ====================================================

        if df_final.empty:

            raise HTTPException(
                status_code=400,
                detail=(
                    "Aucune transaction réussie "
                    "exploitable n'a été trouvée "
                    "dans le fichier Orange API."
                ),
            )

        # ====================================================
        # VALIDATION DU SCHEMA STANDARD
        # ====================================================

        valider_colonnes_standard(
            df_final,
            PARTENAIRE,
        )

        # ====================================================
        # SEPARATION W2B / B2W
        # ====================================================

        w2b, b2w = separer_w2b_b2w(
            df_final
        )

        # ====================================================
        # SAUVEGARDE
        # ====================================================

        sauvegarder_excel_w2b_b2w(
            df_final,
            TABLES,
            engine,
            "app_orange_cofina_mobile_plus",
        )

        print(
            "[app_orange_cofina_mobile_plus] "
            f"Chargement réussi : "
            f"{len(df_final)} lignes "
            f"({len(w2b)} W2B / {len(b2w)} B2W)"
        )

        # ====================================================
        # REPONSE
        # ====================================================

        return respond_sheets(
            {
                "Compilation": df_final,
                "Compilation_W2B": w2b,
                "Compilation_B2W": b2w,
            },
            filename=(
                "Compilation_"
                "ORANGE_COFINA_MOBILE_PLUS.xlsx"
            ),
            format=format,
        )

    except HTTPException:
        raise

    except Exception as exc:

        raise HTTPException(
            status_code=500,
            detail=(
                "Erreur chargement Orange API / "
                f"Cofina Mobile Plus : {exc}"
            ),
        )