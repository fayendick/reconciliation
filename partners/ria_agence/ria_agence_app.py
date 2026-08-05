# ============================================================
# APP RIA - AGENCES — service Excel dédié au mapping des agences RIA
# ------------------------------------------------------------
# VERSION ADAPTÉE pour le fichier export "iplaorder"
# (ex: "..._TransactionDetailReport_...xls"), copié/adapté à
# partir de la version précédente de ria_agence_app.py.
# ============================================================
# uvicorn ria_agence_app:app --port 8032
#
# Ce service reçoit le(s) fichier(s) Excel "RIA / iplaorder"
# (une ligne par transaction, colonnes : SC Numéro du transfert,
# Pin, Mode de livraison, Guichetier, Succursale, Code d'agence,
# Sent Amount, Sending Currency, Pays d'origine, Pays de
# destination, Montant du paiement, Devise du Bénéficiaire,
# Commission SA, Devise Comission SA, Date, Taux, TOB, TTHU,
# TVA2, TTA, Montant a payer, Frais Client, Action) et :
#
#   1. Rapproche chaque ligne du référentiel
#      mapping_agences_code_ria.xlsx pour retrouver CODE_AGENCE
#      (match exact sur "Code d'agence" <-> CODE_COFINA, repli en
#      fuzzy matching sur "Succursale" — inchangé par rapport à la
#      version précédente).
#
#   2. Calcule DEUX colonnes montant, selon le sens de l'opération
#      (colonne "Action" du fichier iplaorder : "Envoyé" ou
#      "Payé") :
#
#        Montant A Envoyé =
#            Sent Amount - (TOB + TTHU + TVA2 + TTA + Frais Client)
#            SI Sent Amount > 0 OU Action == "Envoyé"
#            SINON 0
#
#        Montant A Payé =
#            Montant a payer
#            SI Action == "Payé"
#            SINON 0
#
#      Ce sont ces deux colonnes (agrégées par CODE_AGENCE) qui
#      seront comparées au Flex Oracle par
#      reconciliation_engine.reconcilier_par_agence() :
#          Montant A Envoyé (Partenaire) <-> CREDIT (Flex)
#          Montant A Payé   (Partenaire) <-> DEBIT  (Flex)
#      -> voir config.COLONNES_AGENCE_RIA_AGENCE, à mettre à jour
#         avec ces deux noms de colonnes (voir bas de ce fichier).
#
#   3. Le résultat final embarque, EN PLUS des colonnes d'origine
#      du fichier iplaorder, les colonnes du référentiel de
#      mapping qui n'existent pas déjà dans le fichier uploadé :
#      CODE_AGENCE, NOM_AGENCE_MAPPEE, CODE_COFINA,
#      SUCCURSALE_MONF (+ METHODE_MATCH / SCORE, à titre de
#      diagnostic du matching).
#
# lire_fichier_ria() ne retire AUCUNE colonne d'origine du fichier
# iplaorder : elles sont toutes conservées telles quelles.
# ============================================================

import io
import traceback
import unicodedata
from typing import List

import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import StreamingResponse
from rapidfuzz import process, fuzz

from config import DB_PATH, MAPPING_AGENCES_RIA_PATH, get_partenaire, make_sqlite_engine
from common.http_export import respond_sheets
from common.sqlite_io import ecrire_table, lire_table_json

PARTENAIRE = "RIA_AGENCE"
TABLES = get_partenaire(PARTENAIRE)["tables"]
TABLE_SORTIE = TABLES["excel"]

SEUIL_SCORE = 75

# ------------------------------------------------------------
# Colonnes montant du fichier iplaorder à convertir en numérique +
# valeur absolue (le fichier source peut contenir des montants
# négatifs selon le sens de l'opération, ex: "Montant du paiement" ;
# on ne garde ici que la valeur, jamais le signe). C'est à partir
# de ces colonnes déjà "nettoyées" que sont calculées Montant A
# Envoyé / Montant A Payé plus bas.
# ------------------------------------------------------------
COLONNES_MONTANT_RIA = [
    "Sent Amount",
    "Montant du paiement",
    "Commission SA",
    "TOB",
    "TTHU",
    "TVA2",
    "TTA",
    "Montant a payer",
    "Frais Client",
]

# Colonnes obligatoires pour pouvoir mapper l'agence ET calculer
# les deux montants de réconciliation.
COLONNES_OBLIGATOIRES_RIA = [
    "Succursale",
    "Code d'agence",
    "Sent Amount",
    "TOB",
    "TTHU",
    "TVA2",
    "TTA",
    "Frais Client",
    "Montant a payer",
    "Action",
]

# Noms des deux colonnes calculées, utilisées par
# reconciliation_engine.reconcilier_par_agence() via
# config.COLONNES_AGENCE_RIA_AGENCE.
COL_MONTANT_A_ENVOYE = "Montant A Envoyé"
COL_MONTANT_A_PAYE = "Montant A Payé"

# Colonnes du référentiel de mapping à réinjecter dans le résultat,
# en plus de CODE_AGENCE déjà utilisée pour le matching (aucune de
# ces colonnes n'existe dans le fichier iplaorder d'origine).
COLONNES_MAPPING_A_AJOUTER = ["NOM_AGENCE_MAPPEE", "CODE_COFINA", "SUCCURSALE_MONF"]

print(f"[ria_agence_app.py] Base SQLite utilisée : {DB_PATH}")
print(f"[ria_agence_app.py] Fichier mapping utilisé : {MAPPING_AGENCES_RIA_PATH}")

engine = make_sqlite_engine()

app = FastAPI(title="Excel Upload API — RIA Agences (iplaorder)")


# ==========================================================
# NORMALISATION
# ==========================================================

def normaliser(txt):

    if pd.isna(txt):
        return ""

    txt = str(txt).upper()

    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")

    txt = txt.replace("'", " ")
    txt = txt.replace("-", " ")
    txt = txt.replace("/", " ")

    while "  " in txt:
        txt = txt.replace("  ", " ")

    return txt.strip()


def normaliser_code(txt):
    """Normalisation stricte pour comparer un code d'agence
    (Code d'agence <-> CODE_COFINA) : majuscule + espaces retirés,
    sans toucher aux caractères internes (COFI005 doit rester
    COFI005)."""
    if pd.isna(txt):
        return ""
    return str(txt).strip().upper()


def normaliser_action(txt):
    """'Envoyé' / 'envoyé' / 'ENVOYE' / ' Envoyé ' -> 'ENVOYE' ;
    'Payé' / 'payé' / 'PAYE' -> 'PAYE'. Insensible à la casse et
    aux accents, pour ne jamais rater une ligne à cause d'une
    variante de saisie du fichier source."""
    return normaliser(txt).replace(" ", "")


# ==========================================================
# LECTURE DU FICHIER IPLAORDER / RIA (recherche dynamique de l'entête)
# ------------------------------------------------------------
# Cherche la ligne d'entête contenant "Code d'agence" (le fichier
# peut avoir quelques lignes de titre avant le vrai tableau).
# Toutes les colonnes du fichier sont conservées telles quelles à
# partir de là, puis les deux colonnes Montant A Envoyé / Montant A
# Payé sont calculées.
# ==========================================================

def lire_fichier_ria(fichier: UploadFile) -> pd.DataFrame:

    contenu = fichier.file.read()

    try:
        tmp = pd.read_excel(io.BytesIO(contenu), header=None)
    except Exception:
        raise HTTPException(
            status_code=400,
            detail=f"Impossible de lire le fichier {fichier.filename} (format invalide)."
        )

    header = None
    for i, row in tmp.iterrows():
        valeurs = row.astype(str).str.strip().tolist()
        if "Code d'agence" in valeurs or "Code d’agence" in valeurs:
            header = i
            break

    if header is None:
        raise HTTPException(
            status_code=400,
            detail=f"Ligne d'entête \"Code d'agence\" introuvable dans le fichier {fichier.filename}."
        )

    ria = pd.read_excel(io.BytesIO(contenu), header=header)
    ria.columns = ria.columns.str.strip()

    # tolère l'apostrophe typographique ("Code d’agence")
    ria.columns = [c.replace("’", "'") for c in ria.columns]

    manquantes = [c for c in COLONNES_OBLIGATOIRES_RIA if c not in ria.columns]
    if manquantes:
        raise HTTPException(
            status_code=400,
            detail=f"Colonne(s) manquante(s) dans le fichier {fichier.filename} : {manquantes}"
        )

    # Montants en numérique + valeur absolue (le fichier source peut
    # contenir des négatifs selon le sens de l'opération ; on ne garde
    # que la valeur, le sens est porté par la colonne "Action").
    for col in COLONNES_MONTANT_RIA:
        if col in ria.columns:
            ria[col] = pd.to_numeric(ria[col], errors="coerce").fillna(0).abs()

    # --------------------------------------------------------
    # Calcul des deux colonnes de réconciliation.
    # --------------------------------------------------------
    action_norm = ria["Action"].apply(normaliser_action)
    est_envoye = (ria["Sent Amount"] > 0) | (action_norm == "ENVOYE")
    est_paye = action_norm == "PAYE"
    
    
    

    frais_envoi = ria["TOB"] + ria["TTHU"] + ria["TVA2"] + ria["TTA"] + ria["Frais Client"]

    ria[COL_MONTANT_A_ENVOYE] = 0.0
    ria.loc[est_envoye, COL_MONTANT_A_ENVOYE] = (
        ria.loc[est_envoye, "Sent Amount"] + frais_envoi.loc[est_envoye]
    ).abs()

    #ria[COL_MONTANT_A_PAYE] = 0.0
    #ria.loc[est_paye, COL_MONTANT_A_PAYE] = ria.loc[est_paye, "Montant a payer"].abs()
    
    ria[COL_MONTANT_A_PAYE] = (
    ria["Montant a payer"]
        .where(ria["Montant a payer"] > 0, 0)
        .abs()
    )
    
    
    

    ria["CLE_SUCCURSALE"] = ria["Succursale"].apply(normaliser)
    ria["CLE_CODE"] = ria["Code d'agence"].apply(normaliser_code)
    ria["FICHIER_SOURCE"] = fichier.filename

    return ria


# ==========================================================
# LECTURE DU MAPPING
# ==========================================================

def lire_mapping() -> pd.DataFrame:

    try:
        mapping = pd.read_excel(MAPPING_AGENCES_RIA_PATH)
    except FileNotFoundError:
        raise HTTPException(
            status_code=500,
            detail=f"Fichier de mapping introuvable : {MAPPING_AGENCES_RIA_PATH}"
        )

    mapping.columns = mapping.columns.str.strip()

    for col in ("CODE_AGENCE", "NOM_AGENCE_MAPPEE", "CODE_COFINA"):
        if col not in mapping.columns:
            raise HTTPException(
                status_code=500,
                detail=f"Colonne '{col}' manquante dans le fichier de mapping."
            )

    mapping["CLE_CODE"] = mapping["CODE_COFINA"].apply(normaliser_code)

    # Clé de nom pour le repli fuzzy : priorité au libellé du fichier
    # source du référentiel (SUCCURSALE_MONF) quand il existe, sinon
    # le nom mappé (NOM_AGENCE_MAPPEE).
    if "SUCCURSALE_MONF" in mapping.columns:
        mapping["CLE_NOM"] = mapping["SUCCURSALE_MONF"].where(
            mapping["SUCCURSALE_MONF"].notna(), mapping["NOM_AGENCE_MAPPEE"]
        ).apply(normaliser)
    else:
        mapping["CLE_NOM"] = mapping["NOM_AGENCE_MAPPEE"].apply(normaliser)

    return mapping


# ==========================================================
# MATCHING : code exact d'abord, fuzzy sur le nom en repli
# ------------------------------------------------------------
# Réinjecte, pour chaque ligne matchée, CODE_AGENCE + toutes les
# colonnes du mapping absentes du fichier iplaorder
# (COLONNES_MAPPING_A_AJOUTER), + METHODE_MATCH/SCORE en
# diagnostic.
# ==========================================================

def appliquer_matching_ria(ria: pd.DataFrame, mapping: pd.DataFrame) -> pd.DataFrame:

    mapping_par_code = {
        code: ligne for code, ligne in zip(mapping["CLE_CODE"], mapping.to_dict("records"))
        if code
    }
    liste_noms_mapping = mapping["CLE_NOM"].tolist()

    colonnes_mapping_dispo = [c for c in COLONNES_MAPPING_A_AJOUTER if c in mapping.columns]

    codes = []
    extra = {c: [] for c in colonnes_mapping_dispo}
    methodes, scores = [], []

    def _ajouter_ligne_vide():
        codes.append(None)
        for c in colonnes_mapping_dispo:
            extra[c].append(None)
        methodes.append(None)

    def _ajouter_ligne(ligne):
        codes.append(ligne["CODE_AGENCE"])
        for c in colonnes_mapping_dispo:
            extra[c].append(ligne.get(c))
        methodes.append(ligne.get("_methode"))

    for _, row in ria.iterrows():

        cle_code = row["CLE_CODE"]
        cle_nom = row["CLE_SUCCURSALE"]

        # 1) match exact sur le code (Code d'agence <-> CODE_COFINA)
        if cle_code and cle_code in mapping_par_code:
            ligne = dict(mapping_par_code[cle_code])
            ligne["_methode"] = "CODE"
            _ajouter_ligne(ligne)
            scores.append(100)
            continue

        # 2) repli : fuzzy matching sur le nom de la Succursale
        meilleur = process.extractOne(
            cle_nom,
            liste_noms_mapping,
            scorer=fuzz.token_sort_ratio
        ) if cle_nom else None

        if meilleur is None:
            _ajouter_ligne_vide()
            scores.append(0)
            continue

        nom_trouve, score = meilleur[0], meilleur[1]

        if score >= SEUIL_SCORE:
            ligne = mapping[mapping["CLE_NOM"] == nom_trouve].iloc[0].to_dict()
            ligne["_methode"] = "FUZZY"
            _ajouter_ligne(ligne)
            scores.append(score)
        else:
            _ajouter_ligne_vide()
            scores.append(score)

    ria["CODE_AGENCE"] = codes
    for c in colonnes_mapping_dispo:
        ria[c] = extra[c]
    ria["METHODE_MATCH"] = methodes
    ria["SCORE"] = scores

    ordre = ["CODE_AGENCE"] + colonnes_mapping_dispo + ["METHODE_MATCH", "SCORE"]
    ordre += [c for c in ria.columns if c not in ordre]

    return ria[ordre]


# ==========================================================
# SAUVEGARDE SQLITE
# ==========================================================

def sauvegarder_sqlite(df: pd.DataFrame, table: str, source: str):
    ecrire_table(df, table, engine, log_prefix=source)


# ==========================================================
# ENDPOINTS
# ==========================================================

def nom_feuille(nom_fichier: str, index: int) -> str:
    """Nom de feuille Excel valide (<=31 caractères, sans caractères interdits)."""
    base = nom_fichier.rsplit(".", 1)[0]
    for caractere in ["\\", "/", "?", "*", "[", "]", ":"]:
        base = base.replace(caractere, "_")
    base = base.strip() or f"Fichier_{index}"
    return base[:28] + f"_{index}" if len(base) > 28 else base


@app.post("/map-agences-ria")
async def map_agences_ria(
    files: List[UploadFile] = File(...),
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):

    try:
        mapping = lire_mapping()

        # ------------------------------------------------------
        # 1. TRAITEMENT INDIVIDUEL DE CHAQUE FICHIER
        # ------------------------------------------------------
        dfs_traites = []
        for fichier in files:
            ria = lire_fichier_ria(fichier)
            ria_mappe = appliquer_matching_ria(ria, mapping)
            dfs_traites.append((fichier.filename, ria_mappe))

        # ------------------------------------------------------
        # 2. COMBINAISON DES FICHIERS TRAITÉS (avant réconciliation)
        # ------------------------------------------------------
        df_combine = pd.concat(
            [df for _, df in dfs_traites],
            ignore_index=True
        )

        sauvegarder_sqlite(df_combine, TABLE_SORTIE, "ria_agence_app")

        # ------------------------------------------------------
        # 3. EXPORT : Excel multi-feuilles ou JSON léger (/charger)
        # ------------------------------------------------------
        sheets = {"Compilation": df_combine}
        for i, (nom_fichier, df) in enumerate(dfs_traites, start=1):
            sheets[nom_feuille(nom_fichier, i)] = df

        return respond_sheets(
            sheets if format.strip().lower() not in {"json", "none", "sqlite"} else {"Compilation": df_combine},
            filename="RIA_MAPPE.xlsx",
            format=format,
            json_payload={
                "status": "ok",
                "format": "json",
                "filename": "RIA_MAPPE.xlsx",
                "nb_lignes": int(len(df_combine)),
                "nb_fichiers": len(dfs_traites),
            },
        )

    except HTTPException:
        raise
    except Exception as e:
        return {"status": "error", "message": str(e), "trace": traceback.format_exc()}


@app.get("/db/ria-agence")
def get_ria_agence(limit: int = Query(None), offset: int = Query(0)):
    return lire_table_json(engine, TABLE_SORTIE, limit=limit, offset=offset)


@app.get("/health")
def health():
    return {
        "status": "ok",
        "partenaire": PARTENAIRE,
        "db_path": DB_PATH,
        "mapping": MAPPING_AGENCES_RIA_PATH,
        "colonnes_reconciliation": {
            "col_credit": COL_MONTANT_A_ENVOYE,
            "col_debit": COL_MONTANT_A_PAYE,
        },
    }