# ============================================================
# APP WAVE - AGENCES — service Excel dédié au mapping des agences Wave
# ============================================================
# uvicorn wave_app_agence:app --port 8030
#
# Ce service reçoit le(s) fichier(s) Excel "Wave agence" au format
# TRANSACTIONNEL (relevé d'opérations), c'est-à-dire un tableau qui
# commence par la ligne d'entête :
#
#   Quand | Référence | Quoi | Montant (CFA) | ID Wave et Nom de
#   l'Agent | Contrepartie | Opérateur | Solde (CFA)
#
# (le fichier peut contenir quelques lignes de titre/logo avant
# cette ligne d'entête — elle est donc recherchée dynamiquement,
# comme avant, en cherchant "Quand").
#
# Traitement effectué pour chaque ligne :
#
#   1) La colonne utilisée pour le fuzzy matching est "Opérateur"
#      (ex: "Cofina Kaolack", "COFINA RUFISQUE", "*Cofina Express
#      Pikine*"...), et NON PAS "ID Wave et Nom de l'Agent". Le
#      référentiel mapping_agences_cofina_wave.xlsx contient :
#         Opérateur          -> libellé brut Wave, sert UNIQUEMENT
#                                 au fuzzy matching, à comparer à la
#                                 colonne "Opérateur" du fichier
#                                 partenaire.
#         CODE_AGENCE (ex: 512, 509, 523…) -> code agence numérique
#                                 Cofina, comparé au code agence
#                                 renvoyé par Flex (également
#                                 numérique). On n'utilise JAMAIS un
#                                 code du type "COFI005".
#         NOM_AGENCE_MAPPEE  -> libellé lisible de l'agence, pour
#                                 info uniquement.
#      La colonne "ID Wave et Nom de l'Agent" reste présente dans le
#      détail par transaction (colonne d'origine, jamais supprimée),
#      mais n'est PLUS utilisée pour le rapprochement.
#
#   2) CODE_AGENCE est toujours forcé en entier "Int64" nullable
#      (jamais en flottant type 512.0, jamais en texte type
#      "COFI005") afin qu'il se compare proprement au code Flex lors
#      de la réconciliation.
#
#   3) La colonne "Quoi" pilote la création de deux nouvelles
#      colonnes de montants, calculées en VALEUR ABSOLUE à partir de
#      "Montant (CFA)" (les montants Wave peuvent être négatifs,
#      typiquement les retraits — abs() les neutralise) :
#         - Quoi == "Retrait" -> colonne RETRAIT = abs(Montant (CFA))
#           (sera rapprochée du DEBIT renvoyé par wave_agence_flex_api.py)
#         - Quoi == "Dépôt"   -> colonne DEPOT   = abs(Montant (CFA))
#           (sera rapprochée du CREDIT renvoyé par wave_agence_flex_api.py)
#      Les autres valeurs de "Quoi" (ex: "Achat des UV", "Remboursement
#      des UV") laissent RETRAIT et DEPOT à 0 pour la ligne concernée.
#      [CORRECTIF] RETRAIT/DEPOT ne sont QUE des colonnes informatives
#      (répartition partielle). Le SOLDE_PARTENAIRE réel par agence
#      (voir agreger_par_agence) est calculé à partir du Montant (CFA)
#      SIGNÉ, sommé sur TOUTES les catégories de "Quoi" — sinon les
#      catégories hors Retrait/Dépôt (Achat des UV, Remboursement des
#      UV...) sont perdues et le solde par agence est sous-évalué.
#
#   4) La ligne technique dont la colonne "ID Wave et Nom de l'Agent"
#      correspond à "A12956_SN3 - COFINA SENEGAL" est exclue du
#      fichier avant tout traitement : ce n'est pas une vraie agence
#      (compte technique/pivot Cofina Sénégal), et sa présence
#      faussait les totaux rapprochés avec Flex. La comparaison est
#      faite sur la clé normalisée (insensible à la casse/accents/
#      espaces) et repère la ligne dès que "A12956_SN3" est présent
#      dans la colonne, pour rester robuste à de petites variations
#      de libellé (espaces, tirets...).
#
#   5) Après le fuzzy matching, les lignes dont l'agence n'a PAS été
#      trouvée (CODE_AGENCE = None, score < SEUIL_SCORE ou
#      "Opérateur" vide) sont exclues du résultat (feuilles Excel +
#      table SQLite + Compilation). Ces lignes non matchées n'ont pas
#      de contrepartie fiable côté Flex et faussaient les sommes par
#      agence utilisées pour la réconciliation.
#
# ------------------------------------------------------------
# [CORRECTIF - agrégat par agence]
# ------------------------------------------------------------
# AVANT : SOLDE_PARTENAIRE était calculé avec
#         `wave.groupby("CODE_AGENCE")[...].transform("sum")`.
#         `.transform()` NE RÉDUIT PAS le nombre de lignes : la même
#         somme se retrouvait donc recopiée sur CHAQUE ligne de
#         transaction de l'agence -> une agence avec 40 transactions
#         apparaissait 40 fois avec le même solde répété (ce qui
#         ressemblait à des doublons).
#
# APRÈS : une nouvelle fonction agreger_par_agence() fait un vrai
#         `.groupby("CODE_AGENCE").agg(...)` -> UNE SEULE ligne par
#         agence (ex: 501, 4083235...), avec RETRAIT/DEPOT sommés et
#         SOLDE_PARTENAIRE = RETRAIT - DEPOT (même convention que
#         reconciliation_engine.preparer_agence_partenaire, qui fait
#         col_debit - col_credit = RETRAIT - DEPOT).
#
#         C'est ce DataFrame agrégé (1 ligne/agence) qui est :
#           - sauvegardé dans la table SQLite TABLE_SORTIE
#             (COMPILATION_WAVE_AGENCE), donc ce que lit la
#             réconciliation et l'endpoint /db/wave-agence.
#           - exporté dans la feuille Excel "Compilation".
#
#         Le détail transaction par transaction (matché) reste
#         disponible, pour l'audit, dans :
#           - la feuille Excel "Detail_Compilation" (tous fichiers
#             combinés)
#           - une feuille par fichier envoyé (comme avant)
#         mais n'est plus ce qui est utilisé pour la réconciliation.
#
# Aucune colonne d'origine du fichier (Quand, Référence, Quoi,
# Montant (CFA), ID Wave et Nom de l'Agent, Contrepartie, Opérateur,
# Solde (CFA)) n'est supprimée ni renommée dans le détail.
#
# Plusieurs fichiers peuvent être envoyés en une seule fois
# (/map-agences accepte une liste de fichiers) : chacun est lu et
# matché INDIVIDUELLEMENT, puis tous les détails sont combinés et
# c'est cette combinaison qui est agrégée par agence avant d'être
# sauvegardée en base et exposée à la réconciliation (voir
# reconcilier_par_agence() côté config / reconciliation_engine).
#
# IMPORTANT côté config.py : COLONNES_AGENCE_WAVE_AGENCE doit
# pointer sur "RETRAIT" (face au DEBIT flex) et "DEPOT" (face au
# CREDIT flex).
# ============================================================

import io
import re
import traceback
import unicodedata
from typing import List, Optional

import numpy as np
import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import StreamingResponse
from rapidfuzz import process, fuzz

from config import DB_PATH, MAPPING_AGENCES_WAVE_PATH, get_partenaire, make_sqlite_engine
from common.http_export import respond_sheets, wants_excel
from common.sqlite_io import ecrire_table, lire_table_json

PARTENAIRE = "WAVE_AGENCE"
TABLES = get_partenaire(PARTENAIRE)["tables"]
TABLE_SORTIE = TABLES["excel"]

SEUIL_SCORE = 75

# Libellés (une fois normalisés via normaliser()) utilisés pour repérer
# les retraits / dépôts dans la colonne "Quoi".
LIBELLE_RETRAIT = "RETRAIT"
LIBELLE_DEPOT = "DEPOT"

# Ligne technique à exclure systématiquement du relevé Wave (compte
# pivot Cofina Sénégal, ce n'est pas une vraie agence). On ne compare
# que sur "A12956_SN3" (après normalisation) pour rester robuste aux
# petites variations de libellé autour de ce code.
CLE_LIGNE_TECHNIQUE_EXCLUE = "A12956_SN3"

print(f"[wave_app_agence.py] Base SQLite utilisée : {DB_PATH}")
print(f"[wave_app_agence.py] Fichier mapping utilisé : {MAPPING_AGENCES_WAVE_PATH}")

engine = make_sqlite_engine()

app = FastAPI(title="Excel Upload API — Wave Agences")


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
    txt = txt.replace("’", " ")
    txt = txt.replace("-", " ")
    txt = txt.replace("/", " ")
    txt = txt.replace("*", " ")

    while "  " in txt:
        txt = txt.replace("  ", " ")

    return txt.strip()


# ==========================================================
# LECTURE DU FICHIER WAVE (recherche dynamique de l'entête "Quand")
# ------------------------------------------------------------
# Cherche la ligne d'entête contenant "Quand" (le fichier Wave
# agence peut avoir quelques lignes de titre/logo avant le vrai
# tableau). Toutes les colonnes du fichier (Quand, Référence, Quoi,
# Montant (CFA), ID Wave et Nom de l'Agent, Contrepartie, Opérateur,
# Solde (CFA)) sont conservées telles quelles à partir de là.
# ==========================================================

def _trouver_colonne(colonnes, doit_contenir: List[str]) -> Optional[str]:
    """Retrouve le nom exact d'une colonne à partir de mots-clés
    (insensible à la casse/accents/apostrophes), pour ne pas dépendre
    du caractère exact utilisé pour les accents (Opérateur / Operateur)
    ou l'apostrophe (' vs ’) dans "ID Wave et Nom de l'Agent"."""
    for col in colonnes:
        cle = normaliser(col)
        if all(normaliser(mot) in cle for mot in doit_contenir):
            return col
    return None


def lire_fichier_wave(fichier: UploadFile) -> pd.DataFrame:

    contenu = fichier.file.read()
    if not contenu:
        raise HTTPException(
            status_code=400,
            detail=f"Le fichier {fichier.filename} est vide."
        )

    try:
        tmp = pd.read_excel(io.BytesIO(contenu), header=None)
    except Exception:
        raise HTTPException(
            status_code=400,
            detail=f"Impossible de lire le fichier {fichier.filename} (format invalide)."
        )

    header = None
    for i, row in tmp.iterrows():
        if "Quand" in row.astype(str).tolist():
            header = i
            break

    if header is None:
        raise HTTPException(
            status_code=400,
            detail=f"Ligne 'Quand' introuvable dans le fichier {fichier.filename}."
        )

    wave = pd.read_excel(io.BytesIO(contenu), header=header)
    wave.columns = wave.columns.str.strip()

    # Colonnes attendues (recherchées par mots-clés pour être robuste
    # aux variations d'apostrophe/accents dans les entêtes)
    col_quand = _trouver_colonne(wave.columns, ["Quand"])
    col_quoi = _trouver_colonne(wave.columns, ["Quoi"])
    col_montant = _trouver_colonne(wave.columns, ["Montant"])
    col_agent = _trouver_colonne(wave.columns, ["ID", "Wave", "Agent"])
    col_operateur = _trouver_colonne(wave.columns, ["Operateur"])

    manquantes = [
        nom for nom, col in [
            ("Quand", col_quand),
            ("Quoi", col_quoi),
            ("Montant (CFA)", col_montant),
            ("ID Wave et Nom de l'Agent", col_agent),
            ("Opérateur", col_operateur),
        ] if col is None
    ]
    if manquantes:
        raise HTTPException(
            status_code=400,
            detail=f"Colonne(s) {manquantes} introuvable(s) dans le fichier {fichier.filename}."
        )

    # ------------------------------------------------------
    # Exclusion de la ligne technique "A12956_SN3 - COFINA SENEGAL"
    # repérée via la colonne "ID Wave et Nom de l'Agent". Ce n'est
    # pas une vraie agence : la laisser dans le fichier faussait les
    # totaux comparés à Flex.
    # ------------------------------------------------------
    cle_agent_normalisee = wave[col_agent].apply(normaliser)
    wave = wave[~cle_agent_normalisee.str.contains(CLE_LIGNE_TECHNIQUE_EXCLUE, na=False)].copy()

    # ------------------------------------------------------
    # 1) Clé de matching = colonne "Opérateur" (ex: "Cofina Kaolack",
    #    "*Cofina Express Pikine*"...), normalisée pour le fuzzy
    #    matching contre le référentiel.
    # ------------------------------------------------------
    wave["CLE"] = wave[col_operateur].apply(normaliser)
    wave["FICHIER_SOURCE"] = fichier.filename

    # ------------------------------------------------------
    # 2) RETRAIT / DEPOT à partir de "Quoi" et "Montant (CFA)".
    #    montant.abs() neutralise les montants Wave négatifs (les
    #    retraits notamment sont souvent négatifs dans le relevé) :
    #    RETRAIT et DEPOT sont donc toujours des valeurs positives,
    #    directement comparables au DEBIT/CREDIT (positifs) de Flex.
    # ------------------------------------------------------
    montant = pd.to_numeric(wave[col_montant], errors="coerce").fillna(0)
    quoi_norm = wave[col_quoi].apply(normaliser)

    wave["RETRAIT"] = np.where(quoi_norm == LIBELLE_RETRAIT, montant.abs(), 0)
    wave["DEPOT"] = np.where(quoi_norm == LIBELLE_DEPOT, montant.abs(), 0)

    # [CORRECTIF] Montant SIGNÉ conservé tel quel (pas de abs()) pour
    # TOUTES les lignes, quel que soit "Quoi" (Retrait, Dépôt, Achat
    # des UV, Remboursement des UV...). C'est LA colonne utilisée pour
    # calculer le solde réel par agence (SOLDE_PARTENAIRE, voir
    # agreger_par_agence ci-dessous) : RETRAIT/DEPOT ne couvrent que
    # deux des catégories de "Quoi" et sous-estiment le solde si on
    # les utilisait seules (une agence peut être créditée/débitée par
    # d'autres types d'opérations qui doivent aussi entrer dans le
    # solde).
    wave["MONTANT"] = montant

    return wave


# ==========================================================
# LECTURE DU MAPPING
# ------------------------------------------------------------
# Nouveau schéma du référentiel mapping_agences_cofina_wave.xlsx :
#   Opérateur | CODE_AGENCE | NOM_AGENCE_MAPPEE
# (CODE_SOURCE / LIBELLE_SOURCE n'existent plus)
# ==========================================================

def lire_mapping() -> pd.DataFrame:

    try:
        mapping = pd.read_excel(MAPPING_AGENCES_WAVE_PATH)
    except FileNotFoundError:
        raise HTTPException(
            status_code=500,
            detail=f"Fichier de mapping introuvable : {MAPPING_AGENCES_WAVE_PATH}"
        )

    mapping.columns = mapping.columns.str.strip()

    col_operateur = _trouver_colonne(mapping.columns, ["Operateur"])
    if col_operateur is None:
        raise HTTPException(
            status_code=500,
            detail="Colonne 'Opérateur' manquante dans le fichier de mapping."
        )

    for col in ("CODE_AGENCE", "NOM_AGENCE_MAPPEE"):
        if col not in mapping.columns:
            raise HTTPException(
                status_code=500,
                detail=f"Colonne '{col}' manquante dans le fichier de mapping."
            )

    # Nom interne stable, quel que soit l'accent exact utilisé dans
    # l'entête source ("Opérateur" / "Operateur"...).
    mapping.rename(columns={col_operateur: "OPERATEUR_MAPPING"}, inplace=True)

    # CODE_AGENCE doit rester un entier "propre" (512, 509, 523…),
    # comparable au code agence renvoyé par Flex. On force le type
    # "Int64" nullable pour éviter tout glissement en flottant
    # (512.0) ou en texte.
    mapping["CODE_AGENCE"] = pd.to_numeric(
        mapping["CODE_AGENCE"], errors="coerce"
    ).astype("Int64")

    mapping["CLE"] = mapping["OPERATEUR_MAPPING"].apply(normaliser)

    # On ignore les lignes de mapping sans "Opérateur" exploitable
    # (NaN), et on déduplique sur la clé normalisée (on garde la
    # première occurrence) pour ne pas fausser le fuzzy matching.
    mapping = mapping[mapping["CLE"] != ""].drop_duplicates(subset="CLE", keep="first")

    return mapping.reset_index(drop=True)


# ==========================================================
# MATCHING FUZZY (détail, transaction par transaction)
# ==========================================================

def appliquer_matching(wave: pd.DataFrame, mapping: pd.DataFrame) -> pd.DataFrame:
    """Retourne le DÉTAIL (une ligne par transaction Wave matchée) —
    utilisé pour l'audit (feuilles Excel). Pour l'agrégat par agence
    (une ligne par CODE_AGENCE, sans doublons), voir
    agreger_par_agence() ci-dessous."""

    liste_mapping = mapping["CLE"].tolist()
    mapping_par_cle = mapping.set_index("CLE")

    codes, agences, scores, libelles = [], [], [], []

    # Cache : un même opérateur revient sur de nombreuses lignes du
    # relevé, inutile de relancer rapidfuzz à chaque fois pour la
    # même clé normalisée.
    cache = {}

    for operateur in wave["CLE"]:

        if operateur in cache:
            code, nom_agence, score, libelle = cache[operateur]
        else:
            meilleur = process.extractOne(
                operateur,
                liste_mapping,
                scorer=fuzz.token_sort_ratio
            )

            if meilleur is None or operateur == "":
                code, nom_agence, score, libelle = None, None, 0, None
            else:
                libelle_trouve, score = meilleur[0], meilleur[1]

                if score >= SEUIL_SCORE:
                    ligne = mapping_par_cle.loc[libelle_trouve]
                    code = ligne["CODE_AGENCE"]
                    nom_agence = ligne["NOM_AGENCE_MAPPEE"]
                    libelle = ligne["OPERATEUR_MAPPING"]
                else:
                    code, nom_agence, libelle = None, None, None

            cache[operateur] = (code, nom_agence, score, libelle)

        codes.append(code)
        agences.append(nom_agence)
        scores.append(score)
        libelles.append(libelle)

    wave["CODE_AGENCE"] = pd.array(codes, dtype="Int64")
    wave["NOM_AGENCE"] = agences
    wave["LIBELLE_MAPPING"] = libelles
    wave["SCORE"] = scores

    # ------------------------------------------------------
    # On exclut les lignes non matchées (CODE_AGENCE manquant :
    # "Opérateur" vide, aucun candidat, ou score sous le seuil). Ces
    # lignes n'ont pas de correspondance fiable côté Flex et
    # faussaient les totaux de réconciliation par agence.
    # ------------------------------------------------------
    wave = wave[wave["CODE_AGENCE"].notna()].copy()

    ordre = ["CODE_AGENCE", "NOM_AGENCE", "LIBELLE_MAPPING", "SCORE", "RETRAIT", "DEPOT"]
    ordre += [c for c in wave.columns if c not in ordre]

    return wave[ordre]


# ==========================================================
# [CORRECTIF] AGRÉGAT PAR AGENCE — UNE LIGNE PAR CODE_AGENCE
# ------------------------------------------------------------
# Réduit le détail (une ligne par transaction, déjà filtrée sur les
# lignes matchées) à UNE SEULE LIGNE PAR AGENCE, avec les totaux
# RETRAIT / DEPOT et le solde net.
#
# Remplace l'ancien `.groupby(...).transform("sum")`, qui NE
# RÉDUISAIT PAS le nombre de lignes (la même somme était recopiée
# sur chaque transaction -> autant de "doublons apparents" que de
# transactions pour une même agence). Ici, `.agg()` retourne bien
# une seule ligne par CODE_AGENCE.
# ==========================================================

def agreger_par_agence(df: pd.DataFrame) -> pd.DataFrame:

    colonnes_vides = [
        "CODE_AGENCE", "NOM_AGENCE", "LIBELLE_MAPPING",
        "NB_TRANSACTIONS", "RETRAIT", "DEPOT", "SOLDE_PARTENAIRE",
        "FICHIERS_SOURCE",
    ]

    if df.empty:
        return pd.DataFrame(columns=colonnes_vides)

    agg = (
        df.groupby("CODE_AGENCE", dropna=False)
          .agg(
              NOM_AGENCE=("NOM_AGENCE", "first"),
              LIBELLE_MAPPING=("LIBELLE_MAPPING", "first"),
              NB_TRANSACTIONS=("MONTANT", "count"),
              RETRAIT=("RETRAIT", "sum"),
              DEPOT=("DEPOT", "sum"),
              # [CORRECTIF] Solde = somme du MONTANT SIGNÉ, TOUTES
              # catégories de "Quoi" confondues (Retrait, Dépôt, Achat
              # des UV, Remboursement des UV...) — c'est ce total qui
              # correspond au vrai solde de l'agence (ex: 4 083 235
              # pour l'agence 501), pas seulement "RETRAIT - DEPOT"
              # qui ignore les autres types d'opérations.
              SOLDE_PARTENAIRE=("MONTANT", "sum"),
              FICHIERS_SOURCE=("FICHIER_SOURCE", lambda s: ", ".join(sorted(set(s.dropna())))),
          )
          .reset_index()
    )

    # RETRAIT / DEPOT restent dans le résultat à titre INFORMATIF
    # (répartition partielle, ne couvrant que 2 des catégories de
    # "Quoi") — ce n'est plus la formule utilisée pour SOLDE_PARTENAIRE,
    # qui est calculé ci-dessus directement à partir du MONTANT signé.

    return agg[colonnes_vides].sort_values("CODE_AGENCE").reset_index(drop=True)


# ==========================================================
# SAUVEGARDE SQLITE
# ==========================================================

def sauvegarder_sqlite(df: pd.DataFrame, table: str, source: str):
    ecrire_table(df, table, engine, log_prefix=source)


# ==========================================================
# ENDPOINTS
# ==========================================================

def nom_feuille(nom_fichier: str, index: int) -> str:
    """Nom de feuille Excel valide (<=31 caractères, sans caractères
    interdits, TOUJOURS unique grâce au suffixe "_index" — deux
    fichiers dont le nom tronqué serait identique ne produisent donc
    jamais un conflit de feuille)."""
    base = nom_fichier.rsplit(".", 1)[0]
    for caractere in ["\\", "/", "?", "*", "[", "]", ":"]:
        base = base.replace(caractere, "_")
    base = base.strip() or "Fichier"

    suffixe = f"_{index}"
    return base[: 31 - len(suffixe)] + suffixe


@app.post("/map-agences")
async def map_agences(
    files: List[UploadFile] = File(...),
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):

    if not files:
        raise HTTPException(status_code=400, detail="Aucun fichier envoyé.")

    try:
        mapping = lire_mapping()

        # ------------------------------------------------------
        # 1. TRAITEMENT INDIVIDUEL DE CHAQUE FICHIER (détail)
        #    (exclusion de la ligne technique A12956_SN3, calcul
        #    RETRAIT/DEPOT en valeur absolue, fuzzy matching, puis
        #    exclusion des lignes non matchées — tout est fait dans
        #    lire_fichier_wave() / appliquer_matching() ci-dessus)
        # ------------------------------------------------------
        dfs_traites = []
        for fichier in files:
            wave = lire_fichier_wave(fichier)
            wave_mappe = appliquer_matching(wave, mapping)
            dfs_traites.append((fichier.filename, wave_mappe))

        # ------------------------------------------------------
        # 2. DÉTAIL COMBINÉ (tous fichiers confondus) — gardé pour
        #    l'audit dans une feuille Excel dédiée, mais N'EST PLUS
        #    ce qui est sauvegardé en base ni utilisé pour la
        #    réconciliation (voir point 3).
        # ------------------------------------------------------
        df_detail_combine = pd.concat(
            [df for _, df in dfs_traites],
            ignore_index=True
        )

        # ------------------------------------------------------
        # 3. [CORRECTIF] AGRÉGAT PROPRE PAR AGENCE — une seule ligne
        #    par CODE_AGENCE (ex: 501, 4083235...), plus de doublons.
        #    C'est CE DataFrame qui est sauvegardé en base et utilisé
        #    par reconcilier_par_agence().
        # ------------------------------------------------------
        df_agrege = agreger_par_agence(df_detail_combine)

        sauvegarder_sqlite(df_agrege, TABLE_SORTIE, "wave_app_agence")

        # ------------------------------------------------------
        # 4. EXPORT : Excel multi-feuilles (manuel) ou JSON léger
        #    (/charger Module_FED — le corps Excel était ignoré).
        # ------------------------------------------------------
        sheets = {"Compilation": df_agrege, "Detail_Compilation": df_detail_combine}
        for i, (nom_fichier, df) in enumerate(dfs_traites, start=1):
            sheets[nom_feuille(nom_fichier, i)] = df

        if wants_excel(format):
            return respond_sheets(sheets, filename="WAVE_MAPPE.xlsx", format="excel")

        return respond_sheets(
            {"Compilation": df_agrege},
            filename="WAVE_MAPPE.xlsx",
            format="json",
            json_payload={
                "status": "ok",
                "format": "json",
                "filename": "WAVE_MAPPE.xlsx",
                "nb_agences": int(len(df_agrege)),
                "nb_transactions": int(len(df_detail_combine)),
                "nb_fichiers": len(dfs_traites),
            },
        )

    except HTTPException:
        raise
    except Exception as e:
        return {"status": "error", "message": str(e), "trace": traceback.format_exc()}


@app.get("/db/wave-agence")
def get_wave_agence(limit: int = Query(None), offset: int = Query(0)):
    return lire_table_json(engine, TABLE_SORTIE, limit=limit, offset=offset)


@app.get("/health")
def health():
    return {
        "status": "ok",
        "partenaire": PARTENAIRE,
        "db_path": DB_PATH,
        "mapping": MAPPING_AGENCES_WAVE_PATH,
    }