# ============================================================
# APP ORANGE MONEY — Upload MULTI-FICHIERS + Mapping MSISDN Agence
# ------------------------------------------------------------
# uvicorn app_orange:app --port 8010
#
# CORRECTIF IMPORTANT (vs version précédente) :
# /process-excel traitait UNIQUEMENT files[0] et ignorait
# silencieusement tous les autres fichiers uploadés — c'est ce qui
# provoquait une table COMPILATION_ORANGE vide ou incomplète dès
# que le premier fichier de la liste ne contenait pas de lignes
# exploitables (ex: fichier "TOTAL" seul, ou format différent).
#
# Comportement désormais :
#   1. Chaque fichier uploadé est traité INDIVIDUELLEMENT par
#      traiter_orange() (recherche de sa propre ligne d'entête,
#      son propre mapping MSISDN -> agence, etc.)
#   2. Les DataFrames obtenus sont CONCATÉNÉS (pd.concat) en un
#      seul DataFrame final.
#   3. Ce DataFrame concaténé est sauvegardé UNE SEULE FOIS dans
#      COMPILATION_ORANGE.
#   4. Si un fichier échoue (mauvais format, colonnes inattendues,
#      ...), il est ignoré et signalé (dans les logs serveur ET
#      dans la réponse), mais les autres fichiers valides sont
#      quand même traités et sauvegardés — un seul fichier
#      problématique ne bloque plus tout le chargement.
#   5. Si AUCUN fichier n'a pu être traité, une erreur 500 claire
#      est renvoyée (au lieu de sauvegarder un DataFrame vide).
#
# NOUVEAU (cette version) :
#   6. Isolation des lignes dont la colonne "nb_cashin" est NULL
#      (NaN / non convertible en nombre) :
#        - ces lignes sont SORTIES du DataFrame principal (donc
#          exclues de COMPILATION_ORANGE),
#        - elles sont regroupées dans un DataFrame séparé
#          (df_nulls_cashin) par fichier, puis concaténées à
#          l'échelle globale,
#        - elles apparaissent dans un onglet dédié
#          "Nb_Cashin_Nuls" du fichier Excel retourné,
#          ainsi qu'un en-tête custom "X-Lignes-Nb-Cashin-Nulles",
#          pour qu'aucune ligne ne disparaisse silencieusement.
#
# CORRECTIF (cette version) :
#   7. PARTENAIRE était réglé sur "ORANGE", une clé qui n'existe
#      PAS dans config.PARTENAIRES (la clé réelle est
#      "ORANGE_AGENCE", voir config.py). Cela provoquait un
#      KeyError: 'ORANGE' au démarrage du module (donc du service
#      uvicorn app_orange:app, port 8010), avant même que
#      l'application FastAPI ne soit créée — le service restait
#      injoignable, d'où les 500/502 "Erreur Upload Excel" renvoyés
#      par reconc.py lors du /charger?partenaire=ORANGE_AGENCE.
#      -> PARTENAIRE est maintenant aligné sur la clé réelle de
#      config.py : "ORANGE_AGENCE".
# ============================================================

import io
import os
import traceback
from typing import List

import pandas as pd

from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import StreamingResponse


from config import DB_PATH, PARTENAIRES, MAPPING_ORANGE_MSISDN_PATH, make_sqlite_engine
from common.excel_common import sauvegarder_sqlite
from common.sqlite_io import lire_table_json
from common.http_export import respond_sheets, wants_excel


# ============================================================
# PARAMETRES
# ============================================================

PARTENAIRE = "ORANGE_AGENCE"  # <-- CORRIGÉ (était "ORANGE", clé inexistante dans config.PARTENAIRES)

TABLES = PARTENAIRES[PARTENAIRE]["tables"]

FICHIER_MAPPING = MAPPING_ORANGE_MSISDN_PATH

NB_COLONNES = 33

print(f"[app_orange.py] Base SQLite : {DB_PATH}")

engine = make_sqlite_engine()

app = FastAPI(
    title="Orange Money Mapping API"
)


# ============================================================
# COLONNES ORANGE MONEY
# ============================================================

COLONNES_ORANGE = [

    "numero",
    "tete_reseau_n1",
    "membre_n2",
    "membre_n3",
    "membre_n4",
    "statut",

    "msisdn",
    "msisdn_tdr",

    "souscriptions",
    "clients_actifs",

    "nb_tx_souscriptions",
    "montant_souscriptions",

    "nb_cashin",
    "montant_cashin",

    "nb_cashout",
    "montant_cashout",

    "nb_tno",
    "montant_tno",

    "nb_c2c_emis",
    "montant_c2c_emis",

    "nb_c2c_recus",
    "montant_c2c_recus",

    "nb_approvisionnements",
    "montant_approvisionnements",

    "nb_remboursements",
    "montant_remboursements",

    "nb_paiements_marchands",
    "montant_paiements_marchands",

    "commission_cashout",
    "commission_cashin",
    "commission_tno",
    "commission_paiements_marchands",
    "commission_totale"
]


# ============================================================
# TRAITEMENT D'UN SEUL FICHIER ORANGE MONEY
# ------------------------------------------------------------
# Reçoit un objet fichier (UploadFile.file) et renvoie DEUX
# DataFrames pour CE fichier uniquement :
#   - df            : lignes valides (nb_cashin non null)
#   - df_nulls      : lignes isolées car nb_cashin est null
#
# Appelée une fois par fichier uploadé par process_excel()
# ci-dessous, qui se charge de la boucle et de la concaténation
# (séparément pour les deux flux).
# ============================================================

def traiter_orange(upload_file):

    # --------------------------------------------------------
    # Lecture brute
    # --------------------------------------------------------

    upload_file.seek(0)

    df_brut = pd.read_excel(
        upload_file,
        header=None
    )

    # --------------------------------------------------------
    # Recherche ligne entête
    # --------------------------------------------------------

    ligne_entete = None

    for i in range(len(df_brut)):

        valeur = str(df_brut.iloc[i, 0]).strip().lower()

        if valeur in ["n°", "nº", "no"]:

            ligne_entete = i
            break

    if ligne_entete is None:

        raise Exception(
            "Impossible de trouver la ligne d'entête Orange Money."
        )

    print(f"Ligne entête détectée : {ligne_entete}")

    # --------------------------------------------------------
    # Lecture du tableau
    # --------------------------------------------------------

    upload_file.seek(0)

    df = pd.read_excel(
        upload_file,
        header=ligne_entete
    )

    if len(df.columns) != NB_COLONNES:

        raise Exception(
            f"""
Structure inattendue

Colonnes trouvées : {len(df.columns)}
Colonnes attendues : {NB_COLONNES}

{list(df.columns)}
"""
        )

    df.columns = COLONNES_ORANGE

    # --------------------------------------------------------
    # Suppression TOTAL
    # --------------------------------------------------------

    df = df[
        df["numero"]
        .astype(str)
        .str.strip()
        .str.upper()
        != "TOTAL"
    ]

    df.reset_index(
        drop=True,
        inplace=True
    )

    # --------------------------------------------------------
    # Nettoyage MSISDN
    # --------------------------------------------------------

    df["msisdn"] = (
        df["msisdn"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    # --------------------------------------------------------
    # Conversion numérique + ISOLATION DES NULLS SUR nb_cashin
    # ----------------------------------------------------------
    # On convertit d'abord "nb_cashin" en numérique (errors="coerce"
    # transforme toute valeur non convertible en NaN, ce qui est
    # considéré comme "null" pour ce filtre).
    # On isole ensuite TOUTES les lignes où "nb_cashin" est null
    # dans un DataFrame séparé (df_nulls_cashin), et on les retire
    # du DataFrame principal : elles ne seront donc PAS incluses
    # dans COMPILATION_ORANGE, mais restent traçables (onglet dédié
    # dans l'Excel retourné).
    # --------------------------------------------------------

    df["nb_cashin"] = pd.to_numeric(
        df["nb_cashin"],
        errors="coerce"
    )

    masque_nb_cashin_null = df["nb_cashin"].isna()

    df_nulls_cashin = df[masque_nb_cashin_null].copy()

    print(
        f"Lignes avec nb_cashin null isolées : "
        f"{len(df_nulls_cashin)} / {len(df)}"
    )

    df = df[~masque_nb_cashin_null].copy()

    df.reset_index(
        drop=True,
        inplace=True
    )

    df_nulls_cashin.reset_index(
        drop=True,
        inplace=True
    )

    # --------------------------------------------------------
    # Création colonne sens
    # --------------------------------------------------------

    df["montant_cashin"] = pd.to_numeric(
        df["montant_cashin"],
        errors="coerce"
    ).fillna(0)

    df["montant_cashout"] = pd.to_numeric(
        df["montant_cashout"],
        errors="coerce"
    ).fillna(0)

    df["sens"] = "AUTRE"

    df.loc[
        df["montant_cashin"] != 0,
        "sens"
    ] = "IN"

    df.loc[
        df["montant_cashout"] != 0,
        "sens"
    ] = "OUT"

    print("\nRépartition colonne sens :")
    print(df["sens"].value_counts())

    # --------------------------------------------------------
    # Chargement Mapping
    # --------------------------------------------------------

    if not os.path.exists(FICHIER_MAPPING):

        raise Exception(
            f"Le fichier {FICHIER_MAPPING} est introuvable "
            f"(chemin recherché : {os.path.abspath(FICHIER_MAPPING)}). "
            f"Vérifie qu'il est bien présent dans le dossier où tourne "
            f"'uvicorn app_orange:app'."
        )

    df_map = pd.read_excel(
        FICHIER_MAPPING
    )

    colonnes_map = [
        "MSISDN",
        "AGENCE",
        "CODE_AGENCE"
    ]

    for c in colonnes_map:

        if c not in df_map.columns:

            raise Exception(
                f"Colonne absente dans Mapp_tab : {c}"
            )

    df_map["MSISDN"] = (
        df_map["MSISDN"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    # --------------------------------------------------------
    # Mapping (LEFT JOIN)
    # --------------------------------------------------------

    df = df.merge(

        df_map[
            [
                "MSISDN",
                "AGENCE",
                "CODE_AGENCE"
            ]
        ],

        how="left",

        left_on="msisdn",

        right_on="MSISDN"

    )

    df.drop(
        columns=["MSISDN"],
        inplace=True
    )

    # --------------------------------------------------------
    # Contrôles
    # --------------------------------------------------------

    print("--------------------------------------")
    print("Nombre lignes Orange :", len(df))
    print("Sans agence :", df["AGENCE"].isna().sum())
    print("Lignes isolées (nb_cashin null) :", len(df_nulls_cashin))
    print("--------------------------------------")

    return df, df_nulls_cashin


# ============================================================
# UPLOAD — MULTI-FICHIERS
# ------------------------------------------------------------
# Traite CHAQUE fichier uploadé séparément avec traiter_orange(),
# puis concatène tous les résultats obtenus en un seul DataFrame
# avant la sauvegarde SQLite (une seule écriture dans
# COMPILATION_ORANGE, table déjà "replace" par sauvegarder_sqlite
# si c'est son comportement — sinon voir la note ci-dessous).
#
# Un fichier en échec (mauvais format, colonnes manquantes...) est
# IGNORÉ individuellement (pas de blocage du chargement global) et
# listé dans "fichiers_erreurs" de la réponse + dans les logs
# serveur avec sa trace complète. Seule l'absence TOTALE de fichier
# exploitable est bloquante (500).
#
# Les lignes dont "nb_cashin" est null sont isolées par
# traiter_orange() et concaténées séparément ici
# (df_nulls_cashin_all) : elles ne vont PAS dans COMPILATION_ORANGE
# mais sont restituées dans l'onglet "Nb_Cashin_Nuls" du fichier
# Excel retourné.
#
# Les erreurs sont levées via HTTPException (status_code 500) plutôt
# qu'un simple dict JSON avec statut 200 implicite : un 200
# "silencieux" en cas d'échec empêchait reconc.py de détecter le
# problème (son contrôle `if excel_response.status_code != 200` ne
# se déclenchait jamais), et sauvegarder_sqlite() n'était alors
# jamais appelé : les tables restaient vides sans qu'aucune erreur
# ne remonte jusqu'à Streamlit.
# ============================================================

@app.post("/process-excel")
async def process_excel(
    files: List[UploadFile] = File(...),
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):

    if len(files) == 0:
        raise HTTPException(
            status_code=400,
            detail="Aucun fichier reçu."
        )

    dataframes_ok = []
    dataframes_nulls_cashin = []
    fichiers_ok = []
    fichiers_erreurs = []

    for f in files:

        try:
            df_fichier, df_nulls_fichier = traiter_orange(f.file)
        except Exception as e:
            print(f"[app_orange.py] Erreur traitement '{f.filename}' :")
            traceback.print_exc()
            fichiers_erreurs.append({"fichier": f.filename, "erreur": str(e)})
            continue

        # Traçabilité : sait de quel fichier source vient chaque ligne,
        # utile pour déboguer un futur écart après concaténation.
        df_fichier["_FICHIER_SOURCE"] = f.filename
        df_nulls_fichier["_FICHIER_SOURCE"] = f.filename

        dataframes_ok.append(df_fichier)
        dataframes_nulls_cashin.append(df_nulls_fichier)
        fichiers_ok.append(f.filename)

    if not dataframes_ok:
        raise HTTPException(
            status_code=500,
            detail=(
                "Aucun fichier n'a pu être traité. "
                f"Détails : {fichiers_erreurs}"
            )
        )

    # --------------------------------------------------------
    # Concaténation de TOUS les fichiers traités avec succès en
    # un seul DataFrame final, avant la sauvegarde SQLite.
    # --------------------------------------------------------
    df_final = pd.concat(dataframes_ok, ignore_index=True)

    # --------------------------------------------------------
    # Concaténation des lignes isolées (nb_cashin null) à
    # l'échelle globale (peut être vide si aucun fichier n'en
    # contenait).
    # --------------------------------------------------------
    if dataframes_nulls_cashin:
        df_nulls_cashin_all = pd.concat(
            dataframes_nulls_cashin,
            ignore_index=True
        )
    else:
        df_nulls_cashin_all = pd.DataFrame()

    print(
        f"[app_orange.py] {len(fichiers_ok)} fichier(s) traité(s) et "
        f"concaténé(s) ({len(df_final)} lignes au total) : {fichiers_ok}"
    )
    if fichiers_erreurs:
        print(f"[app_orange.py] {len(fichiers_erreurs)} fichier(s) ignoré(s) : {fichiers_erreurs}")
    print(
        f"[app_orange.py] {len(df_nulls_cashin_all)} ligne(s) isolée(s) "
        f"globalement (nb_cashin null), exclues de COMPILATION_ORANGE."
    )

    try:
        sauvegarder_sqlite(
            df_final,
            TABLES["excel"],
            engine,
            "app_orange"
        )
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Erreur sauvegarde SQLite (table '{TABLES['excel']}') : {e}"
        )

    try:
        export_headers = {
            "X-Fichiers-Traites": str(len(fichiers_ok)),
            "X-Fichiers-Ignores": str(len(fichiers_erreurs)),
            "X-Lignes-Nb-Cashin-Nulles": str(len(df_nulls_cashin_all)),
        }

        if not wants_excel(format):
            return respond_sheets(
                {"Orange_Mappe": df_final},
                filename="Orange_Mappe.xlsx",
                format="json",
                headers=export_headers,
                json_payload={
                    "status": "ok",
                    "format": "json",
                    "filename": "Orange_Mappe.xlsx",
                    "nb_lignes": int(len(df_final)),
                    "nb_fichiers_traites": len(fichiers_ok),
                    "nb_fichiers_ignores": len(fichiers_erreurs),
                    "nb_lignes_nb_cashin_nulles": int(len(df_nulls_cashin_all)),
                },
            )

        sheets = {
            "Orange_Mappe": df_final,
            "Nb_Cashin_Nuls": df_nulls_cashin_all,
        }
        if fichiers_erreurs:
            sheets["Fichiers_Ignores"] = pd.DataFrame(fichiers_erreurs)

        return respond_sheets(
            sheets,
            filename="Orange_Mappe.xlsx",
            format="excel",
            headers=export_headers,
        )

    except Exception as e:
        # La sauvegarde SQLite a déjà réussi à ce stade : seule la
        # génération du fichier Excel de retour a échoué. On le
        # signale quand même comme une erreur pour que ce soit
        # visible côté appelant.
        raise HTTPException(
            status_code=500,
            detail=f"Sauvegarde SQLite OK, mais erreur génération Excel : {e}"
        )


# ============================================================
# CONSULTATION SQLITE
# ============================================================

@app.get("/db/compilation")
def get_compilation(limit: int = Query(None), offset: int = Query(0)):
    return lire_table_json(engine, TABLES["excel"], limit=limit, offset=offset)


# ============================================================
# HEALTH
# ============================================================

@app.get("/health")
def health():

    return {

        "status": "ok",

        "partenaire": PARTENAIRE,

        "db_path": DB_PATH

    }