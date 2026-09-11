# ============================================================
# APP NIIRPAY — SERVICE EXCEL / CSV
# ============================================================
#
# Rôle :
#   - recevoir le fichier CSV NiirPay
#   - normaliser les colonnes
#   - produire le schéma standard du moteur de réconciliation
#   - séparer W2B / B2W
#   - sauvegarder les données dans SQLite
#   - retourner JSON ou Excel
#
# IMPORTANT :
#   Le gateway envoie TOUJOURS une List[UploadFile].
#
#   Même si NiirPay fournit un seul fichier, on conserve :
#
#       files: List[UploadFile]
#
# ============================================================

import traceback
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

from config import (
    DB_PATH,
    PARTENAIRES,
    make_sqlite_engine,
)

from common.excel_common import sauvegarder_sqlite
from common.sqlite_io import lire_table_json
from common.http_export import respond_sheets, wants_excel


# ============================================================
# CONFIGURATION NIIRPAY
# ============================================================

PARTENAIRE = "NIIRPAY"


if PARTENAIRE not in PARTENAIRES:
    raise RuntimeError(
        f"Le partenaire '{PARTENAIRE}' n'existe pas dans "
        f"config.PARTENAIRES."
    )


TABLES = PARTENAIRES[PARTENAIRE]["tables"]


print(
    f"[app_niirpay.py] Base SQLite utilisée : {DB_PATH}"
)


engine = make_sqlite_engine()


# ============================================================
# APPLICATION FASTAPI
# ============================================================

app = FastAPI(
    title="NiirPay Excel API",
    description=(
        "Service de chargement et normalisation des fichiers "
        "NiirPay pour la réconciliation."
    ),
)


# ============================================================
# OUTILS
# ============================================================

def normaliser_nom_colonne(colonne):
    """
    Normalise le nom d'une colonne :

    - suppression des espaces inutiles
    - passage en majuscules
    - suppression des accents
    """

    if colonne is None:
        return ""

    texte = str(colonne).strip().upper()

    texte = unicodedata.normalize(
        "NFKD",
        texte,
    )

    texte = "".join(
        caractere
        for caractere in texte
        if not unicodedata.combining(caractere)
    )

    texte = " ".join(
        texte.split()
    )

    return texte


def trouver_colonne(
    df: pd.DataFrame,
    noms_possibles,
):
    """
    Recherche une colonne en utilisant plusieurs noms possibles.
    """

    colonnes = {}

    for colonne in df.columns:

        nom_normalise = normaliser_nom_colonne(
            colonne
        )

        colonnes[nom_normalise] = colonne

    for nom in noms_possibles:

        nom_normalise = normaliser_nom_colonne(
            nom
        )

        if nom_normalise in colonnes:

            return colonnes[nom_normalise]

    return None


def supprimer_timezone_pour_excel(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Supprime les timezone éventuelles afin de permettre
    l'export Excel.
    """

    df = df.copy()

    for colonne in df.columns:

        try:

            if pd.api.types.is_datetime64_any_dtype(
                df[colonne]
            ):

                df[colonne] = (
                    pd.to_datetime(
                        df[colonne],
                        errors="coerce",
                    )
                    .dt.tz_localize(None)
                )

        except Exception:
            pass

    return df


# ============================================================
# NETTOYAGE DES COLONNES POUR SQLITE
# ============================================================

def supprimer_doublons_colonnes_sqlite(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    SQLite ne distingue pas correctement certaines colonnes
    qui ne diffèrent que par la casse.

    Exemple :

        commissions_operator_amount
        COMMISSIONS_OPERATOR_AMOUNT

    SQLite considère ces deux noms comme identiques.

    On conserve la première occurrence.
    """

    df = df.copy()

    colonnes_vues = set()

    colonnes_a_conserver = []

    for colonne in df.columns:

        nom_sql = str(
            colonne
        ).strip().upper()

        if nom_sql not in colonnes_vues:

            colonnes_vues.add(
                nom_sql
            )

            colonnes_a_conserver.append(
                colonne
            )

        else:

            print(
                "[app_niirpay] "
                "Colonne dupliquée supprimée : "
                f"{colonne}"
            )

    return df.loc[
        :,
        colonnes_a_conserver,
    ].copy()


# ============================================================
# LECTURE DU FICHIER NIIRPAY
# ============================================================

def lire_fichier_niirpay(
    upload_file,
) -> pd.DataFrame:
    """
    Lit le fichier CSV NiirPay.

    On essaie d'abord UTF-8-SIG puis Latin-1.
    Le séparateur est détecté automatiquement.
    """

    upload_file.seek(0)

    try:

        df = pd.read_csv(
            upload_file,
            sep=None,
            engine="python",
            encoding="utf-8-sig",
        )

    except UnicodeDecodeError:

        upload_file.seek(0)

        df = pd.read_csv(
            upload_file,
            sep=None,
            engine="python",
            encoding="latin-1",
        )

    except Exception as e:

        raise HTTPException(
            status_code=400,
            detail=(
                "Impossible de lire le fichier NiirPay : "
                f"{e}"
            ),
        )

    # --------------------------------------------------------
    # Nettoyage des noms de colonnes
    # --------------------------------------------------------

    df.columns = [
        str(colonne).strip()
        for colonne in df.columns
    ]

    print(
        "=================================================="
    )

    print(
        "[app_niirpay.py] Colonnes détectées :"
    )

    print(
        list(df.columns)
    )

    print(
        "=================================================="
    )

    if df.empty:

        raise HTTPException(
            status_code=400,
            detail=(
                "Le fichier NiirPay "
                "ne contient aucune ligne."
            ),
        )

    return df


# ============================================================
# TRAITEMENT NIIRPAY
# ============================================================

def traiter_niirpay(
    upload_file,
) -> pd.DataFrame:
    """
    Transforme le fichier NiirPay vers le format
    utilisé par l'application de réconciliation.
    """

    df = lire_fichier_niirpay(
        upload_file
    )

    # ========================================================
    # RECHERCHE DES COLONNES
    # ========================================================

    colonne_date = trouver_colonne(
        df,
        [
            "Transaction Date",
            "TRANSACTIONDATE",
            "Date Transaction",
            "DATE TRANSACTION",
        ],
    )

    colonne_niirpay_id = trouver_colonne(
        df,
        [
            "niirPay Transaction ID",
            "NIIRPAY TRANSACTION ID",
            "GUTRANSACTIONID",
        ],
    )

    colonne_operator_id = trouver_colonne(
        df,
        [
            "Operator Transaction ID",
            "OPERATOR TRANSACTION ID",
            "CODE TRANSACTION OPERATEUR",
        ],
    )

    colonne_reference = trouver_colonne(
        df,
        [
            "reference",
            "REFERENCE",
            "REFERENCE NIIRPAY",
        ],
    )

    colonne_amount = trouver_colonne(
        df,
        [
            "amount",
            "AMOUNT",
            "MONTANT",
        ],
    )

    colonne_fees = trouver_colonne(
        df,
        [
            "fees",
            "FRAIS",
        ],
    )

    colonne_commission = trouver_colonne(
        df,
        [
            "commissions_operator_amount",
            "COMMISSIONS OPERATOR AMOUNT",
            "COMMISSION",
        ],
    )

    colonne_currency = trouver_colonne(
        df,
        [
            "currency",
            "CURRENCY",
            "DEVISE",
        ],
    )

    # ========================================================
    # VERIFICATION DES COLONNES OBLIGATOIRES
    # ========================================================

    colonnes_obligatoires = {

        "Transaction Date":
            colonne_date,

        "niirPay Transaction ID":
            colonne_niirpay_id,

        "Operator Transaction ID":
            colonne_operator_id,

        "amount":
            colonne_amount,

        "fees":
            colonne_fees,

        "currency":
            colonne_currency,
    }

    manquantes = [

        nom

        for nom, colonne
        in colonnes_obligatoires.items()

        if colonne is None
    ]

    if manquantes:

        raise HTTPException(
            status_code=400,
            detail=(
                "Colonnes obligatoires NiirPay absentes : "
                f"{manquantes}. "
                f"Colonnes disponibles : "
                f"{list(df.columns)}"
            ),
        )

    # ========================================================
    # DATE TRANSACTION
    # ========================================================

    df["DATE TRANSACTION"] = pd.to_datetime(
        df[colonne_date],
        errors="coerce",
    )

    # ========================================================
    # IDENTIFIANT NIIRPAY
    # ========================================================

    df["NIIRPAY_TRANSACTION_ID"] = (
        df[colonne_niirpay_id]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # IDENTIFIANT OPERATEUR
    # ========================================================

    df["CODE TRANSACTION OPERATEUR"] = (
        df[colonne_operator_id]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # REFERENCE
    # ========================================================

    if colonne_reference is not None:

        df["REFERENCE_NIIRPAY"] = (
            df[colonne_reference]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    else:

        df["REFERENCE_NIIRPAY"] = ""

    # ========================================================
    # MONTANT
    # ========================================================

    df["MONTANT"] = pd.to_numeric(
        df[colonne_amount],
        errors="coerce",
    )

    # ========================================================
    # FRAIS
    # ========================================================

    df["FRAIS"] = (
        pd.to_numeric(
            df[colonne_fees],
            errors="coerce",
        )
        .fillna(0)
    )

    # ========================================================
    # COMMISSION
    # ========================================================

    if colonne_commission is not None:

        df["COMMISSIONS_OPERATOR_AMOUNT"] = (
            pd.to_numeric(
                df[colonne_commission],
                errors="coerce",
            )
            .fillna(0)
        )

    else:

        df["COMMISSIONS_OPERATOR_AMOUNT"] = 0

    # ========================================================
    # DEVISE
    # ========================================================

    df["CURRENCY"] = (
        df[colonne_currency]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # NUMERO COMPTE
    # ========================================================

    df["NUMERO COMPTE"] = ""

    # ========================================================
    # TYPE TRANSACTION
    #
    # PROVISOIRE :
    # NiirPay est actuellement considéré comme débit/W2B.
    # ========================================================

    df["TYPE TRANSACTION"] = "D"

    # ========================================================
    # MONTANT AVEC FRAIS
    # ========================================================

    df["MONTANT_AVEC_FRAIS"] = (
        df["MONTANT"].fillna(0)
        +
        df["FRAIS"].fillna(0)
    )

    # ========================================================
    # FICHIER SOURCE
    # ========================================================

    df["_FICHIER_SOURCE"] = (
        getattr(
            upload_file,
            "filename",
            "",
        )
        or ""
    )

    # ========================================================
    # SUPPRESSION DES LIGNES SANS MONTANT
    # ========================================================

    df = df[
        df["MONTANT"].notna()
    ].copy()

    df = df.reset_index(
        drop=True
    )

    # ========================================================
    # LOGS
    # ========================================================

    print(
        "--------------------------------------------------"
    )

    print(
        "[app_niirpay.py] "
        f"Nombre de transactions : {len(df)}"
    )

    print(
        "[app_niirpay.py] "
        "Montant total : "
        f"{df['MONTANT'].sum():,.2f}"
    )

    print(
        "--------------------------------------------------"
    )

    return df


# ============================================================
# SEPARATION W2B / B2W
# ============================================================

def separer_w2b_b2w(
    df: pd.DataFrame,
):
    """
    NiirPay est traité comme B2W.

    B2W :
        toutes les transactions

    W2B :
        aucune transaction
    """

    df = df.copy()

    df_w2b = df.iloc[
        0:0
    ].copy()

    df_b2w = df.copy()

    return (
        df_w2b,
        df_b2w,
    )

# ============================================================
# SAUVEGARDE SQLITE
# ============================================================

def sauvegarder_niirpay(
    df: pd.DataFrame,
):
    """
    Sauvegarde :

        COMPILATION_NIIRPAY
        COMPILATION_NIIRPAY_W2B
        COMPILATION_NIIRPAY_B2W
    """

    # ========================================================
    # SECURITE SQLITE
    # ========================================================

    df = supprimer_doublons_colonnes_sqlite(
        df
    )

    # ========================================================
    # SEPARATION W2B / B2W
    # ========================================================

    df_w2b, df_b2w = separer_w2b_b2w(
        df
    )

    # ========================================================
    # TABLE PRINCIPALE
    # ========================================================

    try:

        sauvegarder_sqlite(
            df,
            TABLES["excel"],
            engine,
            "app_niirpay",
        )

        print(
            "[app_niirpay] Table "
            f"'{TABLES['excel']}' écrite : "
            f"{len(df)} lignes"
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=(
                "Erreur sauvegarde SQLite "
                f"(table '{TABLES['excel']}') : "
                f"{e}"
            ),
        )

    # ========================================================
    # TABLE W2B
    # ========================================================

    try:

        sauvegarder_sqlite(
            df_w2b,
            TABLES["excel_w2b"],
            engine,
            "app_niirpay",
        )

        print(
            "[app_niirpay] Table "
            f"'{TABLES['excel_w2b']}' écrite : "
            f"{len(df_w2b)} lignes"
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=(
                "Erreur sauvegarde SQLite "
                f"(table '{TABLES['excel_w2b']}') : "
                f"{e}"
            ),
        )

    # ========================================================
    # TABLE B2W
    # ========================================================

    try:

        sauvegarder_sqlite(
            df_b2w,
            TABLES["excel_b2w"],
            engine,
            "app_niirpay",
        )

        print(
            "[app_niirpay] Table "
            f"'{TABLES['excel_b2w']}' écrite : "
            f"{len(df_b2w)} lignes"
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=(
                "Erreur sauvegarde SQLite "
                f"(table '{TABLES['excel_b2w']}') : "
                f"{e}"
            ),
        )

    return (
        df_w2b,
        df_b2w,
    )


# ============================================================
# ENDPOINT PRINCIPAL
# ============================================================

@app.post(
    "/process-excel"
)
async def process_excel(
    files: List[UploadFile] = File(
        ...,
        description=(
            "Fichier(s) CSV NiirPay. "
            "Le gateway envoie une liste ; "
            "NiirPay traite chaque fichier."
        ),
    ),

    format: str = Query(
        "excel",
        description="excel (défaut) ou json",
    ),
):
    """
    Endpoint principal de chargement NiirPay.

    IMPORTANT :

        files est une LISTE.

    Le gateway appelle call_partner_upload(),
    qui construit lui-même une List[UploadFile].
    """

    # ========================================================
    # VERIFICATION
    # ========================================================

    if not files:

        raise HTTPException(
            status_code=400,
            detail=(
                "Aucun fichier NiirPay reçu."
            ),
        )

    # ========================================================
    # VARIABLES
    # ========================================================

    dataframes = []

    fichiers_ok = []

    fichiers_erreurs = []

    # ========================================================
    # TRAITEMENT DES FICHIERS
    # ========================================================

    for fichier in files:

        try:

            print(
                "=================================================="
            )

            print(
                "[app_niirpay.py] "
                f"Fichier reçu : {fichier.filename}"
            )

            print(
                "=================================================="
            )

            # ------------------------------------------------
            # Traitement
            # ------------------------------------------------

            df = traiter_niirpay(
                fichier.file
            )

            # ------------------------------------------------
            # Nom du fichier source
            # ------------------------------------------------

            df["_FICHIER_SOURCE"] = (
                fichier.filename
                or ""
            )

            # ------------------------------------------------
            # Ajout
            # ------------------------------------------------

            dataframes.append(
                df
            )

            fichiers_ok.append(
                fichier.filename
            )

        except HTTPException:

            raise

        except Exception as e:

            print(
                "[app_niirpay.py] "
                f"Erreur traitement "
                f"'{fichier.filename}' :"
            )

            traceback.print_exc()

            fichiers_erreurs.append(
                {
                    "fichier":
                        fichier.filename,

                    "erreur":
                        str(e),
                }
            )

    # ========================================================
    # AUCUN FICHIER TRAITE
    # ========================================================

    if not dataframes:

        raise HTTPException(
            status_code=500,
            detail=(
                "Aucun fichier NiirPay "
                "n'a pu être traité. "
                f"Détails : {fichiers_erreurs}"
            ),
        )

    # ========================================================
    # CONCATENATION
    # ========================================================

    df_final = pd.concat(
        dataframes,
        ignore_index=True,
    )

    # ========================================================
    # SAUVEGARDE SQLITE
    # ========================================================

    df_w2b, df_b2w = sauvegarder_niirpay(
        df_final
    )

    # ========================================================
    # HEADERS
    # ========================================================

    export_headers = {

        "X-Nb-Fichiers":
            str(len(fichiers_ok)),

        "X-Nb-Transactions":
            str(len(df_final)),

        "X-Nb-W2B":
            str(len(df_w2b)),

        "X-Nb-B2W":
            str(len(df_b2w)),
    }

    # ========================================================
    # FORMAT JSON
    # ========================================================

    if not wants_excel(
        format
    ):

        return respond_sheets(
            {
                "NIIRPAY":
                    df_final,

                "NIIRPAY_W2B":
                    df_w2b,

                "NIIRPAY_B2W":
                    df_b2w,
            },

            filename="NIIRPAY.xlsx",

            format="json",

            headers=export_headers,

            json_payload={

                "status":
                    "ok",

                "format":
                    "json",

                "filename":
                    "NIIRPAY.xlsx",

                "fichiers":
                    fichiers_ok,

                "nb_fichiers":
                    len(fichiers_ok),

                "nb_transactions":
                    int(len(df_final)),

                "nb_w2b":
                    int(len(df_w2b)),

                "nb_b2w":
                    int(len(df_b2w)),

                "erreurs":
                    fichiers_erreurs,
            },
        )

    # ========================================================
    # PREPARATION EXCEL
    # ========================================================

    sheets = {

        "NIIRPAY":
            supprimer_timezone_pour_excel(
                df_final
            ),

        "NIIRPAY_W2B":
            supprimer_timezone_pour_excel(
                df_w2b
            ),

        "NIIRPAY_B2W":
            supprimer_timezone_pour_excel(
                df_b2w
            ),
    }

    # ========================================================
    # ERREURS EVENTUELLES
    # ========================================================

    if fichiers_erreurs:

        sheets[
            "Fichiers_Erreurs"
        ] = pd.DataFrame(
            fichiers_erreurs
        )

    # ========================================================
    # GENERATION EXCEL
    # ========================================================

    try:

        return respond_sheets(
            sheets,

            filename="NIIRPAY.xlsx",

            format="excel",

            headers=export_headers,
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=(
                "Sauvegarde SQLite OK, "
                "mais erreur lors de la "
                "génération du fichier Excel : "
                f"{e}"
            ),
        )


# ============================================================
# CONSULTATION SQLITE
# ============================================================

@app.get(
    "/db/compilation"
)
def get_compilation(
    limit: int = Query(None),
    offset: int = Query(0),
):
    """
    Retourne la compilation NiirPay.
    """

    return lire_table_json(
        engine,
        TABLES["excel"],
        limit=limit,
        offset=offset,
    )


# ============================================================
# CONSULTATION W2B
# ============================================================

@app.get(
    "/db/compilation-w2b"
)
def get_compilation_w2b(
    limit: int = Query(None),
    offset: int = Query(0),
):
    """
    Retourne les transactions NiirPay W2B.
    """

    return lire_table_json(
        engine,
        TABLES["excel_w2b"],
        limit=limit,
        offset=offset,
    )


# ============================================================
# CONSULTATION B2W
# ============================================================

@app.get(
    "/db/compilation-b2w"
)
def get_compilation_b2w(
    limit: int = Query(None),
    offset: int = Query(0),
):
    """
    Retourne les transactions NiirPay B2W.
    """

    return lire_table_json(
        engine,
        TABLES["excel_b2w"],
        limit=limit,
        offset=offset,
    )


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get(
    "/health"
)
def health():
    """
    Vérification du service NiirPay.
    """

    return {

        "status":
            "ok",

        "partenaire":
            PARTENAIRE,

        "db_path":
            DB_PATH,

        "tables": {

            "excel":
                TABLES["excel"],

            "excel_w2b":
                TABLES["excel_w2b"],

            "excel_b2w":
                TABLES["excel_b2w"],
        },
    }