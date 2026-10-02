# ============================================================
# APP WESTERN - service Excel dédié au rapprochement par agence
# ============================================================

import io
import re
import traceback
from pathlib import Path
from typing import List, Optional

import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException, Query

from common.config import DB_PATH, make_sqlite_engine
from common.http_export import respond_sheets, wants_excel
from common.sqlite_io import ecrire_table, lire_table_json


# ============================================================
# CONFIGURATION
# ============================================================

PARTENAIRE = "WESTERN"

TABLE_SORTIE = "COMPILATION_WESTERN"

BASE_DIR = Path(__file__).resolve().parents[2]

MAPPING_PATH = BASE_DIR / "data" / "mappings" / "ADJ WU.xlsx"

print(f"[app_western.py] Base SQLite utilisée : {DB_PATH}")
print(f"[app_western.py] Mapping utilisé : {MAPPING_PATH}")


engine = make_sqlite_engine()

app = FastAPI(title="Excel Upload API — Western")


# ============================================================
# OUTILS
# ============================================================

def normaliser(valeur) -> str:
    """
    Normalise une valeur texte :
    - majuscules
    - suppression des accents
    - espaces propres
    """
    if pd.isna(valeur):
        return ""

    texte = str(valeur).upper().strip()

    import unicodedata

    texte = unicodedata.normalize("NFD", texte)
    texte = "".join(
        c for c in texte
        if unicodedata.category(c) != "Mn"
    )

    texte = re.sub(r"\s+", " ", texte)

    return texte.strip()


def trouver_colonne(
    colonnes,
    mots_cles: List[str]
) -> Optional[str]:
    """
    Recherche une colonne de manière robuste.
    """
    for colonne in colonnes:
        colonne_norm = normaliser(colonne)

        if all(
            normaliser(mot) in colonne_norm
            for mot in mots_cles
        ):
            return colonne

    return None


# ============================================================
# LECTURE DU MAPPING ADJ WU
# ============================================================

def lire_mapping() -> pd.DataFrame:

    if not MAPPING_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=(
                f"Fichier de mapping introuvable : "
                f"{MAPPING_PATH}"
            ),
        )

    try:
        mapping = pd.read_excel(MAPPING_PATH)
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=(
                f"Impossible de lire le mapping ADJ WU.xlsx : {e}"
            ),
        )

    mapping.columns = mapping.columns.astype(str).str.strip()

    colonnes_attendues = [
        "ADJ",
        "NOM AGENCE",
        "NUMERO AGENCE",
    ]

    manquantes = [
        colonne
        for colonne in colonnes_attendues
        if colonne not in mapping.columns
    ]

    if manquantes:
        raise HTTPException(
            status_code=500,
            detail=(
                "Colonnes manquantes dans ADJ WU.xlsx : "
                + ", ".join(manquantes)
            ),
        )

    mapping = mapping[
        colonnes_attendues
    ].copy()

    mapping["ADJ"] = (
        mapping["ADJ"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    mapping["NOM AGENCE"] = (
        mapping["NOM AGENCE"]
        .astype(str)
        .str.strip()
    )

    mapping["NUMERO AGENCE"] = pd.to_numeric(
        mapping["NUMERO AGENCE"],
        errors="coerce"
    ).astype("Int64")

    mapping = mapping.dropna(
        subset=["ADJ", "NUMERO AGENCE"]
    )

    mapping = mapping.drop_duplicates(
        subset=["ADJ"],
        keep="first"
    )

    return mapping.reset_index(drop=True)


# ============================================================
# EXTRACTION DE L'ADJ DEPUIS REMARQUES
# ============================================================

def extraire_adj(valeur) -> Optional[str]:
    """
    Extrait un identifiant du type ADJ057871
    depuis la colonne REMARQUES.

    Exemple :
    ADJ056759-100271810003-PAIEMENT WU...
    -> ADJ056759
    """

    texte = normaliser(valeur)

    match = re.search(
        r"\b(ADJ\d+)\b",
        texte
    )

    if match:
        return match.group(1)

    return None


# ============================================================
# LECTURE DU FICHIER WESTERN
# ============================================================

def lire_fichier_western(
    fichier: UploadFile
) -> pd.DataFrame:

    contenu = fichier.file.read()

    if not contenu:
        raise HTTPException(
            status_code=400,
            detail=f"Le fichier {fichier.filename} est vide.",
        )

    try:
        df = pd.read_excel(
            io.BytesIO(contenu)
        )
    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Impossible de lire le fichier "
                f"{fichier.filename} : {e}"
            ),
        )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Recherche robuste des colonnes
    # --------------------------------------------------------

    col_remarques = trouver_colonne(
        df.columns,
        ["REMARQUES"]
    )

    col_credit = trouver_colonne(
        df.columns,
        ["CREDIT"]
    )

    col_debit = trouver_colonne(
        df.columns,
        ["DEBIT"]
    )

    if col_remarques is None:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Colonne 'REMARQUES' introuvable "
                f"dans {fichier.filename}."
            ),
        )

    if col_credit is None:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Colonne 'CRÉDIT' introuvable "
                f"dans {fichier.filename}."
            ),
        )

    if col_debit is None:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Colonne 'DÉBIT' introuvable "
                f"dans {fichier.filename}."
            ),
        )

    # --------------------------------------------------------
    # Conservation du fichier source
    # --------------------------------------------------------

    df["FICHIER_SOURCE"] = fichier.filename

    # --------------------------------------------------------
    # Normalisation REMARQUES
    # --------------------------------------------------------

    df["REMARQUES_NORMALISEES"] = (
        df[col_remarques]
        .apply(normaliser)
    )

    # --------------------------------------------------------
    # Extraction ADJ
    # --------------------------------------------------------

    df["ADJ"] = (
        df[col_remarques]
        .apply(extraire_adj)
    )

    # --------------------------------------------------------
    # Identification de l'opération
    # --------------------------------------------------------

    remarques = df["REMARQUES_NORMALISEES"]

    df["TYPE_OPERATION"] = None

    df.loc[
        remarques.str.contains(
            "PAIEMENT WU",
            na=False
        ),
        "TYPE_OPERATION"
    ] = "PAIEMENT WU"

    df.loc[
        remarques.str.contains(
            "ENVOI WU",
            na=False
        ),
        "TYPE_OPERATION"
    ] = "ENVOI WU"

    # --------------------------------------------------------
    # On ne garde pour la réconciliation que :
    #
    # PAIEMENT WU
    # ENVOI WU
    #
    # Les autres opérations :
    # TTHU
    # COMMISSION
    # REMBOURSEMENT
    # etc.
    # restent disponibles dans le détail mais ne participent
    # pas à l'agrégation de réconciliation.
    # --------------------------------------------------------

    df["A_RECONCILIER"] = (
        df["TYPE_OPERATION"].notna()
    )

        # --------------------------------------------------------
    # NETTOYAGE DES MONTANTS WESTERN
    #
    # Format du fichier :
    #   XOF 2,188,125.00
    #   XOF 847,289.00
    #
    # Résultat souhaité :
    #   2188125
    #   847289
    #
    # Les .00 ne sont pas conservés.
    # --------------------------------------------------------

    def nettoyer_montant(valeur):

        if pd.isna(valeur):
            return 0

        texte = str(valeur).strip().upper()

        # Suppression de la devise
        texte = texte.replace("XOF", "")

        # Suppression des espaces
        texte = texte.replace(" ", "")

        # Suppression des séparateurs de milliers
        texte = texte.replace(",", "")

        if texte in ("", "-", "NAN", "NONE"):
            return 0

        try:
            return int(float(texte))
        except (ValueError, TypeError):
            return 0

    df["CREDIT"] = (
        df[col_credit]
        .apply(nettoyer_montant)
    )

    df["DEBIT"] = (
        df[col_debit]
        .apply(nettoyer_montant)
    )

    # Les montants sont maintenant des XOF entiers.

    # --------------------------------------------------------
    # Le solde sera :
    #
    # DEBIT - CREDIT
    # --------------------------------------------------------

    df["SOLDE"] = (
        df["DEBIT"] -
        df["CREDIT"]
    )

    return df


# ============================================================
# APPLICATION DU MAPPING ADJ -> AGENCE
# ============================================================

def appliquer_mapping(
    df: pd.DataFrame,
    mapping: pd.DataFrame
) -> pd.DataFrame:

    df = df.copy()

    mapping_lookup = mapping.set_index("ADJ")

    df["CODE_AGENCE"] = (
        df["ADJ"]
        .map(mapping_lookup["NUMERO AGENCE"])
    )

    df["NOM_AGENCE"] = (
        df["ADJ"]
        .map(mapping_lookup["NOM AGENCE"])
    )

    df["CODE_AGENCE"] = pd.array(
        df["CODE_AGENCE"],
        dtype="Int64"
    )

    return df


# ============================================================
# AGREGATION PAR AGENCE
# ============================================================

def agreger_par_agence(
    df: pd.DataFrame
) -> pd.DataFrame:

    colonnes_resultat = [
        "CODE_AGENCE",
        "NOM_AGENCE",
        "NB_TRANSACTIONS",
        "NB_PAIEMENT_WU",
        "NB_ENVOI_WU",
        "DEBIT_PARTENAIRE",
        "CREDIT_PARTENAIRE",
        "SOLDE_PARTENAIRE",
    ]

    if df.empty:
        return pd.DataFrame(
            columns=colonnes_resultat
        )

    # --------------------------------------------------------
    # Seulement les opérations destinées à la réconciliation
    # --------------------------------------------------------

    df_recon = df[
        df["A_RECONCILIER"]
        & df["CODE_AGENCE"].notna()
    ].copy()

    if df_recon.empty:
        return pd.DataFrame(
            columns=colonnes_resultat
        )

    # --------------------------------------------------------
    # Colonnes auxiliaires
    # --------------------------------------------------------

    df_recon["NB_PAIEMENT_WU"] = (
        df_recon["TYPE_OPERATION"]
        .eq("PAIEMENT WU")
        .astype(int)
    )

    df_recon["NB_ENVOI_WU"] = (
        df_recon["TYPE_OPERATION"]
        .eq("ENVOI WU")
        .astype(int)
    )

    # --------------------------------------------------------
    # UNE ligne par agence
    # --------------------------------------------------------

    resultat = (
        df_recon
        .groupby(
            [
                "CODE_AGENCE",
                "NOM_AGENCE",
            ],
            dropna=False
        )
        .agg(
            NB_TRANSACTIONS=(
                "TYPE_OPERATION",
                "count"
            ),
            NB_PAIEMENT_WU=(
                "NB_PAIEMENT_WU",
                "sum"
            ),
            NB_ENVOI_WU=(
                "NB_ENVOI_WU",
                "sum"
            ),
            DEBIT_PARTENAIRE=(
                "DEBIT",
                "sum"
            ),
            CREDIT_PARTENAIRE=(
                "CREDIT",
                "sum"
            ),
        )
        .reset_index()
    )

    # --------------------------------------------------------
    # Solde partenaire
    #
    # DEBIT - CREDIT
    # --------------------------------------------------------

    resultat["SOLDE_PARTENAIRE"] = (
        resultat["DEBIT_PARTENAIRE"]
        -
        resultat["CREDIT_PARTENAIRE"]
    )

    # --------------------------------------------------------
    # Colonnes de suivi qui seront alimentées lors de la
    # comparaison avec FLEX.
    #
    # Elles ne sont PAS calculées ici, car FLEX n'est pas
    # encore disponible.
    # --------------------------------------------------------

    resultat["DEBIT_FLEX"] = 0.0
    resultat["CREDIT_FLEX"] = 0.0
    resultat["SOLDE_FLEX"] = 0.0

    resultat["CREDIT_SURPLUS_PARTENAIRE"] = 0.0

    resultat["ECART_SOLDE"] = 0.0

    resultat["TAXE"] = 0.0

    resultat["STATUT"] = "A RAPPROCHER"

    ordre = [
        "CODE_AGENCE",
        "NOM_AGENCE",
        "NB_TRANSACTIONS",
        "NB_PAIEMENT_WU",
        "NB_ENVOI_WU",
        "DEBIT_PARTENAIRE",
        "CREDIT_PARTENAIRE",
        "SOLDE_PARTENAIRE",
        "DEBIT_FLEX",
        "CREDIT_FLEX",
        "SOLDE_FLEX",
        "CREDIT_SURPLUS_PARTENAIRE",
        "ECART_SOLDE",
        "TAXE",
        "STATUT",
    ]

    return (
        resultat[ordre]
        .sort_values("CODE_AGENCE")
        .reset_index(drop=True)
    )


# ============================================================
# SAUVEGARDE SQLITE
# ============================================================

def sauvegarder_sqlite(
    df: pd.DataFrame
):
    ecrire_table(
        df,
        TABLE_SORTIE,
        engine,
        log_prefix="app_western"
    )


# ============================================================
# NOM FEUILLE EXCEL
# ============================================================

def nom_feuille(
    nom_fichier: str,
    index: int
) -> str:

    base = nom_fichier.rsplit(
        ".",
        1
    )[0]

    for caractere in [
        "\\",
        "/",
        "?",
        "*",
        "[",
        "]",
        ":",
    ]:
        base = base.replace(
            caractere,
            "_"
        )

    base = base.strip() or "Fichier"

    suffixe = f"_{index}"

    return (
        base[
            :31 - len(suffixe)
        ]
        + suffixe
    )


def rapprocher_western_par_agence(df_partenaire, df_flex):
    """
    Rapprochement WESTERN par agence.

    ENVOI WU :
        Partenaire DEBIT
        <-> Flex ACCOUNT_NO 101100000001 / TRN_CODE 121 / DEBIT

    PAIEMENT WU :
        Partenaire CREDIT
        <-> Flex ACCOUNT_NO 101100000001 / TRN_CODE 122 / CREDIT

    Le rapprochement est réalisé par agence.
    """

    # ============================================================
    # 1. Préparation partenaire
    # ============================================================

    df_p = df_partenaire.copy()

    df_p["DEBIT"] = pd.to_numeric(
        df_p["DEBIT"], errors="coerce"
    ).fillna(0).astype(int)

    df_p["CREDIT"] = pd.to_numeric(
        df_p["CREDIT"], errors="coerce"
    ).fillna(0).astype(int)

    df_p["CODE_AGENCE"] = pd.to_numeric(
        df_p["CODE_AGENCE"], errors="coerce"
    ).astype("Int64")

    # ============================================================
    # 2. Préparation Flex
    # ============================================================

    df_f = df_flex.copy()

    df_f["DEBIT"] = pd.to_numeric(
        df_f["DEBIT"], errors="coerce"
    ).fillna(0).astype(int)

    df_f["CREDIT"] = pd.to_numeric(
        df_f["CREDIT"], errors="coerce"
    ).fillna(0).astype(int)

    df_f["CODE_AGENCE"] = pd.to_numeric(
        df_f["CODE_AGENCE"], errors="coerce"
    ).astype("Int64")

    df_f["ACCOUNT_NO"] = (
        df_f["ACCOUNT_NO"]
        .astype(str)
        .str.strip()
    )

    df_f["TRN_CODE"] = (
        df_f["TRN_CODE"]
        .astype(str)
        .str.strip()
    )

    # ============================================================
    # 3. Flex utilisé pour les ENVOI WU
    # ============================================================

    flex_envoi = df_f[
        (df_f["ACCOUNT_NO"] == "101100000001")
        & (df_f["TRN_CODE"] == "121")
        & (df_f["DEBIT"] > 0)
    ].copy()

    flex_envoi_agence = (
        flex_envoi
        .groupby("CODE_AGENCE", dropna=False)["DEBIT"]
        .sum()
        .rename("DEBIT_FLEX")
        .reset_index()
    )

    # ============================================================
    # 4. Flex utilisé pour les PAIEMENT WU
    # ============================================================

    flex_paiement = df_f[
        (df_f["ACCOUNT_NO"] == "101100000001")
        & (df_f["TRN_CODE"] == "122")
        & (df_f["CREDIT"] > 0)
    ].copy()

    flex_paiement_agence = (
        flex_paiement
        .groupby("CODE_AGENCE", dropna=False)["CREDIT"]
        .sum()
        .rename("CREDIT_FLEX")
        .reset_index()
    )

    # ============================================================
    # 5. Agrégation partenaire
    # ============================================================

    partenaire_agence = (
        df_p[df_p["A_RECONCILIER"]]
        .groupby(
            ["CODE_AGENCE", "NOM_AGENCE"],
            dropna=False
        )
        .agg(
            NB_TRANSACTIONS=("TYPE_OPERATION", "size"),
            NB_PAIEMENT_WU=(
                "TYPE_OPERATION",
                lambda s: (s == "PAIEMENT WU").sum()
            ),
            NB_ENVOI_WU=(
                "TYPE_OPERATION",
                lambda s: (s == "ENVOI WU").sum()
            ),
            DEBIT_PARTENAIRE=("DEBIT", "sum"),
            CREDIT_PARTENAIRE=("CREDIT", "sum"),
        )
        .reset_index()
    )

    # ============================================================
    # 6. Rapprochement Flex
    # ============================================================

    resultat = partenaire_agence.merge(
        flex_envoi_agence,
        on="CODE_AGENCE",
        how="left"
    )

    resultat = resultat.merge(
        flex_paiement_agence,
        on="CODE_AGENCE",
        how="left"
    )

    resultat["DEBIT_FLEX"] = (
        resultat["DEBIT_FLEX"]
        .fillna(0)
        .astype(int)
    )

    resultat["CREDIT_FLEX"] = (
        resultat["CREDIT_FLEX"]
        .fillna(0)
        .astype(int)
    )

    # ============================================================
    # 7. Calcul des écarts
    # ============================================================

    resultat["ECART_DEBIT"] = (
        resultat["DEBIT_PARTENAIRE"]
        - resultat["DEBIT_FLEX"]
    )

    resultat["ECART_CREDIT"] = (
        resultat["CREDIT_PARTENAIRE"]
        - resultat["CREDIT_FLEX"]
    )

    # ============================================================
    # 8. Solde
    # ============================================================

    resultat["SOLDE_PARTENAIRE"] = (
        resultat["DEBIT_PARTENAIRE"]
        - resultat["CREDIT_PARTENAIRE"]
    )

    resultat["SOLDE_FLEX"] = (
        resultat["DEBIT_FLEX"]
        - resultat["CREDIT_FLEX"]
    )

    # ============================================================
    # 9. TAXE
    #
    # La taxe est déterminée à partir du surplus de crédit
    # constaté entre le partenaire et Flex.
    #
    # Aucune donnée "TAXE" provenant de Flexcube n'est utilisée.
    # ============================================================

    resultat["TAXE"] = (
        resultat["CREDIT_PARTENAIRE"]
        - resultat["CREDIT_FLEX"]
    )
    # ============================================================
    # 10. Crédit surplus partenaire
    # ============================================================

    resultat["CREDIT_SURPLUS_PARTENAIRE"] = (
        resultat["CREDIT_PARTENAIRE"]
        - resultat["CREDIT_FLEX"]
    )

    # ============================================================
    # 11. Statut
    # ============================================================

    def determiner_statut(row):
    
        ecart_credit = abs(int(row["ECART_CREDIT"]))

    # Règle métier WESTERN :
    # l'écart de crédit est considéré comme la TAXE.
    # L'agence est réconciliée si l'écart ne dépasse pas 15 000 XOF.
        if ecart_credit <= 15000:
            return "Réconcilié"

        return "Ecart montant"


    resultat["STATUT"] = resultat.apply(
        determiner_statut,
        axis=1
    )

    # ============================================================
    # 12. Ordre final
    # ============================================================

    colonnes = [
        "CODE_AGENCE",
        "NOM_AGENCE",
        "NB_TRANSACTIONS",
        "NB_PAIEMENT_WU",
        "NB_ENVOI_WU",
        "DEBIT_PARTENAIRE",
        "DEBIT_FLEX",
        "ECART_DEBIT",
        "CREDIT_PARTENAIRE",
        "CREDIT_FLEX",
        "ECART_CREDIT",
        "CREDIT_SURPLUS_PARTENAIRE",
        "TAXE",
        "SOLDE_PARTENAIRE",
        "SOLDE_FLEX",
        "STATUT",
    ]

    return resultat[colonnes].sort_values(
        "CODE_AGENCE"
    ).reset_index(drop=True)


# ============================================================
# ENDPOINT PRINCIPAL
# ============================================================

@app.post("/map-agences")
async def map_agences(
    files: List[UploadFile] = File(...),
    format: str = Query(
        "excel",
        description=(
            "excel (défaut) ou json"
        ),
    ),
):

    if not files:
        raise HTTPException(
            status_code=400,
            detail="Aucun fichier envoyé.",
        )

    try:

        # ----------------------------------------------------
        # Lecture du mapping
        # ----------------------------------------------------

        mapping = lire_mapping()

        dfs_traites = []

        # ----------------------------------------------------
        # Traitement de chaque fichier
        # ----------------------------------------------------

        for fichier in files:

            western = lire_fichier_western(
                fichier
            )

            western = appliquer_mapping(
                western,
                mapping
            )

            dfs_traites.append(
                (
                    fichier.filename,
                    western
                )
            )

        # ----------------------------------------------------
        # Détail combiné
        # ----------------------------------------------------

        df_detail = pd.concat(
            [
                df
                for _, df in dfs_traites
            ],
            ignore_index=True
        )

        # ----------------------------------------------------
        # Agrégation par agence
        # ----------------------------------------------------

        df_agrege = agreger_par_agence(
            df_detail
        )

        # ----------------------------------------------------
        # Sauvegarde SQLite
        # ----------------------------------------------------

        sauvegarder_sqlite(
            df_agrege
        )

        # ----------------------------------------------------
        # Export Excel
        # ----------------------------------------------------

        sheets = {
            "Compilation": df_agrege,
            "Detail_Compilation": df_detail,
        }

        for i, (
            nom_fichier,
            df
        ) in enumerate(
            dfs_traites,
            start=1
        ):
            sheets[
                nom_feuille(
                    nom_fichier,
                    i
                )
            ] = df

        # ----------------------------------------------------
        # Réponse Excel
        # ----------------------------------------------------

        if wants_excel(format):

            return respond_sheets(
                sheets,
                filename="WESTERN_MAPPE.xlsx",
                format="excel",
            )

        # ----------------------------------------------------
        # Réponse JSON
        # ----------------------------------------------------

        return respond_sheets(
            {
                "Compilation": df_agrege
            },
            filename="WESTERN_MAPPE.xlsx",
            format="json",
            json_payload={
                "status": "ok",
                "format": "json",
                "filename": "WESTERN_MAPPE.xlsx",
                "nb_agences": int(
                    len(df_agrege)
                ),
                "nb_transactions": int(
                    len(df_detail)
                ),
                "nb_transactions_reconciliation": int(
                    df_detail[
                        df_detail[
                            "A_RECONCILIER"
                        ]
                        & df_detail[
                            "CODE_AGENCE"
                        ].notna()
                    ].shape[0]
                ),
                "nb_fichiers": len(
                    dfs_traites
                ),
            },
        )

    except HTTPException:
        raise

    except Exception as e:

        return {
            "status": "error",
            "message": str(e),
            "trace": traceback.format_exc(),
        }


# ============================================================
# CONSULTATION SQLITE
# ============================================================

@app.get("/db/western")
def get_western(
    limit: int = Query(None),
    offset: int = Query(0),
):

    return lire_table_json(
        engine,
        TABLE_SORTIE,
        limit=limit,
        offset=offset,
    )


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/health")
def health():

    return {
        "status": "ok",
        "partenaire": PARTENAIRE,
        "db_path": str(DB_PATH),
        "mapping": str(MAPPING_PATH),
        "table": TABLE_SORTIE,
    }