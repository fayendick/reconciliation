# ============================================================
# APP ORANGE USSD — service Excel dédié au partenaire Orange (USSD Partenaire)

# ============================================================
# uvicorn app_orange_ussd:app --port 8040
#
# Contrairement à Wave, le fichier Orange "Daily-ChannelUserTransaction
# Report ... USSD_Partenaire.xls" n'est PAS déjà au format standard :
# il contient un bandeau d'en-tête, des blocs "Commissions cumulées" /
# "Transactions échouées" / totaux à éliminer, deux colonnes "N° de
# Compte" (Agent puis Correspondant) et deux colonnes séparées
# Débit/Crédit au lieu d'une colonne MONTANT unique.
#
# -> nettoyage_ussd_orange.nettoyer_ussd_orange() fait le ménage
#    (garde uniquement les vraies transactions "Succès"), puis ce
#    fichier fait le mapping vers le schéma standard attendu par
#    reconciliation_engine.py :
#
#    DATE_HEURE                  = Date + Heure (concaténées et
#                                   parsées ensemble)
#    TYPE TRANSACTION            = Service -> "W2B" (Cash in) / "B2W" (Cash Out)
#    CODE TRANSACTION OPERATEUR  = Référence
#    NUMERO COMPTE                = N° de Compte (Correspondant) = n°
#                                   de téléphone du client
#    DEBIT / CREDIT               = colonnes Débit / Crédit du fichier
#                                   Orange, conservées SÉPARÉMENT (au
#                                   lieu d'être fusionnées en un MONTANT
#                                   unique) : c'est reconciliation_engine.
#                                   preparer_wave_partenaire() qui choisit
#                                   laquelle des deux comparer selon le
#                                   sens — [CORRECTIF] convention en
#                                   partie double, alignée sur le
#                                   commentaire ORANGE_USSD de config.py :
#                                     W2B -> DEBIT  (comparé à
#                                            MOUVEMENT_CREDIT côté Flex)
#                                     B2W -> CREDIT (comparé à
#                                            MOUVEMENT_DEBIT côté Flex)
#                                   (et non plus W2B->CREDIT / B2W->DEBIT
#                                   comme l'ancienne version du moteur le
#                                   faisait à tort — voir
#                                   reconciliation_engine.py).
#
# ⚠️ Ce nouveau schéma (DATE_HEURE + DEBIT + CREDIT) diffère de
# l'ancien schéma (DATE TRANSACTION + MONTANT) encore utilisé par
# d'autres partenaires (Wave, Wizz, ...). reconciliation_engine.
# preparer_wave_partenaire() détecte automatiquement lequel des deux
# schémas est utilisé selon les colonnes présentes, donc les deux
# schémas cohabitent sans se casser mutuellement. Mais si
# excel_common.valider_colonnes_standard() / config.COLONNES_STANDARD_EXCEL
# imposent une liste de colonnes fixe et identique pour tous les
# partenaires, il faudra les mettre à jour pour accepter ce nouveau
# schéma côté ORANGE_USSD (à vérifier dans ces fichiers — NON VÉRIFIÉ
# faute d'avoir excel_common.py : c'est la piste la plus probable si
# le résultat reste vide malgré le correctif de reconciliation_engine.py).
# ============================================================

import io
import traceback
from typing import List

import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import StreamingResponse

from config import DB_PATH, PARTENAIRES, make_sqlite_engine
from common.excel_common import (
    valider_colonnes_standard,
    compiler_dataframes,
    separer_w2b_b2w,
    sauvegarder_excel_w2b_b2w,
)
from common.http_export import respond_sheets
from common.nettoyage_ussd_orange import nettoyer_ussd_orange
from common.sqlite_io import lire_table_json

PARTENAIRE = "ORANGE_USSD"
TABLES = PARTENAIRES[PARTENAIRE]["tables"]

print(f"[app_orange_ussd.py] Base SQLite utilisée : {DB_PATH}")

engine = make_sqlite_engine()

app = FastAPI(title="Excel Upload API — Orange USSD Partenaire")


# ------------------------------------------------------------
# Normalisation "Cash in" / "Cash Out" -> "W2B" / "B2W".
#
# IMPORTANT : excel_common.separer_w2b_b2w() compare TYPE TRANSACTION
# à la chaîne EXACTE "W2B" ou "B2W" (pas "CASH IN"/"CASH OUT") :
#     type_transaction == "W2B" / type_transaction == "B2W"
# Il faut donc produire ces deux codes précis ici, avec la même
# convention utilisée côté orange_ussd_flex_api.py, sinon les deux
# sous-tables (excel_w2b / excel_b2w) seraient vides.
#
# Convention retenue (comme Wave) :
#   Cash in  (le client dépose du cash chez l'agent) -> "W2B"
#   Cash Out (le client retire du cash)               -> "B2W"
# ------------------------------------------------------------
def _normaliser_type_transaction(service: pd.Series) -> pd.Series:
    return service.astype(str).str.strip().str.upper().replace({
        "CASH IN": "W2B",
        "CASH OUT": "B2W",
        
        
    
    })


def mapper_vers_schema_standard(df_brut: pd.DataFrame) -> pd.DataFrame:
    """Transforme un DataFrame déjà nettoyé (sortie de nettoyer_ussd_orange)
    vers le schéma standard attendu par reconciliation_engine :
    DATE_HEURE (fusion Date+Heure), TYPE TRANSACTION, CODE TRANSACTION
    OPERATEUR, NUMERO COMPTE, DEBIT, CREDIT."""

    df = pd.DataFrame()

    date_heure = (
        df_brut["Date"].astype(str).str.strip()
        + " "
        + df_brut["Heure"].astype(str).str.strip()
    )
    df["DATE_HEURE"] = pd.to_datetime(date_heure, errors="coerce", dayfirst=True)

    df["TYPE TRANSACTION"] = _normaliser_type_transaction(df_brut["Service"])
    df["CODE TRANSACTION OPERATEUR"] = df_brut["Référence"].astype(str).str.strip()

    df["NUMERO COMPTE"] = (
        df_brut["N° de Compte (Correspondant)"].astype(str).str.strip()
    )

    # DEBIT et CREDIT gardées séparées (pas de fusion en MONTANT unique) :
    # c'est reconciliation_engine.preparer_wave_partenaire() qui choisit
    # laquelle comparer selon le sens (voir correctif en tête de fichier).
    df["DEBIT"] = pd.to_numeric(df_brut["Débit"], errors="coerce").fillna(0)
    df["CREDIT"] = pd.to_numeric(df_brut["Crédit"], errors="coerce").fillna(0)

    return df


def lire_et_mapper_fichiers(files: List[UploadFile]) -> List[pd.DataFrame]:
    dfs = []
    for f in files:
        contenu = f.file.read()
        df_brut = nettoyer_ussd_orange(io.BytesIO(contenu), garder_uniquement_succes=True)
        dfs.append(mapper_vers_schema_standard(df_brut))
    return dfs


@app.post("/process-excel")
async def process_excel(
    files: List[UploadFile] = File(...),
    format: str = Query("excel", description="excel (défaut) ou json (skip openpyxl, pour /charger)"),
):

    try:
        # 1. LECTURE BRUTE + NETTOYAGE + MAPPING VERS LE SCHÉMA STANDARD
        dfs = lire_et_mapper_fichiers(files)

        for df in dfs:
            valider_colonnes_standard(df, PARTENAIRE)

        # 2. COMPILATION
        df_final = compiler_dataframes(dfs)

        # 3. SPLIT W2B (Cash In) / B2W (Cash Out)
        compile_w2b, compile_b2w = separer_w2b_b2w(df_final)

        # 4. SAUVEGARDE SQLITE : 1 table + 2 vues
        sauvegarder_excel_w2b_b2w(df_final, TABLES, engine, "app_orange_ussd")

        # 5. EXPORT (Excel si demandé, sinon JSON léger pour /charger)
        return respond_sheets(
            {
                "Compilation": df_final,
                "Compilation_CashIn": compile_w2b,
                "Compilation_CashOut": compile_b2w,
            },
            filename="Compilation_ORANGE_USSD.xlsx",
            format=format,
        )

    except HTTPException:
        raise
    except Exception as e:
        return {"status": "error", "message": str(e), "trace": traceback.format_exc()}


@app.get("/db/compilation")
def get_compilation(limit: int = Query(None), offset: int = Query(0)):
    return lire_table_json(engine, TABLES["excel"], limit=limit, offset=offset)


@app.get("/db/cash-in")
def get_cash_in(limit: int = Query(None), offset: int = Query(0)):
    return lire_table_json(engine, TABLES["excel_w2b"], limit=limit, offset=offset)


@app.get("/db/cash-out")
def get_cash_out(limit: int = Query(None), offset: int = Query(0)):
    return lire_table_json(engine, TABLES["excel_b2w"], limit=limit, offset=offset)


@app.get("/health")
def health():
    return {"status": "ok", "partenaire": PARTENAIRE, "db_path": DB_PATH}