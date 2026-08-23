# ============================================================
# ORANGE USSD FLEX — service Oracle Flexcube dédié au partenaire
# Orange USSD Partenaire (mode "two_pointers")
# ============================================================
# uvicorn orange_ussd_flex_api:app --port 8041
#
# Adapté de la requête Oracle fournie (écritures ACVW_ALL_AC_ENTRIES
# sur le compte 379200000180, jointes au KYC pour récupérer le numéro
# de téléphone du client via RELATED_CUSTOMER).
#
# Produit les colonnes attendues par reconciliation_engine.preparer_wave_flex :
#   DATE_VALEUR                 = date+heure de l'écriture (SAVE_TIMESTAMP)
#   CODE_TRANSACTION_OPERATEUR  = référence de la transaction (TRN_REF_NO)
#   NUMERO_COMPTE                = n° de téléphone du client (MOBILE_NUMBER)
#   MOUVEMENT_DEBIT / MOUVEMENT_CREDIT = montant_Debit / montant_Credit
#   TYPE_TRANSACTION             = "W2B" (Cash In côté app_orange_ussd.py)
#                                   ou "B2W" (Cash Out)
#
# IMPORTANT : ces codes "W2B"/"B2W" doivent être EXACTEMENT les mêmes
# que ceux produits par app_orange_ussd.py (_normaliser_type_transaction),
# car reconciliation_two_pointers() compare littéralement
# wp["TYPE TRANSACTION"] == wf["TYPE_TRANSACTION"].
#
# --------------------------------------------------------------
# [CORRECTIF] INVERSION W2B/B2W (drcr_ind <-> Cash In/Cash Out) :
#   (voir historique complet dans la version précédente du fichier —
#   inchangé, non reproduit ici pour rester concis)
# --------------------------------------------------------------
#
# [CORRECTIF 2026-08-23] BUG "Query(379200000180) not supported" :
#
#   CAUSE RÉELLE (confirmée via gateway/mount_partners.py) : les
#   endpoints Flex sont invoqués EN PROCESS via call_partner_flex(),
#   qui récupère la fonction Python brute derrière la route
#   (route.endpoint) et l'appelle directement avec endpoint(**kwargs)
#   — SANS jamais passer par le mécanisme ASGI / dependency-injection
#   de FastAPI qui résout normalement Query(...) en sa vraie valeur.
#
#   Conséquence : si l'appelant (gateway/reconc.py) ne fournit pas
#   explicitement "compte_agent" dans les params transmis à
#   call_partner_flex(), le paramètre reçoit la valeur par défaut
#   LITTÉRALE du signature Python, qui était `Query(ORANGE_USSD_...)`
#   — l'objet FastAPI lui-même — d'où l'erreur cx_Oracle
#   "Python value of type Query not supported".
#
#   -> FIX : le défaut de `compte_agent` n'utilise plus Query(...)
#      mais une valeur Python simple (ORANGE_USSD_COMPTE_AGENT).
#      Un défaut simple reste parfaitement valide pour FastAPI en
#      cas de vrai appel HTTP (le paramètre reste bien un query
#      param optionnel), et fonctionne aussi correctement en appel
#      direct in-process comme le fait call_partner_flex().
#      Seule la description OpenAPI (visible dans /docs) est perdue,
#      ce qui est un compromis acceptable ici.
#
#   ⚠️ Le même risque existe potentiellement pour tout AUTRE
#   paramètre optionnel à défaut Query(...) dans les autres modules
#   *_flex_api.py (wave, wizz, ria_agence, wave_agence...) s'ils sont
#   eux aussi invoqués via call_partner_flex() sans que reconc.py ne
#   fournisse systématiquement toutes leurs clés optionnelles. À
#   vérifier au cas par cas si un problème similaire réapparaît.
# --------------------------------------------------------------

import os
import traceback
from datetime import date, datetime
from typing import Optional

import pandas as pd
from fastapi import HTTPException, Query
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    register_db_routes,
    save_split_w2b_b2w,
    split_by_type_exact,
)

PARTENAIRE = "ORANGE_USSD"
rt = bootstrap_flex_service(PARTENAIRE, title="Orange USSD Flex API", log_prefix="orange_ussd_flex_api")
app = rt.app
TABLES = rt.tables
TABLE_FLEX = TABLES["flex"]
TABLE_FLEX_W2B = TABLES["flex_w2b"]
TABLE_FLEX_B2W = TABLES["flex_b2w"]

# Compte agent Orange (Tête de réseau) sur lequel filtrer les écritures.
ORANGE_USSD_COMPTE_AGENT = os.getenv("ORANGE_USSD_COMPTE_AGENT", "379200000180")


SQL_ORANGE_USSD_FLEX = text("""
WITH KYC AS (
    SELECT
        p.CUSTOMER_NO,
        p.TELEPHONE AS MOBILE_NUMBER
    FROM CFSFCUBS145.STTM_CUSTOMER sc
    LEFT JOIN CFSFCUBS145.STTM_CUST_PERSONAL p
        ON sc.CUSTOMER_NO = p.CUSTOMER_NO
),
Journal AS (
    SELECT
        TRN_REF_NO, AC_ENTRY_SR_NO, EVENT_SR_NO, EVENT, AC_BRANCH, AC_NO,
        DRCR_IND, TRN_CODE, LCY_AMOUNT, VALUE_DT, TRN_DT, TXN_INIT_DATE,
        RELATED_ACCOUNT, RELATED_CUSTOMER, SAVE_TIMESTAMP, MODULE,
        BATCH_NO, CURR_NO
    FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES
    WHERE MODULE = 'DE'
    UNION
    SELECT
        TRN_REF_NO, AC_ENTRY_SR_NO, EVENT_SR_NO, EVENT, AC_BRANCH, AC_NO,
        DRCR_IND, TRN_CODE, LCY_AMOUNT, VALUE_DT, TRN_DT, TXN_INIT_DATE,
        RELATED_ACCOUNT, RELATED_CUSTOMER, SAVE_TIMESTAMP, MODULE,
        BATCH_NO, CURR_NO
    FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES
    WHERE MODULE <> 'DE'
)
SELECT
    a.SAVE_TIMESTAMP                                    AS DATE_VALEUR,
    a.TRN_REF_NO                                        AS CODE_TRANSACTION_OPERATEUR,
    kyc.MOBILE_NUMBER                                    AS NUMERO_COMPTE,
    -- [CORRECTIF] inversé : voir bloc de commentaire en tête de fichier.
    -- Un Cash In (W2B) partenaire correspond en réalité à une écriture
    -- DÉBIT (drcr_ind='D') côté Flex, et un Cash Out (B2W) à une
    -- écriture CRÉDIT (drcr_ind='C').
    DECODE(a.DRCR_IND, 'C', a.LCY_AMOUNT, 0)             AS MOUVEMENT_DEBIT,
    DECODE(a.DRCR_IND, 'D', a.LCY_AMOUNT, 0)             AS MOUVEMENT_CREDIT,
    DECODE(a.DRCR_IND, 'D', 'W2B', 'C', 'B2W')           AS TYPE_TRANSACTION,
    a.AC_BRANCH                                          AS CODE_AGENCE,
    br.BRANCH_NAME                                        AS LIBELLE_AGENCE
FROM Journal a
LEFT JOIN KYC kyc ON a.RELATED_CUSTOMER = kyc.CUSTOMER_NO
LEFT JOIN CFSFCUBS145.STTM_BRANCH br ON br.BRANCH_CODE = a.AC_BRANCH
LEFT JOIN CFSFCUBS145.gltm_glmaster c ON c.gl_code = a.AC_NO
LEFT JOIN CFSFCUBS145.STTM_CUST_ACCOUNT s ON s.CUST_AC_NO = a.AC_NO
WHERE
    a.TRN_REF_NO IN (
        SELECT TRN_REF_NO FROM CFSFCUBS145.ACVW_ALL_AC_ENTRIES
        WHERE AC_NO = :compte_agent
    )
    AND a.SAVE_TIMESTAMP >= :date_debut
    AND a.SAVE_TIMESTAMP < :date_fin + 1
    AND NVL(c.GL_DESC, s.AC_DESC) NOT IN (
        'COMPTES TTRANSFERT POUR VIREMENT INTERNE',
        'COMPTE ATTENTE ORANGE USSD',
        ' COMPTES DE LIAISON'
    )
    AND a.BATCH_NO IS NULL
ORDER BY a.VALUE_DT DESC
""")


def _parser_date_fr(valeur: str) -> date:
    """Parse une date reçue en DD/MM/YYYY (format envoyé par reconc.py,
    identique à ce qu'attendent les autres services flex du projet).
    Accepte aussi le format ISO YYYY-MM-DD en repli, au cas où."""
    valeur = (valeur or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(valeur, fmt).date()
        except ValueError:
            continue
    raise HTTPException(
        status_code=422,
        detail=f"Date invalide : '{valeur}' (formats acceptés : DD/MM/YYYY ou YYYY-MM-DD)",
    )


def _executer_requete(date_debut: date, date_fin: date, compte_agent: str) -> pd.DataFrame:
    return oracle_query(
        rt,
        SQL_ORANGE_USSD_FLEX,
        {
            "compte_agent": compte_agent,
            "date_debut": datetime.combine(date_debut, datetime.min.time()),
            "date_fin": datetime.combine(date_fin, datetime.min.time()),
        },
    )


def orange_ussd_flex_logique(
    date_debut: str,
    date_fin: Optional[str] = None,
    compte_agent: str = ORANGE_USSD_COMPTE_AGENT,
):
    """Fonction métier PURE (aucune dépendance à fastapi.Query).

    C'est la fonction à utiliser pour tout appel direct en Python
    (ex: depuis gateway/reconc.py), afin de ne jamais recevoir un
    objet Query() non résolu à la place d'une vraie valeur.
    La route FastAPI ci-dessous n'est qu'un adaptateur HTTP autour
    de cette fonction.
    """
    debut = _parser_date_fr(date_debut)
    borne_fin = _parser_date_fr(date_fin) if date_fin else debut

    df = _executer_requete(debut, borne_fin, compte_agent)
    flex_w2b, flex_b2w = split_by_type_exact(df)

    return save_split_w2b_b2w(
        rt,
        df,
        sheets={
            "ORANGE_USSD_FLEX": df,
            "ORANGE_USSD_FLEX_W2B": flex_w2b,
            "ORANGE_USSD_FLEX_B2W": flex_b2w,
        },
        filename="ORANGE_USSD_FLEX.xlsx",
        format="json",
        mode="exact",
        json_payload={
            "status": "ok",
            "total": len(df),
            "cash_in_w2b": len(flex_w2b),
            "cash_out_b2w": len(flex_b2w),
        },
    )


@app.get("/orange-ussd-flex")
def orange_ussd_flex(
    date_debut: str = Query(..., description="Date de début (incluse), format DD/MM/YYYY"),
    date_fin: Optional[str] = Query(None, description="Date de fin (exclue). Par défaut date_debut + 1 jour, format DD/MM/YYYY"),
    # [CORRECTIF] défaut simple (pas Query(...)) : voir bloc de commentaire
    # en tête de fichier — call_partner_flex() appelle cette fonction en
    # process, en contournant la résolution FastAPI de Query(...).
    compte_agent: str = ORANGE_USSD_COMPTE_AGENT,
):
    """Route HTTP : extrait les paramètres de la requête via Query()
    (résolus par FastAPI ici, dans le contexte d'un vrai appel ASGI),
    puis délègue toute la logique à orange_ussd_flex_logique()."""
    try:
        return orange_ussd_flex_logique(date_debut, date_fin, compte_agent)

    except HTTPException:
        raise
    except Exception as e:
        # IMPORTANT : ne jamais renvoyer une erreur avec un code 200 —
        # reconc.py considérerait alors, à tort, que les tables Flex
        # ont bien été écrites alors que rien n'a été sauvegardé.
        raise HTTPException(
            status_code=502,
            detail=f"Échec de la requête Oracle Flex (ORANGE_USSD) : {e}\n{traceback.format_exc()}",
        )


register_db_routes(rt, [
    ("/db/flex", TABLE_FLEX),
    ("/db/flex-cash-in", TABLE_FLEX_W2B),
    ("/db/flex-cash-out", TABLE_FLEX_B2W),
])