# ============================================================
# RECONCILIATION API (GATEWAY) — MULTI-PARTENAIRE UNIFIÉ
# ============================================================
# uvicorn reconc:app --port 8002
#
# API UNIQUE : Module_FED et Streamlit parlent uniquement à :8002.
# Les services Excel / Flex de chaque partenaire sont montés sous
# /svc/<nom>/... (voir mount_partners.py). /charger appelle encore
# upload_url / flex_url (désormais en loopback sur le même port).
#
# NOTE IMPORTANTE : toute la construction visuelle (graphiques
# Plotly ET carte résumé HTML) est faite ICI, côté API. Chaque
# endpoint /db/reconciliation-graphe-* renvoie une figure Plotly
# complète et autonome (fig.to_json()), et /db/reconciliation-carte-resume
# renvoie du HTML déjà prêt à afficher + les valeurs formatées des
# 3 colonnes ajoutées (diff_montant_partenaire, diff_montant_flexcube,
# ecart_difference). Le client (streamlit_app.py ou tout autre
# consommateur) n'a plus qu'à afficher tel quel (pio.from_json +
# st.plotly_chart, ou st.markdown / st.metric) : AUCUNE construction,
# calcul, couleur ou mise en forme à faire de son côté.
#
# DEUX FAMILLES D'ENDPOINTS DE RÉCONCILIATION, LIÉES AU "mode" DE
# CHAQUE PARTENAIRE (voir config.py) :
#   - /reconciliation/run          : Two Pointers, transaction par
#     transaction, W2B + B2W. Réservé aux partenaires dont
#     config["mode"] == "two_pointers" (Wave, Wizz, ...) — ce sont
#     les seuls à avoir des tables excel_w2b/excel_b2w/flex_w2b/
#     flex_b2w réellement peuplées.
#   - /reconciliation/run-agence   : agrégé par CODE_AGENCE, compare
#     DEBIT (Flex) <-> col_debit (Partenaire) et CREDIT (Flex) <->
#     col_credit (Partenaire). Réservé aux partenaires dont
#     config["mode"] == "agence" (Orange Agence, Wave Agence), dont
#     l'Excel partenaire est déjà une compilation par agence (pas de
#     transaction individuelle ni de date à apparier, donc pas de
#     table excel_w2b/excel_b2w).
#     -> nécessite dans config.PARTENAIRES[...]["tables"] une clé
#        supplémentaire "reconciliation_agence" (nom de la table
#        SQLite dédiée à ce résultat), en plus de "excel" et "flex"
#        déjà utilisées par app_orange.py/orange_flex_api.py ou
#        wave_app_agence.py/wave_agence_flex_api.py.
#     -> nécessite aussi une clé "colonnes_agence" (col_credit /
#        col_debit) précisant le nom des deux colonnes montant côté
#        Excel-partenaire pour CE partenaire (voir config.py) —
#        c'est ce qui permet au même moteur générique de servir
#        Orange Agence (montant_cashin/montant_cashout) ET Wave
#        Agence (DEPOT/RETRAIT) sans aucune modification du moteur.
#
# Appeler le mauvais endpoint pour le mauvais mode renvoie désormais
# une erreur 400 explicite (au lieu d'un "no such table" cryptique
# provoqué par la lecture d'une table qui n'a jamais été créée).
#
# ------------------------------------------------------------
# [CORRECTIF] /charger — DÉTECTION PRÉCOCE D'UNE EXTRACTION FLEX VIDE
# ------------------------------------------------------------
# AVANT : /charger renvoyait "success" dès que l'upload Excel (200)
#         ET l'extraction Flex (200) répondaient, SANS vérifier que
#         l'extraction Flex avait réellement trouvé des lignes. Pour
#         un partenaire en mode "agence" (ex: WAVE_AGENCE), si la
#         période demandée ne contenait aucune écriture Oracle (mauvais
#         date_debut/date_fin, filtre SQL trop restrictif...), la table
#         Flex (ex: WAVE_AGENCE_FLEX) était simplement écrasée avec 0
#         ligne -> /charger répondait quand même "succès", et l'erreur
#         n'apparaissait QUE plus tard, au moment de
#         /reconciliation/run-agence, avec un message générique
#         "Table source vide pour WAVE_AGENCE" difficile à rattacher à
#         sa cause réelle (date de chargement incorrecte).
#
# APRÈS : pour un partenaire dont le service Flex expose l'en-tête
#         HTTP "X-Nb-Agences" (c'est le cas de wave_agence_flex_api.py,
#         voir /wave-agence-flex), /charger vérifie cet en-tête et
#         renvoie IMMÉDIATEMENT une 400 explicite si 0 agence a été
#         trouvée, plutôt que de laisser l'erreur remonter plus tard
#         pendant la réconciliation. Si l'en-tête est absent (services
#         plus anciens qui ne l'exposent pas encore), le comportement
#         reste inchangé (pas de vérification, comme avant) — cette
#         vérification est donc rétro-compatible.
# ------------------------------------------------------------
#
# ------------------------------------------------------------
# [CORRECTIF] /reconciliation/run — APPARIEMENT ORANGE USSD
# ------------------------------------------------------------
# Pour Wave/Wizz, CODE_TRANSACTION est la même référence des deux
# côtés (Excel-partenaire et Flex), donc l'appariement historique
# par CODE_TRANSACTION identique fonctionne bien et N'EST PAS
# TOUCHÉ : par défaut, cfg.get("apparier_par_telephone_montant",
# False) vaut False pour ces partenaires (clé absente de leur
# config), donc reconcilier_un_sens() garde exactement son
# comportement d'avant.
#
# Pour Orange USSD, CODE_TRANSACTION désigne DEUX identifiants
# différents (référence opérateur télécom côté Excel vs référence
# interne Oracle TRN_REF_NO côté Flex) qui ne coïncident jamais :
# l'appariement se fait donc par NUMERO_COMPTE (téléphone client)
# normalisé + TYPE TRANSACTION, avec préférence pour un montant
# identique puis pour l'écart de temps le plus faible (voir
# reconciliation_engine.reconciliation_two_pointers). Activé
# uniquement si config.PARTENAIRES["ORANGE_USSD"]
# ["apparier_par_telephone_montant"] == True.
# ------------------------------------------------------------
# ============================================================

import io
import json
from datetime import datetime
from typing import List

import pandas as pd
import plotly.graph_objects as go
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import StreamingResponse, JSONResponse, Response

from common.reconciliation_engine import (
    reconcilier_un_sens,
    construire_table_resume,
    calculer_taux_reussite,
    STATUTS,
    STATUT_COULEURS,
    reconcilier_par_agence,
    construire_table_resume_agence,
)
from config import DB_PATH, PARTENAIRES, get_partenaire, get_mode, MODE_TWO_POINTERS, MODE_AGENCE, make_sqlite_engine
from common.sqlite_io import ecrire_table, df_json_response, df_to_records
from gateway.mount_partners import mount_partner_apps, call_partner_upload, call_partner_flex

app = FastAPI(title="Reconciliation Gateway API — Multi-partenaire")

REQUEST_TIMEOUT = 60

print(f"[reconc.py] Base SQLite utilisée : {DB_PATH}")

engine = make_sqlite_engine()

# Monte Excel + Flex de tous les partenaires sous /svc/...
_MOUNTED = mount_partner_apps(app)
print(f"[reconc.py] Services partenaires montés : {len(_MOUNTED)}")


# ============================================================
# HELPERS
# ============================================================

def safe_query(sql: str, limit: int | None = None, offset: int = 0):
    try:
        df = pd.read_sql(sql, engine)
        # NaN/Inf ne sont pas JSON-compliant (sinon HTTP 500 Starlette)
        # Sérialisation via to_json (évite to_dict orient=records).
        return df_json_response(df, limit=limit, offset=offset)
    except Exception as e:
        print(f"[reconc.py] ERREUR safe_query('{sql}') : {e}")
        return Response(content="[]", media_type="application/json")


def read_table(table_name: str) -> pd.DataFrame:
    try:
        df = pd.read_sql(f"SELECT * FROM {table_name}", engine)
    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail=f"Table '{table_name}' introuvable en base ({e}). "
                    f"As-tu bien lancé le chargement Excel + Flex pour ce partenaire avant ?"
        )
    return df


def cfg_or_400(partenaire: str) -> dict:
    try:
        return get_partenaire(partenaire)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


def read_reconciliation_or_empty(table_name: str) -> pd.DataFrame:
    """Comme read_table, mais renvoie un DataFrame vide (au lieu de lever
    une erreur 400) si la table n'existe pas encore — utile pour les
    endpoints de synthèse/graphe qu'on veut pouvoir appeler avant la
    première réconciliation, sans faire planter la page Streamlit."""
    try:
        return pd.read_sql(f"SELECT * FROM {table_name}", engine)
    except Exception:
        return pd.DataFrame()


def formater_nb(nb) -> str:
    """1234 -> '1 234' (séparateur de milliers façon FR)."""
    try:
        return f"{int(round(nb)):,}".replace(",", " ")
    except (ValueError, TypeError):
        return "0"


def formater_pct(pct: float) -> str:
    """12.34 -> '12,34' (virgule décimale façon FR)."""
    return f"{pct:.2f}".replace(".", ",")


def verifier_mode(partenaire: str, cfg: dict, mode_attendu: str, endpoint_attendu: str):
    """Lève une 400 explicite si le partenaire n'est pas dans le mode
    attendu par l'endpoint appelé, en indiquant le bon endpoint à
    utiliser à la place — plutôt que de laisser un read_table() planter
    plus loin sur une table qui n'a jamais été créée pour ce mode."""
    mode_reel = cfg.get("mode", MODE_TWO_POINTERS)
    if mode_reel != mode_attendu:
        raise HTTPException(
            status_code=400,
            detail=(
                f"'{partenaire}' est configuré en mode '{mode_reel}', pas '{mode_attendu}'. "
                f"{endpoint_attendu}"
            )
        )


# ============================================================
# LISTE DES PARTENAIRES DISPONIBLES (pour le selectbox Streamlit)
# ------------------------------------------------------------
# Expose désormais aussi "mode" ("two_pointers" ou "agence") pour
# que streamlit_app.py puisse adapter son UI (onglets, tableaux
# W2B/B2W affichés ou non) sans dupliquer cette info en dur côté
# client.
# ============================================================

@app.get("/partenaires")
def liste_partenaires():
    return [
        {"key": key, "label": cfg["label"], "mode": cfg.get("mode", MODE_TWO_POINTERS)}
        for key, cfg in PARTENAIRES.items()
    ]


@app.get("/statuts")
def liste_statuts():
    """Liste fixe et ordonnée des statuts possibles — utilisée par
    streamlit_app.py pour toujours afficher les 6 statuts, même à 0."""
    return STATUTS


# ============================================================
# CHARGEMENT COMPLET (Excel dédié + Flex dédié) — PARAMÉTRÉ
# ============================================================

@app.post("/charger")
async def charger(
    partenaire: str = Query(..., description="Ex: WAVE, ORANGE_AGENCE, WIZZ, WAVE_AGENCE"),
    files: List[UploadFile] = File(...),
    date_debut: str = "01/01/2000",
    date_fin: str = "01/01/2100",
):
    """
    1. Envoie les fichiers Excel vers le service Excel DÉDIÉ à ce partenaire
    2. Lance l'extraction Flex sur le service Oracle DÉDIÉ à ce partenaire
    3. [CORRECTIF] Si le service Flex expose l'en-tête "X-Nb-Agences"
       (cas de wave_agence_flex_api.py) et qu'il vaut 0, on échoue
       IMMÉDIATEMENT avec un message explicite, au lieu de laisser
       l'erreur remonter plus tard, de façon plus confuse, au moment
       de la réconciliation ("Table source vide pour ...").
    """

    cfg = cfg_or_400(partenaire)

    files_payload = []
    for f in files:
        content = await f.read()
        files_payload.append(("files", (f.filename, content, f.content_type)))

    # Appels in-process — pas de HTTP loopback sur le même worker.
    excel_response = await call_partner_upload(
        partenaire,
        files_payload,
        params={"format": "json"},
    )

    if excel_response.status_code != 200:
        raise HTTPException(
            status_code=502,
            detail=f"Erreur Upload Excel ({partenaire}) : {excel_response.text}"
        )

    flex_response = await call_partner_flex(
        partenaire,
        params={"date_debut": date_debut, "date_fin": date_fin, "format": "json"},
    )

    if flex_response.status_code != 200:
        raise HTTPException(
            status_code=502,
            detail=f"Erreur extraction Flex ({partenaire}) : {flex_response.text}"
        )

    # --------------------------------------------------------
    # [CORRECTIF] Vérification précoce : 0 ligne côté Flex.
    # L'en-tête "X-Nb-Agences" n'existe que sur les services Flex qui
    # l'exposent explicitement (wave_agence_flex_api.py) — absent
    # ailleurs, donc aucun impact sur les autres partenaires/services
    # plus anciens.
    # --------------------------------------------------------
    nb_agences_header = flex_response.headers.get("X-Nb-Agences")
    if nb_agences_header is not None and nb_agences_header.isdigit() and int(nb_agences_header) == 0:
        nb_ecritures_header = flex_response.headers.get("X-Nb-Ecritures", "?")
        raise HTTPException(
            status_code=400,
            detail=(
                f"Extraction Flex ({partenaire}) : 0 écriture Oracle trouvée "
                f"(nb_ecritures_brutes={nb_ecritures_header}) pour la période "
                f"{date_debut} -> {date_fin}. Vérifie que cette période contient "
                f"bien des transactions (élargis-la si besoin), et que le filtre "
                f"SQL (compte GL 'COMPTE ATTENTE  WAVE', liste de comptes) est "
                f"correct. Rien n'a été écrasé côté Excel-partenaire, mais la "
                f"table Flex de {partenaire} est maintenant vide : ne lance pas "
                f"la réconciliation tant que ce point n'est pas corrigé."
            )
        )

    return {
        "status": "success",
        "partenaire": partenaire,
        "mode": cfg.get("mode", MODE_TWO_POINTERS),
        "message": f"Chargement {cfg['label']} et extraction Flex terminés.",
        "nb_agences_flex": int(nb_agences_header) if nb_agences_header and nb_agences_header.isdigit() else None,
    }


# ============================================================
# DB READ ENDPOINTS — PARAMÉTRÉS PAR PARTENAIRE
# ------------------------------------------------------------
# excel-w2b / excel-b2w / flex-w2b / flex-b2w n'existent que pour
# les partenaires en mode "two_pointers". Pour un partenaire en
# mode "agence" (Orange Agence, Wave Agence), ces tables ne sont
# jamais créées : safe_query() capture l'erreur SQLite et renvoie []
# proprement, donc ces endpoints restent utilisables sans planter —
# mais streamlit_app.py ne doit plus les appeler pour ces partenaires
# (voir adaptation de l'UI selon le mode).
# ============================================================

@app.get("/db/excel")
def excel(partenaire: str = Query(...), limit: int = Query(None), offset: int = Query(0)):
    cfg = cfg_or_400(partenaire)
    return safe_query(f"SELECT * FROM {cfg['tables']['excel']}", limit=limit, offset=offset)


@app.get("/db/excel-w2b")
def excel_w2b(partenaire: str = Query(...), limit: int = Query(None), offset: int = Query(0)):
    cfg = cfg_or_400(partenaire)
    return safe_query(f"SELECT * FROM {cfg['tables']['excel_w2b']}", limit=limit, offset=offset)


@app.get("/db/excel-b2w")
def excel_b2w(partenaire: str = Query(...), limit: int = Query(None), offset: int = Query(0)):
    cfg = cfg_or_400(partenaire)
    return safe_query(f"SELECT * FROM {cfg['tables']['excel_b2w']}", limit=limit, offset=offset)


@app.get("/db/flex")
def flex(partenaire: str = Query(...), limit: int = Query(None), offset: int = Query(0)):
    cfg = cfg_or_400(partenaire)
    return safe_query(f"SELECT * FROM {cfg['tables']['flex']}", limit=limit, offset=offset)


@app.get("/db/flex-w2b")
def flex_w2b(partenaire: str = Query(...), limit: int = Query(None), offset: int = Query(0)):
    cfg = cfg_or_400(partenaire)
    return safe_query(f"SELECT * FROM {cfg['tables']['flex_w2b']}", limit=limit, offset=offset)


@app.get("/db/flex-b2w")
def flex_b2w(partenaire: str = Query(...), limit: int = Query(None), offset: int = Query(0)):
    cfg = cfg_or_400(partenaire)
    return safe_query(f"SELECT * FROM {cfg['tables']['flex_b2w']}", limit=limit, offset=offset)


@app.get("/db/tables")
def db_tables():
    try:
        df = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table'", engine)
        return {"tables": df["name"].tolist()}
    except Exception as e:
        return {"error": str(e)}


# ============================================================
# RÉCONCILIATION — Two Pointers (W2B + B2W), PARAMÉTRÉE
# ------------------------------------------------------------
# Réservée aux partenaires en mode "two_pointers". Pour un
# partenaire en mode "agence" (Orange Agence, Wave Agence), les
# tables excel_w2b / excel_b2w n'ont jamais été créées par son
# service Excel dédié : on bloque donc l'appel ICI, avant même de
# tenter read_table(), avec un message qui renvoie vers le bon
# endpoint.
#
# [CORRECTIF] Le paramètre apparier_par_telephone_montant est lu
# depuis cfg (config.PARTENAIRES[...]["apparier_par_telephone_montant"])
# et transmis à reconcilier_un_sens(). Absent de la config -> False
# par défaut -> AUCUN changement de comportement pour Wave/Wizz.
# Positionné à True uniquement pour ORANGE_USSD (voir config.py).
# ============================================================

@app.post("/reconciliation/run")
def run_reconciliation(partenaire: str = Query(...)):

    cfg = cfg_or_400(partenaire)

    verifier_mode(
        partenaire, cfg, MODE_TWO_POINTERS,
        "Ce partenaire est agrégé par agence : utilise "
        "POST /reconciliation/run-agence à la place."
    )

    t = cfg["tables"]

    wp_w2b_raw = read_table(t["excel_w2b"])
    wp_b2w_raw = read_table(t["excel_b2w"])
    wf_w2b_raw = read_table(t["flex_w2b"])
    wf_b2w_raw = read_table(t["flex_b2w"])

    for nom, df in [
        (f"Excel W2B ({partenaire})", wp_w2b_raw), (f"Excel B2W ({partenaire})", wp_b2w_raw),
        (f"Flex W2B ({partenaire})", wf_w2b_raw), (f"Flex B2W ({partenaire})", wf_b2w_raw),
    ]:
        if df.empty:
            raise HTTPException(
                status_code=400,
                detail=f"La table source '{nom}' est vide. "
                        f"Lance le chargement Excel + Flex pour {partenaire} avant de réconcilier."
            )

    # [CORRECTIF] False par défaut (clé absente pour Wave/Wizz) ->
    # comportement historique inchangé. True uniquement pour
    # ORANGE_USSD (voir config.py).
    apparier_par_telephone_montant = cfg.get("apparier_par_telephone_montant", False)

    try:
        resultat_w2b = reconcilier_un_sens(
            wp_w2b_raw, wf_w2b_raw, "W2B",
            apparier_par_telephone_montant=apparier_par_telephone_montant,
        )
        resultat_w2b["SENS"] = "W2B"

        resultat_b2w = reconcilier_un_sens(
            wp_b2w_raw, wf_b2w_raw, "B2W",
            apparier_par_telephone_montant=apparier_par_telephone_montant,
        )
        resultat_b2w["SENS"] = "B2W"

        resultat_final = pd.concat([resultat_w2b, resultat_b2w], ignore_index=True)

    except KeyError as e:
        raise HTTPException(
            status_code=400,
            detail=f"Colonne attendue absente pour la réconciliation ({partenaire}) : {e}. "
                    f"Vérifie les colonnes des tables sources."
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur pendant la réconciliation ({partenaire}) : {e}")

    # Les doublons (Excel-partenaire ET Flex, W2B et B2W confondus) sont
    # déjà présents dans resultat_final avec STATUT="Doublon" (détectés
    # et retirés AVANT le Two Pointers par reconciliation_engine.py).
    doublons = resultat_final[resultat_final["STATUT"] == "Doublon"].copy()

    ecrire_table(resultat_final, t["reconciliation"], engine, log_prefix="reconc")
    ecrire_table(doublons, t["doublons"], engine, log_prefix="reconc")

    print(f"[reconc.py] Réconciliation {partenaire} terminée : {len(resultat_final)} lignes -> {t['reconciliation']}")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resultat_final.to_excel(writer, sheet_name="Resultat_Global", index=False)

        # Table résumé (schéma métier fixe) dans son propre onglet
        cfg_resume = cfg.get("colonnes_resume", {})
        table_resume = construire_table_resume(resultat_final, cfg_resume)
        table_resume.to_excel(writer, sheet_name="Resume", index=False)

        # Un onglet par statut, TOUJOURS dans le même ordre (STATUTS),
        # et TOUJOURS créé même si 0 ligne (juste les en-têtes) — pour
        # rester cohérent avec le donut/résumé/filtre, qui affichent
        # déjà tous les statuts même à 0.
        for statut in STATUTS:
            subset = resultat_final[resultat_final["STATUT"] == statut]
            sheet_name = str(statut)[:31]
            subset.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Resultat_Reconciliation_{partenaire}.xlsx"}
    )


@app.get("/db/reconciliation")
def get_reconciliation(partenaire: str = Query(...), limit: int = Query(None), offset: int = Query(0)):
    cfg = cfg_or_400(partenaire)
    return safe_query(f"SELECT * FROM {cfg['tables']['reconciliation']}", limit=limit, offset=offset)


# ============================================================
# TABLE RÉSUMÉ (schéma métier fixe)
# ============================================================

@app.get("/db/reconciliation-resume")
def get_reconciliation_resume(partenaire: str = Query(...)):
    """
    Renvoie le résultat de réconciliation aplati vers le schéma
    résumé : Type Transaction, Montant Partenaire, Montant Flex,
    Num_Tel_Client, Nom_Client, Compte, Agence, Ecart Montant,
    Diff Heure, Date Fichier Partenaire, Periode Fichier, Statut.
    """
    cfg = cfg_or_400(partenaire)
    resultat = read_reconciliation_or_empty(cfg["tables"]["reconciliation"])

    if resultat.empty:
        return []

    resume = construire_table_resume(resultat, cfg.get("colonnes_resume", {}))
    for col in ["Periode Fichier", "Date Fichier Partenaire"]:
        if col in resume.columns:
            resume[col] = resume[col].astype(str)

    return json.loads(resume.to_json(orient="records"))


# ============================================================
# TAUX DE RÉUSSITE DE LA RÉCONCILIATION
# ============================================================

@app.get("/db/reconciliation-taux")
def get_reconciliation_taux(partenaire: str = Query(...)):
    """
    Taux de réussite = (Réconcilié + Réconcilié avec tolérance) / total,
    calculé sur TOUS les statuts (écarts, non comptabilisées,
    comptabilisations isolées et doublons inclus dans le total).
    """
    cfg = cfg_or_400(partenaire)
    resultat = read_reconciliation_or_empty(cfg["tables"]["reconciliation"])
    return calculer_taux_reussite(resultat)


@app.get("/db/reconciliation-summary")
def get_reconciliation_summary(partenaire: str = Query(...)):
    """Nombre de lignes par statut — TOUJOURS les 6 statuts de STATUTS,
    avec NB=0 pour ceux qui n'ont aucune ligne (table absente incluse)."""
    cfg = cfg_or_400(partenaire)
    resultat = read_reconciliation_or_empty(cfg["tables"]["reconciliation"])
    detail = calculer_taux_reussite(resultat)
    return [{"STATUT": s, "NB": detail["comptes"][s]} for s in STATUTS]


# ============================================================
# CARTE RÉSUMÉ (HTML prêt à afficher, une seule ligne)
# ------------------------------------------------------------
# Reprend TOUJOURS la feuille Resultat_Global (= table SQLite
# "reconciliation" de ce partenaire) pour calculer :
#   Période fichier, Transactions Partenaire, Montant Partenaire,
#   Transactions Flexcube, Montant Flexcube, Diff. Montant
#   Partenaire (W2B-B2W), Diff. Montant Flexcube (W2B-B2W),
#   Ecart Différence, Matchées, Ecarts, Montant Ecart, Taux de
#   Match, Statut, Dernière maj.
# Renvoie :
#   - "html"                       : tableau HTML déjà entièrement
#                                     mis en forme (couleurs, badge
#                                     de statut...) -> st.markdown()
#   - "diff_montant_partenaire"    : valeur déjà formatée (FR)
#   - "diff_montant_flexcube"      : valeur déjà formatée (FR)
#   - "ecart_difference"           : valeur déjà formatée (FR)
# Le client n'a plus qu'à afficher tel quel (st.markdown / st.metric)
# — AUCUN calcul ni style à faire de son côté.
#
# Diff. Montant Partenaire = somme montant Partenaire (W2B) - somme
# montant Partenaire (B2W). Diff. Montant Flexcube = idem côté
# Flexcube. Ecart Différence = |diff Partenaire - diff Flexcube|.
# Regroupement fait via la colonne "SENS" (ajoutée par
# run_reconciliation), avec repli sur les colonnes TYPE TRANSACTION
# brutes pour un résultat calculé avant l'ajout de SENS (même
# logique que get_reconciliation_graphe_evolution). Concerne
# uniquement le mode "two_pointers" (les partenaires en mode
# "agence", Orange Agence et Wave Agence, n'appellent pas cet
# endpoint : leur carte résumé équivalente n'existe pas encore et
# n'est pas nécessaire, la table résumé par agence suffit).
# ============================================================

@app.get("/db/reconciliation-carte-resume")
def get_reconciliation_carte_resume(partenaire: str = Query(...)):

    cfg = cfg_or_400(partenaire)
    resultat = read_reconciliation_or_empty(cfg["tables"]["reconciliation"])

    if resultat.empty:
        return {
            "html": None,
            "diff_montant_partenaire": None,
            "diff_montant_flexcube": None,
            "ecart_difference": None,
        }

    date_heure_partenaire = pd.to_datetime(resultat.get("WP_DATE_HEURE"), errors="coerce")
    date_heure_flex = pd.to_datetime(resultat.get("WF_DATE_HEURE"), errors="coerce")
    date_heure = date_heure_partenaire.fillna(date_heure_flex).dropna()

    periode = (
        f"{date_heure.min().strftime('%d/%m/%Y')} - {date_heure.max().strftime('%d/%m/%Y')}"
        if not date_heure.empty else "—"
    )

    transactions_partenaire = int(date_heure_partenaire.notna().sum())
    transactions_flexcube = int(date_heure_flex.notna().sum())

    montant_partenaire = pd.to_numeric(resultat.get("WP_MONTANT_COMPARAISON"), errors="coerce").sum()
    montant_flexcube = pd.to_numeric(resultat.get("WF_MONTANT_COMPARAISON"), errors="coerce").sum()

    detail = calculer_taux_reussite(resultat)
    matchees = detail["reconcilies"]
    total = detail["total"]
    ecarts = total - matchees
    taux_match = detail["taux_reussite"]

    montant_ecart = abs(montant_partenaire - montant_flexcube)

    # --------------------------------------------------------
    # Différence W2B / B2W, séparément côté Partenaire et côté
    # Flexcube, + écart entre ces deux différences.
    # --------------------------------------------------------
    if "SENS" in resultat.columns:
        sens = resultat["SENS"].astype(str).str.upper().str.strip()
    else:
        sens = pd.Series(pd.NA, index=resultat.index, dtype=object)
        if "WP_TYPE TRANSACTION" in resultat.columns:
            sens = sens.fillna(resultat["WP_TYPE TRANSACTION"].astype(str).str.upper().str.strip())
        if "WF_TYPE_TRANSACTION" in resultat.columns:
            sens = sens.fillna(resultat["WF_TYPE_TRANSACTION"].astype(str).str.upper().str.strip())

    montant_partenaire_par_sens = pd.to_numeric(
        resultat.get("WP_MONTANT_COMPARAISON"), errors="coerce"
    ).groupby(sens).sum()
    montant_flexcube_par_sens = pd.to_numeric(
        resultat.get("WF_MONTANT_COMPARAISON"), errors="coerce"
    ).groupby(sens).sum()

    montant_partenaire_w2b = float(montant_partenaire_par_sens.get("W2B", 0.0) or 0.0)
    montant_partenaire_b2w = float(montant_partenaire_par_sens.get("B2W", 0.0) or 0.0)
    montant_flexcube_w2b = float(montant_flexcube_par_sens.get("W2B", 0.0) or 0.0)
    montant_flexcube_b2w = float(montant_flexcube_par_sens.get("B2W", 0.0) or 0.0)

    difference_montant_partenaire = montant_partenaire_w2b - montant_partenaire_b2w
    difference_montant_flexcube = montant_flexcube_w2b - montant_flexcube_b2w
    ecart_difference = abs(difference_montant_partenaire - difference_montant_flexcube)

    statut = "Terminé" if total > 0 else "En attente"
    derniere_maj = datetime.now().strftime("%d/%m/%Y %H:%M")

    th_style = "padding:10px 14px; color:#8a8a8a; font-weight:500; font-size:12px; text-transform:uppercase; letter-spacing:0.02em; white-space:nowrap;"
    td_style = "padding:16px 14px; font-size:14px; white-space:nowrap;"

    html = f"""
    <div style="overflow-x:auto; border:1px solid #eee; border-radius:10px;">
      <table style="width:100%; border-collapse:collapse;">
        <thead>
          <tr style="border-bottom:2px solid #f0f0f0; text-align:left;">
            <th style="{th_style}">Période fichier</th>
            <th style="{th_style}">Transactions Partenaire</th>
            <th style="{th_style}">Montant Partenaire (XOF)</th>
            <th style="{th_style}">Transactions Flexcube</th>
            <th style="{th_style}">Montant Flexcube (XOF)</th>
            <th style="{th_style}">Diff. Montant Partenaire (W2B-B2W)</th>
            <th style="{th_style}">Diff. Montant Flexcube (W2B-B2W)</th>
            <th style="{th_style}">Ecart Différence (XOF)</th>
            <th style="{th_style}">Matchées</th>
            <th style="{th_style}">Ecarts</th>
            <th style="{th_style}">Montant Ecart (XOF)</th>
            <th style="{th_style}">Taux de Match</th>
            <th style="{th_style}">Statut</th>
            <th style="{th_style}">Dernière maj</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="{td_style} font-weight:600;">{periode}</td>
            <td style="{td_style}">{formater_nb(transactions_partenaire)}</td>
            <td style="{td_style}">{formater_nb(montant_partenaire)}</td>
            <td style="{td_style}">{formater_nb(transactions_flexcube)}</td>
            <td style="{td_style}">{formater_nb(montant_flexcube)}</td>
            <td style="{td_style}">{formater_nb(difference_montant_partenaire)}</td>
            <td style="{td_style}">{formater_nb(difference_montant_flexcube)}</td>
            <td style="{td_style} color:#E67E22; font-weight:700;">{formater_nb(ecart_difference)}</td>
            <td style="{td_style} color:#1E8449; font-weight:700;">{formater_nb(matchees)}</td>
            <td style="{td_style} color:#E74C3C; font-weight:700;">{formater_nb(ecarts)}</td>
            <td style="{td_style} color:#E74C3C; font-weight:700;">{formater_nb(montant_ecart)}</td>
            <td style="{td_style} font-weight:700;">{formater_pct(taux_match)}%</td>
            <td style="{td_style}">
              <span style="background:#E8F8EF; color:#1E8449; padding:5px 14px; border-radius:14px; font-size:12px; font-weight:600;">{statut}</span>
            </td>
            <td style="{td_style} color:#999;">{derniere_maj}</td>
          </tr>
        </tbody>
      </table>
    </div>
    """

    return {
        "html": html,
        "diff_montant_partenaire": formater_nb(difference_montant_partenaire),
        "diff_montant_flexcube": formater_nb(difference_montant_flexcube),
        "ecart_difference": formater_nb(ecart_difference),
    }


# ============================================================
# GRAPHIQUE 1/2 — DONUT RÉPARTITION PAR STATUT
# ------------------------------------------------------------
# Figure Plotly autonome et complète : couleurs, légende
# ("Statut — nb (pct%)"), taux global au centre. Rien à
# retoucher côté client.
# ============================================================

@app.get("/db/reconciliation-graphe-statut")
def get_reconciliation_graphe_statut(partenaire: str = Query(...)):

    cfg = cfg_or_400(partenaire)
    resultat = read_reconciliation_or_empty(cfg["tables"]["reconciliation"])

    detail = calculer_taux_reussite(resultat)
    comptes = detail["comptes"]
    total = detail["total"]

    if total == 0:
        return JSONResponse(content=None)

    labels_legende = []
    valeurs = []
    couleurs = []

    for statut in STATUTS:
        nb = comptes[statut]
        pct = round(100 * nb / total, 2) if total > 0 else 0.0
        labels_legende.append(f"{statut} — {formater_nb(nb)} ({formater_pct(pct)}%)")
        valeurs.append(nb)
        couleurs.append(STATUT_COULEURS[statut])

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels_legende,
                values=valeurs,
                hole=0.6,
                sort=False,
                direction="clockwise",
                textinfo="none",
                marker=dict(colors=couleurs, line=dict(color="#FFFFFF", width=2)),
                hovertemplate="%{label}<extra></extra>",
            )
        ]
    )

    taux_str = formater_pct(detail["taux_reussite"])

    fig.update_layout(
        title=dict(text="RÉCONCILIATION PAR STATUT", font=dict(size=16, color="#1a237e"), x=0),
        height=380,
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(orientation="v", y=0.5, x=1.02, xanchor="left"),
        annotations=[
            dict(
                # [MODIF] Annotation centrale du donut réduite : taux
                # global plus petit (14 au lieu de 18) et libellé
                # "Taux global" plus discret (10 au lieu d'hériter de
                # la taille du bloc <b>), pour un centre moins imposant.
                text=f"<b style='font-size:14px;'>{taux_str}%</b><br><span style='font-size:10px;'>Taux global</span>",
                x=0.5, y=0.5,
                showarrow=False,
                font=dict(size=14, color="#1a237e"),
            )
        ],
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    return JSONResponse(content=json.loads(fig.to_json()))


# ============================================================
# GRAPHIQUE 2/2 — ÉVOLUTION DU MONTANT (PARTENAIRE vs FLEX)
# ------------------------------------------------------------
# Reprend TOUJOURS la feuille Resultat_Global (= table SQLite
# "reconciliation" de ce partenaire). Deux courbes : montant total
# côté Excel-partenaire (somme de WP_MONTANT_COMPARAISON) et côté
# Flexcube Oracle (somme de WF_MONTANT_COMPARAISON), regroupées sur
# une granularité de temps ADAPTATIVE (minute / heure / jour selon
# l'étendue réelle des données), pour le sens choisi (W2B ou B2W).
# Filtré via la colonne SENS ajoutée par run_reconciliation ; repli
# sur les colonnes TYPE TRANSACTION brutes si un résultat plus
# ancien (sans SENS) est encore en base.
#
# Courbe Partenaire en rouge, fine (les deux lignes restent fines
# pour ne pas se marcher dessus visuellement) ; un très léger
# décalage symétrique est appliqué à l'affichage seul (la valeur
# réelle reste dans le survol) pour que les deux courbes restent
# distinguables même quand les montants sont rigoureusement
# identiques (cas fréquent sur les lignes réconciliées).
#
# Pourquoi une granularité adaptative : sur un export réel étalé
# sur une journée entière avec ~1-2 transactions/minute, regrouper
# par minute donne des centaines de points quasi tous isolés
# (illisible). Regrouper par heure fait ressortir la vraie
# tendance (heures de pointe). Sur un jeu de test (quelques
# minutes), regrouper par heure écraserait tout en un seul point
# -> on reste alors par minute. On choisit donc le pas en
# fonction de l'étendue réelle :
#   <= 2h de données   -> par minute
#   <= 3 jours          -> par heure
#   au-delà             -> par jour
#
# Concerne uniquement le mode "two_pointers" (colonne SENS produite
# par run_reconciliation).
# ============================================================

def _choisir_granularite(etendue: pd.Timedelta):
    if etendue <= pd.Timedelta(hours=2):
        return "min", "%H:%M"
    if etendue <= pd.Timedelta(days=3):
        return "h", "%d/%m %Hh"
    return "D", "%d/%m"


@app.get("/db/reconciliation-graphe-evolution")
def get_reconciliation_graphe_evolution(
    partenaire: str = Query(...),
    type_transaction: str = Query("W2B"),
):

    cfg = cfg_or_400(partenaire)
    resultat = read_reconciliation_or_empty(cfg["tables"]["reconciliation"])

    if resultat.empty:
        return JSONResponse(content=None)

    type_transaction = type_transaction.upper().strip()

    if "SENS" in resultat.columns:
        df = resultat[resultat["SENS"] == type_transaction].copy()
    else:
        # Repli pour un résultat calculé avant l'ajout de la colonne SENS
        masque = pd.Series(False, index=resultat.index)
        if "WP_TYPE TRANSACTION" in resultat.columns:
            masque |= resultat["WP_TYPE TRANSACTION"].astype(str).str.upper().str.strip() == type_transaction
        if "WF_TYPE_TRANSACTION" in resultat.columns:
            masque |= resultat["WF_TYPE_TRANSACTION"].astype(str).str.upper().str.strip() == type_transaction
        df = resultat[masque].copy()

    if df.empty:
        return JSONResponse(content=None)

    # Horodatage : côté partenaire en priorité, repli côté flex (sert
    # uniquement à déterminer l'étendue réelle -> la granularité)
    df["WP_DATE_HEURE"] = pd.to_datetime(df.get("WP_DATE_HEURE"), errors="coerce")
    df["WF_DATE_HEURE"] = pd.to_datetime(df.get("WF_DATE_HEURE"), errors="coerce")
    df["_DATE_HEURE"] = df["WP_DATE_HEURE"].fillna(df["WF_DATE_HEURE"])

    df = df.dropna(subset=["_DATE_HEURE"])

    if df.empty:
        return JSONResponse(content=None)

    etendue = df["_DATE_HEURE"].max() - df["_DATE_HEURE"].min()
    pas, format_ticks = _choisir_granularite(etendue)
    df["_BUCKET"] = df["_DATE_HEURE"].dt.floor(pas)

    # Montant par bucket, séparément pour chaque côté (pas de repli
    # coalescé ici : Partenaire et Flexcube sont deux séries distinctes
    # à comparer visuellement, comme pour les comptages avant).
    df["_MONTANT_PARTENAIRE"] = pd.to_numeric(df.get("WP_MONTANT_COMPARAISON"), errors="coerce")
    df["_MONTANT_FLEX"] = pd.to_numeric(df.get("WF_MONTANT_COMPARAISON"), errors="coerce")

    evolution = (
        df.groupby("_BUCKET")
        .agg(
            PARTENAIRE=("_MONTANT_PARTENAIRE", "sum"),
            FLEXCUBE=("_MONTANT_FLEX", "sum"),
        )
        .sort_index()
    )

    y_partenaire = evolution["PARTENAIRE"]
    y_flexcube = evolution["FLEXCUBE"]

    # --------------------------------------------------------
    # Petit décalage visuel : quand les montants Partenaire et
    # Flexcube sont rigoureusement identiques (cas fréquent sur les
    # lignes réconciliées), les deux courbes se superposent
    # parfaitement et une seule reste visible. On applique un très
    # léger décalage symétrique (uniquement pour l'affichage — le
    # survol continue d'indiquer la vraie valeur via customdata) pour
    # que les deux courbes restent toujours distinguables.
    # --------------------------------------------------------
    amplitude = max(float(y_partenaire.max() or 0), float(y_flexcube.max() or 0))
    decalage = amplitude * 0.004 if amplitude > 0 else 0.0

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=evolution.index,
            y=y_partenaire + decalage,
            customdata=y_partenaire,
            mode="lines+markers",
            name="Partenaire",
            line=dict(color="#E53935", width=1.5),
            marker=dict(size=5),
            hovertemplate="<b>Partenaire</b><br>%{x}<br>Montant : %{customdata:,.0f}<extra></extra>",
        )
    )

    fig.add_trace(
        go.Scatter(
            x=evolution.index,
            y=y_flexcube - decalage,
            customdata=y_flexcube,
            mode="lines+markers",
            name="Flexcube",
            line=dict(color="#1565C0", width=1.5),
            marker=dict(size=5),
            hovertemplate="<b>Flexcube</b><br>%{x}<br>Montant : %{customdata:,.0f}<extra></extra>",
        )
    )

    fig.update_layout(
        title=dict(
            text=f"ÉVOLUTION DU MONTANT — {type_transaction}",
            font=dict(size=16, color="#1a237e"),
            x=0,
        ),
        height=380,
        margin=dict(l=20, r=20, t=60, b=30),
        legend=dict(orientation="h", y=1.15, x=0),
        xaxis=dict(title="Temps", tickformat=format_ticks),
        yaxis=dict(title="Montant"),
        hovermode="x unified",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    return JSONResponse(content=json.loads(fig.to_json()))


# ============================================================
# RÉCONCILIATION PAR AGENCE — CODE_AGENCE / DEBIT / CREDIT
# ------------------------------------------------------------
# Dédiée aux partenaires en mode "agence" (Orange Agence, Wave
# Agence), dont l'Excel est déjà agrégé par agence. Lit directement
# les tables "excel" et "flex" du partenaire (pas de split W2B/B2W
# ici), compare par CODE_AGENCE :
#   DEBIT (Flex)  vs col_debit  (Partenaire)
#   CREDIT (Flex) vs col_credit (Partenaire)
#
# col_credit/col_debit viennent de config.PARTENAIRES[...]
# ["colonnes_agence"] — c'est cette indirection qui permet au même
# moteur générique de servir Orange Agence (montant_cashin/montant_cashout)
# ET Wave Agence (DEPOT/RETRAIT) sans aucune modification du moteur
# ni de ce gateway pour un futur partenaire "agence" supplémentaire :
# il suffira d'ajouter son bloc dans config.py avec sa propre clé
# "colonnes_agence".
#
# Nécessite dans config.PARTENAIRES[...]["tables"] une clé
# "reconciliation_agence" (nom de table SQLite dédié au résultat).
# Bloque explicitement l'appel pour un partenaire en mode
# "two_pointers" (Wave, Wizz), avec un message qui renvoie vers
# /reconciliation/run.
# ============================================================

@app.post("/reconciliation/run-agence")
def run_reconciliation_agence(partenaire: str = Query(...)):

    cfg = cfg_or_400(partenaire)

    verifier_mode(
        partenaire, cfg, MODE_AGENCE,
        "Ce partenaire est réconcilié transaction par transaction : "
        "utilise POST /reconciliation/run à la place."
    )

    t = cfg["tables"]

    if "reconciliation_agence" not in t:
        raise HTTPException(
            status_code=400,
            detail=f"config.py : ajoute la clé 'reconciliation_agence' dans "
                    f"PARTENAIRES['{partenaire}']['tables'] (nom de table SQLite "
                    f"pour stocker ce résultat)."
        )

    wp_raw = read_table(t["excel"])
    wf_raw = read_table(t["flex"])

    if wp_raw.empty or wf_raw.empty:
        cote_vide = []
        if wp_raw.empty:
            cote_vide.append(f"Excel-partenaire ('{t['excel']}')")
        if wf_raw.empty:
            cote_vide.append(f"Flex ('{t['flex']}')")
        raise HTTPException(
            status_code=400,
            detail=f"Table(s) source vide(s) pour {partenaire} : {', '.join(cote_vide)}. "
                    f"Lance le chargement Excel + Flex avant de réconcilier par agence "
                    f"(si c'est le côté Flex, vérifie la période demandée : "
                    f"GET {cfg['flex_url']} avec des dates plus larges pour confirmer "
                    f"qu'il existe bien des écritures Oracle sur la période)."
        )

    # Colonnes montant côté Excel-partenaire pour CE partenaire (voir
    # config.PARTENAIRES[...]["colonnes_agence"]) : Orange fournit
    # montant_cashin/montant_cashout, Wave Agence fournit DEPOT/RETRAIT.
    colonnes_agence = cfg.get("colonnes_agence", {})
    col_credit = colonnes_agence.get("col_credit", "MONTANT_CASHIN")
    col_debit = colonnes_agence.get("col_debit", "MONTANT_CASHOUT")

    try:
        resultat_final = reconcilier_par_agence(
            wp_raw, wf_raw,
            col_credit=col_credit,
            col_debit=col_debit,
        )
    except KeyError as e:
        raise HTTPException(
            status_code=400,
            detail=f"Colonne attendue absente pour la réconciliation par agence ({partenaire}) : {e}."
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur pendant la réconciliation par agence ({partenaire}) : {e}")

    ecrire_table(resultat_final, t["reconciliation_agence"], engine, log_prefix="reconc")

    print(f"[reconc.py] Réconciliation par agence {partenaire} terminée : "
          f"{len(resultat_final)} agence(s) -> {t['reconciliation_agence']}")

    statuts_agence = ["Réconcilié", "Ecart montant", "Non comptabilisée", "Comptabilisation isolée"]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resultat_final.to_excel(writer, sheet_name="Resultat_Agence", index=False)

        table_resume = construire_table_resume_agence(resultat_final)
        table_resume.to_excel(writer, sheet_name="Resume_Agence", index=False)

        for statut in statuts_agence:
            subset = resultat_final[resultat_final["STATUT"] == statut]
            sheet_name = str(statut)[:31]
            subset.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Resultat_Reconciliation_Agence_{partenaire}.xlsx"}
    )


@app.get("/db/reconciliation-agence")
def get_reconciliation_agence(partenaire: str = Query(...), limit: int = Query(None), offset: int = Query(0)):
    cfg = cfg_or_400(partenaire)
    t = cfg["tables"]
    if "reconciliation_agence" not in t:
        return Response(content="[]", media_type="application/json")
    return safe_query(f"SELECT * FROM {t['reconciliation_agence']}", limit=limit, offset=offset)


@app.get("/db/reconciliation-agence-resume")
def get_reconciliation_agence_resume(partenaire: str = Query(...)):
    cfg = cfg_or_400(partenaire)
    t = cfg["tables"]
    if "reconciliation_agence" not in t:
        return []

    resultat = read_reconciliation_or_empty(t["reconciliation_agence"])
    if resultat.empty:
        return []

    resume = construire_table_resume_agence(resultat)
    return json.loads(resume.to_json(orient="records"))


@app.get("/db/reconciliation-agence-taux")
def get_reconciliation_agence_taux(partenaire: str = Query(...)):
    """Taux de réussite calculé sur les statuts de la réconciliation par
    agence (Réconcilié / Ecart montant / Non comptabilisée /
    Comptabilisation isolée) — réutilise calculer_taux_reussite, qui
    gère nativement les statuts absents (comptés à 0)."""
    cfg = cfg_or_400(partenaire)
    t = cfg["tables"]
    if "reconciliation_agence" not in t:
        return calculer_taux_reussite(pd.DataFrame())

    resultat = read_reconciliation_or_empty(t["reconciliation_agence"])
    return calculer_taux_reussite(resultat)


# ============================================================
# RESET — PAR PARTENAIRE (n'efface QUE ses propres tables)
# ============================================================

@app.post("/reset")
def reset_application(partenaire: str = Query(...)):

    cfg = cfg_or_400(partenaire)
    tables = list(cfg["tables"].values())
    supprimees = []

    try:
        with engine.begin() as conn:
            for table in tables:
                try:
                    conn.execute(f"DROP TABLE IF EXISTS {table}")
                    supprimees.append(table)
                except Exception as e:
                    print(f"[reset] Impossible de supprimer {table}: {e}")

        return {
            "status": "success",
            "partenaire": partenaire,
            "message": f"Données {cfg['label']} réinitialisées.",
            "tables_supprimees": supprimees
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur reset : {e}")


# ============================================================
# HEALTHCHECK
# ============================================================

@app.get("/")
@app.get("/health")
def health():
    return {
        "status": "ok",
        "db_path": DB_PATH,
        "partenaires": list(PARTENAIRES.keys()),
        "services_montes": _MOUNTED,
        "mode": "unified",
    }