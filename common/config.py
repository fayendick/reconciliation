# ============================================================
# CONFIG CENTRALE — MULTI-PARTENAIRE (v2 : un app_<partenaire>.py par partenaire)
# ------------------------------------------------------------
# Architecture :
#   - Une API unique (reconc.py, port 8002) monte chaque
#     app_<partenaire>.py et <partenaire>_flex_api.py sous /svc/...
#   - Module_FED / Streamlit n'appellent que :8002.
#   - Chaque fichier partenaire garde son mapping Excel / SQL Oracle.
#
# Pour AJOUTER un partenaire (ex: FREE) :
#   1. Copier app_wizz.py -> app_free.py, adapter le mapping.
#   2. Copier wizz_flex_api.py -> free_flex_api.py, adapter le SQL.
#   3. Ajouter le bloc dans PARTENAIRES + entrées dans mount_partners.py.
#   Rien d'autre à toucher côté FED / Streamlit.
#
# ------------------------------------------------------------
# CLÉ "mode"
# ------------------------------------------------------------
# Deux modes de réconciliation existent dans reconciliation_engine.py,
# et ils ne consomment PAS le même schéma de tables :
#
#   - "two_pointers" : partenaire dont l'Excel liste chaque
#     transaction individuellement (ex: Wave, Wizz). Nécessite
#     les 4 tables détaillées : excel_w2b, excel_b2w, flex_w2b,
#     flex_b2w. Utilise /reconciliation/run.
#
#   - "agence" : partenaire dont l'Excel est déjà une compilation
#     agrégée par agence (ex: Orange Agence, Wave Agence, RIA
#     Agence). PAS de split W2B/B2W côté Excel : les clés
#     excel_w2b/excel_b2w n'existent pas et ne doivent JAMAIS être
#     interrogées pour ce partenaire. Utilise
#     /reconciliation/run-agence, et les tables "excel" + "flex"
#     (compilations complètes) uniquement.
#
# C'est cette clé qui permet à reconc.py de bloquer proprement le
# mauvais endpoint pour le mauvais partenaire, et à streamlit_app.py
# de n'afficher/appeler QUE ce qui existe réellement pour le
# partenaire actif.
#
# ------------------------------------------------------------
# CLÉ "colonnes_agence" (mode "agence" uniquement)
# ------------------------------------------------------------
# reconciliation_engine.reconcilier_par_agence() compare toujours,
# par CODE_AGENCE : un montant "crédit" et un montant "débit" côté
# Partenaire, face à CREDIT/DEBIT côté Flex Oracle. Le NOM de ces
# deux colonnes côté Partenaire change d'un partenaire à l'autre
# (Orange Agence : montant_cashin / montant_cashout ; Wave Agence : DEPOT
# / RETRAIT ; RIA Agence : Montant du paiement / Montant a payer)
# -> chaque partenaire "agence" précise ici où les trouver, sans que
# le moteur générique ait besoin de les connaître à l'avance :
#   - "col_credit" : nom de la colonne Excel-partenaire à comparer
#     à CREDIT côté Flex (ex: un dépôt/cashin qui alimente le compte)
#   - "col_debit"  : nom de la colonne Excel-partenaire à comparer
#     à DEBIT côté Flex (ex: un retrait/cashout qui débite le compte)
#
# ------------------------------------------------------------
# CLÉ "apparier_par_telephone_montant" (mode "two_pointers" uniquement)
# ------------------------------------------------------------
# [NOUVEAU] Par défaut absente (donc False via cfg.get(...)) pour
# TOUS les partenaires "two_pointers" existants (Wave, Wizz) :
# reconciliation_engine.reconciliation_two_pointers() apparie alors
# les lignes par CODE_TRANSACTION identique des deux côtés, comme
# avant — AUCUN changement de comportement pour eux.
#
# Positionnée à True UNIQUEMENT pour ORANGE_USSD, car pour ce
# partenaire CODE_TRANSACTION désigne deux identifiants différents
# d'un côté à l'autre (référence opérateur télécom côté Excel-
# partenaire vs référence interne Oracle TRN_REF_NO côté Flex), qui
# ne coïncident jamais. L'appariement se fait alors par
# NUMERO_COMPTE (téléphone client) normalisé + TYPE TRANSACTION,
# avec préférence pour un montant identique puis pour l'écart de
# temps le plus faible (voir reconciliation_engine.py et
# reconc.py -> run_reconciliation).
# ------------------------------------------------------------

import os

# Racine du projet (parent de common/)
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_CONFIG_DIR = _PROJECT_ROOT


def _load_dotenv(path: str | None = None) -> None:
    """Charge un fichier .env sans écraser les variables déjà présentes
    dans l'environnement du process (comportement standard : une vraie
    variable d'env, définie AVANT le lancement du process, a toujours
    priorité sur le fichier .env).

    [CORRECTIF 2026-08-23] Ce comportement "silencieux" est une source
    de bugs difficiles à diagnostiquer : si une variable (typiquement
    RECON_DB_PATH) a été définie un jour comme variable d'environnement
    Windows PERSISTANTE (setx, Paramètres système, profil PowerShell...),
    alors TOUTE modification du fichier .env est ignorée sans le moindre
    avertissement, et rien dans les logs ne l'indique. C'est exactement
    ce qui s'est produit ici : RECON_DB_PATH pointait vers un ancien
    chemin cassé (C:\\reconciliation\\data\\base.d), et corriger le .env
    n'avait aucun effet tant que cette variable système restait définie.

    -> On journalise maintenant, pour CHAQUE clé du .env, si elle a été
       ignorée parce qu'une variable d'environnement du même nom existait
       déjà — avec les deux valeurs (celle qui gagne vs celle du .env) —
       pour rendre ce cas immédiatement visible dans les logs au
       démarrage, au lieu de devoir le déduire indirectement."""
    env_path = path or os.path.join(_PROJECT_ROOT, ".env")
    if not os.path.isfile(env_path):
        return
    try:
        with open(env_path, encoding="utf-8") as fh:
            for raw in fh:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                if not key:
                    continue
                val = val.strip()
                if len(val) >= 2 and val[0] == val[-1] and val[0] in ("'", '"'):
                    val = val[1:-1]
                if key in os.environ:
                    if os.environ[key] != val:
                        print(
                            f"[config._load_dotenv] ATTENTION : '{key}' est déjà défini "
                            f"comme variable d'environnement système (valeur actuelle "
                            f"utilisée : '{os.environ[key]}') — la valeur du .env "
                            f"('{val}') est IGNORÉE. Si ce n'est pas voulu, supprime "
                            f"cette variable d'environnement système "
                            f"([System.Environment]::SetEnvironmentVariable('{key}', "
                            f"$null, 'User')) puis rouvre un nouveau terminal."
                        )
                    continue
                os.environ[key] = val
    except OSError:
        pass


_load_dotenv()

# ------------------------------------------------------------
# Chemin de la base SQLite partagée par TOUS les services.
# Doit être identique (même variable d'env RECON_DB_PATH) pour
# chaque app_<partenaire>.py, chaque <partenaire>_flex_api.py
# et reconc.py.
# ------------------------------------------------------------
DB_PATH = os.getenv(
    "RECON_DB_PATH",
    os.path.join(_PROJECT_ROOT, "data", "base.db")
)

# [CORRECTIF 2026-08-23] Log explicite de la SOURCE de DB_PATH (variable
# d'env déjà présente au démarrage du process vs valeur par défaut
# calculée depuis _PROJECT_ROOT), pour diagnostiquer en un coup d'œil
# tout chemin inattendu au démarrage, sans avoir à fouiller .env ou les
# variables système à la main.
if "RECON_DB_PATH" in os.environ:
    print(f"[config.py] RECON_DB_PATH défini via l'environnement -> DB_PATH = {DB_PATH}")
else:
    print(f"[config.py] RECON_DB_PATH absent -> valeur par défaut -> DB_PATH = {DB_PATH}")

# [CORRECTIF 2026-08-23] Crée le dossier parent de la base SQLite s'il
# n'existe pas encore. SQLite ne crée JAMAIS le dossier parent tout
# seul (seulement le fichier .db) : sans ce garde-fou, un DB_PATH
# pointant vers un dossier absent provoque
# "sqlite3.OperationalError: unable to open database file" — comme
# rencontré lorsqu'un ancien RECON_DB_PATH (variable d'environnement
# système périmée) pointait vers un dossier jamais créé sur cette
# machine.
_db_dir = os.path.dirname(DB_PATH)
if _db_dir and not os.path.isdir(_db_dir):
    try:
        os.makedirs(_db_dir, exist_ok=True)
        print(f"[config.py] Dossier créé pour la base SQLite : {_db_dir}")
    except OSError as e:
        print(f"[config.py] ATTENTION : impossible de créer le dossier '{_db_dir}' ({e}).")

# Écriture SQLite par lots (to_sql) — utile sur gros extractions Flex.
SQLITE_CHUNKSIZE = int(os.getenv("SQLITE_CHUNKSIZE", "2000"))
# Limite optionnelle des endpoints /db/* (0 = pas de limite, comportement historique).
SQLITE_DB_READ_LIMIT = int(os.getenv("SQLITE_DB_READ_LIMIT", "0"))

# ------------------------------------------------------------
# Oracle Flexcube — credentials UNIQUEMENT via env / .env
# (jamais de mot de passe en dur dans le code).
# Timeouts calibrés sous le REQUEST_TIMEOUT gateway (60s) :
#   connect < pool wait < call < gateway HTTP.
# ------------------------------------------------------------
ORACLE_USER = os.getenv("ORACLE_USER", "report_sn")
ORACLE_PASSWORD = os.getenv("ORACLE_PASSWORD", "")
ORACLE_HOST = os.getenv("ORACLE_HOST", "10.44.221.104")
ORACLE_PORT = os.getenv("ORACLE_PORT", "1522")
ORACLE_SERVICE_NAME = os.getenv("ORACLE_SERVICE_NAME", "FCPRDSNPDB")

ORACLE_POOL_SIZE = int(os.getenv("ORACLE_POOL_SIZE", "5"))
ORACLE_MAX_OVERFLOW = int(os.getenv("ORACLE_MAX_OVERFLOW", "10"))
ORACLE_POOL_RECYCLE = int(os.getenv("ORACLE_POOL_RECYCLE", "1800"))
# Attente max d'une connexion libre dans le pool (secondes).
ORACLE_POOL_TIMEOUT = int(os.getenv("ORACLE_POOL_TIMEOUT", "10"))
# Timeout TCP à l'établissement de session (secondes).
ORACLE_CONNECT_TIMEOUT = int(os.getenv("ORACLE_CONNECT_TIMEOUT", "10"))
# Timeout d'exécution d'une requête Oracle (millisecondes, call_timeout).
ORACLE_CALL_TIMEOUT_MS = int(os.getenv("ORACLE_CALL_TIMEOUT_MS", "50000"))


def get_oracle_url() -> str:
    from urllib.parse import quote_plus

    if not ORACLE_PASSWORD:
        raise RuntimeError(
            "ORACLE_PASSWORD manquant. Définir la variable d'environnement "
            "ou créer un fichier .env (voir .env.example)."
        )

    user = quote_plus(ORACLE_USER)
    password = quote_plus(ORACLE_PASSWORD)
    return (
        f"oracle+cx_oracle://{user}:{password}"
        f"@{ORACLE_HOST}:{ORACLE_PORT}/?service_name={ORACLE_SERVICE_NAME}"
    )


_oracle_engine = None
_sqlite_engine = None


def make_oracle_engine():
    """Engine Oracle singleton (1 pool / process, même si N modules Flex)."""
    global _oracle_engine
    if _oracle_engine is not None:
        return _oracle_engine

    from sqlalchemy import create_engine, event

    engine = create_engine(
        get_oracle_url(),
        pool_pre_ping=True,
        pool_size=ORACLE_POOL_SIZE,
        max_overflow=ORACLE_MAX_OVERFLOW,
        pool_recycle=ORACLE_POOL_RECYCLE,
        pool_timeout=ORACLE_POOL_TIMEOUT,

        #connect_args={
            # cx_Oracle / oracledb : délai max pour ouvrir la session TCP.
           # "tcp_connect_timeout": ORACLE_CONNECT_TIMEOUT,
         # }

        connect_args={},




        # Hint SQLAlchemy (utile surtout si un dialcte le propage) ;
        # le vrai garde-fou requête est call_timeout ci-dessous.
        execution_options={"timeout": max(1, ORACLE_CALL_TIMEOUT_MS // 1000)},
    )

    @event.listens_for(engine, "connect")
    def _set_oracle_call_timeout(dbapi_connection, connection_record):
        # Empêche une requête Flex de bloquer le worker (et le gateway) indéfiniment.
        try:
            dbapi_connection.call_timeout = ORACLE_CALL_TIMEOUT_MS
        except AttributeError:
            pass

    _oracle_engine = engine
    return engine


def make_sqlite_engine():
    """Engine SQLite singleton pointant sur DB_PATH."""
    global _sqlite_engine
    if _sqlite_engine is not None:
        return _sqlite_engine

    from sqlalchemy import create_engine

    _sqlite_engine = create_engine(
        f"sqlite:///{DB_PATH}",
        connect_args={"check_same_thread": False},
        pool_pre_ping=True,
    )
    return _sqlite_engine


# Base URL unique de l'API unifiée (Module_FED / Streamlit / /charger).
GATEWAY_BASE_URL = os.getenv("GATEWAY_BASE_URL", "http://127.0.0.1:8002").rstrip("/")

# ------------------------------------------------------------
# Chemin du référentiel de mapping des agences Wave (utilisé par
# wave_app_agence.py pour le fuzzy matching Agent -> CODE_AGENCE).
# ------------------------------------------------------------
MAPPING_AGENCES_WAVE_PATH = os.getenv(
    "MAPPING_AGENCES_WAVE_PATH",
    os.path.join(_CONFIG_DIR, "data", "mappings", "mapping_agences_cofina_wave.xlsx")
)

# ------------------------------------------------------------
# Chemin du référentiel de mapping des agences RIA (utilisé par
# ria_agence_app.py : match exact sur "Code d'agence" <-> CODE_COFINA,
# puis repli en fuzzy matching sur le nom de la Succursale).
# <-- AJOUT (manquait, provoquait un ImportError au démarrage de
#     ria_agence_app.py -> service jamais up -> 500 depuis reconc.py)
# ------------------------------------------------------------
MAPPING_AGENCES_RIA_PATH = os.getenv(
    "MAPPING_AGENCES_RIA_PATH",
    os.path.join(_CONFIG_DIR, "data", "mappings", "mapping_agences_code_ria.xlsx")
)

MAPPING_ORANGE_MSISDN_PATH = os.getenv(
    "MAPPING_ORANGE_MSISDN_PATH",
    os.path.join(_CONFIG_DIR, "data", "mappings", "Mapp_tab.xlsx")
)

# ------------------------------------------------------------
# Valeurs valides pour la clé "mode" de chaque partenaire.
# ------------------------------------------------------------
MODE_TWO_POINTERS = "two_pointers"
MODE_AGENCE = "agence"
MODES_VALIDES = (MODE_TWO_POINTERS, MODE_AGENCE)

# ------------------------------------------------------------
# Schéma interne standard : c'est le format que produit CHAQUE
# app_<partenaire>.py une fois ses colonnes brutes renommées, et
# que consomme reconciliation_engine.py. Garder ce schéma commun
# est ce qui permet de réutiliser le même moteur de
# réconciliation pour tous les partenaires en mode "two_pointers".
# ------------------------------------------------------------
COLONNES_STANDARD_EXCEL = [
    "DATE TRANSACTION",
    "TYPE TRANSACTION",
    "CODE TRANSACTION OPERATEUR",
    "NUMERO COMPTE",
    "MONTANT",
]

# ------------------------------------------------------------
# COLONNES SUPPLÉMENTAIRES POUR LA TABLE RÉSUMÉ (mode "two_pointers"
# uniquement — voir reconciliation_engine.construire_table_resume)
# ------------------------------------------------------------
COLONNES_RESUME_DEFAUT = {
    "num_tel_client": None,
    "nom_client": None,
    "agence": "WF_LIBELLE_AGENCE",
    "periode_fichier": "WF_DATE_VALEUR",
}

# ------------------------------------------------------------
# WAVE (mode "two_pointers") : colonnes réelles confirmées.
# ------------------------------------------------------------
COLONNES_RESUME_WAVE = {
    "num_tel_client": "WP_NUMERO WALLET",
    "nom_client": ["WP_CLIENT", "WF_DESCRIPTION"],
    "agence": "WF_LIBELLE_AGENCE",
    "periode_fichier": "WF_DATE_VALEUR",
}

# ------------------------------------------------------------
# ORANGE USSD (mode two_pointers)
# ------------------------------------------------------------

COLONNES_RESUME_ORANGE_USSD = {

    # numéro téléphone client provenant du fichier partenaire
    "num_tel_client": "WP_NUMERO COMPTE",

    "nom_client": None,

    # agence provenant du Flex Oracle
    "agence": "WF_LIBELLE_AGENCE",

    # période Oracle
    "periode_fichier": "WF_DATE_VALEUR",

}





#---------- resume orange agence (mode agence -> non utilisé par
# construire_table_resume, conservé pour compat si un jour Orange
# Agence repasse en two_pointers ; sans effet en mode "agence" actuel)

COLONNES_RESUME_ORANGE = {

    "num_tel_client": "WP_msisdn",

    "nom_client": None,

    "agence": [
        "WP_AGENCE",
        "WF_LIBELLE_AGENCE"
    ],

    "periode_fichier": "WF_DATE_VALEUR",

}

# ------------------------------------------------------------
# COLONNES POUR LA RÉCONCILIATION PAR AGENCE (mode "agence"
# uniquement — voir reconciliation_engine.reconcilier_par_agence).
# ------------------------------------------------------------

COLONNES_AGENCE_ORANGE = {
    "col_credit": "MONTANT_CASHIN",
    "col_debit": "MONTANT_CASHOUT",
}

# Wave Agence : le fichier Excel "Wave agence" (traité par
# wave_app_agence.py) contient directement les colonnes DEPOT et
# RETRAIT par agence (en plus de CODE_AGENCE/NOM_AGENCE ajoutées
# par le fuzzy matching) -> DEPOT est comparé à CREDIT côté Flex,
# RETRAIT est comparé à DEBIT côté Flex.
COLONNES_AGENCE_WAVE_AGENCE = {
    "col_credit": "DEPOT",
    "col_debit": "RETRAIT",
}

# RIA Agence : ria_agence_app.py produit "Montant du paiement" et
# "Montant a payer" (déjà en valeur absolue) par agence -> hypothèse
# par défaut, à valider avec la définition métier réelle :
#   Montant du paiement (Partenaire) <-> CREDIT (Flex)
#   Montant a payer      (Partenaire) <-> DEBIT  (Flex)
# <-- AJOUT (manquait, provoquait un ValueError dans get_partenaire
#     appelé par RIA_agence_flex.py au démarrage du module)
#COLONNES_AGENCE_RIA_AGENCE = {
   # "col_credit": "Montant du paiement",
   # "col_debit": "Montant a payer",
#}





COLONNES_AGENCE_RIA_AGENCE = {
    "col_credit": "Montant A Envoyé",   # <-> CREDIT (Flex)
    "col_debit": "Montant A Payé",      # <-> DEBIT (Flex)
}



PARTENAIRES = {

    "WAVE": {
        "label": "Wave",
        "mode": MODE_TWO_POINTERS,
        "upload_url": os.getenv("WAVE_UPLOAD_URL", f"{GATEWAY_BASE_URL}/svc/wave-excel/process-excel"),
        "flex_url": os.getenv("WAVE_FLEX_URL", f"{GATEWAY_BASE_URL}/svc/wave-flex/wave-inter-flex"),
        "tables": {
            "excel": "COMPILATION_WAVE",
            "excel_w2b": "COMPILATION_WAVE_W2B",
            "excel_b2w": "COMPILATION_WAVE_B2W",
            "flex": "WAVE_FLEX",
            "flex_w2b": "WAVE_FLEX_W2B",
            "flex_b2w": "WAVE_FLEX_B2W",
            "reconciliation": "RECONCILIATION_WAVE",
            "doublons": "DOUBLONS_WAVE",
        },
        "colonnes_resume": {**COLONNES_RESUME_WAVE},
        # Pas de "apparier_par_telephone_montant" ici -> False par
        # défaut (voir cfg.get(...) dans reconc.py) -> comportement
        # INCHANGÉ pour Wave (appariement par CODE_TRANSACTION, comme
        # avant ce correctif).
    },

    #====================================USSSD TWO POINTERS



    "ORANGE_AGENCE": {
        "label": "Orange Agence",
        # --------------------------------------------------------
        # MODE "agence" : app_orange.py n'écrit QUE la table "excel"
        # (compilation complète, déjà agrégée par agence via
        # Mapp_tab.xlsx). Il n'y a AUCUNE table excel_w2b / excel_b2w
        # pour Orange Agence -> ne jamais les interroger. Le
        # rapprochement se fait via reconcilier_par_agence() /
        # /reconciliation/run-agence, pas via /reconciliation/run
        # (Two Pointers, réservé au mode "two_pointers").
        # --------------------------------------------------------
        "mode": MODE_AGENCE,
        "upload_url": os.getenv("ORANGE_AGENCE_UPLOAD_URL", f"{GATEWAY_BASE_URL}/svc/orange-excel/process-excel"),
        "flex_url": os.getenv("ORANGE_AGENCE_FLEX_URL", f"{GATEWAY_BASE_URL}/svc/orange-flex/orange-flex"),
        "tables": {
            "excel": "COMPILATION_ORANGE_AGENCE",
            # NOTE : ces deux clés ne correspondent à AUCUNE table
            # réellement créée (app_orange.py ne les écrit jamais).
            # Conservées uniquement pour compat si un jour Orange
            # Agence bascule en mode "two_pointers" ; reconc.py et
            # streamlit_app.py ne les interrogent plus tant que
            # mode == "agence".
            "excel_w2b": "COMPILATION_ORANGE_AGENCE_W2B",
            "excel_b2w": "COMPILATION_ORANGE_AGENCE_B2W",
            "flex": "ORANGE_AGENCE_FLEX",
            "flex_w2b": "ORANGE_AGENCE_FLEX_W2B",   # écritures DEBIT (existe réellement, voir orange_flex_api.py)
            "flex_b2w": "ORANGE_AGENCE_FLEX_B2W",   # écritures CREDIT (existe réellement, voir orange_flex_api.py)
            # NOTE : renommée "..._DETAIL" pour éviter toute collision
            # avec "reconciliation_agence" ci-dessous (les deux tables
            # sont différentes : celle-ci resterait vide tant qu'Orange
            # Agence n'est pas en mode "two_pointers").
            "reconciliation": "RECONCILIATION_ORANGE_AGENCE_DETAIL",
            "doublons": "DOUBLONS_ORANGE_AGENCE",
            # ------------------------------------------------------
            # Table dédiée au résultat de reconcilier_par_agence()
            # (section 11 de reconciliation_engine.py), calculé à
            # partir des colonnes montant_cashin/montant_cashout +
            # CODE_AGENCE produites par app_orange.py (Excel-partenaire
            # mappé) et DEBIT/CREDIT + CODE_AGENCE produites par
            # orange_flex_api.py (Flex Oracle).
            # ------------------------------------------------------
            "reconciliation_agence": "RECONCILIATION_ORANGE_AGENCE",
            "resume": "RESUME_ORANGE_AGENCE",
        },

        "colonnes_resume": {**COLONNES_RESUME_ORANGE},
        "colonnes_agence": {**COLONNES_AGENCE_ORANGE},
    },

    "WIZZ": {
        "label": "Wizz",
        "mode": MODE_TWO_POINTERS,
        "upload_url": os.getenv("WIZZ_UPLOAD_URL", f"{GATEWAY_BASE_URL}/svc/wizz-excel/process-excel"),
        "flex_url": os.getenv("WIZZ_FLEX_URL", f"{GATEWAY_BASE_URL}/svc/wizz-flex/wizz-flex"),
        "tables": {
            "excel": "COMPILATION_WIZZ",
            "excel_w2b": "COMPILATION_WIZZ_W2B",
            "excel_b2w": "COMPILATION_WIZZ_B2W",
            "flex": "WIZZ_FLEX",
            "flex_w2b": "WIZZ_FLEX_W2B",
            "flex_b2w": "WIZZ_FLEX_B2W",
            "reconciliation": "RECONCILIATION_WIZZ",
            "doublons": "DOUBLONS_WIZZ",
        },

        "colonnes_resume": {**COLONNES_RESUME_DEFAUT},
        # Pas de "apparier_par_telephone_montant" ici -> False par
        # défaut -> comportement INCHANGÉ pour Wizz.

    },

    "WAVE_AGENCE": {
        "label": "Wave Agence",
        # --------------------------------------------------------
        # MODE "agence" : wave_app_agence.py écrit une compilation
        # complète par agence (CODE_AGENCE, NOM_AGENCE, DEPOT,
        # RETRAIT, ...), pas de transaction individuelle -> pas de
        # split W2B/B2W, comme Orange. Le rapprochement se fait par
        # SOMME et par CODE_AGENCE via reconcilier_par_agence() /
        # /reconciliation/run-agence :
        #   DEPOT  (Partenaire) <-> CREDIT (Flex)
        #   RETRAIT(Partenaire) <-> DEBIT  (Flex)
        # --------------------------------------------------------
        "mode": MODE_AGENCE,
        "upload_url": os.getenv("WAVE_AGENCE_UPLOAD_URL", f"{GATEWAY_BASE_URL}/svc/wave-agence-excel/map-agences"),
        "flex_url": os.getenv("WAVE_AGENCE_FLEX_URL", f"{GATEWAY_BASE_URL}/svc/wave-agence-flex/wave-agence-flex"),
        "tables": {
            "excel": "COMPILATION_WAVE_AGENCE",
            "flex": "WAVE_AGENCE_FLEX",
            "reconciliation_agence": "RECONCILIATION_WAVE_AGENCE",
            "resume": "RESUME_WAVE_AGENCE",
        },
        "colonnes_resume": {},   # non utilisé en mode "agence"
        "colonnes_agence": {**COLONNES_AGENCE_WAVE_AGENCE},
    },

    # <-- AJOUT : bloc manquant. ria_agence_app.py (port 8032) et
    # RIA_agence_flex.py (port 8033) lisent tous les deux
    # get_partenaire("RIA_AGENCE") dès leur import -> sans ce bloc,
    # ces deux services levaient une exception au démarrage et
    # restaient injoignables, d'où le 500 "après le chargement"
    # renvoyé par reconc.py (upload_url/flex_url injoignables).
    "RIA_AGENCE": {
        "label": "RIA Agence",
        # --------------------------------------------------------
        # MODE "agence" : ria_agence_app.py écrit une compilation
        # complète (une ligne par paiement RIA, déjà mappée à une
        # agence via mapping_agences_code_ria.xlsx), pas de split
        # W2B/B2W. Le rapprochement se fait par SOMME et par
        # CODE_AGENCE via reconcilier_par_agence() /
        # /reconciliation/run-agence :
        #   Montant du paiement (Partenaire) <-> CREDIT (Flex)
        #   Montant a payer      (Partenaire) <-> DEBIT  (Flex)
        # --------------------------------------------------------
        "mode": MODE_AGENCE,
        "upload_url": os.getenv("RIA_AGENCE_UPLOAD_URL", f"{GATEWAY_BASE_URL}/svc/ria-agence-excel/map-agences-ria"),
        "flex_url": os.getenv("RIA_AGENCE_FLEX_URL", f"{GATEWAY_BASE_URL}/svc/ria-agence-flex/ria-agence-flex"),
        "tables": {
            "excel": "COMPILATION_RIA_AGENCE",
            "flex": "RIA_AGENCE_FLEX",
            "reconciliation_agence": "RECONCILIATION_RIA_AGENCE",
            "resume": "RESUME_RIA_AGENCE",
        },
        "colonnes_resume": {},   # non utilisé en mode "agence"
        "colonnes_agence": {**COLONNES_AGENCE_RIA_AGENCE},
    },
}




# ============================================================
# PARTENAIRE ORANGE_USSD
# ------------------------------------------------------------
# Mode "two_pointers" (comme Wave / Wizz) : le fichier USSD
# Partenaire liste chaque transaction individuellement (pas un
# agrégat par agence) -> on réutilise tel quel le moteur
# reconciliation_engine.reconcilier_un_sens / two_pointers, avec
# les 4 tables détaillées (excel_w2b, excel_b2w, flex_w2b,
# flex_b2w).
#
# Mapping vers le schéma standard (COLONNES_STANDARD_EXCEL) fait
# par app_orange_ussd.py :
#   DATE TRANSACTION           = Date + Heure (partenaire) concaténées
#   TYPE TRANSACTION           = Service ("Cash in" -> "W2B",
#                                  "Cash Out" -> "B2W" — codes EXACTS
#                                  attendus par excel_common.separer_w2b_b2w())
#   CODE TRANSACTION OPERATEUR = Référence
#   NUMERO COMPTE               = N° de Compte (Correspondant) = n° de
#                                  téléphone du client (PAS le compte
#                                  Agent 786256338)
#   MONTANT                     = Débit si renseigné, sinon Crédit
#
# Côté Flex Oracle (orange_ussd_flex_api.py), sens :
#   W2B (Wallet -> Banque = argent qui RENTRE, Cash In)  -> compare
#        MONTANT (Débit partenaire) à MOUVEMENT_CREDIT (Crédit Flex)
#   B2W (Banque -> Wallet = argent qui SORT, Cash Out)   -> compare
#        MONTANT (Crédit partenaire) à MOUVEMENT_DEBIT (Débit Flex)
# Convention identique à Wave : à réajuster si l'analyse métier
# montre l'inverse, auquel cas il suffit d'inverser le mapping dans
# _normaliser_type_transaction() (app_orange_ussd.py) ET le
# DECODE(drcr_ind, ...) (orange_ussd_flex_api.py).
#
# NOTE sur "colonnes_resume" : le préfixage WP_/WF_ fait par
# reconciliation_engine.py concatène "WP_"/"WF_" + nom EXACT de la
# colonne d'origine (espaces compris, pas de conversion en underscore).
# Comme COLONNES_STANDARD_EXCEL utilise "NUMERO COMPTE" (avec un
# espace), la colonne résultante est "WP_NUMERO COMPTE" (espace) et
# NON "WP_NUMERO_COMPTE" (underscore) — piège classique à ne pas
# reproduire pour un futur partenaire two_pointers.
#
# NOTE sur "apparier_par_telephone_montant" (voir bloc en tête de
# fichier) : positionnée à True ICI UNIQUEMENT, car CODE_TRANSACTION
# n'est PAS la même référence des deux côtés pour ce partenaire
# (référence opérateur télécom côté Excel vs TRN_REF_NO Oracle côté
# Flex) -> l'appariement bascule sur NUMERO_COMPTE (téléphone)
# normalisé + montant, dans reconciliation_engine.py.
# ------------------------------------------------------------
# Montés sous reconc (:8002) :
#   /svc/orange-ussd-excel/process-excel
#   /svc/orange-ussd-flex/orange-ussd-flex
# ============================================================

PARTENAIRES["ORANGE_USSD"] = {
    "label": "Orange USSD Partenaire",
    "mode": MODE_TWO_POINTERS,
    "upload_url": os.getenv("ORANGE_USSD_UPLOAD_URL", f"{GATEWAY_BASE_URL}/svc/orange-ussd-excel/process-excel"),
    "flex_url": os.getenv("ORANGE_USSD_FLEX_URL", f"{GATEWAY_BASE_URL}/svc/orange-ussd-flex/orange-ussd-flex"),
    "tables": {
        "excel": "COMPILATION_ORANGE_USSD",
        "excel_w2b": "COMPILATION_ORANGE_USSD_W2B",
        "excel_b2w": "COMPILATION_ORANGE_USSD_B2W",
        "flex": "ORANGE_USSD_FLEX",
        "flex_w2b": "ORANGE_USSD_FLEX_W2B",
        "flex_b2w": "ORANGE_USSD_FLEX_B2W",
        "reconciliation": "RECONCILIATION_ORANGE_USSD",
        "doublons": "DOUBLONS_ORANGE_USSD",
    },

    # [CORRECTIF — Orange USSD uniquement] CODE_TRANSACTION (Référence
    # opérateur télécom côté Excel-partenaire) ne correspond à AUCUN
    # identifiant côté Flex (TRN_REF_NO = référence interne Oracle) :
    # comparer ces deux colonnes ne matche jamais rien. On apparie donc
    # par NUMERO_COMPTE (téléphone client) normalisé + TYPE TRANSACTION,
    # avec préférence pour un montant identique puis pour l'écart de
    # temps le plus faible (voir reconciliation_engine.py). Absent chez
    # tous les autres partenaires -> comportement inchangé pour eux.
    "apparier_par_telephone_montant": True,

    "colonnes_resume": {
        "num_tel_client": "WP_NUMERO COMPTE",   # <- espace, pas underscore (voir note ci-dessus)
        "nom_client": None,
        "agence": "WF_LIBELLE_AGENCE",          # <- fourni par le JOIN STTM_BRANCH ajouté dans orange_ussd_flex_api.py
        "periode_fichier": "WF_DATE_VALEUR",
    },
}




def get_partenaire(nom: str) -> dict:
    """Retourne la config d'un partenaire, lève une erreur explicite sinon.

    Alias courants (identifiants Module_FED / typos) → clé gateway.
    """
    nom = (nom or "").upper().strip().replace(" ", "_").replace("-", "_")
    aliases = {
        "WAVE_USSD": "WAVE",          # partenaire mal nommé côté FED (fichier Wave banque)
        "WAVE_INTER": "WAVE",
        "WAVE_INT": "WAVE",
        "ORANGE": "ORANGE_AGENCE",
        "ORANGE_MONEY": "ORANGE_AGENCE",
        "ORANGE_AG": "ORANGE_AGENCE",
        "USSD": "ORANGE_USSD",
        "ORANGE_USSD_PARTENAIRE": "ORANGE_USSD",
        "RIA": "RIA_AGENCE",
    }
    nom = aliases.get(nom, nom)
    if nom not in PARTENAIRES:
        raise ValueError(
            f"Partenaire '{nom}' inconnu. Partenaires disponibles : {list(PARTENAIRES.keys())}"
        )
    return PARTENAIRES[nom]


def get_mode(nom: str) -> str:
    """Retourne le mode ('two_pointers' ou 'agence') d'un partenaire,
    avec repli sur 'two_pointers' si la clé est absente (compat
    ascendante pour un partenaire ajouté sans préciser 'mode')."""
    return get_partenaire(nom).get("mode", MODE_TWO_POINTERS)