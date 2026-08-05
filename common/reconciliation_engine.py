"""
============================================================
 RECONCILIATION ENGINE — recherche du meilleur candidat
============================================================
Moteur 100% générique : ne connaît PAS le partenaire. Il reçoit
des DataFrames déjà chargés (Excel partenaire + Flex Oracle) et
applique toujours la même logique de rapprochement, quel que
soit le partenaire (Wave, Orange Agence, Wave Agence, ...), tant que les
DataFrames respectent les colonnes attendues :
  - Excel partenaire (two_pointers) : DATE TRANSACTION, CODE
    TRANSACTION OPERATEUR, NUMERO COMPTE, MONTANT, TYPE TRANSACTION
  - Flex Oracle (two_pointers)      : DATE_VALEUR, CODE_TRANSACTION_
    OPERATEUR, NUMERO_COMPTE, MOUVEMENT_DEBIT/CREDIT, TYPE_TRANSACTION
  - Excel partenaire (agence)       : CODE_AGENCE + soit une colonne
    "SOLDE_PARTENAIRE" DÉJÀ calculée à la source (ex: Wave Agence,
    qui calcule le solde net à partir du Montant (CFA) signé, toutes
    catégories de "Quoi" confondues), soit une colonne "crédit" et
    une colonne "débit" dont le NOM est précisé par
    config.PARTENAIRES[...]["colonnes_agence"] (ex: Orange Agence ->
    montant_cashin/montant_cashout). Voir preparer_agence_partenaire
    ci-dessous pour le détail des deux cas.
  - Flex Oracle (agence)            : CODE_AGENCE, DEBIT, CREDIT

C'est ce qui permet de réutiliser exactement ce fichier pour
n'importe quel nouveau partenaire, sans aucune modification.

Ce fichier contient DEUX logiques de rapprochement, indépendantes :
  A. Rapprochement TRANSACTION PAR TRANSACTION (Two Pointers),
     pour les partenaires dont l'Excel liste chaque transaction
     individuellement (ex: Wave, Wizz, Orange USSD) — sections 1 à 9
     ci-dessous.
  B. Rapprochement AGRÉGÉ PAR AGENCE (section 11), pour les
     partenaires dont l'Excel est déjà une compilation par agence
     (ex: Orange Agence, Wave Agence) — pas de CODE_TRANSACTION ni
     de DATE_HEURE à apparier, on compare des SOMMES par
     CODE_AGENCE.

------------------------------------------------------------
[CORRECTIF v2] APPARIEMENT ORANGE USSD PAR TELEPHONE + MONTANT
------------------------------------------------------------
Pour Wave/Wizz, CODE_TRANSACTION (Excel-partenaire) et TRN_REF_NO
(Flex, exposé sous CODE_TRANSACTION_OPERATEUR) désignent LA MÊME
référence de transaction des deux côtés : comparer
wp["CODE_TRANSACTION"] == wf["CODE_TRANSACTION"] fonctionne donc
très bien pour ces partenaires, et CE COMPORTEMENT N'EST PAS
TOUCHÉ (voir apparier_par_telephone_montant=False, valeur par
défaut, dans reconciliation_two_pointers / reconcilier_un_sens).

Pour Orange USSD, en revanche :
  - CODE_TRANSACTION côté Excel-partenaire = colonne "Référence"
    du rapport Orange Money = référence de l'OPÉRATEUR TÉLÉCOM.
  - CODE_TRANSACTION côté Flex = TRN_REF_NO = référence INTERNE
    Oracle/Flexcube de l'écriture comptable.
  Ce sont deux identifiants totalement différents qui ne
  coïncident jamais -> comparer ces deux colonnes ne matche RIEN.

  [Fenêtre de recherche] La sélection du "meilleur candidat" scanne
  TOUTES les lignes flex encore disponibles (non utilisées), et non
  une fenêtre glissante calée sur `len(utilises)` (qui confondait
  le NOMBRE de lignes consommées avec leur POSITION réelle dans le
  tableau trié). `fenetre_recherche` reste accepté en paramètre pour
  compatibilité de signature mais n'est plus utilisé pour borner la
  recherche.

------------------------------------------------------------
[CORRECTIF v3] LE TÉLÉPHONE N'EST PLUS UN FILTRE BLOQUANT
------------------------------------------------------------
Constat sur données réelles : côté Flex, NUMERO_COMPTE ne désigne
pas toujours le MSISDN du client — dans certains flux (ex: écritures
au niveau agence), c'est le compte/la référence de l'AGENCE
(WF_CODE_AGENCE / WF_LIBELLE_AGENCE sont d'ailleurs présents à côté).
Exiger l'égalité stricte WP_NUMERO_COMPTE_NORME == WF_NUMERO_COMPTE_
NORME comme condition bloquante (`continue`) excluait alors TOUS
les bons candidats, même quand montant et heure correspondaient
exactement. Le téléphone devient donc un critère de PRÉFÉRENCE
dans le score plutôt qu'un filtre.

------------------------------------------------------------
[CORRECTIF v4] IDENTITÉ MINIMALE REQUISE (anti-vol de candidat)
------------------------------------------------------------
Sans le v3, un candidat pouvait être retenu même sans AUCUN point
commun avec la ligne partenaire — la ligne flex la plus proche en
temps par défaut. Une transaction WP orpheline "volait" ainsi la
ligne flex d'une AUTRE transaction WP qui, elle, avait une vraie
correspondance. Un candidat n'est désormais retenu QUE s'il partage
au moins un signal d'identité : montant identique OU téléphone
identique.

------------------------------------------------------------
[CORRECTIF v5] MONTANT SEUL = SIGNAL FAIBLE AU-DELÀ D'1H
------------------------------------------------------------
Le montant identique seul (sans téléphone identique) reste un
signal statistiquement fragile sur une longue fenêtre : deux
transactions totalement indépendantes partagent souvent un montant
rond (10 000, 20 000, 40 000 CFA...). Observé concrètement : une
transaction WP orpheline (aucune vraie correspondance) s'est vue
apparier avec une transaction flex éloignée de ~2h37, simplement
parce que leurs montants coïncidaient — ce qui a, comme au v4,
"consommé" la ligne flex qui appartenait en réalité à UNE AUTRE
ligne WP arrivée 1 seconde plus tard avec le MÊME téléphone.

Règle ajoutée : un candidat identifié PAR LE MONTANT SEUL (sans
confirmation téléphone) n'est accepté que si l'écart de temps reste
sous 1h. Au-delà d'1h, le téléphone doit AUSSI correspondre pour
que le candidat soit retenu. Le palier de statut "Réconcilié -
écart > 1h" ne peut donc plus apparaître pour un simple hasard de
montant : il suppose désormais que montant ET téléphone coïncident
tous les deux.

  Activé UNIQUEMENT quand apparier_par_telephone_montant=True est
  passé explicitement (voir config.py -> PARTENAIRES["ORANGE_USSD"]
  et reconc.py -> run_reconciliation). Par défaut (Wave, Wizz,
  tout partenaire qui ne positionne pas cette clé), la valeur reste
  False et le comportement est STRICTEMENT IDENTIQUE à avant tous
  ces correctifs.
============================================================
"""

import pandas as pd


# ============================================================
# LISTE FIXE DES STATUTS POSSIBLES (ordre d'affichage commun à
# reconc.py et streamlit_app.py). Toujours les mêmes valeurs,
# même si un statut n'a aucune ligne pour un partenaire donné —
# c'est cette liste qui permet d'afficher "0" plutôt que de ne
# pas afficher la case du tout.
#
# "Réconcilié - écart 8s à 1h" / "Réconcilié - écart > 1h" :
# même code transaction, même type, même montant (diff_montant
# == 0) qu'un "Réconcilié"/"Réconcilié avec tolérance", mais
# l'écart de date dépasse la tolérance stricte (8s) — on les
# distingue plutôt que de les classer en "Non comptabilisée"/
# "Comptabilisation isolée", car TOUT correspond sauf le délai.
#
# NOTE : le rapprochement par agence (section 11) n'utilise que 4
# de ces 8 statuts ("Réconcilié", "Ecart montant", "Non
# comptabilisée", "Comptabilisation isolée") — il n'y a pas de
# notion de délai ni de tolérance temporelle sur un agrégat.
# ============================================================
STATUTS = [
    "Réconcilié",
    "Réconcilié avec tolérance",
    "Réconcilié - écart 8s à 1h",
    "Réconcilié - écart > 1h",
    "Ecart montant",
    "Non comptabilisée",
    "Comptabilisation isolée",
    "Doublon",
]

# ------------------------------------------------------------
# Couleurs hexadécimales alignées sur les émojis utilisés dans
# streamlit_app.py (STATUT_COLORS) :
#   ✅ Réconcilié                    -> vert soutenu
#   🟢 Réconcilié avec tolérance      -> vert clair
#   🟡 Réconcilié - écart 8s à 1h     -> jaune/ambre (alerte légère)
#   🟤 Réconcilié - écart > 1h        -> ambre foncé (alerte forte)
#   🟠 Ecart montant                 -> orange
#   🔵 Non comptabilisée             -> bleu
#   🟣 Comptabilisation isolée       -> violet
#   🔴 Doublon                       -> rouge
# Utilisé par le graphique camembert (/db/reconciliation-graphe-statut)
# pour que les couleurs du graphe correspondent exactement aux
# couleurs des statuts affichés ailleurs dans l'UI.
# ------------------------------------------------------------
STATUT_COULEURS = {
    "Réconcilié": "#1E8449",
    "Réconcilié avec tolérance": "#2ECC71",
    "Réconcilié - écart 8s à 1h": "#F1C40F",
    "Réconcilié - écart > 1h": "#D68910",
    "Ecart montant": "#E67E22",
    "Non comptabilisée": "#3498DB",
    "Comptabilisation isolée": "#9B59B6",
    "Doublon": "#E74C3C",
}


# ============================================================
# 0. NORMALISATION TÉLÉPHONE
# ============================================================

def _normaliser_telephone(serie: pd.Series) -> pd.Series:
    """Rend un numéro de téléphone comparable quel que soit le format
    d'origine (avec/sans indicatif pays, espaces, zéro initial...) :
    ne garde que les chiffres, puis conserve les 9 derniers (format
    local sénégalais), afin qu'un numéro stocké avec l'indicatif
    ('221771234567', Flex/KYC) soit comparable à un numéro stocké
    sans ('771234567' ou '0771234567', fichier Orange partenaire)."""
    s = serie.astype(str).str.replace(r"\D", "", regex=True)
    return s.str.slice(-9)


# ============================================================
# 1. NORMALISATION EXCEL PARTENAIRE (Wave-Partenaire, Orange Agence-Partenaire, ...)
# ============================================================

def preparer_wave_partenaire(df: pd.DataFrame, sens: str = None) -> pd.DataFrame:
    """
    Normalise l'Excel-partenaire vers le schéma interne du moteur
    two_pointers (DATE_HEURE, CODE_TRANSACTION, NUMERO_COMPTE,
    MONTANT_COMPARAISON), quel que soit le schéma source :

      - Ancien schéma (Wave, Wizz, ...) : DATE TRANSACTION + MONTANT
        (une seule colonne montant, déjà orientée à la source).
        -> comportement STRICTEMENT INCHANGÉ.

      - Nouveau schéma (Orange USSD Partenaire) : DATE_HEURE (déjà
        fusionnée Date+Heure par app_orange_ussd.py) + DEBIT/CREDIT
        séparés. `sens` ('W2B' ou 'B2W') est alors OBLIGATOIRE pour
        savoir laquelle des deux colonnes utiliser comme
        MONTANT_COMPARAISON.

        Convention métier (double-entrée, confirmée par le mapping
        documenté dans config.py pour ORANGE_USSD) :
            W2B (Cash In / Wallet -> Banque, argent qui RENTRE)
                -> DEBIT côté partenaire, à comparer à
                   MOUVEMENT_CREDIT côté Flex
            B2W (Cash Out / Banque -> Wallet, argent qui SORT)
                -> CREDIT côté partenaire, à comparer à
                   MOUVEMENT_DEBIT côté Flex
        C'est un rapprochement CROISÉ (débit partenaire <-> crédit
        Flex, et inversement), pas un rapprochement "même nom des
        deux côtés" : un débit chez le partenaire correspond à un
        crédit dans les livres Flex et réciproquement, comme dans
        toute écriture en partie double.
        Voir preparer_wave_flex ci-dessous : le côté Flex, lui,
        était déjà correct (W2B -> MOUVEMENT_CREDIT, B2W ->
        MOUVEMENT_DEBIT) et n'a pas été modifié.

    La détection se fait sur les colonnes présentes (DEBIT/CREDIT ->
    nouveau schéma), donc AUCUN changement de comportement pour un
    partenaire qui n'a pas ces colonnes (Wave, Wizz, ...).
    """
    df = df.copy()

    nouveau_schema = "DEBIT" in df.columns and "CREDIT" in df.columns

    if nouveau_schema:
        df["DATE_HEURE"] = pd.to_datetime(df["DATE_HEURE"], errors="coerce")

        if sens == "W2B":
            # W2B : DEBIT partenaire (<-> MOUVEMENT_CREDIT Flex)
            df["MONTANT_COMPARAISON"] = df["DEBIT"]
        elif sens == "B2W":
            # B2W : CREDIT partenaire (<-> MOUVEMENT_DEBIT Flex)
            df["MONTANT_COMPARAISON"] = df["CREDIT"]
        else:
            raise ValueError(
                "preparer_wave_partenaire : 'sens' ('W2B' ou 'B2W') est "
                "obligatoire pour un Excel-partenaire au format DEBIT/CREDIT "
                "(ex: Orange USSD Partenaire)."
            )

        df.rename(
            columns={
                "CODE TRANSACTION OPERATEUR": "CODE_TRANSACTION",
                "NUMERO COMPTE": "NUMERO_COMPTE",
            },
            inplace=True,
        )
    else:
        # --- comportement historique, inchangé (Wave, Wizz, ...) ---
        df["DATE_HEURE"] = pd.to_datetime(df["DATE TRANSACTION"], errors="coerce")

        df.rename(
            columns={
                "CODE TRANSACTION OPERATEUR": "CODE_TRANSACTION",
                "NUMERO COMPTE": "NUMERO_COMPTE",
                "MONTANT": "MONTANT_COMPARAISON",
            },
            inplace=True,
        )

    for col in ["CODE_TRANSACTION", "NUMERO_COMPTE"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Colonne téléphone normalisée : utilisée comme critère de
    # PRÉFÉRENCE dans l'appariement Orange USSD (voir section 7,
    # CORRECTIF v3/v5) ET conservée pour affichage/audit (ex:
    # Num_Tel_Client dans la table résumé). Ajout pur pour
    # Wave/Wizz : aucun impact, ils ne la lisent jamais dans leur
    # chemin d'exécution.
    if "NUMERO_COMPTE" in df.columns:
        df["NUMERO_COMPTE_NORME"] = _normaliser_telephone(df["NUMERO_COMPTE"])

    return df


# ============================================================
# 2. NORMALISATION FLEX ORACLE (Wave Flex, Orange Agence Flex, ...)
# ------------------------------------------------------------
# [INCHANGÉ] Cette fonction était déjà correcte et n'a pas été
# modifiée : W2B -> MOUVEMENT_CREDIT, B2W -> MOUVEMENT_DEBIT.
# Seul ajout : la colonne NUMERO_COMPTE_NORME (voir section 0),
# lue comme critère de PRÉFÉRENCE (pas de filtrage) dans le mode
# apparier_par_telephone_montant (Orange USSD, voir section 7), en
# plus de servir à l'affichage/audit — sans impact sur le reste.
# NOTE : pour certains flux Flex (écritures niveau agence),
# NUMERO_COMPTE peut être un compte d'agence et non le MSISDN du
# client — voir CORRECTIF v3/v4/v5 en tête de fichier.
# ============================================================

def preparer_wave_flex(df: pd.DataFrame, sens: str) -> pd.DataFrame:

    df = df.copy()

    if "CODE_TRANSACTION" in df.columns:
        df.rename(columns={"CODE_TRANSACTION": "CODE_TRANSACTION_FLEX"}, inplace=True)

    df["DATE_HEURE"] = pd.to_datetime(df["DATE_VALEUR"], errors="coerce")

    if sens == "W2B":
        df["MONTANT_COMPARAISON"] = df["MOUVEMENT_CREDIT"]
    elif sens == "B2W":
        df["MONTANT_COMPARAISON"] = df["MOUVEMENT_DEBIT"]
    else:
        raise ValueError("sens doit être 'W2B' ou 'B2W'")

    df.rename(
        columns={
            "CODE_TRANSACTION_OPERATEUR": "CODE_TRANSACTION",
            "NUMERO_COMPTE": "NUMERO_COMPTE",
        },
        inplace=True,
    )

    for col in ["CODE_TRANSACTION", "NUMERO_COMPTE", "TYPE_TRANSACTION"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Colonne téléphone normalisée : voir note section 0/1 +
    # CORRECTIF v3/v4/v5 en tête de fichier.
    if "NUMERO_COMPTE" in df.columns:
        df["NUMERO_COMPTE_NORME"] = _normaliser_telephone(df["NUMERO_COMPTE"])

    return df


# ============================================================
# 3. NETTOYAGE MONTANT
# ============================================================

def nettoyer_montant(df: pd.DataFrame) -> pd.DataFrame:

    df = df.copy()
    df["MONTANT_COMPARAISON"] = (
        pd.to_numeric(df["MONTANT_COMPARAISON"], errors="coerce").fillna(0)
    )
    return df


# ============================================================
# 4. TRI OBLIGATOIRE AVANT L'APPARIEMENT
# ============================================================

def trier(df: pd.DataFrame) -> pd.DataFrame:

    return df.sort_values(
        ["DATE_HEURE", "CODE_TRANSACTION", "MONTANT_COMPARAISON"]
    ).reset_index(drop=True)


# ============================================================
# 5. DÉTECTION DOUBLONS (avant rapprochement)
# ------------------------------------------------------------
# Appliqué séparément sur CHAQUE côté (Excel-partenaire ET
# Flex Oracle), pour CHAQUE sens (W2B et B2W) — un doublon côté
# Wave-Partenaire n'a rien à voir avec un doublon côté Oracle,
# les deux doivent être détectés indépendamment.
# ============================================================

def detecter_doublons(df: pd.DataFrame, colonnes_cle=None) -> pd.DataFrame:

    colonnes_cle = colonnes_cle or ["CODE_TRANSACTION", "DATE_HEURE", "MONTANT_COMPARAISON"]

    d = df[
        df.duplicated(subset=colonnes_cle, keep=False)
    ].copy()

    d["STATUT"] = "Doublon"

    return d


def retirer_doublons(df: pd.DataFrame, colonnes_cle=None):
    """
    Sépare `df` en (df_sans_doublons, doublons) — les doublons
    sont retirés du jeu qui partira dans l'appariement, pour
    ne jamais fausser la recherche du meilleur candidat.
    """

    doublons = detecter_doublons(df, colonnes_cle)
    df_sans_doublons = df.drop(index=doublons.index).reset_index(drop=True)

    return df_sans_doublons, doublons.reset_index(drop=True)


# ============================================================
# 6. DIFFÉRENCE EN SECONDES
# ============================================================

def difference_secondes(a, b) -> float:
    """
    Retourne la différence absolue en secondes.
    Si une des deux dates est invalide (NaT),
    on retourne +∞ afin d'empêcher toute fausse réconciliation.
    """

    if pd.isna(a) or pd.isna(b):
        return float("inf")

    return abs((a - b).total_seconds())


# ============================================================
# 7. MOTEUR D'APPARIEMENT (recherche du meilleur candidat,
#    garde toutes les colonnes)
# ------------------------------------------------------------
# Pour chaque ligne partenaire, on cherche — parmi toutes les
# lignes flex encore disponibles — un candidat de MÊME type de
# transaction, dont l'écart de date est le plus petit, tant que
# cet écart reste sous tolerance_max_secondes (la fenêtre de
# recherche au-delà de laquelle on considère qu'il n'y a tout
# simplement pas de correspondance).
#
# Le critère d'IDENTITÉ utilisé en plus du type de transaction
# dépend de apparier_par_telephone_montant :
#
#   - False (par défaut, Wave/Wizz, INCHANGÉ) : même CODE_TRANSACTION
#     des deux côtés (la référence de transaction est la même
#     référence chez le partenaire et chez Flex pour ces partenaires).
#
#   - True (Orange USSD) : CODE_TRANSACTION n'est PAS comparable d'un
#     système à l'autre pour ce partenaire. Il n'y a donc PAS de clé
#     d'identité unique et fiable disponible des deux côtés — voir
#     CORRECTIF v3/v4/v5 en tête de fichier :
#       * un candidat n'est même évalué QUE s'il partage au moins un
#         signal d'identité avec la ligne partenaire : montant
#         identique OU téléphone (NUMERO_COMPTE_NORME) identique.
#         Sans aucun des deux -> `continue` (candidat écarté, jamais
#         consommé, laissé disponible pour une autre ligne
#         partenaire qui pourrait vraiment matcher dessus).
#       * [CORRECTIF v5] si SEUL le montant correspond (pas le
#         téléphone), le candidat n'est accepté que si l'écart de
#         temps reste sous 1h -- au-delà, un montant identique isolé
#         n'est plus un signal assez fiable (trop de coïncidences
#         possibles sur des montants ronds). Passé 1h, il faut EN
#         PLUS le téléphone identique.
#       * parmi les candidats retenus, on préfère : montant identique
#         d'abord, puis téléphone identique, puis l'écart de temps
#         le plus faible (tuple de score, comparaison lexicographique).
#
# Une fois le meilleur candidat trouvé, le statut est déterminé
# par PALIERS d'écart de date (uniquement quand le montant
# correspond exactement, diff_montant == 0) :
#   0s                          -> "Réconcilié"
#   ]0s , tolerance_secondes]   -> "Réconcilié avec tolérance"
#   ]tolerance_secondes , 1h]   -> "Réconcilié - écart 8s à 1h"
#   ]1h , tolerance_max]        -> "Réconcilié - écart > 1h"
#     (en mode apparier_par_telephone_montant, ce dernier palier
#     suppose désormais TOUJOURS un téléphone identique en plus du
#     montant, grâce au CORRECTIF v5 ci-dessus)
# Si le montant ne correspond pas mais le compte oui -> "Ecart
# montant" (indépendamment du délai). Sinon -> "Non comptabilisée".
# ============================================================

UNE_HEURE_EN_SECONDES = 3600


def reconciliation_two_pointers(
    partenaire: pd.DataFrame,
    flex: pd.DataFrame,
    tolerance_secondes: int = 8,
    tolerance_max_secondes: int = UNE_HEURE_EN_SECONDES * 24,
    fenetre_recherche: int = 150,
    apparier_par_telephone_montant: bool = False,
) -> pd.DataFrame:
    """
    tolerance_secondes     : au-delà de 0s et jusqu'à cette valeur (8s
                              par défaut) -> "Réconcilié avec tolérance".
    tolerance_max_secondes : fenêtre MAXIMALE de recherche d'un
                              candidat (24h par défaut) — au-delà,
                              on considère qu'il n'y a pas de
                              correspondance (Non comptabilisée /
                              Comptabilisation isolée). Les paliers
                              "8s à 1h" et "> 1h" sont calculés à
                              l'intérieur de cette fenêtre.
    fenetre_recherche       : conservé uniquement pour compatibilité
                              de signature avec les appelants existants
                              (reconcilier_un_sens, config.py, ...) —
                              n'est PLUS utilisé pour borner la
                              recherche. La sélection scanne TOUTES
                              les lignes flex encore disponibles.
    apparier_par_telephone_montant : voir docstring de la section 7 et
                              l'en-tête du fichier (CORRECTIF v2 à v5).
                              False par défaut = comportement
                              HISTORIQUE INCHANGÉ (Wave, Wizz, tout
                              partenaire qui ne positionne pas ce
                              paramètre à True). True = Orange USSD
                              uniquement -> appariement par identité
                              minimale (montant OU téléphone), avec
                              exigence de téléphone en plus du montant
                              au-delà d'1h d'écart, puis score
                              montant > téléphone > écart de temps.
    """

    resultat = []

    utilises = set()

    for _, wp in partenaire.iterrows():

        meilleur = None
        meilleur_idx = None
        meilleur_diff = float("inf")
        # Utilisé uniquement en mode apparier_par_telephone_montant :
        # tuple (0 si montant identique sinon 1, 0 si téléphone
        # identique sinon 1, diff_temps), comparaison lexicographique
        # -> priorité au montant, puis au téléphone, puis au temps.
        meilleur_score = None

        # On scanne TOUT le flex non-utilisé restant, et non une
        # fenêtre glissante calée sur len(utilises) (qui est un
        # COMPTEUR de lignes consommées, pas leur POSITION réelle
        # dans le tableau trié).
        for idx in range(0, len(flex)):

            if idx in utilises:
                continue

            wf = flex.iloc[idx]

            if wp["TYPE TRANSACTION"] != wf["TYPE_TRANSACTION"]:
                continue

            diff_temps = difference_secondes(
                wp["DATE_HEURE"],
                wf["DATE_HEURE"]
            )

            if diff_temps > tolerance_max_secondes:
                continue

            if apparier_par_telephone_montant:
                # [Orange USSD — CORRECTIF v3/v4/v5]
                diff_montant_candidat = abs(
                    wp["MONTANT_COMPARAISON"] - wf["MONTANT_COMPARAISON"]
                )
                meme_montant = diff_montant_candidat == 0
                meme_telephone = (
                    pd.notna(wp.get("NUMERO_COMPTE_NORME"))
                    and pd.notna(wf.get("NUMERO_COMPTE_NORME"))
                    and wp["NUMERO_COMPTE_NORME"] == wf["NUMERO_COMPTE_NORME"]
                )

                # [CORRECTIF v4] Identité minimale obligatoire : sans
                # montant NI téléphone en commun, ce candidat n'est
                # pas plausible -> on ne l'évalue même pas, pour ne
                # jamais le "consommer" à tort au détriment d'une
                # ligne partenaire ultérieure qui, elle, a une vraie
                # correspondance dessus.
                if not meme_montant and not meme_telephone:
                    continue

                # [CORRECTIF v5] Un montant identique SEUL (sans
                # téléphone) n'est un signal fiable qu'à courte
                # distance temporelle. Au-delà d'1h, on exige EN PLUS
                # le téléphone -- sinon deux transactions
                # indépendantes qui partagent juste un montant rond
                # (10000, 20000, 40000...) peuvent s'apparier à tort
                # sur une longue fenêtre, et voler la ligne flex
                # d'une AUTRE transaction partenaire qui, elle, avait
                # une vraie correspondance (même bug de fond que v4).
                if not meme_telephone and diff_temps > UNE_HEURE_EN_SECONDES:
                    continue

                score = (
                    0 if meme_montant else 1,
                    0 if meme_telephone else 1,
                    diff_temps,
                )
                if meilleur_score is None or score < meilleur_score:
                    meilleur = wf
                    meilleur_idx = idx
                    meilleur_diff = diff_temps
                    meilleur_score = score
            else:
                # [Wave/Wizz, INCHANGÉ] identité par référence de
                # transaction, la même des deux côtés pour ces
                # partenaires.
                if wp["CODE_TRANSACTION"] != wf["CODE_TRANSACTION"]:
                    continue

                if diff_temps < meilleur_diff:
                    meilleur = wf
                    meilleur_idx = idx
                    meilleur_diff = diff_temps

        if meilleur is not None:

            ligne = {}

            for c in partenaire.columns:
                ligne["WP_" + c] = wp[c]

            for c in flex.columns:
                ligne["WF_" + c] = meilleur[c]

            diff_montant = abs(
                wp["MONTANT_COMPARAISON"]
                - meilleur["MONTANT_COMPARAISON"]
            )

            if diff_montant == 0:

                if meilleur_diff == 0:
                    ligne["STATUT"] = "Réconcilié"
                elif meilleur_diff <= tolerance_secondes:
                    ligne["STATUT"] = "Réconcilié avec tolérance"
                elif meilleur_diff <= UNE_HEURE_EN_SECONDES:
                    ligne["STATUT"] = "Réconcilié - écart 8s à 1h"
                else:
                    ligne["STATUT"] = "Réconcilié - écart > 1h"

            elif wp["NUMERO_COMPTE"] == meilleur["NUMERO_COMPTE"]:

                ligne["STATUT"] = "Ecart montant"

            else:

                ligne["STATUT"] = "Non comptabilisée"

            resultat.append(ligne)
            utilises.add(meilleur_idx)

        else:

            ligne = {}

            for c in partenaire.columns:
                ligne["WP_" + c] = wp[c]

            ligne["STATUT"] = "Non comptabilisée"

            resultat.append(ligne)

    for idx, wf in flex.iterrows():

        if idx in utilises:
            continue

        ligne = {}

        for c in flex.columns:
            ligne["WF_" + c] = wf[c]

        ligne["STATUT"] = "Comptabilisation isolée"

        resultat.append(ligne)

    return pd.DataFrame(resultat)


# ============================================================
# 8. PIPELINE COMPLET POUR UN SENS (W2B ou B2W)
# ============================================================

def _prefixer_doublons(doublons: pd.DataFrame, prefixe: str) -> pd.DataFrame:
    """Renomme les colonnes d'un DataFrame de doublons avec le préfixe
    WP_/WF_, pour qu'il ait le même schéma que les lignes produites par
    reconciliation_two_pointers (colonnes mélangées + STATUT)."""

    if doublons.empty:
        return pd.DataFrame()

    colonnes_a_prefixer = [c for c in doublons.columns if c != "STATUT"]
    renomme = doublons.rename(columns={c: f"{prefixe}_{c}" for c in colonnes_a_prefixer})
    return renomme


def reconcilier_un_sens(
    wp_raw: pd.DataFrame,
    wf_raw: pd.DataFrame,
    sens: str,
    tolerance_secondes: int = 8,
    tolerance_max_secondes: int = UNE_HEURE_EN_SECONDES * 24,
    fenetre_recherche: int = 150,
    apparier_par_telephone_montant: bool = False,
) -> pd.DataFrame:
    """
    Enchaîne, pour un sens donné (W2B ou B2W) :
      1. préparation + nettoyage des deux côtés
      2. détection ET retrait des doublons (Excel-partenaire ET Flex),
         AVANT l'appariement, pour ne jamais fausser la recherche du
         meilleur candidat avec des lignes en double
      3. tri + appariement (meilleur candidat en temps) sur les
         données dédupliquées, avec paliers de statut selon l'écart
         de date (voir reconciliation_two_pointers)
      4. réinjection des doublons dans le résultat final, avec
         STATUT = "Doublon" et le même schéma de colonnes (WP_/WF_)
         que les autres statuts

    apparier_par_telephone_montant : transmis tel quel à
        reconciliation_two_pointers (voir sa docstring). False par
        défaut = comportement historique inchangé (Wave, Wizz).
        Mettre à True uniquement pour Orange USSD (voir config.py
        et reconc.py -> run_reconciliation).
    """

    wp_prepare = nettoyer_montant(preparer_wave_partenaire(wp_raw, sens=sens))
    wf_prepare = nettoyer_montant(preparer_wave_flex(wf_raw, sens))

    wp_dedup, wp_doublons = retirer_doublons(wp_prepare)
    wf_dedup, wf_doublons = retirer_doublons(wf_prepare)

    wp_tri = trier(wp_dedup)
    wf_tri = trier(wf_dedup)

    resultat_appariement = reconciliation_two_pointers(
        wp_tri, wf_tri,
        tolerance_secondes=tolerance_secondes,
        tolerance_max_secondes=tolerance_max_secondes,
        fenetre_recherche=fenetre_recherche,
        apparier_par_telephone_montant=apparier_par_telephone_montant,
    )

    doublons_wp = _prefixer_doublons(wp_doublons, "WP")
    doublons_wf = _prefixer_doublons(wf_doublons, "WF")

    return pd.concat(
        [resultat_appariement, doublons_wp, doublons_wf],
        ignore_index=True
    )


# ============================================================
# 9. TABLE RÉSUMÉ (schéma métier fixe, lisible)
# ------------------------------------------------------------
# Aplati resultat_final (colonnes WP_/WF_ mélangées, différentes
# selon le statut) vers un schéma fixe et lisible :
#   Type Transaction, Montant Partenaire, Montant Flex,
#   Num_Tel_Client, Nom_Client, Compte, Agence, Ecart Montant,
#   Diff Heure, Date Fichier Partenaire, Periode Fichier, Statut
#
# `colonnes_supplementaires` (venant de config.PARTENAIRES[...]
# ["colonnes_resume"]) indique dans quelle colonne déjà préfixée
# (WP_.../WF_...) trouver Num_Tel_Client / Nom_Client / Agence /
# Periode Fichier pour CE partenaire. Si l'info n'existe pas pour
# ce partenaire (mapping à None ou colonne absente), la colonne
# résultante est simplement vide — jamais d'erreur.
# ============================================================

def construire_table_resume(resultat: pd.DataFrame, colonnes_supplementaires: dict = None) -> pd.DataFrame:

    colonnes_supplementaires = colonnes_supplementaires or {}
    n = len(resultat)

    def _noms(valeur):
        """Normalise une valeur de config (None / str / liste) en liste de noms."""
        if not valeur:
            return []
        return [valeur] if isinstance(valeur, str) else list(valeur)

    def colonne_texte(valeur):
        """Renvoie la première colonne non vide parmi la liste de noms
        candidats (repli WP -> WF, par ex. Nom_Client), sinon une
        colonne vide."""
        serie = pd.Series([None] * n, index=resultat.index, dtype=object)
        for nom in _noms(valeur):
            if nom in resultat.columns:
                serie = serie.where(serie.notna(), resultat[nom])
        return serie

    def colonne_num(nom):
        if nom and nom in resultat.columns:
            return pd.to_numeric(resultat[nom], errors="coerce")
        return pd.Series([pd.NA] * n, index=resultat.index, dtype="Float64")

    def colonne_date(valeur):
        serie = pd.Series([pd.NaT] * n, index=resultat.index)
        for nom in _noms(valeur):
            if nom in resultat.columns:
                serie = serie.where(serie.notna(), pd.to_datetime(resultat[nom], errors="coerce"))
        return serie

    montant_partenaire = colonne_num("WP_MONTANT_COMPARAISON")
    montant_flex = colonne_num("WF_MONTANT_COMPARAISON")

    compte = colonne_texte(["WP_NUMERO_COMPTE", "WF_NUMERO_COMPTE"])

    ecart_montant = (montant_partenaire - montant_flex).abs()

    date_heure_wp = colonne_date("WP_DATE_HEURE")
    date_heure_wf = colonne_date("WF_DATE_HEURE")
    diff_heure_secondes = (date_heure_wp - date_heure_wf).dt.total_seconds().abs()

    num_tel_client = colonne_texte(colonnes_supplementaires.get("num_tel_client"))
    nom_client = colonne_texte(colonnes_supplementaires.get("nom_client"))
    agence = colonne_texte(colonnes_supplementaires.get("agence"))
    periode_fichier = colonne_date(colonnes_supplementaires.get("periode_fichier"))

    # Type de transaction (W2B / B2W) : la colonne SENS (ajoutée par
    # run_reconciliation lors de la concaténation W2B+B2W) est la
    # source la plus fiable ; repli sur les colonnes TYPE TRANSACTION
    # brutes si un résultat plus ancien (sans SENS) est encore en base.
    if "SENS" in resultat.columns:
        type_transaction = resultat["SENS"]
    else:
        type_transaction = colonne_texte(["WP_TYPE TRANSACTION", "WF_TYPE_TRANSACTION"])

    # Date du fichier partenaire : date/heure de la transaction telle
    # que déclarée par l'Excel-partenaire lui-même (par opposition à
    # "Periode Fichier", qui est la date côté Flex Oracle).
    date_fichier_partenaire = date_heure_wp

    resume = pd.DataFrame({
        "Type Transaction": type_transaction,
        "Montant Partenaire": montant_partenaire,
        "Montant Flex": montant_flex,
        "Num_Tel_Client": num_tel_client,
        "Nom_Client": nom_client,
        "Compte": compte,
        "Agence": agence,
        "Ecart Montant": ecart_montant,
        "Diff Heure": diff_heure_secondes,
        "Date Fichier Partenaire": date_fichier_partenaire,
        "Periode Fichier": periode_fichier,
        "Statut": resultat["STATUT"] if "STATUT" in resultat.columns else pd.Series([None] * n),
    })

    return resume


# ============================================================
# 10. TAUX DE RÉUSSITE DE LA RÉCONCILIATION
# ------------------------------------------------------------
# Rapporté à l'ENSEMBLE des lignes (tous statuts confondus,
# y compris les nouveaux paliers "écart 8s à 1h" / "écart > 1h") :
#   taux = (Réconcilié + Réconcilié avec tolérance) / total * 100
#
# Les 4 statuts "Réconcilié*" (délai 0s, ]0,8s], ]8s,1h], >1h) sont
# TOUS comptés comme réconciliés dans le taux global : dans les
# 4 cas, le code transaction, le type et le MONTANT correspondent
# exactement — seul le délai diffère. Seuls "Ecart montant", "Non
# comptabilisée", "Comptabilisation isolée" et "Doublon" sont exclus
# du taux de réussite.
# ============================================================

STATUTS_RECONCILIES = [
    "Réconcilié",
    "Réconcilié avec tolérance",
    "Réconcilié - écart 8s à 1h",
    "Réconcilié - écart > 1h",
]


def calculer_taux_reussite(resultat_ou_comptes) -> dict:
    """
    Accepte soit un DataFrame résultat (colonne STATUT), soit un
    dict déjà agrégé {STATUT: nombre}. Renvoie le détail complet :
    taux (%), nombre de lignes réconciliées, total, et le compte
    par statut (toujours tous les statuts de STATUTS, 0 si absent).
    """

    if isinstance(resultat_ou_comptes, pd.DataFrame):
        comptes = resultat_ou_comptes["STATUT"].value_counts().to_dict() if "STATUT" in resultat_ou_comptes.columns else {}
    else:
        comptes = dict(resultat_ou_comptes or {})

    comptes_complets = {s: int(comptes.get(s, 0)) for s in STATUTS}

    total = sum(comptes_complets.values())
    reconcilies = sum(comptes_complets[s] for s in STATUTS_RECONCILIES)

    taux = round(100 * reconcilies / total, 2) if total > 0 else 0.0

    return {
        "taux_reussite": taux,
        "reconcilies": reconcilies,
        "total": total,
        "comptes": comptes_complets,
    }


# ============================================================
# 11. RÉCONCILIATION AGRÉGÉE PAR AGENCE (SOLDE = DEBIT - CREDIT)
# ------------------------------------------------------------
# [INCHANGÉ] Toute la section 11 (mode "agence" : Orange Agence,
# Wave Agence, RIA Agence) est identique à l'original — les
# correctifs d'appariement Orange USSD ci-dessus ne concernent QUE
# le mode "two_pointers" (preparer_wave_partenaire /
# preparer_wave_flex / reconciliation_two_pointers /
# reconcilier_un_sens), donc Wave, Wizz et Orange USSD. Cette
# section n'est pas touchée.
# ============================================================

CODE_AGENCE_NON_MAPPE = "AGENCE_NON_MAPPEE"


def _normaliser_code_agence(serie: pd.Series) -> pd.Series:
    """
    Rend CODE_AGENCE comparable quel que soit le format d'origine
    ('012', '12', '12.0', ' 12 ' -> '12'), et route explicitement
    vers CODE_AGENCE_NON_MAPPE toute valeur manquante (NaN côté
    Partenaire = agence absente du référentiel de mapping, ou
    CODE_AGENCE vide côté Flex), au lieu de la chaîne "nan" muette.
    """
    s = serie.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.str.replace(r"^0+(?=\d)", "", regex=True)
    s = s.where(~s.str.lower().isin(["nan", "none", ""]), CODE_AGENCE_NON_MAPPE)
    return s


def preparer_agence_partenaire(
    df: pd.DataFrame,
    col_credit: str,
    col_debit: str,
):
    """
    Calcule SOLDE_PARTENAIRE de deux façons possibles :

      - Si `df` contient déjà une colonne "SOLDE_PARTENAIRE" (ex:
        Wave Agence, calculé à la source à partir du montant signé de
        TOUTES les transactions), on la réutilise directement — on ne
        la recalcule JAMAIS à partir de col_debit/col_credit, qui ne
        couvrent qu'une partie des cas et donneraient un solde faux.

      - Sinon (ex: Orange Agence), on calcule
        SOLDE_PARTENAIRE = col_debit - col_credit, comme avant.

    Dans les deux cas, DEBIT_PARTENAIRE / CREDIT_PARTENAIRE restent
    calculées à partir de col_debit / col_credit quand ces colonnes
    existent (à titre informatif / audit uniquement), sinon mises à 0.
    """

    df = df.copy()
    df.columns = df.columns.str.upper().str.strip()

    col_credit_u = (col_credit or "").upper()
    col_debit_u = (col_debit or "").upper()

    a_credit = bool(col_credit_u) and col_credit_u in df.columns
    a_debit = bool(col_debit_u) and col_debit_u in df.columns

    if a_credit:
        df[col_credit_u] = pd.to_numeric(df[col_credit_u], errors="coerce").fillna(0)
    if a_debit:
        df[col_debit_u] = pd.to_numeric(df[col_debit_u], errors="coerce").fillna(0)

    df["CODE_AGENCE"] = _normaliser_code_agence(df["CODE_AGENCE"])

    label = next(
        (c for c in ("AGENCE", "NOM_AGENCE", "LIBELLE_AGENCE") if c in df.columns),
        None
    )

    solde_deja_calcule = "SOLDE_PARTENAIRE" in df.columns

    if solde_deja_calcule:
        # Cas Wave Agence : le solde net est déjà correct à la
        # source, on ne fait que le rendre numérique avant de le
        # ré-agréger par CODE_AGENCE.
        df["SOLDE_PARTENAIRE"] = pd.to_numeric(df["SOLDE_PARTENAIRE"], errors="coerce").fillna(0)
    elif a_debit and a_credit:
        # Cas Orange Agence (et compat. générale) : solde reconstruit
        # à partir de col_debit - col_credit.
        df["SOLDE_PARTENAIRE"] = df[col_debit_u] - df[col_credit_u]
    else:
        raise KeyError(
            "Impossible de calculer SOLDE_PARTENAIRE : ni colonne "
            "'SOLDE_PARTENAIRE' déjà présente, ni les deux colonnes "
            f"col_debit='{col_debit}' / col_credit='{col_credit}' "
            "trouvées dans le fichier partenaire."
        )

    agg_kwargs = {"SOLDE_PARTENAIRE": ("SOLDE_PARTENAIRE", "sum")}
    if a_debit:
        agg_kwargs["DEBIT_PARTENAIRE"] = (col_debit_u, "sum")
    if a_credit:
        agg_kwargs["CREDIT_PARTENAIRE"] = (col_credit_u, "sum")
    agg_kwargs["AGENCE"] = (label, "first") if label else ("CODE_AGENCE", "first")

    agg = (
        df.groupby("CODE_AGENCE", dropna=False)
          .agg(**agg_kwargs)
          .reset_index()
    )

    # DEBIT_PARTENAIRE / CREDIT_PARTENAIRE toujours présentes dans le
    # résultat (0 si la colonne source n'existait pas), pour que le
    # reste du pipeline (construire_table_resume_agence, etc.) n'ait
    # jamais besoin de tester leur existence.
    if "DEBIT_PARTENAIRE" not in agg.columns:
        agg["DEBIT_PARTENAIRE"] = 0.0
    if "CREDIT_PARTENAIRE" not in agg.columns:
        agg["CREDIT_PARTENAIRE"] = 0.0

    return agg


def preparer_agence_flex(df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrège le fichier Flex Oracle (sortie orange_flex_api.py /
    wave_agence_flex_api.py / RIA_agence_flex.py) par CODE_AGENCE :
      - somme de DEBIT / CREDIT (info/audit, conservées)
      - SOLDE_FLEX = somme(DEBIT) - somme(CREDIT)

    Commun à tous les partenaires en mode "agence" : le nom de ces
    deux colonnes est toujours DEBIT/CREDIT côté Flex. Fonctionne
    aussi bien sur un journal brut (une ligne par écriture) que sur
    un agrégat déjà réduit à une ligne par agence (wave_agence_flex_api.py) :
    dans ce dernier cas, le groupby-sum est simplement une identité.
    """

    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.upper()

    for col in ("DEBIT", "CREDIT", "CODE_AGENCE"):
        if col not in df.columns:
            raise KeyError(col)

    df["DEBIT"] = pd.to_numeric(df["DEBIT"], errors="coerce").fillna(0)
    df["CREDIT"] = pd.to_numeric(df["CREDIT"], errors="coerce").fillna(0)
    df["CODE_AGENCE"] = _normaliser_code_agence(df["CODE_AGENCE"])

    agg = df.groupby("CODE_AGENCE", dropna=False).agg(
        DEBIT_FLEX=("DEBIT", "sum"),
        CREDIT_FLEX=("CREDIT", "sum"),
    ).reset_index()

    # Solde flex = débit - crédit
    agg["SOLDE_FLEX"] = agg["DEBIT_FLEX"] - agg["CREDIT_FLEX"]

    return agg


def reconcilier_par_agence(
    wp_raw: pd.DataFrame,
    wf_raw: pd.DataFrame,
    tolerance_montant: float = 0,
    col_credit: str = "MONTANT_CASHIN",
    col_debit: str = "MONTANT_CASHOUT",
) -> pd.DataFrame:
    """
    Rapprochement agrégé par CODE_AGENCE, basé sur UN SEUL écart de
    SOLDE par agence :

        ECART_SOLDE = SOLDE_PARTENAIRE - SOLDE_FLEX

    tolerance_montant : écart absolu toléré sur le SOLDE pour rester
                        classé "Réconcilié" (0 = égalité stricte).
    col_credit/col_debit : noms des colonnes Excel-partenaire utilisées
                        pour calculer SOLDE_PARTENAIRE quand celui-ci
                        n'est pas déjà présent dans wp_raw (voir
                        preparer_agence_partenaire et
                        config.PARTENAIRES[...]["colonnes_agence"]).
                        Pour Wave Agence, wp_raw contient déjà
                        "SOLDE_PARTENAIRE" -> ces deux paramètres ne
                        servent alors qu'à calculer DEBIT_PARTENAIRE/
                        CREDIT_PARTENAIRE à titre informatif.

    Statuts produits (sous-ensemble de STATUTS) :
      - "Réconcilié"              : agence des deux côtés, |écart
                                     solde| <= tolérance
      - "Ecart montant"           : agence des deux côtés, mais écart
                                     de solde au-delà de la tolérance
      - "Non comptabilisée"       : agence présente côté Partenaire
                                     uniquement (absente du Flex)
      - "Comptabilisation isolée" : agence présente côté Flex
                                     uniquement (absente du Partenaire)

    La ligne correspondant aux comptes non mappés à une agence
    (CODE_AGENCE == CODE_AGENCE_NON_MAPPE, "AGENCE_NON_MAPPEE") reste
    exclue du résultat retourné, comme avant.
    """

    partenaire_agg = preparer_agence_partenaire(wp_raw, col_credit=col_credit, col_debit=col_debit)
    flex_agg = preparer_agence_flex(wf_raw)

    fusion = pd.merge(
        partenaire_agg, flex_agg,
        on="CODE_AGENCE", how="outer", indicator=True
    )

    for col in ("CREDIT_PARTENAIRE", "DEBIT_PARTENAIRE", "SOLDE_PARTENAIRE",
                "CREDIT_FLEX", "DEBIT_FLEX", "SOLDE_FLEX"):
        fusion[col] = fusion[col].fillna(0)

    # Un seul écart désormais : celui du solde.
    fusion["ECART_SOLDE"] = fusion["SOLDE_PARTENAIRE"] - fusion["SOLDE_FLEX"]

    def _statut(row):
        if row["_merge"] == "left_only":
            return "Non comptabilisée"
        if row["_merge"] == "right_only":
            return "Comptabilisation isolée"
        if abs(row["ECART_SOLDE"]) <= tolerance_montant:
            return "Réconcilié"
        return "Ecart montant"

    fusion["STATUT"] = fusion.apply(_statut, axis=1)
    fusion.drop(columns=["_merge"], inplace=True)

    fusion = fusion[fusion["CODE_AGENCE"] != CODE_AGENCE_NON_MAPPE].copy()

    return fusion.sort_values("CODE_AGENCE").reset_index(drop=True)


def construire_table_resume_agence(resultat: pd.DataFrame) -> pd.DataFrame:
    """Met le résultat de reconcilier_par_agence en forme lisible :
    met SOLDE Partenaire / SOLDE Flex / Ecart Solde en avant (le
    détail DEBIT/CREDIT reste dispo pour l'audit)."""

    n = len(resultat)

    resume = pd.DataFrame({
        "Code Agence": resultat.get("CODE_AGENCE", pd.Series([None] * n)),
        "Agence": resultat.get("AGENCE", pd.Series([None] * n)),
        "Solde Partenaire": resultat.get("SOLDE_PARTENAIRE", pd.Series([None] * n)),
        "Solde Flex": resultat.get("SOLDE_FLEX", pd.Series([None] * n)),
        "Ecart Solde": resultat.get("ECART_SOLDE", pd.Series([None] * n)),
        "Débit Partenaire": resultat.get("DEBIT_PARTENAIRE", pd.Series([None] * n)),
        "Crédit Partenaire": resultat.get("CREDIT_PARTENAIRE", pd.Series([None] * n)),
        "Débit Flex": resultat.get("DEBIT_FLEX", pd.Series([None] * n)),
        "Crédit Flex": resultat.get("CREDIT_FLEX", pd.Series([None] * n)),
        "Statut": resultat.get("STATUT", pd.Series([None] * n)),
    })

    return resume