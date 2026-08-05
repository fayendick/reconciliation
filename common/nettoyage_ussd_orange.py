# ============================================================
# NETTOYAGE — Rapport "Daily-ChannelUserTransactionReport" (USSD Partenaire Orange)
# ------------------------------------------------------------
# Le fichier brut Orange Money contient, avant les vraies lignes
# de transactions :
#   - un bandeau d'en-tête (Application / Réseau / Période / ...)
#   - PUIS le vrai en-tête de tableau :
#       N° | Date | Heure | Référence | Service | Paiement | Statut
#       | Mode | N° de Compte | Wallet | N° Pseudo | N° de Compte
#       | Wallet | Débit | Crédit | Compte: <NUMERO> | Sous-réseau
#   - PUIS un bloc "Commissions cumulées" + un bloc "Transactions
#     échouées" + une ligne "Solde initial" + une ligne "Commissions"
#     -> tout ceci n'est PAS une transaction réelle.
#   - Les vraies transactions ("Statut" == "Succès") suivent.
#   - En fin de fichier : lignes "Total" et "Solde" (totaux, à
#     exclure elles aussi).
#
# Ce module repère l'en-tête dynamiquement (peu importe le numéro
# de ligne exact ni le numéro de compte affiché dans "Compte: xxx"),
# puis ne garde que les lignes qui sont de VRAIES transactions
# réussies (Statut == "Succès", Référence renseignée, Date valide).
# Cela reproduit exactement le comportement demandé : tout ce qui
# se trouve entre l'en-tête et la dernière ligne "Commissions"
# (récap commissions, transactions échouées, totaux) est éliminé,
# et fonctionne aussi bien pour les prochains exports quotidiens
# (nombre de lignes différent à chaque fois).
# ============================================================

import io
from typing import Union

import pandas as pd

COLONNES_ATTENDUES_DEBUT = ["N°", "Date", "Heure", "Référence", "Service"]


def _trouver_ligne_entete(raw: pd.DataFrame) -> int:
    """Retourne l'index de la ligne d'en-tête réelle du tableau
    (celle qui commence par N° / Date / Heure / Référence / Service)."""

    for i, row in raw.iterrows():
        valeurs = [str(v).strip() for v in row[:5].tolist()]
        if valeurs == COLONNES_ATTENDUES_DEBUT:
            return i

    raise ValueError(
        "Ligne d'en-tête introuvable (attendu : N°, Date, Heure, Référence, "
        "Service en début de ligne). Le format du fichier a peut-être changé."
    )


def lire_ussd_orange_brut(source: Union[str, bytes, io.BytesIO]) -> pd.DataFrame:
    """Lit le fichier .xls brut Orange Money (USSD Partenaire) tel quel,
    sans aucun nettoyage — utile pour debug."""
    return pd.read_excel(source, header=None, engine="xlrd")


def nettoyer_ussd_orange(
    source: Union[str, bytes, io.BytesIO],
    garder_uniquement_succes: bool = True,
) -> pd.DataFrame:
    """
    Nettoie un export Orange Money "Daily-ChannelUserTransactionReport"
    (USSD Partenaire) et retourne un DataFrame propre, une ligne par
    transaction, avec les vrais en-têtes de colonnes.

    garder_uniquement_succes : si True (par défaut), ne garde que les
        lignes Statut == "Succès" (élimine du même coup les transactions
        échouées, les récapitulatifs de commissions, et les lignes
        Total/Solde de fin de fichier — aucune de ces lignes n'a le
        Statut "Succès"). Mettre False pour garder aussi les échecs.
    """

    raw = pd.read_excel(source, header=None, engine="xlrd")

    ligne_entete = _trouver_ligne_entete(raw)
    entetes = [str(v).strip() for v in raw.iloc[ligne_entete].tolist()]

    data = raw.iloc[ligne_entete + 1:].copy()
    data.columns = entetes
    data = data.reset_index(drop=True)

    # Les deux colonnes "N° de Compte" (Agent, puis Correspondant) et les
    # deux colonnes "Wallet" ont le même nom brut -> on les distingue par
    # position pour éviter toute ambiguïté dans la suite du traitement.
    data.columns = _dedupliquer_colonnes(list(data.columns))

    date_parsee = pd.to_datetime(data["Date"], errors="coerce", dayfirst=True)
    reference_ok = (
        data["Référence"].notna()
        & (data["Référence"].astype(str).str.strip() != "")
        & (data["Référence"].astype(str).str.strip().str.lower() != "nan")
    )

    masque = date_parsee.notna() & reference_ok

    if garder_uniquement_succes:
        masque &= data["Statut"].astype(str).str.strip() == "Succès"

    propre = data[masque].copy()
    propre.insert(1, "DATE_PARSEE", date_parsee[masque])
    propre = propre.reset_index(drop=True)

    return propre


def _dedupliquer_colonnes(colonnes: list) -> list:
    """Renomme la 2e occurrence de 'N° de Compte' -> 'N° de Compte (Correspondant)'
    et la 2e occurrence de 'Wallet' -> 'Wallet (Correspondant)', en gardant les
    premières occurrences (Agent) inchangées."""

    vues = {}
    resultat = []

    for c in colonnes:
        vues[c] = vues.get(c, 0) + 1
        if c == "N° de Compte" and vues[c] == 2:
            resultat.append("N° de Compte (Correspondant)")
        elif c == "Wallet" and vues[c] == 2:
            resultat.append("Wallet (Correspondant)")
        else:
            resultat.append(c)

    return resultat


if __name__ == "__main__":
    # Nettoie le fichier fourni et écrit le résultat en xlsx à côté.
    import sys

    chemin_entree = sys.argv[1] if len(sys.argv) > 1 else (
        "/mnt/user-data/uploads/"
        "1785278013491_Copie_de_Daily-ChannelUserTransactionReport-786256338-20260717__002__USSD_Partenaire.xls"
    )
    chemin_sortie = sys.argv[2] if len(sys.argv) > 2 else (
        "/mnt/user-data/outputs/USSD_Partenaire_786256338_20260717_NETTOYE.xlsx"
    )

    df = nettoyer_ussd_orange(chemin_entree)
    df.to_excel(chemin_sortie, index=False)
    print(f"{len(df)} transactions -> {chemin_sortie}")