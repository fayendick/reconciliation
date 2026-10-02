
import pandas as pd
from pathlib import Path

from partners.western.western_flex_api import get_western_flex
from partners.western.app_western import (
    lire_fichier_western,
    lire_mapping,
    appliquer_mapping,
    agreger_par_agence,
)


DATE_DEBUT = "24/09/2026"
DATE_FIN = "24/09/2026"

FICHIER_PARTENAIRE = r"data\2609250M9Z5Z.xls"
FICHIER_MAPPING = r"data\mappings\ADJ WU.xlsx"


# ============================================================
# 1. PARTENAIRE WESTERN
# ============================================================

print("\n" + "=" * 100)
print("                    PARTENAIRE WESTERN")
print("=" * 100)

class FichierTest:
    def __init__(self, chemin):
        self.file = open(chemin, "rb")
        self.filename = Path(chemin).name


fichier_test = FichierTest(FICHIER_PARTENAIRE)

try:
    df_partenaire = lire_fichier_western(fichier_test)
finally:
    fichier_test.file.close()

mapping = lire_mapping()

df_partenaire = appliquer_mapping(
    df_partenaire,
    mapping,
)

df_partenaire_agence = agreger_par_agence(
    df_partenaire
)

# On garde uniquement les agences qui ont des transactions
# à réconcilier.
df_partenaire_agence = df_partenaire_agence[
    df_partenaire_agence["NB_TRANSACTIONS"] > 0
].copy()

print("\n===== PARTENAIRE PAR AGENCE =====")

print(
    df_partenaire_agence[
        [
            "CODE_AGENCE",
            "NOM_AGENCE",
            "NB_TRANSACTIONS",
            "NB_PAIEMENT_WU",
            "NB_ENVOI_WU",
            "DEBIT_PARTENAIRE",
            "CREDIT_PARTENAIRE",
            "SOLDE_PARTENAIRE",
        ]
    ].to_string(index=False)
)


# ============================================================
# 2. EXTRACTION FLEX
# ============================================================

print("\n" + "=" * 100)
print("                    FLEX")
print("=" * 100)

df_flex = get_western_flex(
    date_debut=DATE_DEBUT,
    date_fin=DATE_FIN,
)

if df_flex.empty:
    print("Aucune donnée Flex.")
    raise SystemExit(0)


df_flex["DEBIT"] = pd.to_numeric(
    df_flex["DEBIT"],
    errors="coerce",
).fillna(0)

df_flex["CREDIT"] = pd.to_numeric(
    df_flex["CREDIT"],
    errors="coerce",
).fillna(0)

df_flex["CODE_AGENCE"] = pd.to_numeric(
    df_flex["CODE_AGENCE"],
    errors="coerce",
).astype("Int64")


# ============================================================
# 3. IDENTIFICATION DES ECRITURES DE TAXE
# ============================================================

colonnes_texte = [
    "DESCRIPTION",
    "DESCRIPTION BATCH",
    "LIBELLE_OPER",
]

for colonne in colonnes_texte:
    if colonne not in df_flex.columns:
        df_flex[colonne] = ""

    df_flex[colonne] = (
        df_flex[colonne]
        .fillna("")
        .astype(str)
        .str.upper()
    )


df_flex["EST_TAXE"] = False

for colonne in colonnes_texte:
    df_flex["EST_TAXE"] = (
        df_flex["EST_TAXE"]
        | df_flex[colonne].str.contains(
            "TAXE",
            na=False,
        )
    )


# ============================================================
# 4. TAXE PAR AGENCE
# ============================================================

df_taxe = df_flex[
    df_flex["EST_TAXE"]
].copy()

df_taxe_agence = (
    df_taxe
    .groupby(
        "CODE_AGENCE",
        dropna=False,
    )
    .agg(
        NB_ECRITURES_TAXE=("TRN_REF_NO", "count"),
        DEBIT_TAXE=("DEBIT", "sum"),
        CREDIT_TAXE=("CREDIT", "sum"),
    )
    .reset_index()
)

df_taxe_agence["TAXE"] = (
    df_taxe_agence["DEBIT_TAXE"]
    + df_taxe_agence["CREDIT_TAXE"]
)


print("\n===== TAXES FLEX =====")

if df_taxe_agence.empty:
    print("Aucune écriture de taxe détectée.")
else:
    print(
        df_taxe_agence.to_string(
            index=False
        )
    )


# ============================================================
# 5. FLEX HORS TAXE
# ============================================================

df_flex_hors_taxe = df_flex[
    ~df_flex["EST_TAXE"]
].copy()


df_flex_agence = (
    df_flex_hors_taxe
    .groupby(
        [
            "CODE_AGENCE",
            "LIBELLE_AGENCE",
        ],
        dropna=False,
    )
    .agg(
        NB_ECRITURES_FLEX=("TRN_REF_NO", "count"),
        DEBIT_FLEX=("DEBIT", "sum"),
        CREDIT_FLEX=("CREDIT", "sum"),
    )
    .reset_index()
)


df_flex_agence["SOLDE_FLEX"] = (
    df_flex_agence["DEBIT_FLEX"]
    - df_flex_agence["CREDIT_FLEX"]
)


# ============================================================
# 6. COMPARAISON PAR AGENCE
# ============================================================

resultat = df_partenaire_agence[
    [
        "CODE_AGENCE",
        "NOM_AGENCE",
        "NB_TRANSACTIONS",
        "NB_PAIEMENT_WU",
        "NB_ENVOI_WU",
        "DEBIT_PARTENAIRE",
        "CREDIT_PARTENAIRE",
        "SOLDE_PARTENAIRE",
    ]
].copy()


resultat = resultat.merge(
    df_flex_agence[
        [
            "CODE_AGENCE",
            "NB_ECRITURES_FLEX",
            "DEBIT_FLEX",
            "CREDIT_FLEX",
            "SOLDE_FLEX",
        ]
    ],
    on="CODE_AGENCE",
    how="left",
)


resultat = resultat.merge(
    df_taxe_agence[
        [
            "CODE_AGENCE",
            "NB_ECRITURES_TAXE",
            "TAXE",
        ]
    ],
    on="CODE_AGENCE",
    how="left",
)


# Valeurs manquantes = 0
colonnes_numeriques = [
    "NB_ECRITURES_FLEX",
    "DEBIT_FLEX",
    "CREDIT_FLEX",
    "SOLDE_FLEX",
    "NB_ECRITURES_TAXE",
    "TAXE",
]

for colonne in colonnes_numeriques:
    resultat[colonne] = (
        pd.to_numeric(
            resultat[colonne],
            errors="coerce",
        )
        .fillna(0)
    )


# ============================================================
# 7. SURPLUS CREDIT PARTENAIRE
# ============================================================

resultat["CREDIT_SURPLUS_PARTENAIRE"] = (
    resultat["CREDIT_PARTENAIRE"]
    - resultat["CREDIT_FLEX"]
)


# ============================================================
# 8. ECART DEBIT
# ============================================================

resultat["ECART_DEBIT"] = (
    resultat["DEBIT_PARTENAIRE"]
    - resultat["DEBIT_FLEX"]
)


# ============================================================
# 9. ECART CREDIT
# ============================================================

resultat["ECART_CREDIT"] = (
    resultat["CREDIT_PARTENAIRE"]
    - resultat["CREDIT_FLEX"]
)


# ============================================================
# 10. AFFICHAGE FINAL
# ============================================================

colonnes_resultat = [
    "CODE_AGENCE",
    "NOM_AGENCE",
    "NB_TRANSACTIONS",
    "NB_PAIEMENT_WU",
    "NB_ENVOI_WU",
    "DEBIT_PARTENAIRE",
    "CREDIT_PARTENAIRE",
    "DEBIT_FLEX",
    "CREDIT_FLEX",
    "TAXE",
    "CREDIT_SURPLUS_PARTENAIRE",
    "ECART_DEBIT",
    "ECART_CREDIT",
    "SOLDE_PARTENAIRE",
    "SOLDE_FLEX",
]


print("\n")
print("=" * 120)
print("                    COMPARAISON PARTENAIRE / FLEX")
print("=" * 120)

print(
    resultat[
        colonnes_resultat
    ]
    .sort_values("CODE_AGENCE")
    .to_string(index=False)
)


# ============================================================
# 11. CONTROLE SPECIFIQUE DES TAXES
# ============================================================

print("\n")
print("=" * 120)
print("                    CONTROLE TAXES")
print("=" * 120)

print(
    resultat[
        [
            "CODE_AGENCE",
            "NOM_AGENCE",
            "CREDIT_PARTENAIRE",
            "CREDIT_FLEX",
            "CREDIT_SURPLUS_PARTENAIRE",
            "TAXE",
        ]
    ]
    .sort_values("CODE_AGENCE")
    .to_string(index=False)
)


# ============================================================
# 12. TOTAUX
# ============================================================

print("\n")
print("=" * 120)
print("                    TOTAUX")
print("=" * 120)

print(
    f"DEBIT PARTENAIRE       : "
    f"{resultat['DEBIT_PARTENAIRE'].sum():,.0f}"
)

print(
    f"CREDIT PARTENAIRE     : "
    f"{resultat['CREDIT_PARTENAIRE'].sum():,.0f}"
)

print(
    f"DEBIT FLEX HORS TAXE  : "
    f"{resultat['DEBIT_FLEX'].sum():,.0f}"
)

print(
    f"CREDIT FLEX HORS TAXE: "
    f"{resultat['CREDIT_FLEX'].sum():,.0f}"
)

print(
    f"TAXE FLEX             : "
    f"{resultat['TAXE'].sum():,.0f}"
)

print(
    f"SURPLUS CREDIT        : "
    f"{resultat['CREDIT_SURPLUS_PARTENAIRE'].sum():,.0f}"
)

print(
    f"ECART DEBIT           : "
    f"{resultat['ECART_DEBIT'].sum():,.0f}"
)

print(
    f"ECART CREDIT          : "
    f"{resultat['ECART_CREDIT'].sum():,.0f}"
)


print("\n")
print("=" * 120)
print("                    ANALYSE TERMINÉE")
print("=" * 120)

# ============================================================
# 13. ANALYSE DES MOUVEMENTS FLEX
# ============================================================

print("\n")
print("=" * 120)
print("                    ANALYSE DES MOUVEMENTS FLEX")
print("=" * 120)

df_analyse = df_flex.copy()


# ------------------------------------------------------------
# 13.1 Compte + code transaction
# ------------------------------------------------------------

analyse_compte = (
    df_analyse
    .groupby(
        [
            "ACCOUNT_NO",
            "TRN_CODE",
            "LIBELLE_OPER",
        ],
        dropna=False,
    )
    .agg(
        NB_ECRITURES=("TRN_REF_NO", "count"),
        DEBIT=("DEBIT", "sum"),
        CREDIT=("CREDIT", "sum"),
    )
    .reset_index()
)

print("\n===== PAR COMPTE / TRN_CODE =====")

print(
    analyse_compte
    .sort_values(
        [
            "ACCOUNT_NO",
            "TRN_CODE",
        ]
    )
    .to_string(index=False)
)


# ------------------------------------------------------------
# 13.2 Compte + description
# ------------------------------------------------------------

analyse_description = (
    df_analyse
    .groupby(
        [
            "ACCOUNT_NO",
            "TRN_CODE",
            "DESCRIPTION",
        ],
        dropna=False,
    )
    .agg(
        NB_ECRITURES=("TRN_REF_NO", "count"),
        DEBIT=("DEBIT", "sum"),
        CREDIT=("CREDIT", "sum"),
    )
    .reset_index()
)

print("\n===== PAR COMPTE / TRN_CODE / DESCRIPTION =====")

print(
    analyse_description
    .sort_values(
        [
            "ACCOUNT_NO",
            "TRN_CODE",
        ]
    )
    .to_string(index=False)
)


# ------------------------------------------------------------
# 13.3 Analyse uniquement des agences partenaires
# ------------------------------------------------------------

df_analyse_agences = df_analyse[
    df_analyse["CODE_AGENCE"].isin(
        [
            501,
            504,
            506,
            510,
            512,
            522,
            527,
        ]
    )
].copy()


analyse_agences = (
    df_analyse_agences
    .groupby(
        [
            "CODE_AGENCE",
            "ACCOUNT_NO",
            "TRN_CODE",
            "LIBELLE_OPER",
        ],
        dropna=False,
    )
    .agg(
        NB_ECRITURES=("TRN_REF_NO", "count"),
        DEBIT=("DEBIT", "sum"),
        CREDIT=("CREDIT", "sum"),
    )
    .reset_index()
)

print("\n===== AGENCES PARTENAIRES : COMPTE / TRN_CODE =====")

print(
    analyse_agences
    .sort_values(
        [
            "CODE_AGENCE",
            "ACCOUNT_NO",
            "TRN_CODE",
        ]
    )
    .to_string(index=False)
)

# ============================================================
# 14. COMPARAISON CIBLEE PARTENAIRE / FLEX
# ============================================================

print("\n")
print("=" * 120)
print("          COMPARAISON CIBLEE PARTENAIRE / FLEX")
print("=" * 120)


# ------------------------------------------------------------
# FLEX : uniquement CAISSE TRANSFERT
# ------------------------------------------------------------

df_flex_caisse = df_flex[
    df_flex["ACCOUNT_NO"].astype(str).str.strip()
    == "101100000001"
].copy()


# ------------------------------------------------------------
# Agences du partenaire
# ------------------------------------------------------------

agences = sorted(
    df_partenaire_agence["CODE_AGENCE"]
    .dropna()
    .astype(int)
    .unique()
)


resultats_cibles = []


for code_agence in agences:

    # ========================================================
    # PARTENAIRE
    # ========================================================

    ligne_partenaire = df_partenaire_agence[
        df_partenaire_agence["CODE_AGENCE"] == code_agence
    ]

    if ligne_partenaire.empty:
        continue

    ligne_partenaire = ligne_partenaire.iloc[0]

    debit_partenaire = float(
        ligne_partenaire["DEBIT_PARTENAIRE"]
    )

    credit_partenaire = float(
        ligne_partenaire["CREDIT_PARTENAIRE"]
    )

    nb_paiement = int(
        ligne_partenaire["NB_PAIEMENT_WU"]
    )

    nb_envoi = int(
        ligne_partenaire["NB_ENVOI_WU"]
    )


    # ========================================================
    # FLEX
    # ========================================================

    flex_agence = df_flex_caisse[
        df_flex_caisse["CODE_AGENCE"] == code_agence
    ].copy()


    # Crédit Flex = opérations TRN_CODE 122
    credit_flex = flex_agence[
        flex_agence["TRN_CODE"].astype(str) == "122"
    ]["CREDIT"].sum()


    # Débit Flex = opérations TRN_CODE 121
    debit_flex = flex_agence[
        flex_agence["TRN_CODE"].astype(str) == "121"
    ]["DEBIT"].sum()


    # ========================================================
    # ECARTS
    # ========================================================

    ecart_debit = (
        debit_partenaire
        - debit_flex
    )

    ecart_credit = (
        credit_partenaire
        - credit_flex
    )


    # ========================================================
    # TAXE FLEX
    # ========================================================

    taxe_agence = df_taxe_agence[
        df_taxe_agence["CODE_AGENCE"] == code_agence
    ]

    if taxe_agence.empty:
        taxe = 0
    else:
        taxe = float(
            taxe_agence.iloc[0]["TAXE"]
        )


    resultats_cibles.append(
        {
            "CODE_AGENCE": code_agence,
            "NOM_AGENCE": ligne_partenaire["NOM_AGENCE"],

            "NB_PAIEMENT_WU": nb_paiement,
            "NB_ENVOI_WU": nb_envoi,

            "DEBIT_PARTENAIRE":
                debit_partenaire,

            "DEBIT_FLEX":
                debit_flex,

            "ECART_DEBIT":
                ecart_debit,

            "CREDIT_PARTENAIRE":
                credit_partenaire,

            "CREDIT_FLEX":
                credit_flex,

            "ECART_CREDIT":
                ecart_credit,

            "TAXE":
                taxe,
        }
    )


df_cible = pd.DataFrame(
    resultats_cibles
)


# ============================================================
# AFFICHAGE
# ============================================================

print(
    df_cible.to_string(
        index=False
    )
)


# ============================================================
# CONTROLE DES ECARTS
# ============================================================

print("\n")
print("=" * 120)
print("                    CONTROLE DES ECARTS")
print("=" * 120)

for _, ligne in df_cible.iterrows():

    print(
        f"Agence {int(ligne['CODE_AGENCE'])} "
        f"| {ligne['NOM_AGENCE']}"
    )

    print(
        f"  PAIEMENT WU : "
        f"{int(ligne['NB_PAIEMENT_WU'])}"
    )

    print(
        f"  ENVOI WU    : "
        f"{int(ligne['NB_ENVOI_WU'])}"
    )

    print(
        f"  DEBIT  : partenaire="
        f"{ligne['DEBIT_PARTENAIRE']:,.0f} "
        f"| Flex="
        f"{ligne['DEBIT_FLEX']:,.0f} "
        f"| écart="
        f"{ligne['ECART_DEBIT']:,.0f}"
    )

    print(
        f"  CREDIT : partenaire="
        f"{ligne['CREDIT_PARTENAIRE']:,.0f} "
        f"| Flex="
        f"{ligne['CREDIT_FLEX']:,.0f} "
        f"| écart="
        f"{ligne['ECART_CREDIT']:,.0f}"
    )

    print(
        f"  TAXE FLEX : "
        f"{ligne['TAXE']:,.0f}"
    )

    print("-" * 100)

# ============================================================
# 15. DETAIL DES PAIEMENTS WU
# ============================================================

print("\n")
print("=" * 120)
print("              DETAIL PAIEMENT WU - PARTENAIRE")
print("=" * 120)

df_paiement = df_partenaire[
    df_partenaire["TYPE_OPERATION"] == "PAIEMENT WU"
].copy()

colonnes_partenaire = [
    "CODE_AGENCE",
    "NOM_AGENCE",
    "TYPE_OPERATION",
    "CREDIT",
    "DEBIT",
    "REMARQUES",
]

colonnes_partenaire = [
    c for c in colonnes_partenaire
    if c in df_paiement.columns
]

print(
    df_paiement[
        colonnes_partenaire
    ].to_string(index=False)
)


print("\n")
print("=" * 120)
print("              DETAIL CREDIT FLEX - TRN_CODE 122")
print("=" * 120)

df_paiement_flex = df_flex[
    (
        df_flex["ACCOUNT_NO"].astype(str).str.strip()
        == "101100000001"
    )
    &
    (
        df_flex["TRN_CODE"].astype(str)
        == "122"
    )
    &
    (
        df_flex["CREDIT"] > 0
    )
].copy()

colonnes_flex = [
    "CODE_AGENCE",
    "LIBELLE_AGENCE",
    "CREDIT",
    "DATE_SAISIE",
    "DATE_VALEUR",
    "TRN_REF_NO",
    "TRN_CODE",
    "LIBELLE_OPER",
    "DESCRIPTION",
]

colonnes_flex = [
    c for c in colonnes_flex
    if c in df_paiement_flex.columns
]

print(
    df_paiement_flex[
        colonnes_flex
    ]
    .sort_values(
        [
            "CODE_AGENCE",
            "CREDIT",
        ]
    )
    .to_string(index=False)
)