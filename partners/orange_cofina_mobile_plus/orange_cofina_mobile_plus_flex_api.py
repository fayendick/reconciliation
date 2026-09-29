from fastapi import HTTPException, Query

import pandas as pd
from sqlalchemy import text

from common.flex_common import (
    bootstrap_flex_service,
    oracle_query,
    register_db_routes,
)
from common.sqlite_io import ecrire_table, creer_vues_split
from common.http_export import respond_sheets


PARTENAIRE = "ORANGE_COFINA_MOBILE_PLUS"

rt = bootstrap_flex_service(
    PARTENAIRE,
    title="Orange API / Cofina Mobile Plus Flex API",
    log_prefix="orange_cofina_mobile_plus_flex_api",
)

app = rt.app
TABLES = rt.tables


# ============================================================
# EXTRACTION ORACLE
# ============================================================
# Source confirmÃ©e : OMB_SN.ORANGETRANSACTION
# Rapprochement : RECIPIENTPHONENUMBER + AMOUNT + DATETRANS.
# Le champ TRANSACTIONDATE n'est PAS utilisÃ© pour la clÃ© de temps.
# ============================================================


def get_orange_cofina_mobile_plus_flex(
    date_debut: str,
    date_fin: str,
) -> pd.DataFrame:
    sql = text(
        """
        SELECT
            ID,
            DESCRIPTION,
            DT_CRE_ENREG,
            DT_MAJ_ENREG,
            ISDELETED,
            ISENABLED,
            NAME,
            AMOUNT,
            COMPTEDEBIT,
            COMPTEDEBITTYPE,
            DATETRANS,
            FEES,
            GUTRANSACTIONID,
            NUMCLIENT,
            NUMEROCLIENT,
            NUMERORESERVATION,
            NUMFACTURE,
            PARTNERTRANSACTIONID,
            PROVIDERNOM,
            RECIPIENTPHONENUMBER,
            SERVICE,
            SERVICEID,
            STATUSFINAL,
            STATUTINT,
            TRANSACTIONDATE,
            USERID,
            MSGSTATUS,
            TXNID,
            ENDTOENDID,
            RESPCODE,
            TXNREFNO
        FROM OMB_SN.ORANGETRANSACTION
        WHERE UPPER(TRIM(STATUSFINAL)) = 'SUCCESS'
          AND DATETRANS >= TO_DATE(:date_debut, 'DD/MM/YYYY')
          AND DATETRANS < TO_DATE(:date_fin, 'DD/MM/YYYY')
        ORDER BY DATETRANS
        """
    )

    return oracle_query(
        rt,
        sql,
        {
            "date_debut": date_debut,
            "date_fin": date_fin,
        },
    )


# ============================================================
# PREPARATION FLEX
# ============================================================


def _normaliser_service(valeur) -> str:
    if pd.isna(valeur):
        return ""
    texte = str(valeur).strip().upper()
    texte = " ".join(texte.split())
    return {
        "CASH IN": "W2B",
        "CASHIN": "W2B",
        "CASH-IN": "W2B",
        "CASH OUT": "B2W",
        "CASHOUT": "B2W",
        "CASH-OUT": "B2W",
    }.get(texte, texte)


def preparer_flex(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    required = {
        "AMOUNT",
        "DATETRANS",
        "RECIPIENTPHONENUMBER",
        "SERVICE",
        "STATUSFINAL",
    }
    missing = required - set(df.columns)
    if missing:
        raise HTTPException(
            status_code=500,
            detail=(
                f"Colonnes Oracle Orange API manquantes : {sorted(missing)}. "
                f"Colonnes disponibles : {list(df.columns)}"
            ),
        )

    df = df.copy()

    # O2C / approvisionnement n'a pas de téléphone destinataire et
    # ne fait pas partie du périmètre W2B/B2W. On l'exclut donc avant
    # le mapping du service.
    service_brut = df["SERVICE"].fillna("").astype(str).str.strip().str.upper()
    masque_o2c = service_brut.eq("O2C TRANSFER")
    masque_sans_telephone = (
        df["RECIPIENTPHONENUMBER"].isna()
        | df["RECIPIENTPHONENUMBER"].astype(str).str.strip().eq("")
    )
    df = df.loc[~masque_o2c & ~masque_sans_telephone].copy()

    if df.empty:
        return pd.DataFrame()

    df["AMOUNT"] = pd.to_numeric(df["AMOUNT"], errors="coerce").fillna(0)
    df["FEES"] = pd.to_numeric(df.get("FEES"), errors="coerce").fillna(0)
    df["DATETRANS"] = pd.to_datetime(df["DATETRANS"], errors="coerce")

    df["DATE_VALEUR"] = df["DATETRANS"]
    df["DATE_HEURE"] = df["DATETRANS"]

    # TÃ©lÃ©phone Flex confirmÃ© : RECIPIENTPHONENUMBER.
    df["NUMERO_COMPTE"] = (
        df["RECIPIENTPHONENUMBER"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # RÃ©fÃ©rence informative uniquement.
    reference = df.get("TXNREFNO", df.get("GUTRANSACTIONID", ""))
    df["CODE_TRANSACTION_OPERATEUR"] = (
        pd.Series(reference, index=df.index)
        .fillna("")
        .astype(str)
        .str.strip()
    )
    df["CODE_TRANSACTION"] = df["CODE_TRANSACTION_OPERATEUR"]

    df["TYPE_TRANSACTION"] = df["SERVICE"].map(_normaliser_service)

    inconnus = sorted(
        v for v in df["TYPE_TRANSACTION"].dropna().unique()
        if v not in {"W2B", "B2W"}
    )
    if inconnus:
        raise HTTPException(
            status_code=500,
            detail=(
                "Services Oracle Orange API non reconnus : "
                f"{inconnus}. VÃ©rifie le mapping mÃ©tier de SERVICE."
            ),
        )

    # W2B = crÃ©dit Flex ; B2W = dÃ©bit Flex.
    df["MOUVEMENT_CREDIT"] = 0.0
    df["MOUVEMENT_DEBIT"] = 0.0
    df.loc[df["TYPE_TRANSACTION"] == "W2B", "MOUVEMENT_CREDIT"] = df.loc[
        df["TYPE_TRANSACTION"] == "W2B", "AMOUNT"
    ]
    df.loc[df["TYPE_TRANSACTION"] == "B2W", "MOUVEMENT_DEBIT"] = df.loc[
        df["TYPE_TRANSACTION"] == "B2W", "AMOUNT"
    ]

    df["MONTANT"] = df["AMOUNT"]
    df["MONTANT_COMPARAISON"] = df["AMOUNT"]

    return df


# ============================================================
# API PRINCIPALE
# ============================================================

@app.get("/orange-cofina-mobile-plus-flex")
def export_orange_cofina_mobile_plus_flex(
    date_debut: str,
    date_fin: str,
    format: str = Query("excel", description="excel (dÃ©faut) ou json"),
):
    df = get_orange_cofina_mobile_plus_flex(date_debut, date_fin)

    if df.empty:
        raise HTTPException(
            status_code=404,
            detail="Aucune transaction Orange API / Cofina Mobile Plus trouvÃ©e sur la pÃ©riode.",
        )

    df = preparer_flex(df)

    w2b = df[df["TYPE_TRANSACTION"] == "W2B"].copy()
    b2w = df[df["TYPE_TRANSACTION"] == "B2W"].copy()

    print(
        f"[orange_cofina_mobile_plus_flex_api] Extraction : {len(df)} lignes "
        f"({len(w2b)} W2B / {len(b2w)} B2W)"
    )

    ecrire_table(
        df,
        TABLES["flex"],
        rt.sqlite,
        log_prefix=rt.log_prefix,
    )

    creer_vues_split(
        rt.sqlite,
        TABLES["flex"],
        TABLES["flex_w2b"],
        TABLES["flex_b2w"],
        "UPPER(TRIM(CAST(TYPE_TRANSACTION AS TEXT))) = 'W2B'",
        "UPPER(TRIM(CAST(TYPE_TRANSACTION AS TEXT))) = 'B2W'",
        log_prefix=rt.log_prefix,
    )

    return respond_sheets(
        {
            "ORANGE_COFINA_MOBILE_PLUS_FLEX": df,
            "ORANGE_COFINA_MOBILE_PLUS_FLEX_W2B": w2b,
            "ORANGE_COFINA_MOBILE_PLUS_FLEX_B2W": b2w,
        },
        filename="ORANGE_COFINA_MOBILE_PLUS_FLEX.xlsx",
        format=format,
    )


# ============================================================
# ROUTES SQLITE
# ============================================================

register_db_routes(
    rt,
    [
        ("/db/orange-cofina-mobile-plus", TABLES["flex"]),
        ("/db/orange-cofina-mobile-plus-w2b", TABLES["flex_w2b"]),
        ("/db/orange-cofina-mobile-plus-b2w", TABLES["flex_b2w"]),
    ],
)
