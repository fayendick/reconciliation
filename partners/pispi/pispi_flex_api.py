# ============================================================
# PI/SPI — SERVICE FLEX ORACLE
# ============================================================

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


# ============================================================
# CONFIGURATION
# ============================================================

PARTENAIRE = "PISPI"

rt = bootstrap_flex_service(
    PARTENAIRE,
    title="PI/SPI Flex API",
    log_prefix="pispi_flex_api",
)

app = rt.app
TABLES = rt.tables


# ============================================================
# EXTRACTION ORACLE PI/SPI
# ============================================================

def get_pispi_flex(
    date_debut: str,
    date_fin: str,
) -> pd.DataFrame:

    # ========================================================
    # REQUETE PI/SPI — CONSERVEE TELLE QUE FOURNIE
    # ========================================================
    sql = text("""
    with PI_ENTRANT_SORTANT as (

/* =======================
   PI SORTANT
   ======================= */
SELECT 
    a.AC_ENTRY_SR_NO,
    i.MONTANTTRANSFERT       AS MONTANT,
    a.AC_NO        AS NUMCPT,
    i.TXNREFNO,
    a.RELATED_CUSTOMER,
    a.DRCR_IND,
    a.TRN_DT,
    ( SELECT 'Transfert PI SPI' || ' vers ' || NOMPAYE
        FROM CFSFCUBS145.STTM_CUST_ACCOUNT
       WHERE CUST_AC_NO = i.NCPDEBITEUR
    ) AS DESC_OPS,
    i.FEES         AS FRAIS
FROM OMB_SN.PITRANSACTION i, CFSFCUBS145.ACVW_ALL_AC_ENTRIES  a
WHERE i.TXNREFNO = a.EXTERNAL_REF_NO
  AND a.AC_NO = i.NCPDEBITEUR
  AND i.SENS = 'D'
  and a.TRN_DT between to_date(:date_debut, 'DD/MM/YYYY')
                   and to_date(:date_fin, 'DD/MM/YYYY')
  
  
  
  
 
UNION ALL
 
/* =======================
   PI ENTRANT
   ======================= */
SELECT 
    a.AC_ENTRY_SR_NO,
    i.MONTANTTRANSFERT       AS MONTANT,
    a.AC_NO        AS NUMCPT,
    i.TXNREFNO,
    a.RELATED_CUSTOMER ,
    a.DRCR_IND,
    a.TRN_DT,
    ( SELECT 'Transfert PI SPI' || ' reçu de ' || NOMPAYEUR
        FROM CFSFCUBS145.STTM_CUST_ACCOUNT
       WHERE CUST_AC_NO = i.NCPDEBITEUR
    ) AS DESC_OPS,
    i.FEES         AS FRAIS
FROM OMB_SN.PITRANSACTION i, CFSFCUBS145.ACVW_ALL_AC_ENTRIES a
WHERE i.TXNREFNO = a.EXTERNAL_REF_NO
  AND a.AC_NO = i.COMPTEPAYE
  AND i.SENS = 'C'
  and a.TRN_DT between to_date(:date_debut, 'DD/MM/YYYY')
                 and to_date(:date_fin, 'DD/MM/YYYY')
 
 

  
  ),
  
  
  KYC AS ( 

    SELECT 

        sc.CUSTOMER_NO,

        cat.CUST_CAT_DESC ,
        DECODE(sc.CUSTOMER_TYPE, 'C', 'ENTREPRISE', 'I', 'PARTICULIER')  as CUSTOMER_TYPE

    FROM CFSFCUBS145.STTM_CUSTOMER sc

    LEFT JOIN CFSFCUBS145.STTM_CUSTOMER_CAT cat 
    
        ON cat.CUST_CAT = sc.CUSTOMER_CATEGORY
     LEFT JOIN CFSFCUBS145.STTM_CUST_PERSONAL p 
        ON sc.CUSTOMER_NO = p.CUSTOMER_NO
    
    ),
    PI_FINAL as (
    select 

    kyc.CUST_CAT_DESC,
    kyc.CUSTOMER_TYPE ,
    pi.* 

    from PI_ENTRANT_SORTANT  pi

    left join KYC kyc on kyc.CUSTOMER_NO= pi.RELATED_CUSTOMER
    ),

    ecriture as (
    SELECT
    A.AC_NO,
    cpt.AC_DESC as NOM_CLIENT,
    A.TRN_REF_NO,
    A.EXTERNAL_REF_NO as REF_PI,
    A.LCY_AMOUNT,
    A.TRN_DT,
    A.SAVE_TIMESTAMP
    FROM
    CFSFCUBS145.ACVW_ALL_AC_ENTRIES A
    JOIN CFSFCUBS145.STTM_CUST_ACCOUNT CPT ON A.AC_NO = CPT.CUST_AC_NO
    JOIN CFSFCUBS145.STTM_ACCOUNT_CLASS CL ON CPT.ACCOUNT_CLASS = CL.ACCOUNT_CLASS
    WHERE
    CL.ACCOUNT_CODE IN ('251' ,'253')
    ),

    STATUT_TRANSACTION_PI as ( 
    select ni.TXNREFNO,
    e.LCY_AMOUNT,
    e.NOM_CLIENT,
    e.SAVE_TIMESTAMP,
    CASE WHEN e.LCY_AMOUNT IS NOT NULL THEN 'SUCCESS'ELSE 'FAIL' END AS STATUT
    from  PI_FINAL ni
    LEFT JOIN ecriture e  on ni.TXNREFNO = e.REF_PI
    AND e.AC_NO = ni.NUMCPT

    )
    select i.* ,s.NOM_CLIENT,s.STATUT as "STATUT_TRANSACTION_PI",s.SAVE_TIMESTAMP
    from PI_FINAL i 
    left join STATUT_TRANSACTION_PI s on i.TXNREFNO = s.TXNREFNO
        
    
    """)

    # ========================================================
    # FENETRE DE RECONCILIATION
    # Date debut : 00:00:00
    # Date fin   : 00:00:00
    #
    # La requete SQL ci-dessus n'est pas modifiee.
    # Le filtrage de la periode est applique apres extraction
    # sur SAVE_TIMESTAMP, date/heure retenue pour la reconciliation.
    # ========================================================

    date_debut_ts = pd.Timestamp(date_debut).normalize()
    date_fin_ts = pd.Timestamp(date_fin).normalize()
    
    test_sql = """
    SELECT
        SYS_CONTEXT('USERENV', 'SESSION_USER') AS SESSION_USER,
        SYS_CONTEXT('USERENV', 'CURRENT_SCHEMA') AS CURRENT_SCHEMA
    FROM DUAL
    """

    test_df = oracle_query(rt, test_sql, {})
    print(test_df)

    df = oracle_query(
        rt,
        sql,
        {
            "date_debut": date_debut,
            "date_fin": date_fin,
        
        },
    )

    if df is None or df.empty:
        return df

    # ========================================================
    # COLONNES DE RECONCILIATION
    # ========================================================

    df["SAVE_TIMESTAMP"] = pd.to_datetime(
        df["SAVE_TIMESTAMP"],
        errors="coerce",
    )

    df = df[
        (df["SAVE_TIMESTAMP"] >= date_debut_ts)
        & (df["SAVE_TIMESTAMP"] < date_fin_ts)
    ].copy()

    # ========================================================
    # MAPPING METIER PI/SPI
    #
    # Rapprochement uniquement sur :
    # - NUMERO_COMPTE
    # - MONTANT_COMPARAISON
    # - DATE_HEURE
    #
    # TXNREFNO / CODE_TRANSACTION reste informatif.
    # ========================================================

    df["NUMERO_COMPTE"] = (
        df["NUMCPT"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["MONTANT_COMPARAISON"] = pd.to_numeric(
        df["MONTANT"],
        errors="coerce",
    )

    df["DATE_VALEUR"] = df["SAVE_TIMESTAMP"]
    df["DATE_HEURE"] = df["SAVE_TIMESTAMP"]

    df["CODE_TRANSACTION"] = (
        df["TXNREFNO"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["CODE_TRANSACTION_OPERATEUR"] = df["CODE_TRANSACTION"]

    df["SENS"] = (
        df["DRCR_IND"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df["TYPE_TRANSACTION"] = df["SENS"].map(
        {
            "C": "W2B",
            "D": "B2W",
        }
    )

    df["MOUVEMENT_CREDIT"] = 0.0
    df["MOUVEMENT_DEBIT"] = 0.0

    df.loc[df["SENS"] == "C", "MOUVEMENT_CREDIT"] = df.loc[
        df["SENS"] == "C", "MONTANT_COMPARAISON"
    ]

    df.loc[df["SENS"] == "D", "MOUVEMENT_DEBIT"] = df.loc[
        df["SENS"] == "D", "MONTANT_COMPARAISON"
    ]

    return df


# ============================================================
# API PRINCIPALE
# ============================================================

@app.get("/pispi-flex")
def export_pispi_flex(
    date_debut: str,
    date_fin: str,
    format: str = Query(
        "excel",
        description="excel ou json",
    ),
):

    df = get_pispi_flex(
        date_debut,
        date_fin,
    )

    if df is None or df.empty:
        raise HTTPException(
            status_code=404,
            detail="Aucune transaction PI/SPI trouvée.",
        )

    if "SENS" not in df.columns:
        raise HTTPException(
            status_code=500,
            detail=(
                "Colonne SENS absente. "
                f"Colonnes disponibles : {list(df.columns)}"
            ),
        )

    # ========================================================
    # NORMALISATION DU SENS
    # ========================================================

    df["SENS"] = (
        df["SENS"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ========================================================
    # PI/SPI :
    #
    # CREDIT → W2B
    # DEBIT  → B2W
    # ========================================================

    w2b = df[
        df["SENS"] == "C"
    ].copy()

    b2w = df[
        df["SENS"] == "D"
    ].copy()

    # ========================================================
    # TYPE TRANSACTION
    # ========================================================
    # Le moteur de réconciliation travaille avec :
    # CREDIT → W2B
    # DEBIT  → B2W

    df["TYPE_TRANSACTION"] = df["SENS"].map(
        {
            "C": "W2B",
            "D": "B2W",
        }
    )
    
    # ========================================================
    # MOUVEMENTS COMPTABLES
    # ========================================================
    # PI/SPI :
    # C = Crédit -> montant dans MOUVEMENT_CREDIT
    # D = Débit  -> montant dans MOUVEMENT_DEBIT

    

    w2b["TYPE_TRANSACTION"] = "W2B"
    b2w["TYPE_TRANSACTION"] = "B2W"

    # ========================================================
    # SAUVEGARDE
    # ========================================================

    from common.flex_common import save_split_w2b_b2w

    # Sauvegarde de la table complète
    ecrire_table(
        df,
        TABLES["flex"],
        rt.sqlite,
        log_prefix=rt.log_prefix,
    )

    # PI/SPI :
    # C = Crédit = W2B
    # D = Débit   = B2W
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
            "PISPI_FLEX": df,
            "PISPI_FLEX_W2B": w2b,
            "PISPI_FLEX_B2W": b2w,
        },
        filename="PISPI_FLEX.xlsx",
        format=format,
    )


# ============================================================
# ROUTES SQLITE
# ============================================================

register_db_routes(
    rt,
    [
        ("/db/pispi", TABLES["flex"]),
        ("/db/pispi-w2b", TABLES["flex_w2b"]),
        ("/db/pispi-b2w", TABLES["flex_b2w"]),
    ],
)
