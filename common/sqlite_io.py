# ============================================================
# SQLITE I/O — écriture optimisée + lecture JSON légère
# ------------------------------------------------------------
# - Une seule écriture physique pour la table complète, puis
#   vues W2B/B2W (évite 3× to_sql(replace) du même jeu).
# - chunksize pour les gros DataFrames.
# - Sérialisation via DataFrame.to_json (évite to_dict records).
# ============================================================

from __future__ import annotations

import json
from typing import Optional

import pandas as pd
from fastapi.responses import Response
from sqlalchemy import text

from config import SQLITE_CHUNKSIZE, SQLITE_DB_READ_LIMIT


def qident(name: str) -> str:
    """Identifiant SQLite quoté (supporte espaces / majuscules)."""
    return '"' + str(name).replace('"', '""') + '"'


def _drop_relation(conn, name: str) -> None:
    """Supprime une vue ou une table du même nom (SQLite refuse
    DROP VIEW sur une table et inversement)."""
    ident = qident(name)
    row = conn.execute(
        text("SELECT type FROM sqlite_master WHERE name = :n LIMIT 1"),
        {"n": name},
    ).fetchone()
    if not row:
        return
    kind = row[0]
    if kind == "view":
        conn.execute(text(f"DROP VIEW IF EXISTS {ident}"))
    else:
        conn.execute(text(f"DROP TABLE IF EXISTS {ident}"))


def ecrire_table(
    df: pd.DataFrame,
    table_name: str,
    engine,
    *,
    log_prefix: str = "sqlite_io",
    chunksize: Optional[int] = None,
) -> int:
    """Écrit (replace) une table avec chunksize. Retourne le nb de lignes."""
    cs = SQLITE_CHUNKSIZE if chunksize is None else chunksize
    kwargs = {
        "name": table_name,
        "con": engine,
        "if_exists": "replace",
        "index": False,
    }
    if cs and cs > 0 and len(df) > cs:
        kwargs["chunksize"] = cs

    # Si des vues pointaient encore vers d'anciens noms, on nettoie
    # d'abord la relation cible au cas où ce serait une vue.
    with engine.begin() as conn:
        _drop_relation(conn, table_name)

    df.to_sql(**kwargs)
    print(f"[{log_prefix}] Table '{table_name}' écrite : {len(df)} lignes (chunksize={cs or 'off'})")
    return int(len(df))


def creer_vues_split(
    engine,
    table_full: str,
    table_a: str,
    table_b: str,
    where_a: str,
    where_b: str,
    *,
    log_prefix: str = "sqlite_io",
) -> None:
    """Remplace table_a / table_b par des vues filtrées sur table_full."""
    with engine.begin() as conn:
        _drop_relation(conn, table_a)
        _drop_relation(conn, table_b)
        conn.execute(text(
            f"CREATE VIEW {qident(table_a)} AS "
            f"SELECT * FROM {qident(table_full)} WHERE {where_a}"
        ))
        conn.execute(text(
            f"CREATE VIEW {qident(table_b)} AS "
            f"SELECT * FROM {qident(table_full)} WHERE {where_b}"
        ))
    print(
        f"[{log_prefix}] Vues '{table_a}' / '{table_b}' "
        f"créées sur '{table_full}'"
    )


def ecrire_avec_vues_w2b_b2w(
    df: pd.DataFrame,
    table_full: str,
    table_w2b: str,
    table_b2w: str,
    engine,
    *,
    type_col: str = "TYPE_TRANSACTION",
    mode: str = "contains",
    log_prefix: str = "sqlite_io",
    chunksize: Optional[int] = None,
) -> int:
    """
    Une écriture + 2 vues.

    mode:
      - "contains" : TYPE contient W2B / B2W (Wave / Wizz Flex)
      - "exact"    : TYPE = 'W2B' / 'B2W' (Excel partenaire, Orange USSD)
      - "dc"       : TYPE = 'D' / 'C' (Orange Agence Flex : débit/crédit)
    """
    col = qident(type_col)
    expr = f"UPPER(TRIM(CAST({col} AS TEXT)))"

    if mode == "exact":
        where_w2b = f"{expr} = 'W2B'"
        where_b2w = f"{expr} = 'B2W'"
    elif mode == "dc":
        where_w2b = f"{expr} = 'D'"
        where_b2w = f"{expr} = 'C'"
    else:
        where_w2b = f"{expr} LIKE '%W2B%'"
        where_b2w = f"{expr} LIKE '%B2W%'"

    n = ecrire_table(df, table_full, engine, log_prefix=log_prefix, chunksize=chunksize)
    creer_vues_split(
        engine,
        table_full,
        table_w2b,
        table_b2w,
        where_w2b,
        where_b2w,
        log_prefix=log_prefix,
    )
    return n


def df_json_response(
    df: pd.DataFrame,
    *,
    limit: Optional[int] = None,
    offset: int = 0,
) -> Response:
    """
    Sérialise sans passer par to_dict(orient='records').
    limit=None utilise SQLITE_DB_READ_LIMIT (0 = illimité).
    """
    if limit is None:
        limit = SQLITE_DB_READ_LIMIT if SQLITE_DB_READ_LIMIT > 0 else None

    if offset or limit:
        fin = None if limit is None else offset + limit
        df = df.iloc[offset:fin]

    payload = df.to_json(orient="records", date_format="iso")
    return Response(content=payload, media_type="application/json")


def lire_table_json(
    engine,
    table_name: str,
    *,
    limit: Optional[int] = None,
    offset: int = 0,
) -> Response:
    try:
        df = pd.read_sql(f"SELECT * FROM {qident(table_name)}", engine)
    except Exception:
        return Response(content="[]", media_type="application/json")
    return df_json_response(df, limit=limit, offset=offset)


def df_to_records(df: pd.DataFrame) -> list:
    """Équivalent sûr de to_dict(records) via to_json (NaN/Inf OK)."""
    return json.loads(df.to_json(orient="records", date_format="iso"))
