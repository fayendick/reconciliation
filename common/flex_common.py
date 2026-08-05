# ============================================================
# FLEX COMMON — boilerplate partagé par tous les *_flex_api.py
# ------------------------------------------------------------
# Chaque service garde UNIQUEMENT son SQL Oracle spécifique.
# Ici : bootstrap FastAPI, engines Oracle/SQLite, health,
# routes /db/*, lecture Oracle, sauvegarde + export Excel/JSON.
# ============================================================

from __future__ import annotations

from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from sqlalchemy import text
from sqlalchemy.engine import Engine

from config import DB_PATH, get_partenaire, make_oracle_engine, make_sqlite_engine
from common.http_export import respond_sheets
from common.sqlite_io import ecrire_avec_vues_w2b_b2w, ecrire_table, lire_table_json


class FlexRuntime:
    """Contexte d'un service Flex (1 process = 1 app + 2 engines)."""

    __slots__ = ("app", "partenaire", "tables", "oracle", "sqlite", "log_prefix")

    def __init__(
        self,
        app: FastAPI,
        partenaire: str,
        tables: dict,
        oracle: Engine,
        sqlite: Engine,
        log_prefix: str,
    ):
        self.app = app
        self.partenaire = partenaire
        self.tables = tables
        self.oracle = oracle
        self.sqlite = sqlite
        self.log_prefix = log_prefix


def bootstrap_flex_service(
    partenaire: str,
    *,
    title: str,
    log_prefix: Optional[str] = None,
) -> FlexRuntime:
    """Crée l'app FastAPI, les engines, et enregistre /health."""
    cfg = get_partenaire(partenaire)
    tables = cfg["tables"]
    prefix = log_prefix or f"{partenaire.lower()}_flex"
    app = FastAPI(title=title)
    oracle = make_oracle_engine()
    sqlite = make_sqlite_engine()

    print(f"[{prefix}] Base SQLite utilisée : {DB_PATH}")

    @app.get("/health")
    def health():
        return {
            "status": "ok",
            "db_path": DB_PATH,
            "partenaire": partenaire,
            "table": tables.get("flex"),
        }

    return FlexRuntime(app, partenaire, tables, oracle, sqlite, prefix)


def register_db_routes(
    runtime: FlexRuntime,
    routes: Sequence[Tuple[str, str]],
) -> None:
    """
    Enregistre des GET /db/... qui lisent une table/vue SQLite.
    routes = [("/db/wave", "WAVE_FLEX"), ...]
    """
    for path, table_name in routes:
        # défauts figés pour éviter la capture tardive de la boucle
        def _make(path_=path, table_=table_name):
            @runtime.app.get(path_)
            def _handler(
                limit: int = Query(None),
                offset: int = Query(0),
                _table: str = table_,
            ):
                return lire_table_json(
                    runtime.sqlite, _table, limit=limit, offset=offset
                )

            return _handler

        _make()


def oracle_query(
    runtime: FlexRuntime,
    sql,
    params: Optional[Mapping[str, Any]] = None,
) -> pd.DataFrame:
    """Exécute un SQL Oracle et normalise les noms de colonnes."""
    try:
        with runtime.oracle.connect() as conn:
            df = pd.read_sql(
                sql if hasattr(sql, "text") else text(str(sql)),
                conn,
                params=dict(params or {}),
            )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Erreur Oracle/SQL: {str(e)}")

    df.columns = df.columns.astype(str).str.strip().str.upper()
    return df


def save_split_w2b_b2w(
    runtime: FlexRuntime,
    df: pd.DataFrame,
    *,
    sheets: Dict[str, pd.DataFrame],
    filename: str,
    format: str = "excel",
    mode: str = "contains",
    type_col: str = "TYPE_TRANSACTION",
    headers: Optional[Mapping[str, str]] = None,
    json_payload: Optional[Dict[str, Any]] = None,
):
    """1 écriture flex + vues W2B/B2W, puis export Excel ou JSON."""
    tables = runtime.tables
    try:
        ecrire_avec_vues_w2b_b2w(
            df,
            tables["flex"],
            tables["flex_w2b"],
            tables["flex_b2w"],
            runtime.sqlite,
            type_col=type_col,
            mode=mode,
            log_prefix=runtime.log_prefix,
        )
        return respond_sheets(
            sheets,
            filename=filename,
            format=format,
            headers=headers,
            json_payload=json_payload,
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur export: {str(e)}")


def save_single_flex(
    runtime: FlexRuntime,
    df: pd.DataFrame,
    *,
    sheet_name: str,
    filename: str,
    format: str = "excel",
    headers: Optional[Mapping[str, str]] = None,
    json_payload: Optional[Dict[str, Any]] = None,
    table_key: str = "flex",
):
    """Écriture d'une seule table flex (mode agence) + export."""
    try:
        ecrire_table(
            df,
            runtime.tables[table_key],
            runtime.sqlite,
            log_prefix=runtime.log_prefix,
        )
        return respond_sheets(
            {sheet_name: df},
            filename=filename,
            format=format,
            headers=headers,
            json_payload=json_payload,
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur export: {str(e)}")


def split_by_type_contains(df: pd.DataFrame, col: str = "TYPE_TRANSACTION"):
    series = df[col].fillna("").astype(str).str.upper()
    return (
        df[series.str.contains("W2B")].copy(),
        df[series.str.contains("B2W")].copy(),
    )


def split_by_type_exact(df: pd.DataFrame, col: str = "TYPE_TRANSACTION"):
    series = df[col].fillna("").astype(str).str.upper().str.strip()
    return (
        df[series == "W2B"].copy(),
        df[series == "B2W"].copy(),
    )


def split_by_dc(df: pd.DataFrame, col: str = "TYPE_TRANSACTION"):
    series = df[col].fillna("").astype(str).str.upper().str.strip()
    return (
        df[series == "D"].copy(),
        df[series == "C"].copy(),
    )
