# ============================================================
# HTTP EXPORT — Excel optionnel vs JSON léger
# ------------------------------------------------------------
# Module_FED / reconc.py /charger n'utilisent que le statut HTTP
# et éventuellement quelques en-têtes : le corps Excel est jeté.
# Passer format=json évite openpyxl + sérialisation coûteuse.
# Les appels manuels (navigateur, Streamlit direct) gardent
# format=excel par défaut.
# ============================================================

from __future__ import annotations

import io
from typing import Any, Dict, Mapping, Optional

import pandas as pd
from fastapi.responses import JSONResponse, StreamingResponse


def wants_excel(fmt: Optional[str]) -> bool:
    return str(fmt or "excel").strip().lower() not in {"json", "none", "sqlite"}


def respond_sheets(
    sheets: Dict[str, pd.DataFrame],
    *,
    filename: str,
    format: str = "excel",
    json_payload: Optional[Dict[str, Any]] = None,
    headers: Optional[Mapping[str, str]] = None,
):
    """Répond en Excel (défaut) ou JSON {status, counts} si format=json."""
    hdrs = {str(k): str(v) for k, v in dict(headers or {}).items()}

    if wants_excel(format):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, sheet_name=str(name)[:31], index=False)
        output.seek(0)
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                **hdrs,
                "Content-Disposition": f'attachment; filename="{filename}"',
            },
        )

    payload = json_payload or {
        "status": "ok",
        "format": "json",
        "filename": filename,
        "sheets": {name: int(len(df)) for name, df in sheets.items()},
        "total_rows": int(sum(len(df) for df in sheets.values())),
    }
    return JSONResponse(content=payload, headers=hdrs)
