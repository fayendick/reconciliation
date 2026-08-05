# ============================================================
# MOUNT PARTNERS — une API, N modules partenaires
# ------------------------------------------------------------
# - mount_partner_apps() : expose chaque app sous /svc/<nom>/...
# - call_partner_*()     : invoque Excel/Flex EN PROCESSUS
#   (appel direct du handler, sans HTTP loopback ni httpx).
# ============================================================

from __future__ import annotations

import importlib
import inspect
from io import BytesIO
from typing import Any, Dict, List, Mapping, Optional, Tuple

from fastapi import FastAPI, UploadFile
from fastapi.responses import JSONResponse, Response, StreamingResponse
from starlette.routing import Route

# (préfixe mount, module Python exposant `app`)
PARTNER_MOUNTS: List[Tuple[str, str]] = [
    ("/svc/wave-excel", "partners.wave.app_wave"),
    ("/svc/wave-flex", "partners.wave.wave_inter_flex_api"),
    ("/svc/orange-excel", "partners.orange.app_orange"),
    ("/svc/orange-flex", "partners.orange.orange_flex_api"),
    ("/svc/wizz-excel", "partners.wizz.app_wizz"),
    ("/svc/wizz-flex", "partners.wizz.wizz_flex_api"),
    ("/svc/wave-agence-excel", "partners.wave_agence.wave_app_agence"),
    ("/svc/wave-agence-flex", "partners.wave_agence.wave_agence_flex_api"),
    ("/svc/ria-agence-excel", "partners.ria_agence.ria_agence_app"),
    ("/svc/ria-agence-flex", "partners.ria_agence.RIA_agence_flex"),
    ("/svc/orange-ussd-excel", "partners.orange_ussd.app_orange_ussd"),
    ("/svc/orange-ussd-flex", "partners.orange_ussd.orange_ussd_flex_api"),
]

# Cibles pour /charger (module + chemin interne de la sous-app)
PARTNER_ASGI: Dict[str, Dict[str, str]] = {
    "WAVE": {
        "upload_module": "partners.wave.app_wave",
        "upload_path": "/process-excel",
        "flex_module": "partners.wave.wave_inter_flex_api",
        "flex_path": "/wave-inter-flex",
    },
    "ORANGE_AGENCE": {
        "upload_module": "partners.orange.app_orange",
        "upload_path": "/process-excel",
        "flex_module": "partners.orange.orange_flex_api",
        "flex_path": "/orange-flex",
    },
    "WIZZ": {
        "upload_module": "partners.wizz.app_wizz",
        "upload_path": "/process-excel",
        "flex_module": "partners.wizz.wizz_flex_api",
        "flex_path": "/wizz-flex",
    },
    "WAVE_AGENCE": {
        "upload_module": "partners.wave_agence.wave_app_agence",
        "upload_path": "/map-agences",
        "flex_module": "partners.wave_agence.wave_agence_flex_api",
        "flex_path": "/wave-agence-flex",
    },
    "RIA_AGENCE": {
        "upload_module": "partners.ria_agence.ria_agence_app",
        "upload_path": "/map-agences-ria",
        "flex_module": "partners.ria_agence.RIA_agence_flex",
        "flex_path": "/ria-agence-flex",
    },
    "ORANGE_USSD": {
        "upload_module": "partners.orange_ussd.app_orange_ussd",
        "upload_path": "/process-excel",
        "flex_module": "partners.orange_ussd.orange_ussd_flex_api",
        "flex_path": "/orange-ussd-flex",
    },
}


class PartnerCallResult:
    """Réponse minimaliste compatible avec l'ancien usage requests."""

    __slots__ = ("status_code", "text", "headers", "content")

    def __init__(self, status_code: int, text: str = "", headers: Optional[Mapping[str, str]] = None, content: bytes = b""):
        self.status_code = status_code
        self.text = text
        self.headers = {str(k): str(v) for k, v in dict(headers or {}).items()}
        self.content = content


def mount_partner_apps(app: FastAPI) -> list[str]:
    """Importe et monte chaque app partenaire. Retourne les préfixes OK."""
    mounted: list[str] = []
    for prefix, module_name in PARTNER_MOUNTS:
        try:
            mod = importlib.import_module(module_name)
            sub_app = getattr(mod, "app", None)
            if sub_app is None:
                print(f"[mount_partners] SKIP {module_name} : pas d'attribut app")
                continue
            app.mount(prefix, sub_app)
            mounted.append(prefix)
            print(f"[mount_partners] {module_name} -> {prefix}")
        except Exception as e:
            print(f"[mount_partners] ERREUR montage {module_name} sur {prefix} : {e}")
    return mounted


def _load_app(module_name: str):
    mod = importlib.import_module(module_name)
    sub_app = getattr(mod, "app", None)
    if sub_app is None:
        raise RuntimeError(f"Module {module_name} n'expose pas app")
    return sub_app


def _find_endpoint(sub_app: FastAPI, path: str):
    for route in sub_app.routes:
        if isinstance(route, Route) and route.path == path:
            return route.endpoint
    raise RuntimeError(f"Route {path} introuvable sur {sub_app.title if hasattr(sub_app, 'title') else sub_app}")


def _normalize_result(result: Any) -> PartnerCallResult:
    from fastapi import HTTPException

    if isinstance(result, HTTPException):
        return PartnerCallResult(result.status_code, str(result.detail))

    if isinstance(result, JSONResponse):
        body = result.body
        if isinstance(body, memoryview):
            body = body.tobytes()
        text = body.decode("utf-8", errors="replace") if isinstance(body, (bytes, bytearray)) else str(body)
        return PartnerCallResult(result.status_code, text, result.headers, body if isinstance(body, (bytes, bytearray)) else text.encode())

    if isinstance(result, StreamingResponse):
        chunks = []
        try:
            for chunk in result.body_iterator:
                if isinstance(chunk, memoryview):
                    chunk = chunk.tobytes()
                if isinstance(chunk, str):
                    chunk = chunk.encode()
                chunks.append(chunk)
        except TypeError:
            # async iterator — rare pour format=json
            pass
        content = b"".join(chunks)
        return PartnerCallResult(result.status_code, content.decode("utf-8", errors="replace"), result.headers, content)

    if isinstance(result, Response):
        body = result.body or b""
        if isinstance(body, memoryview):
            body = body.tobytes()
        text = body.decode("utf-8", errors="replace") if isinstance(body, (bytes, bytearray)) else str(body)
        return PartnerCallResult(result.status_code, text, result.headers, body if isinstance(body, (bytes, bytearray)) else b"")

    if isinstance(result, dict):
        import json
        text = json.dumps(result, ensure_ascii=False, default=str)
        return PartnerCallResult(200, text, {"content-type": "application/json"}, text.encode())

    return PartnerCallResult(200, str(result))


async def _ainvoke(endpoint, kwargs: dict) -> PartnerCallResult:
    from fastapi import HTTPException

    try:
        if inspect.iscoroutinefunction(endpoint):
            result = await endpoint(**kwargs)
        else:
            result = endpoint(**kwargs)
        return _normalize_result(result)
    except HTTPException as e:
        return PartnerCallResult(e.status_code, str(e.detail))


async def call_partner_upload(
    partenaire: str,
    files_payload: list,
    params: Optional[Mapping[str, Any]] = None,
) -> PartnerCallResult:
    """POST Excel in-process. files_payload = [(field, (name, bytes, ctype)), ...]"""
    spec = PARTNER_ASGI.get(partenaire)
    if not spec:
        raise KeyError(f"Pas de cible ASGI upload pour {partenaire}")

    sub_app = _load_app(spec["upload_module"])
    endpoint = _find_endpoint(sub_app, spec["upload_path"])

    uploads: List[UploadFile] = []
    for _field, (filename, content, content_type) in files_payload:
        uploads.append(
            UploadFile(
                filename=filename or "upload.xlsx",
                file=BytesIO(content),
                headers={"content-type": content_type or "application/octet-stream"},
            )
        )

    kwargs: Dict[str, Any] = {"files": uploads}
    params = dict(params or {})
    if "format" in params:
        kwargs["format"] = params["format"]

    return await _ainvoke(endpoint, kwargs)


async def call_partner_flex(
    partenaire: str,
    params: Optional[Mapping[str, Any]] = None,
) -> PartnerCallResult:
    """GET Flex in-process."""
    spec = PARTNER_ASGI.get(partenaire)
    if not spec:
        raise KeyError(f"Pas de cible ASGI flex pour {partenaire}")

    sub_app = _load_app(spec["flex_module"])
    endpoint = _find_endpoint(sub_app, spec["flex_path"])

    params = dict(params or {})
    # Ne passer que les kwargs acceptés par le handler
    sig = inspect.signature(endpoint)
    kwargs = {k: v for k, v in params.items() if k in sig.parameters}
    # date_debut / date_fin souvent requis positionnellement via Query
    for required in ("date_debut", "date_fin"):
        if required in sig.parameters and required not in kwargs and required in params:
            kwargs[required] = params[required]

    return await _ainvoke(endpoint, kwargs)
