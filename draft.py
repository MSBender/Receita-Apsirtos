"""
draft.py — Autosave / restauração de rascunho por sessão.

Cada sessão tem um id (guardado na URL pelo app). O rascunho é gravado num
arquivo JSON por id. Quando a página recarrega — inclusive numa reconexão de
WebSocket — o app lê o mesmo id e restaura tudo automaticamente, sem o usuário
perder nada nem precisar clicar.

Stdlib apenas — nenhuma dependência nova.
"""

import os
import re
import json
import tempfile
from datetime import datetime

_DIR = tempfile.gettempdir()


def _path(sid: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_-]", "", str(sid))[:40] or "default"
    return os.path.join(_DIR, f"receita_draft_{safe}.json")


def save_draft(sid: str, fields: dict) -> None:
    """Grava o rascunho de forma atômica. Nunca lança exceção."""
    try:
        payload = {
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "fields": fields,
        }
        p = _path(sid)
        tmp = p + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        os.replace(tmp, p)
    except Exception:
        pass


def load_draft(sid: str) -> dict:
    try:
        with open(_path(sid), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def clear_draft(sid: str) -> None:
    try:
        os.remove(_path(sid))
    except Exception:
        pass
