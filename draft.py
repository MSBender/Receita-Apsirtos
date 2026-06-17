"""
draft.py — Autosave / recuperação de rascunho do formulário.

Salva os campos de texto preenchidos num arquivo JSON local. Se o app
reiniciar no meio do uso (reboot por memória, queda de conexão, app dormindo),
o conteúdo pode ser recuperado em vez de começar tudo do zero.

Stdlib apenas — nenhuma dependência nova.
"""

import os
import json
import tempfile
from datetime import datetime

_DRAFT_PATH = os.path.join(tempfile.gettempdir(), "receita_isabelle_draft.json")


def save_draft(fields: dict) -> None:
    """Grava o rascunho de forma atômica. Nunca lança exceção."""
    try:
        payload = {
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "fields": fields,
        }
        tmp = _DRAFT_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        os.replace(tmp, _DRAFT_PATH)
    except Exception:
        pass  # autosave jamais deve quebrar o app


def load_draft() -> dict:
    try:
        with open(_DRAFT_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def draft_info() -> dict:
    """Retorna {'when': str, 'fields': dict} se houver rascunho com conteúdo."""
    data = load_draft()
    fields = data.get("fields") if data else None
    if not fields:
        return {}
    if not any(str(v).strip() for v in fields.values()):
        return {}
    try:
        when = datetime.fromisoformat(data["saved_at"]).strftime("%d/%m/%Y às %H:%M")
    except Exception:
        when = "?"
    return {"when": when, "fields": fields}


def clear_draft() -> None:
    try:
        os.remove(_DRAFT_PATH)
    except Exception:
        pass
