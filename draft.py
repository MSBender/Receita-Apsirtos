"""
draft.py — Autosave / restauração de rascunho por sessão (texto + imagens).

Cada sessão tem um id (guardado na URL pelo app). O rascunho de texto é gravado
num JSON por id; os prints enviados são gravados como JPEG numa pasta por id.
Quando a página recarrega — inclusive numa reconexão de WebSocket — o app lê o
mesmo id e restaura texto e imagens automaticamente, sem o usuário perder nada.

Stdlib apenas — nenhuma dependência nova.
"""

import os
import re
import json
import tempfile
from datetime import datetime

_DIR = tempfile.gettempdir()


def _safe(s: str, n: int = 40) -> str:
    return re.sub(r"[^A-Za-z0-9_-]", "", str(s))[:n] or "default"


def _path(sid: str) -> str:
    return os.path.join(_DIR, f"receita_draft_{_safe(sid)}.json")


def _img_dir(sid: str) -> str:
    return os.path.join(_DIR, f"receita_imgs_{_safe(sid)}")


# ── Rascunho de texto ────────────────────────────────────────────────────────
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


# ── Prints (imagens) ─────────────────────────────────────────────────────────
def save_image(sid: str, img_hash: str, jpeg_bytes: bytes) -> None:
    """Salva um print (JPEG) identificado pelo hash do conteúdo. Não duplica.
    Prefixo numérico preserva a ordem de envio."""
    try:
        d = _img_dir(sid)
        os.makedirs(d, exist_ok=True)
        h = _safe(img_hash)
        existentes = [fn for fn in os.listdir(d) if fn.endswith(".jpg")]
        for fn in existentes:
            if fn.endswith(f"_{h}.jpg"):
                return  # já salvo
        seq = len(existentes)
        with open(os.path.join(d, f"{seq:03d}_{h}.jpg"), "wb") as f:
            f.write(jpeg_bytes)
    except Exception:
        pass


def list_images(sid: str) -> list:
    """Retorna [(hash, jpeg_bytes), ...] na ordem de envio."""
    out = []
    try:
        d = _img_dir(sid)
        if os.path.isdir(d):
            for fn in sorted(os.listdir(d)):
                if fn.endswith(".jpg"):
                    h = fn[:-4].split("_", 1)[-1]
                    with open(os.path.join(d, fn), "rb") as f:
                        out.append((h, f.read()))
    except Exception:
        pass
    return out


def clear_images(sid: str) -> None:
    try:
        d = _img_dir(sid)
        if os.path.isdir(d):
            for fn in os.listdir(d):
                try:
                    os.remove(os.path.join(d, fn))
                except Exception:
                    pass
            os.rmdir(d)
    except Exception:
        pass
