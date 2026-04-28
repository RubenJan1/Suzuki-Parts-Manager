"""Persistentielaag voor Krat Beheer — kratten worden als JSON opgeslagen."""

from __future__ import annotations
import json
import uuid
from datetime import datetime
from pathlib import Path
from utils.paths import appdata_root


def kratten_dir() -> Path:
    p = appdata_root() / "kratten"
    p.mkdir(parents=True, exist_ok=True)
    return p


def list_kratten() -> list[dict]:
    kratten = []
    for f in sorted(kratten_dir().glob("krat_*.json")):
        try:
            with open(f, "r", encoding="utf-8") as fh:
                kratten.append(json.load(fh))
        except Exception:
            pass
    kratten.sort(key=lambda k: k.get("aangemaakt", ""), reverse=True)
    return kratten


def load_krat(krat_id: str) -> dict | None:
    p = kratten_dir() / f"krat_{krat_id}.json"
    if not p.exists():
        return None
    with open(p, "r", encoding="utf-8") as fh:
        return json.load(fh)


def save_krat(krat: dict) -> None:
    p = kratten_dir() / f"krat_{krat['krat_id']}.json"
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(krat, fh, ensure_ascii=False, indent=2)


def delete_krat(krat_id: str) -> None:
    p = kratten_dir() / f"krat_{krat_id}.json"
    if p.exists():
        p.unlink()


def new_krat(naam: str, locatie: str) -> dict:
    return {
        "krat_id":         str(uuid.uuid4()),
        "naam":            naam,
        "locatie":         locatie,
        "aangemaakt":      datetime.now().isoformat(),
        "status":          "inventarisatie",
        "beprijzing_index": 0,
        "artikelen":       [],
    }
