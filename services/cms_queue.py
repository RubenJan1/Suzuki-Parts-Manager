"""
CMS Weekfactuur Queue
Slaat geleverde orderregels automatisch op na elke afboekrun (277 of 1322).
Op donderdag worden deze regels ingeladen in de Factuurmaker.
"""
from __future__ import annotations

import json
import uuid
from datetime import date
from pathlib import Path

_QUEUE_FILE = Path("data/cms_queue.json")


def _load() -> dict:
    try:
        if _QUEUE_FILE.exists():
            return json.loads(_QUEUE_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {"entries": []}


def _save(data: dict) -> None:
    _QUEUE_FILE.parent.mkdir(parents=True, exist_ok=True)
    _QUEUE_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def add_run(bron: str, regels: list) -> None:
    """
    Voeg een afboekrun toe aan de queue.
    bron: "277" of "1322"
    regels: list van dicts met keys: title, omschrijving, besteld, geleverd, prijs, factuurnummer
    """
    if not regels:
        return
    data = _load()
    data["entries"].append({
        "id": str(uuid.uuid4()),
        "datum": date.today().isoformat(),
        "bron": bron,
        "regels": regels,
        "verwerkt": False,
    })
    _save(data)


def get_pending(bron: str = None) -> list:
    """Geeft alle niet-verwerkte entries, optioneel gefilterd op bron."""
    data = _load()
    entries = [e for e in data.get("entries", []) if not e.get("verwerkt")]
    if bron:
        entries = [e for e in entries if e.get("bron") == bron]
    return entries


def has_pending(bron: str = None) -> bool:
    return bool(get_pending(bron))


def pending_counts() -> dict:
    """Geeft {"277": n_regels, "1322": n_regels} voor de banner."""
    result = {"277": 0, "1322": 0}
    for entry in get_pending():
        b = entry.get("bron", "")
        if b in result:
            result[b] += len(entry.get("regels", []))
    return result


def pending_factuurnummers(bron: str) -> list:
    """Geeft gesorteerde unieke CMS-ordernummers voor de banner-tekst."""
    seen: set = set()
    for entry in get_pending(bron):
        for r in entry.get("regels", []):
            fn = str(r.get("factuurnummer", "")).strip()
            if fn:
                seen.add(fn)
    return sorted(seen)


def mark_verwerkt(bron: str) -> None:
    """Markeer alle pending entries voor een bron als verwerkt."""
    data = _load()
    for entry in data.get("entries", []):
        if not entry.get("verwerkt") and entry.get("bron") == bron:
            entry["verwerkt"] = True
    _save(data)
