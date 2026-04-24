# services/superseded.py
"""
Superseded nummers opzoeken vanuit 'Superseded lijst.xls'.

Suzuki onderdeelnummers zijn gegroepeerd in families: hetzelfde onderdeel
kan onder meerdere nummers bekend zijn (oud model, nieuw model, ander jaar).
De ALT kolom (kolom AJ) is het canonieke basisnummer per familie.

Kolom-indeling XLS (0-gebaseerd):
  R17=0, R16=2, ..., R1=32  — oudere (superseded) nummers
  ALT=35                     — huidig canoniek nummer
  U1=39, U2=41, ..., U9=55  — nieuwere nummers
  FU kolommen (oneven) en kolom 34/36/37 bevatten type-vlaggen, geen nummers.

Formatteringsregel: een ALT dat eindigt op -000 wordt zonder die suffix opgeschreven;
andere suffixen (bv. -027, -294) worden wél meegeschreven.
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Set

import pandas as pd
try:
    import xlrd  # noqa: F401 — explicit import zodat PyInstaller xlrd meebundelt
except ImportError:
    pass

_index: Optional[Dict[str, Set[str]]] = None

_PART_COLS = list(range(0, 33, 2)) + [35] + list(range(39, 56, 2))
_ALT_COL = 35

# Geldig Suzuki-onderdeelnummer: DDDDD-XXXXX-XXX (5-5-3, alfanumeriek)
_PART_RE = re.compile(r"^\d{5}-[A-Z0-9]{5}-[A-Z0-9]{3}$")


def _is_part_number(s: str) -> bool:
    return bool(_PART_RE.match(s.strip().upper()))


def _norm(s: str) -> str:
    return s.strip().upper()


def _fmt(raw: str) -> str:
    """Formatteer voor weergave: strip -000 suffix."""
    s = raw.strip()
    if s.upper().endswith("-000"):
        return s[:-4]
    return s


def _build_index() -> Dict[str, Set[str]]:
    from utils.paths import resource_path

    path = resource_path("assets/Superseded lijst.xls")
    df = pd.read_excel(path, header=None, dtype=str)

    index: Dict[str, Set[str]] = {}

    for row_idx in range(1, len(df)):
        row = df.iloc[row_idx]
        alt_raw = str(row.iloc[_ALT_COL]).strip()
        if alt_raw in ("nan", "") or not _is_part_number(alt_raw):
            continue

        row_family: Set[str] = set()
        for col_idx in _PART_COLS:
            val = str(row.iloc[col_idx]).strip()
            if val in ("nan", "") or not _is_part_number(val):
                continue
            row_family.add(_norm(val))

        for num in row_family:
            if num not in index:
                index[num] = set()
            index[num] |= row_family

    return index


def _get_index() -> Dict[str, Set[str]]:
    global _index
    if _index is None:
        _index = _build_index()
    return _index


def preload_async() -> None:
    """Start het laden van de superseded index in een achtergrondthread."""
    import threading
    t = threading.Thread(target=_get_index, daemon=True, name="superseded-preload")
    t.start()


def lookup_superseded(part_number: str) -> List[str]:
    """
    Geeft alle gerelateerde (superseded) nummers voor een onderdeelnummer.
    Resultaat is gesorteerd en geformatteerd (geen -000 suffix).
    Retourneert lege lijst als het nummer niet gevonden wordt.
    """
    if not part_number or not part_number.strip():
        return []

    index = _get_index()
    raw = _norm(part_number)

    candidates: Set[str] = {raw}
    if raw.endswith("-000"):
        candidates.add(raw[:-4])
    else:
        candidates.add(raw + "-000")

    family: Set[str] = set()
    for candidate in candidates:
        if candidate in index:
            family |= index[candidate]

    if not family:
        return []

    base = raw[:-4] if raw.endswith("-000") else raw
    family = {n for n in family if (n[:-4] if n.endswith("-000") else n) != base}

    return sorted(set(_fmt(n) for n in family))
