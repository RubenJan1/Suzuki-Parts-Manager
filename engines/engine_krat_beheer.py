"""Business logica voor Krat Beheer — WC lookup en XLSX export."""

from __future__ import annotations
import os
from datetime import date
from typing import Optional

import pandas as pd


def _get_col(df: pd.DataFrame, row: pd.Series, *candidates: str) -> str:
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        real = lower_map.get(cand.strip().lower())
        if real is not None:
            return str(row.get(real, "") or "").strip()
    return ""


def wc_lookup(artikelnummer: str, wc_df) -> dict | None:
    """Zoek artikel op in WC export op Naam-kolom. Geeft match-dict of None."""
    if wc_df is None or wc_df.empty or not artikelnummer:
        return None

    naam_col = None
    for c in wc_df.columns:
        if str(c).strip().lower() in ("naam", "name", "title"):
            naam_col = c
            break
    if naam_col is None:
        return None

    mask = wc_df[naam_col].astype(str).str.strip().str.upper() == artikelnummer.strip().upper()
    if not mask.any():
        return None

    row = wc_df[mask].iloc[0]

    prijs_str = _get_col(wc_df, row, "Reguliere prijs", "Regular price", "Prijs")
    try:
        prijs = float(prijs_str.replace(",", ".").replace("€", "").strip() or "0")
    except Exception:
        prijs = 0.0

    voorraad_str = _get_col(wc_df, row, "Voorraad", "Stock")
    try:
        voorraad = int(float(voorraad_str or "0"))
    except Exception:
        voorraad = 0

    cats_raw = _get_col(wc_df, row, "Categorieën", "Productcategorieën", "Categories", "Product categories")
    cat_list = [c.strip() for c in cats_raw.split(",") if c.strip()] if cats_raw else []

    return {
        "wc_id":             str(row.get("ID", "") or "").strip(),
        "naam":              artikelnummer,
        "prijs":             prijs,
        "voorraad":          voorraad,
        "locatie":           _get_col(wc_df, row, "Beschrijving", "Locatie", "Description"),
        "short_description": _get_col(wc_df, row, "Korte beschrijving", "Short Description"),
        "categorieen":       cat_list,
    }


def count_beprijsd(krat: dict) -> tuple[int, int]:
    """Geeft (aantal beprijsd/overgeslagen, totaal) terug."""
    artikelen = krat.get("artikelen", [])
    done = sum(1 for a in artikelen if a.get("prijs_status") is not None)
    return done, len(artikelen)


def _rows_nieuw(krat: dict) -> list[dict]:
    rows = []
    locatie = krat.get("locatie", "")
    for art in krat.get("artikelen", []):
        if art.get("samenvoeg_beslissing") == "update":
            continue
        prijs_status = art.get("prijs_status")
        prijs = art.get("prijs")
        if prijs_status == "overgeslagen" or prijs is None:
            stock, prijs_export, loc = 0, 0.0, ""
        else:
            stock = art.get("voorraad", 0)
            prijs_export = float(prijs or 0)
            loc = locatie
        rows.append({
            "ID":                 "",
            "Title":              art.get("artikelnummer", ""),
            "Productcategorieën": "|".join(art.get("categorieen", [])),
            "Stock":              stock,
            "Short Description":  art.get("omschrijving", ""),
            "Locatie":            loc,
            "Prijs":              prijs_export,
        })
    return rows


def _rows_samenvoeg(krat: dict) -> list[dict]:
    rows = []
    for art in krat.get("artikelen", []):
        if art.get("samenvoeg_beslissing") != "update":
            continue
        wc = art.get("wc_match") or {}
        prijs_status = art.get("prijs_status")
        prijs = art.get("prijs")
        if prijs_status == "overgeslagen" or prijs is None:
            prijs_export = float(wc.get("prijs", 0))
        else:
            prijs_export = float(prijs or 0)
        stock = int(wc.get("voorraad", 0)) + int(art.get("voorraad", 0))
        rows.append({
            "ID":                 wc.get("wc_id", ""),
            "Title":              art.get("artikelnummer", ""),
            "Productcategorieën": "|".join(art.get("categorieen", [])),
            "Stock":              stock,
            "Short Description":  art.get("omschrijving", ""),
            "Locatie":            wc.get("locatie", krat.get("locatie", "")),
            "Prijs":              prijs_export,
        })
    return rows


def export_nieuwe_artikelen(krat: dict, path: str) -> Optional[str]:
    rows = _rows_nieuw(krat)
    if not rows:
        return None
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    pd.DataFrame(rows).to_excel(path, index=False)
    return path


def export_samenvoeg_update(krat: dict, path: str) -> Optional[str]:
    rows = _rows_samenvoeg(krat)
    if not rows:
        return None
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    pd.DataFrame(rows).to_excel(path, index=False)
    return path
