# engines/engine_zoeklijst.py
"""
Engine: Zoeklijst / Mail parser (WC + TLC tradelist)

Update v2:
- TLC bron heet bij jullie 'TLC tradelist' (geen masterlijst)
- TLC tradelist headers: Title | Stock | Prijs | Locatie
- load_tlc_xlsx() herkent nu die headers
- Als TLC Stock <= 0: behandelen als 'niet beschikbaar' (Found=NO met note)
"""

from __future__ import annotations

from pathlib import Path
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
from utils.paths import output_root

LOCATION_HINT_RE = re.compile(r"\b[A-Z]{1,3}\d{1,2}(?:-\d{1,2})?\b")



def normalize_part_number(raw: str) -> Optional[str]:
    if raw is None:
        return None
    s = str(raw).strip().upper()
    if not s:
        return None

    s = re.sub(r"\s+", "", s)
    s = s.strip(",.;:()[]{}<>\"'")

    if "-" in s:
        s2 = re.sub(r"[^A-Z0-9-]", "", s)
        s2 = re.sub(r"-{2,}", "-", s2).strip("-")
        core = s2.replace("-", "")
        if len(core) < 10:
            return None
        return s2

    core = re.sub(r"[^A-Z0-9]", "", s)
    if len(core) < 8:
        return None

    if len(core) == 10:
        return f"{core[:5]}-{core[5:]}"
    if len(core) == 13:
        return f"{core[:5]}-{core[5:10]}-{core[10:]}"
    if len(core) > 10:
        return f"{core[:5]}-{core[5:10]}-{core[10:]}"
    return None


def extract_part_numbers_from_text(text: str) -> List[str]:
    if not text:
        return []
    tokens = re.findall(r"[A-Za-z0-9-]{8,}", text.upper())
    out: List[str] = []
    for t in tokens:
        n = normalize_part_number(t)
        if n:
            out.append(n)
    return out


def extract_part_numbers_from_xlsx(path: str) -> List[str]:
    df = pd.read_excel(path, dtype=str)
    parts: List[str] = []
    for col in df.columns:
        series = df[col].dropna().astype(str)
        for v in series.tolist():
            n = normalize_part_number(v)
            if n:
                parts.append(n)
    return parts


def _best_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {str(c).strip(): c for c in df.columns}
    for cand in candidates:
        if cand in cols:
            return cols[cand]
    lower_map = {str(c).lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def _to_float_safe(x) -> Optional[float]:
    if x is None:
        return None
    try:
        s = str(x).strip().replace(",", ".")
        if s == "" or s.lower() == "nan":
            return None
        return float(s)
    except Exception:
        return None


class EngineZoeklijst:
    def __init__(self, output_dir=str(output_root() / "zoeklijst")):
        # TLC tradelist (databron 2)
        self.tlc_df: Optional[pd.DataFrame] = None
        self._tlc_index: Dict[str, int] = {}
        self._tlc_title_col: Optional[str] = None  # part number
        self._tlc_stock_col: Optional[str] = None
        self._tlc_price_col: Optional[str] = None
        self._tlc_location_col: Optional[str] = None

        # WC export (databron 1)
        self.website_df: Optional[pd.DataFrame] = None

        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self._title_col: Optional[str] = None
        self._short_desc_col: Optional[str] = None
        self._stock_col: Optional[str] = None
        self._price_col: Optional[str] = None
        self._cat_col: Optional[str] = None
        self._brand_col: Optional[str] = None
        self._location_col: Optional[str] = None

        self._index: Dict[str, int] = {}
        self._dup_titles: Dict[str, List[int]] = {}

    # -------------------------
    # Loaders
    # -------------------------
    def load_website_df(self, df: pd.DataFrame) -> None:
        if df is None or df.empty:
            raise ValueError("WC export dataframe is leeg.")

        self.website_df = df.copy()
        self.website_df.columns = [str(c).strip() for c in self.website_df.columns]

        self._title_col = _best_column(self.website_df, ["Title", "Naam", "Name", "post_title"])
        self._short_desc_col = _best_column(self.website_df, ["Korte beschrijving", "Short description", "Short Description"])
        self._stock_col = _best_column(self.website_df, ["Voorraad", "Stock", "Stock quantity", "Stock Quantity"])
        self._price_col = _best_column(self.website_df, ["Reguliere prijs", "Regular price", "Prijs", "Price"])
        self._cat_col = _best_column(self.website_df, ["Categorieën", "Categories", "Productcategorieën", "Product categories"])
        self._brand_col = _best_column(self.website_df, ["Merken", "Brands", "Brand"])

        self._location_col = _best_column(self.website_df, ["Locatie", "Location", "Magazijnlocatie", "Warehouse location", "Opslaglocatie"])
        if not self._location_col:
            best = None
            best_score = 0.0
            for c in self.website_df.columns:
                try:
                    s = self.website_df[c].astype(str)
                except Exception:
                    continue
                score = s.str.contains(LOCATION_HINT_RE, na=False).mean()
                if score > best_score and score > 0.05:
                    best, best_score = c, score
            self._location_col = best

        if not self._title_col:
            raise ValueError("Kon geen Title/Naam kolom vinden in WC export.")

        self._index.clear()
        self._dup_titles.clear()
        titles = self.website_df[self._title_col].astype(str).fillna("")
        for idx, val in enumerate(titles):
            norm = normalize_part_number(val) or re.sub(r"[^A-Z0-9-]", "", str(val).upper())
            if not norm:
                continue
            key = norm.replace("-", "")
            if key in self._index:
                self._dup_titles.setdefault(key, [self._index[key]]).append(idx)
            else:
                self._index[key] = idx

    def load_tlc_xlsx(self, path: str) -> None:
        """Laad TLC tradelist (xlsx) met headers: Title | Stock | Prijs | Locatie."""
        df = pd.read_excel(path, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]

        self._tlc_title_col = _best_column(df, ["Title", "Naam", "Name", "Part number", "Artikelnummer", "SKU"])
        self._tlc_stock_col = _best_column(df, ["Stock", "Voorraad", "Stock quantity", "Aantal"])
        self._tlc_price_col = _best_column(df, ["Prijs", "Price", "Reguliere prijs", "Regular price"])
        self._tlc_location_col = _best_column(df, ["Locatie", "Location", "Magazijnlocatie", "Opslaglocatie"])

        if not self._tlc_title_col:
            raise ValueError("TLC tradelist: kon geen Title/Artikelnummer kolom vinden (verwacht: Title)." )

        self.tlc_df = df
        self._tlc_index.clear()

        titles = df[self._tlc_title_col].astype(str).fillna("")
        for idx, val in enumerate(titles):
            n = normalize_part_number(val)
            if not n:
                continue
            key = n.replace("-", "")
            if key not in self._tlc_index:
                self._tlc_index[key] = idx

    # -------------------------
    # Lookup
    # -------------------------
    def lookup_wc(self, part_number: str) -> Tuple[bool, Optional[pd.Series], List[int]]:
        if self.website_df is None:
            raise RuntimeError("WC export niet geladen.")
        key = part_number.replace("-", "")
        if key in self._index:
            idx = self._index[key]
            dup = self._dup_titles.get(key, [])
            return True, self.website_df.iloc[idx], dup
        return False, None, []

    def lookup_any(self, part_number: str) -> Tuple[str, Optional[pd.Series], List[int]]:
        found, row, dup = self.lookup_wc(part_number)
        if found and row is not None:
            return "WC", row, dup

        if self.tlc_df is not None:
            key = part_number.replace("-", "")
            if key in self._tlc_index:
                idx = self._tlc_index[key]
                return "TLC", self.tlc_df.iloc[idx], []
        return "NO", None, []

    # -------------------------
    # Report
    # -------------------------
    def build_report(self, part_numbers: List[str]) -> pd.DataFrame:
        if self.website_df is None:
            raise RuntimeError("WC export niet geladen.")

        seen = set()
        uniq: List[str] = []
        for p in part_numbers:
            k = p.replace("-", "")
            if k not in seen:
                uniq.append(p)
                seen.add(k)

        def get(row: pd.Series, col: Optional[str]) -> str:
            if not col:
                return ""
            v = row.get(col, "")
            if pd.isna(v):
                return ""
            return str(v).strip()

        rows: List[dict] = []
        for p in uniq:
            source, row, dup_idxs = self.lookup_any(p)

            if source == "NO" or row is None:
                # Probeer via superseded familie (bidirectioneel)
                found_via: Optional[tuple] = None
                sup_numbers: List[str] = []
                try:
                    from services.superseded import lookup_superseded
                    sup_numbers = lookup_superseded(p)
                    for sup_nr in sup_numbers:
                        s2, r2, d2 = self.lookup_any(sup_nr)
                        if s2 != "NO" and r2 is not None:
                            found_via = (s2, r2, d2, sup_nr)
                            break
                except Exception:
                    pass

                if found_via:
                    s2, r2, d2, sup_nr = found_via
                    sup_note = f"Gevonden via superseded → {sup_nr}"
                    if s2 == "WC":
                        notes = [sup_note]
                        if d2:
                            notes.append(f"DUBBEL in WC ({1 + len(d2)}x)")
                        rows.append({
                            "Part number": p,
                            "Source": "WC",
                            "Found": "YES",
                            "Stock": get(r2, self._stock_col),
                            "Price": get(r2, self._price_col),
                            "Locatie": get(r2, self._location_col),
                            "Notes": " | ".join(notes),
                        })
                    else:  # TLC
                        stock_str = get(r2, self._tlc_stock_col)
                        stock_val = _to_float_safe(stock_str)
                        if stock_val is not None and stock_val <= 0:
                            rows.append({
                                "Part number": p,
                                "Source": "TLC",
                                "Found": "NO",
                                "Stock": stock_str,
                                "Price": get(r2, self._tlc_price_col),
                                "Locatie": get(r2, self._tlc_location_col),
                                "Notes": f"{sup_note} | Voorraad = 0 (niet beschikbaar)",
                            })
                        else:
                            rows.append({
                                "Part number": p,
                                "Source": "TLC",
                                "Found": "YES",
                                "Stock": stock_str,
                                "Price": get(r2, self._tlc_price_col),
                                "Locatie": get(r2, self._tlc_location_col),
                                "Notes": sup_note,
                            })
                    continue

                # Echt niet gevonden — vermeld eventuele familie als hint
                not_found_note = "Niet gevonden (WC + TLC)"
                if sup_numbers:
                    not_found_note += f" | Superseded familie: {', '.join(sup_numbers[:4])}"
                rows.append({
                    "Part number": p,
                    "Source": "NO",
                    "Found": "NO",
                    "Stock": "",
                    "Price": "",
                    "Locatie": "",
                    "Notes": not_found_note,
                })
                continue

            if source == "WC":
                notes = []
                if dup_idxs:
                    notes.append(f"DUBBEL in WC ({1+len(dup_idxs)}x)")
                if not get(row, self._location_col):
                    notes.append("Geen locatie in WC export (of niet gedetecteerd)")

                rows.append({
                    "Part number": p,
                    "Source": "WC",
                    "Found": "YES",
                    "Stock": get(row, self._stock_col),
                    "Price": get(row, self._price_col),
                    "Locatie": get(row, self._location_col),
                    "Notes": " | ".join(notes),
                })
            else:
                # TLC tradelist
                stock_str = get(row, self._tlc_stock_col)
                stock_val = _to_float_safe(stock_str)
                if stock_val is not None and stock_val <= 0:
                    # jullie regel: voorraad 0 = hebben we niet
                    rows.append({
                        "Part number": p,
                        "Source": "TLC",
                        "Found": "NO",
                        "Stock": stock_str,
                        "Price": get(row, self._tlc_price_col),
                        "Locatie": get(row, self._tlc_location_col),
                        "Notes": "In TLC tradelist maar voorraad = 0 (niet beschikbaar)",
                    })
                else:
                    rows.append({
                        "Part number": p,
                        "Source": "TLC",
                        "Found": "YES",
                        "Stock": stock_str,
                        "Price": get(row, self._tlc_price_col),
                        "Locatie": get(row, self._tlc_location_col),
                        "Notes": "Gevonden in TLC tradelist",
                    })

        return pd.DataFrame(rows)

    def export_report_xlsx(self, report_df: pd.DataFrame, filename: str = "zoeklijst_report.xlsx") -> Path:
        out_path = self.output_dir / filename
        out_path.parent.mkdir(parents=True, exist_ok=True)
        report_df.to_excel(out_path, index=False)
        return out_path

    @staticmethod
    def _is_uitverkocht(row: pd.Series) -> bool:
        """Rij is uitverkocht/niet beschikbaar als Found=NO, of als stock <= 0."""
        if str(row.get("Found", "")).strip().upper() == "NO":
            return True
        stock = str(row.get("Stock", "")).strip().replace(",", ".")
        try:
            return float(stock) <= 0
        except ValueError:
            return stock == ""

    def export_report_xlsx_splits(self, report_df: pd.DataFrame, stem: str = "zoeklijst_report") -> dict:
        """
        Sla 3 varianten op:
          {stem}_compleet.xlsx      — alles
          {stem}_beschikbaar.xlsx   — zonder uitverkochte regels
          {stem}_uitverkocht.xlsx   — alleen uitverkochte regels
        Geeft een dict terug met keys 'compleet', 'beschikbaar', 'uitverkocht' en 'folder'.
        """
        self.output_dir.mkdir(parents=True, exist_ok=True)

        mask_uit = report_df.apply(self._is_uitverkocht, axis=1)

        paden = {}
        voor = [
            ("compleet",    report_df),
            ("beschikbaar", report_df[~mask_uit]),
            ("uitverkocht", report_df[mask_uit]),
        ]
        for label, df in voor:
            pad = self.output_dir / f"{stem}_{label}.xlsx"
            df.to_excel(pad, index=False)
            paden[label] = pad

        paden["folder"] = self.output_dir
        return paden
