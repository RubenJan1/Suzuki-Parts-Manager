# ============================================================
# engines/engine_tlc_update.py
# TLC Update – Aanvullen / Krat vervangen (offline, hufterproof)
#
# Verwacht:
# - Actieve TLC: output/1322/TLC/TLC_1.xlsx  (met header: Title, Stock, Prijs, Locatie)
# - Updatebestanden (1 of meer): headerloos, vaste vakken:
#   A=Title, B=Stock, C=Locatie, D=Prijs
#
# Regels:
# - NOOIT samenvoegen op Title alleen
# - Alleen match op (Title + Locatie)
# - Mode MERGE: update/voeg toe op Title+Locatie
# - Mode REPLACE_LOC: vervang 1 of meer locaties volledig (krat opnieuw gedaan)
# ============================================================

from __future__ import annotations

import os
import re
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any

import pandas as pd


def ensure_dir(path: Path):
    path.mkdir(parents=True, exist_ok=True)


def run_id() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def to_int(x) -> int:
    try:
        s = str(x).strip()
        if s == "" or s.lower() == "nan":
            return 0
        return int(float(s.replace(",", ".").strip()))
    except Exception:
        return 0


def parse_price(val) -> float:
    if val is None:
        return 0.0
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return 0.0
    s = s.replace("€", "").replace(" ", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def euro(x: float) -> str:
    # NL-format 1.234,56
    try:
        return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"


def clean_title(x: str) -> str:
    return str(x or "").strip()




def clean_loc(x: str) -> str:
    s = str(x or "").strip().upper()
    if not s:
        return ""

    # TLC60 / TLC 60 / TLC060 -> 60
    m = re.match(r"^TLC\s*0*(\d+)$", s, flags=re.IGNORECASE)
    if m:
        return str(int(m.group(1)))

    # alleen getal -> normaliseer leading zeros
    if s.isdigit():
        return str(int(s))

    return s



class TLCUpdateEngine:
    def __init__(self):
        self.base_dir: Optional[str] = None  # output/1322
        self.tlc_filename: str = "TLC_1.xlsx"
        self.update_paths: List[str] = []

    # -----------------------------
    # Setup
    # -----------------------------
    def set_base_dir(self, base_dir: str):
        self.base_dir = base_dir

    def clear_updates(self):
        self.update_paths = []

    def add_update_file(self, path: str):
        if not path or not os.path.exists(path):
            return

        name = os.path.basename(path)
        # Skip Excel temporary lock files
        if name.startswith("~$"):
            return

        self.update_paths.append(path)


    # -----------------------------
    # Paths
    # -----------------------------
    def _paths_for_run(self, rid: str) -> Dict[str, Path]:
        if not self.base_dir:
            raise RuntimeError("base_dir niet gezet. Gebruik set_base_dir().")

        base = Path(self.base_dir)
        tlc_dir = base / "TLC"
        backup_dir = base / "BACKUP_TLC"
        debug_dir = base / "DEBUG" / f"RUN_{rid}"

        ensure_dir(tlc_dir)
        ensure_dir(backup_dir)
        ensure_dir(debug_dir)

        active_tlc = tlc_dir / self.tlc_filename
        day = datetime.now().strftime("%Y-%m-%d")
        backup_tlc = backup_dir / f"TLC_1_BACKUP_{day}_{rid}.xlsx"


        return {
            "base": base,
            "tlc_dir": tlc_dir,
            "backup_dir": backup_dir,
            "debug_dir": debug_dir,
            "active_tlc": active_tlc,
            "backup_tlc": backup_tlc,
        }

    # -----------------------------
    # Readers
    # -----------------------------
    def _read_active_tlc(self, path: str) -> pd.DataFrame:
        """
        TLC_1.xlsx (jullie) heeft header: Title, Stock, Prijs, Locatie.
        We ondersteunen ook headerloos (auto-detect), dan nemen we:
          [Title, Stock, Prijs, Locatie] uit de eerste 4 kolommen.
        """
        df0 = pd.read_excel(path, header=None, dtype=str)
        first = str(df0.iat[0, 0] or "").strip().lower()

        if first == "title":
            df = pd.read_excel(path, header=0, dtype=str)
        else:
            df = df0.iloc[:, :4].copy()
            df.columns = ["Title", "Stock", "Prijs", "Locatie"]

        df.columns = [str(c).strip() for c in df.columns]
        needed = ["Title", "Stock", "Prijs", "Locatie"]
        for c in needed:
            if c not in df.columns:
                raise RuntimeError(f"TLC mist kolom: {c}")

        df = df[needed].copy()
        df["Title"] = df["Title"].apply(clean_title)
        df["Locatie"] = df["Locatie"].apply(clean_loc)
        df["Stock"] = df["Stock"].apply(to_int)
        df["Prijs"] = df["Prijs"].apply(parse_price)

        df = df[(df["Title"] != "") & (df["Locatie"] != "")].copy()
        return df

    def _read_update_file(self, path: str) -> pd.DataFrame:
        """
        Updatebestand: headerloos met vaste vakken:
        A=Title, B=Stock, C=Locatie, D=Prijs
        """
        try:
            df = pd.read_excel(path, header=None, dtype=str, engine="openpyxl")
        except Exception as e:
            raise RuntimeError(
                f"Kan updatebestand niet lezen:\n{path}\n\n"
                f"Excel fout: {e}\n\n"
                "Check:\n"
                "• Is dit een echte .xlsx (geen csv hernoemd)?\n"
                "• Is het bestand beschadigd/leeg?\n"
                "• Open het bestand in Excel en 'Opslaan als' -> .xlsx.\n"
                "• Upload niet het ~$.xlsx lock-bestand."
            )

        if df.shape[0] == 0 and df.shape[1] == 0:
            raise RuntimeError(
                f"Updatebestand is leeg of heeft geen sheet:\n{path}\n\n"
                "Open in Excel en sla opnieuw op als .xlsx."
            )

        # Forceer 4 kolommen
        if df.shape[1] < 4:
            raise RuntimeError(
                f"Updatebestand heeft te weinig kolommen (minimaal A-D nodig):\n{path}\n"
                f"Gevonden kolommen: {df.shape[1]}"
            )

        df = df.iloc[:, :4].copy()
        df.columns = ["Title", "Stock", "Locatie", "Prijs"]

        df["Title"] = df["Title"].apply(clean_title)
        df["Locatie"] = df["Locatie"].apply(clean_loc)
        df["Stock"] = df["Stock"].apply(to_int)
        df["Prijs"] = df["Prijs"].apply(parse_price)

        # --- NEW: support blok-koppen zoals "TLC60" ---
        current_loc = ""
        keep_rows = []

        for _, r in df.iterrows():
            t = clean_title(r["Title"])
            loc = clean_loc(r["Locatie"])

            # kopregel detectie: "TLC60" in Title, en verder leeg
            m = re.match(r"^TLC\s*0*(\d+)$", t.strip(), flags=re.IGNORECASE)
            is_header = bool(m) and (to_int(r["Stock"]) == 0) and (parse_price(r["Prijs"]) == 0.0) and (loc == "")

            if is_header:
                current_loc = str(int(m.group(1)))
                continue  # kopregel niet meenemen als product

            # locatie invullen vanuit kopregel als leeg
            if loc == "" and current_loc:
                loc = current_loc

            keep_rows.append({
                "Title": t,
                "Stock": to_int(r["Stock"]),
                "Locatie": loc,
                "Prijs": parse_price(r["Prijs"]),
                "Source": r.get("Source", "")
            })

        out = pd.DataFrame(keep_rows)

        # drop echte lege regels
        out = out[(out["Title"] != "") & (out["Locatie"] != "")].copy()
        return out


    # -----------------------------
    # Validation
    # -----------------------------
    def _validate_updates(self, upd_all: pd.DataFrame, replace_locs: Optional[List[str]]):
        errors: List[str] = []
        warnings: List[str] = []

        missing_title = upd_all["Title"].eq("")
        missing_loc = upd_all["Locatie"].eq("")
        if missing_title.any():
            errors.append(f"{int(missing_title.sum())} regels missen Title (kolom A).")
        if missing_loc.any():
            errors.append(f"{int(missing_loc.sum())} regels missen Locatie (kolom C).")

        # conflicts inside update: same key different stock/price
        upd_all["_key"] = upd_all["Title"] + "||" + upd_all["Locatie"]
        grp = upd_all.groupby("_key", as_index=False).agg(
            stock_min=("Stock", "min"),
            stock_max=("Stock", "max"),
            price_min=("Prijs", "min"),
            price_max=("Prijs", "max"),
            count=("Stock", "count"),
        )
        conflicts = grp[
            (grp["count"] > 1) &
            ((grp["stock_min"] != grp["stock_max"]) | (grp["price_min"] != grp["price_max"]))
        ].copy()

        conflict_keys = []
        if not conflicts.empty:
            # maak lijst van keys (Title||Locatie)
            conflict_keys = conflicts["_key"].head(20).tolist()

            msg = (
                f"Updatebestanden bevatten conflicten: {len(conflicts)} keys met zelfde Title+Locatie "
                f"maar verschillende Stock/Prijs.\n"
                "Voorbeelden:\n- " + "\n- ".join(conflict_keys)
            )
            errors.append(msg)


        if replace_locs:
            for loc in replace_locs:
                if not (upd_all["Locatie"] == loc).any():
                    errors.append(f"Locatie '{loc}' gekozen voor vervangen, maar komt niet voor in updatebestanden.")

        # warnings
        title_locs = upd_all.groupby("Title")["Locatie"].nunique()
        multi = title_locs[title_locs > 1]
        if not multi.empty:
            warnings.append(f"{len(multi)} titles komen voor op meerdere locaties in de update (OK, maar check).")

        extreme = upd_all["Stock"].abs() > 500
        if extreme.any():
            warnings.append(f"{int(extreme.sum())} regels hebben voorraad > 500 (check).")

        return errors, warnings

    # -----------------------------
    # Apply logic
    # -----------------------------
    def run_update(self, mode: str = "MERGE", replace_locations: Optional[str] = None) -> Dict[str, Any]:
        """
        mode:
          - "MERGE"
          - "REPLACE_LOC"
        replace_locations:
          - None of string like "63" or "63,64"
        """
        if not self.update_paths:
            raise RuntimeError("Geen updatebestand(en) toegevoegd.")

        rid = run_id()
        P = self._paths_for_run(rid)

        active_tlc_path = str(P["active_tlc"])
        if not os.path.exists(active_tlc_path):
            raise RuntimeError(
                f"Actieve TLC ontbreekt: {active_tlc_path}\n"
                f"Plaats {self.tlc_filename} in: {P['tlc_dir']}"
            )

        replace_locs: Optional[List[str]] = None
        if mode == "REPLACE_LOC":
            if not replace_locations or not str(replace_locations).strip():
                raise RuntimeError("Vervang-modus gekozen maar geen locatie ingevuld.")
            replace_locs = [clean_loc(x) for x in re.split(r"[,\s;]+", str(replace_locations).strip()) if x.strip()]
            if not replace_locs:
                raise RuntimeError("Geen geldige locaties gevonden om te vervangen.")

        # Backup vóór wijzigingen
        shutil.copy2(active_tlc_path, str(P["backup_tlc"]))

        # Debug kopieën
        try:
            shutil.copy2(active_tlc_path, str(P["debug_dir"] / f"INPUT_{self.tlc_filename}"))
            for i, src in enumerate(self.update_paths, start=1):
                shutil.copy2(src, str(P["debug_dir"] / f"INPUT_UPDATE_{i}_{Path(src).name}"))
        except Exception:
            pass

        tlc = self._read_active_tlc(active_tlc_path)

        upd_frames = []
        for fp in self.update_paths:
            dfu = self._read_update_file(fp)
            dfu["Source"] = Path(fp).name
            upd_frames.append(dfu)

        upd_all = pd.concat(upd_frames, ignore_index=True)

        errors, warnings = self._validate_updates(upd_all.copy(), replace_locs)

        report_path = str(P["debug_dir"] / f"TLC_UPDATE_REPORT_{rid}.xlsx")

        # Stop on errors (write report)
        if errors:
            with pd.ExcelWriter(report_path, engine="openpyxl") as xw:
                pd.DataFrame({"Error": errors}).to_excel(xw, index=False, sheet_name="ERRORS")
                pd.DataFrame({"Warning": warnings}).to_excel(xw, index=False, sheet_name="WARNINGS")
                upd_all.to_excel(xw, index=False, sheet_name="STAGING")
            raise RuntimeError(f"Update gestopt door fouten. Zie report:\n{report_path}")

        # Apply
        updated_rows = []
        added_rows = []
        removed_rows = []

        if mode == "REPLACE_LOC" and replace_locs:
            mask = tlc["Locatie"].isin(replace_locs)
            removed_rows = tlc[mask].copy().to_dict("records")
            tlc = tlc[~mask].copy()
            upd_all = upd_all[upd_all["Locatie"].isin(replace_locs)].copy()

        # Deduplicate updates by exact key (conflicts already blocked)
        upd_all["_key"] = upd_all["Title"] + "||" + upd_all["Locatie"]
        upd_all = upd_all.drop_duplicates(subset=["_key"], keep="last").copy()

        tlc["_key"] = tlc["Title"] + "||" + tlc["Locatie"]
        tlc_index = {k: i for i, k in enumerate(tlc["_key"].tolist())}

        for _, u in upd_all.iterrows():
            title = u["Title"]
            loc = u["Locatie"]
            stock_new = int(u["Stock"])
            price_new = float(u["Prijs"])
            key = u["_key"]

            if key in tlc_index:
                i = tlc_index[key]
                stock_old = int(tlc.at[i, "Stock"])
                price_old = float(tlc.at[i, "Prijs"])

                # OVERWRITE voorraad (veilig)
                tlc.at[i, "Stock"] = stock_new

                # prijs: overschrijven alleen als update prijs > 0, anders laten staan
                if price_new > 0:
                    tlc.at[i, "Prijs"] = price_new

                updated_rows.append({
                    "Title": title,
                    "Locatie": loc,
                    "Stock_oud": stock_old,
                    "Stock_nieuw": stock_new,
                    "Prijs_oud": euro(price_old),
                    "Prijs_nieuw": euro(float(tlc.at[i, "Prijs"])),
                    "Source": u.get("Source", "")
                })
            else:
                tlc = pd.concat([tlc, pd.DataFrame([{
                    "Title": title,
                    "Stock": stock_new,
                    "Prijs": price_new,
                    "Locatie": loc,
                    "_key": key
                }])], ignore_index=True)
                added_rows.append({
                    "Title": title,
                    "Locatie": loc,
                    "Stock": stock_new,
                    "Prijs": euro(price_new),
                    "Source": u.get("Source", "")
                })

        # Remove <=0
        tlc = tlc[tlc["Stock"] > 0].copy()

        # Sort: Locatie then Title
        def loc_sort(x):
            s = str(x).strip().upper()
            if s.isdigit():
                return (0, int(s), s)
            return (1, 999999, s)

        tlc["_loc_sort"] = tlc["Locatie"].apply(loc_sort)
        tlc = tlc.sort_values(by=["_loc_sort", "Title"], ascending=True).drop(columns=["_loc_sort"])

        # Write TLC back with header
        out_tlc = tlc[["Title", "Stock", "Prijs", "Locatie"]].copy()
        out_tlc["Prijs"] = out_tlc["Prijs"].apply(lambda x: euro(float(x)))
        out_tlc.to_excel(active_tlc_path, index=False)

        # Report
        summary = {
            "run_id": rid,
            "mode": mode,
            "replace_locations": ", ".join(replace_locs) if replace_locs else "",
            "updates_files": ", ".join([Path(x).name for x in self.update_paths]),
            "updated_rows": len(updated_rows),
            "added_rows": len(added_rows),
            "removed_rows": len(removed_rows) if mode == "REPLACE_LOC" else 0,
            "warnings_count": len(warnings),
            "active_tlc": active_tlc_path,
            "backup_tlc": str(P["backup_tlc"]),
            "debug_dir": str(P["debug_dir"]),
        }
        with pd.ExcelWriter(report_path, engine="openpyxl") as xw:
            pd.DataFrame(list(summary.items()), columns=["Key", "Value"]).to_excel(xw, index=False, sheet_name="SUMMARY")
            pd.DataFrame({"Warning": warnings}).to_excel(xw, index=False, sheet_name="WARNINGS")
            pd.DataFrame(updated_rows).to_excel(xw, index=False, sheet_name="UPDATED")
            pd.DataFrame(added_rows).to_excel(xw, index=False, sheet_name="ADDED")
            if mode == "REPLACE_LOC":
                pd.DataFrame(removed_rows).to_excel(xw, index=False, sheet_name="REMOVED")
            upd_all.drop(columns=["_key"], errors="ignore").to_excel(xw, index=False, sheet_name="STAGING_USED")

        return {
            "rid": rid,
            "active_tlc": active_tlc_path,
            "backup_tlc": str(P["backup_tlc"]),
            "debug_dir": str(P["debug_dir"]),
            "report": report_path,
            "updated": len(updated_rows),
            "added": len(added_rows),
            "removed": int(len(removed_rows)) if mode == "REPLACE_LOC" else 0,
            "warnings": len(warnings),
        }
