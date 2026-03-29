# ============================================================
# engines/engine_tlc_1322.py
# TLC / 1322 – Interne verkoop (MAANDLIJST)
# CORRECTE LOGICA: regel-voor-regel afboeken
# ============================================================
import re
import os
from datetime import datetime
import pandas as pd
import html
import shutil
from pathlib import Path


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------

import html

def _clean_loc_text(loc: str) -> str:
    """Opschonen van locatie: HTML tags + entities + (escaped) newlines eruit, whitespace normaliseren."""
    if not isinstance(loc, str):
        return ""
    s = html.unescape(loc)
    s = re.sub(r"<[^>]+>", " ", s)
    s = s.replace("\\n", " ").replace("\\r", " ").replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip().upper()
    return s


def location_sort_key(loc):
    """
    Natuurlijke sortering voor magazijnlocaties.
    - Excel-achtig: D1, D2, ... D10, D11 ...
    - Samengestelde locaties: 'GS 550/D4' -> D4
    - BGT/BGS/B... worden samengevoegd als 1 groep (B)
    """
    s = _clean_loc_text(loc)
    if not s:
        return (999, "", 999999, "")

    if "/" in s:
        parts = [p.strip() for p in s.split("/") if p.strip()]
        for p in reversed(parts):
            if re.search(r"[A-Z]+\s*\d+", p):
                s = p
                break

    s = re.sub(r"\bP\s+D\b", "PD", s)
    s = re.sub(r"\bPL\s*D\b", "PLD", s)
    s = s.replace(" ", "")

    if s.startswith("BGT") or s.startswith("BGS") or s == "B":
        prefix, num = "B", 0
    elif s.startswith("BR"):
        prefix, num = "BR", 0
    else:
        m = re.search(r"^([A-Z]+)(\d+)$", s)
        if m:
            prefix, num = m.group(1), int(m.group(2))
        else:
            m2 = re.search(r"([A-Z]+)(\d+)", s)
            if m2:
                prefix, num = m2.group(1), int(m2.group(2))
            else:
                prefix, num = s, 999999

    priority_map = {
        "BR": 10,
        "B": 20,
        "D": 30,
        "PD": 35,
        "PLD": 40,
        "GT": 50,
        "GS": 60,
        "GR": 65,
        "GTR": 70,
        "GB": 75,
        "H": 80,
        "Y": 90,
        "RGV": 100,
    }
    pri = priority_map.get(prefix, 500)
    return (pri, prefix, num, s)


def clean_location(loc):
    return _clean_loc_text(loc)


def location_sort_key_1322(loc):
    """
    TLC/1322 sort:
    - locaties zijn meestal alleen nummers -> klein naar groot
    - soms 'D100' -> na nummers, ook klein naar groot
    """
    s = _clean_loc_text(loc)
    if not s:
        return (999, 999999, "")
    s0 = s.replace(" ", "")
    if re.match(r"^\d+$", s0):
        return (0, int(s0), s0)
    m = re.match(r"^D(\d+)$", s0, flags=re.I)
    if m:
        return (1, int(m.group(1)), s0.upper())
    m2 = re.match(r"^([A-Z]+)(\d+)$", s0)
    if m2:
        return (2, int(m2.group(2)), s0.upper())
    return (3, 999999, s0.upper())


def _primary_loc_from_multi(loc_used: str) -> str:
    """
    Picklijst locatie kan zijn: '12(1) + 18(2)'.
    Voor sorteren willen we de eerste locatie ('12').
    """
    s = str(loc_used or "").strip()
    if not s:
        return ""

    first = s.split("+")[0].strip()          # '12(1)'
    first = re.sub(r"\(\d+\)", "", first)    # '(1)' weg -> '12'
    return first.strip()


# --- Model helpers (voor D-locaties) ---
def extract_model_from_categories(cat_text: str) -> str:
    """
    Haal een model zoals GS550 / GT750 / T500 uit categorie-tekst.
    We kiezen bij voorkeur een 'eindmodel' (bv 'GS550' i.p.v. 'GS series').
    """
    if not isinstance(cat_text, str):
        return ""
    s = cat_text.upper()
    # match bekende families met cijfers
    m = re.search(r"\b(GSX\-?R|GSX|GSF|GS|GT|T|TS|DR|SP|RM|RF|RGV)\s*([0-9]{2,4})\b", s)
    if not m:
        return ""
    fam = m.group(1).replace("-", "")
    num = m.group(2)
    return f"{fam}{num}"

def model_sort_key(model: str):
    m = (model or "").upper()
    if not m:
        return (999, "", 999999)
    fam_map = {
        "GT": 10, "T": 12, "TS": 14, "RGV": 18,
        "GS": 30, "GSX": 32, "GSXR": 33, "GSF": 34,
        "DR": 50, "SP": 52, "RM": 54, "RF": 56,
    }
    fam = re.match(r"^[A-Z]+", m).group(0)
    pri = fam_map.get(fam, 500)
    num_m = re.search(r"(\d+)", m)
    num = int(num_m.group(1)) if num_m else 999999
    return (pri, fam, num)

def pick_location_only_d(loc_raw: str, cat_text: str) -> str:
    """1322: géén model/categorie-prefix meer; locatie blijft locatie."""
    return clean_location(loc_raw)

def pick_sort_key(loc_raw: str, cat_text: str):
    loc = pick_location_only_d(loc_raw, cat_text)
    return location_sort_key_1322(loc)



def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def run_id():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def to_int(x):
    try:
        return int(float(str(x).replace(",", ".")))
    except:
        return 0


def parse_price(val):
    if val is None:
        return 0.0

    s = str(val)

    # verwijder euro, spaties, rare tekens
    s = s.replace("€", "").replace(" ", "").strip()

    # NL → EN decimaal
    s = s.replace(",", ".")

    try:
        return float(s)
    except:
        return 0.0
    
def _join_fact(series):
    vals = sorted({str(x).strip() for x in series if str(x).strip()})
    return ", ".join(vals)





def euro(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ------------------------------------------------------------
# Engine
# ------------------------------------------------------------
class TLC1322Engine:
    def __init__(self):
        # --- NEW: vaste rootmap voor waterproof systeem ---
        self.base_dir = None
        self.tlc_filename = "TLC_1.xlsx"

        # (oude velden)
        self.tlc_path = None          # houden we voor backward compat, maar we gebruiken 'm niet meer
        self.cms_paths = []
        self.wc_df = None  # optioneel: WC-export voor categorie/model lookup


    def set_tlc_path(self, path):
        self.tlc_path = path

    def add_cms_1322(self, path):
        self.cms_paths.append(path)

    def clear(self):
        self.tlc_path = None
        self.cms_paths = []
        self.wc_df = None  # optioneel: WC-export voor categorie/model lookup

    def _primary_loc_from_multi(loc_used: str) -> str:
        s = str(loc_used or "").strip()
        if not s:
            return ""
        first = s.split("+")[0].strip()          # '24(2)'
        first = re.sub(r"\(\d+\)", "", first)    # -> '24'
        return first.strip()
    
    # --------------------------------------------------------
    # NEW: Waterproof filesystem helpers
    # --------------------------------------------------------
    def set_base_dir(self, base_dir: str):
        self.base_dir = base_dir

    def _paths_for_run(self, rid: str):
        if not self.base_dir:
            raise RuntimeError("base_dir is niet gezet. Gebruik set_base_dir().")

        base = Path(self.base_dir)
        tlc_dir = base / "TLC"
        backup_dir = base / "BACKUP_TLC"
        pick_dir = base / "PICKLIJST"
        debug_run_dir = base / "DEBUG" / f"RUN_{rid}"

        tlc_dir.mkdir(parents=True, exist_ok=True)
        backup_dir.mkdir(parents=True, exist_ok=True)
        pick_dir.mkdir(parents=True, exist_ok=True)
        debug_run_dir.mkdir(parents=True, exist_ok=True)

        active_tlc = tlc_dir / self.tlc_filename
        backup_tlc = backup_dir / f"TLC_1_{rid}.xlsx"

        return {
            "base": base,
            "tlc_dir": tlc_dir,
            "backup_dir": backup_dir,
            "pick_dir": pick_dir,
            "debug_dir": debug_run_dir,
            "active_tlc": active_tlc,
            "backup_tlc": backup_tlc,
        }


    def _write_runlog(self, debug_dir: Path, rid: str, meta: dict):
        """Schrijf een simpele runlog.xlsx in debug map."""
        try:
            rows = [[k, str(v)] for k, v in meta.items()]
            df = pd.DataFrame(rows, columns=["Key", "Value"])
            df.to_excel(debug_dir / f"RUNLOG_{rid}.xlsx", index=False)
        except Exception:
            pass


    # --------------------------------------------------------
    # RUN
    # --------------------------------------------------------
    def run(self):
        if not self.cms_paths:
            raise RuntimeError("Geen CMS 1322-bestellingen geladen")

        rid = run_id()
        P = self._paths_for_run(rid)

        active_tlc_path = str(P["active_tlc"])
        if not os.path.exists(active_tlc_path):
            raise RuntimeError(
                f"Actieve TLC ontbreekt: {active_tlc_path}\n"
                f"Zet {self.tlc_filename} in: {P['tlc_dir']}"
            )

        # backup vóór afboeken
        shutil.copy2(active_tlc_path, str(P["backup_tlc"]))

        # debug kopieën (handig bij fouten)
        try:
            shutil.copy2(active_tlc_path, str(P["debug_dir"] / f"INPUT_{self.tlc_filename}"))
            for i, src in enumerate(self.cms_paths, start=1):
                shutil.copy2(src, str(P["debug_dir"] / f"INPUT_ORDER_{i}_{Path(src).name}"))
        except Exception:
            pass

        # gebruik deze TLC als basis voor de bestaande read_excel code
        self.tlc_path = active_tlc_path


        # WC categorie lookup (optioneel, voor MODEL bij D-locaties)
        wc_cat_by_title = {}
        if isinstance(self.wc_df, pd.DataFrame) and not self.wc_df.empty:
            # probeer titel en categorie kolommen te vinden
            cols = {c.lower(): c for c in self.wc_df.columns}
            col_title = cols.get("naam") or cols.get("title") or cols.get("post_title")
            col_cat = cols.get("categorieën") or cols.get("categories") or cols.get("cat")
            if col_title and col_cat:
                tmp = self.wc_df[[col_title, col_cat]].copy()
                tmp[col_title] = tmp[col_title].astype(str).str.strip()
                tmp[col_cat] = tmp[col_cat].astype(str)
                wc_cat_by_title = dict(zip(tmp[col_title], tmp[col_cat]))

        # ----------------------------
        # TLC basis (startvoorraad)
        # ----------------------------
        tlc = pd.read_excel(
            self.tlc_path,
            header=None,
            names=["Title", "Stock", "Prijs", "Locatie"],
            dtype=str
        )

        tlc["Title"] = tlc["Title"].fillna("").astype(str).str.strip()
        tlc = tlc[tlc["Title"] != ""].copy()
        # --- NIEUW: verwijder header-achtige rij als die als data is ingelezen ---
        tlc = tlc[~tlc["Title"].astype(str).str.strip().str.upper().eq("TITLE")].copy()

        tlc["Stock"] = tlc["Stock"].apply(to_int)
        tlc["Prijs"] = tlc["Prijs"].apply(parse_price)
        tlc["Locatie"] = tlc["Locatie"].fillna("").astype(str)


        # ----------------------------
        # CMS regels (GEEN groupby!)
        # ----------------------------
        cms_rows = []
        raw_lines_count = 0
   


        for p in self.cms_paths:
            df = pd.read_excel(
                p,
                header=None,
                names=["Title", "Omschrijving", "Aantal", "Prijs", "Factuur"],
                dtype=str
            )

            raw_lines_count += len(df)

            df["Title"] = df["Title"].fillna("").astype(str).str.strip()
            df = df[df["Title"] != ""].copy()
            df["Aantal"] = df["Aantal"].apply(to_int)
            df["Factuur"] = df["Factuur"].fillna("").astype(str)
            df["Omschrijving"] = df["Omschrijving"].fillna("").astype(str)

            for _, r in df.iterrows():
                cms_rows.append(r)
        # ----------------------------
        # Orders samenvoegen (dubbele regels)
        # - zelfde Title+Omschrijving+Prijs -> aantallen optellen
        # ----------------------------
        if cms_rows:
            cms_df = pd.DataFrame(cms_rows)
            cms_df["Title"] = cms_df["Title"].fillna("").astype(str).str.strip()
            cms_df["Omschrijving"] = cms_df["Omschrijving"].fillna("").astype(str)
            cms_df["Aantal"] = cms_df["Aantal"].apply(to_int)
            cms_df["Prijs"] = cms_df["Prijs"].apply(parse_price)
            cms_df["Factuur"] = cms_df["Factuur"].fillna("").astype(str)

            def _join_text(series):
                vals = [str(x).strip() for x in series if str(x).strip()]
                # uniek houden, volgorde ongeveer behouden
                seen = set()
                out = []
                for v in vals:
                    if v not in seen:
                        out.append(v); seen.add(v)
                return " | ".join(out)

            cms_df = (
                cms_df.groupby(["Title", "Prijs"], as_index=False)
                    .agg({
                        "Aantal": "sum",
                        "Factuur": _join_fact,
                        "Omschrijving": _join_text,
                    })
            )

            cms_rows = cms_df.to_dict("records")
            merged_lines_count = int(len(cms_rows))
            duplicates_merged = int(raw_lines_count - merged_lines_count)

        # ----------------------------
        # Output lijsten
        # ----------------------------
        tlc_nieuw = []
        picklijst = []
        uitverkocht = []


        # ----------------------------
        # Verwerking: REGEL VOOR REGEL (met ondersteuning voor dubbele Titles in TLC)
        # ----------------------------
        found_count = 0
        not_found_count = 0
        tekort_count = 0

        # Index: Title -> TLC rijen (meerdere mogelijk)
        tlc_by_title = {}
        for idx, rr in tlc.iterrows():
            t = str(rr["Title"]).strip()
            if not t:
                continue
            tlc_by_title.setdefault(t, []).append(idx)

        for r in cms_rows:
            title = str(r["Title"]).strip()
            gevraagd = int(r["Aantal"])
            factuur = r.get("Factuur", "")
            oms = r.get("Omschrijving", "")

            cand_idxs = tlc_by_title.get(title, [])

            # Kandidaten kiezen: eerst zelfde prijs als orderprijs (als meegegeven), daarna locatie klein->groot
            order_price = parse_price(r.get("Prijs", 0.0))

            def cand_key(i):
                pr = float(tlc.at[i, "Prijs"])
                price_match = 0 if (order_price > 0 and abs(pr - order_price) < 0.001) else 1
                loc = tlc.at[i, "Locatie"]
                return (price_match, location_sort_key_1322(loc), pr, i)

            cand_sorted = sorted(cand_idxs, key=cand_key)

            remaining = gevraagd
            geleverd_total = 0

            # nieuw: echte allocaties [(locatie, aantal)]
            alloc_parts = []
            first_price = None

            # totaal voorraad oud (som over alle TLC regels met dit title)
            stock_oud_total = sum(int(tlc.at[i, "Stock"]) for i in cand_idxs) if cand_idxs else 0
            # --- NIEUW: alle locaties waar voorraad lag (vóór afboeken) ---
            all_loc_parts = []
            for i in cand_sorted:
                st = int(tlc.at[i, "Stock"])
                if st <= 0:
                    continue
                loc = clean_location(tlc.at[i, "Locatie"]) or "LOCATIE ONTBREEKT"
                all_loc_parts.append((loc, st))

            if len(all_loc_parts) == 0:
                loc_all = ""
            elif len(all_loc_parts) == 1:
                loc_all = f"{all_loc_parts[0][0]}({all_loc_parts[0][1]})"
            else:
                loc_all = " + ".join([f"{loc}({st})" for loc, st in all_loc_parts])

            
            for i in cand_sorted:
                if remaining <= 0:
                    break

                stock_oud = int(tlc.at[i, "Stock"])
                if stock_oud <= 0:
                    continue

                take = min(stock_oud, remaining)
                tlc.at[i, "Stock"] = stock_oud - take

                geleverd_total += take
                remaining -= take

                loc = clean_location(tlc.at[i, "Locatie"]) or "LOCATIE ONTBREEKT"
                alloc_parts.append((loc, take))

                if first_price is None:
                    first_price = float(tlc.at[i, "Prijs"])

            # tekort na afboeken
            tekort = remaining
            if tekort > 0:
                tekort_count += 1

            gevonden = bool(cand_idxs)
            if gevonden:
                found_count += 1
            else:
                not_found_count += 1

            nieuw_total = stock_oud_total - geleverd_total

            # Bouw locatiestring voor picklijst (PAS NA de loop!)
            if len(alloc_parts) == 0:
                loc_used = ""
            elif len(alloc_parts) == 1:
                loc_used = alloc_parts[0][0]  # alleen locatie, geen "(x)"
            else:
                loc_used = " + ".join([f"{loc}({q})" for loc, q in alloc_parts])


            opm_parts = []

            # waarschuwing: meerdere locaties
            if len(all_loc_parts) > 1:
                opm_parts.append("⚠️ Artikel ligt op meerdere locaties")

            if not gevonden:
                opm_parts.append("Niet gevonden in TLC basislijst")
            elif stock_oud_total == 0:
                opm_parts.append("Voorraad was al 0 (maandlijst)")
            elif tekort > 0:
                opm_parts.append(f"Tekort: {tekort} (maandlijst, mogelijk tussentijds verkocht)")

            opm = " | ".join(opm_parts)


            picklijst.append([
                title, oms, gevraagd, geleverd_total, stock_oud_total, nieuw_total,
                loc_used, loc_all, euro(float(first_price or 0.0)), factuur, opm
            ])


            # ---- Uitverkocht debug ----
            if (not gevonden) or (tekort > 0) or (stock_oud_total > 0 and nieuw_total == 0):
                uitverkocht.append([
                    str(r["Title"]).strip(),
                    stock_oud_total,
                    gevraagd,
                    tekort,
                    loc_used,
                    factuur,
                    opm
                ])

        # ----------------------------
        # TLC_NIEUW: alleen > 0
        # ----------------------------
        # ----------------------------
        for _, r in tlc.iterrows():
            if int(r["Stock"]) > 0:
                tlc_nieuw.append([
                    str(r["Title"]).strip(),
                    int(r["Stock"]),
                    euro(float(r["Prijs"])),
                    str(r["Locatie"])
                ])


        print("DEBUG cms_rows:", len(cms_rows))
        print("DEBUG uitverkocht:", len(uitverkocht), "picklijst:", len(picklijst), "tlc_nieuw:", len(tlc_nieuw))

        # ----------------------------
        # Output schrijven (waterproof mappen)
        # ----------------------------
        p_tlc = str(P["active_tlc"])  # overschrijft TLC/TLC_1.xlsx
        p_pick = str(P["pick_dir"] / f"TLC_PICKLIJST_1322_{rid}.xlsx")
        p_out = str(P["debug_dir"] / f"TLC_UITVERKOCHT_1322_{rid}.xlsx")


        tlc_df = pd.DataFrame(
            tlc_nieuw,
            columns=["Title", "Stock", "Prijs", "Locatie"]
        )
        tlc_df["_loc_sort"] = tlc_df["Locatie"].apply(location_sort_key_1322)
        tlc_df = tlc_df.sort_values(by=["_loc_sort", "Title"], ascending=True).drop(columns=["_loc_sort"])
        tlc_df.to_excel(p_tlc, index=False)

        pick_df = pd.DataFrame(picklijst, columns=[
            "Title","Omschrijving","Aantal gevraagd","Aantal geleverd","Voorraad oud",
            "Voorraad nieuw","Locatie","Alle locaties","Prijs","Factuurnummer","Opmerking"
        ])

        pick_df = pick_df.rename(columns={
            "Aantal gevraagd": "Besteld",
            "Aantal geleverd": "Levering",
        })

        # Sorteer: MODEL+D eerst (alleen bij D-locaties), daarna normale locatie-sort
        pick_df["_loc_primary"] = pick_df["Locatie"].apply(_primary_loc_from_multi)
        pick_df["_loc_sort"] = pick_df["_loc_primary"].apply(location_sort_key_1322)

        pick_df = pick_df.sort_values(
            by=["_loc_sort", "Title"],
            ascending=True
        ).drop(columns=["_loc_primary", "_loc_sort"])



        pick_df.to_excel(p_pick, index=False)


        pd.DataFrame(
            uitverkocht,
            columns=[
                "Title", "Voorraad oud", "Aantal gevraagd",
                "Tekort", "Locatie", "Factuurnummer", "Reden"
            ]
        ).to_excel(p_out, index=False)
        
        # ----------------------------
        # Run stats (controle) -> naar DEBUG folder
        # ----------------------------
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            p_stats = os.path.join(str(P["debug_dir"]), f"RUN_STATS_{ts}.xlsx")

            stats = [
                ["CMS bestanden", len(self.cms_paths)],
                ["Orderregels (raw)", int(raw_lines_count)],
                ["Regels na samenvoegen", int(merged_lines_count if 'merged_lines_count' in locals() else len(cms_rows))],
                ["Dubbele regels samengevoegd", int(duplicates_merged if 'duplicates_merged' in locals() else 0)],
                ["Gevonden in TLC basislijst", int(found_count)],
                ["Niet gevonden in TLC basislijst", int(not_found_count)],
                ["Regels met tekort", int(tekort_count)],
            ]
            pd.DataFrame(stats, columns=["Metric", "Value"]).to_excel(p_stats, index=False)
        except Exception:
            pass


        self._write_runlog(P["debug_dir"], rid, {
            "run_id": rid,
            "active_tlc": p_tlc,
            "backup_tlc": str(P["backup_tlc"]),
            "orders_count": len(self.cms_paths),
            "orders_files": ", ".join([Path(x).name for x in self.cms_paths]),
            "picklijst": p_pick,
            "uitverkocht": p_out,
        })


        return [p_tlc, p_pick, p_out]
    def set_wc_df(self, wc_df):
        """Zet WC-export DataFrame voor model/categorie lookup (optioneel)."""
        self.wc_df = wc_df


