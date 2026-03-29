from pathlib import Path
from datetime import datetime
import pandas as pd

from engines.engine_website_277 import clean_text

def _safe_read_excel(path: str) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Bestand niet gevonden: {path}")
    return pd.read_excel(p)


def load_changes(path: str) -> pd.DataFrame:
    df = _safe_read_excel(path)

    # verwachte kolommen uit CHANGES_277
    expected = [
        "ID", "Title", "Stock_oud", "Besteld", "Geleverd",
        "Stock_nieuw", "Tekort", "Factuur"
    ]

    for col in expected:
        if col not in df.columns:
            if col == "Factuur":
                df[col] = ""
            else:
                raise RuntimeError(f"Kolom ontbreekt in CHANGES bestand: {col}")

    df["ID"] = df["ID"].astype(str).str.strip()
    df["Title"] = df["Title"].astype(str).str.strip()

    for col in ["Stock_oud", "Besteld", "Geleverd", "Stock_nieuw", "Tekort"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df["Factuur"] = df["Factuur"].astype(str).fillna("").str.strip()
    return df

def format_categories_for_wpimport(raw_categories: str) -> str:
    """
    Zet WC categoriepad om naar WP All Import formaat.

    Voorbeeld input:
    "Originele onderdelen > 2-takt > TS > TS100 / TS125 / TS125X"

    Output:
    "2-takt|TS|TS100 / TS125 / TS125X"

    Regels:
    - meerdere paden uit WC zijn vaak komma-gescheiden
    - per pad halen we de hiërarchie eruit
    - 'Originele onderdelen' laten we weg
    - dubbele stukken voorkomen
    - volgorde behouden
    """
    if not raw_categories:
        return ""

    ignore_names = {
        "originele onderdelen",
        "motoren te koop",
        "hot parts",
        "verschillende merken",
    }

    out = []
    seen = set()

    # WC export heeft vaak meerdere paden gescheiden door komma's
    for path in str(raw_categories).split(","):
        parts = [clean_text(p).strip() for p in str(path).split(">")]
        parts = [p for p in parts if p]

        for p in parts:
            if p.lower() in ignore_names:
                continue
            key = p.lower()
            if key not in seen:
                seen.add(key)
                out.append(p)

    return "|".join(out)


def pick_best_short_description(series) -> str:
    """
    Pak de beste short description uit de WC data:
    - HTML eruit
    - lege waarden negeren
    - langste bruikbare tekst kiezen
    """
    vals = []
    for x in series:
        s = clean_text(x)
        if s.strip():
            vals.append(s)

    if not vals:
        return ""

    return max(vals, key=len)


def merge_changes(old_df: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    """
    Merge op product-ID.
    Regels:
    - laatste Stock_nieuw wint
    - eerste Stock_oud bewaren uit oudste batch
    - Besteld / Geleverd / Tekort optellen
    - Factuur samenvoegen uniek
    - Title van nieuwste regel wint als die gevuld is
    """
    old_df = old_df.copy()
    new_df = new_df.copy()

    old_df["_source_order"] = 1
    new_df["_source_order"] = 2

    all_df = pd.concat([old_df, new_df], ignore_index=True)

    rows = []
    for pid, grp in all_df.groupby("ID", dropna=False):
        grp = grp.sort_values("_source_order")
        first_row = grp.iloc[0]
        last_row = grp.iloc[-1]

        facturen = []
        for f in grp["Factuur"].tolist():
            if not f:
                continue
            for part in str(f).split(","):
                part = part.strip()
                if part and part not in facturen:
                    facturen.append(part)

        row = {
            "ID": str(pid).strip(),
            "Title": str(last_row.get("Title", "") or "").strip() or str(first_row.get("Title", "") or "").strip(),
            "Stock_oud": int(first_row.get("Stock_oud", 0)),
            "Besteld": int(grp["Besteld"].sum()),
            "Geleverd": int(grp["Geleverd"].sum()),
            "Stock_nieuw": int(last_row.get("Stock_nieuw", 0)),
            "Tekort": int(grp["Tekort"].sum()),
            "Factuur": ", ".join(facturen),
        }
        rows.append(row)

    merged = pd.DataFrame(rows)

    if not merged.empty:
        merged = merged.sort_values(["Title", "ID"], kind="stable").reset_index(drop=True)

    return merged


def build_update_from_changes(changes_df: pd.DataFrame, wc_df: pd.DataFrame) -> pd.DataFrame:
    """
    Bouw merged website-update in exact dezelfde layout als de gewone 277 export:

    ID | Title | Productcategorieën | Stock | Short Description | Locatie | Prijs

    Extra regels:
    - Short Description HTML-cleanen
    - Productcategorieën naar WP All Import pad-formaat met |
    - bij stock 0 -> Locatie en Prijs leeg
    """
    wc = wc_df.copy()
    wc.columns = [str(c).strip() for c in wc.columns]

    # tolerante kolomdetectie
    col_id = None
    for c in ["ID", "Id", "post_id", "Post ID", "Product ID"]:
        if c in wc.columns:
            col_id = c
            break
    if not col_id:
        raise RuntimeError("Geen ID kolom gevonden in WC-export voor merge")

    col_title = None
    for c in ["Title", "Naam", "post_title"]:
        if c in wc.columns:
            col_title = c
            break
    if not col_title:
        raise RuntimeError("Geen Title kolom gevonden in WC-export voor merge")

    col_cat = None
    for c in ["Categorieën", "Productcategorieën"]:
        if c in wc.columns:
            col_cat = c
            break
    if not col_cat:
        raise RuntimeError("Geen categorie kolom gevonden in WC-export voor merge")

    col_short = None
    for c in ["Korte beschrijving", "Short description", "Short Description"]:
        if c in wc.columns:
            col_short = c
            break
    if not col_short:
        raise RuntimeError("Geen short description kolom gevonden in WC-export voor merge")

    col_locatie = None
    for c in ["Beschrijving", "Locatie"]:
        if c in wc.columns:
            col_locatie = c
            break
    if not col_locatie:
        raise RuntimeError("Geen locatie kolom gevonden in WC-export voor merge")

    col_price = None
    for c in ["Reguliere prijs", "Prijs"]:
        if c in wc.columns:
            col_price = c
            break
    if not col_price:
        raise RuntimeError("Geen prijs kolom gevonden in WC-export voor merge")

    wc[col_id] = wc[col_id].astype(str).str.strip()
    changes = changes_df.copy()
    changes["ID"] = changes["ID"].astype(str).str.strip()

    # Alleen producten meenemen die in merged changes zitten
    merged = wc.merge(
        changes[["ID", "Stock_nieuw"]],
        left_on=col_id,
        right_on="ID",
        how="inner"
    )

    rows = []
    for _, r in merged.iterrows():
        prod_id = str(r[col_id]).strip()
        title = clean_text(r[col_title])
        categories_clean = format_categories_for_wpimport(r[col_cat])
        stock_new = int(r["Stock_nieuw"])
        short_desc = clean_text(r[col_short])
        locatie = clean_text(r[col_locatie])
        prijs = r[col_price]

        # zelfde regel als gewone export
        if stock_new == 0:
            locatie = ""
            prijs = ""

        rows.append([
            prod_id,
            title,
            categories_clean,
            stock_new,
            short_desc,
            locatie,
            prijs,
        ])

    upd_df = pd.DataFrame(
        rows,
        columns=[
            "ID",
            "Title",
            "Productcategorieën",
            "Stock",
            "Short Description",
            "Locatie",
            "Prijs",
        ]
    )

    # zelfde sortering: locatie -> title
    try:
        def _loc_sort_key(val):
            s = "" if pd.isna(val) else str(val).strip().upper()
            return s

        upd_df["_loc_sort"] = upd_df["Locatie"].apply(_loc_sort_key)
        upd_df = upd_df.sort_values(
            by=["_loc_sort", "Title"],
            ascending=True
        ).drop(columns=["_loc_sort"])
    except Exception:
        pass

    return upd_df

def save_merged_files(
    merged_changes_df: pd.DataFrame,
    merged_update_df: pd.DataFrame,
    output_dir: str = "output/277"
):
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    debug_dir = out_dir / "DEBUG"
    debug_dir.mkdir(parents=True, exist_ok=True)

    rid = datetime.now().strftime("%Y%m%d_%H%M%S")
    update_path = out_dir / f"WEBSITE_UPDATE_277_MERGED_{rid}.xlsx"
    changes_path = debug_dir / f"CHANGES_277_MERGED_{rid}.xlsx"

    merged_update_df.to_excel(update_path, index=False)
    merged_changes_df.to_excel(changes_path, index=False)

    return {
        "batch_id": f"277_MERGED_{rid}",
        "update_path": str(update_path),
        "debug_changes_path": str(changes_path),
        "rid": rid,
    }