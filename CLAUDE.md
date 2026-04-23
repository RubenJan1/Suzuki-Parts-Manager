# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this app does

Suzuki Parts Manager is a PySide6 desktop application for **Vlaandere Motoren** (a Suzuki motorcycle parts dealer in Bolsward, NL). It manages parts inventory, generates WP All Import–ready Excel files, processes CMS orders, and produces PDF invoices.

The app is distributed as a Windows `.exe` (PyInstaller). Current version: `2.0.11` (see `version.py`).

## Running the app

```bash
python main.py
```

Dependencies (no requirements.txt – install manually):
```
PySide6, pandas, openpyxl, reportlab, Pillow
```

## Architecture

### Data flow

1. User loads a **WooCommerce CSV export** on the Start tab → stored in `AppState.wc_df` (read-only)
2. All other tabs receive `AppState` and use `wc_df` as source of truth
3. Engines generate output XLSX/PDF files to `%LOCALAPPDATA%\Suzuki Parts Manager\output\`

### Layer separation

| Layer | Location | Role |
|---|---|---|
| UI (tabs) | `tabs/tab_*.py` | PySide6 QWidget, calls engines, shows results |
| Business logic | `engines/engine_*.py` | Pure data processing, no UI dependencies |
| Shared state | `app_state.py` | `AppState` singleton passed to all tabs |
| Services | `services/` | Update checker, batch state, auto-updater, superseded lookup |
| Utils | `utils/` | Paths, theming, helpers |

### AppState

`AppState` is the only shared state object. It holds:
- `wc_df` – loaded WooCommerce DataFrame (read-only source)
- `wc_path` – path of the loaded CSV
- `output_dir` – `%LOCALAPPDATA%\Suzuki Parts Manager\output\`
- `tradelist_path` – optionally set by the Tradelist tab, used by Website 277 engine

### Tabs and their engines

| Tab name | Tab file | Engine file | What it does |
|---|---|---|---|
| Start | `tab_intro.py` | — | Load WC CSV export; unlocks all other tabs |
| Inboeken | `tab_inboeken.py` | `engine_inboeken.py` | Add/update products; exports XLSX for WP All Import |
| Tradelist | `tab_tradelist.py` | `engine_tradelist.py` | Generate ONSZELF/CMS/GoParts tradelists |
| 1322 | `tab_tlc_1322.py` | `engine_tlc_1322.py` | Internal TLC sales (monthly); deducts from TLC_1.xlsx |
| TLC Update | `tab_tlc_update.py` | `engine_tlc_update.py` | Update/replace TLC stock (MERGE or REPLACE_LOC mode) |
| Website 277 | `tab_website_277.py` | `engine_website_277.py` | Deduct stock for CMS 277 orders from WC export |
| Factuurmaker | `tab_factuurmaker.py` | `engine_factuurmaker.py` | PDF invoices/creditnotes (21% VAT, Vlaandere Motoren) |
| Zoeklijst | `tab_zoeklijst.py` | `engine_zoeklijst.py` | Mail parser / parts search against WC + TLC |

## Key domain rules

**WooCommerce column names** (Dutch): `Naam`, `Reguliere prijs`, `Voorraad`, `Korte beschrijving`, `Beschrijving` (used for location!), `Categorieën`. Engines use `_find_col()` to match tolerantly.

**Locatie (location) is stored in the `Beschrijving` (description) column** in WooCommerce – not a dedicated field. Always clean with `clean_location()` / `_clean_loc_text()` before use.

**Category paths** are exported as full WooCommerce paths separated by `|`:
`Originele onderdelen > 2-takt > GT series > GT750|Originele onderdelen > 2-takt > GT series`
The category tree is hardcoded as JSON in both `engine_inboeken.py` and `engine_website_277.py` (the two copies must stay in sync).

**Pattern parts**: Products ending in `-p` or with "Pattern parts" in categories are namaak/aftermarket. The Website 277 engine never deducts stock from pattern products. The Tradelist engine filters them out entirely.

**Uitverkocht (out of stock)**: When stock goes to 0, price is set to 0 and location is cleared in WP All Import output files.

**Pricing**: Tradelist uses `CMS = 86%` and `GoParts = 81%` of the base price.

**TLC (internal stock list)**: Maintained as `TLC_1.xlsx` (no headers, columns: Title, Stock, Prijs, Locatie) inside a structured folder. The engine makes automatic backups before each run.

**Batch state** (`data/batch_state.json`): Tracks which Website 277 or other runs are `PENDING_IMPORT` / `IMPORTED` / `MERGED`. Managed by `services/batch_state.py`.

**Superseded nummers**: Suzuki onderdeelnummers bestaan in families (oud/nieuw model, ander jaar). `services/superseded.py` bouwt een lazy index vanuit `assets/Superseded lijst.xls` (kolom AJ = canoniek ALT-nummer, kolommen R1–R17 en U1–U9 = verwante nummers). `lookup_superseded(part_number)` geeft alle gerelateerde nummers gesorteerd terug, zonder `-000` suffix. De Inboeken-tab laadt de index alvast in de achtergrond via `preload_async()` bij opstarten. Bij zoeken wordt de korte beschrijving automatisch aangevuld met een `Superseded to: ...` regel; als een artikelnummer niet direct gevonden wordt, zoekt de tab alsnog via superseded nummers in de WC export.

**Website 277 – laatste stap**: Na het genereren van het output-bestand opent de tab automatisch de map (`%LOCALAPPDATA%\Suzuki Parts Manager\output\277\`) in Windows Verkenner via `os.startfile()`, zodat de gebruiker het bestand direct kan uploaden naar de website.

## File paths at runtime

```
%LOCALAPPDATA%\Suzuki Parts Manager\
  output\
    tradelist\          # Tradelist engine output
    277\                # Website 277 engine output (+ DEBUG subfolder)
    1322\               # TLC 1322 engine output
    facturen\           # PDF invoices + sequence.json + facturen_log.jsonl
  app.lock              # Single-instance lock
```

Assets (`assets/`) and the executable itself are resolved via `utils/paths.py:resource_path()`, which handles both dev (`__file__`-relative) and PyInstaller (`sys._MEIPASS`) contexts.

## Theming

`utils/theme.py:apply_theme(widget)` detects dark/light mode from the system palette and applies a full QSS stylesheet. Called at the top of each tab's `_build_ui()`. Button variants are controlled via `setObjectName`: `"primary"` (blue), `"secondary"` (grey), `"danger"` (red).

## Auto-update

On startup, `main.py` calls `services/update_checker.py:check_github_release()` against `RubenJan1/Suzuki-Parts-Manager`. If a newer tag is found (semver comparison), the user is prompted to download. `services/auto_updater.py:run_updater()` downloads the zip and launches `updater.exe` (a separate PyInstaller-compiled script from `updater.py`).
