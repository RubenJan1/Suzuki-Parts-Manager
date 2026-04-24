# tabs/tab_inboeken.py
"""
Suzuki Parts Manager — Tab Inboeken (V17)
UX-focus (Nielsen heuristics / foutpreventie):
- Form-validatie live: Save is disabled totdat het 'echt klopt'
- Duidelijke status bovenin (wat mist er nog?)
- Enter op Title = popup-overzicht (geen dropdown)
- Quicksearch = popup-overzicht
- Categorieboom uit Woo JSON structuur, elke node individueel selecteerbaar
  (GEEN tristate / GEEN bulk-toggle)
- Zedder/Reiner detectie: merge + blijft staan

Belangrijk voor WP All Import:
- We exporteren 'Productcategorieën' als volledige paden gescheiden door '|'
  zodat Woo categorieën exact matchen.
"""

from __future__ import annotations
from datetime import date
from typing import List, Optional
import os
import subprocess
import sys
import traceback
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QTextEdit, QPushButton, QFileDialog,
    QMessageBox, QTreeWidget, QTreeWidgetItem, QSplitter, QDialog, QTableWidget,
    QTableWidgetItem, QAbstractItemView, QHeaderView, QGroupBox, QMenu
)

from engines.engine_inboeken import InboekenEngine, SearchHit, CATEGORIES_TREE, parse_price, round_up_to_5cent
from utils.theme import apply_theme

SEARCH_LIMIT = 500


class ChooseHitDialog(QDialog):
    def __init__(self, hits: List[SearchHit], parent=None, title: str = "Zoekresultaten"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.hits = hits
        self.selected_hit: Optional[SearchHit] = None

        self.current_wc_id: Optional[str] = None  # gevuld als je een product uit WC export laadt
        self._force_new: bool = False
        self._editing_product_id: Optional[str] = None
       
        lay = QVBoxLayout(self)

        info = QLabel("Dubbelklik of selecteer een regel en klik 'Kies'.")
        lay.addWidget(info)

        self.table = QTableWidget(self)
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Source", "Title", "Stock", "Prijs", "Locatie", "Beschrijving", "Categorieën"
        ])
        self.table.setRowCount(len(hits))
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for i, h in enumerate(hits):
            self.table.setItem(i, 0, QTableWidgetItem(getattr(h, "source", "")))
            self.table.setItem(i, 1, QTableWidgetItem(getattr(h, "title", "")))
            self.table.setItem(i, 2, QTableWidgetItem(str(getattr(h, "stock", ""))))
            self.table.setItem(i, 3, QTableWidgetItem(f"{getattr(h, 'prijs', 0.0):.2f}"))
            self.table.setItem(i, 4, QTableWidgetItem(getattr(h, "locatie", "")))
            self.table.setItem(i, 5, QTableWidgetItem(getattr(h, "short_description", "")))
            self.table.setItem(i, 6, QTableWidgetItem("|".join(getattr(h, "category_paths", []) or [])))

        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.Stretch)
        hdr.setSectionResizeMode(6, QHeaderView.Stretch)

        self.table.doubleClicked.connect(self._accept_selected)

        lay.addWidget(self.table)

        row = QHBoxLayout()
        btn_ok = QPushButton("Kies")
        btn_cancel = QPushButton("Annuleren")
        btn_ok.clicked.connect(self._accept_selected)
        btn_cancel.clicked.connect(self.reject)
        row.addStretch(1)
        row.addWidget(btn_ok)
        row.addWidget(btn_cancel)
        lay.addLayout(row)

    def _accept_selected(self):
        r = self.table.currentRow()
        if r < 0:
            return
        self.selected_hit = self.hits[r]
        self.accept()


class TabInboeken(QWidget):
    def __init__(self, app_state, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self.engine = InboekenEngine()
        # --- WC ID state (altijd bestaan, ook als je start met nieuw product) ---
        self.current_wc_id: Optional[str] = None
        self._force_new: bool = False
        self.selected_hit = None
        self._editing_product_id: Optional[str] = None

        if getattr(self.app_state, "wc_df", None) is not None:
            self.engine.set_website_df(self.app_state.wc_df)

        # categorie-selectie is lijst van FULL PATHS (zoals Woo)
        self.selected_category_paths: set[str] = set()
        self._item_by_path: dict[str, QTreeWidgetItem] = {}

        self._build_ui()
        self._install_solid_context_menus()
        self._refresh_validation()

        try:
            from services.superseded import preload_async
            preload_async()
        except Exception:
            pass

    def _clean_text(self, value) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        if text.lower() in ("nan", "none", "null"):
            return ""
        return text

    # ---------------- UI ----------------

    def _show_solid_context_menu(self, widget, pos):
        menu = widget.createStandardContextMenu()
        menu.setAttribute(Qt.WA_TranslucentBackground, False)
        menu.setWindowOpacity(1.0)
        menu.setStyleSheet("""
            QMenu {
                background-color: #f4f4f4;
                border: 1px solid #8a8a8a;
                padding: 4px;
                color: #111111;
            }
            QMenu::item {
                background-color: transparent;
                padding: 6px 24px 6px 24px;
                margin: 1px;
            }
            QMenu::item:selected {
                background-color: #d9d9d9;
                color: #111111;
            }
            QMenu::separator {
                height: 1px;
                background: #bdbdbd;
                margin: 4px 8px;
            }
        """)
        menu.exec(widget.mapToGlobal(pos))

    def _install_solid_context_menus(self):
        widgets = [
            getattr(self, "ed_cat_filter", None),
            getattr(self, "ed_quicksearch", None),
            getattr(self, "ed_title", None),
            getattr(self, "ed_stock", None),
            getattr(self, "ed_price", None),
            getattr(self, "ed_location", None),
            getattr(self, "ed_desc", None),
            getattr(self, "zedder_text", None),
            getattr(self, "log", None),
        ]
        for widget in widgets:
            if widget is None:
                continue
            widget.setContextMenuPolicy(Qt.CustomContextMenu)
            widget.customContextMenuRequested.connect(
                lambda pos, w=widget: self._show_solid_context_menu(w, pos)
            )

    def _build_ui(self):
        apply_theme(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # =========================
        # TITEL
        # =========================
        title = QLabel("Inboeken — Producten beheren")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        root.addWidget(title)

        # =========================
        # STATUS
        # =========================
        self.lbl_status = QLabel("")
        f = QFont()
        f.setBold(True)
        self.lbl_status.setFont(f)
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setStyleSheet("""
            QLabel {
                background: palette(base);
                border: 1px solid palette(mid);
                border-radius: 8px;
                padding: 10px;
            }
        """)
        root.addWidget(self.lbl_status)

        # =========================
        # HOOFDWERKGEBIED
        # =========================
        splitter = QSplitter(Qt.Horizontal)
        root.addWidget(splitter, stretch=1)

        # ======================================================
        # LINKS — ZOEKEN + CATEGORIEËN
        # ======================================================
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setSpacing(10)
        ll.setContentsMargins(0, 0, 0, 0)

        # -------------------------
        # Zoeken / Laden
        # -------------------------
        gb_search = QGroupBox("Zoeken / Laden")
        sbl = QVBoxLayout(gb_search)

        r2 = QHBoxLayout()
        r2.addWidget(QLabel("Zoeken op alles:"))
        self.ed_quicksearch = QLineEdit()
        self.ed_quicksearch.setPlaceholderText("Zoek op nummer, titel, beschrijving of locatie")
        self.ed_quicksearch.returnPressed.connect(self.do_quicksearch)
        r2.addWidget(self.ed_quicksearch)

        btn2 = QPushButton("Popup")
        btn2.setObjectName("secondary")
        btn2.clicked.connect(self.do_quicksearch)
        r2.addWidget(btn2)

        sbl.addLayout(r2)
        ll.addWidget(gb_search)

        # -------------------------
        # Categorieën
        # -------------------------
        gb_cat = QGroupBox("Categorieën (Woo structuur)")
        gbl = QVBoxLayout(gb_cat)

        row = QHBoxLayout()
        row.addWidget(QLabel("Filter:"))

        self.ed_cat_filter = QLineEdit()
        self.ed_cat_filter.setPlaceholderText("Typ om te filteren (bv. GT750, 2-takt, GSX...)")
        self.ed_cat_filter.textChanged.connect(self._apply_tree_filter)
        row.addWidget(self.ed_cat_filter)

        btn_clear_filter = QPushButton("X")
        btn_clear_filter.setObjectName("secondary")
        btn_clear_filter.clicked.connect(lambda: self.ed_cat_filter.setText(""))
        row.addWidget(btn_clear_filter)

        gbl.addLayout(row)

        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setUniformRowHeights(True)
        self.tree.setAnimated(False)
        self.tree.itemChanged.connect(self.on_tree_item_changed)
        gbl.addWidget(self.tree)

        ll.addWidget(gb_cat, stretch=1)

        splitter.addWidget(left)

        # ======================================================
        # RECHTS — PRODUCTGEGEVENS + ACTIES + EXTRA + LOG
        # ======================================================
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setSpacing(10)
        rl.setContentsMargins(0, 0, 0, 0)

        # -------------------------
        # Productgegevens
        # -------------------------

        gb_data = QGroupBox("Productgegevens")
        dbl = QVBoxLayout(gb_data)

        r_title = QHBoxLayout()
        r_title.addWidget(QLabel("Title (onderdeelnummer):"))

        self.ed_title = QLineEdit()
        self.ed_title.textChanged.connect(self._refresh_validation)
        self.ed_title.setPlaceholderText("bv. 10001-18804")
        self.ed_title.returnPressed.connect(self.do_search_title)
        r_title.addWidget(self.ed_title)

        btn_load = QPushButton("Zoek / Laad")
        btn_load.setObjectName("primary")
        btn_load.clicked.connect(self.do_search_title)
        r_title.addWidget(btn_load)

        dbl.addLayout(r_title)

        r3 = QHBoxLayout()
        r3.addWidget(QLabel("Voorraad:"))

        self.ed_stock = QLineEdit()
        self.ed_stock.textChanged.connect(self._refresh_validation)
        r3.addWidget(self.ed_stock)

        r3.addWidget(QLabel("Prijs:"))

        self.ed_price = QLineEdit()
        self.ed_price.textChanged.connect(self._refresh_validation)
        r3.addWidget(self.ed_price)

        btn_kort_11 = QPushButton("-11%")
        btn_kort_11.setObjectName("secondary")
        btn_kort_11.setFixedWidth(55)
        btn_kort_11.clicked.connect(lambda: self.apply_discount(11))
        r3.addWidget(btn_kort_11)

        r3.addWidget(QLabel("Locatie:"))

        self.ed_location = QLineEdit()
        self.ed_location.setPlaceholderText("bv. D33, PD12, GB 1, K 13")
        self.ed_location.textChanged.connect(self._refresh_validation)
        r3.addWidget(self.ed_location)

        dbl.addLayout(r3)

        dbl.addWidget(QLabel("Korte beschrijving:"))

        self.ed_desc = QTextEdit()
        self.ed_desc.setMinimumHeight(90)
        self.ed_desc.textChanged.connect(self._refresh_validation)
        dbl.addWidget(self.ed_desc)

        rl.addWidget(gb_data)

        self.lbl_selected = QLabel("Geselecteerde categorieën: (0)")
        self.lbl_selected.setWordWrap(True)
        rl.addWidget(self.lbl_selected)

        # -------------------------
        # Acties op product
        # -------------------------
        row_actions = QHBoxLayout()

        self.btn_save = QPushButton("Toevoegen / Updaten")
        self.btn_save.setObjectName("primary")
        self.btn_save.clicked.connect(self.on_add_update)
        row_actions.addWidget(self.btn_save)

        btn_clear = QPushButton("Leegmaken")
        btn_clear.setObjectName("secondary")
        btn_clear.clicked.connect(self.on_clear)
        row_actions.addWidget(btn_clear)

        row_actions.addStretch(1)
        rl.addLayout(row_actions)

        # -------------------------
        # Bestanden / output
        # -------------------------
        row_files = QHBoxLayout()

        btn_save_out = QPushButton("Output opslaan…")
        btn_save_out.setObjectName("secondary")
        btn_save_out.clicked.connect(self.on_save_output)
        row_files.addWidget(btn_save_out)

        btn_open_out = QPushButton("Open output map")
        btn_open_out.setObjectName("secondary")
        btn_open_out.clicked.connect(
            lambda: self._open_output_folder(self.engine.output_dir)
        )
        row_files.addWidget(btn_open_out)

        row_files.addStretch(1)
        rl.addLayout(row_files)

        # -------------------------
        # Zedder / Reiner
        # -------------------------
        gb_ext = QGroupBox("Zedder / Reiner")
        exl = QVBoxLayout(gb_ext)

        exl.addWidget(QLabel("Zedder (plakken):"))

        self.zedder_text = QTextEdit()
        self.zedder_text.setMinimumHeight(110)
        exl.addWidget(self.zedder_text)

        rz = QHBoxLayout()

        btn_z_models = QPushButton("Zedder → Categorieën")
        btn_z_models.setObjectName("secondary")
        btn_z_models.clicked.connect(self.on_zedder_categories)

        btn_z_fill = QPushButton("Zedder → Title + Desc")
        btn_z_fill.setObjectName("secondary")
        btn_z_fill.clicked.connect(self.on_zedder_fill)

        btn_reiners = QPushButton("Reiner → Prijs + Categorieën")
        btn_reiners.setObjectName("secondary")
        btn_reiners.clicked.connect(self.on_reiners)

        rz.addWidget(btn_z_models)
        rz.addWidget(btn_z_fill)
        rz.addWidget(btn_reiners)
        rz.addStretch(1)

        exl.addLayout(rz)
        rl.addWidget(gb_ext)

        # -------------------------
        # Log
        # -------------------------
        rl.addWidget(QLabel("Log"))

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(120)
        rl.addWidget(self.log)

        splitter.addWidget(right)

        # Verhoudingen: links voldoende ruimte voor zoeken + categorieën
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        self._build_tree_from_json(CATEGORIES_TREE)
        self._update_selected_label()
    # ------------- Categories Tree -------------

    def _build_tree_from_json(self, nodes: list):
        """Bouw de Woo categorieboom.
        Belangrijk: elke node is individueel checkbaar (GEEN tristate, GEEN cascading).
        """
        self.tree.blockSignals(True)
        try:
            self.tree.clear()
            self._item_by_path = {}

            def add_node(parent, node, prefix):
                name = str(node.get("name","")).strip()
                if not name:
                    return
                children = node.get("children", []) or []
                path_list = prefix + [name]
                path_str = " > ".join(path_list)

                it = QTreeWidgetItem([name])
                it.setData(0, Qt.UserRole, path_str)

                # Individueel checkbaar (ook als er children zijn), MAAR: geen Qt tri-state gedrag
                flags = it.flags() | Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable

                # Qt6/Qt5: haal (auto) tri-state weg zodat parent/child nooit automatisch mee togglen
                if hasattr(Qt, "ItemIsTristate"):
                    flags = flags & ~Qt.ItemIsTristate
                if hasattr(Qt, "ItemIsAutoTristate"):
                    flags = flags & ~Qt.ItemIsAutoTristate

                it.setFlags(flags)
                it.setCheckState(0, Qt.Unchecked)

                if parent is None:
                    self.tree.addTopLevelItem(it)
                else:
                    parent.addChild(it)

                self._item_by_path[path_str] = it

                for c in children:
                    add_node(it, c, path_list)

                # default open tot 3 levels voor snelheid
                if len(path_list) <= 3:
                    it.setExpanded(True)

            for n in nodes:
                add_node(None, n, [])

        finally:
            self.tree.blockSignals(False)

    def _apply_tree_filter(self, text: str):
        q = (text or "").strip().lower()

        def match_item(it: QTreeWidgetItem) -> bool:
            name = (it.text(0) or "").lower()
            path = (it.data(0, Qt.UserRole) or "").lower()
            return (q in name) or (q in path)

        def walk(it: QTreeWidgetItem) -> bool:
            any_child_visible = False
            for i in range(it.childCount()):
                if walk(it.child(i)):
                    any_child_visible = True
            visible = True if not q else (match_item(it) or any_child_visible)
            it.setHidden(not visible)
            if visible and q and match_item(it):
                it.setExpanded(True)
            return visible

        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i))

    def on_toggle_export_mode(self):
        self.engine.export_leaf_names = bool(self.chk_leaf_export.isChecked())
        mode = "LEAF (GT750|T350)" if self.engine.export_leaf_names else "FULL PATH (A > B > C)"
        self._log(f"EXPORT MODE | {mode}")

    def on_save_output(self):
        if self.engine.df is None or self.engine.df.empty:
            QMessageBox.information(self, "Output opslaan", "Er is nog geen output. Boek eerst minstens 1 regel in.")
            return
        today = date.today().isoformat()
        default_name = f"inboeken_output_{today}.xlsx"
        out_dir = self.engine.output_dir
        default_path = os.path.join(out_dir, default_name)

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Output opslaan",
            default_path,
            "Excel (*.xlsx);;CSV (*.csv)"
        )
        if not path:
            return
        try:
            saved = self.engine.export_output(path)
            self.engine.clear_autosave()
            self._log(f"OUTPUT | opgeslagen: {saved}")
            self._log("AUTOSAVE | gewist — volgende sessie begint schoon")
            QMessageBox.information(
                self,
                "Output opgeslagen",
                f"Opgeslagen:\n{saved}\n\nDe sessielijst is gewist. Je kunt weer vers beginnen."
            )
        except Exception:
            QMessageBox.warning(self, "Opslaan mislukt", "Kan output niet opslaan.")

    def on_tree_item_changed(self, item, col):
        path = item.data(0, Qt.UserRole)
        if not path:
            return

        if item.checkState(0) == Qt.Checked:
            self.selected_category_paths.add(path)
        else:
            self.selected_category_paths.discard(path)

        self._update_selected_label()
        self._refresh_validation()



    def _sync_tree_checks(self):
        self.tree.blockSignals(True)
        try:
            for path, it in self._item_by_path.items():
                if not (it.flags() & Qt.ItemIsUserCheckable):
                    continue
                it.setCheckState(0, Qt.Checked if path in self.selected_category_paths else Qt.Unchecked)
        finally:
            self.tree.blockSignals(False)


    def _update_selected_label(self):
        paths = sorted(self.selected_category_paths)
        self.lbl_selected.setText(f"Geselecteerde categorieën: ({len(paths)})")
        self._log("CATS | " + (" | ".join(paths[:8]) + (" ..." if len(paths) > 8 else "")))
    
    def _set_field_state(self, widget, state: str):
        """
        state: 'ok' | 'bad' | 'neutral'
        """
        if state == "bad":
            widget.setStyleSheet("""
                QLineEdit, QTextEdit {
                    border: 2px solid #e74c3c;
                    border-radius: 6px;
                    background-color: rgba(231, 76, 60, 0.10);
                    color: palette(text);
                }
            """)
        elif state == "ok":
            widget.setStyleSheet("""
                QLineEdit, QTextEdit {
                    border: 2px solid #2ecc71;
                    border-radius: 6px;
                    background-color: rgba(46, 204, 113, 0.08);
                    color: palette(text);
                }
            """)
        else:
            widget.setStyleSheet("")  # terug naar default theme


    # ------------- Validation (error-prevention) -------------

    def _refresh_validation(self):
        """Live validatie: Save-knop alleen aan als de invoer consistent is."""
        title = (self.ed_title.text() or "").strip()
        stock_txt = (self.ed_stock.text() or "").strip()
        price_txt = (self.ed_price.text() or "").strip()
        loc = (self.ed_location.text() or "").strip()
        desc = (self.ed_desc.toPlainText() or "").strip()

        problems = []

        # -----------------------------
        # Helper: simpele field highlight
        # (werkt ook zonder extra helper-functies)
        # -----------------------------
        def set_bad(w):
            w.setStyleSheet("""
                border: 2px solid #e74c3c;
                border-radius: 6px;
                background-color: rgba(231, 76, 60, 0.10);
                color: palette(text);
            """)

        def set_ok(w):
            w.setStyleSheet("""
                border: 2px solid #2ecc71;
                border-radius: 6px;
                background-color: rgba(46, 204, 113, 0.08);
                color: palette(text);
            """)

        def set_neutral(w):
            w.setStyleSheet("")

        # reset eerst alles
        for w in (self.ed_title, self.ed_stock, self.ed_price, self.ed_location, self.ed_desc):
            set_neutral(w)

        # -----------------------------
        # Title 

        # -----------------------------
        if not title:
            problems.append("Title ontbreekt")
            set_bad(self.ed_title)
        else:
            set_ok(self.ed_title)

        # -----------------------------
        # Stock (verplicht; 0 mag, maar niet leeg)
        # -----------------------------
        stock = None
        if not stock_txt:
            problems.append("Voorraad is verplicht (0 mag, maar niet leeg)")
            set_bad(self.ed_stock)
        else:
            try:
                stock = int(float(stock_txt.replace(",", ".")))
                set_ok(self.ed_stock)
            except Exception:
                stock = None
                problems.append("Voorraad is geen geldig getal")
                set_bad(self.ed_stock)

        # -----------------------------
        # Categorieën verplicht
        # -----------------------------
        if not self.selected_category_paths:
            problems.append("Geen categorieën geselecteerd")

        # -----------------------------
        # Bij voorraad ≠ 0: locatie + prijs + beschrijving verplicht
        # -----------------------------
        if stock is not None and stock != 0:
            # Locatie
            if not loc:
                problems.append("Locatie verplicht bij voorraad ≠ 0")
                set_bad(self.ed_location)
            else:
                set_ok(self.ed_location)

            # Prijs
            if not price_txt:
                problems.append("Prijs verplicht bij voorraad ≠ 0")
                set_bad(self.ed_price)
            else:
                try:
                    p = float(price_txt.replace(",", ".").replace("€", "").strip())
                    if p <= 0:
                        problems.append("Prijs moet > 0 bij voorraad ≠ 0")
                        set_bad(self.ed_price)
                    else:
                        set_ok(self.ed_price)
                except Exception:
                    problems.append("Prijs is geen geldig getal")
                    set_bad(self.ed_price)

            # Beschrijving
            if not desc:
                problems.append("Korte beschrijving is verplicht bij voorraad ≠ 0")
                set_bad(self.ed_desc)
            else:
                set_ok(self.ed_desc)
        else:
            # voorraad = 0 of None → locatie/prijs/desc zijn niet verplicht in UI-validatie
            # (engine zal bij save alsnog streng zijn waar nodig)
            if loc:
                set_ok(self.ed_location)
            if price_txt:
                # alleen groen maken als parsebaar
                try:
                    float(price_txt.replace(",", ".").replace("€", "").strip())
                    set_ok(self.ed_price)
                except Exception:
                    set_bad(self.ed_price)
            if desc:
                set_ok(self.ed_desc)

        # -----------------------------
        # Eindresultaat
        # -----------------------------
        ok = len(problems) == 0

        if ok:
            self.lbl_status.setText("✅ Klaar om op te slaan")
            self.lbl_status.setStyleSheet("color: #1b7f3a;")
        else:
            self.lbl_status.setText(
                "⚠️ Mist nog: " + " • ".join(problems[:3]) + (" …" if len(problems) > 3 else "")
            )
            self.lbl_status.setStyleSheet("color: #b00020;")

        self.btn_save.setEnabled(ok)


    # ------------- Search -------------

    def _choose_from_hits(self, hits: List[SearchHit], title: str):
        if len(hits) == 1:
            return hits[0]
        dlg = ChooseHitDialog(hits, self, title=title)
        if dlg.exec() == QDialog.Accepted and dlg.selected_hit:
            return dlg.selected_hit
        return None

    def do_search_title(self):
        self._force_new = False
        q = (self.ed_title.text() or "").strip()
        if not q:
            return
        hits = self.engine.exact_title_hits(q, limit=SEARCH_LIMIT)
        if not hits:
            hits = self.engine.search(q, limit=SEARCH_LIMIT)

        if not hits:
            st = self.engine.not_found_status(q)
            self._log(f"NIET GEVONDEN | {q} | website={st['in_web']} eigen={st['in_own']}")
            sup_hits = self._search_via_superseded(q)
            if sup_hits:
                chosen = self._choose_from_hits(sup_hits, title=f"Superseded nummer gevonden voor: {q} ({len(sup_hits)})")
                if chosen:
                    self._apply_hit(chosen)
                return
            QMessageBox.information(self, "Zoek", f"Niets gevonden voor: {q}")
            self._fill_superseded()
            return

        chosen = self._choose_from_hits(hits, title=f"Zoekresultaten ({len(hits)})")
        if chosen:
            self._apply_hit(chosen)

    def do_quicksearch(self):
        self._force_new = False
        q = (self.ed_quicksearch.text() or "").strip()
        if not q:
            return
        hits = self.engine.search(q, limit=SEARCH_LIMIT)
        if not hits:
            self._log(f"NIET GEVONDEN | {q} (quicksearch)")
            QMessageBox.information(self, "Quicksearch", f"Niets gevonden voor: {q}")
            return
        chosen = self._choose_from_hits(hits, title=f"Quicksearch ({len(hits)})")
        if chosen:
            self._apply_hit(chosen)

    def _apply_hit(self, hit: SearchHit):
        data = self.engine.load_product(hit)

        self._force_new = False
        self.current_wc_id = None
        self._editing_product_id = None

        try:
            if hit.row is not None:
                found_id = str(hit.row.get("ID", "")).strip() or None
                self.current_wc_id = found_id
                self._editing_product_id = found_id
        except Exception:
            self.current_wc_id = None
            self._editing_product_id = None

        self.ed_title.setText(self._clean_text(data.get("Title", "")))
        self.ed_stock.setText(self._clean_text(data.get("Stock", "")))
        self.ed_price.setText(self._clean_text(data.get("Prijs", "")))
        self.ed_location.setText(self._clean_text(data.get("Locatie", "")))
        self.ed_desc.setPlainText(self._clean_text(data.get("Short Description", "")))

        self.selected_category_paths.clear()

        for p in data.get("SelectedCategoryPaths", []) or []:
            if p:
                self.selected_category_paths.add(p)

        self._sync_tree_checks()
        self._update_selected_label()
        self._log(f"GELADEN | {data.get('Title','')} ({data.get('Source','')})")
        self._fill_superseded()
        self._refresh_validation()

    def _fill_superseded(self):
        """Zoek superseded nummers op voor de huidige titel en vul de korte beschrijving aan."""
        try:
            from services.superseded import lookup_superseded
            title = (self.ed_title.text() or "").strip()
            if not title:
                return
            related = lookup_superseded(title)
            if not related:
                return
            superseded_line = "Superseded to: " + ", ".join(related)
            current = (self.ed_desc.toPlainText() or "").strip()
            lines = [l for l in current.splitlines() if not l.strip().startswith("Superseded to:")]
            lines = [l for l in lines if l.strip()]
            new_desc = "\n".join(lines + [superseded_line]) if lines else superseded_line
            self.ed_desc.setPlainText(new_desc)
            self.ed_desc.moveCursor(QTextCursor.End)
            self._log(f"SUPERSEDED | {superseded_line}")
        except Exception as e:
            self._log(f"SUPERSEDED | fout: {e}")

    def _search_via_superseded(self, q: str) -> list:
        """Zoek via superseded nummers in de WC export als het originele nummer niet gevonden wordt."""
        try:
            from services.superseded import lookup_superseded
            related = lookup_superseded(q)
            if not related:
                return []
            hits = []
            seen = set()
            for num in related:
                for hit in self.engine.exact_title_hits(num, limit=SEARCH_LIMIT):
                    key = (getattr(hit, "source", ""), getattr(hit, "title", ""))
                    if key not in seen:
                        seen.add(key)
                        hits.append(hit)
            return hits
        except Exception:
            return []

    # ------------- Save -------------

    def on_add_update(self):
        wc_id_to_use = None if self._force_new else self.current_wc_id
        title = (self.ed_title.text() or "").strip()
        stock = (self.ed_stock.text() or "").strip()
        prijs = (self.ed_price.text() or "").strip()
        locatie = (self.ed_location.text() or "").strip()
        short_desc = (self.ed_desc.toPlainText() or "").strip()
        cats = sorted(self.selected_category_paths)
        
        # uitverkocht confirm
        try:
            s_int = int(float(stock.replace(",", "."))) if stock else 0
        except Exception:
            s_int = 0

        if s_int == 0:
            r = QMessageBox.question(
                self,
                "Uitverkocht?",
                "Voorraad is 0. Wil je dit product UITVERKOCHT zetten?\n\nDit zet prijs=0 en locatie leeg.",
                QMessageBox.Yes | QMessageBox.No
            )
            if r == QMessageBox.Yes:
                prijs = "0"
                locatie = ""

        # ✅ engine call + debug bij crash
        try:
            wc_id_to_use = None if self._force_new else self.current_wc_id

            res = self.engine.add_or_update(
                title=title,
                selected_category_paths=cats,
                stock=stock,
                short_description=short_desc,
                locatie=locatie,
                prijs=prijs,
                wc_id=wc_id_to_use,
            )
        except Exception as e:
            # debug log (mag nooit extra crashen)
            try:
                self.engine.debug.bump("errors")
                self.engine.debug.event("ERROR", {
                    "where": "TabInboeken -> add_or_update",
                    "error": str(e),
                    "trace": traceback.format_exc()[:4000],
                })
                self.engine.debug.write_stats_xlsx()
            except Exception:
                pass

            QMessageBox.critical(self, "Fout", str(e))
            self._refresh_validation()
            return

        # ✅ success flow (alleen als res bestaat)
        self._log(f"{res.actie} | {res.title} | stock={res.stock} | €{res.prijs:.2f} | loc={res.locatie}")

        self.current_wc_id = None
        self._force_new = False
        self.selected_hit = None
        self._editing_product_id = None
        self.on_clear()

        self._refresh_validation()
        
    def on_clear(self):
        has_data = any([
            (self.ed_title.text() or "").strip(),
            (self.ed_stock.text() or "").strip(),
            (self.ed_price.text() or "").strip(),
            (self.ed_location.text() or "").strip(),
            (self.ed_desc.toPlainText() or "").strip(),
            bool(self.selected_category_paths),
            (self.zedder_text.toPlainText() or "").strip(),
            (self.ed_quicksearch.text() or "").strip(),
            (self.ed_cat_filter.text() or "").strip(),
        ])

        if has_data:
            r = QMessageBox.question(
                self,
                "Formulier leegmaken?",
                "Alle ingevulde gegevens worden gewist. Doorgaan?",
                QMessageBox.Yes | QMessageBox.No
            )
            if r != QMessageBox.Yes:
                return

        self.current_wc_id = None
        self._force_new = False
        self.selected_hit = None
        self._editing_product_id = None

        self.ed_title.clear()
        self.ed_stock.clear()
        self.ed_price.clear()
        self.ed_location.clear()
        self.ed_desc.clear()
        self.ed_quicksearch.clear()
        self.ed_cat_filter.clear()
        self.zedder_text.clear()

        self.selected_category_paths.clear()
        self._sync_tree_checks()
        self._apply_tree_filter("")
        self._update_selected_label()

        self._refresh_validation()
        self._log("Formulier leeggemaakt: nieuwe invoer gestart.")
        self.lbl_status.setText("Geen product geladen")
    

    def _parse_price_from_ui(self, s: str) -> float:
        s = (s or "").strip()
        if not s:
            return 0.0
        # NL input: "12,34"
        s = s.replace("€", "").strip().replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0

    def apply_discount(self, pct: int) -> None:
        base = parse_price(self.ed_price.text() or "")
        if base <= 0:
            return
        new_price = base * (1.0 - (pct / 100.0))
        new_price = round_up_to_5cent(new_price)
        self.ed_price.setText(f"{new_price:.2f}".replace(".", ","))
        self._refresh_validation()

    # -------- Zedder / Reiner --------

    def on_zedder_fill(self):
        t = self.zedder_text.toPlainText()
        title, desc = self.engine.zedder_fill_title_and_desc(t, current_title=self.ed_title.text())
        if title:
            self.ed_title.setText(title)
        if desc:
            self.ed_desc.setPlainText(desc)
        self._log("ZEDDER | title/desc gevuld")
        self._refresh_validation()

    def on_zedder_categories(self):
        t = self.zedder_text.toPlainText()
        paths = self.engine.zedder_detect_model_category_paths(t)
        before = set(self.selected_category_paths)
        self.selected_category_paths.update(paths)
        self._sync_tree_checks()
        self._update_selected_label()
        added = sorted(self.selected_category_paths - before)
        self._log("ZEDDER | categorieën toegevoegd: " + (", ".join(added) if added else "—"))
        models = getattr(self.engine, "last_zedder_models", set())
        unmapped = getattr(self.engine, "last_zedder_unmapped", set())
        if models:
            self._log("ZEDDER | modellen gevonden: " + ", ".join(sorted(models)))
        if unmapped:
            self._log("ZEDDER | NIET gemapt: " + ", ".join(sorted(unmapped)))

        self._refresh_validation()

    def on_reiners(self):
        pn = (self.ed_title.text() or "").strip()
        if not pn:
            QMessageBox.information(self, "Reiner", "Vul eerst een part number in bij Title.")
            return
        hit = self.engine.reiners_lookup(pn)
        if not hit:
            self._log(f"REINER | niet gevonden: {pn}")
            QMessageBox.information(self, "Reiner", f"Niets gevonden voor: {pn}")
            return

        if "prijs" in hit and hit["prijs"] is not None:
            self.ed_price.setText(f"{hit['prijs']:.2f}".replace(".", ","))

        before = set(self.selected_category_paths)
        self.selected_category_paths.update(hit.get("category_paths", []) or [])
        self._sync_tree_checks()
        self._update_selected_label()
        added = sorted(self.selected_category_paths - before)
        self._log("REINER | gevuld (added: " + (", ".join(added) if added else "—") + ")")
        self._refresh_validation()

    def _open_output_folder(self, folder: str) -> None:
        if not os.path.isdir(folder):
            return

        if sys.platform.startswith("win"):
            os.startfile(folder)
        elif sys.platform == "darwin":
            subprocess.call(["open", folder])
        else:
            subprocess.call(["xdg-open", folder])


        

    # ------------- Log -------------

    def _log(self, msg: str):
        self.log.append(msg)