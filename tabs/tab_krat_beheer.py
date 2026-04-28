# tabs/tab_krat_beheer.py
"""
Krat Beheer — Twee-fase inventarisatie en beprijzing

Fase 1  Inventarisatie (medewerker, zonder baas)
  • Krat aanmaken met naam + locatie (WC Beschrijving-waarde)
  • Per artikel: nummer + omschrijving + categorie + voorraad
  • WC-check: bestaat al? Nieuw product of update bestaande?

Fase 2  Beprijzing (samen met baas)
  • Artikelen één voor één, in inventarisatievolgorde
  • Keyboard-first: Enter = opslaan en volgende
  • Voortgang opgeslagen na elk artikel (pauze/hervatten)

Fase 3  Export
  • Export A: nieuwe WC producten  (WP All Import XLSX)
  • Export B: update bestaande WC producten (samenvoeg)
"""

from __future__ import annotations
import os
from datetime import date
from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView,
    QInputDialog, QMessageBox, QDialog, QDialogButtonBox, QFormLayout, QSplitter,
    QTreeWidget, QTreeWidgetItem, QGroupBox, QTextEdit, QTextBrowser,
    QFrame, QScrollArea, QStackedWidget,
)

from engines.engine_inboeken import CATEGORIES_TREE, InboekenEngine
from engines.engine_krat_beheer import (
    wc_lookup, export_nieuwe_artikelen, export_samenvoeg_update, count_beprijsd,
)
from services.krat_state import list_kratten, load_krat, save_krat, delete_krat, new_krat
from utils.paths import output_root
from utils.theme import apply_theme

STATUS_LABELS = {
    "inventarisatie": "🔵 Inventarisatie bezig",
    "wacht_op_prijs": "🟠 Wacht op prijs",
    "beprijzing":     "🟡 Beprijzing bezig",
    "klaar":          "🟢 Klaar voor export",
    "geexporteerd":   "✅ Geëxporteerd",
}


# ──────────────────────────────────────────────────────────────
# Dialogen
# ──────────────────────────────────────────────────────────────

class NieuwKratDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Nieuw krat aanmaken")
        self.setMinimumWidth(380)
        lay = QVBoxLayout(self)

        form = QFormLayout()
        self.ed_naam = QLineEdit()
        self.ed_naam.setPlaceholderText("bijv. Krat April A")
        self.ed_locatie = QLineEdit()
        self.ed_locatie.setPlaceholderText("bijv. D33, PD12, GB 1")
        form.addRow("Naam:", self.ed_naam)
        form.addRow("Locatie (WC locatieveld):", self.ed_locatie)
        lay.addLayout(form)

        info = QLabel(
            "De locatie wordt het WC-locatieveld (Beschrijving) voor alle\n"
            "nieuwe artikelen in dit krat bij export."
        )
        info.setStyleSheet("color: palette(mid); font-size: 11px;")
        lay.addWidget(info)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._on_ok)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _on_ok(self):
        if not self.ed_naam.text().strip():
            QMessageBox.warning(self, "Naam vereist", "Vul een naam in voor het krat.")
            return
        self.accept()

    def get_values(self) -> tuple[str, str]:
        return self.ed_naam.text().strip(), self.ed_locatie.text().strip()


# ──────────────────────────────────────────────────────────────
# Help-dialoog
# ──────────────────────────────────────────────────────────────

class KratHelpDialog(QDialog):
    _CONTENT = {
        0: ("Krat Beheer — Hoe werkt het?", """
<h3>Wat is Krat Beheer?</h3>
<p>Gebruik deze tab om kratten vol onderdelen <b>zonder prijs</b> te inventariseren,
beprijzen en exporteren naar WooCommerce.</p>

<h3>De 3 stappen op een rij</h3>
<ol>
  <li><b>Inventarisatie</b> — Typ elk artikelnummer, kies of het nieuw of
      bestaand is in WC, noteer de voorraad en selecteer categorieën.</li>
  <li><b>Beprijzing</b> — Samen met de baas: typ prijs per artikel,
      <b>Enter</b> = opslaan en volgende.</li>
  <li><b>Export</b> — Genereer XLSX-bestanden voor WP All Import en upload ze
      op de website.</li>
</ol>

<h3>Hoe beginnen?</h3>
<ol>
  <li>Klik <b>Nieuw krat</b>.</li>
  <li>Geef het krat een naam (bijv. <i>Krat April A</i>) en een locatieCode
      (bijv. <i>D33</i> — dit wordt het WC-locatieveld voor nieuwe artikelen).</li>
  <li>Selecteer het krat in de lijst en klik <b>Inventariseer</b>.</li>
</ol>

<h3>Statussen</h3>
<ul>
  <li>🔵 <b>Inventarisatie bezig</b> — Je bent nog artikelen aan het inventariseren.</li>
  <li>🟠 <b>Wacht op prijs</b> — Inventarisatie klaar; beprijzing nog niet gestart.</li>
  <li>🟡 <b>Beprijzing bezig</b> — Samen met de baas aan het beprijzen.</li>
  <li>🟢 <b>Klaar voor export</b> — Alle artikelen beprijsd of overgeslagen.</li>
  <li>✅ <b>Geëxporteerd</b> — Upload gedaan, krat is klaar.</li>
</ul>
"""),
        1: ("Inventarisatie — Hoe werkt het?", """
<h3>Stap voor stap</h3>
<ol>
  <li>Typ het <b>artikelnummer</b> en druk <b>Enter</b> of klik
      <i>Controleer WC</i>.</li>
  <li>De app controleert of het artikel al in WooCommerce staat:
    <ul>
      <li><b>Gevonden</b> → kies <i>Samenvoegen met bestaande</i>
          (voorraad wordt opgeteld bij export) of <i>Aanmaken als nieuw</i>.</li>
      <li><b>Niet gevonden</b> → klik <i>Toevoegen als nieuw</i>.</li>
    </ul>
  </li>
  <li>Vul de <b>omschrijving</b> in (bijv. <i>O-RING</i>).</li>
  <li>Vul de <b>voorraad</b> in — het aantal stuks dat in het krat ligt.</li>
  <li>Vink rechts de juiste <b>categorieën</b> aan. Gebruik de filterbalk om
      snel te zoeken.</li>
  <li>Druk op <b>Toevoegen</b> — het formulier wordt leeggemaakt voor het
      volgende artikel.</li>
  <li>Herhaal voor elk artikel in het krat.</li>
  <li>Als het krat leeg is: klik <b>Klaar met inventarisatie →</b></li>
</ol>

<h3>Zedder / Reiner</h3>
<p>Plak tekst uit Zedder of Reiner in het tekstvak linksonder. Gebruik dan:</p>
<ul>
  <li><b>Zedder → Categorieën</b> — detecteert automatisch het model en vinkt
      de juiste categorieën aan.</li>
  <li><b>Zedder → Titel + Omschr.</b> — vult omschrijving in vanuit de tekst.</li>
</ul>

<h3>Tips</h3>
<ul>
  <li>De locatie van het krat (bijv. <i>D33</i>) wordt het WC-locatieveld voor
      alle <i>nieuwe</i> artikelen. Bij <i>samenvoegen</i> blijft de bestaande
      WC-locatie bewaard.</li>
  <li>Alles wordt direct opgeslagen na elk artikel — je kunt de app
      tussentijds sluiten zonder gegevensverlies.</li>
  <li>Vergissing? Selecteer het artikel in de onderste tabel en klik
      <i>Verwijder geselecteerd artikel</i>.</li>
</ul>
"""),
        2: ("Beprijzing — Hoe werkt het?", """
<h3>Stap voor stap</h3>
<ol>
  <li>Elk artikel wordt één voor één getoond — het artikelnummer staat groot
      in beeld.</li>
  <li>Typ de <b>prijs</b> in euro's (gebruik komma of punt als
      decimaalteken).</li>
  <li>Druk <b>Enter</b> → prijs wordt opgeslagen en je gaat automatisch naar
      het volgende artikel.</li>
  <li>Druk <b>Escape</b> (of klik <i>Overslaan</i>) als je de prijs nu niet
      weet → het artikel wordt als uitverkocht (€0) geëxporteerd.</li>
  <li>Klik <b>← Vorige</b> om een al ingevuld artikel te corrigeren.</li>
  <li>Dubbelklik in de <b>voortgangslijst</b> rechts om direct naar een
      specifiek artikel te springen.</li>
  <li>Na het laatste artikel: klik <b>Klaar → Ga naar export</b>.</li>
</ol>

<h3>Samenvoeg-artikelen</h3>
<p>Als een artikel als <i>Samenvoegen</i> is aangemerkt, zie je de huidige
WC-prijs alvast ingevuld. Je kunt die overnemen (gewoon Enter drukken) of
een nieuwe prijs typen.</p>

<h3>Overgeslagen artikelen</h3>
<p>Overgeslagen artikelen worden bij de export meegenomen als
<b>uitverkocht</b> (prijs €0, voorraad 0). Je kunt ze later nog corrigeren
door opnieuw de beprijzing te openen en er naartoe te dubbelklikken.</p>

<h3>Voortgang bewaren</h3>
<p>Na elk artikel wordt de voortgang direct opgeslagen. Je kunt pauzeren en
later hervatten — de app pikt het op waar je gebleven was.</p>
"""),
        3: ("Export — Hoe werkt het?", """
<h3>Twee soorten exports</h3>
<ul>
  <li><b>Export A — Nieuwe artikelen</b><br>
      Artikelen die nog <i>niet</i> in WooCommerce staan. Upload dit bestand
      in WP All Import als <i>nieuwe producten aanmaken</i>.</li>
  <li><b>Export B — Samenvoeg update</b><br>
      Artikelen die al in WC staan. De voorraad wordt opgeteld bij de
      bestaande WC-voorraad. Upload dit als <i>update van bestaande
      producten</i>.</li>
</ul>

<h3>Stap voor stap</h3>
<ol>
  <li>Controleer de samenvatting bovenaan (hoeveel nieuw, hoeveel update,
      hoeveel overgeslagen).</li>
  <li>Klik <b>Export A</b> voor nieuwe artikelen, <b>Export B</b> voor
      samenvoeg-updates, of <b>Exporteer beide tegelijk</b>.</li>
  <li>De map opent automatisch in Windows Verkenner.</li>
  <li>Upload de XLSX-bestanden in WP All Import op de website.</li>
  <li>Klik daarna <b>Markeer als geëxporteerd</b> zodat het krat als
      ✅ Geëxporteerd gemarkeerd wordt.</li>
</ol>

<h3>Let op</h3>
<ul>
  <li><b>Export B</b> is grijs als er geen samenvoeg-artikelen zijn.</li>
  <li>Overgeslagen artikelen worden als uitverkocht (€0, voorraad 0)
      geëxporteerd — controleer de samenvatting.</li>
</ul>
"""),
    }

    def __init__(self, page: int = 0, parent=None):
        super().__init__(parent)
        title, html = self._CONTENT.get(page, self._CONTENT[0])
        self.setWindowTitle(title)
        self.setMinimumWidth(520)
        self.setMinimumHeight(460)
        self.resize(560, 520)

        lay = QVBoxLayout(self)
        lay.setSpacing(12)

        tb = QTextBrowser()
        tb.setOpenExternalLinks(False)
        tb.setHtml(html)
        lay.addWidget(tb)

        btns = QDialogButtonBox(QDialogButtonBox.Close)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)


# ──────────────────────────────────────────────────────────────
# Hoofd-tab
# ──────────────────────────────────────────────────────────────

class TabKratBeheer(QWidget):
    def __init__(self, app_state, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self._active_krat: Optional[dict] = None
        self._active_wc_match: Optional[dict] = None
        self._inboeken_engine = InboekenEngine()

        self.selected_category_paths: set[str] = set()
        self._item_by_path: dict[str, QTreeWidgetItem] = {}

        self._build_ui()

    # ──────────────────────────────────────────────────────────
    # UI opbouw
    # ──────────────────────────────────────────────────────────

    def _build_ui(self):
        apply_theme(self)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.stack = QStackedWidget()
        root.addWidget(self.stack)

        self.stack.addWidget(self._build_overzicht_page())       # 0
        self.stack.addWidget(self._build_inventarisatie_page())  # 1
        self.stack.addWidget(self._build_beprijzing_page())      # 2
        self.stack.addWidget(self._build_export_page())          # 3

        self.stack.setCurrentIndex(0)
        self._refresh_overzicht()

    # ──────────────────────────────────────────────────────────
    # PAGE 0 — Overzicht
    # ──────────────────────────────────────────────────────────

    def _build_overzicht_page(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(12)

        hdr_row = QHBoxLayout()
        title = QLabel("Krat Beheer")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        hdr_row.addWidget(title)
        hdr_row.addStretch()
        btn_help_ov = QPushButton("?")
        btn_help_ov.setObjectName("secondary")
        btn_help_ov.setFixedWidth(44)
        btn_help_ov.setToolTip("Uitleg & stappenplan")
        btn_help_ov.clicked.connect(lambda: KratHelpDialog(0, self).exec())
        hdr_row.addWidget(btn_help_ov)
        lay.addLayout(hdr_row)

        sub = QLabel(
            "Inventariseer kratten zonder prijs, bepaal daarna samen met de baas de prijs per artikel, "
            "en exporteer als WP All Import XLSX."
        )
        sub.setWordWrap(True)
        lay.addWidget(sub)

        self.ov_table = QTableWidget()
        self.ov_table.setColumnCount(6)
        self.ov_table.setHorizontalHeaderLabels(
            ["Naam", "Locatie", "Artikelen", "Beprijsd", "Status", "Aangemaakt"]
        )
        self.ov_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.ov_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.ov_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.ov_table.verticalHeader().setVisible(False)
        hdr = self.ov_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)
        for col in range(1, 6):
            hdr.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        lay.addWidget(self.ov_table, stretch=1)

        row = QHBoxLayout()
        btn_nieuw = QPushButton("Nieuw krat")
        btn_nieuw.setObjectName("primary")
        btn_nieuw.clicked.connect(self.on_nieuw_krat)

        btn_inv = QPushButton("Inventariseer")
        btn_inv.setObjectName("secondary")
        btn_inv.clicked.connect(self.on_open_inventarisatie)

        btn_prijs = QPushButton("Start beprijzing")
        btn_prijs.setObjectName("primary")
        btn_prijs.clicked.connect(self.on_open_beprijzing)

        btn_export = QPushButton("Exporteer")
        btn_export.setObjectName("secondary")
        btn_export.clicked.connect(self.on_open_export)

        btn_del = QPushButton("Verwijder krat")
        btn_del.setObjectName("danger")
        btn_del.clicked.connect(self.on_verwijder_krat)

        row.addWidget(btn_nieuw)
        row.addWidget(btn_inv)
        row.addWidget(btn_prijs)
        row.addWidget(btn_export)
        row.addStretch()
        row.addWidget(btn_del)
        lay.addLayout(row)

        return page

    def _refresh_overzicht(self):
        kratten = list_kratten()
        self._overzicht_kratten = kratten
        self.ov_table.setRowCount(len(kratten))
        for i, k in enumerate(kratten):
            done, total = count_beprijsd(k)
            self.ov_table.setItem(i, 0, QTableWidgetItem(k.get("naam", "")))
            self.ov_table.setItem(i, 1, QTableWidgetItem(k.get("locatie", "")))
            self.ov_table.setItem(i, 2, QTableWidgetItem(str(total)))
            self.ov_table.setItem(i, 3, QTableWidgetItem(f"{done}/{total}"))
            status = k.get("status", "")
            self.ov_table.setItem(i, 4, QTableWidgetItem(STATUS_LABELS.get(status, status)))
            self.ov_table.setItem(i, 5, QTableWidgetItem(k.get("aangemaakt", "")[:10]))

    def _selected_krat(self) -> Optional[dict]:
        row = self.ov_table.currentRow()
        if row < 0 or not hasattr(self, "_overzicht_kratten"):
            return None
        if row >= len(self._overzicht_kratten):
            return None
        return load_krat(self._overzicht_kratten[row]["krat_id"])

    def on_nieuw_krat(self):
        dlg = NieuwKratDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return
        naam, locatie = dlg.get_values()
        krat = new_krat(naam, locatie)
        save_krat(krat)
        self._refresh_overzicht()

    def on_open_inventarisatie(self):
        krat = self._selected_krat()
        if krat is None:
            QMessageBox.information(self, "Geen selectie", "Selecteer eerst een krat.")
            return
        self._show_inventarisatie(krat)

    def on_open_beprijzing(self):
        krat = self._selected_krat()
        if krat is None:
            QMessageBox.information(self, "Geen selectie", "Selecteer eerst een krat.")
            return
        if not krat.get("artikelen"):
            QMessageBox.information(
                self, "Leeg krat",
                "Dit krat heeft nog geen artikelen.\nVoer eerst de inventarisatie uit."
            )
            return
        self._show_beprijzing(krat)

    def on_open_export(self):
        krat = self._selected_krat()
        if krat is None:
            QMessageBox.information(self, "Geen selectie", "Selecteer eerst een krat.")
            return
        self._show_export(krat)

    def on_verwijder_krat(self):
        krat = self._selected_krat()
        if krat is None:
            QMessageBox.information(self, "Geen selectie", "Selecteer eerst een krat.")
            return
        r = QMessageBox.question(
            self, "Verwijder krat",
            f"Weet je zeker dat je '{krat.get('naam', '')}' wilt verwijderen?\n"
            "Alle inventarisatie- en beprijzingsdata gaat verloren.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if r == QMessageBox.Yes:
            delete_krat(krat["krat_id"])
            self._refresh_overzicht()

    # ──────────────────────────────────────────────────────────
    # PAGE 1 — Inventarisatie
    # ──────────────────────────────────────────────────────────

    def _build_inventarisatie_page(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(12, 12, 12, 12)
        lay.setSpacing(8)

        inv_hdr_row = QHBoxLayout()
        self.inv_header = QLabel("Inventarisatie")
        self.inv_header.setStyleSheet("font-size: 18px; font-weight: bold;")
        inv_hdr_row.addWidget(self.inv_header)
        inv_hdr_row.addStretch()
        btn_help_inv = QPushButton("?")
        btn_help_inv.setObjectName("secondary")
        btn_help_inv.setFixedWidth(44)
        btn_help_inv.setToolTip("Uitleg inventarisatie")
        btn_help_inv.clicked.connect(lambda: KratHelpDialog(1, self).exec())
        inv_hdr_row.addWidget(btn_help_inv)
        lay.addLayout(inv_hdr_row)

        splitter = QSplitter(Qt.Horizontal)
        lay.addWidget(splitter, stretch=1)

        # ─── Links ───────────────────────────────────────────
        left_w = QWidget()
        ll = QVBoxLayout(left_w)
        ll.setContentsMargins(0, 0, 4, 0)
        ll.setSpacing(8)

        gb_art = QGroupBox("Artikel toevoegen")
        gl = QVBoxLayout(gb_art)

        r1 = QHBoxLayout()
        r1.addWidget(QLabel("Artikelnummer:"))
        self.inv_nummer = QLineEdit()
        self.inv_nummer.setPlaceholderText("bijv. 09320-05006")
        self.inv_nummer.returnPressed.connect(self._inv_check_wc)
        r1.addWidget(self.inv_nummer)
        btn_check = QPushButton("Controleer WC")
        btn_check.setObjectName("secondary")
        btn_check.clicked.connect(self._inv_check_wc)
        r1.addWidget(btn_check)
        gl.addLayout(r1)

        r2 = QHBoxLayout()
        r2.addWidget(QLabel("Omschrijving:"))
        self.inv_omschr = QLineEdit()
        self.inv_omschr.setPlaceholderText("bijv. O-RING")
        r2.addWidget(self.inv_omschr)
        gl.addLayout(r2)

        r3 = QHBoxLayout()
        r3.addWidget(QLabel("Voorraad:"))
        self.inv_voorraad = QLineEdit()
        self.inv_voorraad.setPlaceholderText("bijv. 3")
        self.inv_voorraad.setFixedWidth(90)
        r3.addWidget(self.inv_voorraad)
        r3.addStretch()
        gl.addLayout(r3)

        # WC-match panel (verborgen tot check)
        self.inv_wc_panel = QFrame()
        self.inv_wc_panel.setFrameShape(QFrame.StyledPanel)
        self.inv_wc_panel.setStyleSheet("""
            QFrame {
                background: rgba(255, 200, 50, 0.12);
                border: 1px solid rgba(180, 130, 0, 0.4);
                border-radius: 6px;
            }
        """)
        wcp = QVBoxLayout(self.inv_wc_panel)
        wcp.setContentsMargins(10, 8, 10, 8)
        self.inv_wc_label = QLabel("")
        self.inv_wc_label.setWordWrap(True)
        wcp.addWidget(self.inv_wc_label)
        wc_btns = QHBoxLayout()
        self.btn_als_nieuw = QPushButton("Aanmaken als nieuw")
        self.btn_als_nieuw.setObjectName("primary")
        self.btn_als_nieuw.clicked.connect(lambda: self._inv_toevoegen("nieuw"))
        self.btn_samenvoeg = QPushButton("Samenvoegen met bestaande")
        self.btn_samenvoeg.setObjectName("secondary")
        self.btn_samenvoeg.clicked.connect(lambda: self._inv_toevoegen("update"))
        wc_btns.addWidget(self.btn_als_nieuw)
        wc_btns.addWidget(self.btn_samenvoeg)
        wcp.addLayout(wc_btns)
        self.inv_wc_panel.setVisible(False)
        gl.addWidget(self.inv_wc_panel)

        self.btn_toevoegen = QPushButton("Toevoegen")
        self.btn_toevoegen.setObjectName("primary")
        self.btn_toevoegen.clicked.connect(lambda: self._inv_toevoegen("nieuw"))
        gl.addWidget(self.btn_toevoegen)

        ll.addWidget(gb_art)

        # Zedder
        gb_zed = QGroupBox("Zedder / Reiner")
        zl = QVBoxLayout(gb_zed)
        self.inv_zedder = QTextEdit()
        self.inv_zedder.setMinimumHeight(80)
        self.inv_zedder.setMaximumHeight(130)
        self.inv_zedder.setPlaceholderText("Plak Zedder tekst hier...")
        zl.addWidget(self.inv_zedder)
        zr = QHBoxLayout()
        btn_z_cat = QPushButton("Zedder → Categorieën")
        btn_z_cat.setObjectName("secondary")
        btn_z_cat.clicked.connect(self._inv_zedder_categories)
        btn_z_fill = QPushButton("Zedder → Titel + Omschr.")
        btn_z_fill.setObjectName("secondary")
        btn_z_fill.clicked.connect(self._inv_zedder_fill)
        zr.addWidget(btn_z_cat)
        zr.addWidget(btn_z_fill)
        zr.addStretch()
        zl.addLayout(zr)
        ll.addWidget(gb_zed)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left_scroll.setWidget(left_w)
        splitter.addWidget(left_scroll)

        # ─── Rechts ──────────────────────────────────────────
        right_w = QWidget()
        rl = QVBoxLayout(right_w)
        rl.setContentsMargins(4, 0, 0, 0)
        rl.setSpacing(8)

        gb_cat = QGroupBox("Categorieën")
        cl = QVBoxLayout(gb_cat)
        cr = QHBoxLayout()
        cr.addWidget(QLabel("Filter:"))
        self.inv_cat_filter = QLineEdit()
        self.inv_cat_filter.setPlaceholderText("bijv. GT750, GSX, 2-takt...")
        self.inv_cat_filter.textChanged.connect(self._inv_filter_tree)
        cr.addWidget(self.inv_cat_filter)
        btn_clr = QPushButton("X")
        btn_clr.setObjectName("secondary")
        btn_clr.setFixedWidth(28)
        btn_clr.clicked.connect(lambda: self.inv_cat_filter.setText(""))
        cr.addWidget(btn_clr)
        cl.addLayout(cr)

        self.inv_tree = QTreeWidget()
        self.inv_tree.setHeaderHidden(True)
        self.inv_tree.setUniformRowHeights(True)
        self.inv_tree.setAnimated(False)
        self.inv_tree.itemChanged.connect(self._inv_on_tree_changed)
        cl.addWidget(self.inv_tree)

        self.inv_lbl_cats = QLabel("Geselecteerde categorieën: (0)")
        self.inv_lbl_cats.setWordWrap(True)
        cl.addWidget(self.inv_lbl_cats)
        rl.addWidget(gb_cat, stretch=1)

        lbl_art = QLabel("Toegevoegde artikelen:")
        lbl_art.setStyleSheet("font-weight: bold;")
        rl.addWidget(lbl_art)

        self.inv_table = QTableWidget()
        self.inv_table.setColumnCount(5)
        self.inv_table.setHorizontalHeaderLabels(["#", "Artikelnummer", "Omschrijving", "Aantal", "Type"])
        self.inv_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.inv_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.inv_table.verticalHeader().setVisible(False)
        th = self.inv_table.horizontalHeader()
        th.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        th.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        th.setSectionResizeMode(2, QHeaderView.Stretch)
        th.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        th.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.inv_table.setMinimumHeight(130)
        self.inv_table.doubleClicked.connect(self._inv_edit_aantal)
        rl.addWidget(self.inv_table)

        btn_del_art = QPushButton("Verwijder geselecteerd artikel")
        btn_del_art.setObjectName("danger")
        btn_del_art.clicked.connect(self._inv_verwijder_artikel)
        rl.addWidget(btn_del_art)

        splitter.addWidget(right_w)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 6)

        nav = QHBoxLayout()
        btn_terug = QPushButton("← Terug naar overzicht")
        btn_terug.setObjectName("secondary")
        btn_terug.clicked.connect(self._nav_overzicht)
        self.btn_klaar_inv = QPushButton("Klaar met inventarisatie →")
        self.btn_klaar_inv.setObjectName("primary")
        self.btn_klaar_inv.clicked.connect(self._inv_klaar)
        nav.addWidget(btn_terug)
        nav.addStretch()
        nav.addWidget(self.btn_klaar_inv)
        lay.addLayout(nav)

        self._build_inv_tree()
        return page

    def _build_inv_tree(self):
        self.inv_tree.blockSignals(True)
        try:
            self.inv_tree.clear()
            self._item_by_path = {}

            def add_node(parent, node, prefix):
                name = str(node.get("name", "")).strip()
                if not name:
                    return
                path_list = prefix + [name]
                path_str = " > ".join(path_list)
                it = QTreeWidgetItem([name])
                it.setData(0, Qt.UserRole, path_str)
                flags = (
                    it.flags()
                    | Qt.ItemIsEnabled
                    | Qt.ItemIsSelectable
                    | Qt.ItemIsUserCheckable
                )
                if hasattr(Qt, "ItemIsAutoTristate"):
                    flags = flags & ~Qt.ItemIsAutoTristate
                it.setFlags(flags)
                it.setCheckState(0, Qt.Unchecked)
                if parent is None:
                    self.inv_tree.addTopLevelItem(it)
                else:
                    parent.addChild(it)
                self._item_by_path[path_str] = it
                for child in node.get("children", []) or []:
                    add_node(it, child, path_list)
                if len(path_list) <= 3:
                    it.setExpanded(True)

            for n in CATEGORIES_TREE:
                add_node(None, n, [])
        finally:
            self.inv_tree.blockSignals(False)

    def _inv_filter_tree(self, text: str):
        q = (text or "").strip().lower()

        def walk(it: QTreeWidgetItem) -> bool:
            any_child = False
            for i in range(it.childCount()):
                if walk(it.child(i)):
                    any_child = True
            name = (it.text(0) or "").lower()
            path = (it.data(0, Qt.UserRole) or "").lower()
            visible = True if not q else ((q in name or q in path) or any_child)
            it.setHidden(not visible)
            if visible and q and (q in name or q in path):
                it.setExpanded(True)
            return visible

        for i in range(self.inv_tree.topLevelItemCount()):
            walk(self.inv_tree.topLevelItem(i))

    def _inv_on_tree_changed(self, item: QTreeWidgetItem, _col: int):
        path = item.data(0, Qt.UserRole)
        if not path:
            return
        if item.checkState(0) == Qt.Checked:
            self.selected_category_paths.add(path)
        else:
            self.selected_category_paths.discard(path)
        self._inv_update_cats_label()

    def _inv_sync_tree(self):
        self.inv_tree.blockSignals(True)
        try:
            for path, it in self._item_by_path.items():
                if not (it.flags() & Qt.ItemIsUserCheckable):
                    continue
                it.setCheckState(
                    0, Qt.Checked if path in self.selected_category_paths else Qt.Unchecked
                )
        finally:
            self.inv_tree.blockSignals(False)

    def _inv_update_cats_label(self):
        self.inv_lbl_cats.setText(f"Geselecteerde categorieën: ({len(self.selected_category_paths)})")

    def _inv_check_wc(self):
        nummer = self.inv_nummer.text().strip()
        if not nummer:
            return

        wc_df = getattr(self.app_state, "wc_df", None)
        match = wc_lookup(nummer, wc_df)

        if match:
            prijs = match.get("prijs", 0.0)
            voorraad = match.get("voorraad", 0)
            locatie = match.get("locatie", "") or "—"
            short = match.get("short_description", "")

            lines = [
                f"✅ Gevonden in WooCommerce",
                f"Prijs: € {prijs:.2f}   ·   Voorraad: {voorraad}   ·   Locatie: {locatie}",
            ]
            if short:
                lines.append(f"Omschrijving: {short}")
            self.inv_wc_label.setText("\n".join(lines))
            self.inv_wc_panel.setVisible(True)
            self.btn_toevoegen.setVisible(False)
            self._active_wc_match = match

            # Auto-aanvullen
            if short and not self.inv_omschr.text().strip():
                self.inv_omschr.setText(short)

            # Auto-aanvinken categorieën uit WC
            wc_cats = [c for c in match.get("categorieen", []) if c in self._item_by_path]
            if wc_cats:
                self.selected_category_paths.update(wc_cats)
                self._inv_sync_tree()
                self._inv_update_cats_label()
        else:
            info = "Niet gevonden in WooCommerce — wordt aangemaakt als nieuw product."
            if wc_df is None or (hasattr(wc_df, "empty") and wc_df.empty):
                info = "Geen WC export geladen — kan niet controleren. Wordt als nieuw behandeld."
            self.inv_wc_label.setText(f"ℹ️ {info}")
            self.inv_wc_panel.setVisible(True)
            self.btn_toevoegen.setVisible(False)
            self._active_wc_match = None

            # Toon alsnog alleen "als nieuw"-knop in panel
            self.btn_samenvoeg.setVisible(False)
            self.btn_als_nieuw.setText("Toevoegen als nieuw")

    def _inv_toevoegen(self, beslissing: str):
        nummer = self.inv_nummer.text().strip()
        if not nummer:
            QMessageBox.warning(self, "Artikelnummer vereist", "Vul een artikelnummer in.")
            return

        # Duplicaat-check: staat dit nummer al in het krat?
        bestaand = next(
            (a for a in (self._active_krat.get("artikelen", []) if self._active_krat else [])
             if a.get("artikelnummer", "").upper() == nummer.upper()),
            None,
        )
        if bestaand is not None:
            r = QMessageBox.question(
                self, "Al in krat",
                f"'{nummer}' staat al in dit krat (aantal: {bestaand.get('voorraad', 0)}).\n\n"
                "Wil je het aantal ophogen?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if r == QMessageBox.Yes:
                extra, ok = QInputDialog.getInt(
                    self, "Aantal ophogen",
                    f"Hoeveel stuks extra van {nummer}?",
                    value=1, min=1, max=9999,
                )
                if ok:
                    bestaand["voorraad"] = bestaand.get("voorraad", 0) + extra
                    save_krat(self._active_krat)
                    self._inv_refresh_table()
                    self._inv_reset_form()
            return

        omschr = self.inv_omschr.text().strip()
        try:
            voorraad = int(float(self.inv_voorraad.text().replace(",", ".") or "0"))
        except Exception:
            voorraad = 0

        wc_match = self._active_wc_match if beslissing == "update" else None

        art = {
            "positie":              len(self._active_krat.get("artikelen", [])),
            "artikelnummer":        nummer,
            "omschrijving":         omschr,
            "categorieen":          sorted(self.selected_category_paths),
            "voorraad":             voorraad,
            "prijs":                None,
            "prijs_status":         None,
            "wc_match":             wc_match,
            "samenvoeg_beslissing": beslissing,
        }

        self._active_krat.setdefault("artikelen", []).append(art)
        save_krat(self._active_krat)
        self._inv_refresh_table()
        self._inv_reset_form()

    def _inv_reset_form(self):
        self.inv_nummer.clear()
        self.inv_omschr.clear()
        self.inv_voorraad.clear()
        self.inv_zedder.clear()
        self.inv_cat_filter.clear()
        self.selected_category_paths.clear()
        self._inv_sync_tree()
        self._inv_update_cats_label()
        self.inv_wc_panel.setVisible(False)
        self.btn_toevoegen.setVisible(True)
        self.btn_samenvoeg.setVisible(True)
        self.btn_als_nieuw.setText("Aanmaken als nieuw")
        self._active_wc_match = None
        self.inv_nummer.setFocus()

    def _inv_refresh_table(self):
        artikelen = self._active_krat.get("artikelen", []) if self._active_krat else []
        self.inv_table.setRowCount(len(artikelen))
        for i, art in enumerate(artikelen):
            type_str = (
                "🔀 Samenvoegen" if art.get("samenvoeg_beslissing") == "update" else "✅ Nieuw"
            )
            self.inv_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.inv_table.setItem(i, 1, QTableWidgetItem(art.get("artikelnummer", "")))
            self.inv_table.setItem(i, 2, QTableWidgetItem(art.get("omschrijving", "")))
            self.inv_table.setItem(i, 3, QTableWidgetItem(str(art.get("voorraad", 0))))
            self.inv_table.setItem(i, 4, QTableWidgetItem(type_str))

    def _inv_verwijder_artikel(self):
        row = self.inv_table.currentRow()
        if row < 0:
            return
        artikelen = self._active_krat.get("artikelen", [])
        if row >= len(artikelen):
            return
        art = artikelen[row]
        r = QMessageBox.question(
            self, "Verwijder artikel",
            f"Verwijder '{art.get('artikelnummer', '')}' uit het krat?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if r == QMessageBox.Yes:
            del artikelen[row]
            for j, a in enumerate(artikelen):
                a["positie"] = j
            save_krat(self._active_krat)
            self._inv_refresh_table()

    def _inv_edit_aantal(self, index):
        """Dubbel-klik op Aantal-cel → pas voorraad aan."""
        if index.column() != 3:
            return
        row = index.row()
        artikelen = self._active_krat.get("artikelen", []) if self._active_krat else []
        if row >= len(artikelen):
            return
        art = artikelen[row]
        huidig = art.get("voorraad", 0)
        nieuw, ok = QInputDialog.getInt(
            self, "Voorraad aanpassen",
            f"Nieuw aantal voor {art.get('artikelnummer', '')}:",
            value=huidig, min=0, max=9999,
        )
        if ok and nieuw != huidig:
            art["voorraad"] = nieuw
            save_krat(self._active_krat)
            self._inv_refresh_table()

    def _inv_zedder_categories(self):
        t = self.inv_zedder.toPlainText()
        paths = self._inboeken_engine.zedder_detect_model_category_paths(t)
        self.selected_category_paths.update(paths)
        self._inv_sync_tree()
        self._inv_update_cats_label()

    def _inv_zedder_fill(self):
        t = self.inv_zedder.toPlainText()
        title, desc = self._inboeken_engine.zedder_fill_title_and_desc(
            t, current_title=self.inv_nummer.text()
        )
        if title:
            self.inv_nummer.setText(title)
        if desc:
            self.inv_omschr.setText(desc)

    def _inv_klaar(self):
        if not self._active_krat:
            return
        if not self._active_krat.get("artikelen"):
            QMessageBox.information(
                self, "Leeg krat",
                "Voeg eerst minimaal één artikel toe voor je de inventarisatie afsluit."
            )
            return
        self._active_krat["status"] = "wacht_op_prijs"
        save_krat(self._active_krat)
        QMessageBox.information(
            self, "Inventarisatie klaar",
            f"Krat '{self._active_krat.get('naam', '')}' staat nu klaar voor beprijzing."
        )
        self._nav_overzicht()

    # ──────────────────────────────────────────────────────────
    # PAGE 2 — Beprijzing
    # ──────────────────────────────────────────────────────────

    def _build_beprijzing_page(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(12)

        bp_hdr_row = QHBoxLayout()
        self.bp_header = QLabel("Beprijzing")
        self.bp_header.setStyleSheet("font-size: 18px; font-weight: bold;")
        bp_hdr_row.addWidget(self.bp_header)
        bp_hdr_row.addStretch()
        btn_help_bp = QPushButton("?")
        btn_help_bp.setObjectName("secondary")
        btn_help_bp.setFixedWidth(44)
        btn_help_bp.setToolTip("Uitleg beprijzing")
        btn_help_bp.clicked.connect(lambda: KratHelpDialog(2, self).exec())
        bp_hdr_row.addWidget(btn_help_bp)
        lay.addLayout(bp_hdr_row)

        self.bp_progress_lbl = QLabel("")
        self.bp_progress_lbl.setStyleSheet("color: palette(mid);")
        lay.addWidget(self.bp_progress_lbl)

        center = QSplitter(Qt.Horizontal)
        lay.addWidget(center, stretch=1)

        # ─── Links: artikel + prijs ───────────────────────────
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 12, 0)
        ll.setSpacing(12)

        self.bp_nummer_lbl = QLabel("")
        f = QFont()
        f.setPointSize(30)
        f.setBold(True)
        self.bp_nummer_lbl.setFont(f)
        self.bp_nummer_lbl.setAlignment(Qt.AlignCenter)
        ll.addWidget(self.bp_nummer_lbl)

        self.bp_info_lbl = QLabel("")
        self.bp_info_lbl.setWordWrap(True)
        self.bp_info_lbl.setAlignment(Qt.AlignCenter)
        self.bp_info_lbl.setStyleSheet("color: palette(mid); font-size: 13px;")
        ll.addWidget(self.bp_info_lbl)

        self.bp_wc_lbl = QLabel("")
        self.bp_wc_lbl.setWordWrap(True)
        self.bp_wc_lbl.setAlignment(Qt.AlignCenter)
        self.bp_wc_lbl.setStyleSheet("""
            QLabel {
                background: rgba(255, 200, 50, 0.12);
                border: 1px solid rgba(180, 130, 0, 0.4);
                border-radius: 6px;
                padding: 6px 10px;
            }
        """)
        self.bp_wc_lbl.setVisible(False)
        ll.addWidget(self.bp_wc_lbl)

        ll.addStretch()

        prijs_row = QHBoxLayout()
        prijs_row.addStretch()
        lbl_eur = QLabel("Prijs (€):")
        lbl_eur.setStyleSheet("font-size: 15px;")
        prijs_row.addWidget(lbl_eur)
        self.bp_prijs = QLineEdit()
        self.bp_prijs.setPlaceholderText("bijv. 12,50")
        self.bp_prijs.setFixedWidth(180)
        self.bp_prijs.setStyleSheet("font-size: 18px; padding: 6px;")
        self.bp_prijs.returnPressed.connect(self._bp_volgende)
        prijs_row.addWidget(self.bp_prijs)
        prijs_row.addStretch()
        ll.addLayout(prijs_row)

        hint = QLabel("Enter = opslaan en volgende  ·  Escape = overslaan")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("color: palette(mid); font-size: 11px;")
        ll.addWidget(hint)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_vorige = QPushButton("← Vorige")
        btn_vorige.setObjectName("secondary")
        btn_vorige.clicked.connect(self._bp_vorige)
        btn_skip = QPushButton("Overslaan")
        btn_skip.setObjectName("secondary")
        btn_skip.clicked.connect(self._bp_overslaan)
        self.btn_bp_volgende = QPushButton("Volgende →")
        self.btn_bp_volgende.setObjectName("primary")
        self.btn_bp_volgende.clicked.connect(self._bp_volgende)
        btn_row.addWidget(btn_vorige)
        btn_row.addWidget(btn_skip)
        btn_row.addWidget(self.btn_bp_volgende)
        btn_row.addStretch()
        ll.addLayout(btn_row)

        center.addWidget(left)

        # ─── Rechts: voortgangslijst ──────────────────────────
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(4, 0, 0, 0)

        lbl_vg = QLabel("Voortgang (dubbelklik = spring naar artikel):")
        lbl_vg.setStyleSheet("font-weight: bold;")
        rl.addWidget(lbl_vg)

        self.bp_voortgang = QTableWidget()
        self.bp_voortgang.setColumnCount(3)
        self.bp_voortgang.setHorizontalHeaderLabels(["", "Artikelnummer", "Prijs"])
        self.bp_voortgang.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.bp_voortgang.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.bp_voortgang.verticalHeader().setVisible(False)
        vh = self.bp_voortgang.horizontalHeader()
        vh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        vh.setSectionResizeMode(1, QHeaderView.Stretch)
        vh.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.bp_voortgang.doubleClicked.connect(
            lambda idx: self._bp_show_artikel(idx.row())
        )
        rl.addWidget(self.bp_voortgang)

        center.addWidget(right)
        center.setStretchFactor(0, 6)
        center.setStretchFactor(1, 4)

        nav = QHBoxLayout()
        btn_terug = QPushButton("← Terug naar overzicht")
        btn_terug.setObjectName("secondary")
        btn_terug.clicked.connect(self._nav_overzicht)
        self.btn_bp_export = QPushButton("Klaar → Ga naar export")
        self.btn_bp_export.setObjectName("primary")
        self.btn_bp_export.clicked.connect(self._bp_klaar)
        nav.addWidget(btn_terug)
        nav.addStretch()
        nav.addWidget(self.btn_bp_export)
        lay.addLayout(nav)

        return page

    def _bp_show_artikel(self, idx: int):
        artikelen = self._active_krat.get("artikelen", [])
        if not artikelen:
            return

        idx = max(0, min(idx, len(artikelen) - 1))
        self._active_krat["beprijzing_index"] = idx
        save_krat(self._active_krat)

        art = artikelen[idx]
        done, total = count_beprijsd(self._active_krat)

        self.bp_header.setText(f"Beprijzing — {self._active_krat.get('naam', '')}")
        self.bp_progress_lbl.setText(
            f"Artikel {idx + 1} van {total}   ·   {done} beprijsd / overgeslagen"
        )

        self.bp_nummer_lbl.setText(art.get("artikelnummer", ""))

        cats = art.get("categorieen", [])
        cat_str = cats[-1].split(" > ")[-1] if cats else ""
        info_parts = [p for p in [art.get("omschrijving", ""), cat_str, f"Voorraad: {art.get('voorraad', 0)}"] if p]
        self.bp_info_lbl.setText("   ·   ".join(info_parts))

        wc = art.get("wc_match")
        if wc and art.get("samenvoeg_beslissing") == "update":
            self.bp_wc_lbl.setText(
                f"🔀 Samenvoegen   ·   WC prijs: € {wc.get('prijs', 0):.2f}"
                f"   ·   WC locatie: {wc.get('locatie', '') or '—'}"
            )
            self.bp_wc_lbl.setVisible(True)
        else:
            self.bp_wc_lbl.setVisible(False)

        # Pre-fill prijs
        prijs = art.get("prijs")
        if prijs is not None:
            self.bp_prijs.setText(f"{prijs:.2f}".replace(".", ","))
        elif wc and art.get("samenvoeg_beslissing") == "update" and wc.get("prijs"):
            self.bp_prijs.setText(f"{wc['prijs']:.2f}".replace(".", ","))
        else:
            self.bp_prijs.clear()

        self.bp_prijs.setFocus()
        self.bp_prijs.selectAll()

        if idx >= len(artikelen) - 1:
            self.btn_bp_volgende.setText("Klaar →")
        else:
            self.btn_bp_volgende.setText("Volgende →")

        self._bp_refresh_voortgang(idx)

    def _bp_refresh_voortgang(self, current_idx: int):
        artikelen = self._active_krat.get("artikelen", [])
        self.bp_voortgang.setRowCount(len(artikelen))
        for i, art in enumerate(artikelen):
            status = art.get("prijs_status")
            if i == current_idx:
                icon = "▶"
            elif status == "beprijsd":
                icon = "✅"
            elif status == "overgeslagen":
                icon = "⏭"
            else:
                icon = "○"
            prijs = art.get("prijs")
            if prijs is not None:
                prijs_str = f"€ {prijs:.2f}"
            elif status == "overgeslagen":
                prijs_str = "overgeslagen"
            else:
                prijs_str = ""
            self.bp_voortgang.setItem(i, 0, QTableWidgetItem(icon))
            self.bp_voortgang.setItem(i, 1, QTableWidgetItem(art.get("artikelnummer", "")))
            self.bp_voortgang.setItem(i, 2, QTableWidgetItem(prijs_str))

        item = self.bp_voortgang.item(current_idx, 0)
        if item:
            self.bp_voortgang.scrollToItem(item)

    def _bp_save_current(self, status: str):
        artikelen = self._active_krat.get("artikelen", [])
        idx = self._active_krat.get("beprijzing_index", 0)
        if idx >= len(artikelen):
            return
        art = artikelen[idx]
        art["prijs_status"] = status
        if status == "beprijsd":
            txt = self.bp_prijs.text().strip().replace(",", ".").replace("€", "").strip()
            try:
                art["prijs"] = float(txt)
            except Exception:
                art["prijs"] = None
                art["prijs_status"] = "overgeslagen"
        save_krat(self._active_krat)

    def _bp_volgende(self):
        if not self.bp_prijs.text().strip():
            self._bp_overslaan()
            return
        self._bp_save_current("beprijsd")
        artikelen = self._active_krat.get("artikelen", [])
        idx = self._active_krat.get("beprijzing_index", 0)
        if idx >= len(artikelen) - 1:
            self._bp_klaar()
        else:
            self._bp_show_artikel(idx + 1)

    def _bp_vorige(self):
        idx = self._active_krat.get("beprijzing_index", 0)
        if idx > 0:
            self._bp_show_artikel(idx - 1)

    def _bp_overslaan(self):
        self._bp_save_current("overgeslagen")
        artikelen = self._active_krat.get("artikelen", [])
        idx = self._active_krat.get("beprijzing_index", 0)
        if idx >= len(artikelen) - 1:
            self._bp_klaar()
        else:
            self._bp_show_artikel(idx + 1)

    def _bp_klaar(self):
        if not self._active_krat:
            return
        self._active_krat["status"] = "klaar"
        save_krat(self._active_krat)
        self._show_export(self._active_krat)

    def keyPressEvent(self, event):
        if self.stack.currentIndex() == 2 and event.key() == Qt.Key_Escape:
            self._bp_overslaan()
        else:
            super().keyPressEvent(event)

    # ──────────────────────────────────────────────────────────
    # PAGE 3 — Export
    # ──────────────────────────────────────────────────────────

    def _build_export_page(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(12)

        exp_hdr_row = QHBoxLayout()
        self.exp_header = QLabel("Export")
        self.exp_header.setStyleSheet("font-size: 18px; font-weight: bold;")
        exp_hdr_row.addWidget(self.exp_header)
        exp_hdr_row.addStretch()
        btn_help_exp = QPushButton("?")
        btn_help_exp.setObjectName("secondary")
        btn_help_exp.setFixedWidth(44)
        btn_help_exp.setToolTip("Uitleg export")
        btn_help_exp.clicked.connect(lambda: KratHelpDialog(3, self).exec())
        exp_hdr_row.addWidget(btn_help_exp)
        lay.addLayout(exp_hdr_row)

        self.exp_summary = QLabel("")
        self.exp_summary.setWordWrap(True)
        self.exp_summary.setStyleSheet("""
            QLabel {
                background: palette(base);
                border: 1px solid palette(mid);
                border-radius: 8px;
                padding: 12px;
                font-family: monospace;
            }
        """)
        lay.addWidget(self.exp_summary)

        btn_a = QPushButton("Export A — Nieuwe artikelen (WP All Import)")
        btn_a.setObjectName("primary")
        btn_a.setMinimumHeight(40)
        btn_a.clicked.connect(lambda: self._exp_do("nieuw"))

        self.btn_exp_b = QPushButton("Export B — Samenvoeg update (WP All Import)")
        self.btn_exp_b.setObjectName("secondary")
        self.btn_exp_b.setMinimumHeight(40)
        self.btn_exp_b.clicked.connect(lambda: self._exp_do("update"))

        btn_ab = QPushButton("Exporteer beide tegelijk")
        btn_ab.setObjectName("secondary")
        btn_ab.clicked.connect(lambda: self._exp_do("beide"))

        lay.addWidget(btn_a)
        lay.addWidget(self.btn_exp_b)
        lay.addWidget(btn_ab)

        self.exp_log = QTextEdit()
        self.exp_log.setReadOnly(True)
        self.exp_log.setMinimumHeight(120)
        self.exp_log.setPlaceholderText("Export log verschijnt hier...")
        lay.addWidget(self.exp_log, stretch=1)

        nav = QHBoxLayout()
        btn_terug = QPushButton("← Terug naar overzicht")
        btn_terug.setObjectName("secondary")
        btn_terug.clicked.connect(self._nav_overzicht)
        btn_mark = QPushButton("Markeer als geëxporteerd")
        btn_mark.setObjectName("secondary")
        btn_mark.clicked.connect(self._exp_markeer)
        nav.addWidget(btn_terug)
        nav.addStretch()
        nav.addWidget(btn_mark)
        lay.addLayout(nav)

        return page

    def _exp_refresh_summary(self):
        if not self._active_krat:
            return
        artikelen = self._active_krat.get("artikelen", [])
        n_nieuw    = sum(1 for a in artikelen if a.get("samenvoeg_beslissing") != "update")
        n_update   = sum(1 for a in artikelen if a.get("samenvoeg_beslissing") == "update")
        n_beprijsd = sum(1 for a in artikelen if a.get("prijs_status") == "beprijsd")
        n_skip     = sum(1 for a in artikelen if a.get("prijs_status") == "overgeslagen")

        self.exp_header.setText(f"Export — {self._active_krat.get('naam', '')}")
        self.exp_summary.setText(
            f"Nieuwe artikelen       (Export A):  {n_nieuw}\n"
            f"Samenvoeg update       (Export B):  {n_update}\n"
            f"Beprijsd:                            {n_beprijsd}\n"
            f"Overgeslagen (→ uitverkocht €0):     {n_skip}\n"
            f"\nLocatie nieuwe artikelen:  {self._active_krat.get('locatie', '—')}"
        )
        self.btn_exp_b.setEnabled(n_update > 0)

    def _exp_do(self, mode: str):
        if not self._active_krat:
            return

        out_dir = output_root() / "kratten"
        out_dir.mkdir(parents=True, exist_ok=True)

        naam_safe = self._active_krat.get("naam", "krat").replace(" ", "_")
        today = date.today().isoformat()
        exported = []

        if mode in ("nieuw", "beide"):
            path = str(out_dir / f"{naam_safe}_{today}_export_A_nieuw.xlsx")
            result = export_nieuwe_artikelen(self._active_krat, path)
            if result:
                exported.append(result)
                self.exp_log.append(f"✅ Export A: {result}")
            else:
                self.exp_log.append("ℹ️ Export A: geen nieuwe artikelen.")

        if mode in ("update", "beide"):
            path = str(out_dir / f"{naam_safe}_{today}_export_B_samenvoeg.xlsx")
            result = export_samenvoeg_update(self._active_krat, path)
            if result:
                exported.append(result)
                self.exp_log.append(f"✅ Export B: {result}")
            else:
                self.exp_log.append("ℹ️ Export B: geen samenvoeg-artikelen.")

        if exported:
            try:
                os.startfile(str(out_dir))
            except Exception:
                pass
            QMessageBox.information(
                self, "Export klaar",
                f"Opgeslagen in:\n{out_dir}\n\nDe map is geopend."
            )

    def _exp_markeer(self):
        if not self._active_krat:
            return
        self._active_krat["status"] = "geexporteerd"
        save_krat(self._active_krat)
        self.exp_log.append("✅ Krat gemarkeerd als geëxporteerd.")

    # ──────────────────────────────────────────────────────────
    # Navigatie
    # ──────────────────────────────────────────────────────────

    def _nav_overzicht(self):
        self._active_krat = None
        self._refresh_overzicht()
        self.stack.setCurrentIndex(0)

    def _show_inventarisatie(self, krat: dict):
        self._active_krat = krat
        krat["status"] = "inventarisatie"
        save_krat(krat)
        locatie = krat.get("locatie", "—")
        self.inv_header.setText(
            f"Inventarisatie — {krat.get('naam', '')}   (Locatie: {locatie})"
        )
        self._inv_reset_form()
        self._inv_refresh_table()
        self.stack.setCurrentIndex(1)

    def _show_beprijzing(self, krat: dict):
        self._active_krat = krat
        krat["status"] = "beprijzing"
        save_krat(krat)
        idx = krat.get("beprijzing_index", 0)
        artikelen = krat.get("artikelen", [])
        idx = max(0, min(idx, len(artikelen) - 1))
        self._bp_show_artikel(idx)
        self.stack.setCurrentIndex(2)

    def _show_export(self, krat: dict):
        self._active_krat = krat
        self._exp_refresh_summary()
        self.exp_log.clear()
        self.stack.setCurrentIndex(3)
