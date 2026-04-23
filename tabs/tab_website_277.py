# tabs/tab_website_277.py
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt, QObject, QEvent
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QFrame, QAbstractItemView, QScrollArea, QApplication,
    QSizePolicy,
)

from engines.engine_website_277 import Website277Engine
from services.batch_state import BatchStore
from utils.paths import output_root
from utils.theme import apply_theme, is_dark_mode


# Stuur muisscroll van een tabel door naar de outer scroll area
class _WheelForwarder(QObject):
    def __init__(self, target, parent=None):
        super().__init__(parent)
        self._target = target

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Wheel:
            QApplication.sendEvent(self._target, event)
            return True
        return False


class TabWebsite277(QWidget):

    S_IDLE         = "idle"
    S_BESTELLING   = "bestelling"   # stap 1 klaar, stap 2 actief
    S_AFGEBOEKT    = "afgeboekt"    # stap 2 klaar, stap 3 actief
    S_WACHT_IMPORT = "wacht"        # export klaar, wacht op "ik heb het gedaan"

    def __init__(self, app_state):
        super().__init__()
        self.app_state = app_state
        self.engine    = Website277Engine(app_state)
        self.store     = BatchStore()

        self._state       = self.S_IDLE
        self._result      = None
        self._update_path = None

        self._build_ui()
        self._restore_openstaande_bestelling()

    # ─────────────────────────────────────────────────────────
    # UI bouwen
    # ─────────────────────────────────────────────────────────

    def _build_ui(self):
        apply_theme(self)
        self._dark = is_dark_mode(self)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(12)

        lbl = QLabel("Bestelling verwerken")
        lbl.setStyleSheet("font-size: 22px; font-weight: bold;")
        outer.addWidget(lbl)

        self.banner = self._maak_banner()
        outer.addWidget(self.banner)
        self.banner.hide()

        self.stap_indicator = self._maak_stap_indicator()
        outer.addWidget(self.stap_indicator)

        # Scrollbaar gebied — stappen verticaal onder elkaar
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        content = QWidget()
        content.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        vbox = QVBoxLayout(content)
        vbox.setContentsMargins(2, 2, 2, 2)
        vbox.setSpacing(14)

        self.frame_stap1 = self._maak_stap1()
        self.frame_stap2 = self._maak_stap2()
        self.frame_stap3 = self._maak_stap3()

        vbox.addWidget(self.frame_stap1)
        vbox.addWidget(self.frame_stap2)
        vbox.addWidget(self.frame_stap3)

        self._scroll.setWidget(content)
        outer.addWidget(self._scroll, stretch=1)

        # Stuur scroll van tabellen door naar de outer scroll area
        forwarder = _WheelForwarder(self._scroll.viewport(), self)
        self.tbl_order.viewport().installEventFilter(forwarder)
        self.tbl_update.viewport().installEventFilter(forwarder)

        self._ververs_ui()

    # ── Banner ────────────────────────────────────────────────

    def _maak_banner(self):
        frame = QFrame()
        frame.setObjectName("banner")
        frame.setStyleSheet("""
            QFrame#banner {
                background: rgba(234,88,12,0.12);
                border: 2px solid #ea580c;
                border-radius: 10px;
            }
        """)
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(10)

        self.banner_lbl = QLabel()
        self.banner_lbl.setWordWrap(True)
        tekst_kleur = "#fb923c" if self._dark else "#9a3412"
        self.banner_lbl.setStyleSheet(
            f"font-size: 13px; font-weight: bold; color: {tekst_kleur};"
            "background: transparent; border: none;"
        )
        lay.addWidget(self.banner_lbl)

        rij = QHBoxLayout()
        self.btn_banner_ja  = QPushButton("Ja, dat was al gedaan")
        self.btn_banner_nee = QPushButton("Nee, ik doe dat eerst")
        self.btn_banner_ja.setObjectName("secondary")
        self.btn_banner_nee.setObjectName("primary")
        self.btn_banner_ja.setMinimumHeight(40)
        self.btn_banner_nee.setMinimumHeight(40)
        self.btn_banner_ja.clicked.connect(self._banner_al_gedaan)
        self.btn_banner_nee.clicked.connect(self._banner_nog_doen)
        rij.addWidget(self.btn_banner_nee)
        rij.addWidget(self.btn_banner_ja)
        rij.addStretch()
        lay.addLayout(rij)
        return frame

    # ── Stap-indicator ────────────────────────────────────────

    def _maak_stap_indicator(self):
        frame = QFrame()
        lay = QHBoxLayout(frame)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)

        self._prog = []
        namen = ["① Bestelling laden", "② Afboeken", "③ Website bijwerken"]
        for i, naam in enumerate(namen):
            lbl = QLabel(naam)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setMinimumHeight(38)
            self._prog.append(lbl)
            lay.addWidget(lbl, stretch=1)
            if i < len(namen) - 1:
                pijl = QLabel("›")
                pijl.setAlignment(Qt.AlignCenter)
                pijl.setStyleSheet("font-size: 20px; color: palette(mid); background: transparent;")
                lay.addWidget(pijl)
        return frame

    # ── Stap 1 ────────────────────────────────────────────────

    def _maak_stap1(self):
        frame = self._stap_frame()
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(10)

        lay.addWidget(self._stap_titel("Stap 1 — Laad de bestelling"))
        lay.addWidget(self._info_lbl(
            "Klik hieronder en kies het Excel-bestand dat je via e-mail hebt ontvangen."
        ))

        btn = QPushButton("Laad bestelling")
        btn.setObjectName("primary")
        btn.setMinimumHeight(52)
        btn.setStyleSheet("font-size: 13pt;")
        btn.clicked.connect(self._laad_bestelling)
        lay.addWidget(btn)

        self.lbl_geladen = QLabel("Nog geen bestelling geladen.")
        self.lbl_geladen.setWordWrap(True)
        self.lbl_geladen.setStyleSheet(self._info_box_style())
        lay.addWidget(self.lbl_geladen)

        self.btn_volgende1 = QPushButton("Volgende  →")
        self.btn_volgende1.setObjectName("primary")
        self.btn_volgende1.setMinimumHeight(44)
        self.btn_volgende1.setEnabled(False)
        self.btn_volgende1.clicked.connect(lambda: self._zet_staat(self.S_BESTELLING))
        lay.addWidget(self.btn_volgende1)

        return frame

    # ── Stap 2 ────────────────────────────────────────────────

    def _maak_stap2(self):
        frame = self._stap_frame()
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(10)

        lay.addWidget(self._stap_titel("Stap 2 — Controleer en afboeken"))
        lay.addWidget(self._info_lbl(
            "Hieronder zie je wat er besteld is. Controleer of de aantallen kloppen."
        ))

        self.tbl_order = QTableWidget()
        self.tbl_order.setColumnCount(3)
        self.tbl_order.setHorizontalHeaderLabels(["Artikelnummer", "Aantal", "Factuurnummer"])
        self.tbl_order.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tbl_order.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tbl_order.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl_order.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_order.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tbl_order.setMinimumHeight(60)
        lay.addWidget(self.tbl_order)

        btn = QPushButton("Afboeken en picklijst openen")
        btn.setObjectName("primary")
        btn.setMinimumHeight(56)
        btn.setStyleSheet("font-size: 13pt;")
        btn.clicked.connect(self._afboeken)
        lay.addWidget(btn)

        lbl_hint = QLabel("De picklijst opent automatisch zodra het afboeken klaar is.")
        lbl_hint.setStyleSheet("color: palette(mid); font-style: italic; font-size: 10pt;")
        lay.addWidget(lbl_hint)

        return frame

    # ── Stap 3 ────────────────────────────────────────────────

    def _maak_stap3(self):
        frame = self._stap_frame()
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(10)

        lay.addWidget(self._stap_titel("Stap 3 — Zet de wijzigingen op de website"))
        lay.addWidget(self._info_lbl(
            "Hieronder zie je de nieuwe voorraad per product. "
            "Klopt er een getal niet? Dubbelklik erop en pas het aan."
        ))

        self.tbl_update = QTableWidget()
        self.tbl_update.setColumnCount(3)
        self.tbl_update.setHorizontalHeaderLabels(["Artikelnummer", "Nieuwe voorraad", "Locatie"])
        self.tbl_update.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tbl_update.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tbl_update.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl_update.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tbl_update.setMinimumHeight(60)
        lay.addWidget(self.tbl_update)

        btn_export = QPushButton("Maak bestand klaar voor de website")
        btn_export.setObjectName("primary")
        btn_export.setMinimumHeight(52)
        btn_export.setStyleSheet("font-size: 13pt;")
        btn_export.clicked.connect(self._exporteer_update)
        lay.addWidget(btn_export)

        self.lbl_export_klaar = QLabel("")
        self.lbl_export_klaar.setWordWrap(True)
        self.lbl_export_klaar.hide()
        tekst_kleur = "#86efac" if self._dark else "#14532d"
        bg_kleur    = "rgba(22,163,74,0.18)" if self._dark else "rgba(22,163,74,0.10)"
        self.lbl_export_klaar.setStyleSheet(f"""
            QLabel {{
                background: {bg_kleur};
                border: 2px solid #16a34a;
                border-radius: 8px;
                padding: 12px;
                font-size: 12pt;
                font-weight: bold;
                color: {tekst_kleur};
            }}
        """)
        lay.addWidget(self.lbl_export_klaar)

        self.btn_klaar = QPushButton("✅   Ik heb het geïmporteerd — Klaar!")
        self.btn_klaar.setObjectName("primary")
        self.btn_klaar.setMinimumHeight(58)
        self.btn_klaar.setStyleSheet(
            "font-size: 14pt; font-weight: bold;"
            "background-color: #16a34a; border-color: #16a34a; color: white;"
        )
        self.btn_klaar.hide()
        self.btn_klaar.clicked.connect(self._markeer_klaar)
        lay.addWidget(self.btn_klaar)

        return frame

    # ─────────────────────────────────────────────────────────
    # State machine
    # ─────────────────────────────────────────────────────────

    def _zet_staat(self, staat):
        self._state = staat
        self._ververs_ui()

    def _ververs_ui(self):
        s = self._state

        ACTIEF   = ("font-size:11pt; font-weight:bold; padding:6px 14px;"
                    "background:#2563EB; color:white; border:2px solid #2563EB; border-radius:4px;")
        KLAAR    = ("font-size:11pt; font-weight:bold; padding:6px 14px;"
                    "background:#16a34a; color:white; border:2px solid #16a34a; border-radius:4px;")
        INACTIEF = ("font-size:11pt; padding:6px 14px;"
                    "background:palette(base); color:palette(mid); border:1px solid palette(mid); border-radius:4px;")

        if s == self.S_IDLE:
            stijlen = [ACTIEF, INACTIEF, INACTIEF]
        elif s == self.S_BESTELLING:
            stijlen = [KLAAR, ACTIEF, INACTIEF]
        else:
            stijlen = [KLAAR, KLAAR, ACTIEF]

        for lbl, stijl in zip(self._prog, stijlen):
            lbl.setStyleSheet(stijl)

        actief_rand   = "QFrame { border:2px solid #2563EB; border-radius:10px; background:palette(base); }"
        inactief_rand = "QFrame { border:2px solid palette(mid); border-radius:10px; background:palette(window); }"

        stap1_aan = s == self.S_IDLE
        stap2_aan = s == self.S_BESTELLING
        stap3_aan = s in (self.S_AFGEBOEKT, self.S_WACHT_IMPORT)

        self.frame_stap1.setStyleSheet(actief_rand if stap1_aan else inactief_rand)
        self.frame_stap2.setStyleSheet(actief_rand if stap2_aan else inactief_rand)
        self.frame_stap3.setStyleSheet(actief_rand if stap3_aan else inactief_rand)

        # Frames altijd klikbaar zodat handlers een duidelijke melding tonen
        self.frame_stap1.setEnabled(True)
        self.frame_stap2.setEnabled(True)
        self.frame_stap3.setEnabled(True)

        self.btn_klaar.setVisible(s == self.S_WACHT_IMPORT)

    # ─────────────────────────────────────────────────────────
    # Banner: openstaande bestelling
    # ─────────────────────────────────────────────────────────

    def _restore_openstaande_bestelling(self):
        batch = self.store.get_latest_open_batch("277")
        if not batch:
            return

        try:
            dt    = datetime.fromisoformat(batch.get("created_at", ""))
            datum = dt.strftime("%d-%m-%Y om %H:%M")
        except Exception:
            datum = batch.get("created_at", "")

        self.banner_lbl.setText(
            f"⚠️  Let op — Er staat nog een niet-afgeronde bestelling open van {datum}.\n\n"
            "Heb je het bestand al in WP All Import op de website gezet?"
        )
        self.banner.show()
        self._result      = batch
        self._update_path = batch.get("update_path", "")

    def _banner_al_gedaan(self):
        batch = self.store.get_latest_open_batch("277")
        if batch:
            self.store.mark_imported(batch["batch_id"])
        self.banner.hide()
        self._reset()

    def _banner_nog_doen(self):
        self.banner.hide()
        if self._update_path and os.path.exists(self._update_path):
            self._laad_update_tabel(self._update_path)
            self.lbl_export_klaar.setText(
                f"Bestand staat klaar.\n\n"
                f"Ga naar WP All Import en importeer:\n\n"
                f"{Path(self._update_path).name}\n\n"
                f"Volledige locatie:\n{self._update_path}"
            )
            self.lbl_export_klaar.show()
            self._zet_staat(self.S_WACHT_IMPORT)
        else:
            QMessageBox.warning(
                self, "Bestand niet gevonden",
                "Het bestand van de vorige bestelling is niet meer te vinden.\n"
                "Je moet de bestelling opnieuw afboeken."
            )
            batch = self.store.get_latest_open_batch("277")
            if batch:
                self.store.mark_imported(batch["batch_id"])
            self._reset()

    # ─────────────────────────────────────────────────────────
    # Stap 1 — Bestelling laden
    # ─────────────────────────────────────────────────────────

    def _laad_bestelling(self):
        # Als je al verder bent, vraag bevestiging voor reset
        if self._state != self.S_IDLE:
            antw = QMessageBox.question(
                self, "Nieuwe bestelling laden?",
                "Je bent al bezig met een bestelling.\n\n"
                "Als je een nieuw bestand laadt wordt de huidige voortgang gewist.\n\n"
                "Weet je het zeker?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if antw != QMessageBox.Yes:
                return
            self._reset()

        paden, _ = QFileDialog.getOpenFileNames(
            self, "Kies het bestellingbestand", "", "Excel bestanden (*.xlsx)"
        )
        if not paden:
            return

        self.engine.clear()
        for p in paden:
            self.engine.add_cms_277(p)

        self._vul_order_tabel()

        namen = ", ".join(Path(p).name for p in paden)
        n     = len(paden)
        self.lbl_geladen.setText(
            f"✅  {n} bestand{'en' if n > 1 else ''} geladen:\n{namen}"
        )
        self.btn_volgende1.setEnabled(True)

    # Bekende kolomnamen die nooit een artikelnummer zijn
    _SKIP_HEADERS = {
        "id", "title", "artikelnummer", "naam", "omschrijving",
        "aantal", "prijs", "factuur", "factuurnummer", "description",
        "productcategorieën", "stock", "locatie", "short description",
    }

    def _vul_order_tabel(self):
        rijen = []
        for p in self.engine.cms_paths:
            try:
                df = pd.read_excel(p, header=None, dtype=str).fillna("")
                df.columns = range(len(df.columns))
                for _, r in df.iterrows():
                    title   = str(r.get(0, "")).strip()
                    aantal  = str(r.get(2, "")).strip()
                    factuur = str(r.get(4, "")).strip()
                    if title and title.lower() not in self._SKIP_HEADERS:
                        rijen.append((title, aantal, factuur))
            except Exception:
                pass

        self.tbl_order.setRowCount(len(rijen))
        for i, (title, aantal, factuur) in enumerate(rijen):
            for kolom, tekst in enumerate([title, aantal, factuur]):
                item = QTableWidgetItem(tekst)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.tbl_order.setItem(i, kolom, item)

        self._pas_tabel_hoogte_aan(self.tbl_order)

    # ─────────────────────────────────────────────────────────
    # Stap 2 — Afboeken
    # ─────────────────────────────────────────────────────────

    def _afboeken(self):
        if self._state != self.S_BESTELLING:
            QMessageBox.warning(
                self, "Stap 1 niet voltooid",
                "Voltooi eerst stap 1:\n\n"
                "1. Klik op 'Laad bestelling' en kies het Excel-bestand.\n"
                "2. Klik daarna op 'Volgende →'."
            )
            return

        if getattr(self.app_state, "wc_df", None) is None:
            QMessageBox.warning(
                self, "WooCommerce export ontbreekt",
                "Ga eerst naar het tabblad 'Start' en laad de WooCommerce export."
            )
            return

        try:
            result = self.engine.run()
        except Exception as e:
            QMessageBox.critical(self, "Er ging iets mis", str(e))
            return

        self._result      = result
        self._update_path = result.get("update_path", "")

        self.store.create_batch(result)

        pick = result.get("pick_path", "")
        if pick and os.path.exists(pick):
            try:
                os.startfile(pick)
            except Exception:
                pass

        if self._update_path and os.path.exists(self._update_path):
            self._laad_update_tabel(self._update_path)

        self._zet_staat(self.S_AFGEBOEKT)

        QMessageBox.information(
            self, "Afboeken klaar",
            "Klaar!\n\n"
            "De picklijst is automatisch geopend.\n\n"
            "Ga de producten pakken en kom daarna terug\n"
            "voor stap 3 om de website bij te werken."
        )

    def _laad_update_tabel(self, pad: str):
        try:
            df = pd.read_excel(pad, dtype=str).fillna("")
        except Exception:
            return

        col_title = next((c for c in df.columns if str(c).lower() == "title"), None)
        col_stock = next((c for c in df.columns if str(c).lower() == "stock"), None)
        col_loc   = next((c for c in df.columns if str(c).lower() == "locatie"), None)

        if not col_title or not col_stock:
            return

        self._update_df_kolommen = (col_title, col_stock, col_loc)

        self.tbl_update.setRowCount(len(df))
        for i, (_, r) in enumerate(df.iterrows()):
            title = str(r.get(col_title, ""))
            stock = str(r.get(col_stock, ""))
            loc   = str(r.get(col_loc, "")) if col_loc else ""

            item_t = QTableWidgetItem(title)
            item_t.setFlags(item_t.flags() & ~Qt.ItemIsEditable)

            item_s = QTableWidgetItem(stock)

            item_l = QTableWidgetItem(loc)
            item_l.setFlags(item_l.flags() & ~Qt.ItemIsEditable)

            self.tbl_update.setItem(i, 0, item_t)
            self.tbl_update.setItem(i, 1, item_s)
            self.tbl_update.setItem(i, 2, item_l)

        self._pas_tabel_hoogte_aan(self.tbl_update)

    # ─────────────────────────────────────────────────────────
    # Stap 3 — Website update exporteren
    # ─────────────────────────────────────────────────────────

    def _exporteer_update(self):
        if self._state not in (self.S_AFGEBOEKT, self.S_WACHT_IMPORT):
            QMessageBox.warning(
                self, "Stap 2 niet voltooid",
                "Voltooi eerst stap 1 en stap 2 voordat je het website-bestand aanmaakt."
            )
            return

        if not self._update_path or not os.path.exists(self._update_path):
            QMessageBox.warning(self, "Bestand niet gevonden",
                                "Het website-updatebestand is niet gevonden.")
            return

        try:
            df = pd.read_excel(self._update_path, dtype=str).fillna("")
            col_title, col_stock, _ = getattr(self, "_update_df_kolommen", ("Title", "Stock", "Locatie"))

            if col_stock in df.columns and col_title in df.columns:
                for rij in range(self.tbl_update.rowCount()):
                    t_item = self.tbl_update.item(rij, 0)
                    s_item = self.tbl_update.item(rij, 1)
                    if t_item and s_item and rij < len(df):
                        df.at[df.index[rij], col_stock] = s_item.text()

            df.to_excel(self._update_path, index=False)

            # Open de map zodat het bestand direct zichtbaar is
            try:
                os.startfile(str(Path(self._update_path).parent))
            except Exception:
                pass

            naam = Path(self._update_path).name
            self.lbl_export_klaar.setText(
                f"✅  Bestand is klaar!\n\n"
                f"Ga nu naar WP All Import en importeer dit bestand:\n\n"
                f"📄  {naam}\n\n"
                f"Locatie:\n{self._update_path}"
            )
            self.lbl_export_klaar.show()
            self._zet_staat(self.S_WACHT_IMPORT)

        except Exception as e:
            QMessageBox.critical(self, "Fout bij opslaan", str(e))

    def _markeer_klaar(self):
        batch = self.store.get_latest_open_batch("277")
        if batch:
            self.store.mark_imported(batch["batch_id"])
        self._reset()
        QMessageBox.information(
            self, "Klaar!",
            "Goed gedaan!\n\n"
            "De bestelling is volledig verwerkt en de website is bijgewerkt.\n\n"
            "Je kunt nu een nieuwe bestelling laden."
        )

    def _reset(self):
        self.engine.clear()
        self._result      = None
        self._update_path = None

        self.tbl_order.setRowCount(0)
        self.tbl_update.setRowCount(0)
        self.lbl_geladen.setText("Nog geen bestelling geladen.")
        self.btn_volgende1.setEnabled(False)
        self.lbl_export_klaar.hide()
        self.banner.hide()

        self._zet_staat(self.S_IDLE)

    # ─────────────────────────────────────────────────────────
    # Hulpfuncties
    # ─────────────────────────────────────────────────────────

    @staticmethod
    def _stap_frame():
        f = QFrame()
        f.setStyleSheet(
            "QFrame { border:2px solid palette(mid); border-radius:10px; background:palette(base); }"
        )
        return f

    @staticmethod
    def _stap_titel(tekst):
        lbl = QLabel(tekst)
        lbl.setStyleSheet("font-size:15px; font-weight:bold; border:none; background:transparent;")
        return lbl

    @staticmethod
    def _info_lbl(tekst):
        lbl = QLabel(tekst)
        lbl.setWordWrap(True)
        lbl.setStyleSheet("border:none; background:transparent;")
        return lbl

    @staticmethod
    def _info_box_style():
        return """
            QLabel {
                background: palette(window);
                border: 1px solid palette(mid);
                border-radius: 6px;
                padding: 8px;
            }
        """

    @staticmethod
    def _pas_tabel_hoogte_aan(tabel: QTableWidget):
        """Zet de tabel op exact de hoogte die nodig is voor alle rijen — geen interne scroll."""
        hoogte = (
            tabel.horizontalHeader().height()
            + tabel.verticalHeader().length()
            + tabel.frameWidth() * 2
            + 4  # kleine buffer voor afrondingen
        )
        tabel.setFixedHeight(max(hoogte, 60))
