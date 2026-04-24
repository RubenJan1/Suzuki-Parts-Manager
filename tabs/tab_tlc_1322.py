# tabs/tab_tlc_1322.py
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

from engines.engine_tlc_1322 import TLC1322Engine
from utils.paths import output_root
from utils.theme import apply_theme, is_dark_mode

BASE_DIR = str(output_root() / "1322")


class _WheelForwarder(QObject):
    def __init__(self, target, parent=None):
        super().__init__(parent)
        self._target = target

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Wheel:
            QApplication.sendEvent(self._target, event)
            return True
        return False


class TabTLC1322(QWidget):

    S_IDLE       = "idle"
    S_GELADEN    = "geladen"
    S_VERWERKT   = "verwerkt"

    def __init__(self, app_state):
        super().__init__()
        self.app_state = app_state
        self.engine    = TLC1322Engine()
        self.engine.set_base_dir(BASE_DIR)

        self._state = self.S_IDLE

        self._build_ui()
        self._ververs_ui()

    # ─────────────────────────────────────────────────────────
    # UI bouwen
    # ─────────────────────────────────────────────────────────

    def _build_ui(self):
        apply_theme(self)
        self._dark = is_dark_mode(self)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(12)

        lbl = QLabel("TLC / 1322 — Bestelling verwerken")
        lbl.setStyleSheet("font-size: 22px; font-weight: bold;")
        outer.addWidget(lbl)

        self.stap_indicator = self._maak_stap_indicator()
        outer.addWidget(self.stap_indicator)

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

        forwarder = _WheelForwarder(self._scroll.viewport(), self)
        self.tbl_order.viewport().installEventFilter(forwarder)
        self.tbl_result.viewport().installEventFilter(forwarder)

    # ── Stap-indicator ────────────────────────────────────────

    def _maak_stap_indicator(self):
        frame = QFrame()
        lay = QHBoxLayout(frame)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)

        self._prog = []
        namen = ["① TLC controleren", "② Bestelling laden", "③ Verwerken"]
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

        lay.addWidget(self._stap_titel("Stap 1 — TLC controleren"))
        lay.addWidget(self._info_lbl(
            "De TLC-basislijst wordt automatisch gevonden. "
            "Als die ontbreekt, klik dan op 'Open TLC map' en zet het bestand er in."
        ))

        self.lbl_tlc_status = QLabel("")
        self.lbl_tlc_status.setWordWrap(True)
        self.lbl_tlc_status.setStyleSheet(self._info_box_style())
        lay.addWidget(self.lbl_tlc_status)

        rij = QHBoxLayout()
        btn_open_tlc = QPushButton("Open TLC map")
        btn_open_tlc.setObjectName("secondary")
        btn_open_tlc.clicked.connect(self._open_tlc_map)
        rij.addWidget(btn_open_tlc)
        rij.addStretch()
        lay.addLayout(rij)

        self.btn_volgende1 = QPushButton("Volgende  →")
        self.btn_volgende1.setObjectName("primary")
        self.btn_volgende1.setMinimumHeight(44)
        self.btn_volgende1.clicked.connect(lambda: self._zet_staat(self.S_GELADEN))
        lay.addWidget(self.btn_volgende1)

        return frame

    # ── Stap 2 ────────────────────────────────────────────────

    def _maak_stap2(self):
        frame = self._stap_frame()
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(10)

        lay.addWidget(self._stap_titel("Stap 2 — Laad de CMS 1322-bestelling(en)"))
        lay.addWidget(self._info_lbl(
            "Kies het Excel-bestand dat je van CMS hebt ontvangen. "
            "Je kunt meerdere bestanden tegelijk laden."
        ))

        btn_laad = QPushButton("Laad CMS 1322-bestelling(en)")
        btn_laad.setObjectName("primary")
        btn_laad.setMinimumHeight(52)
        btn_laad.setStyleSheet("font-size: 13pt;")
        btn_laad.clicked.connect(self._laad_bestelling)
        lay.addWidget(btn_laad)

        self.lbl_geladen = QLabel("Nog geen bestelling geladen.")
        self.lbl_geladen.setWordWrap(True)
        self.lbl_geladen.setStyleSheet(self._info_box_style())
        lay.addWidget(self.lbl_geladen)

        self.tbl_order = QTableWidget()
        self.tbl_order.setColumnCount(3)
        self.tbl_order.setHorizontalHeaderLabels(["Artikelnummer", "Besteld aantal", "Factuurnummer"])
        self.tbl_order.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tbl_order.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tbl_order.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl_order.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_order.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tbl_order.setMinimumHeight(60)
        lay.addWidget(self.tbl_order)

        btn_verwerken = QPushButton("Afboeken en picklijst openen  →")
        btn_verwerken.setObjectName("primary")
        btn_verwerken.setMinimumHeight(56)
        btn_verwerken.setStyleSheet("font-size: 13pt;")
        btn_verwerken.clicked.connect(self._verwerken)
        lay.addWidget(btn_verwerken)

        return frame

    # ── Stap 3 ────────────────────────────────────────────────

    def _maak_stap3(self):
        frame = self._stap_frame()
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(10)

        lay.addWidget(self._stap_titel("Stap 3 — Controleer het resultaat"))
        lay.addWidget(self._info_lbl(
            "Hieronder zie je wat er geleverd is. "
            "Bij een tekort staat het aantal in de kolom 'Tekort'."
        ))

        self.tbl_result = QTableWidget()
        self.tbl_result.setColumnCount(5)
        self.tbl_result.setHorizontalHeaderLabels(
            ["Artikelnummer", "Besteld", "Geleverd", "Tekort", "Locatie"]
        )
        self.tbl_result.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tbl_result.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tbl_result.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl_result.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.tbl_result.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.tbl_result.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_result.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tbl_result.setMinimumHeight(60)
        lay.addWidget(self.tbl_result)

        self.lbl_opgeslagen = QLabel("")
        self.lbl_opgeslagen.setWordWrap(True)
        self.lbl_opgeslagen.hide()
        tekst_kleur = "#86efac" if self._dark else "#14532d"
        bg_kleur    = "rgba(22,163,74,0.18)" if self._dark else "rgba(22,163,74,0.10)"
        self.lbl_opgeslagen.setStyleSheet(f"""
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
        lay.addWidget(self.lbl_opgeslagen)

        self.btn_klaar = QPushButton("✅   Klaar — nieuwe bestelling beginnen")
        self.btn_klaar.setObjectName("primary")
        self.btn_klaar.setMinimumHeight(58)
        self.btn_klaar.setStyleSheet(
            "font-size: 14pt; font-weight: bold;"
            "background-color: #16a34a; border-color: #16a34a; color: white;"
        )
        self.btn_klaar.hide()
        self.btn_klaar.clicked.connect(self._reset)
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
        elif s == self.S_GELADEN:
            stijlen = [KLAAR, ACTIEF, INACTIEF]
        else:
            stijlen = [KLAAR, KLAAR, ACTIEF]

        for lbl, stijl in zip(self._prog, stijlen):
            lbl.setStyleSheet(stijl)

        actief_rand   = "QFrame { border:2px solid #2563EB; border-radius:10px; background:palette(base); }"
        inactief_rand = "QFrame { border:2px solid palette(mid); border-radius:10px; background:palette(window); }"

        self.frame_stap1.setStyleSheet(actief_rand if s == self.S_IDLE    else inactief_rand)
        self.frame_stap2.setStyleSheet(actief_rand if s == self.S_GELADEN  else inactief_rand)
        self.frame_stap3.setStyleSheet(actief_rand if s == self.S_VERWERKT else inactief_rand)

        # TLC status altijd verversen
        self._ververs_tlc_status()

        # Stap 1: Volgende-knop alleen actief als TLC aanwezig
        tlc_ok = os.path.exists(self._tlc_pad())
        self.btn_volgende1.setEnabled(tlc_ok)

        self.btn_klaar.setVisible(s == self.S_VERWERKT)
        self.lbl_opgeslagen.setVisible(s == self.S_VERWERKT)

    def _ververs_tlc_status(self):
        p = self._tlc_pad()
        if os.path.exists(p):
            mtime = datetime.fromtimestamp(os.path.getmtime(p)).strftime("%d-%m-%Y %H:%M")
            self.lbl_tlc_status.setText(f"✅  TLC gevonden\n{p}\nLaatst gewijzigd: {mtime}")
            self.lbl_tlc_status.setStyleSheet(self._info_box_style("#16a34a"))
        else:
            self.lbl_tlc_status.setText(
                f"❌  TLC niet gevonden\nVerwacht: {p}\n\n"
                "Klik op 'Open TLC map' en zet TLC_1.xlsx in die map."
            )
            self.lbl_tlc_status.setStyleSheet(self._info_box_style("#dc2626"))

    # ─────────────────────────────────────────────────────────
    # Acties
    # ─────────────────────────────────────────────────────────

    def _tlc_pad(self) -> str:
        return os.path.join(BASE_DIR, "TLC", "TLC_1.xlsx")

    def _open_tlc_map(self):
        folder = os.path.join(BASE_DIR, "TLC")
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)
        # Ververs na openen
        self._ververs_tlc_status()
        self.btn_volgende1.setEnabled(os.path.exists(self._tlc_pad()))

    def _laad_bestelling(self):
        if self._state != self.S_GELADEN:
            QMessageBox.warning(self, "Stap 1 niet voltooid",
                                "Controleer eerst de TLC en klik op 'Volgende →'.")
            return

        paden, _ = QFileDialog.getOpenFileNames(
            self, "Kies CMS 1322-bestelling(en)", "", "Excel bestanden (*.xlsx)"
        )
        if not paden:
            return

        self.engine.clear()
        self.engine.set_base_dir(BASE_DIR)
        for p in paden:
            self.engine.add_cms_1322(p)

        self._vul_order_tabel()

        n = len(paden)
        namen = ", ".join(Path(p).name for p in paden)
        self.lbl_geladen.setText(f"✅  {n} bestand{'en' if n > 1 else ''} geladen:\n{namen}")

    _SKIP = {
        "title", "artikelnummer", "naam", "omschrijving", "aantal",
        "prijs", "factuur", "factuurnummer", "id",
    }

    def _vul_order_tabel(self):
        rijen = []
        for p in self.engine.cms_paths:
            try:
                df = pd.read_excel(p, header=None, dtype=str).fillna("")
                for _, r in df.iterrows():
                    title   = str(r.get(0, "")).strip()
                    aantal  = str(r.get(2, "")).strip()
                    factuur = str(r.get(4, "")).strip()
                    if title and title.lower() not in self._SKIP:
                        rijen.append((title, aantal, factuur))
            except Exception:
                pass

        self.tbl_order.setRowCount(len(rijen))
        for i, (title, aantal, factuur) in enumerate(rijen):
            for k, tekst in enumerate([title, aantal, factuur]):
                item = QTableWidgetItem(tekst)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.tbl_order.setItem(i, k, item)
        self._pas_tabel_hoogte_aan(self.tbl_order)

    def _verwerken(self):
        if self._state != self.S_GELADEN:
            QMessageBox.warning(self, "Stap 1 niet voltooid",
                                "Controleer eerst de TLC en klik op 'Volgende →'.")
            return

        if not self.engine.cms_paths:
            QMessageBox.warning(self, "Geen bestelling",
                                "Laad eerst een CMS 1322-bestelling.")
            return

        if getattr(self.app_state, "wc_df", None) is None:
            QMessageBox.warning(self, "WooCommerce export ontbreekt",
                                "Ga naar het tabblad 'Start' en laad de WooCommerce export.")
            return

        try:
            self.engine.set_wc_df(self.app_state.wc_df)
            resultaten = self.engine.run()
        except Exception as e:
            QMessageBox.critical(self, "Fout bij verwerken", str(e))
            return

        # Picklijst openen
        pick = next((r for r in resultaten if "PICKLIJST" in Path(r).name.upper()), None)
        if pick and os.path.exists(pick):
            try:
                os.startfile(pick)
            except Exception:
                pass

        # Resultaattabel vullen vanuit last_invoice_lines
        self._vul_result_tabel()

        # Auto-save naar CMS queue
        try:
            from services.cms_queue import add_run
            lines = getattr(self.engine, "last_invoice_lines", [])
            if lines:
                add_run("1322", lines)
                self.lbl_opgeslagen.setText(
                    f"✅  Verwerkt en opgeslagen voor weekfactuur\n\n"
                    f"{len(lines)} regel(s) opgeslagen voor de CMS factuur.\n"
                    "Ga op donderdag naar de Factuurmaker om de factuur op te maken."
                )
            else:
                self.lbl_opgeslagen.setText(
                    "✅  Verwerkt — geen leverbare artikelen gevonden."
                )
        except Exception:
            self.lbl_opgeslagen.setText("✅  Verwerkt.")

        self._zet_staat(self.S_VERWERKT)

        QMessageBox.information(
            self, "1322 klaar",
            "Afboeken is klaar!\n\n"
            "De picklijst is automatisch geopend.\n\n"
            "De geleverde artikelen zijn opgeslagen voor de weekfactuur.\n"
            "Ga op donderdag naar de Factuurmaker."
        )

    def _vul_result_tabel(self):
        lines = getattr(self.engine, "last_invoice_lines", [])
        self.tbl_result.setRowCount(len(lines))
        for i, r in enumerate(lines):
            tekort = max(0, r["besteld"] - r["geleverd"])
            waarden = [
                r["title"],
                str(r["besteld"]),
                str(r["geleverd"]),
                str(tekort) if tekort else "",
                "",
            ]
            for k, tekst in enumerate(waarden):
                item = QTableWidgetItem(tekst)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                if k == 3 and tekort:
                    from PySide6.QtGui import QColor
                    item.setBackground(QColor(255, 220, 220))
                self.tbl_result.setItem(i, k, item)
        self._pas_tabel_hoogte_aan(self.tbl_result)

    def _reset(self):
        self.engine.clear()
        self.engine.set_base_dir(BASE_DIR)
        self.tbl_order.setRowCount(0)
        self.tbl_result.setRowCount(0)
        self.lbl_geladen.setText("Nog geen bestelling geladen.")
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
    def _info_box_style(border_color: str = None) -> str:
        border = border_color or "palette(mid)"
        return f"""
            QLabel {{
                background: palette(window);
                border: 1px solid {border};
                border-radius: 6px;
                padding: 8px;
            }}
        """

    @staticmethod
    def _pas_tabel_hoogte_aan(tabel: QTableWidget):
        hoogte = (
            tabel.horizontalHeader().height()
            + tabel.verticalHeader().length()
            + tabel.frameWidth() * 2
            + 4
        )
        tabel.setFixedHeight(max(hoogte, 60))
