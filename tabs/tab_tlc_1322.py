# ============================================================
# tabs/tab_tlc_1322.py
# TLC / 1322 – Interne verkoop (MAANDLIJST)
#
# Werkt samen met engine_tlc_1322.py
#
# Flow:
# 1. Upload TLC basislijst
# 2. Upload CMS 1322 bestellingen (meerdere mogelijk)
# 3. Run → Picklijst + TLC_NIEUW + UITVERKOCHT
# ============================================================

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog,
    QMessageBox, QTextEdit
)
from engines.engine_tlc_1322 import TLC1322Engine
from PySide6.QtCore import Qt
import os
from datetime import datetime
from utils.paths import output_root

BASE_DIR = str(output_root() / "1322")


class TabTLC1322(QWidget):
    def __init__(self, app_state):
        super().__init__()
        self.app_state = app_state  # niet functioneel nodig, wel consistent
        self.engine = TLC1322Engine()
        self.engine.set_base_dir(BASE_DIR)
        self._build_ui()

    # --------------------------------------------------------
    # UI
    # --------------------------------------------------------
    def _build_ui(self):
        main = QVBoxLayout(self)
        main.setSpacing(14)
        main.setContentsMargins(20, 20, 20, 20)

        # ----------------------------------------------------
        # Titel
        # ----------------------------------------------------
        title = QLabel("TLC / 1322 – Interne verkoop (maandlijst)")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")

        uitleg = QLabel(
            "Deze tool verwerkt interne verkopen (1322).\n\n"
            "• De PICKLIJST bevat ALLE bestelde artikelen\n"
            "• TLC_NIEUW bevat alleen volledig leverbare artikelen\n"
            "• UITVERKOCHT toont tekorten en artikelen met voorraad 0\n\n"
            "Omdat dit een maandlijst is, kunnen tekorten voorkomen."
        )
        uitleg.setWordWrap(True)
        uitleg.setStyleSheet("color: #555;")

        main.addWidget(title)
        main.addWidget(uitleg)

        # ----------------------------------------------------
        # Stap 1 – TLC basislijst
        # ----------------------------------------------------
        step1 = QLabel("① TLC basislijst (automatisch)")
        step1.setStyleSheet("font-size: 14px; font-weight: bold;")
        main.addWidget(step1)

        row1 = QHBoxLayout()
        btn_tlc = QPushButton("📂 Open TLC map (alleen bekijken)")
        btn_tlc.setFixedHeight(30)
        btn_tlc.setStyleSheet("font-size: 11px;")
        btn_tlc.clicked.connect(self.on_upload_tlc)

        btn_reset = QPushButton("Reset alles")
        btn_reset.clicked.connect(self.on_reset)

        row1.addWidget(btn_tlc)
        row1.addWidget(btn_reset)
        row1.addStretch()
        main.addLayout(row1)

        self.lbl_tlc = QLabel("❌ Nog geen TLC basislijst geladen")
        self.lbl_tlc.setWordWrap(True)
        self.lbl_tlc.setStyleSheet(
            """
            QLabel {
                background-color: palette(base);
                color: palette(text);
                border: 1px solid palette(mid);
                padding: 8px;
            }
            """
        )

        main.addWidget(self.lbl_tlc)
        self.lbl_tlc_warning = QLabel(
            "⚠️ Let op: open TLC_1.xlsx NIET in Excel tijdens het verwerken.\n"
            "Als het bestand open staat krijg je 'Permission denied'."
        )
        self.lbl_tlc_warning.setWordWrap(True)
        self.lbl_tlc_warning.setStyleSheet("color:#aa0000; font-weight:bold;")
        main.addWidget(self.lbl_tlc_warning)

        # ----------------------------------------------------
        # Stap 2 – CMS 1322
        # ----------------------------------------------------
        step2 = QLabel("② CMS 1322 bestellingen")
        step2.setStyleSheet("font-size: 14px; font-weight: bold;")
        main.addWidget(step2)

        row2 = QHBoxLayout()
        btn_cms = QPushButton("Add CMS 1322 orders (.xlsx)")
        btn_cms.setFixedHeight(36)
        btn_cms.clicked.connect(self.on_add_cms)

        row2.addWidget(btn_cms)
        row2.addStretch()
        main.addLayout(row2)

        self.lbl_cms = QLabel("0 CMS 1322 bestand(en) toegevoegd")
        self.lbl_cms.setStyleSheet("color:#555;")
        main.addWidget(self.lbl_cms)

        # ----------------------------------------------------
        # Stap 3 – Run
        # ----------------------------------------------------
        step3 = QLabel("③ Verwerken")
        step3.setStyleSheet("font-size: 14px; font-weight: bold;")
        main.addWidget(step3)

        row3 = QHBoxLayout()
        btn_run = QPushButton("Start 1322 verwerking")
        btn_run.setFixedHeight(42)
        btn_run.setStyleSheet("font-weight: bold;")
        btn_run.clicked.connect(self.on_run)

        btn_open = QPushButton("Open output folder")
        btn_open.clicked.connect(self.on_open_output)

        row3.addWidget(btn_run)
        row3.addWidget(btn_open)
        row3.addStretch()
        main.addLayout(row3)

        # ----------------------------------------------------
        # Log
        # ----------------------------------------------------
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(180)
        self.log.setPlaceholderText("Log verschijnt hier...")
        main.addWidget(self.log, stretch=1)

        self._sync_labels()

    # --------------------------------------------------------
    # Helpers
    # --------------------------------------------------------
    def _sync_labels(self):
        p = self._active_tlc_path()
        if os.path.exists(p):
            mtime = os.path.getmtime(p)
            laatst = datetime.fromtimestamp(mtime).strftime("%d-%m-%Y %H:%M:%S")

            self.lbl_tlc.setText(
                "✅ TLC automatisch ingeladen\n"
                f"Bestand: {p}\n"
                f"Laatst gewijzigd: {laatst}"
            )
        else:
            self.lbl_tlc.setText(
                "❌ TLC ontbreekt\n"
                f"Verwacht: {p}\n"
                "Plaats hier het bestand TLC_1.xlsx"
            )


        self.lbl_cms.setText(
            f"{len(self.engine.cms_paths)} CMS 1322 bestand(en) toegevoegd"
        )
    def _active_tlc_path(self):
        return os.path.join(BASE_DIR, "TLC", "TLC_1.xlsx")

    # --------------------------------------------------------
    # Actions
    # --------------------------------------------------------
    def on_upload_tlc(self):
        folder = os.path.join(BASE_DIR, "TLC")
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

    def on_add_cms(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecteer CMS 1322 bestellingen",
            "",
            "Excel bestanden (*.xlsx)"
        )
        if not paths:
            return

        for p in paths:
            self.engine.add_cms_1322(p)

        self.log.append(f"{len(paths)} CMS 1322 bestand(en) toegevoegd")
        self._sync_labels()

    def on_run(self):
        self.engine.set_base_dir(BASE_DIR)
        try:
            # WC-export is al geladen via Start-tab; gebruiken voor model bij D-locaties
            self.engine.set_wc_df(self.app_state.wc_df)
            results = self.engine.run()
        except Exception as e:
            QMessageBox.critical(self, "Fout", str(e))
            return

        self.log.append("\nVerwerking voltooid. Output:")
        for r in results:
            self.log.append(f"- {r}")

        QMessageBox.information(
            self,
            "1322 klaar",
            "1322 verwerking is succesvol afgerond."
        )

        # na run resetten voor nieuwe batch
        self.on_reset()

    def on_reset(self):
        self.engine.clear()
        self.log.clear()
        self._sync_labels()

    def on_open_output(self):
        folder = output_root() / "1322"
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))
