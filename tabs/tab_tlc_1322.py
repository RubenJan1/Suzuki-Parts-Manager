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
from utils.theme import apply_theme

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
        apply_theme(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        # =========================
        # TITEL
        # =========================
        title = QLabel("TLC / 1322 – Interne verkoop")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        root.addWidget(title)

        # =========================
        # UITLEG
        # =========================
        uitleg = QLabel(
            "Verwerkt interne verkopen (1322) op basis van de actieve TLC.\n\n"
            "① Controleer of TLC_1.xlsx aanwezig is\n"
            "② Voeg CMS 1322 bestanden toe\n"
            "③ Start verwerking\n\n"
            "Output:\n"
            "• Picklijst = alle bestelde artikelen\n"
            "• TLC_NIEUW = volledig leverbare artikelen\n"
            "• UITVERKOCHT = tekorten en voorraad 0"
        )
        uitleg.setWordWrap(True)
        root.addWidget(uitleg)

        # =========================
        # STATUS BLOK
        # =========================
        self.lbl_tlc = QLabel("❌ Nog geen TLC basislijst geladen")
        self.lbl_tlc.setWordWrap(True)
        self.lbl_tlc.setStyleSheet("""
            QLabel {
                background: palette(base);
                border: 1px solid palette(mid);
                border-radius: 8px;
                padding: 10px;
            }
        """)
        root.addWidget(self.lbl_tlc)

        self.lbl_tlc_warning = QLabel(
            "⚠️ Let op: open TLC_1.xlsx niet in Excel tijdens verwerken."
        )
        self.lbl_tlc_warning.setWordWrap(True)
        self.lbl_tlc_warning.setStyleSheet("""
            QLabel {
                background: rgba(255, 170, 0, 0.14);
                border: 1px solid rgba(255, 170, 0, 0.35);
                border-radius: 8px;
                padding: 10px;
                color: #b06a00;
                font-weight: bold;
            }
        """)
        root.addWidget(self.lbl_tlc_warning)

        self.lbl_cms = QLabel("0 CMS 1322 bestand(en) toegevoegd")
        self.lbl_cms.setWordWrap(True)
        self.lbl_cms.setStyleSheet("""
            QLabel {
                background: palette(base);
                border: 1px solid palette(mid);
                border-radius: 8px;
                padding: 10px;
            }
        """)
        root.addWidget(self.lbl_cms)

        # =========================
        # ACTIES BOVEN
        # =========================
        row_actions = QHBoxLayout()

        btn_tlc = QPushButton("Open TLC map")
        btn_tlc.setObjectName("secondary")
        btn_tlc.clicked.connect(self.on_upload_tlc)

        btn_cms = QPushButton("Add CMS 1322 orders")
        btn_cms.setObjectName("primary")
        btn_cms.clicked.connect(self.on_add_cms)

        btn_reset = QPushButton("Reset alles")
        btn_reset.setObjectName("secondary")
        btn_reset.clicked.connect(self.on_reset)

        row_actions.addWidget(btn_tlc)
        row_actions.addWidget(btn_cms)
        row_actions.addWidget(btn_reset)
        row_actions.addStretch()

        root.addLayout(row_actions)

        # =========================
        # RUN ACTIES
        # =========================
        row_run = QHBoxLayout()

        btn_run = QPushButton("Start 1322 verwerking")
        btn_run.setObjectName("primary")
        btn_run.setMinimumHeight(40)
        btn_run.clicked.connect(self.on_run)

        btn_open = QPushButton("Open output folder")
        btn_open.setObjectName("secondary")
        btn_open.clicked.connect(self.on_open_output)

        row_run.addWidget(btn_run)
        row_run.addWidget(btn_open)
        row_run.addStretch()

        root.addLayout(row_run)

        # =========================
        # LOG
        # =========================
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(180)
        self.log.setPlaceholderText("Log verschijnt hier...")
        root.addWidget(self.log, stretch=1)

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

        
        self.on_reset()

    def on_reset(self):
        self.engine.clear()
        self.log.clear()
        self._sync_labels()

    def on_open_output(self):
        folder = output_root() / "1322"
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))
