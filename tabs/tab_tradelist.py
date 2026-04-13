# ============================================================
# tab_tradelist.py
# UI-tab voor Tradelist Maker
# Gebruikt ALTIJD de centrale WC-export uit AppState
# (geen eigen CSV-upload meer)
# ============================================================

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton,
    QMessageBox, QTextEdit
)
from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from pathlib import Path


from engines.engine_tradelist import TradelistEngine
from utils.paths import output_root
from utils.theme import apply_theme

class TabTradelist(QWidget):
    def __init__(self, app_state):
        super().__init__()
        self.app_state = app_state
        self.engine = TradelistEngine(app_state)

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
        title = QLabel("Tradelist Maker")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        root.addWidget(title)

        # =========================
        # UITLEG
        # =========================
        uitleg = QLabel(
            "Genereer tradelists voor leveranciers op basis van de geladen "
            "WooCommerce export.\n\n"
            "① Controleer of de juiste WC export geladen is\n"
            "② Start genereren\n"
            "③ Open de output map voor de bestanden"
        )
        uitleg.setWordWrap(True)
        root.addWidget(uitleg)

        # =========================
        # STATUS BLOK
        # =========================
        self.lbl_wc = QLabel("")
        self.lbl_wc.setWordWrap(True)
        self.lbl_wc.setStyleSheet("""
            QLabel {
                background: palette(base);
                border: 1px solid palette(mid);
                border-radius: 8px;
                padding: 10px;
            }
        """)
        root.addWidget(self.lbl_wc)
        self._update_wc_label()

        # =========================
        # ACTIES
        # =========================
        row_actions = QHBoxLayout()

        btn_run = QPushButton("Genereer tradelist")
        btn_run.setObjectName("primary")
        btn_run.setMinimumHeight(40)
        btn_run.clicked.connect(self.run_tradelist)

        btn_open_output = QPushButton("Open output folder")
        btn_open_output.setObjectName("secondary")
        btn_open_output.clicked.connect(self.open_output_folder)

        row_actions.addWidget(btn_run)
        row_actions.addWidget(btn_open_output)
        row_actions.addStretch()

        root.addLayout(row_actions)

        # =========================
        # LOG
        # =========================
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Log verschijnt hier...")
        root.addWidget(self.log, stretch=1)
    # --------------------------------------------------------
    # Helpers
    # --------------------------------------------------------
    def _update_wc_label(self):
        if self.app_state.wc_path:
            self.lbl_wc.setText(f"WC-export in gebruik: {self.app_state.wc_path}")
        else:
            self.lbl_wc.setText("❌ Geen WC-export geladen")


    def open_output_folder(self):
        """Open de tradelist output map in de file explorer."""
        # Engine default is output/tradelist
        output_dir = output_root() / "tradelist"

        # Als we al resultaten hebben in de log, probeer laatste pad te pakken
        # (niet verplicht, maar handig als je ooit met een andere output_dir draait)
        last_path = getattr(self, "_last_output_path", None)
        if last_path:
            try:
                p = Path(last_path)
                if p.exists():
                    output_dir = p.parent
            except Exception:
                pass

        output_dir = output_dir.resolve()
        if not output_dir.exists():
            QMessageBox.information(
                self,
                "Nog geen output",
                f"De output map bestaat nog niet:\n{output_dir}\n\nGenereer eerst een tradelist."
            )
            return

        QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_dir)))

    # --------------------------------------------------------
    # Actie
    # --------------------------------------------------------
    def run_tradelist(self):
        if self.app_state.wc_df is None:
            QMessageBox.warning(
                self,
                "Geen WC-export",
                "Er is geen WooCommerce-export geladen. Ga eerst naar de Start-tab."
            )
            return

        self.log.append("Start tradelist maker...")

        try:
            results = self.engine.run()
        except Exception as e:
            QMessageBox.critical(self, "Fout", str(e))
            self.log.append(f"FOUT: {e}")
            return

        self.log.append("\nKlaar! Gegenereerde bestanden:")
        for name, path in results.items():
            self.log.append(f"- {name}: {path}")


        # onthoud laatste output bestand (voor Open output folder)
     
        if isinstance(results, dict):
            # pak bij voorkeur de ONSZELF tradelist
            self._last_output_path = str(results.get('onszelf'))



        QMessageBox.information(
            self,
            "Tradelist klaar",
            "De tradelists zijn succesvol gegenereerd."
        )