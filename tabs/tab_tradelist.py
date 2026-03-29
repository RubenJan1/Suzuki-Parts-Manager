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
        layout = QVBoxLayout(self)

        title = QLabel("Suzuki Parts – Tradelist Maker")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: bold;")
        layout.addWidget(title)

        uitleg = QLabel(
            "Genereer tradelists voor uw leveranciers op basis van de centrale "
            "WooCommerce-export.\n\n"
            "De gegenereerde bestanden worden opgeslagen in de map 'output/tradelists'."
        )
        uitleg.setWordWrap(True)
        uitleg.setAlignment(Qt.AlignCenter)
        layout.addWidget(uitleg)

        self.lbl_wc = QLabel("")
        self.lbl_wc.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lbl_wc)
        self._update_wc_label()
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        # Buttons (rechtsonder)
        btn_open_output = QPushButton("Open output folder")
        btn_open_output.setFixedHeight(40)
        btn_open_output.clicked.connect(self.open_output_folder)

        btn_run = QPushButton("Genereer Tradelist")
        btn_run.setFixedHeight(40)
        btn_run.clicked.connect(self.run_tradelist)

        bottom = QHBoxLayout()
        bottom.addStretch()
        bottom.addWidget(btn_open_output)
        bottom.addWidget(btn_run)
        layout.addLayout(bottom)
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
        output_dir = Path("output") / "tradelist"

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