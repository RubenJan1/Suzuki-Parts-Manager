# tabs/tab_intro.py
import pandas as pd
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel,
    QPushButton, QFileDialog, QMessageBox
)

class TabIntro(QWidget):
    def __init__(self, app_state, on_loaded_callback=None):
        super().__init__()
        self.app_state = app_state
        self.on_loaded_callback = on_loaded_callback

        layout = QVBoxLayout()

        title = QLabel("Welkom – Suzuki Parts Tools")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")

        uitleg = QLabel(
            "Upload hier de WooCommerce product-export (CSV).\n\n"
            "Dit bestand is een momentopname van de website en wordt\n"
            "gebruikt door Inboeken, Website 277 en TLC 1322."
        )
        uitleg.setWordWrap(True)

        btn_upload = QPushButton("Upload WooCommerce CSV export")
        btn_upload.clicked.connect(self.load_wc_export)

        self.status = QLabel("Nog geen WooCommerce export geladen")

        layout.addWidget(title)
        layout.addWidget(uitleg)
        layout.addSpacing(20)
        layout.addWidget(btn_upload)
        layout.addWidget(self.status)
        layout.addStretch()

        self.setLayout(layout)

    def load_wc_export(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecteer WooCommerce product-export",
            "",
            "CSV bestanden (*.csv)"
        )

        if not path:
            return

        try:
            df = pd.read_csv(path, dtype=str, low_memory=False)
        except Exception as e:
            QMessageBox.critical(
                self,
                "Fout",
                f"Kan WooCommerce CSV niet laden:\n{e}"
            )
            return

        self.app_state.set_wc_export(df, path)
        self.status.setText(f"Geladen: {path}")

        QMessageBox.information(
            self,
            "Succes",
            "WooCommerce export succesvol geladen.\n"
            "Alle tabs kunnen nu gebruikt worden."
        )

        if self.on_loaded_callback:
            self.on_loaded_callback()
