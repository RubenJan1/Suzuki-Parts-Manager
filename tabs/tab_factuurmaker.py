from __future__ import annotations

import os
import pandas as pd

from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem,
    QLineEdit, QTextEdit, QDoubleSpinBox, QComboBox, QRadioButton, QButtonGroup,
    QSplitter, QSizePolicy, QHeaderView
)

from engines.engine_factuurmaker import FactuurMakerEngine
from utils.paths import output_root
from utils.theme import apply_theme

def currency(value) -> str:
    """Format euro with European separators."""
    try:
        v = float(value)
        return f"€ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "€ 0,00"



class TabFactuurmaker(QWidget):
    """
    Factuurmaker (Factuur + Creditfactuur in één tab)

    PRO v3 UI fixes:
    - Preview krijgt standaard veel meer ruimte op 14–15,6"
    - Type toggle + Actions in één rij (minder “lege lucht”)
    - Slimme kolombreedtes (minder horizontale scrollbar)
    - Splitter handle breder (makkelijk slepen)
    """

    def __init__(self, app_state):
        super().__init__()
        self.app_state = app_state
        self.engine = FactuurMakerEngine()
        self._build_ui()

    def _build_ui(self):
        apply_theme(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # =========================
        # TITEL
        # =========================
        title = QLabel("Factuurmaker")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        root.addWidget(title)

        # =========================
        # UITLEG
        # =========================
        uitleg = QLabel(
            "Maak facturen of creditfacturen op basis van CMS-bestellingen.\n\n"
            "① Voeg één of meer CMS-bestanden toe\n"
            "② Controleer documentgegevens\n"
            "③ Controleer preview\n"
            "④ Genereer PDF"
        )
        uitleg.setWordWrap(True)
        root.addWidget(uitleg)

        # =========================
        # HOOFDWERKGEBIED
        # =========================
        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(10)
        root.addWidget(splitter, stretch=1)

        # =========================
        # TOP: BESTANDEN + GEGEVENS + ACTIES
        # =========================
        top = QWidget()
        top_l = QVBoxLayout(top)
        top_l.setSpacing(10)
        top_l.setContentsMargins(0, 0, 0, 0)

        # CMS bestanden
        lbl_step1 = QLabel("CMS orders")
        lbl_step1.setStyleSheet("font-size: 14px; font-weight: bold;")
        top_l.addWidget(lbl_step1)

        row_files = QHBoxLayout()

        btn_add = QPushButton("Add CMS orders (.xlsx)")
        btn_add.setObjectName("primary")
        btn_add.clicked.connect(self.on_add_files)

        btn_reset = QPushButton("Reset")
        btn_reset.setObjectName("danger")
        btn_reset.clicked.connect(self.on_reset)

        self.lbl_status = QLabel("No CMS orders loaded")

        row_files.addWidget(btn_add)
        row_files.addWidget(btn_reset)
        row_files.addStretch()
        row_files.addWidget(self.lbl_status)

        top_l.addLayout(row_files)

        # Documentgegevens
        lbl_step2 = QLabel("Document details")
        lbl_step2.setStyleSheet("font-size: 14px; font-weight: bold;")
        top_l.addWidget(lbl_step2)

        form = QGridLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)

        self.txt_invoice = QLineEdit(self.engine.invoice_number)
        self.txt_invoice.setPlaceholderText("Leeg laten voor automatisch nummer")

        self.txt_supplier = QLineEdit()
        self.txt_supplier.setPlaceholderText("Optioneel (CMS ref / klantnummer)")

        self.spin_shipping = QDoubleSpinBox()
        self.spin_shipping.setDecimals(2)
        self.spin_shipping.setMaximum(9999)
        self.spin_shipping.setSuffix(" €")
        self.spin_shipping.valueChanged.connect(self._update_engine)

        self.txt_billto = QLineEdit(self.engine.bill_to)

        self.txt_address = QTextEdit()
        self.txt_address.setPlainText(self.engine.billing_address)
        self.txt_address.setFixedHeight(72)

        form.addWidget(QLabel("Documentnummer"), 0, 0)
        form.addWidget(self.txt_invoice, 0, 1)

        form.addWidget(QLabel("Klant-/CMS nummer"), 0, 2)
        form.addWidget(self.txt_supplier, 0, 3)

        form.addWidget(QLabel("Verzend/handling"), 1, 0)
        form.addWidget(self.spin_shipping, 1, 1)

        form.addWidget(QLabel("Bill to"), 1, 2)
        form.addWidget(self.txt_billto, 1, 3)

        form.addWidget(QLabel("Billing address"), 2, 0)
        form.addWidget(self.txt_address, 2, 1, 1, 3)

        top_l.addLayout(form)

        # Type + acties
        lbl_step3 = QLabel("Type en acties")
        lbl_step3.setStyleSheet("font-size: 14px; font-weight: bold;")
        top_l.addWidget(lbl_step3)

        row_type_actions = QHBoxLayout()
        row_type_actions.setSpacing(10)

        self.rb_invoice = QRadioButton("Factuur")
        self.rb_credit = QRadioButton("Creditfactuur")
        self.rb_invoice.setChecked(True)

        self.doc_group = QButtonGroup(self)
        self.doc_group.addButton(self.rb_invoice)
        self.doc_group.addButton(self.rb_credit)
        self.rb_invoice.toggled.connect(self.on_doc_type_changed)
        self.rb_credit.toggled.connect(self.on_doc_type_changed)

        self.lbl_original_invoice = QLabel("Originele factuur")
        self.txt_original_invoice = QLineEdit()
        self.txt_original_invoice.setPlaceholderText("Factuurnummer waar deze credit bij hoort (verplicht)")

        self.lbl_credit_reason = QLabel("Reden")
        self.cmb_credit_reason = QComboBox()
        self.cmb_credit_reason.addItems(["Retour", "Te veel gefactureerd", "Verkeerd artikel", "Niet geleverd", "Overig"])

        left = QVBoxLayout()
        left.setSpacing(6)

        row_type = QHBoxLayout()
        row_type.addWidget(QLabel("Type"))
        row_type.addSpacing(8)
        row_type.addWidget(self.rb_invoice)
        row_type.addWidget(self.rb_credit)
        row_type.addStretch()
        left.addLayout(row_type)

        row_credit = QGridLayout()
        row_credit.setHorizontalSpacing(12)
        row_credit.setVerticalSpacing(6)
        row_credit.addWidget(self.lbl_original_invoice, 0, 0)
        row_credit.addWidget(self.txt_original_invoice, 0, 1)
        row_credit.addWidget(self.lbl_credit_reason, 0, 2)
        row_credit.addWidget(self.cmb_credit_reason, 0, 3)
        left.addLayout(row_credit)

        left_wrap = QWidget()
        left_wrap.setLayout(left)

        actions = QHBoxLayout()
        actions.setSpacing(10)

        self.btn_generate = QPushButton("Generate invoice (PDF)")
        self.btn_generate.setObjectName("primary")
        self.btn_generate.clicked.connect(self.on_generate)

        btn_open = QPushButton("Open output folder")
        btn_open.setObjectName("secondary")
        btn_open.clicked.connect(self.on_open_output)

        actions.addWidget(self.btn_generate)
        actions.addWidget(btn_open)

        actions_wrap = QWidget()
        actions_wrap.setLayout(actions)

        row_type_actions.addWidget(left_wrap, 1)
        row_type_actions.addStretch()
        row_type_actions.addWidget(actions_wrap, 0)

        top_l.addLayout(row_type_actions)

        splitter.addWidget(top)

        # =========================
        # BOTTOM: PREVIEW
        # =========================
        bottom = QWidget()
        bottom_l = QVBoxLayout(bottom)
        bottom_l.setSpacing(8)
        bottom_l.setContentsMargins(0, 0, 0, 0)

        lbl_preview = QLabel("Preview")
        lbl_preview.setStyleSheet("font-size: 14px; font-weight: bold;")
        bottom_l.addWidget(lbl_preview)

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Search part number...")
        self.txt_search.textChanged.connect(self.on_search_changed)
        bottom_l.addWidget(self.txt_search)

        self.table = QTableWidget(0, 5)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.table.itemChanged.connect(self.on_table_edited)
        self.table.setHorizontalHeaderLabels(["Part Number", "Description", "Qty", "Price", "Total"])
        self.table.setAlternatingRowColors(True)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.table.setVerticalScrollMode(QTableWidget.ScrollPerPixel)
        self.table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        self.table.verticalHeader().setDefaultSectionSize(28)

        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        bottom_l.addWidget(self.table, stretch=1)

        row_edit = QHBoxLayout()

        btn_add_row = QPushButton("Add manual item")
        btn_add_row.setObjectName("secondary")
        btn_add_row.clicked.connect(self.on_add_row)

        btn_remove = QPushButton("Remove selected row")
        btn_remove.setObjectName("danger")
        btn_remove.clicked.connect(self.on_remove_row)

        row_edit.addWidget(btn_add_row)
        row_edit.addWidget(btn_remove)
        row_edit.addStretch()

        bottom_l.addLayout(row_edit)

        splitter.addWidget(bottom)

        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([260, 740])

        self.on_doc_type_changed()
    # =====================================================
    # LOGIC
    # =====================================================

    def _update_engine(self):
        self.engine.verzendkosten = float(self.spin_shipping.value())

    def on_doc_type_changed(self):
        is_credit = self.rb_credit.isChecked()

        self.lbl_original_invoice.setVisible(is_credit)
        self.txt_original_invoice.setVisible(is_credit)
        self.lbl_credit_reason.setVisible(is_credit)
        self.cmb_credit_reason.setVisible(is_credit)

        self.btn_generate.setText("Generate credit invoice (PDF)" if is_credit else "Generate invoice (PDF)")
        self.refresh_preview()

    def on_add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select CMS orders", "", "Excel files (*.xlsx)")
        if not paths:
            return
        for p in paths:
            self.engine.add_cms_bestelling(p)
        self.refresh_preview()

    def on_reset(self):
        if QMessageBox.question(self, "Reset", "Remove all CMS orders?") != QMessageBox.Yes:
            return
        self.engine.clear_bestellingen()
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        self.table.blockSignals(False)
        self.lbl_status.setText("No CMS orders loaded")

    def refresh_preview(self):
        df = self.engine.work_df
        if df is None:
            return

        is_credit = self.rb_credit.isChecked()
        sign = -1 if is_credit else 1

        self.table.blockSignals(True)
        self.table.setRowCount(0)
        self.lbl_status.setText(f"{len(self.engine.cms_paths)} CMS order file(s) loaded")

        for _, r in df.iterrows():
            row = self.table.rowCount()
            self.table.insertRow(row)

            qty = int(r["Aantal"]) * sign
            price = float(r["Prijs"])
            total = qty * price

            self.table.setItem(row, 0, QTableWidgetItem(str(r["Artikel"])))
            self.table.setItem(row, 1, QTableWidgetItem(str(r["Omschrijving"])))
            self.table.setItem(row, 2, QTableWidgetItem(str(qty)))
            self.table.setItem(row, 3, QTableWidgetItem(currency(price)))
            self.table.setItem(row, 4, QTableWidgetItem(currency(total)))

        self.table.blockSignals(False)
        self.on_search_changed(self.txt_search.text())

    def _parse_float_eu(self, s: str) -> float:
        s = (s or "").strip().replace("€", "").strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        return float(s)

    def on_generate(self):
        try:
            bill_to = (self.txt_billto.text() or "").strip()
            addr = (self.txt_address.toPlainText() or "").strip()
            if not bill_to:
                QMessageBox.warning(self, "Missing info", "Vul 'Bill to' in.")
                return
            if not addr:
                QMessageBox.warning(self, "Missing info", "Vul 'Billing address' in.")
                return

            is_credit = self.rb_credit.isChecked()

            inv_no = (self.txt_invoice.text() or "").strip()
            self.engine.invoice_number = inv_no  # engine may auto-number if blank

            self.engine.supplier_number = (self.txt_supplier.text() or "").strip()
            self.engine.bill_to = bill_to
            self.engine.billing_address = addr
            self.engine.logo_path = "assets/logo.png"
            self._update_engine()

            self.engine.document_type = "credit" if is_credit else "invoice"
            self.engine.original_invoice_number = (self.txt_original_invoice.text() or "").strip()
            self.engine.credit_reason = (self.cmb_credit_reason.currentText() or "").strip()

            if self.engine.document_type == "credit" and not self.engine.original_invoice_number:
                QMessageBox.warning(self, "Missing info", "Voor een creditfactuur is 'Originele factuur' verplicht.")
                return

            path = self.engine.generate_pdf()
            self.txt_invoice.setText(self.engine.invoice_number)

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return

        QMessageBox.information(self, "Document created", f"PDF generated:\n{path}")

    def on_open_output(self):
        folder = output_root() / "facturen"
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))

    def on_search_changed(self, text: str):
        q = (text or "").strip().lower()
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            part = item.text().lower() if item else ""
            self.table.setRowHidden(r, bool(q) and (q not in part))

    def on_table_edited(self, item):
        row = item.row()
        col = item.column()
        mapping = {0: "Artikel", 1: "Omschrijving", 2: "Aantal", 3: "Prijs"}
        if col not in mapping:
            return

        field = mapping[col]
        try:
            if field == "Aantal":
                v = int(float(item.text().strip()))
                self.engine.work_df.at[row, field] = abs(v)
            elif field == "Prijs":
                self.engine.work_df.at[row, field] = self._parse_float_eu(item.text())
            else:
                self.engine.work_df.at[row, field] = item.text()
        except Exception:
            return

        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()

    def on_remove_row(self):
        row = self.table.currentRow()
        if row < 0:
            return
        if QMessageBox.question(self, "Remove item", "Remove selected row from document?") != QMessageBox.Yes:
            return
        self.engine.work_df = self.engine.work_df.drop(self.engine.work_df.index[row]).reset_index(drop=True)
        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()

    def on_add_row(self):
        if self.engine.work_df is None:
            self.engine.clear_bestellingen()
        new_row = {"Artikel": "MANUAL", "Omschrijving": "Handmatig toegevoegd", "Aantal": 1, "Prijs": 0.0}
        self.engine.work_df = pd.concat([self.engine.work_df, pd.DataFrame([new_row])], ignore_index=True)
        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()
