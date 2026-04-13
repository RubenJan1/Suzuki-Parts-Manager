from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem,
    QLineEdit, QTextEdit, QDoubleSpinBox, QComboBox, QRadioButton, QButtonGroup
)
import os
import pandas as pd
from engines.engine_factuurmaker import FactuurMakerEngine
from utils.paths import output_root
from utils.theme import apply_theme


def currency(value):
    try:
        return f"€ {float(value):,.2f}".replace(".", ",")
    except:
        return "€ 0,00"


# =====================================================
# TAB FACTUURMAKER
# =====================================================

class TabFactuurmaker(QWidget):
    """
    Factuurmaker – Desktop (dark UI)
    - Meerdere CMS-bestellingen combineren
    - Supplier Number: 1322 / 277
    """

    def __init__(self, app_state):
        super().__init__()
        self.app_state = app_state
        self.engine = FactuurMakerEngine()
        self._build_ui()

    def _build_ui(self):
        # Gebruik systeemtheme (dark/light) automatisch
        apply_theme(self)
        main = QVBoxLayout(self)
        main.setSpacing(18)
        main.setContentsMargins(16, 16, 16, 16)

        # =====================================================
        # STAP 1 – CMS ORDERS
        # =====================================================
        lbl_step1 = QLabel("① CMS orders")
        lbl_step1.setStyleSheet("font-size: 14pt; font-weight: bold;")
        main.addWidget(lbl_step1)

        row_files = QHBoxLayout()

        btn_add = QPushButton("Add CMS orders (.xlsx)")
        btn_add.setObjectName("primary")
        btn_add.setFixedHeight(40)
        btn_add.clicked.connect(self.on_add_files)

        btn_reset = QPushButton("Reset")
        btn_reset.setObjectName("danger")
        btn_reset.clicked.connect(self.on_reset)

        self.lbl_status = QLabel("No CMS orders loaded")
        self.lbl_status.setStyleSheet("color: #9CA3AF;")

        row_files.addWidget(btn_add)
        row_files.addWidget(btn_reset)
        row_files.addStretch()
        row_files.addWidget(self.lbl_status)

        main.addLayout(row_files)

        # =====================================================
        # STAP 2 – FACTUURGEGEVENS
        # =====================================================
        lbl_step2 = QLabel("② Invoice details")
        lbl_step2.setStyleSheet("font-size: 14pt; font-weight: bold;")
        main.addWidget(lbl_step2)

        form = QGridLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)

        self.txt_invoice = QLineEdit(self.engine.invoice_number)
        self.txt_supplier = QLineEdit()
        self.spin_shipping = QDoubleSpinBox()
        self.spin_shipping.setDecimals(2)
        self.spin_shipping.setMaximum(9999)
        self.spin_shipping.setSuffix(" €")
        self.spin_shipping.valueChanged.connect(self._update_engine)

        self.txt_billto = QLineEdit("CMS")
        self.txt_address = QTextEdit(
            "Artemisweg 245\n8239 DD Lelystad\nNetherlands"
        )
        self.txt_address.setFixedHeight(70)

        # Document type (Invoice / Credit)
        self.rb_invoice = QRadioButton("Invoice")
        self.rb_credit = QRadioButton("Credit invoice")
        self.rb_invoice.setChecked(True)

        self.doc_group = QButtonGroup(self)
        self.doc_group.addButton(self.rb_invoice)
        self.doc_group.addButton(self.rb_credit)

        self.txt_original_invoice = QLineEdit()
        self.txt_original_invoice.setPlaceholderText("Original invoice number (required for credit)")

        self.cmb_credit_reason = QComboBox()
        self.cmb_credit_reason.addItems([
            "Return",
            "Too much invoiced",
            "Wrong item",
            "Not delivered",
            "Other",
        ])

        # Toggle visibility
        self.rb_invoice.toggled.connect(self.on_doc_type_changed)
        self.rb_credit.toggled.connect(self.on_doc_type_changed)

        form.addWidget(QLabel("Invoice number"), 0, 0)
        form.addWidget(self.txt_invoice, 0, 1)

        form.addWidget(QLabel("Supplier number"), 0, 2)
        form.addWidget(self.txt_supplier, 0, 3)

        form.addWidget(QLabel("Shipping/Handling"), 1, 0)
        form.addWidget(self.spin_shipping, 1, 1)

        form.addWidget(QLabel("Bill to"), 1, 2)
        form.addWidget(self.txt_billto, 1, 3)

        form.addWidget(QLabel("Billing address"), 2, 0)
        form.addWidget(self.txt_address, 2, 1, 1, 3)


        # Document type row
        doc_row = QHBoxLayout()
        doc_row.addWidget(self.rb_invoice)
        doc_row.addWidget(self.rb_credit)
        doc_row.addStretch()

        form.addWidget(QLabel("Document type"), 3, 0)
        form.addLayout(doc_row, 3, 1, 1, 3)

        form.addWidget(QLabel("Original invoice"), 4, 0)
        form.addWidget(self.txt_original_invoice, 4, 1)

        form.addWidget(QLabel("Credit reason"), 4, 2)
        form.addWidget(self.cmb_credit_reason, 4, 3)
 
        self.on_doc_type_changed()

        main.addLayout(form)

        # =====================================================
        # STAP 3 – PREVIEW
        # =====================================================
        lbl_step3 = QLabel("③ Invoice preview")
        lbl_step3.setStyleSheet("font-size: 14pt; font-weight: bold;")
        main.addWidget(lbl_step3)

        # Search (filter op artikelnummer)
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Search part number...")
        self.txt_search.textChanged.connect(self.on_search_changed)
        main.addWidget(self.txt_search)

        self.table = QTableWidget(0, 5)
        self.table.setEditTriggers(
            QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed
        )
        self.table.itemChanged.connect(self.on_table_edited)

        self.table.setHorizontalHeaderLabels(
            ["Part Number", "Description", "Qty", "Price", "Total"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)

        main.addWidget(self.table, stretch=1)
        row_edit = QHBoxLayout()

        btn_remove = QPushButton("Remove selected row")
        btn_remove.setObjectName("danger")
        btn_remove.clicked.connect(self.on_remove_row)

        btn_add = QPushButton("Add manual item")
        btn_add.setObjectName("secondary")
        btn_add.clicked.connect(self.on_add_row)

        row_edit.addWidget(btn_add)
        row_edit.addWidget(btn_remove)
        row_edit.addStretch()

        main.addLayout(row_edit)


        # =====================================================
        # STAP 4 – ACTIES
        # =====================================================
        lbl_step4 = QLabel("④ Actions")
        lbl_step4.setStyleSheet("font-size: 14pt; font-weight: bold;")
        main.addWidget(lbl_step4)

        row_actions = QHBoxLayout()

        btn_generate = QPushButton("Generate invoice (PDF)")
        btn_generate.setObjectName("primary")
        btn_generate.setFixedHeight(40)
        btn_generate.clicked.connect(self.on_generate)

        btn_open = QPushButton("Open output folder")
        btn_open.setObjectName("secondary")
        btn_open.clicked.connect(self.on_open_output)

        row_actions.addStretch()
        row_actions.addWidget(btn_generate)
        row_actions.addWidget(btn_open)

        main.addLayout(row_actions)

    # =====================================================
    # LOGICA
    # =====================================================

    def on_add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select CMS orders",
            "",
            "Excel files (*.xlsx)"
        )
        if not paths:
            return

        for p in paths:
            self.engine.add_cms_bestelling(p)

        self.refresh_preview()

    def on_reset(self):
        if QMessageBox.question(
            self,
            "Reset",
            "Remove all CMS orders?"
        ) != QMessageBox.Yes:
            return

        self.engine.clear_bestellingen()
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        self.lbl_status.setText("No CMS orders loaded")

    def refresh_preview(self):
        df = self.engine.work_df
        if df is None:
            return

        self.table.blockSignals(True)
        self.table.setRowCount(0)

        self.lbl_status.setText(
            f"{len(self.engine.cms_paths)} CMS order file(s) loaded"
        )

        for _, r in df.iterrows():
            row = self.table.rowCount()
            self.table.insertRow(row)

            qty = int(r["Aantal"])
            price = float(r["Prijs"])
            total = qty * price

            self.table.setItem(row, 0, QTableWidgetItem(str(r["Artikel"])))
            self.table.setItem(row, 1, QTableWidgetItem(str(r["Omschrijving"])))
            self.table.setItem(row, 2, QTableWidgetItem(str(qty)))
            self.table.setItem(row, 3, QTableWidgetItem(currency(price)))
            self.table.setItem(row, 4, QTableWidgetItem(currency(total)))

        self.table.blockSignals(False)
        self.on_search_changed(self.txt_search.text() if hasattr(self, "txt_search") else "")

    def _update_engine(self):
        self.engine.verzendkosten = float(self.spin_shipping.value())

    def on_generate(self):
        try:
            self.engine.invoice_number = self.txt_invoice.text().strip()
            self.engine.supplier_number = self.txt_supplier.text().strip()
            self.engine.bill_to = self.txt_billto.text().strip()
            self.engine.billing_address = self.txt_address.toPlainText()
            self.engine.logo_path = "assets/logo.png"

            # Document type
            self.engine.document_type = "credit" if self.rb_credit.isChecked() else "invoice"
            self.engine.original_invoice_number = self.txt_original_invoice.text().strip()
            self.engine.credit_reason = self.cmb_credit_reason.currentText().strip()

            if self.engine.document_type == "credit" and not self.engine.original_invoice_number:
                QMessageBox.warning(
                    self,
                    "Missing info",
                    "For a credit invoice, 'Original invoice number' is required."
                )
                return

            path = self.engine.generate_pdf()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return

        QMessageBox.information(
            self,
            "Invoice created",
            f"Invoice successfully generated:\n{path}"
        )

    def on_open_output(self):
        folder = output_root() / "facturen"
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))


    def on_search_changed(self, text: str):
        q = (text or "").strip().lower()
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            part = item.text().lower() if item else ""
            hide = bool(q) and (q not in part)
            self.table.setRowHidden(r, hide)

    def on_table_edited(self, item):
        row = item.row()
        col = item.column()

        mapping = {
            0: "Artikel",
            1: "Omschrijving",
            2: "Aantal",
            3: "Prijs",
        }

        if col not in mapping:
            return

        field = mapping[col]
        value = item.text().replace("€", "").replace(",", ".").strip()

        try:
            if field == "Aantal":
                value = int(value)
            elif field == "Prijs":
                value = float(value)
        except:
            return

        self.engine.work_df.at[row, field] = value
        # Na elke wijziging: dubbelen samenvoegen + sorteren + totals opnieuw berekenen
        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()

        # Na aanpassen artikelnummer: her-sorteer zodat lijst logisch blijft
        if field == "Artikel":
            self.engine.sort_work_df()
            self.refresh_preview()

    def on_remove_row(self):
        row = self.table.currentRow()
        if row < 0:
            return

        if QMessageBox.question(
            self,
            "Remove item",
            "Remove selected item from invoice?"
        ) != QMessageBox.Yes:
            return

        self.engine.work_df = self.engine.work_df.drop(
            self.engine.work_df.index[row]
        ).reset_index(drop=True)
        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()

    
    def on_add_row(self):
        new_row = {
            "Artikel": "MANUAL",
            "Omschrijving": "Handmatig toegevoegd",
            "Aantal": 1,
            "Prijs": 0.0,
        }

        self.engine.work_df = pd.concat(
            [self.engine.work_df, pd.DataFrame([new_row])],
            ignore_index=True
        )
        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()

    def on_doc_type_changed(self):
        is_credit = self.rb_credit.isChecked()
        self.txt_original_invoice.setVisible(is_credit)
        self.cmb_credit_reason.setVisible(is_credit)
        

