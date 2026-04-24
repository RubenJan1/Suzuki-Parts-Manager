from __future__ import annotations

import json
import os
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Dict

import pandas as pd
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from utils.paths import output_root, resource_path

# =====================================================
# CONSTANTS
# =====================================================

VAT_PERCENT = 21.0

COMPANY_INFO = [
    "Vlaandere Motoren - de Marne 136 B",
    "8701 MC - Bolsward",
    "Tel: +316-41484547",
    "IBAN: NL49 RABO 0372 0041 64",
    "VAT: 8077 51 911 B01 | C.O.C: 01018576",
]

FOOTER_INVOICE = [
    "Payment term: 14 days.",
    "Returns only allowed when returned within 30 days of the billing date.",
    "Vlaandere Motoren — IBAN: NL49 RABO 0372 0041 64",
    "VATnumber 8077 51 911 B01 | C.O.C.number 01018576",
]

FOOTER_CREDIT = [
    "This credit note corrects the referenced original invoice.",
    "Settlement will be handled against the original invoice unless a refund is agreed.",
    "Vlaandere Motoren — IBAN: NL49 RABO 0372 0041 64",
    "VATnumber 8077 51 911 B01 | C.O.C.number 01018576",
]


class FactuurMakerEngine:
    """
    Engine voor CMS facturen/creditfacturen:
    - combineert meerdere CMS-bestellingen
    - sorteert Suzuki artikelnummer logisch
    - merge duplicates (Artikel + Prijs) -> Aantal optellen
    - genereert PDF (factuur of creditfactuur)

    Professionele aanvullingen:
    - doorlopende nummering (per jaar + per type)
    - validaties (lege klantgegevens, totaal=0)
    - audit log (JSONL)
    """

    def __init__(self):
        self.cms_paths: list[str] = []
        self.work_df: pd.DataFrame | None = None
        self.verzendkosten: float = 0.0

        # If user leaves blank, engine will auto-number via sequence store.
        self.invoice_number: str = f"INV-{datetime.now().strftime('%Y%m%d%H%M')}"
        self.supplier_number: str = ""
        self.bill_to: str = "CMS"
        self.billing_address: str = "Artemisweg 245\n8239 DD Lelystad\nNetherlands"

        # Document type
        self.document_type: str = "invoice"  # "invoice" | "credit"
        self.original_invoice_number: str = ""
        self.credit_reason: str = ""

        self.logo_path: str = str(resource_path("assets/logo.png"))

        # Storage paths
        self.output_dir = output_root() / "facturen"
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.sequence_path = self.output_dir / "sequence.json"
        self.audit_log_path = self.output_dir / "facturen_log.jsonl"

    # -------------------------------------------------
    # DATA
    # -------------------------------------------------

    def add_cms_bestelling(self, path: str):
        if path not in self.cms_paths:
            self.cms_paths.append(path)
            self._rebuild_work_df()

    def clear_bestellingen(self):
        self.cms_paths.clear()
        self.work_df = pd.DataFrame(columns=["Artikel", "Omschrijving", "Aantal", "Prijs"])

    def _sort_key(self, artikel):
        """Suzuki-onderdelen correct sorteren: elk numeriek segment telt mee."""
        try:
            parts = str(artikel).strip().split("-")
            nums = []
            for p in parts:
                digits = "".join(ch for ch in p if ch.isdigit())
                nums.append(int(digits) if digits else 0)
            return tuple(nums)
        except Exception:
            return (0,)

    def _rebuild_work_df(self):
        """Read all CMS order files and build one combined, globally-sorted list."""
        dfs = []
        for p in self.cms_paths:
            df = pd.read_excel(
                p,
                header=None,
                usecols=[0, 1, 2, 3],
                names=["Artikel", "Omschrijving", "Aantal", "Prijs"],
            )
            dfs.append(df)

        if not dfs:
            self.work_df = pd.DataFrame(columns=["Artikel", "Omschrijving", "Aantal", "Prijs"])
            return

        merged = pd.concat(dfs, ignore_index=True)

        merged["Artikel"] = merged["Artikel"].astype(str).str.strip()
        merged["Omschrijving"] = merged["Omschrijving"].astype(str)
        merged["Aantal"] = pd.to_numeric(merged["Aantal"], errors="coerce").fillna(0).astype(int)
        merged["Prijs"] = pd.to_numeric(merged["Prijs"], errors="coerce").fillna(0.0).astype(float)

        merged = self._merge_duplicates_df(merged)

        merged["_sort"] = merged["Artikel"].apply(self._sort_key)
        merged = merged.sort_values("_sort", kind="mergesort").drop(columns=["_sort"]).reset_index(drop=True)

        self.work_df = merged

    def sort_work_df(self):
        """Sort current working list again by article number."""
        if self.work_df is None or self.work_df.empty:
            return
        df = self.work_df.copy()
        df["_sort"] = df["Artikel"].apply(self._sort_key)
        self.work_df = df.sort_values("_sort", kind="mergesort").drop(columns=["_sort"]).reset_index(drop=True)

    def merge_work_df(self):
        """Merge duplicate lines on (Artikel + Prijs)."""
        if self.work_df is None or self.work_df.empty:
            return
        self.work_df = self._merge_duplicates_df(self.work_df)

    def _merge_duplicates_df(self, df: pd.DataFrame) -> pd.DataFrame:
        for col in ["Artikel", "Omschrijving", "Aantal", "Prijs"]:
            if col not in df.columns:
                df[col] = "" if col in ["Artikel", "Omschrijving"] else 0

        tmp = df.copy()
        tmp["Artikel"] = tmp["Artikel"].astype(str).str.strip()
        tmp["Omschrijving"] = tmp["Omschrijving"].fillna("").astype(str)
        tmp["Aantal"] = pd.to_numeric(tmp["Aantal"], errors="coerce").fillna(0).astype(int)
        tmp["Prijs"] = pd.to_numeric(tmp["Prijs"], errors="coerce").fillna(0.0).astype(float)

        def pick_omschrijving(series: pd.Series) -> str:
            for v in series.astype(str):
                v = (v or "").strip()
                if v and v.lower() != "nan":
                    return v
            return ""

        grouped = (
            tmp.groupby(["Artikel", "Prijs"], as_index=False, sort=False)
            .agg({"Omschrijving": pick_omschrijving, "Aantal": "sum"})
        )

        grouped = grouped[["Artikel", "Omschrijving", "Aantal", "Prijs"]]
        return grouped.reset_index(drop=True)

    def _format_address_lines(self, address_text: str):
        """
        Make address lines nice:
        - if multi-line, keep lines
        - if single line with commas, split
        - try to detect NL postcode to split lines
        """
        s = str(address_text or "").strip()
        if not s:
            return []
        s = re.sub(r"[ \t]+", " ", s)
        if "\n" in s:
            lines = [ln.strip() for ln in s.split("\n") if ln.strip()]
            return lines
        if "," in s:
            parts = [p.strip() for p in s.split(",") if p.strip()]
            return parts

        m_pc = re.search(r"\b(\d{4}\s*[A-Z]{2})\b", s.upper())
        if m_pc:
            pc = m_pc.group(1).replace(" ", "")
            before = s[:m_pc.start()].strip()
            after = s[m_pc.end():].strip()
            parts_after = [p for p in after.split(" ") if p]
            country = ""
            if parts_after:
                last = parts_after[-1].strip().title()
                if last in {"Netherlands", "Nederland", "Germany", "Deutschland", "Belgium", "Belgie", "België", "France", "Italia", "Italy", "Spain"}:
                    country = last
                    city = " ".join(parts_after[:-1]).strip()
                else:
                    city = " ".join(parts_after).strip()
            else:
                city = ""
            line1 = before
            line2 = f"{pc[:4]} {pc[4:]} {city}".strip()
            lines = [ln for ln in [line1, line2, country] if ln]
            if lines:
                return lines

        return [s]

    def merged_df(self) -> pd.DataFrame:
        if self.work_df is None:
            return pd.DataFrame(columns=["Artikel", "Omschrijving", "Aantal", "Prijs"])
        return self.work_df.copy()

    # -------------------------------------------------
    # NUMBERING + AUDIT LOG
    # -------------------------------------------------

    def _load_sequence(self) -> Dict[str, Any]:
        if not self.sequence_path.exists():
            return {}
        try:
            return json.loads(self.sequence_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_sequence(self, data: Dict[str, Any]) -> None:
        self.sequence_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    def _next_document_number(self, doc_type: str) -> str:
        """
        Generate next number in a per-year sequence.
        Example:
          INV-2026-000123
          CR-2026-000045
        """
        year = datetime.now().strftime("%Y")
        prefix = "INV" if doc_type == "invoice" else "CR"
        seq = self._load_sequence()
        key = f"{prefix}-{year}"
        n = int(seq.get(key, 0)) + 1
        seq[key] = n
        self._save_sequence(seq)
        return f"{prefix}-{year}-{n:06d}"

    def _ensure_document_number(self) -> None:
        """
        If invoice_number is empty, auto-number.
        If user already typed something, we keep it.
        """
        if str(self.invoice_number or "").strip():
            return
        self.invoice_number = self._next_document_number(self.document_type)

    def _append_audit_log(self, entry: Dict[str, Any]) -> None:
        self.audit_log_path.parent.mkdir(parents=True, exist_ok=True)
        with self.audit_log_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")

    # -------------------------------------------------
    # PDF
    # -------------------------------------------------

    def _fmt_money(self, v: float) -> str:
        return f"€ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def generate_pdf(self) -> str:
        """
        Generate invoice/credit PDF.

        Validations:
        - bill_to and billing_address must be present
        - grand total cannot be 0.00 (prevents empty/mistake docs)
        """
        # Validate required client fields (defensive)
        if not str(self.bill_to or "").strip():
            raise RuntimeError("Bill to is empty.")
        if not str(self.billing_address or "").strip():
            raise RuntimeError("Billing address is empty.")

        self.merge_work_df()
        df = self.merged_df().copy()

        df["_sort"] = df["Artikel"].apply(self._sort_key)
        df = df.sort_values("_sort", kind="mergesort").reset_index(drop=True).drop(columns=["_sort"])

        if df.empty:
            raise RuntimeError("No CMS orders loaded")

        is_credit = (getattr(self, "document_type", "invoice") == "credit")
        sign = -1 if is_credit else 1

        # Totals (for validation + later display)
        df["Total"] = (df["Aantal"] * df["Prijs"]) * sign
        subtotal = float(df["Total"].sum())
        shipping = float(self.verzendkosten) * sign
        total_ex_vat = subtotal + shipping
        vat_amount = total_ex_vat * (VAT_PERCENT / 100)
        grand_total = total_ex_vat + vat_amount

        if abs(grand_total) < 0.005:
            raise RuntimeError("Grand total is 0.00 — refusing to generate an empty document.")

        # Ensure numbering after validation (so sequence isn't consumed if validation fails)
        self._ensure_document_number()

        self.output_dir.mkdir(parents=True, exist_ok=True)

        doc_label = "CREDITFACTUUR" if is_credit else "FACTUUR"
        file_prefix = "Credit_" if is_credit else "Invoice_"
        output_path = str(self.output_dir / f"{file_prefix}{self.invoice_number}.pdf")

        buffer = BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=1.5 * cm,
            leftMargin=1.5 * cm,
            topMargin=2 * cm,
            bottomMargin=1.5 * cm,
        )

        styles = getSampleStyleSheet()
        title = ParagraphStyle("Title", parent=styles["Heading1"], fontSize=24, textColor=colors.HexColor("#111827"))
        normal = ParagraphStyle("Normal", parent=styles["Normal"], fontSize=10, textColor=colors.HexColor("#111827"))
        bold = ParagraphStyle("Bold", parent=styles["Normal"], fontSize=10, textColor=colors.HexColor("#111827"), fontName="Helvetica-Bold")

        elements = []

        # Logo (non-blocking)
        if self.logo_path and os.path.exists(self.logo_path):
            try:
                img = Image.open(self.logo_path)
                bio = BytesIO()
                img.save(bio, format="PNG")
                bio.seek(0)
                elements.append(RLImage(bio, width=12 * cm, height=5 * cm, kind="proportional"))
            except Exception:
                pass

        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"<b>{doc_label}</b>", title))
        elements.append(Spacer(1, 10))

        for line in COMPANY_INFO:
            elements.append(Paragraph(line, normal))
        elements.append(Spacer(1, 14))

        bill_block = [Paragraph(f"<b>{self.bill_to}</b>", bold)]
        for l in self._format_address_lines(self.billing_address):
            bill_block.append(Paragraph(l, normal))

        number_label = "Credit Note Number:" if is_credit else "Invoice Number:"
        date_label = "Credit Date:" if is_credit else "Invoice Date:"

        info_block = [
            ["Document:", doc_label],
            [number_label, self.invoice_number],
            ["Klant-/CMS nummer:", self.supplier_number],
            [date_label, datetime.now().strftime("%d-%m-%Y")],
            ["VAT:", f"{VAT_PERCENT}%"],
        ]

        if is_credit:
            orig = (getattr(self, "original_invoice_number", "") or "").strip()
            reason = (getattr(self, "credit_reason", "") or "").strip()
            if orig:
                info_block.insert(2, ["Original Invoice:", orig])
            if reason:
                info_block.insert(3, ["Reason:", reason])

        bill_table = Table([[bill_block]], colWidths=[doc.width * 0.55])
        inv_table = Table(info_block, colWidths=[5.2 * cm, doc.width * 0.30])
        inv_table.setStyle([("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)])

        elements.append(Table([[bill_table, inv_table]], colWidths=[doc.width * 0.55, doc.width * 0.45], style=[("VALIGN", (0, 0), (-1, -1), "TOP")]))
        elements.append(Spacer(1, 18))

        # Product table
        table_data = [["Part Number", "Description", "Qty", "Price", "Total"]]
        for _, r in df.iterrows():
            qty = int(r["Aantal"]) * sign
            table_data.append([
                str(r["Artikel"]),
                str(r["Omschrijving"]),
                str(qty),
                self._fmt_money(float(r["Prijs"])),
                self._fmt_money(float(r["Total"])),
            ])

        # Summary
        table_data.append(["", "", "", "Subtotal", self._fmt_money(subtotal)])
        if abs(self.verzendkosten) >= 0.005:
            table_data.append(["", "", "", "Shipping/Handling", self._fmt_money(shipping)])
        table_data.append(["", "", "", "Total Excl. VAT", self._fmt_money(total_ex_vat)])
        table_data.append(["", "", "", f"VAT ({VAT_PERCENT}%)", self._fmt_money(vat_amount)])
        table_data.append(["", "", "", "Grand Total", self._fmt_money(grand_total)])

        col_widths = [3.5 * cm, doc.width - (3.5 + 2 + 3 + 3) * cm, 2 * cm, 3 * cm, 3 * cm]
        product_table = Table(table_data, colWidths=col_widths, repeatRows=1)
        product_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ]))

        elements.append(product_table)
        elements.append(Spacer(1, 18))

        footer_lines = FOOTER_CREDIT if is_credit else FOOTER_INVOICE
        for line in footer_lines:
            elements.append(Paragraph(line, normal))

        doc.build(elements)
        buffer.seek(0)

        with open(output_path, "wb") as f:
            f.write(buffer.read())

        # Audit log entry (append-only)
        self._append_audit_log({
            "ts": datetime.now().isoformat(timespec="seconds"),
            "doc_type": "credit" if is_credit else "invoice",
            "doc_no": self.invoice_number,
            "original_invoice": (self.original_invoice_number or "").strip(),
            "reason": (self.credit_reason or "").strip(),
            "bill_to": (self.bill_to or "").strip(),
            "supplier_number": (self.supplier_number or "").strip(),
            "line_count": int(len(df)),
            "shipping": float(self.verzendkosten) * sign,
            "subtotal": subtotal,
            "vat_percent": VAT_PERCENT,
            "vat_amount": float(vat_amount),
            "grand_total": float(grand_total),
            "pdf_path": os.path.abspath(output_path),
        })

        return output_path
