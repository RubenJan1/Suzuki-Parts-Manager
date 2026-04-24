"""
Parse een eerder gegenereerde factuur-PDF terug naar bewerkbare data.

Verwacht PDF-structuur zoals aangemaakt door engine_factuurmaker.py:
- Tabelkolommen: Part Number | Description | Qty | Price | Total
- Labels: "Invoice Number:" / "Credit Note Number:", "Klant-/CMS nummer:", "Bill to"
- Documenttype herkenning via tekst "CREDITFACTUUR" of "FACTUUR"
"""

from __future__ import annotations

import re
import pandas as pd


def _parse_euro(s: str) -> float:
    """Zet '€ 12.345,67' om naar float 12345.67."""
    s = (s or "").strip().replace("€", "").strip()
    # Europese notatie: punt = duizendteken, komma = decimaal
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_invoice_pdf(path: str) -> dict:
    """
    Parseer een factuur-PDF en geef bewerkbare data terug.

    Geeft dict terug met dezelfde sleutels als load_draft_json():
      invoice_number, supplier_number, bill_to, billing_address,
      document_type, original_invoice_number, credit_reason,
      verzendkosten, df (DataFrame met Artikel/Omschrijving/Aantal/Prijs)

    Gooit RuntimeError als het bestand niet geparseerd kan worden.
    """
    try:
        import pdfplumber
    except ImportError:
        raise RuntimeError(
            "pdfplumber is niet geïnstalleerd.\n"
            "Installeer het via:  pip install pdfplumber"
        )

    with pdfplumber.open(path) as pdf:
        full_text = "\n".join(
            (page.extract_text() or "") for page in pdf.pages
        )

        # Tabellen uit alle pagina's samenvoegen
        all_tables = []
        for page in pdf.pages:
            tbls = page.extract_tables()
            if tbls:
                all_tables.extend(tbls)

    # ── Document type ──────────────────────────────────────────────
    is_credit = "CREDITFACTUUR" in full_text.upper()
    document_type = "credit" if is_credit else "invoice"

    # ── Factuurnummer ──────────────────────────────────────────────
    invoice_number = ""
    if is_credit:
        m = re.search(r"Credit Note Number:\s*(\S+)", full_text, re.IGNORECASE)
    else:
        m = re.search(r"Invoice Number:\s*(\S+)", full_text, re.IGNORECASE)
    if m:
        invoice_number = m.group(1).strip()

    # ── Originele factuur (alleen bij credit) ──────────────────────
    original_invoice_number = ""
    m = re.search(r"Original Invoice:\s*(\S+)", full_text, re.IGNORECASE)
    if m:
        original_invoice_number = m.group(1).strip()

    # ── Reden (alleen bij credit) ──────────────────────────────────
    credit_reason = ""
    m = re.search(r"Reason:\s*(.+)", full_text, re.IGNORECASE)
    if m:
        credit_reason = m.group(1).strip()

    # ── CMS/klantnummer ───────────────────────────────────────────
    supplier_number = ""
    m = re.search(r"Klant-/CMS nummer:\s*(\S+)", full_text, re.IGNORECASE)
    if m:
        supplier_number = m.group(1).strip()

    # ── Bill to + adres ───────────────────────────────────────────
    # De "bill_to" staat als vette naam vóór het adresblok.
    # In de PDF is het de eerste niet-lege regel na "FACTUUR"/"CREDITFACTUUR"
    # tot aan de eerste info-label ("Document:", "Invoice Number:", …).
    bill_to = ""
    billing_address = ""
    _info_labels = r"(Document:|Invoice Number:|Credit Note Number:|Klant-/CMS nummer:|Invoice Date:|Credit Date:|VAT:|Original Invoice:|Reason:|Payment term:)"
    m_doc = re.search(r"(FACTUUR|CREDITFACTUUR)", full_text, re.IGNORECASE)
    if m_doc:
        after_doc = full_text[m_doc.end():]
        # Alles vóór de eerste info-label
        m_info = re.search(_info_labels, after_doc, re.IGNORECASE)
        address_block = after_doc[:m_info.start()] if m_info else after_doc[:300]
        addr_lines = [ln.strip() for ln in address_block.strip().splitlines() if ln.strip()]
        # Verwijder regels die bedrijfsinfo zijn (Vlaandere, Tel:, IBAN:, VAT:)
        addr_lines = [
            ln for ln in addr_lines
            if not any(kw in ln for kw in ["Vlaandere", "Tel:", "IBAN:", "VAT:", "C.O.C", "de Marne"])
        ]
        if addr_lines:
            bill_to = addr_lines[0]
            billing_address = "\n".join(addr_lines[1:]) if len(addr_lines) > 1 else ""

    # ── Producttabel ───────────────────────────────────────────────
    # Zoek de tabel waarvan de header begint met "Part Number"
    product_rows = []
    HEADER = ["Part Number", "Description", "Qty", "Price", "Total"]
    SUMMARY_LABELS = {"Subtotal", "Shipping/Handling", "Total Excl. VAT", "Grand Total"}

    verzendkosten = 0.0

    for tbl in all_tables:
        if not tbl or len(tbl) < 2:
            continue
        header = [str(c or "").strip() for c in tbl[0]]
        if header[:5] != HEADER:
            continue

        for row in tbl[1:]:
            if not row or len(row) < 5:
                continue
            col0 = str(row[0] or "").strip()
            col1 = str(row[1] or "").strip()
            col2 = str(row[2] or "").strip()
            col3 = str(row[3] or "").strip()
            col4 = str(row[4] or "").strip()

            # Samenvattingsrijen overslaan (lege Part Number + summary label)
            if not col0 and col3 in SUMMARY_LABELS:
                if col3 == "Shipping/Handling":
                    verzendkosten = abs(_parse_euro(col4))
                continue
            if not col0 and not col1:
                continue

            try:
                qty = abs(int(col2)) if col2 else 0
            except ValueError:
                qty = 0

            price = _parse_euro(col3)

            product_rows.append({
                "Artikel":      col0,
                "Omschrijving": col1,
                "Aantal":       qty,
                "Prijs":        price,
            })

    if not product_rows:
        raise RuntimeError(
            "Geen producttabel gevonden in de PDF.\n"
            "Controleer of dit een door Suzuki Parts Manager gegenereerde factuur is."
        )

    df = pd.DataFrame(product_rows, columns=["Artikel", "Omschrijving", "Aantal", "Prijs"])
    df["Aantal"] = pd.to_numeric(df["Aantal"], errors="coerce").fillna(0).astype(int)
    df["Prijs"]  = pd.to_numeric(df["Prijs"],  errors="coerce").fillna(0.0).astype(float)

    return {
        "invoice_number":          invoice_number,
        "supplier_number":         supplier_number,
        "bill_to":                 bill_to,
        "billing_address":         billing_address,
        "document_type":           document_type,
        "original_invoice_number": original_invoice_number,
        "credit_reason":           credit_reason,
        "verzendkosten":           verzendkosten,
        "df":                      df,
    }
