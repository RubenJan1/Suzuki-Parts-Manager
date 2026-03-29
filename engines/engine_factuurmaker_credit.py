import re
import os
import pandas as pd
from datetime import datetime
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, Image as RLImage
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image
from utils.paths import output_root

# =====================================================
# CONSTANTEN
# =====================================================

VAT_PERCENT = 21.0

COMPANY_INFO = [
    "Vlaandere Motoren - de Marne 136 B",
    "8701 MC - Bolsward",
    "Tel: +316-41484547",
    "IBAN: NL49 RABO 0372 0041 64",
    "VAT: 8077 51 911 B01 | C.O.C: 01018576",
]

FOOTER_LINES = [
    "Payment term: 14 days.",
    "Returns allowed within 15 days after receiving the item.",
    "Vlaandere Motoren — IBAN: NL49 RABO 0372 0041 64",
    "VATnumber 8077 51 911 B01 | C.O.C.number 01018576",
]


# =====================================================
# ENGINE
# =====================================================

class FactuurMakerEngine:
    """
    Engine voor CMS facturen:
    - combineert meerdere CMS-bestellingen
    - genereert professionele PDF
    """

    def __init__(self):
        self.cms_paths = []
        self.work_df = None 
        self.verzendkosten = 0.0

        self.invoice_number = f"INV-{datetime.now().strftime('%Y%m%d%H%M')}"
        self.supplier_number = ""
        self.bill_to = "CMS"
        self.billing_address = (
            "Artemisweg 245\n"
            "8239 DD Lelystad\n"
            "Netherlands"
        )

        # Document type
        self.document_type = "invoice"  # "invoice" | "credit"
        self.original_invoice_number = ""
        self.credit_reason = ""

        self.logo_path = "assets/logo.png"
        self.output_dir = output_root() / "facturen"
        self.output_dir.mkdir(parents=True, exist_ok=True)

    # -------------------------------------------------
    # DATA
    # -------------------------------------------------

    def add_cms_bestelling(self, path: str):
        if path not in self.cms_paths:
            self.cms_paths.append(path)
            self._rebuild_work_df()


    def clear_bestellingen(self):
        self.cms_paths.clear()
        self.work_df = pd.DataFrame(
            columns=["Artikel", "Omschrijving", "Aantal", "Prijs"]
        )


    def _sort_key(self, artikel):
        """
        Suzuki-onderdelen correct sorteren:
        elk numeriek segment telt mee
        """
        try:
            parts = str(artikel).strip().split("-")
            nums = []

            for p in parts:
                # alleen cijfers houden
                digits = "".join(ch for ch in p if ch.isdigit())
                nums.append(int(digits) if digits else 0)

            return tuple(nums)
        except:
            return (0,)



    def _rebuild_work_df(self):
        """
        Leest ALLE CMS-bestellingen in,
        combineert ze tot één lijst
        en sorteert GLOBAAL op artikelnummer
        """
        dfs = []

        for p in self.cms_paths:
            df = pd.read_excel(
                p,
                header=None,
                usecols=[0, 1, 2, 3],
                names=["Artikel", "Omschrijving", "Aantal", "Prijs"]
            )
            dfs.append(df)

        if not dfs:
            self.work_df = pd.DataFrame(
                columns=["Artikel", "Omschrijving", "Aantal", "Prijs"]
            )
            return

        merged = pd.concat(dfs, ignore_index=True)

        # normaliseren
        merged["Artikel"] = merged["Artikel"].astype(str).str.strip()
        merged["Omschrijving"] = merged["Omschrijving"].astype(str)
        merged["Aantal"] = pd.to_numeric(
            merged["Aantal"], errors="coerce"
        ).fillna(0).astype(int)
        merged["Prijs"] = pd.to_numeric(
            merged["Prijs"], errors="coerce"
        ).fillna(0.0).astype(float)

        # dubbelen samenvoegen op (Artikel + Prijs)
        merged = self._merge_duplicates_df(merged)

        # 🔥 HIER gebeurt de ENIGE sortering 🔥
        merged["_sort"] = merged["Artikel"].apply(self._sort_key)
        merged = merged.sort_values("_sort", kind="mergesort")
        merged = merged.drop(columns=["_sort"]).reset_index(drop=True)

        self.work_df = merged

    def sort_work_df(self):
        """
        Sorteer de huidige werklijst opnieuw op artikelnummer.
        Handig na verwijderen/handmatig toevoegen/edits in de UI.
        """
        if self.work_df is None or self.work_df.empty:
            return
        df = self.work_df.copy()
        df["_sort"] = df["Artikel"].apply(self._sort_key)
        df = df.sort_values("_sort", kind="mergesort").drop(columns=["_sort"]).reset_index(drop=True)
        self.work_df = df

    def merge_work_df(self):
        """
        Combineer dubbele regels op (Artikel + Prijs) door Aantal op te tellen.
        Alleen samenvoegen als zowel artikelnummer als prijs exact overeenkomen.
        """
        if self.work_df is None or self.work_df.empty:
            return
        self.work_df = self._merge_duplicates_df(self.work_df)

    def _merge_duplicates_df(self, df: pd.DataFrame) -> pd.DataFrame:
        # defensief: zorg dat kolommen bestaan
        for col in ["Artikel", "Omschrijving", "Aantal", "Prijs"]:
            if col not in df.columns:
                df[col] = "" if col in ["Artikel", "Omschrijving"] else 0

        tmp = df.copy()

        # normaliseren voor betrouwbare vergelijking
        tmp["Artikel"] = tmp["Artikel"].astype(str).str.strip()
        tmp["Omschrijving"] = tmp["Omschrijving"].fillna("").astype(str)
        tmp["Aantal"] = pd.to_numeric(tmp["Aantal"], errors="coerce").fillna(0).astype(int)
        tmp["Prijs"] = pd.to_numeric(tmp["Prijs"], errors="coerce").fillna(0.0).astype(float)

        def pick_omschrijving(series: pd.Series) -> str:
            # pak de eerste niet-lege omschrijving
            for v in series.astype(str):
                v = (v or "").strip()
                if v and v.lower() != "nan":
                    return v
            return ""

        grouped = (
            tmp.groupby(["Artikel", "Prijs"], as_index=False, sort=False)
            .agg({"Omschrijving": pick_omschrijving, "Aantal": "sum"})
        )

        # kolomvolgorde terug netjes
        grouped = grouped[["Artikel", "Omschrijving", "Aantal", "Prijs"]]
        return grouped.reset_index(drop=True)


    def _format_address_lines(self, address_text: str):
        """
        Zorgt dat een adres altijd netjes onder elkaar komt in de PDF.
        - Als user 1 regel plakt, splitsen we op komma's.
        - Meerdere spaties -> 1 spatie.
        """
        s = str(address_text or "").strip()
        if not s:
            return []
        s = re.sub(r"[ \t]+", " ", s)
        if "\n" in s:
            lines = [ln.strip() for ln in s.split("\n") if ln.strip()]
            return lines
        # één regel: split op komma's
        if "," in s:
            parts = [p.strip() for p in s.split(",") if p.strip()]
            return parts

        # Nederlandse postcode patroon (1234 AB)
        m_pc = re.search(r"\b(\d{4}\s*[A-Z]{2})\b", s.upper())
        if m_pc:
            pc = m_pc.group(1).replace(" ", "")
            # probeer te splitsen: <straat+nr> <postcode> <plaats> <land>
            before = s[:m_pc.start()].strip()
            after = s[m_pc.end():].strip()
            # land = laatste woord als het duidelijk een landnaam is
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
        """
        Geeft de huidige gecombineerde werklijst terug.
        GEEN logica, GEEN sortering.
        """
        if self.work_df is None:
            return pd.DataFrame(
                columns=["Artikel", "Omschrijving", "Aantal", "Prijs"]
            )

        return self.work_df.copy()



    # -------------------------------------------------
    # PDF
    # -------------------------------------------------

    def generate_pdf(self) -> str:
        self.merge_work_df()
        df = self.merged_df().copy()

        # 🔥 DEFINITIEVE GLOBALE SORTERING VOOR FACTUUR 🔥
        df["_sort"] = df["Artikel"].apply(self._sort_key)
        df = df.sort_values("_sort", kind="mergesort").reset_index(drop=True)
        df = df.drop(columns=["_sort"])


        if df.empty:
            raise RuntimeError("No CMS orders loaded")

        output_dir = self.output_dir
        output_path = str(output_dir / file_name)

        is_credit = (getattr(self, "document_type", "invoice") == "credit")
        doc_label = "CREDITFACTUUR" if is_credit else "FACTUUR"
        file_prefix = "Credit_" if is_credit else "Invoice_"
        file_name = f"{file_prefix}{self.invoice_number}.pdf"
        output_path = os.path.join(output_dir, file_name)

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

        title = ParagraphStyle(
            "Title",
            parent=styles["Heading1"],
            fontSize=24,
            textColor=colors.HexColor("#111827"),
        )

        normal = ParagraphStyle(
            "Normal",
            parent=styles["Normal"],
            fontSize=10,
            textColor=colors.HexColor("#111827"),
        )

        bold = ParagraphStyle(
            "Bold",
            parent=styles["Normal"],
            fontSize=10,
            textColor=colors.HexColor("#111827"),
            fontName="Helvetica-Bold",
        )

        elements = []

        # -------------------------------------------------
        # LOGO
        # -------------------------------------------------
        if os.path.exists(self.logo_path):
            img = Image.open(self.logo_path)
            bio = BytesIO()
            img.save(bio, format="PNG")
            bio.seek(0)

            logo = RLImage(bio, width=12 * cm, height=5 * cm, kind="proportional")
            elements.append(logo)

        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"<b>{doc_label}</b>", title))
        if is_credit and getattr(self, "original_invoice_number", ""):
            elements.append(Paragraph(f"Reference: {self.original_invoice_number}", normal))
        if is_credit and getattr(self, "credit_reason", ""):
            elements.append(Paragraph(f"Reason: {self.credit_reason}", normal))
        elements.append(Spacer(1, 10))


        # -------------------------------------------------
        # COMPANY INFO
        # -------------------------------------------------
        for line in COMPANY_INFO:
            elements.append(Paragraph(line, normal))

        elements.append(Spacer(1, 14))

        # -------------------------------------------------
        # BILLING + INVOICE INFO
        # -------------------------------------------------
        bill_block = [Paragraph(f"<b>{self.bill_to}</b>", bold)]
        for l in self._format_address_lines(self.billing_address):
            bill_block.append(Paragraph(l, normal))

        invoice_block = [
            ["Document:", doc_label],
            ["Invoice Number:", self.invoice_number],
            ["Supplier Number:", self.supplier_number],
            ["Invoice Date:", datetime.now().strftime("%d-%m-%Y")],
            ["VAT:", f"{VAT_PERCENT}%"],
        ]

        if is_credit:
            orig = getattr(self, "original_invoice_number", "").strip()
            reason = getattr(self, "credit_reason", "").strip()
            if orig:
                invoice_block.insert(2, ["Original Invoice:", orig])
            if reason:
                invoice_block.insert(3, ["Reason:", reason])

        bill_table = Table([[bill_block]], colWidths=[doc.width * 0.55])
        inv_table = Table(invoice_block, colWidths=[4 * cm, doc.width * 0.25])

        inv_table.setStyle([
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ])

        elements.append(
            Table(
                [[bill_table, inv_table]],
                colWidths=[doc.width * 0.55, doc.width * 0.45],
                style=[("VALIGN", (0, 0), (-1, -1), "TOP")],
            )
        )

        elements.append(Spacer(1, 18))

        # -------------------------------------------------
        # PRODUCT TABLE
        # -------------------------------------------------
        sign = -1 if is_credit else 1
        df["Total"] = (df["Aantal"] * df["Prijs"]) * sign

        subtotal = df["Total"].sum()
        total_ex_vat = subtotal + (self.verzendkosten * sign)
        vat_amount = total_ex_vat * (VAT_PERCENT / 100)
        grand_total = total_ex_vat + vat_amount

        table_data = [
            ["Part Number", "Description", "Qty", "Price", "Total"]
        ]

        for _, r in df.iterrows():
            table_data.append([
                str(r["Artikel"]),
                str(r["Omschrijving"]),
                str(int(r["Aantal"]) * sign),
                f"€ {r['Prijs']:.2f}".replace(".", ","),
                f"€ {r['Total']:.2f}".replace(".", ","),
            ])

        table_data += [
            ["", "", "", "Subtotal", f"€ {subtotal:.2f}".replace(".", ",")],
            ["", "", "", "Shipping/Handling", f"€ {(self.verzendkosten * sign):.2f}".replace(".", ",")],
            ["", "", "", "Total Excl. VAT", f"€ {total_ex_vat:.2f}".replace(".", ",")],
            ["", "", "", f"VAT ({VAT_PERCENT}%)", f"€ {vat_amount:.2f}".replace(".", ",")],
            ["", "", "", "Grand Total", f"€ {grand_total:.2f}".replace(".", ",")],
        ]

        col_widths = [
            3.5 * cm,
            doc.width - (3.5 + 2 + 3 + 3) * cm,
            2 * cm,
            3 * cm,
            3 * cm,
        ]

        product_table = Table(table_data, colWidths=col_widths, repeatRows=1)
        product_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ]))

        elements.append(product_table)
        elements.append(Spacer(1, 18))

        # -------------------------------------------------
        # FOOTER
        # -------------------------------------------------
        for line in FOOTER_LINES:
            elements.append(Paragraph(line, normal))

        doc.build(elements)
        buffer.seek(0)

        with open(output_path, "wb") as f:
            f.write(buffer.read())

        return output_path
