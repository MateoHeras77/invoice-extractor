"""Parse CPKC (Canadian Pacific Railway) freight invoices billed to Purolator.

The input is a consolidated PDF where each invoice spans two or more pages
(``Page 1/N`` ... ``Page N/N``). Text is extracted with ``pypdfium2`` which, for
this template, renders clean word spacing and keeps each field label adjacent to
its value -- far more reliable than layout-based extraction for this font.

Public API:
    parse_pdf(pdf_bytes, source_file) -> (invoices_df, charges_df)
    build_workbook(invoices_df, charges_df) -> bytes  (an .xlsx workbook)
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import pandas as pd
import pypdfium2 as pdfium

# --------------------------------------------------------------------------- #
# Output schema (also drives column order in the Excel workbook)
# --------------------------------------------------------------------------- #
INVOICE_COLUMNS = [
    # Traceability
    "source_file", "source_page_start", "source_page_end", "parse_warnings",
    # Identifiers
    "cpkc_invoice_number", "original_invoice_number", "account_number",
    "customer_reference", "waybill_number",
    # Dates
    "invoice_date", "original_invoice_date", "due_date", "waybill_date",
    # Parties
    "bill_to_name", "bill_to_address", "shipper_name", "shipper_address",
    "consignee_name", "consignee_address",
    # Routing
    "route", "contract_numbers", "tariff_reference", "origin", "destination",
    "commodity_code", "remarks",
    # Equipment
    "unit_number", "car_type", "plan", "length", "marked_capacity",
    # References
    "load_order_number", "bill_of_lading", "purchase_order", "seal_no",
    "terms_of_sale_number", "conveying_car", "ultimate_consignee",
    # Pickup / delivery (page 2)
    "number_of_pickups", "number_of_deliveries", "first_pickup", "first_delivery",
    # Totals
    "currency", "total_charges", "total_payable", "tax_note",
    # Other
    "instructions",
]

CHARGE_COLUMNS = [
    "cpkc_invoice_number", "line_no", "charge_description", "quantity", "weight",
    "rate", "rate_type", "currency", "charge", "exchange_rate", "total",
]

# Compact, finance-focused view (one row per invoice) for accounting.
FINANCE_COLUMNS = [
    "source_file", "source_page_start", "source_page_end",
    "cpkc_invoice_number", "original_invoice_number", "customer_reference",
    "invoice_date", "due_date", "currency",
    "invoice_amount", "discount", "fuel_surcharge", "fuel_surcharge_pct",
    "total_charges", "total_payable", "tax",
]

# Labels whose value is the single line immediately following the label line.
SCALAR_LABELS = {
    "Original Invoice Number": "original_invoice_number",
    "CPKC Invoice Number": "cpkc_invoice_number",
    "Account Number": "account_number",
    "Customer Reference": "customer_reference",
    "Waybill Number": "waybill_number",
    "Original Invoice Date": "original_invoice_date",
    "Invoice Date": "invoice_date",
    "Due Date": "due_date",
    "Waybill Date": "waybill_date",
    "Origin": "origin",
    "Destination": "destination",
    "Commodity Code": "commodity_code",
    "Tariff Reference": "tariff_reference",
    "Total Payable": "_total_payable_raw",
}

# Labels whose value is the address-style block of following lines, up to the
# next recognised label. First line is the party name; the rest is the address.
BLOCK_LABELS = {
    "Bill To": "bill_to",
    "Shipper": "shipper",
    "Consignee": "consignee",
}

# Labels where the value sits on the SAME line, right after the label text.
INLINE_LABELS = {
    "Load/Order Number": "load_order_number",
    "Seal No.:": "seal_no",
    "Terms of Sale Number": "terms_of_sale_number",
    "Purchase Order:": "purchase_order",
    "Master Bill of Lading": "master_bill_of_lading",
    "Conveying Car:": "conveying_car",
    "Number of pick ups": "number_of_pickups",
    "Number of deliveries": "number_of_deliveries",
}

# Every line that should terminate a BLOCK_LABELS capture.
_STOP_LINES = (
    set(SCALAR_LABELS) | set(BLOCK_LABELS) | {
        "Route", "Shipper's Routing", "Remit to:", "Inquiries to:", "Remarks",
        "References", "Instructions", "Freight Invoice", "No Tax Applied",
        "Number of pick ups", "Number of deliveries", "First Pick Up",
        "First Delivery", "Total Charges", "Total Payable",
    }
)
_INLINE_PREFIXES = tuple(INLINE_LABELS) + (
    "Bill of Lading No.:", "Picked Up from", "Delivered to", "Ultimate Consignee:",
)

_PAGE_FOOTER_RE = re.compile(r"Page\s*(\d+)\s*/\s*(\d+)")
_MONEY_RE = re.compile(r"-?\s?[\d,]+\.\d+")
_UNIT_RE = re.compile(
    r"^([A-Z]{2,4}\s?\d+)\s+([A-Z])\s+(\d+)\s+(\d+)\s+(.*)$"
)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _to_number(text: str):
    """'4,183.28 CAD' / '- 295.76' -> float; return None when not numeric."""
    if text is None:
        return None
    m = _MONEY_RE.search(text)
    if not m:
        return None
    return float(m.group(0).replace(",", "").replace(" ", ""))


def _is_stop(line: str) -> bool:
    return line in _STOP_LINES or line.startswith(_INLINE_PREFIXES)


@dataclass
class _Invoice:
    page_start: int
    page_end: int
    lines: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# PDF -> pages -> invoices
# --------------------------------------------------------------------------- #
def extract_pages(pdf_bytes: bytes) -> list[list[str]]:
    """Return a list of pages, each a list of non-empty, stripped text lines."""
    pages: list[list[str]] = []
    pdf = pdfium.PdfDocument(pdf_bytes)
    try:
        for page in pdf:
            tp = page.get_textpage()
            text = tp.get_text_range()
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            pages.append(lines)
    finally:
        pdf.close()
    return pages


def group_invoices(pages: list[list[str]]) -> list[_Invoice]:
    """Group pages into invoices using the ``Page x/y`` footer.

    A new invoice starts on a page whose footer reads ``Page 1/y`` (or when no
    open invoice exists); it closes when ``x == y``. Pages with no footer are
    appended to the current invoice so nothing is silently dropped.
    """
    invoices: list[_Invoice] = []
    current: _Invoice | None = None

    for idx, lines in enumerate(pages, start=1):
        page_no = total = None
        for ln in lines:
            m = _PAGE_FOOTER_RE.search(ln)
            if m:
                page_no, total = int(m.group(1)), int(m.group(2))
                break

        if page_no == 1 or current is None:
            current = _Invoice(page_start=idx, page_end=idx, lines=list(lines))
            invoices.append(current)
        else:
            current.lines.extend(lines)
            current.page_end = idx

        if page_no is not None and total is not None and page_no >= total:
            current = None  # invoice complete; next page starts a new one

    return invoices


# --------------------------------------------------------------------------- #
# Charge-table parsing
# --------------------------------------------------------------------------- #
def _parse_charge_line(line: str) -> dict | None:
    """Parse one charge row, e.g.
    'FAK 1 25,000 LBS 3,697.0000 Per Car CAD 3,697.00 3,697.00 CAD'.
    Returns None if the line is not a charge row.
    """
    tail = re.search(r"\b(?P<curr>[A-Z]{3})\s+(?P<mid>(?:-?\s?[\d,]+\.\d+\s+)+)"
                     r"(?P=curr)\s*$", line)
    if not tail:
        return None

    nums = _MONEY_RE.findall(tail.group("mid"))
    if not nums:
        return None
    charge = _to_number(nums[0])
    total = _to_number(nums[-1])
    exch = _to_number(nums[1]) if len(nums) >= 3 else None

    head = line[: tail.start()].rstrip()

    rate = rate_type = None
    rate_m = re.search(r"(-?[\d,]+\.\d{3,4})\s*(Per Car|Per \w+|Percent|Flat\w*)?",
                       head)
    qw_m = re.search(r"\s(\d+)\s+([\d,]+)\s+(LBS|KGS)\b", head)

    cut_points = [len(head)]
    if rate_m:
        rate = _to_number(rate_m.group(1))
        rate_type = (rate_m.group(2) or "").strip() or None
        cut_points.append(rate_m.start())
    if qw_m:
        cut_points.append(qw_m.start())
    description = head[: min(cut_points)].strip()

    quantity = int(qw_m.group(1)) if qw_m else None
    weight = float(qw_m.group(2).replace(",", "")) if qw_m else None

    return {
        "charge_description": description,
        "quantity": quantity,
        "weight": weight,
        "rate": rate,
        "rate_type": rate_type,
        "currency": tail.group("curr"),
        "charge": charge,
        "exchange_rate": exch,
        "total": total,
    }


def _parse_charge_table(lines: list[str]) -> list[dict]:
    charges: list[dict] = []
    in_table = False
    for ln in lines:
        if ln.startswith("Charge Description"):
            in_table = True
            continue
        if not in_table:
            continue
        if ln.startswith("Total Charges") or ln.startswith("Total Payable"):
            break
        if ln == "CNT" or ln.startswith("No Tax"):
            continue
        row = _parse_charge_line(ln)
        if row:
            charges.append(row)
    return charges


# --------------------------------------------------------------------------- #
# Single-invoice parsing
# --------------------------------------------------------------------------- #
def parse_invoice(inv: _Invoice) -> tuple[dict, list[dict], list[str]]:
    lines = inv.lines
    h: dict = {c: "" for c in INVOICE_COLUMNS}
    warnings: list[str] = []
    bols: list[str] = []

    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]

        # Scalar label -> next line value (first occurrence wins; page 1 first).
        if line in SCALAR_LABELS and i + 1 < n:
            key = SCALAR_LABELS[line]
            if not h.get(key):
                h[key] = lines[i + 1]
            i += 1

        # Address block -> following lines until a stop line.
        elif line in BLOCK_LABELS:
            base = BLOCK_LABELS[line]
            block: list[str] = []
            j = i + 1
            while j < n and not _is_stop(lines[j]):
                block.append(lines[j])
                j += 1
            if not h.get(f"{base}_name") and block:
                if base == "bill_to":
                    h["bill_to_name"] = block[0]
                    h["bill_to_address"] = " | ".join(block[1:])
                else:
                    h[f"{base}_name"] = block[0]
                    h[f"{base}_address"] = " | ".join(block[1:])
            i = j - 1

        # Inline "label value" on the same line.
        elif line.startswith("Bill of Lading No.:"):
            bols.append(line.split(":", 1)[1].strip())
        elif line.startswith("Ultimate Consignee:"):
            h["ultimate_consignee"] = line.split(":", 1)[1].strip()
        elif line.startswith("Picked Up from") and not h["first_pickup"]:
            h["first_pickup"] = line[len("Picked Up from"):].strip()
        elif line.startswith("Delivered to") and not h["first_delivery"]:
            h["first_delivery"] = line[len("Delivered to"):].strip()
        else:
            for label, key in INLINE_LABELS.items():
                if line.startswith(label):
                    val = line[len(label):].strip()
                    if val and not h.get(key):
                        h[key] = val
                    break

        # Route value: the line after "Route" that is not the contract-number
        # column ("Shipper's Routing" sometimes intervenes).
        if line == "Route" and i + 1 < n and not h["route"]:
            nxt = lines[i + 1]
            if not _is_stop(nxt) and not nxt.isdigit():
                h["route"] = nxt

        # Contract numbers: the digit line(s) following "Contract Numbers".
        if line == "Contract Numbers" and not h["contract_numbers"]:
            vals = []
            j = i + 1
            while j < n and re.match(r"^[\dA-Z]+$", lines[j]):
                vals.append(lines[j])
                j += 1
            h["contract_numbers"] = " ".join(vals)

        # Equipment / unit row.
        if not h["unit_number"]:
            m = _UNIT_RE.match(line)
            if m:
                h["unit_number"] = m.group(1)
                h["car_type"] = m.group(2)
                h["plan"] = m.group(3)
                h["length"] = m.group(4)
                h["marked_capacity"] = m.group(5).strip()

        # Remarks (single value line) and Instructions block.
        if line == "Remarks" and i + 1 < n and not h["remarks"]:
            if not _is_stop(lines[i + 1]):
                h["remarks"] = lines[i + 1]
        if line == "Instructions" and not h["instructions"]:
            block, j = [], i + 1
            while j < n and not _is_stop(lines[j]) \
                    and not _PAGE_FOOTER_RE.search(lines[j]):
                block.append(lines[j])
                j += 1
            h["instructions"] = " ".join(block)

        i += 1

    # Conveying Car may appear inline as "Conveying Car: DTTX 748279".
    for ln in lines:
        if ln.startswith("Conveying Car:") and not h["conveying_car"]:
            h["conveying_car"] = ln.split(":", 1)[1].strip()

    h["bill_of_lading"] = "; ".join(dict.fromkeys(bols))  # dedupe, keep order

    # Totals & currency.
    total_payable_raw = h.pop("_total_payable_raw", "")
    h["total_payable"] = _to_number(total_payable_raw)
    cur_m = re.search(r"\b([A-Z]{3})\b", total_payable_raw or "")
    h["currency"] = cur_m.group(1) if cur_m else ""
    for ln in lines:
        if ln.startswith("Total Charges"):
            h["total_charges"] = _to_number(ln)
        if ln.startswith("No Tax Applied"):
            h["tax_note"] = "No Tax Applied"

    # Charges.
    charges = _parse_charge_table(lines)
    for k, row in enumerate(charges, start=1):
        row["line_no"] = k
        row["cpkc_invoice_number"] = h["cpkc_invoice_number"]

    # Traceability + data-quality checks.
    h["source_page_start"] = inv.page_start
    h["source_page_end"] = inv.page_end
    if not h["cpkc_invoice_number"]:
        warnings.append("missing CPKC invoice number")
    if not charges:
        warnings.append("no charge line items parsed")
    if h["total_charges"] is not None and charges:
        line_sum = round(sum(r["total"] or 0 for r in charges), 2)
        if abs(line_sum - h["total_charges"]) > 0.01:
            warnings.append(
                f"charge sum {line_sum} != total_charges {h['total_charges']}")
    h["parse_warnings"] = "; ".join(warnings)
    return h, charges, warnings


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def parse_pdf(pdf_bytes: bytes, source_file: str):
    """Parse a consolidated PDF into (invoices_df, charges_df)."""
    pages = extract_pages(pdf_bytes)
    invoices = group_invoices(pages)

    inv_rows, charge_rows = [], []
    for inv in invoices:
        header, charges, _ = parse_invoice(inv)
        header["source_file"] = source_file
        inv_rows.append(header)
        charge_rows.extend(charges)

    invoices_df = pd.DataFrame(inv_rows, columns=INVOICE_COLUMNS)
    charges_df = pd.DataFrame(charge_rows, columns=CHARGE_COLUMNS)
    return invoices_df, charges_df


def build_financial_summary(invoices_df: pd.DataFrame,
                            charges_df: pd.DataFrame) -> pd.DataFrame:
    """Compact per-invoice money view: amounts, fuel surcharge, discount, tax.

    Pure aggregation over data already produced by ``parse_pdf`` — no PDF access.
    """
    rows = []
    for _, inv in invoices_df.iterrows():
        cpkc = inv["cpkc_invoice_number"]
        lines = charges_df[charges_df["cpkc_invoice_number"] == cpkc]

        fuel = lines[lines["charge_description"].astype(str)
                     .str.startswith("FUEL SURCHARGE")]
        fuel_total = round(float(fuel["total"].sum()), 2) if len(fuel) else 0.0
        fuel_pct = None
        if len(fuel):
            m = re.search(r"(\d+(?:\.\d+)?)%", str(fuel.iloc[0]["charge_description"]))
            fuel_pct = float(m.group(1)) if m else None

        reduction = lines[lines["charge_description"].astype(str) == "REDUCTION"]
        discount = round(float(reduction["total"].sum()), 2) if len(reduction) else 0.0

        tax = 0.0 if inv["tax_note"] == "No Tax Applied" else _to_number(inv["tax_note"])

        rows.append({
            "source_file": inv["source_file"],
            "source_page_start": inv["source_page_start"],
            "source_page_end": inv["source_page_end"],
            "cpkc_invoice_number": cpkc,
            "original_invoice_number": inv["original_invoice_number"],
            "customer_reference": inv["customer_reference"],
            "invoice_date": inv["invoice_date"],
            "due_date": inv["due_date"],
            "currency": inv["currency"],
            "invoice_amount": inv["total_payable"],
            "discount": discount,
            "fuel_surcharge": fuel_total,
            "fuel_surcharge_pct": fuel_pct,
            "total_charges": inv["total_charges"],
            "total_payable": inv["total_payable"],
            "tax": tax,
        })
    return pd.DataFrame(rows, columns=FINANCE_COLUMNS)


def build_workbook(invoices_df: pd.DataFrame, charges_df: pd.DataFrame,
                   finance_df: pd.DataFrame | None = None) -> bytes:
    """Write all sheets to a styled in-memory .xlsx workbook."""
    if finance_df is None:
        finance_df = build_financial_summary(invoices_df, charges_df)

    buffer = io.BytesIO()
    sheets = (
        ("Invoices", invoices_df),
        ("Charges", charges_df),
        ("Financial Summary", finance_df),
    )
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheets:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "A2"
            for col_idx, col in enumerate(df.columns, start=1):
                width = max(
                    len(str(col)),
                    *(len(str(v)) for v in df[col].head(200)) if len(df) else [0],
                )
                ws.column_dimensions[
                    ws.cell(row=1, column=col_idx).column_letter
                ].width = min(max(width + 2, 10), 60)
    buffer.seek(0)
    return buffer.getvalue()
