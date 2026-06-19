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

# --- Interest Statement and Invoice (a DIFFERENT CPKC document type) --------- #
# One row per statement (header level).
INTEREST_STMT_COLUMNS = [
    "source_file", "source_page_start", "source_page_end", "parse_warnings",
    "cpkc_invoice_number", "account_number", "invoice_date", "due_date",
    "interest_period", "bill_to_name", "bill_to_address",
    "currency", "line_count", "total_interest", "total_payable",
]
# One row per interest line item (past-due original invoice).
INTEREST_LINE_COLUMNS = [
    "cpkc_invoice_number", "line_no", "original_invoice_no", "reference",
    "waybill_no", "waybill_date", "unit_no", "stcc", "original_due_date",
    "amount", "days", "interest",
]

# Labels whose value is the single line immediately following the label line.
SCALAR_LABELS = {
    "Original Invoice Number": "original_invoice_number",
    "CPKC Invoice Number": "cpkc_invoice_number",
    "CPR Invoice Number": "cpkc_invoice_number",  # Miscellaneous Charges Invoice
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
    "Load Number": "load_order_number",  # Miscellaneous Charges Invoice
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

_PAGE_FOOTER_RE = re.compile(r"Page\s*(\d+)\s*(?:/|of)\s*(\d+)")
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
# Interest Statement parsing
# --------------------------------------------------------------------------- #
_INT_STMT_STOP = {
    "CPKC Invoice Number", "Invoice Date", "Account Number", "Total Payable",
    "Interest Period", "Due Date", "Remit to / Retourner à:",
    "Inquiries to / Pour renseignements:",
}
# A full interest line: invoice# / reference (may contain spaces) / waybill /
# date / unit / stcc / due-date / amount / days / interest.
_INT_ROW_RE = re.compile(
    r"^(?P<invoice_no>\S+)\s+(?P<reference>.*?)\s+(?P<waybill>\d+)\s+"
    r"(?P<date>\d{4}/\d{2}/\d{2})\s+(?P<unit>\S+)\s+(?P<stcc>\d+)\s+"
    r"(?P<due>\d{4}/\d{2}/\d{2})\s+(?P<amount>[\d,]+\.\d{2})\s+"
    r"(?P<days>\d+)\s+(?P<interest>[\d,]+\.\d{2})$"
)
# Just the financial tail (used for irregular rows, e.g. wrapped fee entries).
_INT_TAIL_RE = re.compile(
    r"(?P<amount>[\d,]+\.\d{2})\s+(?P<days>\d+)\s+(?P<interest>[\d,]+\.\d{2})$"
)


def detect_doc_type(lines: list[str]) -> str:
    """Classify a page group as 'interest', 'freight' or 'unknown'."""
    text = "\n".join(lines)
    if "Interest Statement and Invoice" in text:
        return "interest"
    if "Charge Description" in text or "Freight Invoice" in text:
        return "freight"
    return "unknown"


def _int_line(inv_no, ref, wb, date, unit, stcc, due, amount, days, interest):
    return {
        "original_invoice_no": (inv_no or "").strip(),
        "reference": (ref or "").strip(),
        "waybill_no": (wb or "").strip(),
        "waybill_date": (date or "").strip(),
        "unit_no": (unit or "").strip(),
        "stcc": (stcc or "").strip(),
        "original_due_date": (due or "").strip(),
        "amount": _to_number(amount),
        "days": int(days) if days and str(days).isdigit() else None,
        "interest": _to_number(interest),
    }


def parse_interest_statement(inv: _Invoice):
    """Parse one Interest Statement page group -> (header, [lines], warnings)."""
    lines = inv.lines
    n = len(lines)
    h = {c: "" for c in INTEREST_STMT_COLUMNS}
    warnings: list[str] = []

    def val_after(label, off=2):
        for i, l in enumerate(lines):
            if l == label and i + off < n:
                return lines[i + off]
        return ""

    h["cpkc_invoice_number"] = val_after("CPKC Invoice Number")
    h["invoice_date"] = val_after("Invoice Date")
    h["account_number"] = val_after("Account Number")
    h["due_date"] = val_after("Due Date")
    h["interest_period"] = val_after("Interest Period")
    tp_raw = val_after("Total Payable")
    h["total_payable"] = _to_number(tp_raw)
    cur_m = re.search(r"\b([A-Z]{3})\b", tp_raw or "")
    h["currency"] = cur_m.group(1) if cur_m else ""

    # Bill To block: name at +2, address lines until the next known label.
    for i, l in enumerate(lines):
        if l == "Bill To":
            if i + 2 < n:
                h["bill_to_name"] = lines[i + 2]
            addr, j = [], i + 3
            while j < n and lines[j] not in _INT_STMT_STOP:
                addr.append(lines[j])
                j += 1
            h["bill_to_address"] = " | ".join(addr)
            break

    # Line-item table: rows between the 'Intérêts' column header and 'Total Payable'.
    line_rows: list[dict] = []
    in_table = False
    fragments: list[str] = []
    for l in lines:
        if not in_table:
            if l == "Intérêts":
                in_table = True
            continue
        if l.startswith("Total Payable"):
            break
        m = _INT_ROW_RE.match(l)
        if m:
            g = m.groupdict()
            line_rows.append(_int_line(g["invoice_no"], g["reference"], g["waybill"],
                                       g["date"], g["unit"], g["stcc"], g["due"],
                                       g["amount"], g["days"], g["interest"]))
            fragments = []
            continue
        t = _INT_TAIL_RE.search(l)
        if t:
            head = l[: t.start()].strip()
            dates = re.findall(r"\d{4}/\d{2}/\d{2}", head)
            inv_no = "".join(fragments) if fragments else (head.split()[0] if head else "")
            date = dates[0] if dates else ""
            due = dates[-1] if len(dates) > 1 else ""
            line_rows.append(_int_line(inv_no, "", "", date, "", "", due,
                                       t["amount"], t["days"], t["interest"]))
            fragments = []
            continue
        if l.strip():  # invoice-number fragment wrapped onto its own line
            fragments.append(l.strip())

    for k, row in enumerate(line_rows, start=1):
        row["line_no"] = k
        row["cpkc_invoice_number"] = h["cpkc_invoice_number"]

    h["source_page_start"] = inv.page_start
    h["source_page_end"] = inv.page_end
    h["line_count"] = len(line_rows)
    h["total_interest"] = round(sum(r["interest"] or 0 for r in line_rows), 2)

    if not h["cpkc_invoice_number"]:
        warnings.append("missing CPKC invoice number")
    if not line_rows:
        warnings.append("no interest line items parsed")
    if h["total_payable"] is not None and line_rows:
        if abs(h["total_interest"] - h["total_payable"]) > 0.01:
            warnings.append(
                f"interest sum {h['total_interest']} != total_payable "
                f"{h['total_payable']}")
    h["parse_warnings"] = "; ".join(warnings)
    return h, line_rows, warnings


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def parse_pdf_full(pdf_bytes: bytes, source_file: str) -> dict:
    """Parse a PDF, auto-detecting freight invoices vs interest statements.

    Returns a dict of four DataFrames: 'freight_invoices', 'freight_charges',
    'interest_statements', 'interest_lines'. Unknown groups fall back to the
    freight parser (its warnings then flag the mismatch).
    """
    pages = extract_pages(pdf_bytes)
    groups = group_invoices(pages)

    fi_rows, fc_rows, is_rows, il_rows = [], [], [], []
    for g in groups:
        if detect_doc_type(g.lines) == "interest":
            header, line_rows, _ = parse_interest_statement(g)
            header["source_file"] = source_file
            is_rows.append(header)
            il_rows.extend(line_rows)
        else:
            header, charges, _ = parse_invoice(g)
            header["source_file"] = source_file
            fi_rows.append(header)
            fc_rows.extend(charges)

    return {
        "freight_invoices": pd.DataFrame(fi_rows, columns=INVOICE_COLUMNS),
        "freight_charges": pd.DataFrame(fc_rows, columns=CHARGE_COLUMNS),
        "interest_statements": pd.DataFrame(is_rows, columns=INTEREST_STMT_COLUMNS),
        "interest_lines": pd.DataFrame(il_rows, columns=INTEREST_LINE_COLUMNS),
    }


def parse_pdf(pdf_bytes: bytes, source_file: str):
    """Backwards-compatible: return only the freight (invoices, charges) frames."""
    r = parse_pdf_full(pdf_bytes, source_file)
    return r["freight_invoices"], r["freight_charges"]


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


def build_workbook(sheets) -> bytes:
    """Write the given sheets to a styled in-memory .xlsx workbook.

    ``sheets`` is an iterable of ``(sheet_name, dataframe)``; empty/None frames
    are skipped. At least one non-empty sheet is always written.
    """
    pairs = [(name, df) for name, df in sheets if df is not None and not df.empty]
    if not pairs:  # never produce a zero-sheet workbook
        pairs = [("Sheet1", pd.DataFrame())]

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in pairs:
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            ws = writer.sheets[sheet_name[:31]]
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
