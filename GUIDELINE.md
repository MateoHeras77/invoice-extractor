# Technical Guideline — How the Extractor Works

This document explains, in plain terms, how the tool turns a PDF into clean Excel
tables: how it decides which tables to produce, how it locates each value, how it
matches the charge lines, and — importantly — **what happens if the invoice layout
changes** (a field is added, removed, renamed or moved).

It is meant to be readable by a non-developer while staying accurate to the code.
Code lives in two files:

- **`invoice_parser.py`** — all the reading/parsing logic (no UI).
- **`app.py`** — the Streamlit screen (upload, preview, download). It only *displays*
  what the parser produces.

---

## 1. The big picture (the pipeline)

A PDF goes through six stages:

```
PDF bytes
   │
   ▼
1. EXTRACT     read the text of every page                → list of pages (lines of text)
   │
   ▼
2. SPLIT       group pages into individual documents      → one group per invoice/statement
   │            (using the "Page x/y" footer)
   ▼
3. DETECT      classify each group's document type        → "freight" or "interest"
   │
   ▼
4. PARSE       pull each field by its label / position     → header row + line-item rows
   │
   ▼
5. RECONCILE   sanity-check the numbers                    → parse_warnings (empty = clean)
   │
   ▼
6. EXPORT      assemble the sheets into one .xlsx          → Invoices / Charges / etc.
```

The single entry point is `parse_pdf_full(pdf_bytes, source_file)`, which returns
four tables: `freight_invoices`, `freight_charges`, `interest_statements`,
`interest_lines`.

---

## 2. How the text is extracted (stage 1)

Function: `extract_pages()`.

We use the **`pypdfium2`** library to read the text of each page. We chose it because,
for these CPKC documents, it returns **clean word spacing** and — crucially — keeps each
**field label right next to its value**. That adjacency is what makes the rest of the
parsing reliable. Each page becomes a simple list of non-empty text lines.

> Why this matters: these PDFs have no real "spaces" in the font, so a naive reader
> glues words together (e.g. `PUROLATORINC`). `pypdfium2` reconstructs the spacing.

---

## 3. How documents are split apart (stage 2)

Function: `group_invoices()`.

A single uploaded PDF is a stack of many invoices. We separate them using the **page
footer**, e.g. `Page 1/2 … Page 2/2`. A new document starts at `Page 1/…` and closes
when the last page (`Page x/x`) is reached. The footer regex also accepts `Page x of y`
(used by the Miscellaneous Charges invoices).

This is deliberately **not** hardcoded to "2 pages per invoice" — it follows the footer,
so 1-page, 2-page or 3-page documents all work. Each group also records its
`source_page_start` / `source_page_end` for traceability.

---

## 4. How the document type is chosen (stage 3) — "which tables"

Function: `detect_doc_type()`.

For each group we look at the text and classify it:

| If the text contains…                | Type        | Goes to sheets…                       |
|--------------------------------------|-------------|---------------------------------------|
| `Interest Statement and Invoice`     | `interest`  | Interest Statements + Interest Lines  |
| `Charge Description` / `Freight Invoice` | `freight` | Invoices + Charges (+ Financial Summary) |
| neither (unknown)                    | falls back to `freight`, and gets flagged |

So **the tables are chosen automatically per document**. A mixed PDF (freight invoices
*and* interest statements in the same file) is handled — each page is routed to the right
parser. In the workbook, **empty sheets are omitted** (`build_workbook`).

---

## 5. How each value is found (stage 4) — "how we find the data"

The parser is **label-anchored, not position-hardcoded**. Instead of saying "the invoice
number is at x=420, y=90", we say "the value is *next to the label* `CPKC Invoice Number`".
This survives small layout shifts. There are four label patterns (see `parse_invoice`):

1. **Scalar labels** (`SCALAR_LABELS`) — the value is the **next line** after the label.
   Example: `CPKC Invoice Number` → next line `718767433`. Used for invoice number,
   dates, account number, etc.

2. **Block labels** (`BLOCK_LABELS`) — the value is **several lines** (a name + address)
   until the next known label. Example: `Shipper` → company name + address lines.

3. **Inline labels** (`INLINE_LABELS`) — the value is on the **same line**, right after
   the label. Example: `Seal No.: 238232`, `Load/Order Number 53678318`.

4. **Tables** — parsed row by row (see next section).

The set of known labels is just a dictionary near the top of `invoice_parser.py`. Adding,
renaming, or removing a field is mostly a matter of editing those dictionaries.

For interest statements, the header is bilingual (English label, French label, value), so
`parse_interest_statement()` reads the value **two lines below** the English label.

---

## 6. How the charge lines are matched (stage 4, the tables)

Function: `_parse_charge_line()`.

A charge row looks like:

```
FAK                 1   25,000 LBS   3,697.0000 Per Car   CAD   3,697.00   3,697.00 CAD
```

The columns are variable — some rows have a quantity and weight, others don't (carbon
surcharges, reductions). So instead of splitting left-to-right (fragile), we **read from
the right**, where the structure is stable:

1. Find the trailing money block `… CAD <charge> <total> CAD` → gives **charge** and **total**.
2. From what's left, pull the **rate** (a number with 3–4 decimals) and **rate type**
   (`Per Car`, `Percent`, …) if present.
3. Pull **quantity + weight** (`1  25,000 LBS`) if present.
4. Whatever text remains on the left is the **description** (`FAK`, `FUEL SURCHARGE 22.72%`,
   `BC CARBON SURCHARGE`, `REDUCTION`).

This "anchor on the stable part, treat the rest as optional" approach is why irregular
rows don't break it.

The interest table (`parse_interest_statement`) is matched the same spirit: a strict
regex `_INT_ROW_RE` captures the full row (invoice no / reference / waybill / dates /
amount / days / interest), and a looser `_INT_TAIL_RE` handles irregular rows (e.g. a fee
entry whose invoice number wraps onto two lines, which we stitch back together).

---

## 7. How we know it's correct (stage 5) — the safety net

Every row gets a `parse_warnings` column. It is **empty when everything checks out**, and
otherwise explains the problem. The key checks:

- **Freight:** the sum of the charge lines must equal `total_charges`
  (`sum(charges) == total_charges`).
- **Interest:** the sum of the interest column must equal `total_payable`
  (`sum(interest) == total_payable`).
- Missing invoice number, or no line items found, are also flagged.

In the app, a green banner appears only when **all** rows are clean; otherwise the flagged
rows are shown first. This is the "don't silently trust bad data" guard — on the two
provided samples, all rows reconciled to the cent.

---

## 8. What happens if the invoice changes?

This is the most common real-world question. Here is the impact of each kind of change and
what to do. The guiding principle: **because parsing is label-anchored and validated, most
changes either keep working or fail loudly (a `parse_warnings` flag) rather than silently
producing wrong numbers.**

| Change to the invoice | Does it still work? | What to do |
|---|---|---|
| **A value changes** (new amounts, dates, names) | ✅ Yes, automatically | Nothing. This is normal — the parser reads whatever value sits next to the label. |
| **Columns/fields reordered**, minor spacing shifts | ✅ Usually | Nothing. Labels are matched by name, not position. |
| **A new field is added** (a label you want to capture) | ⚠️ Ignored until added | Add the label to the right dictionary (`SCALAR_LABELS`, `INLINE_LABELS`, or `BLOCK_LABELS`) and a column to the schema list. Small change. |
| **A field is removed** | ✅ Yes | That column simply comes back blank. No crash. Optionally drop it from the schema. |
| **A field is renamed** (e.g. `Seal No.` → `Seal Number`) | ⚠️ That field goes blank | Update the label text in the dictionary. One-line change. |
| **A new charge type** appears (e.g. a new surcharge) | ✅ Captured as a charge line | It shows up in the Charges sheet automatically. Only the Financial Summary buckets (fuel/discount) need a tweak if you want it grouped specially. |
| **The number format changes** (currency, negatives, separators) | ✅ Mostly | `_to_number()` already handles commas and spaced negatives; a brand-new format may need a small update there. The reconciliation check will flag a mismatch if something is off. |
| **A new document type** (a layout we've never seen) | ⚠️ Flagged, not silent | It falls back to the freight parser and most rows get a `parse_warnings`. That's the signal to add a new detector in `detect_doc_type()` and a small parser (like we did for interest statements). |
| **The page footer changes** (`Page 1/2` format) | ⚠️ Could mis-group | Update the footer regex `_PAGE_FOOTER_RE`. (We already support both `x/y` and `x of y`.) |

**Bottom line for your boss:** the tool is resilient to everyday variation (new values,
reordered fields, missing fields). Structural changes (renamed labels, brand-new field,
brand-new document type) are **small, localized edits** — usually one dictionary entry or
one regex — and the built-in reconciliation means a bad change shows up as a visible
warning instead of a wrong number slipping through.

---

## 9. Quick "where to change things" map

| You want to… | Edit in `invoice_parser.py` |
|---|---|
| Capture/rename/remove a header field | `SCALAR_LABELS` / `INLINE_LABELS` / `BLOCK_LABELS` + the matching `*_COLUMNS` list |
| Change how a charge line is read | `_parse_charge_line()` |
| Support a new document type | `detect_doc_type()` + a new `parse_*` function + new `*_COLUMNS` |
| Adjust the freight money grouping (fuel/discount) | `build_financial_summary()` |
| Change which sheets are exported | `build_workbook()` and the sheet list in `app.py` |
| Fix number parsing | `_to_number()` |
| Change page splitting | `group_invoices()` / `_PAGE_FOOTER_RE` |
