# CPKC / Purolator Freight Invoice Extractor

A small Python + Streamlit tool that turns consolidated PDFs of **CPKC (Canadian
Pacific Railway) freight invoices billed to Purolator** into a structured Excel
workbook. Each invoice (2+ pages) is detected, split out, and parsed into a
per-invoice header row plus per-line charge rows, with traceability back to the
source file and page numbers.

## Setup

Requires [`uv`](https://github.com/astral-sh/uv).

```bash
uv venv
source .venv/bin/activate
uv pip install -r requirements.txt
```

## Run

```bash
source .venv/bin/activate
streamlit run app.py
```

Then open the URL Streamlit prints (default http://localhost:8501), upload one or
more PDFs, review the preview and any warnings, and click **Download Excel
workbook**.

## How it works

- **Extraction** — `pypdfium2` renders the page text. For this invoice template it
  recovers clean word spacing and keeps each field label adjacent to its value.
- **Splitting** — invoices are grouped using the `Page x/y` footer (not a hardcoded
  page count), so 1-page or 3+-page invoices also work.
- **Parsing** — label-anchored, not position-hardcoded, so minor layout shifts
  don't break it. Charge rows are parsed from the right (currency + amounts) so the
  variable description/rate/quantity columns are handled robustly.
- **Validation** — each invoice gets a `parse_warnings` value; it flags a missing
  invoice number, no charge rows, or a charge sum that doesn't reconcile with
  `total_charges`. Clean invoices have an empty `parse_warnings`.

## Output workbook

### Sheet `Invoices` (one row per invoice)

| Group | Columns |
|-------|---------|
| Traceability | `source_file`, `source_page_start`, `source_page_end`, `parse_warnings` |
| Identifiers | `cpkc_invoice_number`, `original_invoice_number`, `account_number`, `customer_reference`, `waybill_number` |
| Dates | `invoice_date`, `original_invoice_date`, `due_date`, `waybill_date` |
| Parties | `bill_to_name`, `bill_to_address`, `shipper_name`, `shipper_address`, `consignee_name`, `consignee_address` |
| Routing | `route`, `contract_numbers`, `tariff_reference`, `origin`, `destination`, `commodity_code`, `remarks` |
| Equipment | `unit_number`, `car_type`, `plan`, `length`, `marked_capacity` |
| References | `load_order_number`, `bill_of_lading`, `purchase_order`, `seal_no`, `terms_of_sale_number`, `conveying_car`, `ultimate_consignee` |
| Pickup / delivery | `number_of_pickups`, `number_of_deliveries`, `first_pickup`, `first_delivery` |
| Totals | `currency`, `total_charges`, `total_payable`, `tax_note` |
| Other | `instructions` |

`bill_of_lading` joins multiple Bill of Lading numbers with `; `. Address lines are
joined with ` | `.

### Sheet `Charges` (one row per line item)

`cpkc_invoice_number` (foreign key), `line_no`, `charge_description`, `quantity`,
`weight`, `rate`, `rate_type`, `currency`, `charge`, `exchange_rate`, `total`.

## Assumptions / scope

- Tuned to the CPKC freight-invoice layout in the provided sample (all `CAD`,
  `No Tax Applied`). Currency and the tax note are captured as fields rather than
  assumed, so other currencies surface correctly if present.
- A non-CPKC PDF parses to zero invoices and the app reports that gracefully
  instead of crashing.
