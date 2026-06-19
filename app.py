"""Streamlit front-end: upload CPKC PDFs (freight invoices, miscellaneous-charge
invoices and/or interest statements) and download a structured Excel workbook.

The document type is auto-detected per page, so a single mixed PDF is handled."""

from __future__ import annotations

import datetime

import pandas as pd
import streamlit as st

import invoice_parser as ip

st.set_page_config(page_title="PDF Data Extractor", page_icon="📊", layout="wide")

# Columns shown by default in the (wide) Invoices table.
CURATED_INVOICE_COLS = [
    "parse_warnings", "cpkc_invoice_number", "invoice_date", "due_date",
    "customer_reference", "account_number", "shipper_name", "consignee_name",
    "origin", "destination", "total_charges", "total_payable", "currency",
    "source_file", "source_page_start",
]

# Money columns (2 decimals) per display "kind".
MONEY_2DP = {
    "Invoices": ["total_charges", "total_payable"],
    "Charges": ["charge", "total"],
    "Financial Summary": ["invoice_amount", "discount", "fuel_surcharge",
                          "total_charges", "total_payable", "tax"],
    "Interest Statements": ["total_interest", "total_payable"],
    "Interest Lines": ["amount", "interest"],
}


def fmt_money(x, dp=2):
    if pd.isna(x):
        return ""
    try:
        return f"{float(str(x).replace(',', '').strip()):,.{dp}f}"
    except (ValueError, TypeError):
        return str(x)  # non-numeric (e.g. an unsupported layout) — show as-is


def fmt_int(x):
    if pd.isna(x):
        return ""
    try:
        return f"{int(round(float(str(x).replace(',', '').strip()))):,}"
    except (ValueError, TypeError):
        return str(x)


def for_display(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    """Return a display-only copy: blanks for missing, formatted numbers."""
    d = df.copy()
    for col in MONEY_2DP.get(kind, []):
        d[col] = d[col].map(lambda v: fmt_money(v, 2))
    if kind == "Charges":
        d["rate"] = d["rate"].map(lambda v: fmt_money(v, 4))
        for col in ("quantity", "weight"):
            d[col] = d[col].map(fmt_int)
    if kind == "Financial Summary":
        d["fuel_surcharge_pct"] = d["fuel_surcharge_pct"].map(
            lambda v: "" if pd.isna(v) else f"{fmt_money(v, 2)}%")
    if kind == "Interest Lines":
        d["days"] = d["days"].map(fmt_int)
    if kind == "Interest Statements":
        d["line_count"] = d["line_count"].map(fmt_int)
    # Any remaining missing values (e.g. empty text/None columns) -> blank.
    return d.astype(object).where(d.notna(), "")


# --- Cached heavy work (instant on rerun: tab switch, download click, etc.) ---
@st.cache_data(show_spinner=False)
def parse_one(file_bytes: bytes, name: str) -> dict:
    return ip.parse_pdf_full(file_bytes, name)


@st.cache_data(show_spinner=False)
def make_workbook(sheets: list) -> bytes:
    return ip.build_workbook(sheets)


def metric_currency(df: pd.DataFrame) -> str:
    return ", ".join(sorted({c for c in df.get("currency", []) if c})) or "—"


def date_range(df: pd.DataFrame) -> str:
    dates = sorted(d for d in df.get("invoice_date", []) if d)
    return f"{dates[0]} → {dates[-1]}" if dates else "—"


# --- Sidebar: upload + display options ----------------------------------------
with st.sidebar:
    st.header("📊 PDF Data Extractor")
    st.caption(
        "Extract structured data from PDF documents into an Excel workbook. "
        "Document types are auto-detected per page, so mixed PDFs are supported."
    )
    uploaded = st.file_uploader(
        "PDF file(s)", type="pdf", accept_multiple_files=True,
        help="Upload one or more PDF files to process.",
    )
    show_all_cols = st.checkbox("Show all columns", value=False)

st.title("📊 PDF Data Extractor")

if not uploaded:
    st.info("⬅️ Upload at least one PDF in the sidebar to begin.")
    st.stop()

# --- Parse every uploaded file ------------------------------------------------
parts = {k: [] for k in ("freight_invoices", "freight_charges",
                         "interest_statements", "interest_lines")}
with st.spinner("Parsing documents…"):
    for file in uploaded:
        try:
            result = parse_one(file.getvalue(), file.name)
            for key in parts:
                parts[key].append(result[key])
        except Exception as exc:  # noqa: BLE001 - surface any parse failure to the user
            st.error(f"Failed to parse **{file.name}**: {exc}")

invoices = pd.concat(parts["freight_invoices"], ignore_index=True) if parts["freight_invoices"] else pd.DataFrame()
charges = pd.concat(parts["freight_charges"], ignore_index=True) if parts["freight_charges"] else pd.DataFrame()
statements = pd.concat(parts["interest_statements"], ignore_index=True) if parts["interest_statements"] else pd.DataFrame()
int_lines = pd.concat(parts["interest_lines"], ignore_index=True) if parts["interest_lines"] else pd.DataFrame()

has_freight = not invoices.empty
has_interest = not statements.empty

if not has_freight and not has_interest:
    st.warning(
        "No supported documents were detected. The file(s) may use a layout this "
        "tool does not support yet."
    )
    st.stop()

finance = ip.build_financial_summary(invoices, charges) if has_freight else pd.DataFrame()

# --- Freight invoices section -------------------------------------------------
if has_freight:
    st.subheader("Documents")
    flagged = int((invoices["parse_warnings"] != "").sum())
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Documents", len(invoices))
    c2.metric(f"Total payable ({metric_currency(invoices)})",
              f"{pd.to_numeric(finance['total_payable'], errors='coerce').sum():,.2f}")
    c3.metric("Total fuel surcharge",
              f"{pd.to_numeric(finance['fuel_surcharge'], errors='coerce').sum():,.2f}")
    c4.metric("Total discount",
              f"{pd.to_numeric(finance['discount'], errors='coerce').sum():,.2f}")
    c5.metric("Rows with warnings", flagged)
    st.caption(f"Charge line items: **{len(charges)}**  ·  Invoice date range: "
               f"**{date_range(invoices)}**")

    if flagged:
        st.warning(f"{flagged} record(s) have parse warnings — review before trusting the export.")
        st.dataframe(
            invoices.loc[invoices["parse_warnings"] != "",
                         ["source_file", "source_page_start", "cpkc_invoice_number", "parse_warnings"]],
            use_container_width=True, hide_index=True,
        )
    else:
        st.success("All records parsed cleanly (totals reconcile).")

    t_inv, t_ch, t_fin = st.tabs([
        f"Invoices ({len(invoices)})", f"Charges ({len(charges)})",
        f"Financial Summary ({len(finance)})"])
    with t_inv:
        inv_view = invoices if show_all_cols else invoices[CURATED_INVOICE_COLS]
        if not show_all_cols:
            st.caption("Curated view — enable *Show all columns* in the sidebar for every field.")
        st.dataframe(for_display(inv_view, "Invoices"), use_container_width=True, hide_index=True)
    with t_ch:
        st.dataframe(for_display(charges, "Charges"), use_container_width=True, hide_index=True)
    with t_fin:
        st.dataframe(for_display(finance, "Financial Summary"), use_container_width=True, hide_index=True)
        chart = finance[["cpkc_invoice_number", "total_payable"]].set_index("cpkc_invoice_number")
        st.caption("Total payable per invoice")
        st.bar_chart(chart, height=280)

# --- Interest statements section ----------------------------------------------
if has_interest:
    st.divider()
    st.subheader("Interest statements")
    iflag = int((statements["parse_warnings"] != "").sum())
    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Statements", len(statements))
    d2.metric(f"Total payable ({metric_currency(statements)})",
              f"{pd.to_numeric(statements['total_payable'], errors='coerce').sum():,.2f}")
    d3.metric("Interest line items", len(int_lines))
    d4.metric("Rows with warnings", iflag)
    st.caption(f"Invoice date range: **{date_range(statements)}**")

    if iflag:
        st.warning(f"{iflag} statement(s) have parse warnings — review before trusting the export.")
        st.dataframe(
            statements.loc[statements["parse_warnings"] != "",
                           ["source_file", "source_page_start", "cpkc_invoice_number", "parse_warnings"]],
            use_container_width=True, hide_index=True,
        )
    else:
        st.success("All interest statements parsed cleanly (interest totals reconcile).")

    t_st, t_ln = st.tabs([f"Interest Statements ({len(statements)})",
                          f"Interest Lines ({len(int_lines)})"])
    with t_st:
        st.dataframe(for_display(statements, "Interest Statements"),
                     use_container_width=True, hide_index=True)
    with t_ln:
        st.dataframe(for_display(int_lines, "Interest Lines"),
                     use_container_width=True, hide_index=True)

# --- Downloads (sidebar, persistent) ------------------------------------------
stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
sheet_specs = [
    ("Invoices", invoices), ("Charges", charges), ("Financial Summary", finance),
    ("Interest Statements", statements), ("Interest Lines", int_lines),
]
workbook = make_workbook([(n, df) for n, df in sheet_specs])

with st.sidebar:
    st.divider()
    st.subheader("Downloads")
    st.download_button(
        "⬇️ Full Excel workbook (.xlsx)", data=workbook,
        file_name=f"cpkc_invoices_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True,
    )
    st.caption("One .xlsx with every populated sheet.")
    for name, df in sheet_specs:
        if df.empty:
            continue
        st.download_button(
            f"{name} (CSV)", df.to_csv(index=False).encode("utf-8"),
            file_name=f"{name.lower().replace(' ', '_')}_{stamp}.csv",
            mime="text/csv", key=f"dl_{name}", use_container_width=True,
        )

    st.divider()
    with st.expander("Files processed", expanded=True):
        rows = []
        for fname in sorted({*invoices.get("source_file", []), *statements.get("source_file", [])}):
            inv_n = int((invoices.get("source_file", pd.Series(dtype=str)) == fname).sum())
            st_n = int((statements.get("source_file", pd.Series(dtype=str)) == fname).sum())
            warn = int((invoices.loc[invoices.get("source_file", pd.Series(dtype=str)) == fname, "parse_warnings"] != "").sum()) if has_freight else 0
            warn += int((statements.loc[statements.get("source_file", pd.Series(dtype=str)) == fname, "parse_warnings"] != "").sum()) if has_interest else 0
            rows.append({"file": fname, "invoices": inv_n, "statements": st_n, "warnings": warn})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
