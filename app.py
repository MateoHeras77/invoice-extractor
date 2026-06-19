"""Streamlit front-end: upload consolidated CPKC/Purolator freight-invoice PDFs
and download a structured Excel workbook (Invoices + Charges + Financial Summary)."""

from __future__ import annotations

import datetime

import pandas as pd
import streamlit as st

import invoice_parser as ip

st.set_page_config(page_title="CPKC Invoice Extractor", page_icon="📦", layout="wide")

# Columns shown by default in the (wide) Invoices table.
CURATED_INVOICE_COLS = [
    "parse_warnings", "cpkc_invoice_number", "invoice_date", "due_date",
    "customer_reference", "account_number", "shipper_name", "consignee_name",
    "origin", "destination", "total_charges", "total_payable", "currency",
    "source_file", "source_page_start",
]

# Display formatting: render a clean copy where missing values are blank (not
# "None"/"NaN") and numbers use thousands separators. The underlying dataframes
# keep their real numeric values (used for downloads, KPIs and the chart).
MONEY_2DP = {
    "Invoices": ["total_charges", "total_payable"],
    "Charges": ["charge", "total"],
    "Financial Summary": ["invoice_amount", "discount", "fuel_surcharge",
                          "total_charges", "total_payable", "tax"],
}


def fmt_money(x, dp=2):
    return "" if pd.isna(x) else f"{x:,.{dp}f}"


def fmt_int(x):
    return "" if pd.isna(x) else f"{int(round(x)):,}"


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
            lambda v: "" if pd.isna(v) else f"{v:.2f}%")
    # Any remaining missing values (e.g. empty text/None columns) -> blank.
    return d.astype(object).where(d.notna(), "")


# --- Cached heavy work (instant on rerun: tab switch, download click, etc.) ---
@st.cache_data(show_spinner=False)
def parse_one(file_bytes: bytes, name: str):
    return ip.parse_pdf(file_bytes, name)


@st.cache_data(show_spinner=False)
def make_workbook(invoices: pd.DataFrame, charges: pd.DataFrame,
                  finance: pd.DataFrame) -> bytes:
    return ip.build_workbook(invoices, charges, finance)


# --- Sidebar: upload, file breakdown, downloads, display options --------------
with st.sidebar:
    st.header("📦 CPKC Invoice Extractor")
    st.caption(
        "Convert consolidated CPKC freight-invoice PDFs (billed to Purolator) "
        "into a structured Excel workbook."
    )
    uploaded = st.file_uploader(
        "Invoice PDF(s)", type="pdf", accept_multiple_files=True,
        help="The consolidated 'Invoices Samples.pdf' or any file with the same layout.",
    )
    show_all_cols = st.checkbox("Show all columns (Invoices)", value=False)

st.title("📦 CPKC / Purolator Freight Invoice Extractor")

if not uploaded:
    st.info("⬅️ Upload at least one PDF in the sidebar to begin.")
    st.stop()

# --- Parse every uploaded file ------------------------------------------------
invoice_frames, charge_frames = [], []
with st.spinner("Parsing invoices…"):
    for file in uploaded:
        try:
            inv_df, ch_df = parse_one(file.getvalue(), file.name)
            invoice_frames.append(inv_df)
            charge_frames.append(ch_df)
        except Exception as exc:  # noqa: BLE001 - surface any parse failure to the user
            st.error(f"Failed to parse **{file.name}**: {exc}")

if not invoice_frames:
    st.stop()

invoices = pd.concat(invoice_frames, ignore_index=True)
charges = pd.concat(charge_frames, ignore_index=True)

if invoices.empty:
    st.warning(
        "No invoices were detected. The file(s) may not use the CPKC freight-"
        "invoice layout this tool expects."
    )
    st.stop()

finance = ip.build_financial_summary(invoices, charges)

# --- KPI row ------------------------------------------------------------------
flagged = int((invoices["parse_warnings"] != "").sum())
total_payable = pd.to_numeric(finance["total_payable"], errors="coerce").sum()
total_fuel = pd.to_numeric(finance["fuel_surcharge"], errors="coerce").sum()
total_discount = pd.to_numeric(finance["discount"], errors="coerce").sum()
currency = ", ".join(sorted({c for c in invoices["currency"] if c})) or "—"
dates = sorted(d for d in invoices["invoice_date"] if d)
date_range = f"{dates[0]} → {dates[-1]}" if dates else "—"

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Invoices", len(invoices))
c2.metric(f"Total payable ({currency})", f"{total_payable:,.2f}")
c3.metric("Total fuel surcharge", f"{total_fuel:,.2f}")
c4.metric("Total discount", f"{total_discount:,.2f}")
c5.metric("Rows with warnings", flagged)
st.caption(f"Charge line items: **{len(charges)}**  ·  Invoice date range: **{date_range}**")

# --- Warnings -----------------------------------------------------------------
if flagged:
    st.warning(
        f"{flagged} invoice(s) have parse warnings — review these before trusting "
        "the export."
    )
    st.dataframe(
        invoices.loc[
            invoices["parse_warnings"] != "",
            ["source_file", "source_page_start", "cpkc_invoice_number", "parse_warnings"],
        ],
        use_container_width=True, hide_index=True,
    )
else:
    st.success("All invoices parsed cleanly (charge totals reconcile).")

# --- Previews -----------------------------------------------------------------
tab_inv, tab_ch, tab_fin = st.tabs([
    f"Invoices ({len(invoices)})",
    f"Charges ({len(charges)})",
    f"Financial Summary ({len(finance)})",
])
with tab_inv:
    inv_view = invoices if show_all_cols else invoices[CURATED_INVOICE_COLS]
    if not show_all_cols:
        st.caption("Curated view — enable *Show all columns* in the sidebar for every field.")
    st.dataframe(for_display(inv_view, "Invoices"),
                 use_container_width=True, hide_index=True)
with tab_ch:
    st.dataframe(for_display(charges, "Charges"),
                 use_container_width=True, hide_index=True)
with tab_fin:
    st.dataframe(for_display(finance, "Financial Summary"),
                 use_container_width=True, hide_index=True)
    chart_df = finance[["cpkc_invoice_number", "total_payable"]].copy()
    chart_df = chart_df.set_index("cpkc_invoice_number")
    st.caption("Total payable per invoice")
    st.bar_chart(chart_df, height=280)

# --- Downloads (sidebar, persistent) ------------------------------------------
stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
workbook = make_workbook(invoices, charges, finance)
with st.sidebar:
    st.divider()
    st.subheader("Downloads")
    st.download_button(
        "⬇️ Full Excel workbook (.xlsx)", data=workbook,
        file_name=f"cpkc_invoices_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True,
    )
    st.caption("One .xlsx with all three tabs: Invoices, Charges, Financial Summary.")
    st.download_button(
        "Invoices (CSV)", invoices.to_csv(index=False).encode("utf-8"),
        file_name=f"invoices_{stamp}.csv", mime="text/csv",
        key="dl_inv", use_container_width=True,
    )
    st.download_button(
        "Charges (CSV)", charges.to_csv(index=False).encode("utf-8"),
        file_name=f"charges_{stamp}.csv", mime="text/csv",
        key="dl_ch", use_container_width=True,
    )
    st.download_button(
        "Financial Summary (CSV)", finance.to_csv(index=False).encode("utf-8"),
        file_name=f"financial_summary_{stamp}.csv", mime="text/csv",
        key="dl_fin", use_container_width=True,
    )

    st.divider()
    with st.expander("Files processed", expanded=True):
        breakdown = (
            invoices.groupby("source_file")
            .agg(invoices=("cpkc_invoice_number", "count"),
                 warnings=("parse_warnings", lambda s: int((s != "").sum())))
            .reset_index()
        )
        ch_counts = charges.groupby("cpkc_invoice_number").size()
        inv_to_file = invoices.set_index("cpkc_invoice_number")["source_file"]
        lines_per_file = (
            ch_counts.groupby(inv_to_file).sum().rename("charge_lines")
        )
        breakdown = breakdown.merge(
            lines_per_file, left_on="source_file", right_index=True, how="left"
        ).fillna({"charge_lines": 0})
        breakdown["charge_lines"] = breakdown["charge_lines"].astype(int)
        st.dataframe(breakdown, use_container_width=True, hide_index=True)
