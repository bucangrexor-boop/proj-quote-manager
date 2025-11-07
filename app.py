
# ===============================================================
# Streamlit Project Quotation Manager 
# File: app.py
# ===============================================================

# ----------------------
# Imports
# ----------------------
import io
import json
import math
import time
import pandas as pd
import streamlit as st
import gspread
from datetime import datetime
from google.oauth2 import service_account
from gspread.exceptions import APIError
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ----------------------
# Streamlit Configuration
# ----------------------
st.set_page_config(page_title="Project Quotation Manager", layout="wide")

# ----------------------
# Constants
# ----------------------
GSHEETS_KEY_SECRET = "gsheets_key"
GCP_SA_SECRET = "gcp_service_account"

SHEET_HEADERS = [
    "Item", "Part Number", "Description", "Quantity", "Unit", "Unit Price", "Subtotal"
]

TERMS_LABELS = [
    ("Terms of payment", "I2", "J2"),
    ("Delivery", "I3", "J3"),
    ("Warranty", "I4", "J4"),
    ("Price Validity", "I5", "J5"),
    ("Discount", "I8", "J8")
]

# ===============================================================
# Helper Functions
# ===============================================================

@st.cache_resource
def get_gspread_client():
    if GCP_SA_SECRET not in st.secrets or GSHEETS_KEY_SECRET not in st.secrets:
        st.error("Google secrets are missing. Add 'gcp_service_account' and 'gsheets_key' in Streamlit Secrets.")
        st.stop()
    creds_info = json.loads(st.secrets[GCP_SA_SECRET])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = service_account.Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(credentials)


@st.cache_resource(ttl=600)
def open_spreadsheet():
    client = get_gspread_client()
    key = st.secrets[GSHEETS_KEY_SECRET]
    try:
        return client.open_by_key(key)
    except Exception as e:
        st.error(f"❌ Could not open sheet: {type(e).__name__} - {e}")
        st.stop()


def worksheet_create_with_headers(ss, title: str):
    ws = ss.add_worksheet(title=title, rows=100, cols=20)
    ws.update([SHEET_HEADERS])
    label_updates = [{"range": label_cell, "values": [[label]]} for label, label_cell, _ in TERMS_LABELS]
    ws.batch_update([{"range": u["range"], "values": u["values"]} for u in label_updates])
    return ws


def save_df_to_worksheet(ws, df: pd.DataFrame):
    import gspread
    df = df.copy()
    df["Item"] = [i + 1 for i in range(len(df))]
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0)
    df["Subtotal"] = (df["Quantity"] * df["Unit Price"]).round(2)
    df = df.fillna("").astype(str)

    values = [SHEET_HEADERS] + df[SHEET_HEADERS].values.tolist()
    end_row = len(values)
    end_col = len(SHEET_HEADERS)
    cell_range = f"A1:{gspread.utils.rowcol_to_a1(end_row, end_col)}"

    for attempt in range(3):
        try:
            ws.batch_clear(["A1:G100"])
            ws.update(cell_range, values)
            return
        except gspread.exceptions.APIError as e:
            if attempt < 2:
                time.sleep(2)
            else:
                st.error("❌ Google Sheets API error while saving. Please wait and try again.")
                st.write(str(e))
                return
        except Exception as e:
            st.error(f"❌ Unexpected error while saving: {e}")
            return


def df_from_worksheet(ws) -> pd.DataFrame:
    for attempt in range(3):
        try:
            values = ws.get("A1:O100")
            if not values:
                return pd.DataFrame(columns=SHEET_HEADERS)

            raw_headers = values[0]
            data_rows = values[1:] if len(values) > 1 else []
            headers = raw_headers if len(raw_headers) == len(SHEET_HEADERS) else SHEET_HEADERS.copy()

            normalized = []
            for row in data_rows:
                row = row + [""] * (len(headers) - len(row)) if len(row) < len(headers) else row[:len(headers)]
                normalized.append(row)

            df = pd.DataFrame(normalized, columns=headers)
            for col in SHEET_HEADERS:
                if col not in df.columns:
                    df[col] = ""

            df = df[SHEET_HEADERS]
            df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
            df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0)
            df["Subtotal"] = df["Quantity"] * df["Unit Price"]
            return df

        except APIError as e:
            if attempt < 2:
                time.sleep(1.5)
                continue
            else:
                st.error("❌ Error reading Google Sheet — please wait and try again.")
                st.write(str(e))
                return pd.DataFrame(columns=SHEET_HEADERS)
        except Exception as e:
            st.error(f"❌ Unexpected error while reading sheet: {e}")
            return pd.DataFrame(columns=SHEET_HEADERS)
    return pd.DataFrame(columns=SHEET_HEADERS)


@st.cache_data(ttl=600)
def df_from_worksheet_cached(spreadsheet_key, worksheet_title):
    client = get_gspread_client()
    ss = client.open_by_key(spreadsheet_key)
    ws = ss.worksheet(worksheet_title)
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame(columns=SHEET_HEADERS)

    df = pd.DataFrame(values[1:], columns=values[0])
    for col in SHEET_HEADERS:
        if col not in df.columns:
            df[col] = ""
    df = df[SHEET_HEADERS]
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0)
    df["Subtotal"] = df["Quantity"] * df["Unit Price"]
    return df


def read_terms_from_ws(ws) -> dict:
    terms = {}
    for label, label_cell, value_cell in TERMS_LABELS:
        try:
            val = ws.acell(value_cell).value or ""
        except Exception:
            val = ""
        terms[label] = val
    return terms


def save_terms_to_ws(ws, terms: dict):
    updates = []
    for label, label_cell, value_cell in TERMS_LABELS:
        updates.append({"range": label_cell, "values": [[label]]})
        updates.append({"range": value_cell, "values": [[terms.get(label, "")]]})
    ws.batch_update([{"range": u["range"], "values": u["values"]} for u in updates])


def get_worksheet_with_retry(ss, project, retries=3, delay=1):
    for i in range(retries):
        try:
            return ss.worksheet(project)
        except gspread.exceptions.APIError:
            if i < retries - 1:
                time.sleep(delay)
            else:
                st.error(f"Failed to open worksheet '{project}'. Please try again in a few seconds.")
                st.session_state.page = "welcome"
                st.stop()


# ===============================================================
# Smart sheet update: only update changed rows (and append new)
# ===============================================================
def apply_sheet_updates(ws, old_df: pd.DataFrame, new_df: pd.DataFrame):
    """
    Update only rows that changed from old_df -> new_df.
    old_df/new_df are expected to have the same column order as SHEET_HEADERS.
    This function will:
     - Update each contiguous block of changed rows (A{start}:G{end}) with their new values.
     - If rows were appended, append them.
    """
    import gspread

    # Normalize data to strings for sheet update (gspread expects strings/numbers)
    old = old_df.fillna("").astype(str).reset_index(drop=True)
    new = new_df.fillna("").astype(str).reset_index(drop=True)

    old_len = len(old)
    new_len = len(new)

    # If sheet was empty before, do a full write (header + all rows)
    if old_len == 0 and new_len > 0:
        values = [SHEET_HEADERS] + new[SHEET_HEADERS].values.tolist()
        ws.batch_clear(["A1:G100"])
        ws.update(f"A1:{gspread.utils.rowcol_to_a1(len(values), len(SHEET_HEADERS))}", values)
        return

    # Find differing row indices up to min length
    min_len = min(old_len, new_len)
    changed = [i for i in range(min_len) if not old.loc[i, SHEET_HEADERS].equals(new.loc[i, SHEET_HEADERS])]

    # If there are changed rows within existing range, group contiguous blocks
    def contiguous_blocks(indices):
        if not indices:
            return []
        blocks = []
        start = indices[0]
        end = start
        for idx in indices[1:]:
            if idx == end + 1:
                end = idx
            else:
                blocks.append((start, end))
                start = idx
                end = idx
        blocks.append((start, end))
        return blocks

    blocks = contiguous_blocks(changed)

    for (start_idx, end_idx) in blocks:
        # sheet rows are offset by 2 (header at row 1)
        sheet_start_row = start_idx + 2
        sheet_end_row = end_idx + 2
        block = new.loc[start_idx:end_idx, SHEET_HEADERS]
        # Prepare values (no header)
        values = block.values.tolist()
        a1_start = f"A{sheet_start_row}"
        a1_end = f"{gspread.utils.rowcol_to_a1(sheet_end_row, len(SHEET_HEADERS))}"
        rng = f"{a1_start}:{a1_end}"
        ws.update(rng, values)

    # Handle appended rows
    if new_len > old_len:
        append_block = new.loc[old_len:new_len - 1, SHEET_HEADERS]
        if len(append_block) > 0:
            start_row = old_len + 2  # +2 because header row + 1-index
            end_row = new_len + 1
            values = append_block.values.tolist()
            a1_start = f"A{start_row}"
            a1_end = f"{gspread.utils.rowcol_to_a1(end_row, len(SHEET_HEADERS))}"
            rng = f"{a1_start}:{a1_end}"
            ws.update(rng, values)

    # If new_len < old_len (rows deleted), easiest and safest fallback is to rewrite whole sheet
    if new_len < old_len:
        values = [SHEET_HEADERS] + new[SHEET_HEADERS].values.tolist()
        ws.batch_clear(["A1:G100"])
        ws.update(f"A1:{gspread.utils.rowcol_to_a1(len(values), len(SHEET_HEADERS))}", values)


# ===============================================================
# PDF Generator
# ===============================================================
def generate_pdf(project_name, df, totals, terms, logo_path="90580b01-f401-47f5-aa43-48230c6c1bf2.jpeg"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()

    # Header
    try:
        logo = Image(logo_path, width=1.3 * inch, height=1.3 * inch)
    except Exception:
        logo = None

    company_title = Paragraph("<b>aNTS Technologies, Inc.</b>", ParagraphStyle('title', fontSize=16, leading=18))
    tagline = Paragraph("Solutions for a Small Planet", ParagraphStyle('tagline', fontSize=10, textColor=colors.gray))
    title = Paragraph("<b>PRICE QUOTE</b>", ParagraphStyle('title', fontSize=16, alignment=1))
    project = Paragraph(f"<b>Project:</b> {project_name}", ParagraphStyle('normal', fontSize=11))
    date_p = Paragraph(f"<b>Date:</b> {datetime.now().strftime('%B %d, %Y')}", ParagraphStyle('normal', fontSize=11))
    header_table = Table([[logo, [company_title, tagline, Spacer(1, 6), project, date_p]]], colWidths=[1.5 * inch, 4.5 * inch])
    header_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    elements += [header_table, Spacer(1, 15), title, Spacer(1, 15)]

    # Table
    data = [list(df.columns)] + df.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER")
    ]))
    elements += [table, Spacer(1, 15)]

    # Totals
    total_data = [
        ["Subtotal", f"₱ {totals['subtotal']:.2f}"],
        ["Discount", f"₱ {totals['discount']:.2f}"],
        ["VAT (12%)", f"₱ {totals['vat']:.2f}"],
        ["TOTAL", f"₱ {totals['total']:.2f}"]
    ]
    total_table = Table(total_data, colWidths=[4 * inch, 2.5 * inch])
    total_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.lightgreen)
    ]))
    elements += [total_table, Spacer(1, 20)]

    # Terms
    elements.append(Paragraph("<b>TERMS & CONDITIONS</b>", styles["Heading4"]))
    for key, value in terms.items():
        elements += [Paragraph(f"<b>{key}:</b> {value}", styles["Normal"]), Spacer(1, 4)]

    elements += [
        Spacer(1, 20),
        Paragraph("Prepared by:", styles["Normal"]),
        Spacer(1, 30),
        Paragraph("<b>_________________________</b>", styles["Normal"]),
        Paragraph("aNTS Technologies, Inc.", styles["Normal"]),
        Spacer(1, 10),
        Paragraph("<i>Thank you for doing business with us!</i>", styles["Italic"])
    ]

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ===============================================================
# UI Pages
# ===============================================================

st.title("📋 Project Quotation Manager")

if "page" not in st.session_state:
    st.session_state.page = "welcome"

if "spreadsheet" not in st.session_state:
    st.session_state.spreadsheet = open_spreadsheet()
ss = st.session_state.spreadsheet

# ----------------------
# Welcome Page
# ----------------------
if st.session_state.page == "welcome":
    st.header("Welcome!")
    if st.button("Create a Project Quote", key="btn_create_project_quote"):
        st.session_state.page = "create_project"
        st.rerun()

    st.subheader("Existing Projects")
    worksheets = [ws.title for ws in ss.worksheets()]
    search = st.text_input("Filter projects", key="filter_projects")
    filtered = [w for w in worksheets if search.lower() in w.lower()]

    for name in filtered:
        cols = st.columns([6, 1])
        with cols[0]:
            st.write(f"**{name}**")
        with cols[1]:
            if st.button("Open", key=f"open_{name}"):
                st.session_state.current_project = name
                st.session_state.page = "project"
                st.rerun()

# ----------------------
# Create Project Page
# ----------------------
elif st.session_state.page == "create_project":
    st.header("Create Project Quote")
    project_name = st.text_input("Enter project name (sheet name)", key="input_project_name")

    if st.button("Create", key="btn_create_project"):
        if not project_name:
            st.warning("Please enter a project name.")
        elif project_name in [ws.title for ws in ss.worksheets()]:
            st.error("Project already exists.")
        else:
            ws = worksheet_create_with_headers(ss, project_name)
            st.session_state.current_project = project_name
            st.session_state.page = "project"
            st.rerun()

    if st.button("Back", key="btn_back_to_welcome"):
        st.session_state.page = "welcome"
        st.rerun()

# ----------------------
# Project Page (Optimized)
# ----------------------
# --- Persistent editor state fix ---
data_key = f"project_df_{project}"

if data_key not in st.session_state:
    st.session_state[data_key] = df.copy()
    st.session_state[data_key].reset_index(drop=True, inplace=True)

df_ref = st.session_state[data_key]  # Reference to the live DataFrame

elif st.session_state.page == "project":
    st_autorefresh = st.empty()
    project = st.session_state.get("current_project")

    
    if "ws" not in st.session_state or st.session_state.get("ws_project") != project:
        st.session_state.ws = get_worksheet_with_retry(ss, project)
        st.session_state.ws_project = project

    ws = st.session_state.ws

# --- Load data only once per project ---
    if "project_df" not in st.session_state or st.session_state.get("project_df_project") != project:
        st.session_state.project_df = df_from_worksheet(ws)
        st.session_state.project_df_project = project
        st.session_state.last_edit_timestamp = 0.0
        st.session_state.is_saving_items = False
    
    df = st.session_state.project_df  # Always use this copy for display/edit

# --- Header Buttons ---
    col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])
    with col1:
        st.markdown(f"### 🧾 Project: {project}")

    with col2:
        if st.button("🔄 Refresh", key="refresh_sheet"):
            with st.spinner("Reloading data..."):
                new_df = df_from_worksheet(ws)
                st.session_state[f"project_df_{project}"] = new_df.copy()
                st.session_state.project_df = new_df.copy()
                st.session_state.unsaved_changes = False
            st.toast("✅ Data reloaded from Google Sheets", icon="🔄")
            st_autorefresh.empty()
    
    with col3:
        if st.button("➕ Row", key="add_top"):
            # Add row locally, will be saved after debounce
            new_row = pd.Series([len(df) + 1, "", "", 0, "", 0, 0], index=SHEET_HEADERS)
            df = df.append(new_row, ignore_index=True)
            st.session_state.project_df = df
            st.rerun()

    with col4:
        if st.button("⬅️ Back", key="back_top"):
            st.session_state.page = "welcome"
            st.rerun()

    with col5:
        export_pdf = st.button("📄 Export PDF", key="export_pdf")

    # Main Table - unique key per project
    # --- Initialize session data if not present ---
    if f"project_df_{project}" not in st.session_state:
        st.session_state[f"project_df_{project}"] = df.copy()
    if "last_edit_timestamp" not in st.session_state:
        st.session_state.last_edit_timestamp = 0.0
    if "is_saving_items" not in st.session_state:
        st.session_state.is_saving_items = False
    
# --- Always work on the session copy ---
    current_df = st.session_state[f"project_df_{project}"]

# --- Ensure project_df exists ---
    if "project_df" not in st.session_state:
        st.session_state.project_df = pd.DataFrame(columns=SHEET_HEADERS)

# --- Display editable table ---
    data_key = f"project_df_{project}"
    editor_key = f"editor_{project}" 

    if data_key not in st.session_state:
        st.session_state[data_key] = df.copy()
    
    edited = st.data_editor(
        df_ref,
        num_rows="dynamic",
        use_container_width=True,
        key=f"editor_{project}"
    )

# Only update if actual cell values changed
    if not edited.equals(df_ref):
    # Instead of replacing the DataFrame object,
    # update it *in place* (this prevents Streamlit reset)
        df_ref.loc[:, :] = edited
        st.session_state.project_df = df_ref
        st.session_state.unsaved_changes = True


# --- Show unsaved changes warning ---
    if st.session_state.get("unsaved_changes", False):
        st.warning("⚠️ You have unsaved edits. Click **💾 Save Changes** to commit them to Google Sheets.")

# --- Save button (manual batch save) ---
    save_col1, save_col2 = st.columns([5, 1])
    with save_col2:
        if st.button("💾 Save Changes", key="save_changes"):
            with st.spinner("Saving changes to Google Sheets..."):
                try:
                    new_df = st.session_state[f"project_df_{project}"].copy()
                    old_df = df_from_worksheet(ws)

                # Recalculate numeric columns and totals
                    new_df["Quantity"] = pd.to_numeric(new_df["Quantity"], errors="coerce").fillna(0)
                    new_df["Unit Price"] = pd.to_numeric(new_df["Unit Price"], errors="coerce").fillna(0)
                    new_df["Subtotal"] = (new_df["Quantity"] * new_df["Unit Price"]).round(2)
                    new_df["Item"] = [i + 1 for i in range(len(new_df))]

                # Apply only changed rows (batch update)
                    apply_sheet_updates(ws, old_df, new_df)

                    st.toast("✅ Changes saved to Google Sheets!", icon="💾")
                    st.session_state.unsaved_changes = False  # 🧠 mark clean state
                except Exception as e:
                    st.error(f"❌ Failed to save changes: {e}")


    # Totals
    total = edited["Subtotal"].sum()
    try:
        discount = float(ws.acell("J8").value or 0)
    except Exception:
        discount = 0.0
    vat = total * 0.12
    grand_total = total + vat - discount

    st.markdown("""
        <style>
        .big-metric { font-size: 28px; font-weight: 700; color: #222; }
        .highlight { font-size: 30px; font-weight: 800; color: #0a8754; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown(f"<div class='big-metric'>Total: ₱{total:,.2f}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='big-metric'>Discount: -₱{discount:,.2f}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='big-metric'>VAT (12%): ₱{vat:,.2f}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='highlight'>Grand Total: ₱{grand_total:,.2f}</div>", unsafe_allow_html=True)

    # Terms & Conditions
    st.markdown("---")
    st.subheader("Terms & Conditions")

    terms = read_terms_from_ws(ws)
    col1, col2 = st.columns(2)
    with col1:
        t_payment = st.text_input("Terms of payment", value=terms.get("Terms of payment", ""))
        t_delivery = st.text_input("Delivery", value=terms.get("Delivery", ""))
        t_discount = st.text_input("Discount", value=terms.get("Discount", ""))
    with col2:
        t_warranty = st.text_input("Warranty", value=terms.get("Warranty", ""))
        t_price = st.text_input("Price Validity", value=terms.get("Price Validity", ""))

    if st.button("Save Terms", key="save_terms"):
        save_terms_to_ws(ws, {
            "Terms of payment": t_payment,
            "Delivery": t_delivery,
            "Warranty": t_warranty,
            "Price Validity": t_price,
            "Discount": t_discount
        })
        st.success("Saved terms successfully.")

    if export_pdf:
        terms = read_terms_from_ws(ws)
        totals = {
            "subtotal": total,
            "discount": discount,
            "vat": vat,
            "total": grand_total
        }
        pdf_buffer = generate_pdf(project, edited, totals, terms)
        st.download_button(
            label="⬇️ Download Price Quote PDF",
            data=pdf_buffer,
            file_name=f"{project}_quotation.pdf",
            mime="application/pdf"
        )

# ===============================================================
# End of File
# ===============================================================





















