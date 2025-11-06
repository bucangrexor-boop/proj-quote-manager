# ===============================================================
# Streamlit Project Quotation Manager 
# File: app.py (Optimized Auto-Save)
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
from streamlit_autorefresh import st_autorefresh
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
# PDF Generator (unchanged)
# ===============================================================

def generate_pdf(project_name, df, totals, terms, logo_path="90580b01-f401-47f5-aa43-48230c6c1bf2.jpeg"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()

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

    data = [list(df.columns)] + df.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER")
    ]))
    elements += [table, Spacer(1, 15)]

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

ss = open_spreadsheet()

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
# Project Page
# ----------------------
elif st.session_state.page == "project":
    project = st.session_state.get("current_project")

    col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])
    with col1:
        st.markdown(f"### 🧾 Project: {project}")

    with col3:
        if st.button("➕ Row", key="add_top"):
            ws = get_worksheet_with_retry(ss, project)
            df = df_from_worksheet_cached(st.secrets[GSHEETS_KEY_SECRET], project)
            df.loc[len(df)] = [len(df) + 1, "", "", 0, "", 0, 0]
            save_df_to_worksheet(ws, df)
            st.rerun()

    with col4:
        if st.button("⬅️ Back", key="back_top"):
            st.session_state.page = "welcome"
            st.rerun()

    with col5:
        export_pdf = st.button("📄 Export PDF", key="export_pdf")

    ws = get_worksheet_with_retry(ss, project)
    df = df_from_worksheet_cached(st.secrets[GSHEETS_KEY_SECRET], project)

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="editor_main")

    # ------------------ Auto-Save Optimized ------------------
    if "last_items_df" not in st.session_state:
        st.session_state.last_items_df = df.to_dict()
    if "last_edit_time" not in st.session_state:
        st.session_state.last_edit_time = 0.0
    if "is_saving_items" not in st.session_state:
        st.session_state.is_saving_items = False
    if "unsaved_changes" not in st.session_state:
        st.session_state.unsaved_changes = False

    if edited.to_dict() != st.session_state.last_items_df:
        st.session_state.last_items_df = edited.to_dict()
        st.session_state.last_edit_time = time.time()
        st.session_state.unsaved_changes = True

    DEBOUNCE_DELAY = 15
    time_since_edit = time.time() - st.session_state.last_edit_time
    should_save = (
        st.session_state.unsaved_changes
        and time_since_edit > DEBOUNCE_DELAY
        and not st.session_state.is_saving_items
    )

    if should_save:
        st.session_state.is_saving_items = True
        st.session_state.unsaved_changes = False
        with st.spinner("💾 Auto-saving to Google Sheets..."):
            try:
                old_df = df_from_worksheet(ws)
                diffs = edited.compare(old_df, keep_shape=True, keep_equal=False)
                changed_rows = list(set(diffs.index.get_level_values(0)))
                batch_data = []

                if changed_rows:
                    for r in changed_rows:
                        row_data = edited.iloc[r].tolist()
                        row_num = r + 2
                        batch_data.append({"range": f"A{row_num}:G{row_num}", "values": [row_data]})
                    ws.batch_update(batch_data)
                    st.toast(f"✅ Saved {len(changed_rows)} row(s).", icon="💾")
                else:
                    st.toast("✅ No changes detected.", icon="💾")
            except Exception as e:
                st.warning(f"⚠️ Auto-save failed: {e}")
            finally:
                st.session_state.is_saving_items = False
                st.session_state.last_edit_time = 0.0

    status_placeholder = st.empty()
    if st.session_state.is_saving_items:
        status_placeholder.info("💾 Saving...")
    elif st.session_state.last_edit_time > 0 and time_since_edit <= DEBOUNCE_DELAY:
        remaining = int(DEBOUNCE_DELAY - time_since_edit)
        status_placeholder.caption(f"⌛ Pending auto-save in {remaining}s...")
    else:
        status_placeholder.caption("✅ All changes saved.")

    if st.session_state.last_edit_time > 0 and not st.session_state.is_saving_items:
        st_autorefresh(interval=1000, key="auto_save_refresh")

    # ----------------------
# End of file
