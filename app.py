# Streamlit Project Quotation Manager 
# File: app.py

import io
import json
import math
import time
import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account

st.set_page_config(page_title="Project Quotation Manager", layout="wide")

# ----------------------
# Configuration
# ----------------------

GSHEETS_KEY_SECRET = "gsheets_key"
GCP_SA_SECRET = "gcp_service_account"

SHEET_HEADERS = [
    "Item",
    "Part Number",
    "Description",
    "Quantity",
    "Unit",
    "Unit Price",
    "Subtotal",
]

TERMS_LABELS = [
    ("Terms of payment", "I2", "J2"),
    ("Delivery", "I3", "J3"),
    ("Warranty", "I4", "J4"),
    ("Price Validity", "I5", "J5"),
    ("Discount", "I8","J8")
]

# ----------------------
# Helpers
# ----------------------
from gspread.exceptions import APIError

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
        ss = client.open_by_key(key)
        return ss
    except Exception as e:
        st.error(f"❌ Could not open sheet: {type(e).__name__} - {e}")
        st.stop()
    return client.open_by_key(key)


def worksheet_create_with_headers(ss, title: str):
    ws = ss.add_worksheet(title=title, rows=100, cols=20)
    ws.update([SHEET_HEADERS])

    # Batch label updates (more efficient)
    label_updates = []
    for label, label_cell, _ in TERMS_LABELS:
        label_updates.append({"range": label_cell, "values": [[label]]})
    ws.batch_update([{"range": u["range"], "values": u["values"]} for u in label_updates])
    return ws

def save_df_to_worksheet(ws, df: pd.DataFrame):
    import time
    import gspread

    df = df.copy()
    df["Item"] = [i + 1 for i in range(len(df))]
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0)
    df["Subtotal"] = (df["Quantity"] * df["Unit Price"]).round(2)

    # Convert all to string (prevents invalid cell values)
    df = df.fillna("").astype(str)

    # Prepare data for update
    values = [SHEET_HEADERS] + df[SHEET_HEADERS].values.tolist()

    # Define update range (exact number of rows/cols)
    end_row = len(values)
    end_col = len(SHEET_HEADERS)
    cell_range = f"A1:{gspread.utils.rowcol_to_a1(end_row, end_col)}"

    # Retry logic for stability
    for attempt in range(3):
        try:
            ws.batch_clear(["A1:O100"])  # safer clear — only clears first 1000 rows
            ws.update(cell_range, values)
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
    from gspread.exceptions import APIError

    # Try a few times to avoid transient API failures
    for attempt in range(3):
        try:
            # Read a bounded range to avoid huge responses; adjust "Z" if you need more cols
            values = ws.get("A1:O100")

            # If sheet empty or no values, return empty DF with proper headers
            if not values or len(values) == 0:
                return pd.DataFrame(columns=SHEET_HEADERS)

            # First row considered headers (but may be malformed)
            raw_headers = values[0]
            data_rows = values[1:] if len(values) > 1 else []

            # If the header row doesn't match expected column count, fall back to SHEET_HEADERS
            if len(raw_headers) != len(SHEET_HEADERS):
                headers = SHEET_HEADERS.copy()
            else:
                headers = raw_headers

            # Normalize each row to match header length:
            normalized = []
            for row in data_rows:
                if len(row) < len(headers):
                    # pad short rows
                    row = row + [""] * (len(headers) - len(row))
                elif len(row) > len(headers):
                    # truncate long rows
                    row = row[: len(headers)]
                normalized.append(row)

            # Create DataFrame using the decided headers
            df = pd.DataFrame(normalized, columns=headers)

            # Ensure all expected columns exist (in case we used fallback headers)
            for col in SHEET_HEADERS:
                if col not in df.columns:
                    df[col] = ""

            # Reorder to canonical column order
            df = df[SHEET_HEADERS]

            # Convert numeric columns safely
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

    # Fallback (shouldn't reach here)
    return pd.DataFrame(columns=SHEET_HEADERS)
    
@st.cache_data(ttl=120)
def df_from_worksheet_cached(spreadsheet_key, worksheet_title):
    """Fetch and process worksheet data with caching."""
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

def safe_get_all_values(ws, retries=3, delay=2):
    for i in range(retries):
        try:
            return ws.get_all_values()
        except gspread.exceptions.APIError as e:
            if "Quota exceeded" in str(e) and i < retries - 1:
                time.sleep(delay)
            else:
                raise

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

# ----------------------
# UI Pages
# ----------------------

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
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

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
                
def generate_pdf(project_name, df, total, discount, vat, grand_total, terms):
    """Generate PDF bytes for download."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=2*cm, leftMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    elements = []

    # Title
    title = Paragraph(f"<b>Project Quotation</b><br/>{project_name}", styles["Title"])
    elements.append(title)
    elements.append(Spacer(1, 12))

    # Table Data
    table_data = [list(df.columns)] + df.values.tolist()
    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 12))

    # Totals
    totals_text = f"""
    <b>Total:</b> ₱{total:,.2f}<br/>
    <b>Discount:</b> -₱{discount:,.2f}<br/>
    <b>VAT (12%):</b> ₱{vat:,.2f}<br/>
    <b>Grand Total:</b> ₱{grand_total:,.2f}
    """
    elements.append(Paragraph(totals_text, styles["Normal"]))
    elements.append(Spacer(1, 12))

    # Terms & Conditions
    elements.append(Paragraph("<b>Terms & Conditions</b>", styles["Heading3"]))
    for k, v in terms.items():
        elements.append(Paragraph(f"<b>{k}:</b> {v}", styles["Normal"]))

    doc.build(elements)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

# ✅ start new block
if st.session_state.page == "project":
    project = st.session_state.get("current_project")

   # === Top header row with project title + buttons ===
    col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])

    with col1:
        st.markdown(f"### 🧾 Project: {project}")
    
    with col2:
        if st.button("💾 Save", key="save_top"):
            ws = get_worksheet_with_retry(ss, project)
            df_to_save = df_from_worksheet_cached(st.secrets[GSHEETS_KEY_SECRET], project)
            save_df_to_worksheet(ws, df_to_save)
            st.success("Items saved to Google Sheet.")

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
        if st.button("📄 Export PDF", key="export_pdf"):
            ws = get_worksheet_with_retry(ss, project)
            df = df_from_worksheet_cached(st.secrets[GSHEETS_KEY_SECRET], project)
            
    # ✅ safer worksheet opening (keep this below header)
    ws = get_worksheet_with_retry(ss, project)
        df = df_from_worksheet_cached(st.secrets[GSHEETS_KEY_SECRET], project)
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        total = df["Subtotal"].sum()
        try:
            discount = float(ws.acell("J6").value or 0)
        except Exception:
            discount = 0.0
        vat = total * 0.12
        grand_total = total + vat - discount
        terms = read_terms_from_ws(ws)
        
        pdf_bytes = generate_pdf(project, df, total, discount, vat, grand_total, terms)
        st.download_button(
            label="⬇️ Download PDF",
            data=pdf_bytes,
            file_name=f"{project}_quotation.pdf",
            mime="application/pdf"
        )


    # ✅ safer worksheet opening (keep this below header)
    ws = get_worksheet_with_retry(ss, project)

    df = df_from_worksheet_cached(st.secrets[GSHEETS_KEY_SECRET], project)
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    total = edited["Subtotal"].sum()
    try:
        discount = float(ws.acell("J6").value or 0)
    except Exception:
        discount = 0.0

    vat = total * 0.12
    grand_total = total + vat - discount

    st.markdown("""
        <style>
        .big-metric {
            font-size: 28px;
            font-weight: 700;
            color: #222; /* dark text */
        }
        .highlight {
            font-size: 30px;
            font-weight: 800;
            color: #0a8754; /* green color for emphasis */
        }
        </style>
    """, unsafe_allow_html=True)

    st.markdown(f"<div class='big-metric'>Total: ₱{total:,.2f}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='big-metric'>Discount: -₱{discount:,.2f}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='big-metric'>VAT (12%): ₱{vat:,.2f}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='highlight'>Grand Total: ₱{grand_total:,.2f}</div>", unsafe_allow_html=True)

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

    if st.button("Save Terms"):
        save_terms_to_ws(ws, {
            "Terms of payment": t_payment,
            "Delivery": t_delivery,
            "Warranty": t_warranty,
            "Price Validity": t_price,
            "Discount": t_discount
        })
        st.success("Saved terms successfully.")
# ----------------------
# requirements.txt
# ----------------------
# streamlit
# gspread
# google-auth
# pandas

# ----------------------
# README.md
# ----------------------
# # Project Quotation Manager
#
# A Streamlit app connected to Google Sheets to manage project quotations.
#
# ## Features
# - Create, edit, and delete project quotation sheets
# - Auto-generated Item numbers
# - Dynamic subtotal calculation (Quantity × Unit Price)
# - Editable Terms of Payment, Delivery, Warranty, and Price Validity
# - Data saved directly to Google Sheets
#
# ## Setup
# 1. Create a Google Service Account with Sheets API enabled.
# 2. Share your Google Sheet with the service account email.
# 3. In Streamlit Cloud, add these to **Secrets**:
# ```toml
# [secrets]
# gsheets_key = "your_spreadsheet_key"
# gcp_service_account = <<EOF
# { ...entire JSON... }
# EOF
# ```
# 4. Deploy on [Streamlit Community Cloud](https://streamlit.io/cloud).
# 5. Run the app and manage quotations easily!


































