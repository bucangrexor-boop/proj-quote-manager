# Streamlit Project Quotation Manager
# File: app.py

import io
import json
import math
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
]

# ----------------------
# Helpers
# ----------------------

@st.cache_resource
def get_gspread_client():
    if GCP_SA_SECRET not in st.secrets or GSHEETS_KEY_SECRET not in st.secrets:
        st.error("Google secrets are missing. Add 'gcp_service_account' and 'gsheets_key' in Streamlit Secrets.")
        st.stop()
    creds_info = json.loads(st.secrets[GCP_SA_SECRET])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = service_account.Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(credentials)


def open_spreadsheet():
    client = get_gspread_client()
    key = st.secrets[GSHEETS_KEY_SECRET]
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
            ws.batch_clear(["A1:Z1000"])  # safer clear — only clears first 1000 rows
            ws.update(cell_range, values)
            return st.success("✅ Items saved to Google Sheet successfully!")
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

# ----------------------
# UI Pages
# ----------------------

st.title("📋 Project Quotation Manager")
if "page" not in st.session_state:
    st.session_state.page = "welcome"

ss = open_spreadsheet()
try:
    st.success(f"✅ Connected to Google Sheet: {ss.title}")
except Exception as e:
    st.error(f"❌ Failed to open Google Sheet: {type(e).__name__} - {e}")
    st.stop()

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
import time
import gspread
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


# ✅ start new block
if st.session_state.page == "project":
    project = st.session_state.get("current_project")
    st.header(f"Project: {project}")

    # ✅ safer worksheet opening
    ws = get_worksheet_with_retry(ss, project)

    df = df_from_worksheet(ws)
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    total = edited["Subtotal"].sum()
    st.metric("Total", f"₱{total:.2f}")

    col = st.columns(3)
    with col[0]:
        if st.button("Save Items"):
            save_df_to_worksheet(ws, edited)
            st.success("Items saved to Google Sheet.")
    with col[1]:
        if st.button("Add Row"):
            edited.loc[len(edited)] = [len(edited) + 1, "", "", 0, "", 0, 0]
    with col[2]:
        if st.button("Back"):
            st.session_state.page = "welcome"

    st.markdown("---")
    st.subheader("Terms & Conditions")

    terms = read_terms_from_ws(ws)

    col1, col2 = st.columns(2)
    with col1:
        t_payment = st.text_input("Terms of payment", value=terms.get("Terms of payment", ""))
        t_delivery = st.text_input("Delivery", value=terms.get("Delivery", ""))
    with col2:
        t_warranty = st.text_input("Warranty", value=terms.get("Warranty", ""))
        t_price = st.text_input("Price Validity", value=terms.get("Price Validity", ""))

    if st.button("Save Terms"):
        save_terms_to_ws(ws, {
            "Terms of payment": t_payment,
            "Delivery": t_delivery,
            "Warranty": t_warranty,
            "Price Validity": t_price,
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







