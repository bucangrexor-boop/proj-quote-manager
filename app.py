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
    df = df.copy()
    df["Item"] = [i + 1 for i in range(len(df))]
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0)
    df["Subtotal"] = (df["Quantity"] * df["Unit Price"]).round(2)
    values = [SHEET_HEADERS] + df[SHEET_HEADERS].astype(str).values.tolist()
    ws.clear()
    ws.update(values)


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
    for label, label_cell, value_cell in TERMS_LABELS:
        ws.update(label_cell, label)
        ws.update(value_cell, terms.get(label, ""))

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
elif st.session_state.page == "project":
    project = st.session_state.get("current_project")
    st.header(f"Project: {project}")
    ws = ss.worksheet(project)

    df = df_from_worksheet(ws)
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="data_editor_items")
    total = edited["Subtotal"].sum()
    st.metric("Total", f"₱{total:,.2f}")

    col = st.columns(3)
    with col[0]:
        if st.button("Save Items", key="btn_save_items"):
            save_df_to_worksheet(ws, edited)
            st.success("Items saved to Google Sheet.")
    with col[1]:
        if st.button("Add Row", key="btn_add_row"):
            edited.loc[len(edited)] = [len(edited) + 1, "", "", 0, "", 0, 0]
    with col[2]:
        if st.button("Back", key="btn_back_from_project"):
            st.session_state.page = "welcome"
            st.rerun()

    st.markdown("---")
    st.subheader("Terms & Conditions")
    terms = read_terms_from_ws(ws)

    col1, col2 = st.columns(2)
    with col1:
        t_payment = st.text_input("Terms of payment", value=terms.get("Terms of payment", ""), key="input_terms_payment")
        t_delivery = st.text_input("Delivery", value=terms.get("Delivery", ""), key="input_terms_delivery")
    with col2:
        t_warranty = st.text_input("Warranty", value=terms.get("Warranty", ""), key="input_terms_warranty")
        t_price = st.text_input("Price Validity", value=terms.get("Price Validity", ""), key="input_terms_price_validity")

    if st.button("Save Terms", key="btn_save_terms"):
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
