# ----------------------
# Imports
# ----------------------
import io
import os
import json
import time
import requests
import numpy as np
import pandas as pd
import streamlit as st
import gspread
from io import BytesIO
from datetime import datetime
from google.oauth2 import service_account
from gspread.exceptions import APIError
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepInFrame)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image as RLImage

# Streamlit Configuration
st.set_page_config(page_title="Project Quotation Manager", layout="wide")

# Constants
GSHEETS_KEY_SECRET = "gsheets_key"
GCP_SA_SECRET = "gcp_service_account"

SHEET_HEADERS = [
    "Item", "Part Number", "Description", "Qty", "Unit", "Unit Price", "Subtotal"
]

TERMS_LABELS = [
    ("TERMS OF PAYMENT", "I2", "J2"),
    ("DELIVERY", "I3", "J3"),
    ("WARRANTY", "I4", "J4"),
    ("PRICE VALIDITY", "I5", "J5"),
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
    # write initial terms label cells so sheet is more user-friendly
    label_updates = [{"range": label_cell, "values": [[label]]} for label, label_cell, _ in TERMS_LABELS]
    ws.batch_update([{"range": u["range"], "values": u["values"]} for u in label_updates])
    return ws


def df_from_worksheet(ws) -> pd.DataFrame:
    for attempt in range(3):
        try:
            values = ws.get("A1:O200")
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
            df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
            df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0)
            df["Subtotal"] = (df["Qty"] * df["Unit Price"]).round(2)
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


def apply_sheet_updates(ws, old_df: pd.DataFrame, new_df: pd.DataFrame):
    def get_last_data_row(ws):
        try:
            colA = ws.col_values(1)
        except Exception:
            return len(old_df) + 1
        last = 0
        for i, val in enumerate(colA, start=1):
            if val is not None and str(val).strip() != "":
                last = i
        return last

    old = old_df.replace({np.nan: None}).reset_index(drop=True)
    new = new_df.replace({np.nan: None}).reset_index(drop=True)

    old_len = len(old)
    new_len = len(new)

    if old_len == 0 and new_len > 0:
        values = [SHEET_HEADERS] + new[SHEET_HEADERS].fillna("").astype(str).values.tolist()
        try:
            ws.batch_clear(["A1:G200"])
            ws.update(f"A1:G{len(values)}", values)
        except Exception as e:
            st.error(f"❌ Error writing full sheet: {e}")
        return

    min_len = min(old_len, new_len)
    changed_rows = []
    for i in range(min_len):
        if not new.loc[i, SHEET_HEADERS].equals(old.loc[i, SHEET_HEADERS]):
            changed_rows.append(i)

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

    blocks = contiguous_blocks(changed_rows)

    for (start_idx, end_idx) in blocks:
        sheet_start_row = start_idx + 2
        sheet_end_row = end_idx + 2
        block_df = new.loc[start_idx:end_idx, SHEET_HEADERS].fillna("").astype(str)
        values = block_df.values.tolist()
        try:
            ws.update(f"A{sheet_start_row}:G{sheet_end_row}", values)
        except Exception as e:
            st.error(f"❌ Error updating rows {sheet_start_row}-{sheet_end_row}: {e}")

    if new_len > old_len:
        last_data_row = get_last_data_row(ws)
        start_index = old_len
        append_block = new.loc[start_index:new_len - 1, SHEET_HEADERS].fillna("").astype(str).values.tolist()
        if append_block:
            start_row = last_data_row + 1
            end_row = start_row + len(append_block) - 1
            try:
                st.write(f"Appending rows at sheet rows {start_row}..{end_row} -> {len(append_block)} rows")
                ws.update(f"A{start_row}:G{end_row}", append_block)
            except Exception as e:
                st.error(f"❌ Error appending rows {start_row}-{end_row}: {e}")

    if new_len < old_len:
        values = [SHEET_HEADERS] + new[SHEET_HEADERS].fillna("").astype(str).values.tolist()
        try:
            ws.batch_clear(["A1:G200"])
            ws.update(f"A1:G{len(values)}", values)
        except Exception as e:
            st.error(f"❌ Error rewriting entire sheet: {e}")


def save_totals_to_ws(ws, total, vat, grand_total):
    updates = [
        {"range": "I9", "values": [["Total"]]},
        {"range": "J9", "values": [[str(total)]]},

        {"range": "I10", "values": [["VAT (12%)"]]},
        {"range": "J10", "values": [[str(vat)]]},

        {"range": "I11", "values": [["Grand Total"]]},
        {"range": "J11", "values": [[str(grand_total)]]},
    ]
    ws.batch_update(updates)

# ===============================================================
# PDF Generator
# ===============================================================
# Register Fonts
# Build font path relative to app.py
BASE_DIR = os.path.dirname(__file__)
FONT_DIR = os.path.join(BASE_DIR, "fonts")

def font(file):
    return os.path.join(FONT_DIR, file)
pdfmetrics.registerFont(TTFont('Arial', font('ARIAL.TTF')))
pdfmetrics.registerFont(TTFont('Arial-Bold', font('ARIALBD.TTF')))
pdfmetrics.registerFont(TTFont('Arial-Narrow', font('ARIALN.TTF')))
pdfmetrics.registerFont(TTFont('Calibri', font('CALIBRI.TTF')))
pdfmetrics.registerFont(TTFont('Calibri-Bold', font('CALIBRIB.TTF')))


def generate_pdf(project_name, df, totals, terms, client_info=None,
                 left_logo_path=None, right_logo_path=None):

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=54, leftMargin=54,
                            topMargin=72, bottomMargin=72)

    elements = []
    styles = getSampleStyleSheet()

    # -----------------------
    # Custom Styles
    # -----------------------
    price_quote_style = ParagraphStyle(
        "PriceQuote",
        fontName="Calibri-Bold",
        fontSize=8,
        alignment=1  # CENTER
    )
    ref_style = ParagraphStyle(
        "RefStyle",
        fontName="Arial-Narrow",
        fontSize=7,
        alignment=1  # CENTER
    )
    title_style = ParagraphStyle(
        "TitleStyle",
        fontName="Arial-Bold",
        fontSize=8,
        alignment=0  # LEFT
    )
    office_style = ParagraphStyle(
        "OfficeStyle",
        fontName="Arial-Bold",
        fontSize=7,
        alignment=0,
        leading=10,
        spaceAfter = 0
    )
    normal_style = ParagraphStyle(
        "NormalStyle",
        fontName="Arial",
        fontSize=7,
        alignment=0,
        leading=10,
        spaceAfter = 0
    )
    table_header_style = ParagraphStyle(
        "TableHeader",
        fontName="Arial-Bold",
        fontSize=7,
        alignment=1  # CENTER
    )

    # -----------------------
    # Load logos safely
    # -----------------------
    def load_logo(url_or_path, width=None, height=None):
        try:
            if url_or_path.startswith("http"):
                response = requests.get(url_or_path)
                response.raise_for_status()
                image_data = BytesIO(response.content)
                img = RLImage(image_data, width=width, height=height)  # height=None keeps aspect ratio
            else:
                img = RLImage(url_or_path, width=width, height=height)
            img.hAlign = 'LEFT'
            return img
        except Exception as e:
            print("Failed to load logo:", e)
        # Use a transparent spacer instead of empty string
            return Spacer(width or 50 , height or 20)

    left_logo = load_logo("https://raw.githubusercontent.com/bucangrexor-boop/proj-quote-manager/main/assets/logoants.png",
                        width = 121.32 * 0.75,
                        height = 50 * 0.75)
    right_logo = load_logo("https://raw.githubusercontent.com/bucangrexor-boop/proj-quote-manager/main/assets/antslogo2.png",
                        width= 201.89 * 0.75,
                        height= 17.33 * 0.75)    
    # -----------------------
    # Header (logos)
    # -----------------------
    header_table = Table(
        [[left_logo, right_logo]],
        colWidths=[3*inch, 3*inch]
    )
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0)
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 0))

    # -----------------------
    # Price Quote Title
    # -----------------------
    elements.append(Paragraph("P R I C E   Q U O T E", price_quote_style))
    elements.append(Spacer(1, 0))

    # -----------------------
    # Ref No.
    # -----------------------
    elements.append(Paragraph(f"Ref No. {project_name}", ref_style))
    elements.append(Spacer(1, 2))

    # -----------------------
    # Date
    # -----------------------
    date_str = datetime.now().strftime("%d-%b-%y")
    date_line = Paragraph(
        f'<para alignment="right"><b>Date</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{date_str}</para>',
        ref_style
    )
    elements.append(date_line)
    elements.append(Spacer(1, 12))

    # -----------------------
    # Client Info
    # -----------------------
    if client_info:
        elements.append(Paragraph(f"<b>{client_info.get('Title', '')}</b>", title_style))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(client_info.get('Office', ''), office_style))
        elements.append(Paragraph(client_info.get("Company", ""), normal_style))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Dear Sir:", normal_style))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(client_info.get("Message", ""), normal_style))
        elements.append(Spacer(1, 10))  
    # -----------------------
    # TABLE DATA PREPARATION
    # -----------------------
    table_data = []
    header = df.columns.tolist()
    table_data.append(header)

    styles = getSampleStyleSheet()
    wrap_style = styles["BodyText"]
    wrap_style.fontName = "Arial"
    wrap_style.fontSize = 7
    wrap_style.leading = 9  # make slightly bigger than font size for line spacing

    for i, row in df.reset_index(drop=True).iterrows():
    # --- SAFE ITEM NUMBER HANDLING ---
        raw_item = row.get("Item", None)
        if pd.isna(raw_item) or raw_item is None or str(raw_item).strip() == "":
            item_no = i + 1
        else:
            try:
                item_no = int(float(str(raw_item).strip()))
            except:
                item_no = i + 1
    # --------------------------------

    # WRAP LONG DESCRIPTION
        description = Paragraph(str(row.get("Description", "") or ""), wrap_style)

    # Qty → FORCE INTEGER (NO DECIMALS)
        qty_raw = row.get("Qty", 0)
        try:
            qty = int(float(qty_raw))
        except:
            qty = 0

    # Append row to table
        table_data.append([
            item_no,
            str(row.get("Part Number", "") or ""),
            description,
            qty,  # integer-only Qty
            str(row.get("Unit", "") or ""),
            f"{row.get('Unit Price', 0):.2f}",
            f"{row.get('Subtotal', 0):.2f}",
        ])

# ----------------------------------------------------
# COLUMN WIDTHS BASED ON AVAILABLE PAGE SPACE
# ----------------------------------------------------
    PAGE_WIDTH, PAGE_HEIGHT = A4
    available_width = PAGE_WIDTH - (doc.leftMargin + doc.rightMargin)
    proportions = [0.0748, 0.1247, 0.4423, 0.0399, 0.0474, 0.1338, 0.1371]
    total_prop = sum(proportions)
    proportions = [p / total_prop for p in proportions]
    col_widths = [available_width * p for p in proportions]

# ----------------------------------------------------
# CREATE MAIN TABLE WITH WRAPPING & AUTO-HEIGHT
# ----------------------------------------------------
    table = Table(table_data, colWidths=col_widths, repeatRows=1, rowHeights=None)

    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),

    # Header
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),

    # Body
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("VALIGN", (0, 1), (-1, -1), "MIDDLE"),

    # Alignment per column
        ("ALIGN", (0, 1), (1, -1), "CENTER"),
        ("ALIGN", (2, 1), (2, -1), "LEFT"),
        ("ALIGN", (3, 1), (3, -1), "RIGHT"),
        ("ALIGN", (4, 1), (4, -1), "CENTER"),
        ("ALIGN", (5, 1), (6, -1), "RIGHT"),

    # Padding proportional to font
        ("TOPPADDING", (0, 1), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, 0), 0),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 0),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 0))

    # -----------------------
    # Totals Table (aligned over last 2 columns)
    # ----------------------

    totals_data = [
        ["Subtotal", f"₱ {totals['subtotal']:,.2f}"],
        ["Discount", f"₱ {totals['discount']:,.2f}"],
        ["VAT (12%)", f"₱ {totals['vat']:,.2f}"],
        ["TOTAL", f"₱ {totals['total']:,.2f}"]
    ]

# Create totals table
    totals_table = Table(
        totals_data,
        colWidths=[100, 120],      # ← Make these narrower so table stays right-side
        hAlign='RIGHT'             # ← THIS IS THE IMPORTANT LINE
    )

    totals_table.setStyle(TableStyle([
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('FONTNAME', (0,0), (-1,-2), 'Helvetica'),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgreen),
        ('LINEABOVE', (0,-1), (-1,-1), 1, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.grey),
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
    ]))

# Add a small space before totals appear
    story.append(Spacer(1, 12))

# Add to the story
    story.append(totals_table)


    # -----------------------
    # Terms & Conditions
    # -----------------------
    for k, v in terms.items():
        if k == "Discount":
            continue
        elements.append(Paragraph(f"<b>{k}:</b> {v}", normal_style))
    elements.append(Spacer(1, 12))
    # -----------------------
    # Sign-off
    # -----------------------
    elements.append(Paragraph("Thank you for doing business with us!", normal_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Respectfully yours,", normal_style))
    elements.append(Spacer(1, 30))

    if client_info:
        elements.append(Paragraph(client_info.get("Edited By", ""), normal_style))
        elements.append(Paragraph("Ants Technologies, Inc.", ref_style))
    # -----------------------
    # Build PDF
    # -----------------------
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

    if st.button("Back", key="btn_back_to_welcome"):
        st.session_state.page = "welcome"
        st.rerun()

#-------------Project UI--------- 

elif st.session_state.page == "project":
    project = st.session_state.get("current_project")

    # Get worksheet
    if "ws" not in st.session_state or st.session_state.get("ws_project") != project:
        st.session_state.ws = get_worksheet_with_retry(ss, project)
        st.session_state.ws_project = project
    ws = st.session_state.ws

    # Session DF key
    session_key = f"project_df_{project}"
    if session_key not in st.session_state:
        st.session_state[session_key] = df_from_worksheet(ws).reset_index(drop=True)

    st.markdown(f"### 🧾 Project: {project}")

    # Header buttons
    col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])
    with col2:
        if st.button("⬅️ Back", key="back_top"):
            st.session_state.page = "welcome"

    with col4:
        export_pdf = st.button("📄 Export PDF", key="export_pdf")

    # -----------------------
    # Editable Table (inside form to prevent flicker)
    # -----------------------
    with st.form("save_project_form"):
        edited_df = st.data_editor(
            st.session_state[session_key],
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_{project}_form"
        )

        submit = st.form_submit_button("💾 Save Changes")
        if submit:
            with st.spinner("Saving changes..."):
                try:
                    new_df = edited_df.copy()
                    for col in ["Qty", "Unit Price"]:
                        new_df[col] = pd.to_numeric(new_df[col], errors="coerce").fillna(0)
                    new_df["Subtotal"] = (new_df["Qty"] * new_df["Unit Price"]).round(2)
                    new_df["Item"] = range(1, len(new_df) + 1)

                    old_df = df_from_worksheet(ws).reset_index(drop=True)
                    apply_sheet_updates(ws, old_df, new_df)

                    # Compute totals
                    total = new_df["Subtotal"].sum()
                    try:
                        discount = float(ws.acell("J8").value or 0)
                    except:
                        discount = 0.0
                    vat = total * 0.12
                    grand_total = total + vat - discount

                    save_totals_to_ws(ws, total, vat, grand_total)

                    st.session_state[session_key] = new_df.copy()
                    st.session_state[session_key + "_totals"] = (total, discount, vat, grand_total)
                    st.success("✅ Changes saved to Google Sheets!")
                except Exception as e:
                    st.error(f"❌ Failed to save changes: {e}")

    # -----------------------
    # Totals Display
    # -----------------------
    if st.session_state.get(session_key + "_totals"):
        total, discount, vat, grand_total = st.session_state[session_key + "_totals"]
        st.markdown(f"<div class='big-metric'>Total: ₱{total:,.2f}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='big-metric'>Discount: -₱{discount:,.2f}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='big-metric'>VAT (12%): ₱{vat:,.2f}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='highlight'>Grand Total: ₱{grand_total:,.2f}</div>", unsafe_allow_html=True)

    # -----------------------
    # Terms & Client Info (separate forms)
    # -----------------------
    st.markdown("---")
    st.subheader("Terms & Conditions")
    terms = read_terms_from_ws(ws)
    with st.form("terms_form"):
        col1, col2 = st.columns(2)
        with col1:
            t_payment = st.text_input("TERMS OF PAYMENT", value=terms.get("TERMS OF PAYMENT", ""))
            t_DELIVERY = st.text_input("DELIVERY", value=terms.get("DELIVERY", ""))
            t_discount = st.text_input("Discount", value=terms.get("Discount", ""))
        with col2:
            t_WARRANTY = st.text_input("WARRANTY", value=terms.get("WARRANTY", ""))
            t_price = st.text_input("PRICE VALIDITY", value=terms.get("PRICE VALIDITY", ""))

        submit_terms = st.form_submit_button("Save Terms")
        if submit_terms:
            save_terms_to_ws(ws, {
                "TERMS OF PAYMENT": t_payment,
                "DELIVERY": t_DELIVERY,
                "WARRANTY": t_WARRANTY,
                "PRICE VALIDITY": t_price,
                "Discount": t_discount
            })
            st.success("Saved terms successfully.")

    st.markdown("---")
    st.subheader("Client Information")
    client_fields = ["Title", "Office", "Company", "Message", "Edited By"]
    saved_values = {}
    for i, field in enumerate(client_fields, start=14):
        try:
            saved_values[field] = ws.acell(f"J{i}").value or ""
        except:
            saved_values[field] = ""

    with st.form("client_form"):
        colA, colB = st.columns(2)
        with colA:
            title_input = st.text_input("Title", value=saved_values["Title"])
            office_input = st.text_input("Office", value=saved_values["Office"])
        with colB:
            company_input = st.text_input("Company", value=saved_values["Company"])
            editedby_input = st.text_input("Edited By", value=saved_values["Edited By"])
        message_input = st.text_area("Message", value=saved_values["Message"], height=120)

        submit_client = st.form_submit_button("Save Client Info")
        if submit_client:
            updates = [
                {"range": f"I{i}", "values": [[field]]} for i, field in enumerate(client_fields, start=14)
            ] + [
                {"range": f"J{i}", "values": [[value]]} for i, value in enumerate(
                    [title_input, office_input, company_input, message_input, editedby_input], start=14)
            ]
            ws.batch_update(updates)
            st.success("Client information saved!")

    # -----------------------
    # Export PDF
    # -----------------------
    if export_pdf:
        try:
            sheet_df = st.session_state[session_key]
            terms = read_terms_from_ws(ws)
            totals = {
                "subtotal": sheet_df["Subtotal"].sum(),
                "discount": float(ws.acell("J8").value or 0),
                "vat": sheet_df["Subtotal"].sum() * 0.12,
                "total": sheet_df["Subtotal"].sum() + (sheet_df["Subtotal"].sum() * 0.12) - float(ws.acell("J8").value or 0)
            }
            client_info = {
                "Title": ws.acell("J14").value or "",
                "Office": ws.acell("J15").value or "",
                "Company": ws.acell("J16").value or "",
                "Message": ws.acell("J17").value or "",
                "Edited By": ws.acell("J18").value or ""
            }
            pdf_buffer = generate_pdf(project, sheet_df, totals, terms, client_info=client_info)
            st.download_button(
                label="⬇️ Download Price Quote PDF",
                data=pdf_buffer,
                file_name=f"{project}_quotation.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"❌ Failed to generate PDF: {e}")





# ===============================================================
# End of File
# ===============================================================







