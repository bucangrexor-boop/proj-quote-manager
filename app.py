# ----------------------
# Imports
# ----------------------
import io
import json
import time
import numpy as np
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

# Streamlit Configuration
st.set_page_config(page_title="Project Quotation Manager", layout="wide")

# Constants
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
            df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
            df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0)
            df["Subtotal"] = (df["Quantity"] * df["Unit Price"]).round(2)
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

    try:
        st.write(f"ok")
    except Exception:
        pass

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
def generate_pdf(project_name, df, totals, terms, 
                 left_logo_path=None, right_logo_path=None):

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=30, leftMargin=30,
                            topMargin=40, bottomMargin=30)

    elements = []
    styles = getSampleStyleSheet()

    # Load logos (safe-loading)
    def load_logo(path, width=1.8*inch):
        if not path:
            return ""
        try:
            return Image(path, width=width, preserveAspectRatio=True, hAlign='LEFT')
        except:
            return ""

    left_logo = load_logo(left_logo_path)
    right_logo = load_logo(right_logo_path)

    # ------------------------------------------
    # HEADER (2 columns)
    # ------------------------------------------
    header_table = Table(
        [[left_logo, right_logo]],
        colWidths=[3*inch, 3*inch]
    )
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10)
    ]))

    elements.append(header_table)
    elements.append(Spacer(1, 12))

    # ------------------------------------------
    # CENTER TITLE BLOCK
    # ------------------------------------------
    title = Paragraph("<b>PRICE QUOTE</b>", ParagraphStyle(
        'TitleCenter', fontSize=18, alignment=1, leading=22
    ))

    ref_no = Paragraph(
        f"<b>Ref No:</b> {project_name}",
        ParagraphStyle('Ref', fontSize=11, alignment=1)
    )

    elements.extend([title, Spacer(1, 4), ref_no, Spacer(1, 20)])

    # ------------------------------------------
    # MAIN TABLE
    # ------------------------------------------
    data = [list(df.columns)] + df.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 15))

    # ------------------------------------------
    # TOTALS TABLE
    # ------------------------------------------
    total_data = [
        ["Subtotal", f"₱ {totals['subtotal']:.2f}"],
        ["Discount", f"₱ {totals['discount']:.2f}"],
        ["VAT (12%)", f"₱ {totals['vat']:.2f}"],
        ["TOTAL", f"₱ {totals['total']:.2f}"],
    ]
    totals_table = Table(total_data, colWidths=[4*inch, 2.3*inch])
    totals_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.lightgreen),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]))

    elements.append(totals_table)
    elements.append(Spacer(1, 20))

    # ------------------------------------------
    # TERMS & CONDITIONS
    # ------------------------------------------
    elements.append(Paragraph("<b>TERMS & CONDITIONS</b>", styles["Heading4"]))
    for k, v in terms.items():
        elements.append(Paragraph(f"<b>{k}:</b> {v}", styles["Normal"]))
        elements.append(Spacer(1, 4))

    # ------------------------------------------
    # SIGN-OFF
    # ------------------------------------------
    elements.extend([
        Spacer(1, 20),
        Paragraph("Prepared by:", styles["Normal"]),
        Spacer(1, 30),
        Paragraph("<b>_________________________</b>", styles["Normal"]),
        Paragraph("aNTS Technologies, Inc.", styles["Normal"]),
        Spacer(1, 15),
        Paragraph("<i>Thank you for doing business with us!</i>", styles["Italic"])
    ])

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
# Project Page (Optimized) - REPLACEMENT BLOCK
# ----------------------
elif st.session_state.page == "project":
    refresh_container = st.empty()
    project = st.session_state.get("current_project")

    # Get worksheet
    if "ws" not in st.session_state or st.session_state.get("ws_project") != project:
        st.session_state.ws = get_worksheet_with_retry(ss, project)
        st.session_state.ws_project = project
    ws = st.session_state.ws

    # Initialize / load per-project DataFrame into session state (single source of truth)
    session_key = f"project_df_{project}"
    if session_key not in st.session_state:
        st.session_state[session_key] = df_from_worksheet(ws).reset_index(drop=True)
        st.session_state[f"{session_key}_loaded"] = True

    # df_ref is the live DataFrame the editor uses
    df_clean = st.session_state[session_key].copy()

    df_clean["Quantity"] = pd.to_numeric(df_clean["Quantity"], errors="coerce").fillna(0)
    df_clean["Unit Price"] = pd.to_numeric(df_clean["Unit Price"], errors="coerce").fillna(0)
    df_clean["Subtotal"] = (df_clean["Quantity"] * df_clean["Unit Price"]).round(2)

    # Save cleaned version before editor → prevents "first key disappears"
    st.session_state[session_key] = df_clean.copy()

    # Header Buttons
    col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])
    with col1:
        st.markdown(f"### 🧾 Project: {project}")

    with col2:
        if st.button("🔄 Refresh", key="refresh_sheet"):
            with st.spinner("Reloading data..."):
                reloaded = df_from_worksheet(ws)
                st.session_state[session_key] = reloaded.reset_index(drop=True)
                df_ref = st.session_state[session_key]
                st.session_state.unsaved_changes = False
            st.toast("✅ Data reloaded from Google Sheets", icon="🔄")
            refresh_container.empty()

    # Note: removed the + Row button per user's request

    with col4:
        if st.button("⬅️ Back", key="back_top"):
            st.session_state.page = "welcome"
            st.rerun()

    with col5:
        export_pdf = st.button("📄 Export PDF", key="export_pdf")

    # This is the ONE source of truth for the editor
    current_df = st.session_state[session_key]
    
    # ---------------------------------------------------
    # ✅ Data editor — edit the session DF directly
    # ---------------------------------------------------
    edited_df = st.data_editor(
        current_df,
        num_rows="dynamic",
        use_container_width=True,
        key=f"editor_{project}"
    )

    # ---------------------------------------------------
    # ✅ Only update session DF if something changed
    # ---------------------------------------------------
    if not edited_df.equals(current_df):
        st.session_state[session_key] = edited_df.copy()
        st.session_state.unsaved_changes = True

    # ---------------------------------------------------
    # ✅ Warning for unsaved edits
    # ---------------------------------------------------
    if st.session_state.get("unsaved_changes", False):
        st.warning("⚠️ You have unsaved edits. Click **💾 Save Changes** to commit them to Google Sheets.")
        
    # Always work on a copy to avoid triggering Streamlit reruns
    current_df = st.session_state[session_key].copy()

# Make sure numeric fields are correct
    current_df["Quantity"] = pd.to_numeric(current_df["Quantity"], errors="coerce").fillna(0).astype(float)
    current_df["Unit Price"] = pd.to_numeric(current_df["Unit Price"], errors="coerce").fillna(0).astype(float)
    current_df["Subtotal"] = (current_df["Quantity"] * current_df["Unit Price"]).round(2)
# Save the computed version back to session state (won’t break the editor)
    st.session_state[session_key] = current_df
    # --- Totals --- (compute BEFORE Save button so they exist when saving)
    total = current_df["Subtotal"].sum()
    try:
        discount = float(ws.acell("J8").value or 0)
    except Exception:
        discount = 0.0
    vat = total * 0.12
    grand_total = total + vat - discount

    # Save button (manual batch save)
    save_col1, save_col2 = st.columns([5, 1])
    with save_col2:
        if st.button("💾 Save Changes", key="save_changes"):
            with st.spinner("Saving changes to Google Sheets..."):
                try:
                    session_key = f"project_df_{project}"

                    # ✅ Get current edited DF (this is the one true source)
                    new_df = st.session_state[session_key].copy()

                    # ✅ Load sheet version for diff
                    old_df = df_from_worksheet(ws).reset_index(drop=True)

                    # ✅ Numeric cleanup
                    for col in ["Quantity", "Unit Price", "Subtotal"]:
                        new_df[col] = pd.to_numeric(new_df[col], errors="coerce").fillna(0).astype(float)

                    # ✅ Recompute subtotal
                    new_df["Subtotal"] = (new_df["Quantity"] * new_df["Unit Price"]).round(2)

                    # ✅ Ensure pure Python numbers (avoid numpy in gspread)
                    new_df = new_df.applymap(lambda x: x.item() if hasattr(x, "item") else x)
    
                    # ✅ Ensure Items always correct
                    new_df["Item"] = range(1, len(new_df) + 1)

                    # ✅ Write ONLY the changed rows
                    apply_sheet_updates(ws, old_df, new_df)

                    # ✅ Recompute totals with the saved values (in case J8 changed externally)
                    saved_total = new_df["Subtotal"].sum()
                    saved_vat = saved_total * 0.12
                    try:
                        saved_discount = float(ws.acell("J8").value or 0)
                    except Exception:
                        saved_discount = 0.0
                    saved_grand = saved_total + saved_vat - saved_discount

                    # ✅ Save totals below
                    save_totals_to_ws(ws, saved_total, saved_vat, saved_grand)

                    # ✅ Update session DF so editor stays in sync — critical fix!
                    st.session_state[session_key] = new_df.copy()

                    # ✅ Mark clean
                    st.session_state.unsaved_changes = False

                    st.toast("✅ Changes saved to Google Sheets!", icon="💾")

                except Exception as e:
                    st.error(f"❌ Failed to save changes: {e}")

    # Display metrics
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

    # Terms & Conditions (unchanged)
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
        save_totals_to_ws(ws, total, vat, grand_total)
        st.success("Saved terms successfully.")

    # Export PDF: per your choice (C) always use latest saved sheet version
    if export_pdf:
        try:
            sheet_df = df_from_worksheet(ws)
            terms = read_terms_from_ws(ws)
            totals = {
                "subtotal": sheet_df["Subtotal"].sum(),
                "discount": float(ws.acell("J8").value or 0),
                "vat": sheet_df["Subtotal"].sum() * 0.12,
                "total": sheet_df["Subtotal"].sum() + (sheet_df["Subtotal"].sum() * 0.12) - float(ws.acell("J8").value or 0)
            }
            pdf_buffer = generate_pdf(project, sheet_df, totals, terms)
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















