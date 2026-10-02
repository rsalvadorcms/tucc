import streamlit as st
import pandas as pd
import sqlite3
import io
import urllib.parse
import string
import os
import re
import shutil
import threading
import time as time_module
from datetime import datetime, date, timedelta, time

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

# Optional import for PDF rendering
try:
    from reportlab.lib.pagesizes import A3, A4, portrait
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

# Set page configurations with native default theme formatting
st.set_page_config(page_title="Office Operations Portal", layout="wide")

LOGO1_PATH = "logo.png"
LOGO2_PATH = "logo2.png"
BACKUP_DIR = "backups"
NEWS_DIR = "site_news_uploads"

# Ensure directories exist
if not os.path.exists(BACKUP_DIR):
    os.makedirs(BACKUP_DIR)
if not os.path.exists(NEWS_DIR):
    os.makedirs(NEWS_DIR)

# ==============================================================================
# ⚙️ 1. HELPER FUNCTIONS & DATABASE ENGINE
# ==============================================================================
DB_FILE = "office_operations.db"

@st.dialog("Data Transaction Status")
def show_transaction_dialog(title_text: str, message_text: str, status_type: str = "success"):
    """Displays transaction outcome in a pop-up modal message box."""
    if status_type == "success":
        st.success(f"### {title_text}")
    elif status_type == "error":
        st.error(f"### {title_text}")
    else:
        st.info(f"### {title_text}")
        
    st.write(message_text)
    
    if st.button("OK", type="primary", use_container_width=True):
        if 'tx_dialog_title' in st.session_state:
            del st.session_state['tx_dialog_title']
        if 'tx_dialog_msg' in st.session_state:
            del st.session_state['tx_dialog_msg']
        if 'tx_dialog_type' in st.session_state:
            del st.session_state['tx_dialog_type']
        st.rerun()

def trigger_transaction_dialog():
    """Renders the modal popup box if status data exists in session state."""
    if 'tx_dialog_title' in st.session_state and st.session_state.tx_dialog_title:
        show_transaction_dialog(
            st.session_state.tx_dialog_title,
            st.session_state.get('tx_dialog_msg', ''),
            st.session_state.get('tx_dialog_type', 'success')
        )

def set_transaction_dialog(title: str, message: str, status_type: str = "success"):
    """Stores dialog notification parameters into session state."""
    st.session_state.tx_dialog_title = title
    st.session_state.tx_dialog_msg = message
    st.session_state.tx_dialog_type = status_type

def perform_automated_backup():
    """Creates a timestamped backup of the SQLite database and exports key tables to Excel."""
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)
        
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    if os.path.exists(DB_FILE):
        backup_db_path = os.path.join(BACKUP_DIR, f"backup_db_{timestamp}.db")
        shutil.copy2(DB_FILE, backup_db_path)
        
    try:
        conn = get_db_connection()
        backup_excel_path = os.path.join(BACKUP_DIR, f"backup_data_{timestamp}.xlsx")
        with pd.ExcelWriter(backup_excel_path, engine='openpyxl') as writer:
            pd.read_sql_query("SELECT * FROM daily_transit", conn).to_excel(writer, index=False, sheet_name="daily_transit")
            pd.read_sql_query("SELECT * FROM overtime_requests", conn).to_excel(writer, index=False, sheet_name="overtime_requests")
            pd.read_sql_query("SELECT * FROM room_bookings", conn).to_excel(writer, index=False, sheet_name="room_bookings")
            pd.read_sql_query("SELECT * FROM users", conn).to_excel(writer, index=False, sheet_name="users")
            pd.read_sql_query("SELECT * FROM cars", conn).to_excel(writer, index=False, sheet_name="cars")
            pd.read_sql_query("SELECT * FROM trips", conn).to_excel(writer, index=False, sheet_name="trips")
            pd.read_sql_query("SELECT * FROM site_news", conn).to_excel(writer, index=False, sheet_name="site_news")
            pd.read_sql_query("SELECT * FROM transit_passengers", conn).to_excel(writer, index=False, sheet_name="transit_passengers")
        conn.close()
    except Exception as e:
        print(f"Excel backup export error: {e}")

_last_backup_slot = None

def run_scheduler():
    """Background worker loop running scheduled backups at 11:30 AM and 5:30 PM using standard libraries."""
    global _last_backup_slot
    while True:
        now = datetime.now()
        current_time_str = now.strftime("%H:%M")
        current_date_str = now.strftime("%Y-%m-%d")
        
        if current_time_str in ["11:30", "17:30"]:
            backup_key = f"{current_date_str}_{current_time_str}"
            if _last_backup_slot != backup_key:
                perform_automated_backup()
                _last_backup_slot = backup_key
                
        time_module.sleep(30)

if 'backup_scheduler_started' not in st.session_state:
    st.session_state.backup_scheduler_started = True
    scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
    scheduler_thread.start()

def format_military_time(input_str: str) -> str:
    if not input_str:
        return ""
    clean_digits = re.sub(r'\D', '', str(input_str).strip())
    if len(clean_digits) == 3:
        clean_digits = "0" + clean_digits
    elif len(clean_digits) == 1:
        clean_digits = "0" + clean_digits + "00"
    elif len(clean_digits) == 2:
        clean_digits = clean_digits + "00"
    if len(clean_digits) == 4:
        hours = int(clean_digits[:2])
        minutes = int(clean_digits[2:])
        if 0 <= hours <= 23 and 0 <= minutes <= 59:
            return f"{hours:02d}:{minutes:02d}"
    return None

def render_datalist_options(list_id: str, options: list):
    options_html = "".join([f'<option value="{opt}">' for opt in options if opt])
    st.components.v1.html(
        f"""
        <script>
        const parentDoc = window.parent.document;
        let dl = parentDoc.getElementById('{list_id}');
        if (!dl) {{
            dl = parentDoc.createElement('datalist');
            dl.id = '{list_id}';
            parentDoc.body.appendChild(dl);
        }}
        dl.innerHTML = '{options_html}';
        
        const inputs = parentDoc.querySelectorAll('input[type="text"]');
        inputs.forEach(input => {{
            if (input.placeholder && input.placeholder.includes('{list_id}')) {{
                input.setAttribute('list', '{list_id}');
            }}
        }});
        </script>
        """,
        height=0,
        width=0
    )

def export_df_to_excel(df, sheet_name="Data"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()

def export_overtime_summary_excel(df, title_text="Overtime Schedule"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overtime Summary"
    ws.views.sheetView[0].showGridLines = True
    
    font_title = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
    )
    
    ws.merge_cells("A1:E1")
    t_cell = ws["A1"]
    t_cell.value = title_text
    t_cell.font = font_title
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 25
    
    headers = list(df.columns)
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[3].height = 20
    
    for r_idx, row in enumerate(df.iterrows(), start=4):
        for c_idx, val in enumerate(row[1], start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value="" if pd.isna(val) else val)
            cell.font = font_data
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
    for col_idx in range(1, ws.max_column + 1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        max_len = max((len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(1, ws.max_row + 1)), default=0)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 18)
        
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def export_overtime_summary_pdf(df, title_text="Overtime Schedule"):
    if not HAS_REPORTLAB:
        return None
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=portrait(A4), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=13, leading=16, alignment=1, textColor=colors.HexColor('#1F4E78'))
    cell_style = ParagraphStyle('CellSt', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=11, alignment=1)
    header_style = ParagraphStyle('HeadSt', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, leading=11, alignment=1, textColor=colors.white)
    
    story = [Paragraph(title_text, title_style), Spacer(1, 15)]
    
    headers = list(df.columns)
    table_data = [[Paragraph(h, header_style) for h in headers]]
    
    for _, row in df.iterrows():
        row_cells = [Paragraph(str(val if pd.notna(val) else ""), cell_style) for val in row]
        table_data.append(row_cells)
        
    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E78')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#BFBFBF')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def export_custom_batam_excel(detailed_df, effective_date_str=""):
    wb = openpyxl.Workbook()
    font_title = Font(name="Calibri", size=13, bold=True, color="1F4E78")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold_label = Font(name="Calibri", size=10, bold=True, color="000000")
    font_data = Font(name="Calibri", size=10)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
    )

    ws2 = wb.active
    ws2.title = "Detailed Allocations"
    ws2.views.sheetView[0].showGridLines = True
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A3
    ws2.page_setup.orientation = ws2.ORIENTATION_PORTRAIT
    ws2.print_title_rows = '1:6'
    
    for r in range(1, 6):
        for c in range(1, 7):
            ws2.cell(row=r, column=c).fill = white_fill

    ws2.row_dimensions[2].height = 20
    ws2.row_dimensions[3].height = 20
    ws2.row_dimensions[4].height = 20

    cell_e5 = ws2.cell(row=5, column=5, value="Effective Date:")
    cell_e5.font = font_bold_label
    cell_e5.alignment = Alignment(horizontal="right", vertical="center")
    
    cell_f5 = ws2.cell(row=5, column=6, value=effective_date_str)
    cell_f5.font = font_bold_label
    cell_f5.alignment = Alignment(horizontal="center", vertical="center")

    start_row = 6
    ws2.merge_cells("B2:D4")
    title_cell = ws2["B2"]
    title_cell.value = "Daily Transportation Arrangement - Passenger list\nTUCC PROJECT - BATAM MODULE YARD [MD-1 & MD-4]\nJOB CODE : 0 - 0847 - 00 - 0001"
    title_cell.font = font_title
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    if os.path.exists(LOGO1_PATH):
        try:
            img1_det = OpenpyxlImage(LOGO1_PATH)
            img1_det.width = 110; img1_det.height = 50
            ws2.add_image(img1_det, "A2")
        except Exception:
            pass

    if os.path.exists(LOGO2_PATH):
        try:
            img2_det = OpenpyxlImage(LOGO2_PATH)
            img2_det.width = 130; img2_det.height = 50
            ws2.add_image(img2_det, "F2")
        except Exception:
            pass

    det_headers = ["Vehicle Description", "Driver Name", "Contact Number", "Passenger", "ETD 1", "ETD 2"]
    for col_idx, h_title in enumerate(det_headers, start=1):
        cell = ws2.cell(row=start_row, column=col_idx, value=h_title)
        cell.font = font_header; cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    data_start = start_row + 1
    for r_idx, row in enumerate(detailed_df.iterrows(), start=data_start):
        row_data = row[1]
        v_model = str(row_data.get("Vehicle Model", "") or "Standard Vehicle").strip()
        p_num = str(row_data.get("Plate Number", "") or "N/A").strip()
        v_color = str(row_data.get("Color", "") or "Black").strip()
        v_desc_3lines = f"{v_model}\n{p_num}\nColor : {v_color}"
            
        d_name = row_data.get("Driver Name", "")
        c_num = row_data.get("Contact Number", "")
        p_name = row_data.get("Passenger Name", "")
        etd1 = row_data.get("ETD 1 (From)", "")
        etd2 = row_data.get("ETD 2 (To)", "")
        
        ordered_vals = [v_desc_3lines, d_name, c_num, p_name, etd1, etd2]
        for c_idx, val in enumerate(ordered_vals, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value="" if pd.isna(val) else val)
            cell.font = font_data; cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if not detailed_df.empty:
        current_grp = None
        grp_start = data_start
        tot_rows = len(detailed_df)
        for idx, grp_val in enumerate(detailed_df['Car Group'].values, start=data_start):
            if grp_val != current_grp:
                if current_grp is not None and (idx - 1) > grp_start:
                    for merge_col in [1, 2, 3, 5, 6]:
                        ws2.merge_cells(start_row=grp_start, start_column=merge_col, end_row=idx - 1, end_column=merge_col)
                        ws2.cell(row=grp_start, column=merge_col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                current_grp = grp_val
                grp_start = idx
        if current_grp is not None and (data_start + tot_rows - 1) > grp_start:
            for merge_col in [1, 2, 3, 5, 6]:
                ws2.merge_cells(start_row=grp_start, start_column=merge_col, end_row=data_start + tot_rows - 1, end_column=merge_col)
                ws2.cell(row=grp_start, column=merge_col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col in ws2.columns:
        col_letter = col[0].column_letter
        if col_letter == 'B':
            ws2.column_dimensions['B'].width = 12
        elif col_letter == 'C':
            ws2.column_dimensions['C'].width = 18
        else:
            max_len = max((len(str(cell.value or '')) for cell in col), default=0)
            ws2.column_dimensions[col_letter].width = max(max_len + 4, 18)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def export_custom_batam_pdf(detailed_df, effective_date_str=""):
    if not HAS_REPORTLAB:
        return None
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=portrait(A3), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=15, leading=19, alignment=1, textColor=colors.HexColor('#1F4E78'))
    subtitle_style = ParagraphStyle('DocSubtitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=14, alignment=1, textColor=colors.HexColor('#333333'))
    eff_date_style = ParagraphStyle('EffDateStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=2, textColor=colors.HexColor('#000000'))
    cell_style = ParagraphStyle('CellText', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=11, alignment=1)
    header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1, textColor=colors.white)
    
    story = []
    title_p = Paragraph("Daily Transportation Arrangement - Passenger list", title_style)
    sub1_p = Paragraph("TUCC PROJECT - BATAM MODULE YARD [MD-1 & MD-4]", subtitle_style)
    sub2_p = Paragraph("JOB CODE : 0 - 0847 - 00 - 0001", subtitle_style)
    header_box = [title_p, Spacer(1, 3), sub1_p, Spacer(1, 2), sub2_p]
    
    img1_elem = RLImage(LOGO1_PATH, width=100, height=45) if os.path.exists(LOGO1_PATH) else ""
    img2_elem = RLImage(LOGO2_PATH, width=120, height=45) if os.path.exists(LOGO2_PATH) else ""
    
    top_table = Table([[img1_elem, header_box, img2_elem]], colWidths=[110, 560, 110])
    top_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (0,0), 'LEFT'), ('ALIGN', (1,0), (1,0), 'CENTER'), ('ALIGN', (2,0), (2,0), 'RIGHT')]))
    story.append(top_table)
    story.append(Spacer(1, 10))
    
    eff_p = Paragraph(f"<b>Effective Date:</b> {effective_date_str}", eff_date_style)
    eff_table = Table([[Paragraph("", cell_style), eff_p]], colWidths=[550, 230])
    eff_table.setStyle(TableStyle([('ALIGN', (1,0), (1,0), 'RIGHT'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(eff_table)
    story.append(Spacer(1, 8))
    
    headers = ["Vehicle Description", "Driver Name", "Contact Number", "Passenger", "ETD 1", "ETD 2"]
    table_data = [[Paragraph(h, header_style) for h in headers]]
    table_spans = []
    current_grp = None
    grp_start = 1
    
    for idx, row_data in enumerate(detailed_df.iterrows(), start=1):
        r_val = row_data[1]
        grp_name = r_val.get("Car Group", "")
        v_model = str(r_val.get("Vehicle Model", "") or "Standard Vehicle").strip()
        p_num = str(r_val.get("Plate Number", "") or "N/A").strip()
        v_color = str(r_val.get("Color", "") or "Black").strip()
        v_desc_html = f"{v_model}<br/>{p_num}<br/>Color : {v_color}"
        
        row_cells = [
            Paragraph(v_desc_html, cell_style),
            Paragraph(str(r_val.get("Driver Name", "") or ""), cell_style),
            Paragraph(str(r_val.get("Contact Number", "") or ""), cell_style),
            Paragraph(str(r_val.get("Passenger Name", "") or ""), cell_style),
            Paragraph(str(r_val.get("ETD 1 (From)", "") or ""), cell_style),
            Paragraph(str(r_val.get("ETD 2 (To)", "") or ""), cell_style)
        ]
        table_data.append(row_cells)
        if grp_name != current_grp:
            if current_grp is not None and (idx - 1) > grp_start:
                for col_i in [0, 1, 2, 4, 5]:
                    table_spans.append(('SPAN', (col_i, grp_start), (col_i, idx - 1)))
            current_grp = grp_name
            grp_start = idx

    tot_rows = len(detailed_df)
    if current_grp is not None and tot_rows > grp_start:
        for col_i in [0, 1, 2, 4, 5]:
            table_spans.append(('SPAN', (col_i, grp_start), (col_i, tot_rows)))
            
    main_table = Table(table_data, colWidths=[150, 120, 120, 210, 90, 90], repeatRows=1)
    main_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E78')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#BFBFBF')),
        ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ] + table_spans))
    story.append(main_table)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def export_shuttle_timetable_excel(df, effective_date_str=""):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Shuttle Timetable"
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    
    font_title = Font(name="Calibri", size=13, bold=True, color="1F4E78")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
    )

    for r in range(1, 20):
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = white_fill

    ws.merge_cells("B2:F3")
    title_cell = ws["B2"]
    title_cell.value = "JGC SHUTTLE TIMETABLE\nTUCC PROJECT - BATAM MODULE YARD [MD-1 & MD-4]"
    title_cell.font = font_title
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if os.path.exists(LOGO1_PATH):
        try:
            img1 = OpenpyxlImage(LOGO1_PATH)
            img1.width = 110; img1.height = 45
            ws.add_image(img1, "A2")
        except Exception:
            pass

    if os.path.exists(LOGO2_PATH):
        try:
            img2 = OpenpyxlImage(LOGO2_PATH)
            img2.width = 120; img2.height = 45
            ws.add_image(img2, "G2")
        except Exception:
            pass

    ws.merge_cells("A5:A6")
    ws["A5"] = "Days (s)"
    
    ws.merge_cells("B5:B6")
    ws["B5"] = "UNIT"
    
    ws.merge_cells("C5:C6")
    ws["C5"] = "DRIVER"
    
    ws.merge_cells("D5:D6")
    ws["D5"] = "TRIP NO."

    ws.merge_cells("E5:F5")
    ws["E5"] = "ROUTE"
    ws["E6"] = "YARD - 1"
    ws["F6"] = "YARD - 3"

    ws.merge_cells("G5:G6")
    ws["G5"] = "REMARKS"

    for r in [5, 6]:
        ws.row_dimensions[r].height = 22
        for col in range(1, 8):
            cell = ws.cell(row=r, column=col)
            cell.font = font_header
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

    start_row = 7
    sorted_df = df.sort_values(by="ETD Start") if "ETD Start" in df.columns else df
    
    for idx, row in enumerate(sorted_df.iterrows(), start=start_row):
        r_val = row[1]
        unit_val = str(r_val.get("Unit", "TOYOTA HI-ACE"))
        driver_val = str(r_val.get("Driver", "TBA"))
        trip_no = f"{idx - start_row + 1}st" if (idx - start_row + 1) == 1 else f"{idx - start_row + 1}nd" if (idx - start_row + 1) == 2 else f"{idx - start_row + 1}rd" if (idx - start_row + 1) == 3 else f"{idx - start_row + 1}th"
        etd1 = str(r_val.get("ETD Start", ""))
        etd2 = str(r_val.get("ETD Return", ""))
        remarks = str(r_val.get("Remarks", "DROP-OFF"))

        row_values = ["", unit_val, driver_val, trip_no, etd1, etd2, remarks]
        ws.row_dimensions[idx].height = 20
        for c_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    total_rows = len(sorted_df)
    if total_rows > 0:
        end_row = start_row + total_rows - 1
        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        days_cell = ws.cell(row=start_row, column=1)
        days_cell.value = "MONDAY\nTUESDAY\nWEDNESDAY\nTHURSDAY\nFRIDAY\nSATURDAY"
        days_cell.font = font_data
        days_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        days_cell.border = thin_border

    max_unit_len = max([len(str(r.get("Unit", ""))) for _, r in sorted_df.iterrows()] + [4], default=18)
    col_width_b = max(max_unit_len + 4, 22)

    col_widths = {'A': 18, 'B': col_width_b, 'C': 16, 'D': 12, 'E': 14, 'F': 14, 'G': 18}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY, password TEXT NOT NULL,
            role TEXT NOT NULL, email_recipients TEXT, emp_name TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS holidays (
            holiday_date TEXT PRIMARY KEY, description TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS trips (
            trip TEXT PRIMARY KEY, trip_name TEXT NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS overtime_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, emp_name TEXT,
            ot_date TEXT, start_time TEXT, end_time TEXT, needs_transport TEXT DEFAULT 'Yes',
            origin TEXT, destination TEXT, departure_time TEXT, return_time TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS meeting_rooms (
            room_number TEXT PRIMARY KEY, room_name TEXT, capacity INTEGER, location TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS room_bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT, room_number TEXT, booked_by TEXT,
            booking_date TEXT, start_time TEXT, end_time TEXT, is_recurring TEXT, recurrence_end_date TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cars (
            car_name TEXT PRIMARY KEY, plate_number TEXT UNIQUE NOT NULL, vehicle TEXT, color TEXT DEFAULT 'Black'
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fleet_drivers (
            driver_name TEXT PRIMARY KEY, driver_mobile TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transit_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT, group_name TEXT UNIQUE NOT NULL,
            driver_name TEXT, etd_1 TEXT, etd_2 TEXT,
            FOREIGN KEY (group_name) REFERENCES cars(car_name) ON DELETE CASCADE,
            FOREIGN KEY (driver_name) REFERENCES fleet_drivers(driver_name) ON DELETE SET NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transit_passengers (
            id INTEGER PRIMARY KEY AUTOINCREMENT, group_name TEXT NOT NULL, passengers TEXT NOT NULL,
            FOREIGN KEY (group_name) REFERENCES transit_groups(group_name) ON DELETE CASCADE
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_transit (
            id INTEGER PRIMARY KEY AUTOINCREMENT, transit_date_start TEXT NOT NULL, transit_date_end TEXT NOT NULL,
            group_name TEXT DEFAULT 'TBA', requested_by TEXT, etd_1 TEXT, etd_2 TEXT,
            location_from TEXT, location_to TEXT, daily TEXT DEFAULT 'No', trip TEXT DEFAULT 'Trip A'
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS site_news (
            id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, filename TEXT NOT NULL,
            file_path TEXT NOT NULL, file_type TEXT NOT NULL, uploaded_by TEXT, upload_date TEXT
        )
    ''')
    
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO users VALUES ('admin', 'admin123', 'Admin', 'admin@company.com', 'System Administrator')")

    cursor.execute("SELECT COUNT(*) FROM trips")
    if cursor.fetchone()[0] == 0:
        for t_code, t_desc in [("Trip A", "Yard to Yard"), ("Trip B", "Sunday Panbil - Wasco - Panbil"), ("Trip C", "Panbil - Destination - Panbil"), ("Trip D", "Overtime Dispatch Route")]:
            cursor.execute("INSERT INTO trips (trip, trip_name) VALUES (?, ?)", (t_code, t_desc))
        
    cursor.execute("SELECT COUNT(*) FROM meeting_rooms")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO meeting_rooms VALUES ('101', 'Boardroom', 15, '1st Floor')")
        cursor.execute("INSERT INTO meeting_rooms VALUES ('102', 'Huddle Room Alpha', 6, '2nd Floor')")

    cursor.execute("SELECT COUNT(*) FROM fleet_drivers")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO fleet_drivers VALUES ('John Doe', '+628111222333')")
        cursor.execute("INSERT INTO fleet_drivers VALUES ('Jane Smith', '+628999888777')")

    cursor.execute("SELECT COUNT(*) FROM cars WHERE car_name = 'TBA'")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO cars (car_name, plate_number, vehicle, color) VALUES ('TBA', 'TBA', 'TBA', 'TBA')")

    cursor.execute("SELECT COUNT(*) FROM cars")
    if cursor.fetchone()[0] == 1:
        cursor.execute("INSERT OR IGNORE INTO cars VALUES ('Car A', 'B 1234 ABC', 'Toyota Avanza', 'Black')")
        cursor.execute("INSERT OR IGNORE INTO cars VALUES ('Car B', 'B 5678 XYZ', 'Toyota Innova', 'White')")
        
    conn.commit()
    conn.close()

init_db()

def run_migrations():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(cars)")
    if "color" not in [col[1] for col in cursor.fetchall()]:
        cursor.execute("ALTER TABLE cars ADD COLUMN color TEXT DEFAULT 'Black'")
    conn.commit()
    conn.close()

run_migrations()
trigger_transaction_dialog()

# ==============================================================================
# 🔐 2. AUTHENTICATION USER INTERFACE
# ==============================================================================
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.emp_name = ""

if 'emp_name' not in st.session_state:
    st.session_state.emp_name = st.session_state.get('username', '')

if not st.session_state.logged_in:
    st.title("🏢 Office Operations Portal")
    st.subheader("Login to access scheduling & overtime systems")
    
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log In")
        
        if submitted:
            conn = get_db_connection()
            user = conn.execute("SELECT * FROM users WHERE username=? AND password=?", (username, password)).fetchone()
            conn.close()
            if user:
                st.session_state.logged_in = True
                st.session_state.username = user['username']
                st.session_state.role = user['role']
                st.session_state.emp_name = user['emp_name'] or user['username']
                st.rerun()
            else:
                st.error("Invalid username or password configuration.")
    st.stop()

# ==============================================================================
# 🗂️ 3. MAIN APP CONTROL PANELS & SIDEBAR NAVIGATION
# ==============================================================================
st.sidebar.title(f"👋 Welcome, {st.session_state.username}")
st.sidebar.info(f"Access Level: **{st.session_state.role}**")

st.sidebar.markdown("---")
st.sidebar.subheader("📌 Navigation")
nav_selection = st.sidebar.radio(
    "Go to:",
    [
        "🏠 Home & Overview (Passenger List)",
        "📅 Scheduled Transit Dispatches (Shuttle Format)",
        "⏰ Overtime & Transport",
        "👥 Transit Groups & Passengers",
        "📅 Daily Transit Dispatch Setup",
        "🏢 Meeting Rooms",
        "📢 Site News",
        "🛠️ System Administration"
    ]
)

st.sidebar.markdown("---")
if st.sidebar.button("Logout Profile"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.emp_name = ""
    st.rerun()

# --- VIEW RENDERERS BASED ON SIDEBAR NAVIGATION ---

# 1. HOME & OVERVIEW (Passenger List Layout with Logo Banners & White Background Table Styling)
if nav_selection == "🏠 Home & Overview (Passenger List)":
    # Header Section with Logos & Title matching Excel layout
    h_col1, h_col2, h_col3 = st.columns([1, 4, 1])
    with h_col1:
        if os.path.exists(LOGO1_PATH):
            st.image(LOGO1_PATH, width=130)
    with h_col2:
        st.markdown(
            """
            <div style="text-align: center;">
                <h3 style="color: #1F4E78; margin-bottom: 0px;">Daily Transportation Arrangement - Passenger list</h3>
                <h4 style="color: #333333; margin-top: 2px; margin-bottom: 2px;">TUCC PROJECT - BATAM MODULE YARD [MD-1 & MD-4]</h4>
                <p style="color: #555555; font-size: 14px; font-weight: bold; margin-top: 0px;">JOB CODE : 0 - 0847 - 00 - 0001</p>
            </div>
            """,
            unsafe_allow_html=True
        )
    with h_col3:
        if os.path.exists(LOGO2_PATH):
            st.image(LOGO2_PATH, width=140)

    st.markdown("---")
    
    col_lbl, col_date = st.columns([3, 1])
    with col_date:
        home_effective_date = st.date_input("Effective Date", value=date.today(), key="home_eff_date")
        home_eff_date_str = home_effective_date.strftime("%Y-%m-%d")

    conn = get_db_connection()
    home_unrolled_df = pd.read_sql_query('''
        SELECT tg.group_name AS "Car Group", c.vehicle AS "Vehicle Model", c.plate_number AS "Plate Number", 
               c.color AS "Color", tg.driver_name AS "Driver Name", fd.driver_mobile AS "Contact Number", 
               tg.etd_1 AS "ETD 1 (From)", tg.etd_2 AS "ETD 2 (To)", tp.passengers AS "Passenger Name"
        FROM transit_groups tg
        LEFT JOIN cars c ON tg.group_name = c.car_name
        LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
        LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
        ORDER BY tg.group_name
    ''', conn)
    conn.close()

    if not home_unrolled_df.empty:
        display_home_df = home_unrolled_df.copy()
        display_home_df["Vehicle Description"] = display_home_df.apply(
            lambda r: f"{r['Vehicle Model'] or 'Standard Vehicle'}<br>{r['Plate Number'] or 'N/A'}<br>Color : {r['Color'] or 'Black'}",
            axis=1
        )
        
        # Build clean merged-group HTML table representation to mimic the exact Excel layout visually with white background
        html_table_rows = ""
        current_group = None
        group_rowspan_counts = display_home_df['Car Group'].value_counts()
        
        # We will iterate and render HTML rows with rowspan for shared group properties
        rendered_groups = set()
        for _, row in display_home_df.iterrows():
            g_name = row['Car Group']
            v_desc = row['Vehicle Description']
            d_name = row['Driver Name'] or ""
            c_num = row['Contact Number'] or ""
            p_name = row['Passenger Name'] or ""
            etd1 = row['ETD 1 (From)'] or ""
            etd2 = row['ETD 2 (To)'] or ""
            
            span_count = group_rowspan_counts.get(g_name, 1)
            
            html_table_rows += "<tr>"
            if g_name != current_group:
                current_group = g_name
                rendered_groups.clear()
                html_table_rows += f"<td rowspan='{span_count}' style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'><b>{g_name}</b><br><span style='font-size:11px;'>{v_desc}</span></td>"
                html_table_rows += f"<td rowspan='{span_count}' style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{d_name}</td>"
                html_table_rows += f"<td rowspan='{span_count}' style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{c_num}</td>"
            
            html_table_rows += f"<td style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{p_name}</td>"
            
            if g_name not in rendered_groups:
                rendered_groups.add(g_name)
                html_table_rows += f"<td rowspan='{span_count}' style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{etd1}</td>"
                html_table_rows += f"<td rowspan='{span_count}' style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{etd2}</td>"
            
            html_table_rows += "</tr>"

        full_custom_html = f"""
        <div style="background-color: white; padding: 15px; border-radius: 5px;">
            <div style="text-align: right; font-weight: bold; margin-bottom: 10px; color: black;">Effective Date: {home_eff_date_str}</div>
            <table style="width: 100%; border-collapse: collapse; background-color: white; color: black; font-family: Calibri, sans-serif; font-size: 14px;">
                <thead>
                    <tr style="background-color: #1F4E78; color: white; text-align: center;">
                        <th style="border: 1px solid #BFBFBF; padding: 10px;">Vehicle Description</th>
                        <th style="border: 1px solid #BFBFBF; padding: 10px;">Driver Name</th>
                        <th style="border: 1px solid #BFBFBF; padding: 10px;">Contact Number</th>
                        <th style="border: 1px solid #BFBFBF; padding: 10px;">Passenger</th>
                        <th style="border: 1px solid #BFBFBF; padding: 10px;">ETD 1</th>
                        <th style="border: 1px solid #BFBFBF; padding: 10px;">ETD 2</th>
                    </tr>
                </thead>
                <tbody>
                    {html_table_rows}
                </tbody>
            </table>
        </div>
        """
        st.markdown(full_custom_html, unsafe_allow_html=True)
    else:
        st.info("No transit groups or passenger assignments configured yet. Go to 'Transit Groups & Passengers' in the navigation menu.")

# 2. SCHEDULED TRANSIT DISPATCHES (Shuttle Format - Trips A, B, C with Logo Banners & White Background Table Styling)
elif nav_selection == "📅 Scheduled Transit Dispatches (Shuttle Format)":
    h_col1, h_col2, h_col3 = st.columns([1, 4, 1])
    with h_col1:
        if os.path.exists(LOGO1_PATH):
            st.image(LOGO1_PATH, width=130)
    with h_col2:
        st.markdown(
            """
            <div style="text-align: center;">
                <h3 style="color: #1F4E78; margin-bottom: 0px;">JGC SHUTTLE TIMETABLE</h3>
                <h4 style="color: #333333; margin-top: 2px; margin-bottom: 2px;">TUCC PROJECT - BATAM MODULE YARD [MD-1 & MD-4]</h4>
            </div>
            """,
            unsafe_allow_html=True
        )
    with h_col3:
        if os.path.exists(LOGO2_PATH):
            st.image(LOGO2_PATH, width=140)

    st.markdown("---")

    conn = get_db_connection()
    shuttle_raw_df = pd.read_sql_query('''
        SELECT dt.id, dt.transit_date_start, dt.transit_date_end, 
               dt.trip AS trip_code,
               dt.requested_by, dt.group_name, c.plate_number, c.vehicle, 
               dt.location_from, dt.location_to, dt.etd_1, dt.etd_2, 
               tg.driver_name
        FROM daily_transit dt 
        LEFT JOIN cars c ON dt.group_name = c.car_name
        LEFT JOIN transit_groups tg ON dt.group_name = tg.group_name
        WHERE dt.trip IN ('Trip A', 'Trip B', 'Trip C')
        ORDER BY dt.etd_1 ASC, dt.id DESC
    ''', conn)
    conn.close()

    if not shuttle_raw_df.empty:
        shuttle_display_df = shuttle_raw_df.copy()
        shuttle_display_df['Days'] = "MONDAY<br>TUESDAY<br>WEDNESDAY<br>THURSDAY<br>FRIDAY<br>SATURDAY"
        shuttle_display_df['Unit'] = shuttle_display_df.apply(lambda r: f"{r['vehicle']} - {r['plate_number']}" if pd.notna(r['plate_number']) and r['plate_number'] != 'TBA' else "TOYOTA HI-ACE", axis=1)
        shuttle_display_df['Driver'] = shuttle_display_df['driver_name'].fillna("TBA")
        shuttle_display_df['Trip No.'] = [f"{i}st" if i==1 else f"{i}nd" if i==2 else f"{i}rd" if i==3 else f"{i}th" for i in range(1, len(shuttle_display_df)+1)]
        shuttle_display_df['Route (Yard-1)'] = shuttle_display_df['etd_1']
        shuttle_display_df['Route (Yard-3)'] = shuttle_display_df['etd_2']
        shuttle_display_df['Remarks'] = "DROP-OFF / PICK-UP"

        shuttle_rows_html = ""
        total_rows_count = len(shuttle_display_df)
        
        for idx, row in shuttle_display_df.iterrows():
            u_val = row['Unit']
            d_val = row['Driver']
            t_no = row['Trip No.']
            r_y1 = row['Route (Yard-1)']
            r_y3 = row['Route (Yard-2)' if 'Route (Yard-2)' in row else 'Route (Yard-3)']
            rem = row['Remarks']
            
            shuttle_rows_html += "<tr>"
            if idx == 0:
                shuttle_rows_html += f"<td rowspan='{total_rows_count}' style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 10px; font-weight: bold;'>MONDAY<br>TUESDAY<br>WEDNESDAY<br>THURSDAY<br>FRIDAY<br>SATURDAY</td>"
            
            shuttle_rows_html += f"<td style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{u_val}</td>"
            shuttle_rows_html += f"<td style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{d_val}</td>"
            shuttle_rows_html += f"<td style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{t_no}</td>"
            shuttle_rows_html += f"<td style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{r_y1}</td>"
            shuttle_rows_html += f"<td style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{r_y3}</td>"
            shuttle_rows_html += f"<td style='background-color: white; border: 1px solid #BFBFBF; text-align: center; vertical-align: middle; padding: 8px;'>{rem}</td>"
            shuttle_rows_html += "</tr>"

        full_shuttle_html = f"""
        <div style="background-color: white; padding: 15px; border-radius: 5px;">
            <table style="width: 100%; border-collapse: collapse; background-color: white; color: black; font-family: Calibri, sans-serif; font-size: 14px;">
                <thead>
                    <tr style="background-color: #1F4E78; color: white; text-align: center;">
                        <th rowspan="2" style="border: 1px solid #BFBFBF; padding: 10px;">Days (s)</th>
                        <th rowspan="2" style="border: 1px solid #BFBFBF; padding: 10px;">UNIT</th>
                        <th rowspan="2" style="border: 1px solid #BFBFBF; padding: 10px;">DRIVER</th>
                        <th rowspan="2" style="border: 1px solid #BFBFBF; padding: 10px;">TRIP NO.</th>
                        <th colspan="2" style="border: 1px solid #BFBFBF; padding: 10px;">ROUTE</th>
                        <th rowspan="2" style="border: 1px solid #BFBFBF; padding: 10px;">REMARKS</th>
                    </tr>
                    <tr style="background-color: #1F4E78; color: white; text-align: center;">
                        <th style="border: 1px solid #BFBFBF; padding: 8px;">YARD - 1</th>
                        <th style="border: 1px solid #BFBFBF; padding: 8px;">YARD - 3</th>
                    </tr>
                </thead>
                <tbody>
                    {shuttle_rows_html}
                </tbody>
            </table>
        </div>
        """
        st.markdown(full_shuttle_html, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        shuttle_excel_df = shuttle_raw_df.copy()
        shuttle_excel_df['Unit'] = shuttle_excel_df.apply(lambda r: f"{r['vehicle']} - {r['plate_number']}" if pd.notna(r['plate_number']) and r['plate_number'] != 'TBA' else "TOYOTA HI-ACE", axis=1)
        shuttle_excel_df['Driver'] = shuttle_excel_df['driver_name'].fillna("TBA")
        shuttle_excel_df['ETD Start'] = shuttle_excel_df['etd_1']
        shuttle_excel_df['ETD Return'] = shuttle_excel_df['etd_2']
        shuttle_excel_df['Remarks'] = "DROP-OFF"

        st.download_button(
            "📥 Download Shuttle Timetable Excel (.xlsx)",
            data=export_shuttle_timetable_excel(shuttle_excel_df),
            file_name=f"JGC_Shuttle_Timetable_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No scheduled transit dispatches found for Trip A, Trip B, or Trip C.")

# 3. OVERTIME & TRANSPORT
elif nav_selection == "⏰ Overtime & Transport":
    st.header("Request Overtime & Logistics Tracking")
    conn = get_db_connection()
    holiday_list = pd.read_sql_query("SELECT holiday_date FROM holidays", conn)['holiday_date'].tolist()
    all_emp_names = pd.read_sql_query("SELECT emp_name FROM users WHERE emp_name IS NOT NULL AND emp_name != ''", conn)['emp_name'].tolist()
    conn.close()
    
    current_user_emp = st.session_state.get("emp_name", "") or st.session_state.get("username", "")
    if not all_emp_names:
        all_emp_names = [current_user_emp] if current_user_emp else ["Default Employee"]
    logged_in_emp = current_user_emp if current_user_emp in all_emp_names else all_emp_names[0]

    today_date = date.today()
    tomorrow_date = today_date + timedelta(days=1)
    tomorrow_is_sunday = tomorrow_date.weekday() == 6
    tomorrow_is_holiday = tomorrow_date.strftime("%Y-%m-%d") in holiday_list
    
    allowed_ot_dates = [today_date]
    if tomorrow_is_sunday or tomorrow_is_holiday:
        allowed_ot_dates.append(tomorrow_date)

    col1, col2 = st.columns(2)
    with col1:
        selected_staff_members = st.multiselect("Select Staff Member(s) for Overtime", options=all_emp_names, default=[logged_in_emp])
        ot_date = st.selectbox("Select Target Date", options=allowed_ot_dates, format_func=lambda d: d.strftime("%Y-%m-%d"))
        date_str = ot_date.strftime("%Y-%m-%d") if ot_date else ""
        
        is_sunday = ot_date.weekday() == 6 if ot_date else False
        is_holiday = date_str in holiday_list if date_str else False
        
        if is_sunday or is_holiday:
            default_start, default_end, default_origin, default_dest, default_dep_time_str = time(7, 0), time(15, 0), "Panbil", "Yard-1 Office", "07:00"
            st.caption("ℹ️ Baseline Rule: **Sunday/Holiday (07:00 - 15:00)**.")
        else:
            default_start, default_end, default_origin, default_dest, default_dep_time_str = time(17, 30), time(19, 0), "Yard-1 Office", "Panbil", "19:00"
            st.caption("ℹ️ Baseline Rule: **Weekday/Saturday (17:30 - 19:00)**.")
        
        start_time = st.time_input("OT Start Time (24-hr Military Time)", value=default_start)
        end_time = st.time_input("OT End Time (24-hr Military Time)", value=default_end)
        needs_transport = st.selectbox("Require Individual Transportation Logistics?", ["Yes", "No"], index=0)
        
    with col2:
        if needs_transport == "Yes":
            origin = st.text_input("Origin Address", value=default_origin)
            destination = st.text_input("Target Destination", value=default_dest)
            dep_time_str = st.text_input("Departure Timeline Estimate (24-hr)", value=default_dep_time_str)
            ret_time_str = st.text_input("Return Timeline Estimate (24-hr)", value="")
        else:
            origin, destination, dep_time_str, ret_time_str = ["", "", "", ""]

    if st.button("Submit New Overtime Request"):
        if not ot_date or not date_str:
            set_transaction_dialog("Data Transaction Unsuccessful", "Target Date cannot be left blank.", "error")
        elif not selected_staff_members:
            set_transaction_dialog("Data Transaction Unsuccessful", "Please select at least one staff member.", "error")
        else:
            try:
                conn = get_db_connection()
                for staff in selected_staff_members:
                    conn.execute('''
                        INSERT INTO overtime_requests (username, emp_name, ot_date, start_time, end_time, needs_transport, origin, destination, departure_time, return_time)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (st.session_state.username, staff, date_str, start_time.strftime("%H:%M"), end_time.strftime("%H:%M"), needs_transport, origin, destination, dep_time_str, ret_time_str))
                
                existing_dispatch = conn.execute(
                    "SELECT id FROM daily_transit WHERE transit_date_start = ? AND transit_date_end = ? AND trip = 'Trip D'", 
                    (date_str, date_str)
                ).fetchone()
                
                if not existing_dispatch:
                    conn.execute('''
                        INSERT INTO daily_transit (transit_date_start, transit_date_end, group_name, requested_by, etd_1, etd_2, location_from, location_to, daily, trip)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (date_str, date_str, "TBA", st.session_state.username, dep_time_str if dep_time_str else "19:00", None, "Yard-1 Office", "Panbil", "No", "Trip D"))
                
                conn.commit()
                conn.close()
                set_transaction_dialog("Data Transaction Successful", f"Overtime request logged for {len(selected_staff_members)} staff member(s) & transit dispatch schedule created.", "success")
            except Exception as e:
                set_transaction_dialog("Data Transaction Unsuccessful", f"Failed to save record: {str(e)}", "error")
        st.rerun()

    st.subheader("📋 Overtime Submission History Log")
    conn = get_db_connection()
    if st.session_state.role in ["Admin", "Owner"]:
        ot_df = pd.read_sql_query("SELECT id, username, emp_name AS 'Employee Name', ot_date AS 'Date', start_time AS 'Start Time', end_time AS 'End Time', needs_transport AS 'Needs Transport', origin AS 'Origin', destination AS 'Destination', departure_time AS 'Departure Time', return_time AS 'Return Time' FROM overtime_requests", conn)
    else:
        ot_df = pd.read_sql_query("SELECT id, username, emp_name AS 'Employee Name', ot_date AS 'Date', start_time AS 'Start Time', end_time AS 'End Time', needs_transport AS 'Needs Transport', origin AS 'Origin', destination AS 'Destination', departure_time AS 'Departure Time', return_time AS 'Return Time' FROM overtime_requests WHERE username = ?", conn, params=[st.session_state.username])
    conn.close()
    
    if not ot_df.empty:
        st.dataframe(ot_df, use_container_width=True)
        
        st.markdown("##### 📥 Export Current & Next Date Overtime Staff List")
        target_export_dates = [today_date.strftime("%Y-%m-%d")]
        if tomorrow_is_sunday or tomorrow_is_holiday:
            target_export_dates.append(tomorrow_date.strftime("%Y-%m-%d"))
            
        conn = get_db_connection()
        placeholders = ','.join(['?'] * len(target_export_dates))
        summary_query = f"SELECT emp_name AS 'Employee Name', ot_date AS 'Date', start_time AS 'Start Time', end_time AS 'End Time', needs_transport AS 'Needs Transport', origin AS 'Origin', destination AS 'Destination' FROM overtime_requests WHERE ot_date IN ({placeholders})"
        curr_next_ot_df = pd.read_sql_query(summary_query, conn, params=target_export_dates)
        conn.close()
        
        if not curr_next_ot_df.empty:
            exp_col1, exp_col2 = st.columns(2)
            dates_label_str = " & ".join(target_export_dates)
            with exp_col1:
                st.download_button(
                    "📥 Export Current/Next OT Staff to Excel (.xlsx)",
                    data=export_overtime_summary_excel(curr_next_ot_df, title_text=f"Overtime Staff List ({dates_label_str})"),
                    file_name=f"Overtime_Staff_List_{datetime.today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with exp_col2:
                if HAS_REPORTLAB:
                    st.download_button(
                        "📄 Export Current/Next OT Staff to PDF (.pdf)",
                        data=export_overtime_summary_pdf(curr_next_ot_df, title_text=f"Overtime Staff List ({dates_label_str})"),
                        file_name=f"Overtime_Staff_List_{datetime.today().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

        if st.session_state.role in ["Admin", "Owner"]:
            with st.expander("✏ Manage / Remove Overtime Submissions"):
                conn = get_db_connection()
                ot_raw = pd.read_sql_query("SELECT id, username, ot_date, emp_name, start_time, end_time, departure_time FROM overtime_requests ORDER BY id DESC", conn)
                conn.close()
                
                if not ot_raw.empty:
                    today_str = date.today().strftime("%Y-%m-%d")
                    tomorrow_str = (date.today() + timedelta(days=1)).strftime("%Y-%m-%d")
                    
                    for _, o_row in ot_raw.iterrows():
                        o_id, o_uname, o_date, o_emp, o_start, o_end, o_dep = o_row['id'], o_row['username'], o_row['ot_date'], o_row['emp_name'], o_row['start_time'], o_row['end_time'], o_row['departure_time']
                        o_col1, o_col2, o_col3, o_col4 = st.columns([3, 2, 1, 1])
                        with o_col1:
                            st.text(f"ID #{o_id} | Date: {o_date} | Staff: {o_emp}")
                        with o_col2:
                            st.text(f"Time: {o_start} - {o_end}")
                        with o_col3:
                            if o_date in [today_str, tomorrow_str]:
                                if st.button("🚗 Transport", key=f"transport_ot_{o_id}", type="secondary"):
                                    try:
                                        conn = get_db_connection()
                                        existing_dt = conn.execute(
                                            "SELECT id FROM daily_transit WHERE transit_date_start = ? AND trip = 'Trip D'",
                                            (o_date,)
                                        ).fetchone()
                                        
                                        if not existing_dt:
                                            conn.execute('''
                                                INSERT INTO daily_transit (transit_date_start, transit_date_end, group_name, requested_by, etd_1, etd_2, location_from, location_to, daily, trip)
                                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                            ''', (o_date, o_date, "TBA", o_uname, o_dep if o_dep else "19:00", None, "Yard-1 Office", "Panbil", "No", "Trip D"))
                                            conn.commit()
                                            set_transaction_dialog("Transport Request Added", f"Daily transit dispatch (Trip D) created for date {o_date}.", "success")
                                        else:
                                            set_transaction_dialog("Already Exists", f"A Trip D transit dispatch for date {o_date} already exists.", "info")
                                        conn.close()
                                    except Exception as e:
                                        set_transaction_dialog("Action Unsuccessful", f"Failed: {str(e)}", "error")
                                    st.rerun()
                            else:
                                st.text("Past date")
                        with o_col4:
                            if st.button("🗑️ Remove", key=f"del_ot_{o_id}", type="primary"):
                                try:
                                    conn = get_db_connection()
                                    conn.execute("DELETE FROM overtime_requests WHERE id = ?", (o_id,))
                                    conn.commit()
                                    conn.close()
                                    set_transaction_dialog("Deletion Successful", f"Removed overtime record #{o_id}.", "success")
                                except Exception as e:
                                    set_transaction_dialog("Deletion Unsuccessful", f"Failed: {str(e)}", "error")
                                st.rerun()
                else:
                    st.info("No overtime submissions found.")

        st.download_button("📥 Export Overtime Log to Excel (.xlsx)", data=export_df_to_excel(ot_df, sheet_name="Overtime_Requests"), file_name=f"Overtime_Requests_{datetime.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 4. TRANSIT GROUPS & PASSENGERS
elif nav_selection == "👥 Transit Groups & Passengers":
    st.header("👥 Transit Groups & Passengers Management")
    conn = get_db_connection()
    drivers_list = [d['driver_name'] for d in conn.execute("SELECT driver_name FROM fleet_drivers").fetchall()]
    cars_list = [c['car_name'] for c in conn.execute("SELECT car_name FROM cars").fetchall()]
    employees_list = [u['emp_name'] for u in conn.execute("SELECT emp_name FROM users WHERE emp_name IS NOT NULL AND emp_name != ''").fetchall()]
    groups_list = [g['group_name'] for g in conn.execute("SELECT group_name FROM transit_groups").fetchall()]
    conn.close()

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Create Transit Group")
        selected_car_group = st.selectbox("Group Name (Select from Cars)", cars_list if cars_list else ["No cars available"], key="g_car_select")
        selected_driver = st.selectbox("Assign Driver (from Fleet Drivers)", drivers_list if drivers_list else ["No drivers available"])
        raw_etd1_grp = st.text_input("ETD 1 (From) [e.g. 0545 or 05:45]", value="05:45", key="grp_etd1_in")
        raw_etd2_grp = st.text_input("ETD 2 (To) [e.g. 1730 or 17:30]", value="17:30", key="grp_etd2_in")
        
        if st.button("Save Transit Group"):
            etd1_grp_formatted, etd2_grp_formatted = format_military_time(raw_etd1_grp), format_military_time(raw_etd2_grp)
            if not etd1_grp_formatted or not etd2_grp_formatted:
                set_transaction_dialog("Data Transaction Unsuccessful", "Invalid time format.", "error")
            else:
                try:
                    conn = get_db_connection()
                    conn.execute("INSERT INTO transit_groups (group_name, driver_name, etd_1, etd_2) VALUES (?, ?, ?, ?)", (selected_car_group, selected_driver, etd1_grp_formatted, etd2_grp_formatted))
                    conn.commit()
                    conn.close()
                    set_transaction_dialog("Data Transaction Successful", f"Group '{selected_car_group}' created.", "success")
                except Exception as e:
                    set_transaction_dialog("Data Transaction Unsuccessful", f"Failed: {str(e)}", "error")
            st.rerun()

    with col2:
        st.subheader("2. Assign Passengers to Group")
        if groups_list and employees_list:
            selected_group_for_p = st.selectbox("Select Group Name", groups_list, key="p_group_select")
            selected_passengers = st.multiselect("Select Passenger(s)", employees_list)
            if st.button("Assign Passengers to Group"):
                if selected_passengers:
                    try:
                        conn = get_db_connection()
                        for p in selected_passengers:
                            conn.execute("INSERT INTO transit_passengers (group_name, passengers) VALUES (?, ?)", (selected_group_for_p, p))
                        conn.commit()
                        conn.close()
                        set_transaction_dialog("Data Transaction Successful", "Passengers assigned.", "success")
                    except Exception as e:
                        set_transaction_dialog("Data Transaction Unsuccessful", f"Failed: {str(e)}", "error")
                st.rerun()

    st.markdown("---")
    title_col, eff_date_col = st.columns([2, 1])
    with title_col:
        st.subheader("📋 Configured Groups & Assigned Passengers")
    with eff_date_col:
        target_effective_date = st.date_input("Target Effective Date", value=date.today(), key="eff_date_picker")
        eff_date_str = target_effective_date.strftime("%Y-%m-%d") if target_effective_date else ""

    conn = get_db_connection()
    unrolled_df = pd.read_sql_query('''
        SELECT tg.group_name AS "Car Group", c.vehicle AS "Vehicle Model", c.plate_number AS "Plate Number", 
               c.color AS "Color", tg.driver_name AS "Driver Name", fd.driver_mobile AS "Contact Number", 
               tg.etd_1 AS "ETD 1 (From)", tg.etd_2 AS "ETD 2 (To)", tp.passengers AS "Passenger Name"
        FROM transit_groups tg
        LEFT JOIN cars c ON tg.group_name = c.car_name
        LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
        LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
        ORDER BY tg.group_name
    ''', conn)
    conn.close()
    
    if not unrolled_df.empty:
        st.dataframe(unrolled_df, use_container_width=True)
        
        with st.expander("✏️ Manage / Remove Passenger Assignments"):
            conn = get_db_connection()
            passengers_raw = pd.read_sql_query("SELECT id, group_name, passengers FROM transit_passengers", conn)
            conn.close()
            
            if not passengers_raw.empty:
                for _, p_row in passengers_raw.iterrows():
                    p_id, p_grp, p_name = p_row['id'], p_row['group_name'], p_row['passengers']
                    p_col1, p_col2, p_col3 = st.columns([2, 2, 1])
                    with p_col1:
                        st.text(f"Group: {p_grp}")
                    with p_col2:
                        st.text(f"Passenger: {p_name}")
                    with p_col3:
                        if st.button("🗑️ Remove", key=f"del_passenger_{p_id}", type="primary"):
                            try:
                                conn = get_db_connection()
                                conn.execute("DELETE FROM transit_passengers WHERE id = ?", (p_id,))
                                conn.commit()
                                conn.close()
                                set_transaction_dialog("Deletion Successful", f"Removed passenger '{p_name}' from group '{p_grp}'.", "success")
                            except Exception as e:
                                set_transaction_dialog("Deletion Unsuccessful", f"Failed: {str(e)}", "error")
                            st.rerun()
            else:
                st.info("No passenger assignments found.")

        col_ex, col_pdf, col_wa = st.columns(3)
        with col_ex:
            st.download_button(
                "📥 Export to Excel (.xlsx)", 
                data=export_custom_batam_excel(unrolled_df, effective_date_str=eff_date_str), 
                file_name=f"Daily_Transportation_Arrangement_TUCC_{eff_date_str}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True
            )
        with col_pdf:
            if HAS_REPORTLAB:
                st.download_button(
                    "📄 Export to PDF (.pdf)", 
                    data=export_custom_batam_pdf(unrolled_df, effective_date_str=eff_date_str), 
                    file_name=f"Daily_Transportation_Arrangement_TUCC_{eff_date_str}.pdf", 
                    mime="application/pdf", 
                    use_container_width=True
                )
        with col_wa:
            wa_message = urllib.parse.quote(
                f"📢 *TUCC PJ Batam - Daily Transportation Arrangement*\n"
                f"📅 Effective Date: {eff_date_str}\n"
                f"Please find the attached passenger layout schedule above."
            )
            wa_url = f"https://wa.me/?text={wa_message}"
            st.link_button(
                "💬 Share on WhatsApp", 
                url=wa_url, 
                use_container_width=True
            )

# 5. DAILY TRANSIT DISPATCH SETUP
elif nav_selection == "📅 Daily Transit Dispatch Setup":
    st.header("📅 Daily Transit Dispatch Schedule & Route Setting")
    conn = get_db_connection()
    all_emp_names = pd.read_sql_query("SELECT emp_name FROM users WHERE emp_name IS NOT NULL AND emp_name != ''", conn)['emp_name'].tolist()
    cars_db_df = pd.read_sql_query("SELECT car_name, plate_number, vehicle FROM cars", conn)
    trips_db_df = pd.read_sql_query("SELECT trip, trip_name FROM trips ORDER BY trip", conn)
    origins_df = pd.read_sql_query("SELECT DISTINCT location_from FROM daily_transit WHERE location_from IS NOT NULL", conn)
    dests_df = pd.read_sql_query("SELECT DISTINCT location_to FROM daily_transit WHERE location_to IS NOT NULL", conn)
    conn.close()
    
    default_locations = ["Yard-1 Office", "Yard-3 Office", "Panbil", "Batam Center", "Hang Nadim Airport"]
    origin_list = sorted(list(set(default_locations + origins_df['location_from'].tolist())))
    dest_list = sorted(list(set(default_locations + dests_df['location_to'].tolist())))
    render_datalist_options("origin_list_dl", origin_list)
    render_datalist_options("dest_list_dl", dest_list)

    current_user_emp = st.session_state.get("emp_name", "") or st.session_state.get("username", "")
    if not all_emp_names:
        all_emp_names = [current_user_emp] if current_user_emp else ["Default Employee"]
    default_req_by = current_user_emp if current_user_emp in all_emp_names else all_emp_names[0]
    is_admin_or_owner = st.session_state.get("role", "") in ["Admin", "Owner"]

    trip_option_labels = [f"{r['trip']} ({r['trip_name']})" for _, r in trips_db_df.iterrows()] if not trips_db_df.empty else ["Trip A (Yard to Yard)"]
    trip_code_map = {lbl: lbl.split(" (")[0] for lbl in trip_option_labels}

    car_option_labels = ["TBA - To Be Assigned"] + [f"{r['car_name']} - {r['plate_number']}" + (f" ({r['vehicle']})" if r['vehicle'] else "") for _, r in cars_db_df.iterrows() if r['car_name'] != 'TBA']
    car_label_to_group = {"TBA - To Be Assigned": "TBA"}
    for _, r in cars_db_df.iterrows():
        lbl = f"{r['car_name']} - {r['plate_number']}" + (f" ({r['vehicle']})" if r['vehicle'] else "")
        car_label_to_group[lbl] = r['car_name']

    if "dispatch_reset_counter" not in st.session_state:
        st.session_state.dispatch_reset_counter = 0
    reset_id = st.session_state.dispatch_reset_counter

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Configure Transit Request")
        requested_by = st.selectbox("Requested By", options=all_emp_names, index=all_emp_names.index(default_req_by), key=f"dt_req_{reset_id}")
        selected_trip_label = st.selectbox("Trip Category", options=trip_option_labels, index=0, key=f"dt_trip_{reset_id}")
        selected_trip_code = trip_code_map.get(selected_trip_label, "Trip A")

        is_daily = st.selectbox("Daily / Recurring Journey?", options=["No", "Yes"] if is_admin_or_owner else ["No"], index=0, disabled=not is_admin_or_owner, key=f"dt_daily_{reset_id}")
        dispatch_date_start = st.date_input("Select Transit Start Date", value=date.today(), min_value=date.today(), key=f"reg_dt_start_{reset_id}")
        
        if is_daily == "Yes":
            dispatch_date_end = st.date_input("Select Transit End Date", value=dispatch_date_start + timedelta(days=1), min_value=dispatch_date_start + timedelta(days=1), key=f"reg_dt_end_{reset_id}")
        else:
            dispatch_date_end = dispatch_date_start

        location_from = st.text_input("Origin Location", value="Yard-1 Office" if reset_id == 0 else "", placeholder="[origin_list_dl] Type origin...", key=f"txt_origin_{reset_id}")
        location_to = st.text_input("Target Location", value="Yard-3 Office" if reset_id == 0 else "", placeholder="[dest_list_dl] Type target...", key=f"txt_dest_{reset_id}")

    with col2:
        st.subheader("Schedule & Vehicle Allocation")
        raw_etd1 = st.text_input("ETD 1 (Start Time)", value="08:00" if reset_id == 0 else "", placeholder="0800", key=f"etd1_{reset_id}")
        raw_etd2 = st.text_input("ETD 2 (Return Time)", value="17:00" if reset_id == 0 else "", placeholder="1700", key=f"etd2_{reset_id}")
        selected_car_label = st.selectbox("Assigned Group / Car Name", options=car_option_labels, index=0, disabled=not is_admin_or_owner, key=f"car_{reset_id}")
        selected_group = car_label_to_group.get(selected_car_label, "TBA") if is_admin_or_owner else "TBA"

    b_col1, b_col2 = st.columns([1, 4])
    with b_col1:
        if st.button("Submit Transit Request", type="primary"):
            fmt_etd1, fmt_etd2 = format_military_time(raw_etd1), format_military_time(raw_etd2)
            if not fmt_etd1 or not fmt_etd2:
                set_transaction_dialog("Data Transaction Unsuccessful", "Invalid ETD Time format.", "error")
            else:
                try:
                    conn = get_db_connection()
                    conn.execute('''
                        INSERT INTO daily_transit (transit_date_start, transit_date_end, group_name, requested_by, etd_1, etd_2, location_from, location_to, daily, trip)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (dispatch_date_start.strftime("%Y-%m-%d"), dispatch_date_end.strftime("%Y-%m-%d"), selected_group, requested_by, fmt_etd1, fmt_etd2, location_from.strip(), location_to.strip(), is_daily, selected_trip_code))
                    conn.commit()
                    conn.close()
                    set_transaction_dialog("Data Transaction Successful", "Transit dispatch logged successfully.", "success")
                except Exception as e:
                    set_transaction_dialog("Data Transaction Unsuccessful", f"Error: {str(e)}", "error")
            st.rerun()
    with b_col2:
        if st.button("🧹 Clear Form Inputs"):
            st.session_state.dispatch_reset_counter += 1
            st.rerun()

    st.markdown("---")
    st.subheader("📊 Scheduled Transit Dispatches Log")
    conn = get_db_connection()
    daily_raw_df = pd.read_sql_query('''
        SELECT dt.id, dt.transit_date_start, dt.transit_date_end, 
               t.trip || ' (' || t.trip_name || ')' AS trip_display, 
               dt.trip AS trip_code,
               dt.requested_by, dt.group_name, c.plate_number, c.vehicle, 
               dt.location_from, dt.location_to, dt.etd_1, dt.etd_2, dt.daily,
               tg.driver_name
        FROM daily_transit dt 
        LEFT JOIN cars c ON dt.group_name = c.car_name
        LEFT JOIN trips t ON dt.trip = t.trip
        LEFT JOIN transit_groups tg ON dt.group_name = tg.group_name
        ORDER BY dt.transit_date_start DESC, dt.id DESC
    ''', conn)
    conn.close()
    
    if not daily_raw_df.empty:
        display_df = daily_raw_df.copy()
        display_df['Group / Car'] = display_df.apply(lambda r: f"{r['group_name']} - {r['plate_number']}" if pd.notna(r['plate_number']) and r['group_name'] != 'TBA' else r['group_name'], axis=1)
        export_df = display_df[["id", "transit_date_start", "transit_date_end", "trip_display", "requested_by", "Group / Car", "location_from", "location_to", "etd_1", "etd_2", "daily"]].rename(columns={"id": "Dispatch ID", "transit_date_start": "Start Date", "transit_date_end": "End Date", "trip_display": "Trip Category", "requested_by": "Requested By", "location_from": "Origin", "location_to": "Destination", "etd_1": "ETD Start", "etd_2": "ETD Return", "daily": "Recurring"})
        st.dataframe(export_df, use_container_width=True)
        
        with st.expander("✏️ Manage / Remove Scheduled Dispatches"):
            conn = get_db_connection()
            dispatches_raw = pd.read_sql_query("SELECT id, transit_date_start, requested_by, location_from, location_to, trip FROM daily_transit ORDER BY id DESC", conn)
            conn.close()
            
            if not dispatches_raw.empty:
                for _, d_row in dispatches_raw.iterrows():
                    d_id, d_date, d_req, d_from, d_to, d_trip = d_row['id'], d_row['transit_date_start'], d_row['requested_by'], d_row['location_from'], d_row['location_to'], d_row['trip']
                    d_col1, d_col2, d_col3 = st.columns([3, 2, 1])
                    with d_col1:
                        st.text(f"ID #{d_id} | Date: {d_date} | {d_from} ➔ {d_to}")
                    with d_col2:
                        st.text(f"Req: {d_req} | Trip: {d_trip}")
                    with d_col3:
                        if st.button("🗑️ Remove", key=f"del_dispatch_{d_id}", type="primary"):
                            try:
                                conn = get_db_connection()
                                conn.execute("DELETE FROM daily_transit WHERE id = ?", (d_id,))
                                conn.commit()
                                conn.close()
                                set_transaction_dialog("Deletion Successful", f"Removed dispatch record #{d_id}.", "success")
                            except Exception as e:
                                set_transaction_dialog("Deletion Unsuccessful", f"Failed: {str(e)}", "error")
                            st.rerun()
            else:
                st.info("No dispatch records found.")

        st.download_button("📥 Export Dispatch Log to Excel (.xlsx)", data=export_df_to_excel(export_df, sheet_name="Daily_Dispatches"), file_name=f"Daily_Dispatch_Schedule_{datetime.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 6. MEETING ROOMS
elif nav_selection == "🏢 Meeting Rooms":
    st.header("Meeting Space Reservations Desk")
    conn = get_db_connection()
    rooms = conn.execute("SELECT * FROM meeting_rooms").fetchall()
    conn.close()
    
    if rooms:
        room_options = {f"{r['room_name']} (Room {r['room_number']} - Capacity: {r['capacity']})": r['room_number'] for r in rooms}
        selected_room_num = room_options[st.selectbox("Choose Target Room Venue", list(room_options.keys()))]
        
        col1, col2 = st.columns(2)
        with col1:
            book_date = st.date_input("Reservation Date", value=date.today())
            b_start, b_end = st.time_input("Start Time", value=time(9, 0)), st.time_input("End Time", value=time(10, 0))
        with col2:
            recurrence = st.selectbox("Recurrence Pattern", ["None", "Daily", "Weekly", "Monthly"])
            recurrence_end = st.date_input("Recurrence End Target", value=book_date + timedelta(days=7))

        if st.button("Confirm Room Block Assignment"):
            target_dates = [book_date]
            if recurrence != "None":
                curr = book_date
                while True:
                    curr += timedelta(days=1 if recurrence == "Daily" else 7 if recurrence == "Weekly" else 30)
                    if curr <= recurrence_end:
                        target_dates.append(curr)
                    else:
                        break
            
            conn = get_db_connection()
            conflict = False
            for td in target_dates:
                if conn.execute("SELECT * FROM room_bookings WHERE room_number = ? AND booking_date = ? AND NOT (start_time >= ? OR end_time <= ?)", (selected_room_num, td.strftime("%Y-%m-%d"), b_end.strftime("%H:%M"), b_start.strftime("%H:%M"))).fetchone():
                    conflict = True
                    break
            
            if conflict:
                set_transaction_dialog("Data Transaction Unsuccessful", "Schedule conflict detected.", "error")
            else:
                for td in target_dates:
                    conn.execute("INSERT INTO room_bookings (room_number, booked_by, booking_date, start_time, end_time, is_recurring, recurrence_end_date) VALUES (?, ?, ?, ?, ?, ?, ?)", (selected_room_num, st.session_state.username, td.strftime("%Y-%m-%d"), b_start.strftime("%H:%M"), b_end.strftime("%H:%M"), recurrence, recurrence_end.strftime("%Y-%m-%d")))
                conn.commit()
                set_transaction_dialog("Data Transaction Successful", f"Room reserved for {len(target_dates)} date(s).", "success")
            conn.close()
            st.rerun()

    st.subheader("📊 Master Room Allocation Schedules")
    conn = get_db_connection()
    bookings_df = pd.read_sql_query("SELECT b.id, r.room_name, b.room_number, b.booked_by, b.booking_date, b.start_time, b.end_time, b.is_recurring FROM room_bookings b JOIN meeting_rooms r ON b.room_number = r.room_number", conn)
    conn.close()
    if not bookings_df.empty:
        st.dataframe(bookings_df, use_container_width=True)

# 7. SITE NEWS
elif nav_selection == "📢 Site News":
    st.header("📢 TUCC PJ Batam Site News & Announcements")
    is_admin_or_owner = st.session_state.get("role", "") in ["Admin", "Owner"]
    
    if is_admin_or_owner:
        with st.expander("📤 Upload New Site News Bulletin (Admin / Owner Only)", expanded=False):
            with st.form("news_upload_form", clear_on_submit=True):
                news_title = st.text_input("News Title / Description")
                uploaded_news_file = st.file_uploader("Upload Bulletin File (PDF or PNG only)", type=["pdf", "png"])
                news_submitted = st.form_submit_button("Publish Bulletin")
                
                if news_submitted:
                    if not news_title.strip():
                        set_transaction_dialog("Upload Unsuccessful", "Please provide a title for the news bulletin.", "error")
                    elif uploaded_news_file is None:
                        set_transaction_dialog("Upload Unsuccessful", "Please select a file to upload.", "error")
                    else:
                        file_ext = uploaded_news_file.name.split('.')[-1].lower()
                        if file_ext not in ["pdf", "png"]:
                            set_transaction_dialog("Upload Unsuccessful", "Only PDF and PNG files are allowed.", "error")
                        else:
                            try:
                                timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                                safe_filename = f"{timestamp_str}_{uploaded_news_file.name}"
                                file_path = os.path.join(NEWS_DIR, safe_filename)
                                
                                with open(file_path, "wb") as f:
                                    f.write(uploaded_news_file.getbuffer())
                                    
                                conn = get_db_connection()
                                conn.execute('''
                                    INSERT INTO site_news (title, filename, file_path, file_type, uploaded_by, upload_date)
                                    VALUES (?, ?, ?, ?, ?, ?)
                                ''', (news_title.strip(), uploaded_news_file.name, file_path, file_ext, st.session_state.username, datetime.now().strftime("%Y-%m-%d %H:%M")))
                                conn.commit()
                                conn.close()
                                
                                set_transaction_dialog("Upload Successful", f"Bulletin '{news_title}' published successfully.", "success")
                            except Exception as e:
                                set_transaction_dialog("Upload Unsuccessful", f"Failed to save file: {str(e)}", "error")
                            st.rerun()

    st.subheader("📋 Published Site News Bulletins (Latest to Oldest)")
    conn = get_db_connection()
    news_df = pd.read_sql_query("SELECT * FROM site_news ORDER BY id DESC", conn)
    conn.close()
    
    if not news_df.empty:
        for _, row in news_df.iterrows():
            news_id, title, filename, file_path, file_type, uploaded_by, upload_date = row['id'], row['title'], row['filename'], row['file_path'], row['file_type'], row['uploaded_by'], row['upload_date']
            
            with st.container(border=True):
                col_info, col_action = st.columns([3, 1])
                with col_info:
                    st.markdown(f"### 📌 {title}")
                    st.caption(f"📅 Published: **{upload_date}** | 👤 By: **{uploaded_by}** | 📄 File: `{filename}`")
                with col_action:
                    if os.path.exists(file_path):
                        with open(file_path, "rb") as f:
                            file_bytes = f.read()
                        st.download_button(
                            label=f"📥 View / Download ({file_type.upper()})",
                            data=file_bytes, file_name=filename,
                            mime="application/pdf" if file_type == "pdf" else "image/png",
                            key=f"dl_news_{news_id}"
                        )
                    else:
                        st.error("File missing.")
                if is_admin_or_owner:
                    if st.button(f"🗑️ Delete Bulletin #{news_id}", key=f"del_news_{news_id}"):
                        try:
                            if os.path.exists(file_path):
                                os.remove(file_path)
                            conn = get_db_connection()
                            conn.execute("DELETE FROM site_news WHERE id = ?", (news_id,))
                            conn.commit()
                            conn.close()
                            set_transaction_dialog("Deletion Successful", "Bulletin deleted.", "success")
                        except Exception as e:
                            set_transaction_dialog("Deletion Unsuccessful", f"Failed: {str(e)}", "error")
                        st.rerun()
    else:
        st.info("No site news bulletins published yet.")

# 8. SYSTEM ADMINISTRATION
elif nav_selection == "🛠️ System Administration":
    current_role = st.session_state.get("role", "")
    if current_role not in ["Admin", "Owner"]:
        st.error("🛡️ Restricted Access Control: Admin or Owner clearance required.")
    else:
        st.header("Admin Control Dashboard Engine")
        
        st.markdown("---")
        st.subheader("💾 Database Backups, Restores & Disaster Recovery")
        b_col1, b_col2 = st.columns(2)
        with b_col1:
            if st.button("🔄 Trigger Manual Backup Now"):
                perform_automated_backup()
                set_transaction_dialog("Data Transaction Successful", "Manual backup snapshot created successfully.", "success")
                st.rerun()
        with b_col2:
            backup_files = sorted(os.listdir(BACKUP_DIR), reverse=True) if os.path.exists(BACKUP_DIR) else []
            if backup_files:
                selected_backup = st.selectbox("Select Backup Snapshot", backup_files)
                if selected_backup:
                    with open(os.path.join(BACKUP_DIR, selected_backup), "rb") as f:
                        st.download_button(f"📥 Download Selected Backup ({selected_backup})", data=f, file_name=selected_backup, mime="application/octet-stream")
            else:
                st.info("No backup snapshots found yet.")

        if current_role == "Owner":
            with st.expander("📥 Restore Database from Backup File (.db)", expanded=False):
                st.warning("⚠️ **Caution:** Uploading and restoring a database backup will overwrite the current live database file (`office_operations.db`). A safety backup of your current database will be created automatically before the restore takes place.")
                uploaded_db_file = st.file_uploader("Upload Database Backup (.db)", type=["db"])
                
                if uploaded_db_file is not None:
                    if st.button("🚀 Confirm and Restore Database", type="primary"):
                        try:
                            if os.path.exists(DB_FILE):
                                safety_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                shutil.copy2(DB_FILE, os.path.join(BACKUP_DIR, f"pre_restore_safety_{safety_timestamp}.db"))
                            
                            with open(DB_FILE, "wb") as f:
                                f.write(uploaded_db_file.getbuffer())
                                
                            set_transaction_dialog("Database Restored Successfully", "The database has been successfully replaced and restored from the backup file.", "success")
                        except Exception as e:
                            set_transaction_dialog("Restore Unsuccessful", f"Failed to restore database: {str(e)}", "error")
                        st.rerun()
        else:
            st.info("ℹ️ Note: **Restore Database from Backup** is restricted to **Owner** accounts only.")

        st.markdown("---")
        st.subheader("📤 Bulk Import Data via Excel (.xlsx)")
        import_table = st.selectbox("Select Database Table", ["users", "holidays", "trips", "fleet_drivers", "cars", "transit_passengers", "site_news"], key="import_tbl_sel")
        
        table_schemas = {
            "users": ["username", "password", "role", "email_recipients", "emp_name"],
            "holidays": ["holiday_date", "description"],
            "trips": ["trip", "trip_name"],
            "fleet_drivers": ["driver_name", "driver_mobile"],
            "cars": ["car_name", "plate_number", "vehicle", "color"],
            "transit_passengers": ["group_name", "passengers"],
            "site_news": ["title", "filename", "file_path", "file_type", "uploaded_by", "upload_date"]
        }
        
        req_cols = table_schemas[import_table]
        buffer_template = io.BytesIO()
        pd.DataFrame(columns=req_cols).to_excel(buffer_template, index=False, sheet_name=f"{import_table}_Template")
        st.download_button(f"📥 Download Excel Template for {import_table}", data=buffer_template.getvalue(), file_name=f"{import_table}_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        uploaded_file = st.file_uploader(f"Upload Excel for '{import_table}'", type=["xlsx", "xls"])
        if uploaded_file is not None:
            try:
                import_df = pd.read_excel(uploaded_file)
                import_df.columns = [str(c).strip().lower() for c in import_df.columns]
                if all(c in import_df.columns for c in req_cols):
                    st.dataframe(import_df[req_cols], use_container_width=True)
                    if st.button(f"🚀 Import Records into '{import_table}'"):
                        conn = get_db_connection()
                        cursor = conn.cursor()
                        for _, row in import_df.iterrows():
                            vals = [None if pd.isna(row[c]) else str(row[c]).strip() for c in req_cols]
                            cursor.execute(f"INSERT OR REPLACE INTO {import_table} ({', '.join(req_cols)}) VALUES ({', '.join(['?']*len(req_cols))})", vals)
                        conn.commit()
                        conn.close()
                        set_transaction_dialog("Data Transaction Successful", "Import completed.", "success")
                        st.rerun()
                else:
                    st.error("Missing required column headers.")
            except Exception as e:
                st.error(f"Error: {str(e)}")

        st.markdown("---")
        st.subheader("🗃️ Master Data Tables Inline CRUD Editor")
        selected_table = st.selectbox("Choose Database Table to Manage", ["users", "holidays", "trips", "overtime_requests", "meeting_rooms", "room_bookings", "fleet_drivers", "cars", "transit_groups", "transit_passengers", "daily_transit", "site_news"])
        
        conn = get_db_connection()
        table_df = pd.read_sql_query(f"SELECT * FROM {selected_table}", conn)
        original_passwords = dict(zip(table_df["username"], table_df["password"])) if selected_table == "users" else {}
        if selected_table == "users":
            table_df["password"] = "••••••••"
        
        drivers_list = [d['driver_name'] for d in conn.execute("SELECT driver_name FROM fleet_drivers").fetchall()]
        cars_list = [c['car_name'] for c in conn.execute("SELECT car_name FROM cars").fetchall()]
        groups_list = [g['group_name'] for g in conn.execute("SELECT group_name FROM transit_groups").fetchall()]
        emp_list = [u['emp_name'] for u in conn.execute("SELECT emp_name FROM users WHERE emp_name IS NOT NULL").fetchall()]
        conn.close()
        
        is_owner = (current_role == "Owner")
        if is_owner:
            table_df.insert(0, "🗑️ Delete", False)
            column_config = {
                "🗑️ Delete": st.column_config.CheckboxColumn("Delete?", help="Check to delete this specific row", default=False)
            }
        else:
            column_config = {}

        if selected_table == "transit_groups":
            column_config["group_name"] = st.column_config.SelectboxColumn("Group Name", options=cars_list)
            column_config["driver_name"] = st.column_config.SelectboxColumn("Driver Name", options=drivers_list)
        elif selected_table == "transit_passengers":
            column_config["group_name"] = st.column_config.SelectboxColumn("Group Name", options=groups_list)
            column_config["passengers"] = st.column_config.SelectboxColumn("Passenger", options=emp_list)

        edited_df = st.data_editor(table_df, num_rows="dynamic", use_container_width=True, column_config=column_config, key=f"editor_{selected_table}")
        
        save_button_label = f"Save Grid Changes & Delete Selected ({selected_table})" if is_owner else f"Save Grid Changes ({selected_table})"
        if st.button(save_button_label):
            try:
                conn = get_db_connection()
                cursor = conn.cursor()
                cursor.execute(f"DELETE FROM {selected_table}")
                
                rows_to_save = edited_df.copy()
                if is_owner and "🗑️ Delete" in rows_to_save.columns:
                    rows_to_save = rows_to_save[rows_to_save["🗑️ Delete"] != True]
                    rows_to_save = rows_to_save.drop(columns=["🗑️ Delete"])
                
                for _, row in rows_to_save.iterrows():
                    row_dict = row.to_dict()
                    if selected_table == "users" and row_dict.get("password") == "••••••••":
                        row_dict["password"] = original_passwords.get(row_dict.get("username"), "")
                    cols = [k for k in row_dict.keys() if row_dict[k] is not None and str(row_dict[k]) != "nan"]
                    cursor.execute(f"INSERT INTO {selected_table} ({', '.join(cols)}) VALUES ({', '.join(['?']*len(cols))})", [row_dict[k] for k in cols])
                
                conn.commit()
                conn.close()
                set_transaction_dialog("Data Transaction Successful", f"Table '{selected_table}' updated successfully.", "success")
            except Exception as e:
                set_transaction_dialog("Data Transaction Unsuccessful", f"Failed: {str(e)}", "error")
            st.rerun()
