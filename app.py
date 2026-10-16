import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import plotly.express as px
import streamlit as st

# WeasyPrint check for PDF generation
try:
  from weasyprint import HTML

  WEASYPRINT_AVAILABLE = True
except Exception:
  WEASYPRINT_AVAILABLE = False

# Page Setup
st.set_page_config(
    page_title="E-Mojani Filtered Pending Report Generator & Dashboard",
    layout="wide",
)

# Custom Styling (CSS) for Streamlit UI
st.markdown(
    """
    <style>
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 12px;
        padding: 18px;
        border-left: 6px solid #2b9348;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-title { font-size: 14px; color: #555; font-weight: 600; }
    .metric-value { font-size: 26px; color: #1b4332; font-weight: bold; margin-top: 5px; }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("📊 भूमि अभिलेख विभाग - अमरावती | प्रलंबित मोजणी अहवाल व डॅशबोर्ड")


# EXCEL HELPER FOR INDIVIDUAL SHEETS (WITH TITLE HEADER)
def convert_df_to_excel(df, report_title="प्रलंबित मोजणी अहवाल"):
  output = io.BytesIO()
  current_date_str = datetime.date.today().strftime("%d/%m/%Y")

  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Report"

  font_dept = Font(name="Calibri", size=14, bold=True, color="1F4E78")
  font_sub = Font(name="Calibri", size=11, bold=True, color="000000")
  font_th = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  fill_th = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )

  thin_border = Border(
      left=Side(style="thin", color="CCCCCC"),
      right=Side(style="thin", color="CCCCCC"),
      top=Side(style="thin", color="CCCCCC"),
      bottom=Side(style="thin", color="CCCCCC"),
  )

  is_multi = isinstance(df.index, pd.MultiIndex)
  num_idx_cols = len(df.index.names) if is_multi else 1
  total_cols = num_idx_cols + len(df.columns)

  # Row 1: Department Title
  ws.merge_cells(
      start_row=1, start_column=1, end_row=1, end_column=max(total_cols, 2)
  )
  cell_dept = ws.cell(row=1, column=1, value="भूमि अभिलेख विभाग - अमरावती")
  cell_dept.font = font_dept
  cell_dept.alignment = Alignment(horizontal="center", vertical="center")

  # Row 2: Subtitle with Date
  ws.merge_cells(
      start_row=2, start_column=1, end_row=2, end_column=max(total_cols, 2)
  )
  cell_sub = ws.cell(
      row=2, column=1, value=f"{report_title} ( दिनांक {current_date_str} )"
  )
  cell_sub.font = font_sub
  cell_sub.alignment = Alignment(horizontal="center", vertical="center")

  # Table Headers
  start_row = 4
  if is_multi:
    for i, idx_name in enumerate(df.index.names):
      cell = ws.cell(row=start_row, column=i + 1, value=idx_name or "")
      cell.font, cell.fill = font_th, fill_th
      cell.alignment = Alignment(horizontal="center", vertical="center")
  else:
    cell = ws.cell(
        row=start_row, column=1, value=df.index.name or "तालुका"
    )
    cell.font, cell.fill = font_th, fill_th
    cell.alignment = Alignment(horizontal="center", vertical="center")

  for j, col_name in enumerate(df.columns):
    cell = ws.cell(
        row=start_row, column=num_idx_cols + j + 1, value=str(col_name)
    )
    cell.font, cell.fill = font_th, fill_th
    cell.alignment = Alignment(horizontal="center", vertical="center")

  # Rows Data
  current_r = start_row + 1
  for idx_val, row in df.iterrows():
    if is_multi:
      for k, val in enumerate(idx_val):
        cell = ws.cell(row=current_r, column=k + 1, value=str(val))
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = thin_border
    else:
      cell = ws.cell(row=current_r, column=1, value=str(idx_val))
      cell.alignment = Alignment(horizontal="left", vertical="center")
      cell.border = thin_border

    for j, col_name in enumerate(df.columns):
      val = row[col_name]
      val_clean = (
          ""
          if (
              pd.isna(val)
              or val == 0
              or str(val).strip() in ["", "0", "nan", "0.0"]
          )
          else (int(val) if isinstance(val, (int, float)) else val)
      )
      cell = ws.cell(
          row=current_r, column=num_idx_cols + j + 1, value=val_clean
      )
      cell.alignment = Alignment(horizontal="center", vertical="center")
      cell.border = thin_border

    current_r += 1

  # Auto Column Width
  for col in ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = openpyxl.utils.get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

  wb.save(output)
  return output.getvalue()


# MASTER EXCEL HELPER: COMBINES ALL REPORTS INTO SINGLE EXCEL
def convert_all_reports_to_single_excel(df_dict):
  output = io.BytesIO()
  current_date_str = datetime.date.today().strftime("%d/%m/%Y")
  wb = openpyxl.Workbook()
  wb.remove(wb.active)  # Remove default active sheet

  font_dept = Font(name="Calibri", size=14, bold=True, color="1F4E78")
  font_sub = Font(name="Calibri", size=11, bold=True, color="000000")
  font_th = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  fill_th = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )

  thin_border = Border(
      left=Side(style="thin", color="CCCCCC"),
      right=Side(style="thin", color="CCCCCC"),
      top=Side(style="thin", color="CCCCCC"),
      bottom=Side(style="thin", color="CCCCCC"),
  )

  for sheet_name, (df_data, report_title) in df_dict.items():
    ws = wb.create_sheet(title=sheet_name[:31])

    is_multi = isinstance(df_data.index, pd.MultiIndex)
    num_idx_cols = len(df_data.index.names) if is_multi else 1
    total_cols = num_idx_cols + len(df_data.columns)

    # Title Headers
    ws.merge_cells(
        start_row=1, start_column=1, end_row=1, end_column=max(total_cols, 2)
    )
    c1 = ws.cell(row=1, column=1, value="भूमि अभिलेख विभाग - अमरावती")
    c1.font, c1.alignment = font_dept, Alignment(
        horizontal="center", vertical="center"
    )

    ws.merge_cells(
        start_row=2, start_column=1, end_row=2, end_column=max(total_cols, 2)
    )
    c2 = ws.cell(
        row=2, column=1, value=f"{report_title} ( दिनांक {current_date_str} )"
    )
    c2.font, c2.alignment = font_sub, Alignment(
        horizontal="center", vertical="center"
    )

    # Table Headers
    start_row = 4
    if is_multi:
      for i, idx_name in enumerate(df_data.index.names):
        cell = ws.cell(row=start_row, column=i + 1, value=idx_name or "")
        cell.font, cell.fill = font_th, fill_th
        cell.alignment = Alignment(horizontal="center", vertical="center")
    else:
      cell = ws.cell(
          row=start_row, column=1, value=df_data.index.name or "तालुका"
      )
      cell.font, cell.fill = font_th, fill_th
      cell.alignment = Alignment(horizontal="center", vertical="center")

    for j, col_name in enumerate(df_data.columns):
      cell = ws.cell(
          row=start_row, column=num_idx_cols + j + 1, value=str(col_name)
      )
      cell.font, cell.fill = font_th, fill_th
      cell.alignment = Alignment(horizontal="center", vertical="center")

    # Table Data
    current_r = start_row + 1
    for idx_val, row in df_data.iterrows():
      if is_multi:
        for k, val in enumerate(idx_val):
          cell = ws.cell(row=current_r, column=k + 1, value=str(val))
          cell.alignment = Alignment(horizontal="left", vertical="center")
          cell.border = thin_border
      else:
        cell = ws.cell(row=current_r, column=1, value=str(idx_val))
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = thin_border

      for j, col_name in enumerate(df_data.columns):
        val = row[col_name]
        val_clean = (
            ""
            if (
                pd.isna(val)
                or val == 0
                or str(val).strip() in ["", "0", "nan", "0.0"]
            )
            else (int(val) if isinstance(val, (int, float)) else val)
        )
        cell = ws.cell(
            row=current_r, column=num_idx_cols + j + 1, value=val_clean
        )
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

      current_r += 1

    # Auto Column Width
    for col in ws.columns:
      max_len = max(len(str(cell.value or "")) for cell in col)
      col_letter = openpyxl.utils.get_column_letter(col[0].column)
      ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

  wb.save(output)
  return output.getvalue()


# PROFESSIONAL & CLEAN PDF GENERATOR FOR ALL REPORTS
def generate_report_pdf_html(df, subtitle_text=""):
  headers = list(df.columns)
  current_date_str = datetime.date.today().strftime("%d/%m/%Y")
  is_multi_index = isinstance(df.index, pd.MultiIndex)

  html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{ size: A4 portrait; margin: 8mm; }}
            body {{ font-family: 'Helvetica Neue', 'Arial', sans-serif; margin: 0; padding: 0; background-color: #ffffff; color: #111; }}
            .header-container {{ text-align: center; margin-bottom: 8px; padding-bottom: 4px; border-bottom: 2px solid #1a365d; }}
            .dept-title {{ font-size: 15px; font-weight: bold; color: #1a365d; text-transform: uppercase; margin-bottom: 2px; }}
            .report-title {{ font-size: 12px; font-weight: 600; color: #2b6cb0; }}
            table {{ width: 100%; border-collapse: collapse; border: 1px solid #4a5568; margin-top: 4px; font-size: 9.5px; table-layout: fixed; }}
            thead {{ display: table-header-group; }}
            tr {{ page-break-inside: avoid; }}
            th {{ background-color: #1a365d; color: #ffffff; font-weight: bold; font-size: 9.5px; padding: 4px 2px; border: 1px solid #2d3748; text-align: center; line-height: 1.1; }}
            td {{ padding: 3px 2px; font-size: 9px; border: 1px solid #cbd5e0; text-align: center; color: #2d3748; line-height: 1.1; }}
            tbody tr:nth-child(even) {{ background-color: #f7fafc; }}
            .text-col {{ text-align: left !important; padding-left: 4px !important; font-weight: 500; }}
            .total-row td {{ font-weight: bold !important; background-color: #edf2f7 !important; color: #1a365d !important; border-top: 1.5px solid #1a365d !important; }}
            .grand-total-col {{ background-color: #feebc8 !important; font-weight: bold !important; color: #742a2a !important; }}
        </style>
    </head>
    <body>
        <div class="header-container">
            <div class="dept-title">भूमि अभिलेख विभाग - अमरावती</div>
            <div class="report-title">{subtitle_text} ( दिनांक {current_date_str} )</div>
        </div>
        <table>
            <thead>
                <tr>
    """

  if is_multi_index:
    for idx_name in df.index.names:
      width = "18%" if idx_name == "तालुका" else "26%"
      html += f'<th style="width: {width};">{idx_name}</th>'
  else:
    html += '<th style="width: 22%;">तालुका</th>'

  for h in headers:
    html += f"<th>{h}</th>"

  html += """
                </tr>
            </thead>
            <tbody>
    """

  for idx, row in df.iterrows():
    idx_str = str(idx)
    is_total = (
        "Total" in idx_str or "एकूण" in idx_str or "Grand Total" in idx_str
    )
    row_class = "total-row" if is_total else ""

    html += f'<tr class="{row_class}">'

    if is_multi_index:
      html += f'<td class="text-col">{idx[0]}</td>'
      html += f'<td class="text-col">{idx[1]}</td>'
    else:
      html += f'<td class="text-col">{idx}</td>'

    for col_name in headers:
      val = row[col_name]
      val_str = (
          ""
          if (
              pd.isna(val)
              or val == 0
              or str(val).strip() in ["", "0", "nan", "0.0"]
          )
          else str(int(val) if isinstance(val, (int, float)) else val)
      )

      col_class = (
          "grand-total-col"
          if col_name in ["Grand Total", "एकूण", "Officer Total", "Stage Total"]
          else ""
      )
      html += f'<td class="{col_class}">{val_str}</td>'

    html += "</tr>"

  html += """
            </tbody>
        </table>
    </body>
    </html>
    """

  if WEASYPRINT_AVAILABLE:
    try:
      return HTML(string=html).write_pdf()
    except Exception:
      return html.encode("utf-8")
  else:
    return html.encode("utf-8")


# Sidebar File Upload
st.sidebar.header("⚙️ File Upload & Options")
uploaded_file = st.sidebar.file_uploader(
    "📁 Main E-Mojani Raw Excel (.xlsx)", type=["xlsx"], key="single_file"
)


def parse_excel_date(val):
  if pd.isna(val) or str(val).strip() in ["", "nan", "NaT"]:
    return pd.NaT
  try:
    val_num = float(val)
    return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=val_num)
  except Exception:
    return pd.to_datetime(val, dayfirst=True, errors="coerce")


if uploaded_file is not None:
  df_raw = pd.read_excel(uploaded_file)
  if "मोजणी तारीख" not in df_raw.columns:
    df_raw = pd.read_excel(uploaded_file, header=1)

  if "तालुका" in df_raw.columns:
    df_raw["तालुका"] = df_raw["तालुका"].astype(str).str.strip()
    df_raw = df_raw[
        df_raw["तालुका"].notna()
        & (~df_raw["तालुका"].isin(["", "nan", "None", "0"]))
    ]

  col_mojni_date = (
      "मोजणी तारीख"
      if "मोजणी तारीख" in df_raw.columns
      else df_raw.columns[8]
      if len(df_raw.columns) >= 9
      else None
  )
  col_status = (
      "स्थिती"
      if "स्थिती" in df_raw.columns
      else df_raw.columns[10]
      if len(df_raw.columns) >= 11
      else None
  )

  df_raw["mojni_date_parsed"] = df_raw[col_mojni_date].apply(parse_excel_date)

  col_k_prat_date = "क प्रत निकाली दिनाक"
  if col_k_prat_date in df_raw.columns:
    df_raw["k_prat_date_parsed"] = df_raw[col_k_prat_date].apply(
        parse_excel_date
    )
  else:
    df_raw["k_prat_date_parsed"] = df_raw["mojni_date_parsed"]

  min_excel_date = df_raw["mojni_date_parsed"].min()
  max_excel_date = df_raw["mojni_date_parsed"].max()

  min_excel_date = (
      datetime.date(2023, 1, 1)
      if pd.isna(min_excel_date)
      else min_excel_date.date()
  )
  max_excel_date = (
      datetime.date.today()
      if pd.isna(max_excel_date)
      else max_excel_date.date()
  )
  today_file_str = datetime.date.today().strftime("%d-%m-%Y")

  st.sidebar.markdown("---")
  st.sidebar.subheader("📅 Report 1 कालावधी (Date Range)")
  col_d1, col_d2 = st.sidebar.columns(2)
  start_date = col_d1.date_input(
      "पासून (From)", value=min_excel_date, key="r1_start"
  )
  end_date = col_d2.date_input(
      "पर्यंत (To)", value=datetime.date.today(), key="r1_end"
  )

  st.sidebar.markdown("---")
  st.sidebar.subheader("📌 स्थिती (Status) Filter")
  all_statuses = df_raw[col_status].dropna().unique().tolist()
  completed_defaults = [
      "क प्रत",
      "विनाकार्यवाही",
      "प्रस्तावित बिगरशेती/गुंठेवारी मोजणी पूर्ण",
  ]
  default_selected = [s for s in all_statuses if s not in completed_defaults]

  selected_statuses = st.sidebar.multiselect(
      "Select Status:", options=all_statuses, default=default_selected
  )

  with st.spinner("डेटा प्रोसेस होत आहे... कृपया वाट पहा..."):
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Report 1 (तालुका व दिवसनिहाय)",
        "📋 Report 2 (टप्पा व अधिकारी निहाय)",
        "📈 Executive Dashboard",
        "📋 Report 3 (भूकरमापक निहाय)",
        "📅 Report 4 (मोजणी होणेवर शिल्लक)",
        "🚀 Report 5 (मागील ३ महिने क प्रत)",
    ])

    from_str = start_date.strftime("%d/%m/%Y")
    to_str = end_date.strftime("%d/%m/%Y")

    # --- REPORT 1 GENERATION ---
    df1 = df_raw[df_raw[col_status].isin(selected_statuses)].copy()
    start_dt, end_dt = pd.to_datetime(start_date), pd.to_datetime(end_date)
    df1 = df1[
        (df1["mojni_date_parsed"] >= start_dt)
        & (df1["mojni_date_parsed"] <= end_dt)
    ].copy()
    df1["Pending Days"] = (end_dt - df1["mojni_date_parsed"]).dt.days

    def day_bucket(days):
      if pd.isna(days):
        return "तारीख उपलब्ध नाही"
      if days <= 15:
        return "15 दिवस वर"
      elif days <= 30:
        return "30 दिवस वर"
      elif days <= 60:
        return "60 दिवस वर"
      elif days <= 90:
        return "90 दिवस वर"
      else:
        return "90 दिवसांपेक्षा जास्त"

    df1["दिवस वर"] = df1["Pending Days"].apply(day_bucket)
    bucket_order = [
        "15 दिवस वर",
        "30 दिवस वर",
        "60 दिवस वर",
        "90 दिवस वर",
        "90 दिवसांपेक्षा जास्त",
    ]
    pivot_df = pd.crosstab(
        df1["तालुका"], df1["दिवस वर"], margins=True, margins_name="एकूण"
    )
    existing_cols = [c for c in bucket_order if c in pivot_df.columns]
    if "तारीख उपलब्ध नाही" in pivot_df.columns:
      existing_cols.append("तारीख उपलब्ध नाही")
    if "एकूण" in pivot_df.columns:
      existing_cols.append("एकूण")
    pivot_df = pivot_df[existing_cols]

    # --- REPORT 2 GENERATION ---
    r2_start_dt, r2_end_dt = pd.to_datetime(min_excel_date), pd.to_datetime(
        datetime.date.today()
    )
    today_dt = pd.to_datetime(datetime.date.today())
    df_r2_filtered = df_raw[
        (df_raw["mojni_date_parsed"] >= r2_start_dt)
        & (df_raw["mojni_date_parsed"] <= r2_end_dt)
    ].copy()
    col_yn, col_mojni_type = (
        "क्षेत्र अभिलेखाशी मेळात आहे का?",
        "मोजणीचा प्रकार(Mojni Type)",
    )
    taluka_list = sorted(df_raw["तालुका"].dropna().unique().tolist())
    report2_data = []

    for tal in taluka_list:
      df_t = df_r2_filtered[df_r2_filtered["तालुका"] == tal]
      df_taluka_all = df_raw[df_raw["तालुका"] == tal]
      chanani = len(df_t[df_t[col_status] == "छाननी लिपिक यांनी तपासले"])
      shirastedar = len(
          df_t[
              df_t[col_status]
              == "शिरस्तेदार/मुख्यालय सहाय्यक यांनी तपासले"
          ]
      )
      up_bhoo = len(
          df_t[
              df_t[col_status] == "ऊप.अ. भू. अ/ न .भू अ यांच्या मान्यतेवर"
          ]
      )
      off_total = chanani
