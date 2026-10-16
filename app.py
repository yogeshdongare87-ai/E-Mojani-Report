import datetime
import io
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


# Excel Download Helper
def convert_df_to_excel(df, sheet_name="Report"):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name=sheet_name)
  return output.getvalue()


# 🌟 PROFESSIONAL & CLEAN PDF GENERATOR FOR ALL REPORTS
def generate_report_pdf_html(df, subtitle_text=""):
  headers = list(df.columns)
  current_date_str = datetime.date.today().strftime("%d/%m/%Y")

  # Detect MultiIndex (e.g. Taluka + Bhukarmapak in Report 3)
  is_multi_index = isinstance(df.index, pd.MultiIndex)

  html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{
                size: A4 portrait;
                margin: 8mm;
            }}
            body {{
                font-family: 'Helvetica Neue', 'Arial', 'DejaVu Sans', sans-serif;
                margin: 0;
                padding: 0;
                background-color: #ffffff;
                color: #111111;
            }}
            .header-container {{
                text-align: center;
                margin-bottom: 8px;
                padding-bottom: 4px;
                border-bottom: 2px solid #1a365d;
            }}
            .dept-title {{
                font-size: 15px;
                font-weight: bold;
                color: #1a365d;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                margin-bottom: 2px;
            }}
            .report-title {{
                font-size: 12px;
                font-weight: 600;
                color: #2b6cb0;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                border: 1px solid #4a5568;
                margin-top: 4px;
                font-size: 9.5px;
                table-layout: fixed;
            }}
            thead {{
                display: table-header-group;
            }}
            tr {{
                page-break-inside: avoid;
            }}
            th {{
                background-color: #1a365d;
                color: #ffffff;
                font-weight: bold;
                font-size: 9.5px;
                padding: 4px 2px;
                border: 1px solid #2d3748;
                text-align: center;
                word-wrap: break-word;
                line-height: 1.1;
            }}
            td {{
                padding: 3px 2px;
                font-size: 9px;
                border: 1px solid #cbd5e0;
                text-align: center;
                color: #2d3748;
                word-wrap: break-word;
                line-height: 1.1;
            }}
            tbody tr:nth-child(even) {{
                background-color: #f7fafc;
            }}
            .text-col {{
                text-align: left !important;
                padding-left: 4px !important;
                font-weight: 500;
            }}
            .total-row td {{
                font-weight: bold !important;
                background-color: #edf2f7 !important;
                color: #1a365d !important;
                border-top: 1.5px solid #1a365d !important;
                border-bottom: 1.5px solid #1a365d !important;
            }}
            .grand-total-col {{
                background-color: #feebc8 !important;
                font-weight: bold !important;
                color: #742a2a !important;
            }}
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


# Sidebar - File Upload & Filters
st.sidebar.header("⚙️ File Upload & Options")

uploaded_file = st.sidebar.file_uploader(
    "📁 Main E-Mojani Raw Excel (.xlsx) [All Reports 1 to 5]",
    type=["xlsx"],
    key="single_file",
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

  # Text Cleaning
  if "तालुका" in df_raw.columns:
    df_raw["तालुका"] = df_raw["तालुका"].astype(str).str.strip()
    df_raw = df_raw[
        df_raw["तालुका"].notna()
        & (~df_raw["तालुका"].isin(["", "nan", "None", "0"]))
    ]

  # Dynamic Column Detection
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

  # Sidebar Filters
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
      "Select Status:",
      options=all_statuses,
      default=default_selected,
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

    # ---------------------------------------------------------------------
    # TAB 1: REPORT 1
    # ---------------------------------------------------------------------
    with tab1:
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

      st.success(f"✅ Total Filtered Cases: **{len(df1)}**")
      st.info(f"📅 **Selected Pending Period:** {from_str} ते {to_str}")
      st.subheader("📋 तालुका व दिवसनिहाय प्रलंबित प्रकरणे Table")
      st.dataframe(pivot_df, use_container_width=True)

      btn_col1, btn_col2 = st.columns(2)
      with btn_col1:
        excel_r1 = convert_df_to_excel(pivot_df, sheet_name="Report_1")
        st.download_button(
            label="📥 Download Report 1 (Excel)",
            data=excel_r1,
            file_name=f"Report_1_Talukani_hay_{today_file_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      with btn_col2:
        pdf_r1 = generate_report_pdf_html(
            pivot_df,
            subtitle_text="तालुका व दिवसनिहाय प्रलंबित मोजणी प्रकरणे अहवाल",
        )
        st.download_button(
            label="📄 Download Report 1 (PDF / Print)",
            data=pdf_r1,
            file_name=(
                f"Report_1_Talukani_hay_{today_file_str}.pdf"
                if WEASYPRINT_AVAILABLE
                else f"Report_1_Talukani_hay_{today_file_str}.html"
            ),
            mime="application/pdf" if WEASYPRINT_AVAILABLE else "text/html",
        )

    # ---------------------------------------------------------------------
    # TAB 2: REPORT 2
    # ---------------------------------------------------------------------
    with tab2:
      st.subheader("📋 Report 2 - टप्पा व अधिकारी निहाय अहवाल")
      col_r2_d1, col_r2_d2 = st.columns(2)
      r2_start_date = col_r2_d1.date_input(
          "पासून (From)", value=min_excel_date, key="r2_start"
      )
      r2_end_date = col_r2_d2.date_input(
          "पर्यंत (To)", value=datetime.date.today(), key="r2_end"
      )

      r2_from_str = r2_start_date.strftime("%d/%m/%Y")
      r2_to_str = r2_end_date.strftime("%d/%m/%Y")
      st.info(f"🗓 **Report 2 Period:** {r2_from_str} ते {r2_to_str}")

      r2_start_dt, r2_end_dt = pd.to_datetime(r2_start_date), pd.to_datetime(
          r2_end_date
      )
      today_dt = pd.to_datetime(datetime.date.today())

      df_r2_filtered = df_raw[
          (df_raw["mojni_date_parsed"] >= r2_start_dt)
          & (df_raw["mojni_date_parsed"] <= r2_end_dt)
      ].copy()

      col_yn = "क्षेत्र अभिलेखाशी मेळात आहे का?"
      col_mojni_type = "मोजणीचा प्रकार(Mojni Type)"

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
                df_t[col_status]
                == "ऊप.अ. भू. अ/ न .भू अ यांच्या मान्यतेवर"
            ]
        )
        off_total = chanani + shirastedar + up_bhoo

        yes_no = (
            len(
                df_t[
                    df_t[col_yn].notna()
                    & (df_t[col_yn].astype(str).str.strip() != "")
                ]
            )
            if col_yn in df_t.columns
            else 0
        )

        if col_mojni_type in df_t.columns:
          haddi = len(
              df_t[
                  (df_t[col_mojni_type] == "ह्द्दकायम")
                  & (df_t[col_status] == "मोजणीची माहिती
