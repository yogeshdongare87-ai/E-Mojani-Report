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

# Custom Styling (CSS)
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

st.title("📊 भूमि अभिलेख विभाग - प्रलंबित मोजणी अहवाल व डैशबोर्ड")

# Sidebar - File Uploads & Filters
st.sidebar.header("⚙️ Options & File Uploads")

uploaded_file_raw = st.sidebar.file_uploader(
    "1️⃣ Main E-Mojani Raw Excel (.xlsx) [Reports 1, 2, 3 साठी]",
    type=["xlsx"],
    key="file_raw",
)

uploaded_file_citizen = st.sidebar.file_uploader(
    "2️⃣ Citizen Application Excel (.xlsx) [Report 4 साठी]",
    type=["xlsx"],
    key="file_citizen",
)


def parse_excel_date(val):
  if pd.isna(val) or str(val).strip() in ["", "nan", "NaT"]:
    return pd.NaT
  try:
    val_num = float(val)
    return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=val_num)
  except Exception:
    return pd.to_datetime(val, dayfirst=True, errors="coerce")


if uploaded_file_raw is not None:
  df_raw = pd.read_excel(uploaded_file_raw)
  if "मोजणी तारीख" not in df_raw.columns:
    df_raw = pd.read_excel(uploaded_file_raw, header=1)

  # Text Cleaning
  if "तालुका" in df_raw.columns:
    df_raw["तालुका"] = df_raw["तालुका"].astype(str).str.strip()
    df_raw = df_raw[
        df_raw["तालुका"].notna()
        & (~df_raw["तालुका"].isin(["", "nan", "None", "0"]))
    ]

  if "स्थिती" in df_raw.columns:
    df_raw["स्थिती"] = df_raw["स्थिती"].astype(str).str.strip()

  df_raw["mojni_date_parsed"] = df_raw["मोजणी तारीख"].apply(parse_excel_date)

  min_excel_date = df_raw["mojni_date_parsed"].min()
  max_excel_date = df_raw["mojni_date_parsed"].max()

  if pd.isna(min_excel_date):
    min_excel_date = datetime.date(2023, 1, 1)
  else:
    min_excel_date = min_excel_date.date()

  if pd.isna(max_excel_date):
    max_excel_date = datetime.date.today()
  else:
    max_excel_date = max_excel_date.date()

  # --- SIDEBAR FILTERS ---
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

  all_statuses = df_raw["स्थिती"].dropna().unique().tolist()
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
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Report 1 (तालुका व दिवसनिहाय)",
        "📋 Report 2 (टप्पा व अधिकारी निहाय)",
        "📈 Executive Dashboard",
        "📋 Report 3 (भूकरमापक निहाय)",
        "📅 Report 4 (मोजणी होणेवर शिल्लक)",
    ])

    from_str = start_date.strftime("%d/%m/%Y")
    to_str = end_date.strftime("%d/%m/%Y")

    # ---------------------------------------------------------------------
    # TAB 1: REPORT 1 (तालुका व दिवसनिहाय)
    # ---------------------------------------------------------------------
    with tab1:
      df1 = df_raw[df_raw["स्थिती"].isin(selected_statuses)].copy()

      start_dt = pd.to_datetime(start_date)
      end_dt = pd.to_datetime(end_date)

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

      if WEASYPRINT_AVAILABLE:
        rows_html = ""
        for taluka, row in pivot_df.iterrows():
          is_total = taluka == "एकूण"
          tr_style = (
              "background-color: #d9f0a3; font-weight: bold;" if is_total else ""
          )
          tds = f"<td><b>{taluka}</b></td>"
          for col in existing_cols:
            tds += f"<td style='text-align: center;'>{row[col]}</td>"
          rows_html += f"<tr style='{tr_style}'>{tds}</tr>"

        headers_html = "<th>तालुका</th>" + "".join(
            [f"<th>{c}</th>" for c in existing_cols]
        )

        html_content = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <style>
                        @page {{ size: A4 portrait; margin: 15mm; }}
                        body {{ font-family: 'Gargi', 'DejaVu Sans', sans-serif; font-size: 11px; }}
                        h2 {{ text-align: center; margin-bottom: 5px; color: #1b4332; }}
                        .date-header {{ text-align: center; font-size: 13px; font-weight: bold; margin-bottom: 15px; }}
                        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
                        th {{ background-color: #2b9348; color: white; padding: 6px; border: 1px solid #555; font-size: 10px; }}
                        td {{ padding: 5px; border: 1px solid #888; font-size: 10px; }}
                        tr:nth-child(even) {{ background-color: #f9f9f9; }}
                    </style>
                </head>
                <body>
                    <h2>भूमि अभिलेख विभाग - अमरावती</h2>
                    <div class="date-header">तालुका व दिवसनिहाय प्रलंबित मोजणी प्रकरणे अहवाल (दिनांक {from_str} ते {to_str})</div>
                    <table>
                        <thead><tr>{headers_html}</tr></thead>
                        <tbody>{rows_html}</tbody>
                    </table>
                </body>
                </html>
                """
        try:
          pdf_bytes = HTML(string=html_content).write_pdf()
          st.download_button(
              label="📥 Download PDF Report 1",
              data=pdf_bytes,
              file_name=(
                  "Mojani_Pending_Report1_"
                  f"{from_str.replace('/', '-')}_to_{to_str.replace('/', '-')}.pdf"
              ),
              mime="application/pdf",
          )
        except Exception as e:
          st.error(f"PDF जनरेट करताना त्रुटी आली: {e}")

    # ---------------------------------------------------------------------
    # TAB 2: REPORT 2 (टप्पा व अधिकारी निहाय)
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

      r2_start_dt = pd.to_datetime(r2_start_date)
      r2_end_dt = pd.to_datetime(r2_end_date)
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

        chanani = len(df_t[df_t["स्थिती"] == "छाननी लिपिक यांनी तपासले"])
        shirastedar = len(
            df_t[
                df_t["स्थिती"]
                == "शिरस्तेदार/मुख्यालय सहाय्यक यांनी तपासले"
            ]
        )
        up_bhoo = len(
            df_t[
                df_t["स्थिती"]
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
                  & (df_t["स्थिती"] == "मोजणीची माहिती")
              ]
          )
          jama = len(
              df_t[
                  (df_t["स्थिती"] == "मोजणीची माहिती")
                  & (df_t[col_mojni_type] != "ह्द्दकायम")
              ]
          )
        else:
          haddi = len(df_t[df_t["स्थिती"] == "मोजणीची माहिती"])
          jama = 0

        shillak = len(
            df_taluka_all[
                (df_taluka_all["mojni_date_parsed"] > today_dt)
                & (df_taluka_all["स्थिती"] == "मोजणीची माहिती")
            ]
        )

        stage_total = yes_no + jama + haddi + shillak

        report2_data.append({
            "तालुका": tal,
            "Yes/No": yes_no,
            "जमा करणेवर": jama,
            "हददी दाखविणेवर": haddi,
            "शिल्लक प्रकरणे": shillak,
            "Grand Total": stage_total,
            "छाननी लिपीक": chanani,
            "शिरस्तेदार/मुख्यालय सहाय्यक": shirastedar,
            "उप अ भू अ/ भू अ": up_bhoo,
            "Grand Total ": off_total,
        })

      df_rep2 = pd.DataFrame(report2_data)

      total_row = {
          "तालुका": "एकूण",
          "Yes/No": df_rep2["Yes/No"].sum(),
          "जमा करणेवर": df_rep2["जमा करणेवर"].sum(),
          "हददी दाखविणेवर": df_rep2["हददी दाखविणेवर"].sum(),
          "शिल्लक प्रकरणे": df_rep2["शिल्लक प्रकरणे"].sum(),
          "Grand Total": df_rep2["Grand Total"].sum(),
          "छाननी लिपीक": df_rep2["छाननी लिपीक"].sum(),
          "शिरस्तेदार/मुख्यालय सहाय्यक": df_rep2[
              "शिरस्तेदार/मुख्यालय सहाय्यक"
          ].sum(),
          "उप अ भू अ/ भू अ": df_rep2["उप अ भू अ/ भू अ"].sum(),
          "Grand Total ": df_rep2["Grand Total "].sum(),
      }

      df_rep2 = pd.concat(
          [df_rep2, pd.DataFrame([total_row])], ignore_index=True
      )
      st.dataframe(df_rep2, use_container_width=True)

    # ---------------------------------------------------------------------
    # TAB 3: EXECUTIVE DASHBOARD
    # ---------------------------------------------------------------------
    with tab3:
      st.markdown("## 🎨 Executive Visual Dashboard")

      today = datetime.date.today()
      first_day_of_curr_month = today.replace(day=1)
      last_day_prev_month = first_day_of_curr_month - datetime.timedelta(
          days=1
      )
      first_day_prev_month = last_day_prev_month.replace(day=1)

      st.info(
          f"📅 **विश्लेषण कालावधी (मागील महिना):**"
          f" {first_day_prev_month.strftime('%d/%m/%Y')} ते"
          f" {last_day_prev_month.strftime('%d/%m/%Y')}"
      )

      p_start_dt = pd.to_datetime(first_day_prev_month)
      p_end_dt = pd.to_datetime(last_day_prev_month)

      df_prev_kprat = df_raw[
          (df_raw["mojni_date_parsed"] >= p_start_dt)
          & (df_raw["mojni_date_parsed"] <= p_end_dt)
          & (df_raw["स्थिती"] == "क प्रत")
      ].copy()

      # KPI Cards
      c1, c2, c3 = st.columns(3)
      with c1:
        st.markdown(
            f"""
                    <div class="metric-card" style="border-left-color: #2a9d8f;">
                        <div class="metric-title">🏆 मागील महिन्यात 'क प्रत' निकाल</div>
                        <div class="metric-value">{len(df_prev_kprat)}</div>
                    </div>
                """,
            unsafe_allow_html=True,
        )
      with c2:
        st.markdown(
            f"""
                    <div class="metric-card" style="border-left-color: #e76f51;">
                        <div class="metric-title">⏳ एकूण प्रलंबित प्रकरणे</div>
                        <div class="metric-value">{len(df_raw[~df_raw["स्थिती"].isin(completed_defaults)])}</div>
                    </div>
                """,
            unsafe_allow_html=True,
        )
      with c3:
        st.markdown(
            f"""
                    <div class="metric-card" style="border-left-color: #f4a261;">
                        <div class="metric-title">📌 आजच्या तारखेनंतरची शिल्लक प्रकरणे</div>
                        <div class="metric-value">{len(df_raw[(df_raw["mojni_date_parsed"] > today_dt) & (df_raw["स्थिती"] == "मोजणीची माहिती")])}</div>
                    </div>
                """,
            unsafe_allow_html=True,
        )

      st.markdown("<br>", unsafe_allow_html=True)

      col_chart1, col_chart2 = st.columns(2)

      with col_chart1:
        st.markdown("### 🥇 Top 3 तालुका ('क प्रत')")
        if not df_prev_kprat.empty:
          top_t = (
              df_prev_kprat["तालुका"].value_counts().head(3).reset_index()
          )
          top_t.columns = ["तालुका", "संख्या"]

          fig_taluka = px.bar(
              top_t,
              x="तालुका",
              y="संख्या",
              text="संख्या",
              color="तालुका",
              color_discrete_sequence=px.colors.qualitative.Set2,
              title="Top 3 तालुका - 'क प्रत' कामगिरी",
          )
          fig_taluka.update_traces(textposition="outside")
          fig_taluka.update_layout(showlegend=False, height=350)
          st.plotly_chart(fig_taluka, use_container_width=True)
        else:
          st.warning("मागील महिन्यात 'क प्रत' चा डेटा उपलब्ध नाही.")

      with col_chart2:
        st.markdown("### 👷 Top भूकरमापक ('क प्रत')")

        surveyor_col = None
        for col_name in ["भूकरमापक", "कर्मचारी/अधिकारी चे नाव"]:
          if col_name in df_raw.columns:
            surveyor_col = col_name
            break

        if surveyor_col and not df_prev_kprat.empty:
          df_surveyor_filtered = df_prev_kprat[
              df_prev_kprat[surveyor_col].notna()
              & (df_prev_kprat[surveyor_col].astype(str).str.strip() != "")
          ]

          top_s = (
              df_surveyor_filtered[surveyor_col]
              .value_counts()
              .head(5)
              .reset_index()
          )
          top_s.columns = ["भूकरमापक नाव", "संख्या"]

          if not top_s.empty:
            fig_surv = px.bar(
                top_s,
                x="संख्या",
                y="भूकरमापक नाव",
                orientation="h",
                text="संख्या",
                color="भूकरमापक नाव",
                color_discrete_sequence=px.colors.qualitative.Pastel,
                title="सर्वोत्तम कामगिरी करणारे भूकरमापक",
            )
            fig_surv.update_traces(textposition="outside")
            fig_surv.update_layout(
                showlegend=False,
                height=350,
                yaxis=dict(autorange="reversed"),
            )
            st.plotly_chart(fig_surv, use_container_width=True)

    # ---------------------------------------------------------------------
    # TAB 4: REPORT 3 (भूकरमापक निहाय व दिवसनिहाय प्रलंबित प्रकरणे)
    # ---------------------------------------------------------------------
    with tab4:
      st.markdown(
          "## 📋 रिपोर्ट ३: भूकरमापक निहाय व दिवसनिहाय प्रलंबित मोजणी प्रकरणांचा"
          " आढावा"
      )

      surveyor_col_r3 = None
      for col_name in ["भूकरमापक", "कर्मचारी/अधिकारी चे नाव"]:
        if col_name in df_raw.columns:
          surveyor_col_r3 = col_name
          break

      if not surveyor_col_r3:
        st.error(
            "❌ Upload केलेल्या Excel फाईलमध्ये 'भूकरमापक' चा कॉलम आढळला नाही."
        )
      else:
        df_r3_clean = df_raw[
            df_raw["तालुका"].notna() & df_raw[surveyor_col_r3].notna()
        ].copy()

        df_r3_pending = df_r3_clean[
            df_r3_clean["स्थिती"].isin(selected_statuses)
        ].copy()

        curr_today = pd.to_datetime(datetime.date.today())
        df_r3_pending = df_r3_pending[
            df_r3_pending["mojni_date_parsed"].notna()
            & (df_r3_pending["mojni_date_parsed"] <= curr_today)
        ].copy()

        if df_r3_pending.empty:
          st.warning(
              "निवडलेल्याフィルターनुसार कोणताही प्रलंबित डेटा उपलब्ध नाही."
          )
        else:
          df_r3_pending["R3_Pending_Days"] = (
              curr_today - df_r3_pending["mojni_date_parsed"]
          ).dt.days

          def r3_ageing_bucket(days):
            if pd.isna(days):
              return "(blank)"
            elif days <= 15:
              return "0–15 दिवस वर"
            elif days <= 30:
              return "16–30 दिवस वर"
            elif days <= 60:
              return "31–60 दिवस वर"
            elif days <= 90:
              return "61–90 दिवस वर"
            else:
              return "90 दिवसांपेक्षा जास्त"

          df_r3_pending["Ageing_Bucket"] = df_r3_pending[
              "R3_Pending_Days"
          ].apply(r3_ageing_bucket)

          pivot_r3 = pd.pivot_table(
              df_r3_pending,
              index=["तालुका", surveyor_col_r3],
              columns="Ageing_Bucket",
              values=(
                  "अर्ज क्र.(Application No)"
                  if "अर्ज क्र.(Application No)" in df_r3_pending.columns
                  else df_r3_pending.columns[0]
              ),
              aggfunc="count",
              fill_value=0,
              margins=True,
              margins_name="Grand Total",
          )

          r3_bucket_order = [
              "0–15 दिवस वर",
              "16–30 दिवस वर",
              "31–60 दिवस वर",
              "61–90 दिवस वर",
              "90 दिवसांपेक्षा जास्त",
              "(blank)",
              "Grand Total",
          ]

          existing_r3_cols = [
              c for c in r3_bucket_order if c in pivot_r3.columns
          ]
          pivot_r3 = pivot_r3[existing_r3_cols]

          curr_date_fmt = datetime.date.today().strftime("%d/%m/%Y")
          st.markdown(
              f"### **मोजणी प्रकरणाबाबत आढावा दिनांक: {curr_date_fmt}**"
          )
          st.dataframe(pivot_r3, use_container_width=True, height=500)

    # ---------------------------------------------------------------------
    # TAB 5: REPORT 4 (मोजणी होणेवर शिल्लक - Citizen Application File)
    # ---------------------------------------------------------------------
    with tab5:
      st.markdown("## 📅 रिपोर्ट ४: मोजणी होणेवर शिल्लक अहवाल")

      df_source_r4 = None
      if uploaded_file_citizen is not None:
        st.success(
            "✅ 'Citizen Application' Excel फाईल वापरून रिपोर्ट ४ जनरेट होत आहे."
        )
        df_cit_temp = pd.read_excel(uploaded_file_citizen)
        if "मोजणी तारीख" not in df_cit_temp.columns:
          df_cit_temp = pd.read_excel(uploaded_file_citizen, header=1)

        if "तालुका" in df_cit_temp.columns:
          df_cit_temp["तालुका"] = (
              df_cit_temp["तालुका"].astype(str).str.strip()
          )
          df_cit_temp = df_cit_temp[
              df_cit_temp["तालुका"].notna()
              & (~df_cit_temp["तालुका"].isin(["", "nan", "None", "0"]))
          ]
        if "स्थिती" in df_cit_temp.columns:
          df_cit_temp["स्थिती"] = df_cit_temp["स्थिती"].astype(str).str.strip()

        df_cit_temp["mojni_date_parsed"] = df_cit_temp["मोजणी तारीख"].apply(
            parse_excel_date
        )
        df_source_r4 = df_cit_temp
      else:
        st.info(
            "💡 टीप: डाव्या Sidebar मधील २ऱ्या Uploader मध्ये 'Citizen"
            " Application' Excel Upload करू शकता. सध्या १ल्या Excel वरून डेटा"
            " दाखवत आहे."
        )
        df_source_r4 = df_raw.copy()

      today_dt_r4 = pd.to_datetime(datetime.date.today())

      # Filter: mojni_date > today & status == "मोजणीची माहिती"
      df_r4 = df_source_r4[
          (df_source_r4["mojni_date_parsed"].notna())
          & (df_source_r4["mojni_date_parsed"] > today_dt_r4)
          & (df_source_r4["स्थिती"] == "मोजणीची माहिती")
      ].copy()

      if df_r4.empty:
        st.warning(
            "आजच्या तारखेच्या पुढे 'मोजणीची माहिती' स्थितीत असलेली कोणतीही"
            " प्रकरणे आढळली नाहीत."
        )
      else:
        df_r4["Month_Year"] = df_r4["mojni_date_parsed"].dt.strftime("%b")
        df_r4["Year_Month_Sort"] = df_r4["mojni_date_parsed"].dt.strftime(
            "%Y-%m"
        )

        sorted_months_df = (
            df_r4[["Year_Month_Sort", "Month_Year"]]
            .drop_duplicates()
            .sort_values("Year_Month_Sort")
        )
        sorted_month_names = sorted_months_df["Month_Year"].tolist()

        pivot_r4 = pd.pivot_table(
            df_r4,
            index="तालुका",
            columns="Month_Year",
            values=(
                "अर्ज क्र.(Application No)"
                if "अर्ज क्र.(Application No)" in df_r4.columns
                else df_r4.columns[0]
            ),
            aggfunc="count",
            fill_value=0,
        )

        cols_in_pivot = [m for m in sorted_month_names if m in pivot_r4.columns]
        pivot_r4 = pivot_r4[cols_in_pivot]

        standard_talukas = [
            "अंजनगाव",
            "अचलपूर",
            "अमरावती",
            "चांदुर बा.",
            "चिखलदरा",
            "तिवसा",
            "दर्यापूर",
            "धामणगाव",
            "धारणी",
            "नांदगाव",
            "भातकुली",
            "मोर्शी",
            "वरुड",
        ]

        pivot_r4 = pivot_r4.reindex(standard_talukas, fill_value=0)

        # Calculate Row Grand Total
        pivot_r4["Grand Total"] = pivot_r4.sum(axis=1)

        # Calculate District Row (अमरावती)
        district_total = pivot_r4.sum(axis=0)
        df_district_row = pd.DataFrame([district_total], index=["अमरावती"])

        final_r4_df = pd.concat([df_district_row, pivot_r4])

        display_r4_df = final_r4_df.astype(object)
        for col in display_r4_df.columns:
          display_r4_df[col] = display_r4_df[col].apply(
              lambda x: "" if x == 0 or pd.isna(x) else int(x)
          )

        st.subheader("📋 मोजणी होणेवर शिल्लक प्रकरणे (तालुका व महिना निहाय)")
        st.dataframe(display_r4_df, use_container_width=True)

        col_r4_down1, col_r4_down2 = st.columns(2)

        with col_r4_down1:
          import openpyxl
          from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

          def generate_report4_excel(df_data):
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "मोजणी होणेवर शिल्लक"

            navy_fill = PatternFill(
                start_color="1F4E79", end_color="1F4E79", fill_type="solid"
            )
            yellow_fill = PatternFill(
                start_color="FFE699", end_color="FFE699", fill_type="solid"
            )
            white_bold_font = Font(
                name="Calibri", size=11, bold=True, color="FFFFFF"
            )
            black_bold_font = Font(
                name="Calibri", size=11, bold=True, color="000000"
            )
            thin_border = Border(
                left=Side(style="thin", color="000000"),
                right=Side(style="thin", color="000000"),
                top=Side(style="thin", color="000000"),
                bottom=Side(style="thin", color="000000"),
            )

            tot_cols = len(df_data.columns) + 1
            ws.merge_cells(
                start_row=1, start_column=1, end_row=1, end_column=tot_cols
            )
            main_title = ws.cell(row=1, column=1)
            main_title.value = "मोजणी होणेवर शिल्लक"
            main_title.font = Font(name="Calibri", size=14, bold=True)
            main_title.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws.row_dimensions[1].height = 28

            headers = ["जिल्हा"] + list(df_data.columns)
            ws.append(headers)
            ws.row_dimensions[2].height = 35

            for col_idx in range(1, len(headers) + 1):
              cell = ws.cell(row=2, column=col_idx)
              cell.fill = navy_fill
              cell.font = white_bold_font
              cell.alignment = Alignment(
                  horizontal="center", vertical="center", wrap_text=True
              )
              cell.border = thin_border

            for row_label, row_data in df_data.iterrows():
              row_vals = [row_label] + [
                  "" if v == 0 or pd.isna(v) else int(v)
                  for v in row_data.values
              ]
              ws.append(row_vals)
              curr_r = ws.max_row
              is_district = row_label == "अमरावती"

              for c_idx in range(1, len(row_vals) + 1):
                cell = ws.cell(row=curr_r, column=c_idx)
                cell.border = thin_border

                if is_district:
                  cell.font = black_bold_font
                  cell.alignment = Alignment(
                      horizontal="center" if c_idx > 1 else "left",
                      vertical="center",
                  )
                else:
                  if c_idx == len(row_vals):
                    cell.fill = yellow_fill
                    cell.font = black_bold_font
                  else:
                    cell.font = Font(name="Calibri", size=10)

                  cell.alignment = Alignment(
                      horizontal="center" if c_idx > 1 else "left",
                      vertical="center",
                  )

            for col in ws.columns:
              max_len = max(len(str(cell.value or "")) for cell in col)
              col_letter = openpyxl.utils.get_column_letter(col[0].column)
              ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            wb.save(output)
            return output.getvalue()

          excel_bytes_r4 = generate_report4_excel(final_r4_df)
          st.download_button(
              label="🎨 Download Report 4 Excel (.xlsx)",
              data=excel_bytes_r4,
              file_name="Report_4_Mojani_Honevar_Shillak.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )

        with col_r4_down2:
          if WEASYPRINT_AVAILABLE:
            rows_html_r4 = ""
            for row_label, row_data in final_r4_df.iterrows():
              is_dist = row_label == "अमरावती"
              tr_st = (
                  "background-color: #e2efda; font-weight: bold;"
                  if is_dist
                  else ""
              )

              tds = f"<td style='text-align: left;'><b>{row_label}</b></td>"
              for col_name in final_r4_df.columns:
                val = row_data[col_name]
                val_str = "" if val == 0 or pd.isna(val) else str(int(val))
                bg_style = (
                    "background-color: #fff2cc; font-weight: bold;"
                    if (col_name == "Grand Total" and not is_dist)
                    else ""
                )
                tds += f"<td style='{bg_style}'>{val_str}</td>"

              rows_html_r4 += f"<tr style='{tr_st}'>{tds}</tr>"

            headers_html_r4 = "<th>जिल्हा</th>" + "".join(
                [f"<th>{c}</th>" for c in final_r4_df.columns]
            )

            html_content_r4 = f"""
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <style>
                                @page {{ size: A4 portrait; margin: 12mm; }}
                                body {{ font-family: 'Gargi', 'DejaVu Sans', sans-serif; font-size: 11px; text-align: center; color: #222; }}
                                .main-title {{ font-size: 18px; font-weight: bold; margin-bottom: 12px; color: #1f4e79; }}
                                table {{ width: 85%; margin: 0 auto; border-collapse: collapse; margin-top: 5px; }}
                                th {{ background-color: #1f4e79; color: white; padding: 8px; border: 1px solid #000; font-size: 11px; font-weight: bold; text-align: center; }}
                                td {{ padding: 6px; border: 1px solid #000; font-size: 11px; text-align: center; }}
                                tr:nth-child(even) {{ background-color: #f9f9f9; }}
                            </style>
                        </head>
                        <body>
                            <div class="main-title">मोजणी होणेवर शिल्लक</div>
                            <table>
                                <thead>
                                    <tr>{headers_html_r4}</tr>
                                </thead>
                                <tbody>
                                    {rows_html_r4}
                                </tbody>
                            </table>
                        </body>
                        </html>
                        """
            try:
              pdf_bytes_r4 = HTML(string=html_content_r4).write_pdf()
              st.download_button(
                  label="📄 Download Report 4 PDF",
                  data=pdf_bytes_r4,
                  file_name="Report_4_Mojani_Honevar_Shillak.pdf",
                  mime="application/pdf",
              )
            except Exception as e:
              st.error(f"PDF जनरेट करताना त्रुटी आली: {e}")
          else:
            st.warning("⚠️ PDF जनरेशन उपलब्ध नाही (WeasyPrint नाही).")

else:
  st.info(
      "👈 कृपया डाव्या बाजूच्या Sidebar मधून E-Mojani ची Raw Excel फाईल Upload"
      " करा."
  )
