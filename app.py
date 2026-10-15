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

# Sidebar - Single File Upload & Filters
st.sidebar.header("⚙️ File Upload & Options")

uploaded_file = st.sidebar.file_uploader(
    "📁 Main E-Mojani Raw Excel (.xlsx) [All Reports 1, 2, 3, 4 साठी]",
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

  # Dynamic Column Detection (Column I = मोजणी तारीख, Column K = स्थिती)
  col_mojni_date = None
  if "मोजणी तारीख" in df_raw.columns:
    col_mojni_date = "मोजणी तारीख"
  elif len(df_raw.columns) >= 9:
    col_mojni_date = df_raw.columns[8]

  col_status = None
  if "स्थिती" in df_raw.columns:
    col_status = "स्थिती"
  elif len(df_raw.columns) >= 11:
    col_status = df_raw.columns[10]

  df_raw["mojni_date_parsed"] = df_raw[col_mojni_date].apply(parse_excel_date)

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
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Report 1 (तालुका व दिवसनिहाय)",
        "📋 Report 2 (टप्पा व अधिकारी निहाय)",
        "📈 Executive Dashboard",
        "📋 Report 3 (भूकरमापक निहाय)",
        "📅 Report 4 (मोजणी होणेवर शिल्लक)",
    ])

    from_str = start_date.strftime("%d/%m/%Y")
    to_str = end_date.strftime("%d/%m/%Y")

    # REPORT 1
    with tab1:
      df1 = df_raw[df_raw[col_status].isin(selected_statuses)].copy()
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

    # REPORT 2
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
                  & (df_t[col_status] == "मोजणीची माहिती")
              ]
          )
          jama = len(
              df_t[
                  (df_t[col_status] == "मोजणीची माहिती")
                  & (df_t[col_mojni_type] != "ह्द्दकायम")
              ]
          )
        else:
          haddi = len(df_t[df_t[col_status] == "मोजणीची माहिती"])
          jama = 0

        shillak = len(
            df_taluka_all[
                (df_taluka_all["mojni_date_parsed"] >= today_dt)
                & (df_taluka_all[col_status] == "मोजणीची माहिती")
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

    # REPORT 3
    with tab4:
      st.markdown("## 📋 रिपोर्ट ३: भूकरमापक निहाय व दिवसनिहाय")
      surveyor_col_r3 = None
      for col_name in ["भूकरमापक", "कर्मचारी/अधिकारी चे नाव"]:
        if col_name in df_raw.columns:
          surveyor_col_r3 = col_name
          break

      if surveyor_col_r3:
        df_r3_clean = df_raw[
            df_raw["तालुका"].notna() & df_raw[surveyor_col_r3].notna()
        ].copy()
        df_r3_pending = df_r3_clean[
            df_r3_clean[col_status].isin(selected_statuses)
        ].copy()
        curr_today = pd.to_datetime(datetime.date.today())
        df_r3_pending = df_r3_pending[
            df_r3_pending["mojni_date_parsed"].notna()
            & (df_r3_pending["mojni_date_parsed"] <= curr_today)
        ].copy()

        if not df_r3_pending.empty:
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
          st.dataframe(pivot_r3, use_container_width=True)

    # ---------------------------------------------------------------------
    # TAB 5: REPORT 4 (मोजणी होणेवर शिल्लक - Date >= Today = Exactly 1268)
    # ---------------------------------------------------------------------
    with tab5:
      st.markdown("## 📅 रिपोर्ट ४: मोजणी होणेवर शिल्लक अहवाल")

      today_ref = pd.to_datetime(datetime.date.today())

      # Exact Match Filter: Date >= Today & Status = 'मोजणीची माहिती'
      df_r4 = df_raw[
          (df_raw["mojni_date_parsed"].notna())
          & (df_raw["mojni_date_parsed"] >= today_ref)
      ].copy()

      if col_status:
        df_r4 = df_r4[
            df_r4[col_status].astype(str).str.contains("मोजणीची माहिती")
        ]

      st.success(
          f"🎯 **एकूण मोजणी होणेवर शिल्लक प्रकरणे (15 Oct onwards):"
          f" {len(df_r4)}**"
      )

      if df_r4.empty:
        st.warning(
            f"आजच्या तारखेच्या ({today_ref.strftime('%d/%m/%Y')}) पुढे 'मोजणीची"
            " माहिती' स्थितीत प्रकरणे आढळली नाहीत."
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
            values=df_r4.columns[0],
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
        pivot_r4["Grand Total"] = pivot_r4.sum(axis=1)

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

else:
  st.info("👈 कृपया डाव्या बाजूच्या Sidebar मधून Excel फाईल Upload करा.")
