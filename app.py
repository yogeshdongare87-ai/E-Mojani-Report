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

# Helper functions for downloading Excel and PDF
def convert_df_to_excel(df, sheet_name="Report"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name)
    return output.getvalue()


def convert_df_to_pdf_html(df, title="Report"):
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            h2 {{ color: #1b4332; text-align: center; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #dddddd; text-align: center; padding: 8px; font-size: 12px; }}
            th {{ background-color: #2b9348; color: white; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>
        <h2>{title}</h2>
        {df.to_html(classes='data', header=True, index=True)}
    </body>
    </html>
    """
    if WEASYPRINT_AVAILABLE:
        try:
            return HTML(string=html_content).write_pdf()
        except Exception:
            return html_content.encode('utf-8')
    else:
        return html_content.encode('utf-8')


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

    # Dynamic Column Detection
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

        # ---------------------------------------------------------------------
        # TAB 1: REPORT 1 (तालुका व दिवसनिहाय)
        # ---------------------------------------------------------------------
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

            # Download Buttons
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                excel_r1 = convert_df_to_excel(pivot_df, sheet_name="Report_1")
                st.download_button(
                    label="📥 Download Report 1 (Excel)",
                    data=excel_r1,
                    file_name="Report_1_Talukani hay.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with btn_col2:
                pdf_r1 = convert_df_to_pdf_html(pivot_df, title="Report 1 - तालुका व दिवसनिहाय")
                st.download_button(
                    label="📄 Download Report 1 (PDF / Print)",
                    data=pdf_r1,
                    file_name="Report_1.pdf" if WEASYPRINT_AVAILABLE else "Report_1.html",
                    mime="application/pdf" if WEASYPRINT_AVAILABLE else "text/html",
                )

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
                df_t =
