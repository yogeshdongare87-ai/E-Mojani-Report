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


# Excel Download Helper
def convert_df_to_excel(df, sheet_name="Report"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name)
    return output.getvalue()


# PDF / Print Layout Generator
def generate_report_pdf_html(df, title="अहवाल"):
    headers = list(df.columns)

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{
                size: A4 portrait;
                margin: 15mm;
            }}
            body {{
                font-family: 'Arial', 'DejaVu Sans', sans-serif;
                margin: 0;
                padding: 10px;
                background-color: #ffffff;
            }}
            .report-title {{
                text-align: center;
                font-size: 22px;
                font-weight: bold;
                color: #000000;
                margin-bottom: 15px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                border: 1.5px solid #000000;
            }}
            th {{
                background-color: #1F4E78;
                color: #ffffff;
                font-weight: bold;
                font-size: 14px;
                padding: 10px 8px;
                border: 1px solid #000000;
                text-align: center;
            }}
            td {{
                padding: 7px 10px;
                font-size: 13px;
                border: 1px solid #000000;
                text-align: center;
                color: #000000;
            }}
            .taluka-col {{
                text-align: left !important;
                padding-left: 12px !important;
                font-weight: 500;
            }}
            .total-row td {{
                font-weight: bold !important;
                background-color: #ffffff !important;
            }}
            .grand-total-col {{
                background-color: #FFF2CC !important;
                font-weight: bold !important;
            }}
        </style>
    </head>
    <body>
        <div class="report-title">{title}</div>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">जिल्हा / तालुका</th>
    """
    for h in headers:
        html += f"<th>{h}</th>"
    html += """
                </tr>
            </thead>
            <tbody>
    """

    for idx, row in df.iterrows():
        is_total = str(idx).strip() == "Total" or str(idx).strip() == "एकूण"
        row_class = "total-row" if is_total else ""

        html += f'<tr class="{row_class}">'
        html += f'<td class="taluka-col">{idx}</td>'

        for col_name in headers:
            val = row[col_name]
            val_str = (
                ""
                if (
                    pd.isna(val)
                    or val == 0
                    or str(val).strip() in ["", "0", "nan"]
                )
                else str(int(val) if isinstance(val, (int, float)) else val)
            )
            col_class = (
                "grand-total-col"
                if col_name == "Grand Total" or col_name == "एकूण"
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


# Sidebar - Single File Upload & Filters
st.sidebar.header("⚙️ File Upload & Options")

uploaded_file = st.sidebar.file_uploader(
    "📁 Main E-Mojani Raw Excel (.xlsx) [All Reports 1, 2, 3, 4, 5 साठी]",
    type=["xlsx"],
    key="single_file",
)


def parse_excel_date(val):
    if pd.isna(val) or str(val).strip() in ["", "nan", "NaT"]:
        return pd.NaT
    try:
        val_num = float(val)
        return datetime.datetime(1899, 12, 30) + datetime.timedelta(
            days=val_num
        )
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

    # 'क प्रत निकाली दिनाक' parsing if available
    col_k_prat_date = "क प्रत निकाली दिनाक"
    if col_k_prat_date in df_raw.columns:
        df_raw["k_prat_date_parsed"] = df_raw[col_k_prat_date].apply(
            parse_excel_date
        )
    else:
        df_raw["k_prat_date_parsed"] = df_raw["mojni_date_parsed"]

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

        # TAB 1: REPORT 1
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
                df1["तालुका"],
                df1["दिवस वर"],
                margins=True,
                margins_name="एकूण",
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
                    file_name="Report_1_Talukani_hay.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                )
            with btn_col2:
                pdf_r1 = generate_report_pdf_html(
                    pivot_df, title="Report 1 - तालुका व दिवसनिहाय"
                )
                st.download_button(
                    label="📄 Download Report 1 (PDF / Print)",
                    data=pdf_r1,
                    file_name=(
                        "Report_1.pdf"
                        if WEASYPRINT_AVAILABLE
                        else "Report_1.html"
                    ),
                    mime=(
                        "application/pdf"
                        if WEASYPRINT_AVAILABLE
                        else "text/html"
                    ),
                )

        # TAB 2: REPORT 2
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

                chanani = len(
                    df_t[df_t[col_status] == "छाननी लिपिक यांनी तपासले"]
                )
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
                    "Stage Total": stage_total,
                    "छाननी लिपीक": chanani,
                    "शिरस्तेदार/मुख्यालय सहाय्यक": shirastedar,
                    "उप अ भू अ/ भू अ": up_bhoo,
                    "Officer Total": off_total,
                })

            df_rep2 = pd.DataFrame(report2_data)
            total_row = {
                "तालुका": "एकूण",
                "Yes/No": df_rep2["Yes/No"].sum(),
                "जमा करणेवर": df_rep2["जमा करणेवर"].sum(),
                "हददी दाखविणेवर": df_rep2["हददी दाखविणेवर"].sum(),
                "शिल्लक प्रकरणे": df_rep2["शिल्लक प्रकरणे"].sum(),
                "Stage Total": df_rep2["Stage Total"].sum(),
                "छाननी लिपीक": df_rep2["छाननी लिपीक"].sum(),
                "शिरस्तेदार/मुख्यालय सहाय्यक": df_rep2[
                    "शिरस्तेदार/मुख्यालय सहाय्यक"
                ].sum(),
                "उप अ भू अ/ भू अ": df_rep2["उप अ भू अ/ भू अ"].sum(),
                "Officer Total": df_rep2["Officer Total"].sum(),
            }

            df_rep2 = pd.concat(
                [df_rep2, pd.DataFrame([total_row])], ignore_index=True
            )
            st.dataframe(df_rep2, use_container_width=True)

            btn_r2_1, btn_r2_2 = st.columns(2)
            with btn_r2_1:
                excel_r2 = convert_df_to_excel(df_rep2, sheet_name="Report_2")
                st.download_button(
                    label="📥 Download Report 2 (Excel)",
                    data=excel_r2,
                    file_name="Report_2_Stage_Officer.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                )
            with btn_r2_2:
                pdf_r2 = generate_report_pdf_html(
                    df_rep2, title="Report 2 - टप्पा व अधिकारी निहाय"
                )
                st.download_button(
                    label="📄 Download Report 2 (PDF / Print)",
                    data=pdf_r2,
                    file_name=(
                        "Report_2.pdf"
                        if WEASYPRINT_AVAILABLE
                        else "Report_2.html"
                    ),
                    mime=(
                        "application/pdf"
                        if WEASYPRINT_AVAILABLE
                        else "text/html"
                    ),
                )

        # TAB 3: EXECUTIVE DASHBOARD
        with tab3:
            st.subheader("📈 Executive Dashboard Overview")

            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                st.markdown(
                    f'<div class="metric-card"><div class="metric-title">एकूण अपलोड प्रकरणे</div><div class="metric-value">{len(df_raw)}</div></div>',
                    unsafe_allow_html=True,
                )
            with col_m2:
                st.markdown(
                    f'<div class="metric-card"><div class="metric-title">प्रलंबित प्रकरणे (Filtered)</div><div class="metric-value">{len(df1)}</div></div>',
                    unsafe_allow_html=True,
                )
            with col_m3:
                today_ref_dash = pd.to_datetime(datetime.date.today())
                r4_count_dash = len(
                    df_raw[
                        (df_raw["mojni_date_parsed"].notna())
                        & (df_raw["mojni_date_parsed"] >= today_ref_dash)
                        & (
                            df_raw[col_status]
                            .astype(str)
                            .str.contains("मोजणीची माहिती")
                        )
                    ]
                )
                st.markdown(
                    f'<div class="metric-card"><div class="metric-title">मोजणी होणेवर शिल्लक</div><div class="metric-value">{r4_count_dash}</div></div>',
                    unsafe_allow_html=True,
                )

            st.markdown("---")
            if "तालुका" in df_raw.columns:
                st.subheader("📊 तालुका निहाय प्रलंबित प्रकरणे चार्ट")
                tal_chart_df = df1["तालुका"].value_counts().reset_index()
                tal_chart_df.columns = ["तालुका", "प्रकरणे"]
                fig_tal = px.bar(
                    tal_chart_df,
                    x="तालुका",
                    y="प्रकरणे",
                    color="प्रकरणे",
                    text_auto=True,
                    color_continuous_scale="Viridis",
                )
                st.plotly_chart(fig_tal, use_container_width=True)

        # TAB 4: REPORT 3
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

                    btn_r3_1, btn_r3_2 = st.columns(2)
                    with btn_r3_1:
                        excel_r3 = convert_df_to_excel(
                            pivot_r3, sheet_name="Report_3"
                        )
                        st.download_button(
                            label="📥 Download Report 3 (Excel)",
                            data=excel_r3,
                            file_name="Report_3_Bhukarmapak.xlsx",
                            mime=(
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            ),
                        )
                    with btn_r3_2:
                        pdf_r3 = generate_report_pdf_html(
                            pivot_r3, title="Report 3 - भूकरमापक निहाय"
                        )
                        st.download_button(
                            label="📄 Download Report 3 (PDF / Print)",
                            data=pdf_r3,
                            file_name=(
                                "Report_3.pdf"
                                if WEASYPRINT_AVAILABLE
                                else "Report_3.html"
                            ),
                            mime=(
                                "application/pdf"
                                if WEASYPRINT_AVAILABLE
                                else "text/html"
                            ),
                        )

        # TAB 5: REPORT 4 (मोजणी होणेवर शिल्लक)
        with tab5:
            st.markdown("## 📅 रिपोर्ट ४: मोजणी होणेवर शिल्लक अहवाल")

            today_ref = pd.to_datetime(datetime.date.today())

            df_r4 = df_raw[
                (df_raw["mojni_date_parsed"].notna())
                & (df_raw["mojni_date_parsed"] >= today_ref)
            ].copy()

            if col_status in df_r4.columns:
                df_r4 = df_r4[
                    df_r4[col_status].astype(str).str.contains("मोजणीची माहिती")
                ]

            st.success(f"🎯 **एकूण मोजणी होणेवर शिल्लक प्रकरणे: {len(df_r4)}**")

            if df_r4.empty:
                st.warning(
                    f"आजच्या तारखेच्या ({today_ref.strftime('%d/%m/%Y')}) पुढे 'मोजणीची माहिती' स्थितीत प्रकरणे आढळली नाहीत."
                )
            else:
                df_r4["Month_Year"] = df_r4["mojni_date_parsed"].dt.strftime("%B")
                df_r4["Year_Month_Sort"] = df_r4[
                    "mojni_date_parsed"
                ].dt.strftime("%Y-%m")

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

                cols_in_pivot = [
                    m for m in sorted_month_names if m in pivot_r4.columns
                ]
                pivot_r4 = pivot_r4[cols_in_pivot]

                standard_talukas = [
                    "अंजनगाव",
                    "अचलपूर",
                    "अमरावती",
                    "चांदुर बा.",
                    "चांदुर रे.",
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

                pivot_r4 = pivot_r4.rename(
                    index={"अचलपुर": "अचलपूर", "दर्यापुर": "दर्यापूर"}
                )

                present_talukas = pivot_r4.index.tolist()
                for t in present_talukas:
                    if t not in standard_talukas:
                        standard_talukas.append(t)

                pivot_r4 = pivot_r4.reindex(standard_talukas, fill_value=0)
                pivot_r4["Grand Total"] = pivot_r4.sum(axis=1)

                total_row = pivot_r4.sum(axis=0)
                pivot_r4.loc["Total"] = total_row

                display_r4_df = pivot_r4.astype(object)
                for col in display_r4_df.columns:
                    display_r4_df[col] = display_r4_df[col].apply(
                        lambda x: "" if x == 0 or pd.isna(x) else int(x)
                    )

                st.subheader("📋 मोजणी होणेवर शिल्लक")
                st.dataframe(display_r4_df, use_container_width=True)

                btn_r4_1, btn_r4_2 = st.columns(2)
                with btn_r4_1:
                    excel_r4 = convert_df_to_excel(
                        pivot_r4, sheet_name="Report_4"
                    )
                    st.download_button(
                        label="📥 Download Report 4 (Excel)",
                        data=excel_r4,
                        file_name="Report_4_Mojni_Shillak.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ),
                    )
                with btn_r4_2:
                    pdf_r4 = generate_report_pdf_html(
                        pivot_r4, title="मोजणी होणेवर शिल्लक"
                    )
                    st.download_button(
                        label="📄 Download Report 4 (PDF / Print Layout)",
                        data=pdf_r4,
                        file_name=(
                            "Report_4_Mojni_Shillak.pdf"
                            if WEASYPRINT_AVAILABLE
                            else "Report_4_Mojni_Shillak.html"
                        ),
                        mime=(
                            "application/pdf"
                            if WEASYPRINT_AVAILABLE
                            else "text/html"
                        ),
                    )

        # TAB 6: REPORT 5 (मागील ३ महिने 'क प्रत' प्रोग्रेस अहवाल)
        with tab6:
            st.markdown(
                "## 🚀 रिपोर्ट ५: मागील ३ महिन्यांचा 'क प्रत' प्रोग्रेस अहवाल"
                " (आजपर्यंत)"
            )

            today_ref_r5 = pd.to_datetime(datetime.date.today())
            three_months_ago_r5 = today_ref_r5 - pd.Timedelta(days=90)

            df_r5 = df_raw[
                (df_raw["k_prat_date_parsed"].notna())
                & (df_raw["k_prat_date_parsed"] >= three_months_ago_r5)
                & (df_raw["k_prat_date_parsed"] <= today_ref_r5)
            ].copy()

            st.success(
                f"🎯 **मागील ३ महिन्यात पूर्ण झालेली एकूण 'क प्रत' प्रकरणे:"
                f" {len(df_r5)}**"
            )

            if df_r5.empty:
                st.warning(
                    "मागील ३ महिन्यात 'क प्रत' दिनांक असलेले प्रकरणे आढळली"
                    " नाहीत."
                )
            else:
                df_r5["Month_Year"] = df_r5["k_prat_date_parsed"].dt.strftime(
                    "%B"
                )
                df_r5["Year_Month_Sort"] = df_r5[
                    "k_prat_date_parsed"
                ].dt.strftime("%Y-%m")

                sorted_months_r5 = (
                    df_r5[["Year_Month_Sort", "Month_Year"]]
                    .drop_duplicates()
                    .sort_values("Year_Month_Sort")
                )
                sorted_month_names_r5 = sorted_months_r5["Month_Year"].tolist()

                pivot_r5 = pd.pivot_table(
                    df_r5,
                    index="तालुका",
                    columns="Month_Year",
                    values=df_r5.columns[0],
                    aggfunc="count",
                    fill_value=0,
                )

                cols_in_pivot_r5 = [
                    m for m in sorted_month_names_r5 if m in pivot_r5.columns
                ]
                pivot_r5 = pivot_r5[cols_in_pivot_r5]

                standard_talukas = [
                    "अंजनगाव",
                    "अचलपूर",
                    "अमरावती",
                    "चांदुर बा.",
                    "चांदुर रे.",
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

                pivot_r5 = pivot_r5.rename(
                    index={"अचलपुर": "अचलपूर", "दर्यापुर": "दर्यापूर"}
                )

                present_talukas = pivot_r5.index.tolist()
                for t in present_talukas:
                    if t not in standard_talukas:
                        standard_talukas.append(t)

                pivot_r5 = pivot_r5.reindex(standard_talukas, fill_value=0)
                pivot_r5["Grand Total"] = pivot_r5.sum(axis=1)

                total_row_r5 = pivot_r5.sum(axis=0)
                pivot_r5.loc["Total"] = total_row_r5

                display_r5_df = pivot_r5.astype(object)
                for col in display_r5_df.columns:
                    display_r5_df[col] = display_r5_df[col].apply(
                        lambda x: "" if x == 0 or pd.isna(x) else int(x)
                    )

                st.subheader(
                    "📋 मागील ३ महिने 'क प्रत' झालेली प्रकरणे (तालुका व महिना"
                    " निहाय)"
                )
                st.dataframe(display_r5_df, use_container_width=True)

                btn_r5_1, btn_r5_2 = st.columns(2)
                with btn_r5_1:
                    excel_r5 = convert_df_to_excel(
                        pivot_r5, sheet_name="Report_5"
                    )
                    st.download_button(
                        label="📥 Download Report 5 (Excel)",
                        data=excel_r5,
                        file_name="Report_5_K_Prat_Progress.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ),
                    )
                with btn_r5_2:
                    pdf_r5 = generate_report_pdf_html(
                        pivot_r5,
                        title="मागील ३ महिने 'क प्रत' प्रोग्रेस अहवाल",
                    )
                    st.download_button(
                        label="📄 Download Report 5 (PDF / Print Layout)",
                        data=pdf_r5,
                        file_name=(
                            "Report_5_K_Prat_Progress.pdf"
                            if WEASYPRINT_AVAILABLE
                            else "Report_5_K_Prat_Progress.html"
                        ),
                        mime=(
                            "application/pdf"
                            if WEASYPRINT_AVAILABLE
                            else "text/html"
                        ),
                    )

else:
    st.info("👈 कृपया डाव्या बाजूच्या Sidebar मधून Excel फाईल Upload करा.")
