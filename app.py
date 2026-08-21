import datetime
import pandas as pd
import streamlit as st
from weasyprint import HTML

# Page Setup
st.set_page_config(
    page_title="E-Mojani Filtered Pending Report Generator & Dashboard", layout="wide"
)

st.title("📊 भूमि अभिलेख विभाग - प्रलंबित मोजणी अहवाल व डैशबोर्ड")

# Sidebar - Filters
st.sidebar.header("⚙️ Filter Options")

uploaded_file = st.file_uploader(
    "Upload Raw E-Mojani Excel File (.xlsx)", type=["xlsx"]
)

def parse_excel_date(val):
    if pd.isna(val):
        return pd.NaT
    try:
        val_num = float(val)
        return datetime.datetime(1899, 12, 30) + datetime.timedelta(
            days=val_num
        )
    except:
        return pd.to_datetime(val, errors="coerce")

if uploaded_file is not None:
    df_raw = pd.read_excel(uploaded_file)
    if "मोजणी तारीख" not in df_raw.columns:
        df_raw = pd.read_excel(uploaded_file, header=1)

    if "तालुका" in df_raw.columns:
        df_raw = df_raw[df_raw["तालुका"].notna() & (df_raw["तालुका"] != "")]

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
    start_date = col_d1.date_input("पासून (From)", value=min_excel_date, key="r1_start")
    end_date = col_d2.date_input("पर्यंत (To)", value=datetime.date.today(), key="r1_end")

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

    with st.spinner("Data process ho raha hai..."):
        # 3 TABS: TAB 1, TAB 2, AND NEW DASHBOARD TAB 3
        tab1, tab2, tab3 = st.tabs(
            [
                "📊 Report 1 (तालुका व दिवसनिहाय)", 
                "📋 Report 2 (टप्पा व अधिकारी निहाय)",
                "📈 Executive Dashboard (Top Performance)"
            ]
        )

        from_str = start_date.strftime("%d/%m/%Y")
        to_str = end_date.strftime("%d/%m/%Y")

        # ---------------------------------------------------------------------
        # REPORT 1
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

            rows_html = ""
            for taluka, row in pivot_df.iterrows():
                is_total = taluka == "एकूण"
                tr_style = "background-color: #d9f0a3; font-weight: bold;" if is_total else ""
                tds = f"<td><b>{taluka}</b></td>"
                for col in existing_cols:
                    tds += f"<td style='text-align: center;'>{row[col]}</td>"
                rows_html += f"<tr style='{tr_style}'>{tds}</tr>"

            headers_html = "<th>तालुका</th>" + "".join([f"<th>{c}</th>" for c in existing_cols])

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
            pdf_bytes = HTML(string=html_content).write_pdf()

            st.download_button(
                label="📥 Download PDF Report 1",
                data=pdf_bytes,
                file_name=f"Mojani_Pending_Report1_{from_str.replace('/', '-')}_to_{to_str.replace('/', '-')}.pdf",
                mime="application/pdf",
            )

        # ---------------------------------------------------------------------
        # REPORT 2
        # ---------------------------------------------------------------------
        with tab2:
            st.subheader("📋 Report 2 - टप्पा व अधिकारी निहाय अहवाल")

            col_r2_d1, col_r2_d2 = st.columns(2)
            r2_start_date = col_r2_d1.date_input("पासून (From)", value=min_excel_date, key="r2_start")
            r2_end_date = col_r2_d2.date_input("पर्यंत (To)", value=datetime.date.today(), key="r2_end")

            r2_from_str = r2_start_date.strftime("%d/%m/%Y")
            r2_to_str = r2_end_date.strftime("%d/%m/%Y")

            st.info(f"🗓️ **Report 2 Period:** {r2_from_str} ते {r2_to_str}")

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
                shirastedar = len(df_t[df_t["स्थिती"] == "शिरस्तेदार/मुख्यालय सहाय्यक यांनी तपासले"])
                up_bhoo = len(df_t[df_t["स्थिती"] == "ऊप.अ. भू. अ/ न .भू अ यांच्या मान्यतेवर"])
                off_total = chanani + shirastedar + up_bhoo

                yes_no = len(df_t[df_t[col_yn].notna() & (df_t[col_yn] != "")]) if col_yn in df_t.columns else 0

                if col_mojni_type in df_t.columns:
                    haddi = len(df_t[(df_t[col_mojni_type] == "ह्द्दकायम") & (df_t["स्थिती"] == "मोजणीची माहिती")])
                    jama = len(df_t[(df_t["स्थिती"] == "मोजणीची माहिती") & (df_t[col_mojni_type] != "ह्द्दकायम")])
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
                "शिरस्तेदार/मुख्यालय सहाय्यक": df_rep2["शिरस्तेदार/मुख्यालय सहाय्यक"].sum(),
                "उप अ भू अ/ भू अ": df_rep2["उप अ भू अ/ भू अ"].sum(),
                "Grand Total ": df_rep2["Grand Total "].sum(),
            }

            df_rep2 = pd.concat([df_rep2, pd.DataFrame([total_row])], ignore_index=True)
            st.dataframe(df_rep2, use_container_width=True)

        # ---------------------------------------------------------------------
        # TAB 3: EXECUTIVE DASHBOARD & TOP PERFORMANCE (NEW)
        # ---------------------------------------------------------------------
        with tab3:
            st.subheader("📈 'क प्रत' कामगिरी व विश्लेषण डैशबोर्ड")

            # 1. Calculate Previous Month Dates Automatically
            today = datetime.date.today()
            first_day_of_curr_month = today.replace(day=1)
            last_day_prev_month = first_day_of_curr_month - datetime.timedelta(days=1)
            first_day_prev_month = last_day_prev_month.replace(day=1)

            st.markdown(
                f"##### 🗓️ **पिछले महीने का विश्लेषण:** ({first_day_prev_month.strftime('%d/%m/%Y')} ते {last_day_prev_month.strftime('%d/%m/%Y')})"
            )

            # Filter data for previous month & 'क प्रत' Status
            p_start_dt = pd.to_datetime(first_day_prev_month)
            p_end_dt = pd.to_datetime(last_day_prev_month)

            df_prev_month_kprat = df_raw[
                (df_raw["mojni_date_parsed"] >= p_start_dt)
                & (df_raw["mojni_date_parsed"] <= p_end_dt)
                & (df_raw["स्थिती"] == "क प्रत")
            ].copy()

            col_dash1, col_dash2 = st.columns(2)

            with col_dash1:
                st.markdown("### 🏆 Top 3 तालुका ('क प्रत' निकाली)")
                if not df_prev_month_kprat.empty:
                    top_talukas = (
                        df_prev_month_kprat["तालुका"]
                        .value_counts()
                        .head(3)
                        .reset_index()
                    )
                    top_talukas.columns = ["तालुका", "क प्रत निकाली संख्या"]
                    st.dataframe(top_talukas, use_container_width=True, hide_index=True)
                else:
                    st.warning("पिछले महीने में 'क प्रत' का कोई डेटा उपलब्ध नहीं है।")

            with col_dash2:
                st.markdown("### 👷 Top सर्वेक्षक / अधिकारी ('क प्रत' निकाली)")
                
                # Check potential Surveyor/Employee column names in Excel
                surveyor_col = None
                possible_cols = ["कर्मचारी/अधिकारी चे नाव", "मोजणीदार", "कर्मचारी नाव", "अधिकारी नाव"]
                for c in possible_cols:
                    if c in df_raw.columns:
                        surveyor_col = c
                        break

                if surveyor_col and not df_prev_month_kprat.empty:
                    top_surveyors = (
                        df_prev_month_kprat[surveyor_col]
                        .value_counts()
                        .head(5)
                        .reset_index()
                    )
                    top_surveyors.columns = ["सर्वेक्षक / कर्मचारी नाव", "क प्रत निकाली संख्या"]
                    st.dataframe(top_surveyors, use_container_width=True, hide_index=True)
                elif not surveyor_col:
                    st.info("💡 Excel में सर्वेक्षक के नाम वाला कॉलम दर्ज नहीं मिला। (जैसे: 'कर्मचारी/अधिकारी चे नाव')")
                else:
                    st.warning("पिछले महीने में 'क प्रत' का कोई डेटा उपलब्ध नहीं है।")

            st.markdown("---")
            
            # Quick Performance Summary Bar
            st.markdown("### 📊 त्वरित सारांश (Quick Metrics)")
            m1, m2, m3 = st.columns(3)
            m1.metric(
                label="कुल 'क प्रत' (पिछले महीने)", 
                value=len(df_prev_month_kprat)
            )
            m2.metric(
                label="कुल प्रलंबित केस (Overall Pending)", 
                value=len(df_raw[~df_raw["स्थिती"].isin(completed_defaults)])
            )
            m3.metric(
                label="आज की तिथि तक शिल्लक प्रकरणे", 
                value=len(df_raw[(df_raw["mojni_date_parsed"] > today_dt) & (df_raw["स्थिती"] == "मोजणीची माहिती")])
            )
