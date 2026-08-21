import datetime
import pandas as pd
import streamlit as st
from weasyprint import HTML

# Page Setup
st.set_page_config(
    page_title="E-Mojani Filtered Pending Report Generator", layout="wide"
)

st.title("📊 भूमि अभिलेख विभाग - प्रलंबित मोजणी अहवाल")

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
    # Read Excel Data
    df_raw = pd.read_excel(uploaded_file)
    if "मोजणी तारीख" not in df_raw.columns:
        df_raw = pd.read_excel(uploaded_file, header=1)

    # Clean empty taluka rows
    if "तालुका" in df_raw.columns:
        df_raw = df_raw[df_raw["तालुका"].notna() & (df_raw["तालुका"] != "")]

    # Parse Column I (मोजणी तारीख)
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

    # --- REPORT 1 SIDEBAR DATES ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 Report 1 कालावधी (Date Range)")

    col_d1, col_d2 = st.sidebar.columns(2)
    start_date = col_d1.date_input(
        "पासून (From)", value=min_excel_date, key="r1_start"
    )
    end_date = col_d2.date_input(
        "पर्यंत (To)", value=datetime.date.today(), key="r1_end"
    )

    # --- STATUS FILTER IN SIDEBAR ---
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
        tab1, tab2 = st.tabs(
            ["📊 Report 1 (तालुका व दिवसनिहाय)", "📋 Report 2 (टप्पा व अधिकारी निहाय)"]
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
                tr_style = (
                    "background-color: #d9f0a3; font-weight: bold;"
                    if is_total
                    else ""
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
                    <thead>
                        <tr>{headers_html}</tr>
                    </thead>
                    <tbody>
                        {rows_html}
                    </tbody>
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
            r2_start_date = col_r2_d1.date_input(
                "पासून (From)", value=min_excel_date, key="r2_start"
            )
            r2_end_date = col_r2_d2.date_input(
                "पर्यंत (To)", value=datetime.date.today(), key="r2_end"
            )

            r2_from_str = r2_start_date.strftime("%d/%m/%Y")
            r2_to_str = r2_end_date.strftime("%d/%m/%Y")

            st.info(f"🗓️ **Report 2 Period:** {r2_from_str} ते {r2_to_str}")

            r2_start_dt = pd.to_datetime(r2_start_date)
            r2_end_dt = pd.to_datetime(r2_end_date)
            today_dt = pd.to_datetime(datetime.date.today())

            # Date Range Filtered Data (For Officers, Yes/No, Jama, Haddi)
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

                # Officer counts
                chanani = len(
                    df_t[df_t["स्थिती"] == "छाननी लिपिक यांनी तपासले"]
                )
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

                # 1. Yes/No Count (Col Q)
                if col_yn in df_t.columns:
                    yes_no = len(df_t[df_t[col_yn].notna() & (df_t[col_yn] != "")])
                else:
                    yes_no = 0

                # 2. हददी दाखविणेवर Count
                if col_mojni_type in df_t.columns:
                    haddi = len(
                        df_t[
                            (df_t[col_mojni_type] == "ह्द्दकायम")
                            & (df_t["स्थिती"] == "मोजणीची माहिती")
                        ]
                    )
                else:
                    haddi = len(df_t[df_t["स्थिती"] == "मोजणीची माहिती"])

                # 3. जमा करणेवर Count
                if col_mojni_type in df_t.columns:
                    jama = len(
                        df_t[
                            (df_t["स्थिती"] == "मोजणीची माहिती")
                            & (df_t[col_mojni_type] != "ह्द्दकायम")
                        ]
                    )
                else:
                    jama = 0

                # 4. 🔥 शिल्लक प्रकरणे (YOUR EXACT LOGIC)
                # Column I (mojni_date_parsed) > Today Date AND Column K (स्थिती) == 'मोजणीची माहिती'
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

            # Total Row
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

            # PDF Generation
            rows_html_r2 = ""
            for idx, row in df_rep2.iterrows():
                is_tot = row["तालुका"] == "एकूण"
                st_cls = (
                    "background-color: #d3d3d3; font-weight: bold;"
                    if is_tot
                    else ""
                )
                rows_html_r2 += f"""
                <tr style="{st_cls}">
                    <td style="text-align:left;"><b>{row['तालुका']}</b></td>
                    <td>{row['Yes/No']}</td>
                    <td>{row['जमा करणेवर']}</td>
                    <td>{row['हददी दाखविणेवर']}</td>
                    <td>{row['शिल्लक प्रकरणे']}</td>
                    <td style="background-color: #f0f0f0;"><b>{row['Grand Total']}</b></td>
                    <td>{row['छाननी लिपीक']}</td>
                    <td>{row['शिरस्तेदार/मुख्यालय सहाय्यक']}</td>
                    <td>{row['उप अ भू अ/ भू अ']}</td>
                    <td style="background-color: #f0f0f0;"><b>{row['Grand Total ']}</b></td>
                </tr>
                """

            html_content_r2 = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    @page {{ size: A4 landscape; margin: 10mm; }}
                    body {{ font-family: 'Gargi', 'DejaVu Sans', sans-serif; font-size: 11px; text-align: center; }}
                    .title {{ font-size: 16px; font-weight: bold; margin-bottom: 15px; text-align: center; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 5px; }}
                    th {{ background-color: #bfbfbf; color: black; padding: 6px; border: 1px solid #000; font-size: 10px; text-align: center; }}
                    td {{ padding: 5px; border: 1px solid #000; font-size: 10px; text-align: center; }}
                    tr:nth-child(even) {{ background-color: #fdfdfd; }}
                </style>
            </head>
            <body>
                <div class="title">दिनांक {r2_from_str} ते {r2_to_str}</div>
                <table>
                    <thead>
                        <tr>
                            <th>तालुका</th>
                            <th>Yes/No</th>
                            <th>जमा करणेवर</th>
                            <th>हददी दाखविणेवर</th>
                            <th>शिल्लक प्रकरणे</th>
                            <th>Grand Total</th>
                            <th>छाननी लिपीक</th>
                            <th>शिरस्तेदार/<br>मुख्यालय सहाय्यक</th>
                            <th>उप अ भू अ/<br>भू अ</th>
                            <th>Grand Total</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows_html_r2}
                    </tbody>
                </table>
            </body>
            </html>
            """

            pdf_bytes_r2 = HTML(string=html_content_r2).write_pdf()

            st.download_button(
                label="📥 Download Report 2 PDF",
                data=pdf_bytes_r2,
                file_name=(
                    f"Report_2_{r2_from_str.replace('/', '-')}_to_{r2_to_str.replace('/', '-')}.pdf"
                ),
                mime="application/pdf",
            )
