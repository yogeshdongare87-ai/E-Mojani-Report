import datetime
import pandas as pd
import streamlit as st
from weasyprint import HTML

# Page Setup
st.set_page_config(
    page_title="E-Mojani Consolidated Report Generator", layout="wide"
)

st.title("📊 भूमि अभिलेख विभाग - प्रलंबित मोजणी अहवाल")

# Sidebar Filters
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

    # Clean Taluka
    if "तालुका" in df_raw.columns:
        df_raw = df_raw[df_raw["तालुका"].notna() & (df_raw["तालुका"] != "")]

    # Date parsing
    df_raw["mojni_date_parsed"] = df_raw["मोजणी तारीख"].apply(parse_excel_date)

    # Calculate default min and max dates from Excel
    min_date = df_raw["mojni_date_parsed"].min()
    max_date = df_raw["mojni_date_parsed"].max()

    default_start = (
        min_date.date()
        if pd.notna(min_date)
        else datetime.date(2023, 1, 1)
    )
    default_end = datetime.date.today()

    # --- DATE RANGE SELECTION (कालावधी निवडा) ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 अहवाल कालावधी (Date Range)")

    col_d1, col_d2 = st.sidebar.columns(2)
    start_date = col_d1.date_input("पासून (From)", value=default_start)
    end_date = col_d2.date_input("पर्यंत (To)", value=default_end)

    formatted_start = start_date.strftime("%d/%m/%Y")
    formatted_end = end_date.strftime("%d/%m/%Y")

    # Filter Completed Cases
    completed_statuses = [
        "क प्रत",
        "विनाकार्यवाही",
        "प्रस्तावित बिगरशेती/गुंठेवारी मोजणी पूर्ण",
    ]

    # Filter Excel Data between selected Date Range
    start_dt = pd.to_datetime(start_date)
    end_dt = pd.to_datetime(end_date)

    df_filtered = df_raw[
        (df_raw["mojni_date_parsed"] >= start_dt)
        & (df_raw["mojni_date_parsed"] <= end_dt)
    ].copy()

    with st.spinner("Selected Date Range ke hisab se report generate ho rahi hai..."):
        tab1, tab2 = st.tabs(
            ["📊 Report 1 (दिवसनिहाय)", "📋 Report 2 (टप्पा व अधिकारी निहाय)"]
        )

        # ---------------------------------------------------------------------
        # REPORT 1: DAYWISE PENDING REPORT
        # ---------------------------------------------------------------------
        with tab1:
            st.subheader(
                f"1️⃣ तालुका व दिवसनिहाय प्रलंबित प्रकरणे ({formatted_start} ते {formatted_end})"
            )

            df1 = df_filtered[~df_filtered["स्थिती"].isin(completed_statuses)].copy()
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

            pivot1 = pd.crosstab(
                df1["तालुका"],
                df1["दिवस वर"],
                margins=True,
                margins_name="एकूण",
            )
            cols1 = [c for c in bucket_order if c in pivot1.columns]
            if "तारीख उपलब्ध नाही" in pivot1.columns:
                cols1.append("तारीख उपलब्ध नाही")
            if "एकूण" in pivot1.columns:
                cols1.append("एकूण")

            pivot1 = pivot1[cols1]
            st.dataframe(pivot1, use_container_width=True)

        # ---------------------------------------------------------------------
        # REPORT 2: STAGE & OFFICER-WISE SUMMARY REPORT
        # ---------------------------------------------------------------------
        with tab2:
            st.subheader(f"दिनांक {formatted_start} ते {formatted_end}")

            taluka_list = sorted(df_raw["तालुका"].dropna().unique().tolist())
            report2_data = []

            for tal in taluka_list:
                df_t = df_filtered[df_filtered["तालुका"] == tal]

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

                yes_no = len(
                    df_t[
                        df_t["स्थिती"].isin(
                            ["क्षेत्र अभिलेखाशी मेळात आहे का?"]
                        )
                    ]
                )
                jama = len(df_t[df_t["स्थिती"] == "सादर केलेला अर्ज"])
                haddi = len(df_t[df_t["स्थिती"] == "मोजणीची माहिती"])

                all_active = df_t[~df_t["स्थिती"].isin(completed_statuses)]
                shillak = len(
                    all_active[
                        ~all_active["स्थिती"].isin(
                            [
                                "छाननी लिपिक यांनी तपासले",
                                "शिरस्तेदार/मुख्यालय सहाय्यक यांनी तपासले",
                                "ऊप.अ. भू. अ/ न .भू अ यांच्या मान्यतेवर",
                            ]
                        )
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

            # PDF HTML Generation
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
                <div class="title">दिनांक {formatted_start} ते {formatted_end}</div>
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
                    f"Report_From_{formatted_start.replace('/', '-')}_to_{formatted_end.replace('/', '-')}.pdf"
                ),
                mime="application/pdf",
            )
