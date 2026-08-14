import datetime
import pandas as pd
import streamlit as st
from weasyprint import HTML

# Page Setup
st.set_page_config(
    page_title="E-Mojani Consolidated Report Generator", layout="wide"
)

st.title("📊 भूमि अभिलेख विभाग - प्रलंबित मोजणी अहवाल (Report 1 & Report 2)")

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

    # Date Selector
    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 अहवाल दिनांक (Report Date)")
    report_date = st.sidebar.date_input(
        "अहवाल तारीख निवडा", datetime.date.today()
    )
    formatted_report_date = report_date.strftime("%d/%m/%Y")

    # Filter Completed Cases
    completed_statuses = [
        "क प्रत",
        "विनाकार्यवाही",
        "प्रस्तावित बिगरशेती/गुंठेवारी मोजणी पूर्ण",
    ]

    with st.spinner("Dono Reports Generate ho rahe hain..."):
        # --- TAB 1 & TAB 2 IN STREAMLIT ---
        tab1, tab2 = st.tabs(
            ["📊 Report 1 (दिवसनिहाय)", "📋 Report 2 (टप्पा व अधिकारी निहाय)"]
        )

        # ---------------------------------------------------------------------
        # REPORT 1: DAYWISE PENDING REPORT
        # ---------------------------------------------------------------------
        with tab1:
            st.subheader(
                f"1️⃣ तालुका व दिवसनिहाय प्रलंबित प्रकरणे (दिनांक:"
                f" {formatted_report_date})"
            )

            df1 = df_raw[~df_raw["स्थिती"].isin(completed_statuses)].copy()
            as_on_dt = pd.to_datetime(report_date)
            df1["Pending Days"] = (as_on_dt - df1["mojni_date_parsed"]).dt.days

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
            st.subheader(
                f"2️⃣ प्रलंबित टप्पा व अधिकारी निहाय अहवाल (दिनांक:"
                f" {formatted_report_date})"
            )

            # Taluka list
            taluka_list = sorted(df_raw["तालुका"].dropna().unique().tolist())

            report2_data = []

            for tal in taluka_list:
                df_t = df_raw[df_raw["तालुका"] == tal]

                # Officer counts from 'स्थिती'
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

                # Stage counts logic
                yes_no = len(
                    df_t[
                        df_t["स्थिती"].isin(
                            ["क्षेत्र अभिलेखाशी मेळात आहे का?"]
                        )
                    ]
                )
                jama = len(df_t[df_t["स्थिती"] == "सादर केलेला अर्ज"])
                haddi = len(df_t[df_t["स्थिती"] == "मोजणीची माहिती"])

                # Remaining pending
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
                    "Grand Total (Stage)": stage_total,
                    "छाननी लिपीक": chanani,
                    "शिरस्तेदार/मुख्यालय सहाय्यक": shirastedar,
                    "उप अ भू अ/ भू अ": up_bhoo,
                    "Grand Total (Officer)": off_total,
                })

            df_rep2 = pd.DataFrame(report2_data)

            # Sum Row
            total_row = {
                "तालुका": "एकूण",
                "Yes/No": df_rep2["Yes/No"].sum(),
                "जमा करणेवर": df_rep2["जमा करणेवर"].sum(),
                "हददी दाखविणेवर": df_rep2["हददी दाखविणेवर"].sum(),
                "शिल्लक प्रकरणे": df_rep2["शिल्लक प्रकरणे"].sum(),
                "Grand Total (Stage)": df_rep2["Grand Total (Stage)"].sum(),
                "छाननी लिपीक": df_rep2["छाननी लिपीक"].sum(),
                "शिरस्तेदार/मुख्यालय सहाय्यक": df_rep2[
                    "शिरस्तेदार/मुख्यालय सहाय्यक"
                ].sum(),
                "उप अ भू अ/ भू अ": df_rep2["उप अ भू अ/ भू अ"].sum(),
                "Grand Total (Officer)": df_rep2["Grand Total (Officer)"].sum(),
            }

            df_rep2 = pd.concat(
                [df_rep2, pd.DataFrame([total_row])], ignore_index=True
            )

            st.dataframe(df_rep2, use_container_width=True)

            # --- PDF GENERATION FOR REPORT 2 ---
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
                    <td><b>{row['तालुका']}</b></td>
                    <td style="text-align:center;">{row['Yes/No']}</td>
                    <td style="text-align:center;">{row['जमा करणेवर']}</td>
                    <td style="text-align:center;">{row['हददी दाखविणेवर']}</td>
                    <td style="text-align:center;">{row['शिल्लक प्रकरणे']}</td>
                    <td style="text-align:center; background-color: #eef;"><b>{row['Grand Total (Stage)']}</b></td>
                    <td style="text-align:center;">{row['छाननी लिपीक']}</td>
                    <td style="text-align:center;">{row['शिरस्तेदार/मुख्यालय सहाय्यक']}</td>
                    <td style="text-align:center;">{row['उप अ भू अ/ भू अ']}</td>
                    <td style="text-align:center; background-color: #eef;"><b>{row['Grand Total (Officer)']}</b></td>
                </tr>
                """

            html_content_r2 = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    @page {{ size: A4 landscape; margin: 10mm; }}
                    body {{ font-family: 'Gargi', 'DejaVu Sans', sans-serif; font-size: 10px; }}
                    h2 {{ text-align: center; margin-bottom: 2px; color: #1b4332; }}
                    .date-header {{ text-align: center; font-size: 12px; font-weight: bold; margin-bottom: 10px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 5px; }}
                    th {{ background-color: #4a5568; color: white; padding: 5px; border: 1px solid #333; font-size: 9px; text-align: center; }}
                    td {{ padding: 4px; border: 1px solid #777; font-size: 9px; }}
                    tr:nth-child(even) {{ background-color: #f8f9fa; }}
                </style>
            </head>
            <body>
                <h2>भूमि अभिलेख विभाग - प्रलंबित अहवाल</h2>
                <div class="date-header">दिनांक {formatted_report_date}</div>
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
                            <th>शिरस्तेदार / मुख्यालय सहाय्यक</th>
                            <th>उप अ भू अ / भू अ</th>
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
                    f"Stage_Officer_Pending_Report_{formatted_report_date.replace('/', '-')}.pdf"
                ),
                mime="application/pdf",
            )
