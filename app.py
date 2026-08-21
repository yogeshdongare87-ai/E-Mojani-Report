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


def parse_mojni_date(val):
    if pd.isna(val) or val == "" or str(val).strip() == "":
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
    if "तालुका" not in df_raw.columns:
        df_raw = pd.read_excel(uploaded_file, header=1)

    # Clean empty taluka
    if "तालुका" in df_raw.columns:
        df_raw = df_raw[df_raw["तालुका"].notna() & (df_raw["तालुका"] != "")]

    col_mojni_date = "मोजणी तारीख"  # Column I
    col_mojni_type = "मोजणीचा प्रकार(Mojni Type)"
    col_yn = "क्षेत्र अभिलेखाशी मेळात आहे का?"

    # Parse Column I (मोजणी तारीख)
    df_raw["parsed_mojni_date"] = (
        df_raw[col_mojni_date].apply(parse_mojni_date).dt.date
    )

    valid_dates = df_raw["parsed_mojni_date"].dropna()
    min_excel_date = (
        valid_dates.min() if not valid_dates.empty else datetime.date(2023, 1, 1)
    )
    max_excel_date = (
        valid_dates.max() if not valid_dates.empty else datetime.date.today()
    )

    # Global Completed Status List
    completed_defaults = [
        "क प्रत",
        "विनाकार्यवाही",
        "प्रस्तावित बिगरशेती/गुंठेवारी मोजणी पूर्ण",
    ]

    with st.spinner("Data process ho raha hai..."):
        tab1, tab2 = st.tabs(
            ["📊 Report 1 (तालुका व दिवसनिहाय)", "📋 Report 2 (टप्पा व अधिकारी निहाय)"]
        )

        # ---------------------------------------------------------------------
        # REPORT 1
        # ---------------------------------------------------------------------
        with tab1:
            st.sidebar.subheader("📅 Report 1 कालावधी")
            col_d1, col_d2 = st.sidebar.columns(2)
            start_date_r1 = col_d1.date_input(
                "पासून (From)", value=min_excel_date, key="r1_start"
            )
            end_date_r1 = col_d2.date_input(
                "पर्यंत (To)", value=max_excel_date, key="r1_end"
            )

            all_statuses = df_raw["स्थिती"].dropna().unique().tolist()
            default_selected = [
                s for s in all_statuses if s not in completed_defaults
            ]
            selected_statuses = st.sidebar.multiselect(
                "Select Status:",
                options=all_statuses,
                default=default_selected,
            )

            df1 = df_raw[df_raw["स्थिती"].isin(selected_statuses)].copy()
            df1 = df1[
                (df1["parsed_mojni_date"] >= start_date_r1)
                & (df1["parsed_mojni_date"] <= end_date_r1)
            ].copy()

            df1["Pending Days"] = (
                end_date_r1 - df1["parsed_mojni_date"]
            ).apply(lambda x: x.days if pd.notna(x) else None)

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
            st.dataframe(pivot_df, use_container_width=True)

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
                "पर्यंत (To)", value=max_excel_date, key="r2_end"
            )

            r2_from_str = r2_start_date.strftime("%d/%m/%Y")
            r2_to_str = r2_end_date.strftime("%d/%m/%Y")

            taluka_list = sorted(df_raw["तालुका"].dropna().unique().tolist())
            report2_data = []

            for tal in taluka_list:
                # Full Data for Taluka (Unfiltered for Shillak Cases)
                df_taluka_all = df_raw[df_raw["तालुका"] == tal]

                # Filtered Data for Date Range
                df_taluka_filtered = df_taluka_all[
                    (df_taluka_all["parsed_mojni_date"] >= r2_start_date)
                    & (df_taluka_all["parsed_mojni_date"] <= r2_end_date)
                ]

                # 1. Officers Count (From Date Filtered Data)
                chanani = len(
                    df_taluka_filtered[
                        df_taluka_filtered["स्थिती"] == "छाननी लिपिक यांनी तपासले"
                    ]
                )
                shirastedar = len(
                    df_taluka_filtered[
                        df_taluka_filtered["स्थिती"]
                        == "शिरस्तेदार/मुख्यालय सहाय्यक यांनी तपासले"
                    ]
                )
                up_bhoo = len(
                    df_taluka_filtered[
                        df_taluka_filtered["स्थिती"]
                        == "ऊप.अ. भू. अ/ न .भू अ यांच्या मान्यतेवर"
                    ]
                )
                off_total = chanani + shirastedar + up_bhoo

                # 2. Yes/No Count (Col Q)
                if col_yn in df_taluka_filtered.columns:
                    yes_no = len(
                        df_taluka_filtered[
                            df_taluka_filtered[col_yn].notna()
                            & (df_taluka_filtered[col_yn] != "")
                        ]
                    )
                else:
                    yes_no = 0

                # 3. हददी दाखविणेवर Count (Date Filtered)
                if col_mojni_type in df_taluka_filtered.columns:
                    haddi = len(
                        df_taluka_filtered[
                            (df_taluka_filtered[col_mojni_type] == "ह्द्दकायम")
                            & (
                                df_taluka_filtered["स्थिती"]
                                == "मोजणीची माहिती"
                            )
                        ]
                    )
                else:
                    haddi = len(
                        df_taluka_filtered[
                            df_taluka_filtered["स्थिती"] == "मोजणीची माहिती"
                        ]
                    )

                # 4. जमा करणेवर Count (Date Filtered)
                if col_mojni_type in df_taluka_filtered.columns:
                    jama = len(
                        df_taluka_filtered[
                            (
                                df_taluka_filtered["स्थिती"]
                                == "मोजणीची माहिती"
                            )
                            & (
                                df_taluka_filtered[col_mojni_type]
                                != "ह्द्दकायम"
                            )
                        ]
                    )
                else:
                    jama = 0

                # 5. शिल्लक प्रकरणे Count (TOTAL ACTIVE CASES IN EXCEL WITHOUT DATE FILTER)
                all_active = df_taluka_all[
                    ~df_taluka_all["स्थिती"].isin(completed_defaults)
                ]
                shillak = len(
                    all_active[
                        ~all_active["स्थिती"].isin(
                            [
                                "छाननी लिपिक यांनी तपासले",
                                "शिरस्तेदार/मुख्यालय सहाय्यक यांनी तपासले",
                                "ऊप.अ. भू. अ/ न .भू अ यांच्या मान्यतेवर",
                                "मोजणीची माहिती",
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
