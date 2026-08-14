import datetime
import pandas as pd
import streamlit as st
from weasyprint import HTML

# Page Setup
st.set_page_config(
    page_title="E-Mojani Filtered Pending Report Generator", layout="wide"
)

st.title("📊 तालुका व दिवसनिहाय प्रलंबित मोजणी अहवाल")

# Sidebar - Filters
st.sidebar.header("⚙️ Filter Options")

# 1. Report As On Date
report_date = st.sidebar.date_input("🗓️ Report Date (अहवाल तारीख)", datetime.date.today())

# 2. File Upload
uploaded_file = st.file_uploader("Upload Raw E-Mojani Excel File (.xlsx)", type=["xlsx"])


def parse_excel_date(val):
    if pd.isna(val):
        return pd.NaT
    try:
        val_num = float(val)
        return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=val_num)
    except:
        return pd.to_datetime(val, errors='coerce')


if uploaded_file is not None:
    # Read Excel Data
    df_raw = pd.read_excel(uploaded_file)
    if 'मोजणी तारीख' not in df_raw.columns:
        df_raw = pd.read_excel(uploaded_file, header=1)

    # Filter empty taluka rows
    if 'तालुका' in df_raw.columns:
        df_raw = df_raw[df_raw['तालुका'].notna() & (df_raw['तालुका'] != '')]

    # --- STATUS FILTER IN SIDEBAR ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("📌 स्थिती (Status) Filter")

    # Get unique statuses from Excel
    all_statuses = df_raw['स्थिती'].dropna().unique().tolist()

    # Default statuses (Excluding completed cases by default)
    completed_defaults = ['क प्रत', 'विनाकार्यवाही', 'प्रस्तावित बिगरशेती/गुंठेवारी मोजणी पूर्ण']
    default_selected = [s for s in all_statuses if s not in completed_defaults]

    # Multiselect filter button / dropdown
    selected_statuses = st.sidebar.multiselect(
        "Kiski report nikalni hai select karein:",
        options=all_statuses,
        default=default_selected
    )

    # Quick Select Buttons
    col_btn1, col_btn2 = st.sidebar.columns(2)
    if col_btn1.button("Select All"):
        selected_statuses = all_statuses
    if col_btn2.button("Clear All"):
        selected_statuses = []

    # --- DATA PROCESSING ---
    with st.spinner('Data process ho raha hai...'):
        # Filter by selected Statuses
        df = df_raw[df_raw['स्थिती'].isin(selected_statuses)].copy()

        # Date Parsing
        df['mojni_date_parsed'] = df['मोजणी तारीख'].apply(parse_excel_date)

        # Calculate pending days based on Report Date
        as_on_dt = pd.to_datetime(report_date)
        df['Pending Days'] = (as_on_dt - df['mojni_date_parsed']).dt.days

        # Bucket into Day ranges
        def day_bucket(days):
            if pd.isna(days):
                return 'तारीख उपलब्ध नाही'
            if days <= 15:
                return '15 दिवस वर'
            elif days <= 30:
                return '30 दिवस वर'
            elif days <= 60:
                return '60 दिवस वर'
            elif days <= 90:
                return '90 दिवस वर'
            else:
                return '90 दिवसांपेक्षा जास्त'

        df['दिवस वर'] = df['Pending Days'].apply(day_bucket)

        bucket_order = [
            '15 दिवस वर',
            '30 दिवस वर',
            '60 दिवस वर',
            '90 दिवस वर',
            '90 दिवसांपेक्षा जास्त',
        ]

        # Pivot Table Creation
        pivot_df = pd.crosstab(df['तालुका'], df['दिवस वर'], margins=True, margins_name='एकूण')

        # Reorder columns
        existing_cols = [c for c in bucket_order if c in pivot_df.columns]
        if 'तारीख उपलब्ध नाही' in pivot_df.columns:
            existing_cols.append('तारीख उपलब्ध नाही')
        if 'एकूण' in pivot_df.columns:
            existing_cols.append('एकूण')

        pivot_df = pivot_df[existing_cols]

        st.success(f'✅ Total Cases Filtered: **{len(df)}**')

        # Display Date Range Info
        min_date = df['mojni_date_parsed'].min()
        st.info(f"📅 **Pending Period:** {min_date.strftime('%d/%m/%Y') if pd.notna(min_date) else 'N/A'} से {report_date.strftime('%d/%m/%Y')} तक")

        st.subheader("📋 तालुका व दिवसनिहाय प्रलंबित प्रकरणे Table")
        st.dataframe(pivot_df, use_container_width=True)

        # PDF Generation
        formatted_report_date = report_date.strftime('%d/%m/%Y')

        rows_html = ""
        for taluka, row in pivot_df.iterrows():
            is_total = (taluka == 'एकूण')
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
                @page {{
                    size: A4 portrait;
                    margin: 15mm;
                }}
                body {{
                    font-family: 'Gargi', 'DejaVu Sans', sans-serif;
                    font-size: 11px;
                }}
                h2 {{
                    text-align: center;
                    margin-bottom: 5px;
                    color: #1b4332;
                }}
                .date-header {{
                    text-align: center;
                    font-size: 13px;
                    font-weight: bold;
                    margin-bottom: 15px;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 10px;
                }}
                th {{
                    background-color: #2b9348;
                    color: white;
                    padding: 6px;
                    border: 1px solid #555;
                    font-size: 10px;
                }}
                td {{
                    padding: 5px;
                    border: 1px solid #888;
                    font-size: 10px;
                }}
                tr:nth-child(even) {{
                    background-color: #f9f9f9;
                }}
            </style>
        </head>
        <body>
            <h2>भूमि अभिलेख विभाग - अमरावती</h2>
            <div class="date-header">तालुका व दिवसनिहाय प्रलंबित मोजणी प्रकरणे अहवाल (दिनांक {formatted_report_date} पर्यंत)</div>
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
            label="📥 Download PDF Report",
            data=pdf_bytes,
            file_name=f"Mojani_Pending_Report_{formatted_report_date.replace('/', '-')}.pdf",
            mime="application/pdf",
        )
