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


# PDF / Print Layout Matching the Exact Image Design
def generate_report4_pdf_html(df, title="मोजणी होणेवर शिल्लक"):
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
                    <th style="width: 25%;">जिल्हा</th>
    """
    for h in headers:
        html += f'<th>{h}</th>'
    html += """
                </tr>
            </thead>
            <tbody>
    """

    for idx, row in df.iterrows():
        is_total = str(idx).strip() == 'Total'
        row_class = 'total-row' if is_total else ''

        html += f'<tr class="{row_class}">'
        html += f'<td class="taluka-col">{idx}</td>'

        for col_name in headers:
            val = row[col_name]
            val_str = (
                ''
                if (
                    pd.isna(val)
                    or val == 0
                    or str(val).strip() in ['', '0', 'nan']
                )
                else str(int(val) if isinstance(val, (int, float)) else val)
            )
            col_class = 'grand-total-col' if col_name == 'Grand Total' else ''
            html += f'<td class="{col_class}">{val_str}</td>'

        html += '</tr>'

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
            return html.encode('utf-8')
    else:
        return html.encode('utf-8')


# Sidebar - Single File Upload & Filters
st.sidebar.header('⚙️ File Upload & Options')

uploaded_file = st.sidebar.file_uploader(
    '📁 Main E-Mojani Raw Excel (.xlsx) [All Reports 1, 2, 3, 4 साठी]',
    type=['xlsx'],
    key='single_file',
)


def parse_excel_date(val):
    if pd.isna(val) or str(val).strip() in ['', 'nan', 'NaT']:
        return pd.NaT
    try:
        val_num = float(val)
        return datetime.datetime(1899, 12, 30) + datetime.timedelta(
            days=val_num
        )
    except Exception:
        return pd.to_datetime(val, dayfirst=True, errors='coerce')


if uploaded_file is not None:
    df_raw = pd.read_excel(uploaded_file)
    if 'मोजणी तारीख' not in df_raw.columns:
        df_raw = pd.read_excel(uploaded_file, header=1)

    # Text Cleaning
    if 'तालुका' in df_raw.columns:
        df_raw['तालुका'] = df_raw['तालुका'].astype(str).str.strip()
        df_raw = df_raw[
            df_raw['तालुका'].notna()
            & (~df_raw['तालुका'].isin(['', 'nan', 'None', '0']))
        ]

    # Dynamic Column Detection
    col_mojni_date = None
    if 'मोजणी तारीख' in df_raw.columns:
        col_mojni_date = 'मोजणी तारीख'
    elif len(df_raw.columns) >= 9:
        col_mojni_date = df_raw.columns[8]

    col_status = None
    if 'स्थिती' in df_raw.columns:
        col_status = 'स्थिती'
    elif len(df_raw.columns) >= 11:
        col_status = df_raw.columns[10]

    df_raw['mojni_date_parsed'] = df_raw[col_mojni_date].apply(parse_excel_date)

    min_excel_date = df_raw['mojni_date_parsed'].min()
    max_excel_date = df_raw['mojni_date_parsed'].max()

    if pd.isna(min_excel_date):
        min_excel_date = datetime.date(2023, 1, 1)
    else:
        min_excel_date = min_excel_date.date()

    if pd.isna(max_excel_date):
        max_excel_date = datetime.date.today()
    else:
        max_excel_date = max_excel_date.date()

    # Sidebar Filters
    st.sidebar.markdown('---')
    st.sidebar.subheader('📅 Report 1 कालावधी (Date Range)')

    col_d1, col_d2 = st.sidebar.columns(2)
    start_date = col_d1.date_input(
        'पासून (From)', value=min_excel_date, key='r1_start'
    )
    end_date = col_d2.date_input(
        'पर्यंत (To)', value=datetime.date.today(), key='r1_end'
    )

    st.sidebar.markdown('---')
    st.sidebar.subheader('📌 स्थिती (Status) Filter')

    all_statuses = df_raw[col_status].dropna().unique().tolist()
    completed_defaults = [
        'क प्रत',
        'विनाकार्यवाही',
        'प्रस्तावित बिगरशेती/गुंठेवारी मोजणी पूर्ण',
    ]
    default_selected = [s for s in all_statuses if s not in completed_defaults]

    selected_statuses = st.sidebar.multiselect(
        'Select Status:',
        options=all_statuses,
        default=default_selected,
    )

    with st.spinner('डेटा प्रोसेस होत आहे... कृपया वाट पहा...'):
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            '📊 Report 1 (तालुका व दिवसनिहाय)',
            '📋 Report 2 (टप्पा व अधिकारी निहाय)',
            '📈 Executive Dashboard',
            '📋 Report 3 (भूकरमापक निहाय)',
            '📅 Report 4 (मोजणी होणेवर शिल्लक)',
        ])

        from_str = start_date.strftime('%d/%m/%Y')
        to_str = end_date.strftime('%d/%m/%Y')

        # TAB 1: REPORT 1
        with tab1:
            df1 = df_raw[df_raw[col_status].isin(selected_statuses)].copy()
            start_dt = pd.to_datetime(start_date)
            end_dt = pd.to_datetime(end_date)

            df1 = df1[
                (df1['mojni_date_parsed'] >= start_dt)
                & (df1['mojni_date_parsed'] <= end_dt)
            ].copy()

            df1['Pending Days'] = (end_dt - df1['mojni_date_parsed']).dt.days

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

            df1['दिवस वर'] = df1['Pending Days'].apply(day_bucket)
            bucket_order = [
                '15 दिवस वर',
                '30 दिवस वर',
                '60 दिवस वर',
                '90 दिवस वर',
                '90 दिवसांपेक्षा जास्त',
            ]

            pivot_df = pd.crosstab(
                df1['तालुका'],
                df1['दिवस वर'],
                margins=True,
                margins_name='एकूण',
            )
            existing_cols = [c for c in bucket_order if c in pivot_df.columns]
            if 'तारीख उपलब्ध नाही' in pivot_df.columns:
                existing_cols.append('तारीख उपलब्ध नाही')
            if 'एकूण' in pivot_df.columns:
