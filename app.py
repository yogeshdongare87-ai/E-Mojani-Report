import datetime
import io
import pandas as pd
import streamlit as st
from weasyprint import HTML

# Page Setup
st.set_page_config(page_title="E-Mojani Report Generator", layout="wide")

st.title("📊 तालुका व दिवसनिहाय प्रलंबित मोजणी अहवाल")
st.write(
    "💡 Raw E-Mojani Excel file upload karein. App automatic 'मोजणी तारीख' se din calculate karke report banaye-ga."
)

# 1. File Upload
uploaded_file = st.file_uploader(
    "Upload Raw E-Mojani Excel File (.xlsx)", type=["xlsx"]
)


def process_excel(file):
    # Read Excel (skip header if needed, but standard read first)
    df = pd.read_excel(file)

    # If first row contains column headers like 'अर्ज क्र.(Application No)'
    if 'मोजणी तारीख' not in df.columns:
        # Re-read with header=1 if headers are on row 2
        df = pd.read_excel(file, header=1)

    # Filter out empty Taluka rows
    if 'तालुका' in df.columns:
        df = df[df['तालुका'].notna() & (df['तालुका'] != '')]

    # Convert 'मोजणी तारीख' to datetime
    # Handles Excel serial date numbers (e.g., 46244) or date strings
    def parse_excel_date(val):
        if pd.isna(val):
            return pd.NaT
        try:
            val_num = float(val)
            return datetime.datetime(1899, 12, 30) + datetime.timedelta(
                days=val_num
            )
        except:
            return pd.to_datetime(val, errors='coerce')

    df['mojni_date_parsed'] = df['मोजणी तारीख'].apply(parse_excel_date)

    # Today's date
    today = datetime.datetime.now()

    # Calculate pending days
    df['Pending Days'] = (today - df['mojni_date_parsed']).dt.days

    # Bucket into Day ranges
    def day_bucket(days):
        if pd.isna(days):
            return 'तारीख उपलब्ध नाही'
        if days <= 15:
            return '15 दिवस वर'
        elif days <= 30:
            return '30 दिवस वर'
        elif days <= 90:
            return '90 दिवस वर'
        else:
            return '90 दिवसांपेक्षा जास्त'

    df['दिवस वर'] = df['Pending Days'].apply(day_bucket)

    return df


if uploaded_file is not None:
    with st.spinner('Excel process ho raha hai...'):
        df = process_excel(uploaded_file)

        # Create Pivot Table: Taluka vs Day Buckets
        bucket_order = [
            '15 दिवस वर',
            '30 दिवस वर',
            '90 दिवस वर',
            '90 दिवसांपेक्षा जास्त',
        ]

        pivot_df = pd.crosstab(
            df['तालुका'], df['दिवस वर'], margins=True, margins_name='एकूण'
        )

        # Reorder columns logically
        existing_cols = [c for c in bucket_order if c in pivot_df.columns]
        if 'तारीख उपलब्ध नाही' in pivot_df.columns:
            existing_cols.append('तारीख उपलब्ध नाही')
        existing_cols.append('एकूण')

        pivot_df = pivot_df[existing_cols]

        st.success('✅ Data successfully calculate ho gaya!')

        # Display Table on Screen
        st.subheader("📋 तालुका व दिवसनिहाय प्रलंबित प्रकरणे Table")
        st.dataframe(pivot_df, use_container_width=True)

        # Generate PDF Report
        today_str = datetime.datetime.now().strftime('%d/%m/%Y')

        # Convert pivot to HTML table rows
        rows_html = ""
        for taluka, row in pivot_df.iterrows():
            is_total = taluka == 'एकूण'
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
                @page {{
                    size: A4 portrait;
                    margin: 15mm;
                }}
                body {{
                    font-family: 'Gargi', 'DejaVu Sans', sans-serif;
                    font-size: 12px;
                }}
                h2 {{
                    text-align: center;
                    margin-bottom: 5px;
                    color: #1b4332;
                }}
                .date-header {{
                    text-align: center;
                    font-size: 14px;
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
                    padding: 8px;
                    border: 1px solid #555;
                    font-size: 11px;
                }}
                td {{
                    padding: 6px;
                    border: 1px solid #888;
                    font-size: 11px;
                }}
                tr:nth-child(even) {{
                    background-color: #f9f9f9;
                }}
            </style>
        </head>
        <body>
            <h2>भूमि अभिलेख विभाग - अमरावती</h2>
            <div class="date-header">तालुका व दिवसनिहाय प्रलंबित मोजणी प्रकरणे अहवाल (दिनांक {today_str})</div>
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

        # Generate PDF using WeasyPrint
        pdf_bytes = HTML(string=html_content).write_pdf()

        # PDF Download Button
        st.download_button(
            label="📥 Download PDF Report",
            data=pdf_bytes,
            file_name=f"Mojani_Pending_Daywise_Report_{today_str.replace('/', '-')}.pdf",
            mime="application/pdf",
        )
