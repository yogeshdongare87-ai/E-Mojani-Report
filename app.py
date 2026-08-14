import pandas as pd
import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from weasyprint import HTML

st.set_page_config(page_title="Mojni Pending Report Generator", layout="wide")

st.title("📊 तालुका व प्रलंबित दिवसानिहाय मोजणी अहवाल Tool")
st.write("Aapni Excel file upload karein aur instant PDF report download karein.")

# 1. Excel File Upload
uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

def load_excel_raw(file):
    # Parse Excel file smoothly without version issue
    with zipfile.ZipFile(file, 'r') as z:
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ss_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for t in ss_tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                shared_strings.append(t.text if t.text else '')

        sheet_tree = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        rows_data = []
        sheetData = sheet_tree.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData')
        for row in sheetData.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
            row_vals = {}
            for c in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                cell_ref = c.attrib.get('r')
                cell_type = c.attrib.get('t', '')
                v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                val = v.text if v is not None else ''
                if cell_type == 's' and val != '':
                    val = shared_strings[int(val)] if int(val) < len(shared_strings) else val
                col_letter = ''.join([char for char in cell_ref if char.isalpha()])
                row_vals[col_letter] = val
            rows_data.append(row_vals)

    df_raw = pd.DataFrame(rows_data)
    
    # Sort columns by Excel letters
    def col_key(col_str):
        num = 0
        for char in col_str:
            num = num * 26 + (ord(char) - ord('A') + 1)
        return num

    sorted_cols = sorted(df_raw.columns, key=col_key)
    df_raw = df_raw[sorted_cols]
    
    df_clean = df_raw[1:].copy()
    df_clean.columns = df_raw.iloc[0].values
    return df_clean

if uploaded_file is not None:
    try:
        df = load_excel_raw(uploaded_file)
        
        # Detect Taluka and Divas War Columns
        taluka_col = [c for c in df.columns if 'तालुका' in str(c)][0]
        divas_col = [c for c in df.columns if 'दिवस' in str(c)][0]
        
        # 2. Pivot Table Generation (Exact like user screenshot)
        pivot = pd.crosstab(df[taluka_col], df[divas_col], margins=True, margins_name='Grand Total')
        
        st.subheader("📋 Generated Summary Pivot Table")
        st.dataframe(pivot, use_container_width=True)
        
        # 3. Create PDF Button
        if st.button("📄 Generate PDF Report"):
            table_html = """
            <table style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px;">
                <thead>
                    <tr style="background-color: #b8cce4; text-align: center; font-weight: bold;">
                        <th style="border: 1px solid #000; padding: 6px;">तालुका</th>
                        <th style="border: 1px solid #000; padding: 6px;">15 दिवस वर</th>
                        <th style="border: 1px solid #000; padding: 6px;">30 दिवस वर</th>
                        <th style="border: 1px solid #000; padding: 6px;">60 दिवस वर</th>
                        <th style="border: 1px solid #000; padding: 6px;">90 दिवस वर</th>
                        <th style="border: 1px solid #000; padding: 6px;">90 दिवसा जास्त</th>
                        <th style="border: 1px solid #000; padding: 6px;">Grand Total</th>
                    </tr>
                </thead>
                <tbody>
            """
            
            for idx, row in pivot.iterrows():
                is_grand = (idx == 'Grand Total')
                bg = "background-color: #dce6f1; font-weight: bold;" if is_grand else ""
                table_html += f"""
                    <tr style="{bg}">
                        <td style="border: 1px solid #000; padding: 5px;">{idx if idx!='' else '(blank)'}</td>
                        <td style="border: 1px solid #000; padding: 5px; text-align: right;">{row.get('15 दिवस वर', 0)}</td>
                        <td style="border: 1px solid #000; padding: 5px; text-align: right;">{row.get('30 दिवस वर', 0)}</td>
                        <td style="border: 1px solid #000; padding: 5px; text-align: right;">{row.get('60 दिवस वर', 0)}</td>
                        <td style="border: 1px solid #000; padding: 5px; text-align: right;">{row.get('90 दिवस वर', 0)}</td>
                        <td style="border: 1px solid #000; padding: 5px; text-align: right;">{row.get('90 दिवसा जास्त', 0)}</td>
                        <td style="border: 1px solid #000; padding: 5px; text-align: right; font-weight: bold;">{row.get('Grand Total', 0)}</td>
                    </tr>
                """
            
            table_html += "</tbody></table>"
            
            pdf_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
            <meta charset="utf-8">
            <style>
                @page {{ size: A4 landscape; margin: 15mm; }}
                body {{ font-family: 'Arial', sans-serif; }}
                .header {{ text-align: center; margin-bottom: 20px; }}
                .header h2 {{ margin: 0; color: #1e3a8a; font-size: 18px; }}
            </style>
            </head>
            <body>
                <div class="header">
                    <h2>तालुका व प्रलंबित दिवसानिहाय मोजणी अहवाल (Mojni Pending Report)</h2>
                </div>
                {table_html}
            </body>
            </html>
            """
            
            pdf_data = HTML(string=pdf_html).write_pdf()
            
            st.download_button(
                label="📥 Download PDF Report",
                data=pdf_data,
                file_name="Taluka_Pending_Days_Report.pdf",
                mime="application/pdf"
            )
            st.success("PDF Report Taiyar hai!")
            
    except Exception as e:
        st.error(f"File process karne me error aaya: {e}")