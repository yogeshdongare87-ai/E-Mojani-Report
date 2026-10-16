# 🌟 UPDATED PROFESSIONAL & CLEAN PDF / PRINT LAYOUT GENERATOR FOR ALL REPORTS
def generate_report_pdf_html(df, subtitle_text=""):
  headers = list(df.columns)
  current_date_str = datetime.date.today().strftime("%d/%m/%Y")

  # Detect if report contains multi-index (like Taluka + Bhukarmapak in Report 3)
  is_multi_index = isinstance(df.index, pd.MultiIndex)

  html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{
                size: A4 portrait;
                margin: 8mm;
            }}
            body {{
                font-family: 'Helvetica Neue', 'Arial', 'DejaVu Sans', sans-serif;
                margin: 0;
                padding: 0;
                background-color: #ffffff;
                color: #111111;
            }}
            .header-container {{
                text-align: center;
                margin-bottom: 8px;
                padding-bottom: 4px;
                border-bottom: 2px solid #1a365d;
            }}
            .dept-title {{
                font-size: 15px;
                font-weight: bold;
                color: #1a365d;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                margin-bottom: 2px;
            }}
            .report-title {{
                font-size: 12px;
                font-weight: 600;
                color: #2b6cb0;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                border: 1px solid #4a5568;
                margin-top: 4px;
                font-size: 9.5px;
                table-layout: fixed;
            }}
            thead {{
                display: table-header-group; /* Ensures header repeats on every page */
            }}
            tr {{
                page-break-inside: avoid; /* Prevents rows from breaking across pages */
            }}
            th {{
                background-color: #1a365d;
                color: #ffffff;
                font-weight: bold;
                font-size: 9.5px;
                padding: 4px 2px;
                border: 1px solid #2d3748;
                text-align: center;
                word-wrap: break-word;
                line-height: 1.1;
            }}
            td {{
                padding: 3px 2px; /* Reduced row height */
                font-size: 9px;
                border: 1px solid #cbd5e0;
                text-align: center;
                color: #2d3748;
                word-wrap: break-word;
                line-height: 1.1;
            }}
            tbody tr:nth-child(even) {{
                background-color: #f7fafc; /* Alternating zebra rows */
            }}
            .text-col {{
                text-align: left !important;
                padding-left: 4px !important;
                font-weight: 500;
            }}
            .total-row td {{
                font-weight: bold !important;
                background-color: #edf2f7 !important;
                color: #1a365d !important;
                border-top: 1.5px solid #1a365d !important;
                border-bottom: 1.5px solid #1a365d !important;
            }}
            .grand-total-col {{
                background-color: #feebc8 !important; /* Soft highlight for total columns */
                font-weight: bold !important;
                color: #742a2a !important;
            }}
        </style>
    </head>
    <body>
        <div class="header-container">
            <div class="dept-title">भूमि अभिलेख विभाग - अमरावती</div>
            <div class="report-title">{subtitle_text} ( दिनांक {current_date_str} )</div>
        </div>
        <table>
            <thead>
                <tr>
    """

  # Add Table Headers dynamically
  if is_multi_index:
    for idx_name in df.index.names:
      width = "18%" if idx_name == "तालुका" else "26%"
      html += f'<th style="width: {width};">{idx_name}</th>'
  else:
    html += '<th style="width: 22%;">तालुका</th>'

  for h in headers:
    html += f"<th>{h}</th>"

  html += """
                </tr>
            </thead>
            <tbody>
    """

  # Generate Rows
  for idx, row in df.iterrows():
    idx_str = str(idx)
    is_total = (
        "Total" in idx_str or "एकूण" in idx_str or "Grand Total" in idx_str
    )
    row_class = "total-row" if is_total else ""

    html += f'<tr class="{row_class}">'

    if is_multi_index:
      html += f'<td class="text-col">{idx[0]}</td>'
      html += f'<td class="text-col">{idx[1]}</td>'
    else:
      html += f'<td class="text-col">{idx}</td>'

    for col_name in headers:
      val = row[col_name]
      val_str = (
          ""
          if (
              pd.isna(val)
              or val == 0
              or str(val).strip() in ["", "0", "nan", "0.0"]
          )
          else str(int(val) if isinstance(val, (int, float)) else val)
      )

      col_class = (
          "grand-total-col"
          if col_name in ["Grand Total", "एकूण", "Officer Total", "Stage Total"]
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
