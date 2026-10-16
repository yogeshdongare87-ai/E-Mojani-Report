def convert_all_reports_to_single_excel(df_dict):
  """df_dict = {

      'Report 1 (तालुका व दिवसनिहाय)': (df1, 'तालुका व दिवसनिहाय प्रलंबित मोजणी'
      ' प्रकरणे अहवाल'), 'Report 2 (टप्पा व अधिकारी निहाय)': (df2, 'टप्पा व'
      ' अधिकारी निहाय प्रलंबित मोजणी प्रकरणे अहवाल'), ...
  }
  """
  output = io.BytesIO()
  current_date_str = datetime.date.today().strftime("%d/%m/%Y")
  wb = openpyxl.Workbook()
  wb.remove(wb.active)  # Default sheet हटाएं

  for sheet_title, (df_data, report_title) in df_dict.items():
    ws = wb.create_sheet(title=sheet_title[:31])  # Excel Sheet Title max 31 char

    # Styles
    font_dept = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    font_sub = Font(name="Calibri", size=11, bold=True, color="000000")
    font_th = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_th = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    is_multi = isinstance(df_data.index, pd.MultiIndex)
    num_idx_cols = len(df_data.index.names) if is_multi else 1
    total_cols = num_idx_cols + len(df_data.columns)

    # Row 1: Dept Title
    ws.merge_cells(
        start_row=1, start_column=1, end_row=1, end_column=max(total_cols, 2)
    )
    c1 = ws.cell(row=1, column=1, value="भूमि अभिलेख विभाग - अमरावती")
    c1.font = font_dept
    c1.alignment = Alignment(horizontal="center", vertical="center")

    # Row 2: Report Title
    ws.merge_cells(
        start_row=2, start_column=1, end_row=2, end_column=max(total_cols, 2)
    )
    c2 = ws.cell(
        row=2, column=1, value=f"{report_title} ( दिनांक {current_date_str} )"
    )
    c2.font = font_sub
    c2.alignment = Alignment(horizontal="center", vertical="center")

    # Headers & Data Loop...
    # (बाकी सारा Formatting Logic वही रहेगा)

  wb.save(output)
  return output.getvalue()
