# ---------------------------------------------------------------------
        # TAB 6: REPORT 5 (मागील ३ महिने 'क प्रत' प्रोग्रेस अहवाल)
        # ---------------------------------------------------------------------
        with tab6:
            st.markdown(
                "## 🚀 रिपोर्ट ५: मागील ३ महिन्यांचा 'क प्रत' प्रोग्रेस अहवाल"
                " (आजपर्यंत)"
            )

            today_ref_r5 = pd.to_datetime(datetime.date.today())
            three_months_ago_r5 = today_ref_r5 - pd.Timedelta(days=90)

            # Correct Filter: Only valid 'क प्रत निकाली दिनाक' in last 3 months
            df_r5 = df_raw[
                (df_raw["k_prat_date_parsed"].notna())
                & (df_raw["k_prat_date_parsed"] >= three_months_ago_r5)
                & (df_raw["k_prat_date_parsed"] <= today_ref_r5)
            ].copy()

            st.success(
                f"🎯 **मागील ३ महिन्यात पूर्ण झालेली एकूण 'क प्रत' प्रकरणे:"
                f" {len(df_r5)}**"
            )

            if df_r5.empty:
                st.warning(
                    "मागील ३ महिन्यात 'क प्रत' दिनांक असलेले प्रकरणे आढळली"
                    " नाहीत."
                )
            else:
                df_r5["Month_Year"] = df_r5["k_prat_date_parsed"].dt.strftime(
                    "%B"
                )
                df_r5["Year_Month_Sort"] = df_r5[
                    "k_prat_date_parsed"
                ].dt.strftime("%Y-%m")

                sorted_months_r5 = (
                    df_r5[["Year_Month_Sort", "Month_Year"]]
                    .drop_duplicates()
                    .sort_values("Year_Month_Sort")
                )
                sorted_month_names_r5 = sorted_months_r5["Month_Year"].tolist()

                pivot_r5 = pd.pivot_table(
                    df_r5,
                    index="तालुका",
                    columns="Month_Year",
                    values=df_r5.columns[0],
                    aggfunc="count",
                    fill_value=0,
                )

                cols_in_pivot_r5 = [
                    m for m in sorted_month_names_r5 if m in pivot_r5.columns
                ]
                pivot_r5 = pivot_r5[cols_in_pivot_r5]

                standard_talukas = [
                    "अंजनगाव",
                    "अचलपूर",
                    "अमरावती",
                    "चांदुर बा.",
                    "चांदुर रे.",
                    "चिखलदरा",
                    "तिवसा",
                    "दर्यापूर",
                    "धामणगाव",
                    "धारणी",
                    "नांदगाव",
                    "भातकुली",
                    "मोर्शी",
                    "वरुड",
                ]

                pivot_r5 = pivot_r5.rename(
                    index={"अचलपुर": "अचलपूर", "दर्यापुर": "दर्यापूर"}
                )

                present_talukas = pivot_r5.index.tolist()
                for t in present_talukas:
                    if t not in standard_talukas:
                        standard_talukas.append(t)

                pivot_r5 = pivot_r5.reindex(standard_talukas, fill_value=0)
                pivot_r5["Grand Total"] = pivot_r5.sum(axis=1)

                total_row_r5 = pivot_r5.sum(axis=0)
                pivot_r5.loc["Total"] = total_row_r5

                display_r5_df = pivot_r5.astype(object)
                for col in display_r5_df.columns:
                    display_r5_df[col] = display_r5_df[col].apply(
                        lambda x: "" if x == 0 or pd.isna(x) else int(x)
                    )

                st.subheader(
                    "📋 मागील ३ महिने 'क प्रत' झालेली प्रकरणे (तालुका व महिना"
                    " निहाय)"
                )
                st.dataframe(display_r5_df, use_container_width=True)

                btn_r5_1, btn_r5_2 = st.columns(2)
                with btn_r5_1:
                    excel_r5 = convert_df_to_excel(
                        pivot_r5, sheet_name="Report_5"
                    )
                    st.download_button(
                        label="📥 Download Report 5 (Excel)",
                        data=excel_r5,
                        file_name="Report_5_K_Prat_Progress.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ),
                    )
                with btn_r5_2:
                    pdf_r5 = generate_report_pdf_html(
                        pivot_r5,
                        title="मागील ३ महिने 'क प्रत' प्रोग्रेस अहवाल",
                    )
                    st.download_button(
                        label="📄 Download Report 5 (PDF / Print Layout)",
                        data=pdf_r5,
                        file_name=(
                            "Report_5_K_Prat_Progress.pdf"
                            if WEASYPRINT_AVAILABLE
                            else "Report_5_K_Prat_Progress.html"
                        ),
                        mime=(
                            "application/pdf"
                            if WEASYPRINT_AVAILABLE
                            else "text/html"
                        ),
                    )
