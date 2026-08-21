# ---------------------------------------------------------------------
        # TAB 3: STYLISH DASHBOARD WITH CHARTS & COLORFUL CARDS (UPDATED LOGIC)
        # ---------------------------------------------------------------------
        with tab3:
            st.markdown("## 🎨 Executive Visual Dashboard")

            # Dates for Previous Month
            today = datetime.date.today()
            first_day_of_curr_month = today.replace(day=1)
            last_day_prev_month = first_day_of_curr_month - datetime.timedelta(days=1)
            first_day_prev_month = last_day_prev_month.replace(day=1)

            st.info(f"📅 **विश्लेषण कालावधी (मागील महिना):** {first_day_prev_month.strftime('%d/%m/%Y')} ते {last_day_prev_month.strftime('%d/%m/%Y')}")

            p_start_dt = pd.to_datetime(first_day_prev_month)
            p_end_dt = pd.to_datetime(last_day_prev_month)

            # 🔥 FILTER ONLY 'क प्रत' CASES FOR PREVIOUS MONTH FIRST
            df_prev_kprat = df_raw[
                (df_raw["mojni_date_parsed"] >= p_start_dt)
                & (df_raw["mojni_date_parsed"] <= p_end_dt)
                & (df_raw["स्थिती"] == "क प्रत")
            ].copy()

            # --- COLORFUL KPI CARDS ---
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #2a9d8f;">
                        <div class="metric-title">🏆 मागील महिन्यात 'क प्रत' निकाल</div>
                        <div class="metric-value">{len(df_prev_kprat)}</div>
                    </div>
                """, unsafe_allow_html=True)
            with c2:
                st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #e76f51;">
                        <div class="metric-title">⏳ एकूण प्रलंबित प्रकरणे</div>
                        <div class="metric-value">{len(df_raw[~df_raw["स्थिती"].isin(completed_defaults)])}</div>
                    </div>
                """, unsafe_allow_html=True)
            with c3:
                st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #f4a261;">
                        <div class="metric-title">📌 आजच्या तारखेनंतरची शिल्लक प्रकरणे</div>
                        <div class="metric-value">{len(df_raw[(df_raw["mojni_date_parsed"] > today_dt) & (df_raw["स्थिती"] == "मोजणीची माहिती")])}</div>
                    </div>
                """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # --- CHARTS SECTION ---
            col_chart1, col_chart2 = st.columns(2)

            with col_chart1:
                st.markdown("### 🥇 Top 3 तालुका ('क प्रत')")
                if not df_prev_kprat.empty:
                    top_t = df_prev_kprat["तालुका"].value_counts().head(3).reset_index()
                    top_t.columns = ["तालुका", "संख्या"]

                    fig_taluka = px.bar(
                        top_t,
                        x="तालुका",
                        y="संख्या",
                        text="संख्या",
                        color="तालुका",
                        color_discrete_sequence=px.colors.qualitative.Set2,
                        title="Top 3 तालुका - 'क प्रत' कामगिरी"
                    )
                    fig_taluka.update_traces(textposition="outside")
                    fig_taluka.update_layout(showlegend=False, height=350)
                    st.plotly_chart(fig_taluka, use_container_width=True)
                else:
                    st.warning("मागील महिन्यात 'क प्रत' चा डेटा उपलब्ध नाही.")

            with col_chart2:
                st.markdown("### 👷 Top भूकरमापक / मोजणीदार ('क प्रत')")
                
                # Check column names in Excel for Surveyor Name
                surveyor_col = None
                possible_surveyor_cols = [
                    "कर्मचारी/अधिकारी चे नाव", 
                    "भूकरमापक", 
                    "भूकरमापकाचे नाव", 
                    "मोजणीदार", 
                    "कर्मचारी नाव", 
                    "अधिकारी नाव"
                ]
                
                for c in possible_surveyor_cols:
                    if c in df_raw.columns:
                        surveyor_col = c
                        break

                if surveyor_col and not df_prev_kprat.empty:
                    # Clean empty/null values in surveyor names
                    df_surveyor_filtered = df_prev_kprat[
                        df_prev_kprat[surveyor_col].notna() 
                        & (df_prev_kprat[surveyor_col] != "")
                    ]
                    
                    top_s = df_surveyor_filtered[surveyor_col].value_counts().head(5).reset_index()
                    top_s.columns = ["भूकरमापक नाव", "संख्या"]

                    if not top_s.empty:
                        fig_surv = px.bar(
                            top_s,
                            x="संख्या",
                            y="भूकरमापक नाव",
                            orientation="h",
                            text="संख्या",
                            color="भूकरमापक नाव",
                            color_discrete_sequence=px.colors.qualitative.Pastel,
                            title="सर्वात जास्त 'क प्रत' काढणारे भूकरमापक"
                        )
                        fig_surv.update_traces(textposition="outside")
                        fig_surv.update_layout(showlegend=False, height=350, yaxis=dict(autorange="reversed"))
                        st.plotly_chart(fig_surv, use_container_width=True)
                    else:
                        st.warning("'क प्रत' वाले केसेस मध्ये भूकरमापकांचे नाव सापडले नाही.")
                elif not surveyor_col:
                    st.info("💡 Excel मध्ये 'भूकरमापक' किंवा 'कर्मचारी/अधिकारी चे नाव' चा कॉलम सापडला नाही.")
                else:
                    st.warning("मागील महिन्यात 'क प्रत' चा डेटा उपलब्ध नाही.")
