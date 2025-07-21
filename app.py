if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        sheet = st.selectbox("Select Sheet", sheet_names)
        df = xls.parse(sheet)

        if {'Working as', 'Building No', 'Status'}.issubset(df.columns):
            # Raw data processing (original logic)
            df_cleaned = df[['Working as', 'Building No', 'Status']].copy()
            df_cleaned.dropna(subset=['Working as', 'Building No', 'Status'], inplace=True)

            df_day = df_cleaned[df_cleaned['Status'].str.lower() == 'day present']
            df_night = df_cleaned[df_cleaned['Status'].str.lower() == 'night present']

            pivot_day = pd.pivot_table(df_day, index='Working as', columns='Building No', aggfunc='size', fill_value=0)
            pivot_night = pd.pivot_table(df_night, index='Working as', columns='Building No', aggfunc='size', fill_value=0)

            st.subheader("📊 Day Present Pivot Table")
            st.dataframe(pivot_day, use_container_width=True)

            st.subheader("🌙 Night Present Pivot Table")
            st.dataframe(pivot_night, use_container_width=True)

            # --- Group-wise totals (same logic) ---
            def process_group_building(pivot_df):
                df = pivot_df.copy()
                df.reset_index(inplace=True)
                df['Main Group'] = df['Working as'].map(sub_trade_to_group)
                df = df.dropna(subset=['Main Group'])
                df_grouped = df.groupby('Main Group').sum(numeric_only=True)
                df_grouped.loc['Total'] = df_grouped.sum()
                return df_grouped.reset_index()

            group_building_day = process_group_building(pivot_day)
            group_building_night = process_group_building(pivot_night)

            st.subheader("📊 Group-wise Building Count (Day Present)")
            st.dataframe(group_building_day, use_container_width=True)

            st.subheader("🌙 Group-wise Building Count (Night Present)")
            st.dataframe(group_building_night, use_container_width=True)

        else:
            st.error("❌ The uploaded file does not contain expected columns like 'Working as', 'Building No', 'Status'. Please upload raw attendance data.")
            st.dataframe(df.head(), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
