import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Daily Manpower Report", layout="wide")

st.title("📊 Daily Manpower Report - FREESIA")

# Trade to group mapping
sub_trade_to_group = {
    'Asst Electrician': 'ELE',
    'Electrician': 'ELE',
    'Ac Tech': 'HVAC',
    'Ac-Pipe-Fitter': 'HVAC',
    'Asst Ductman': 'HVAC',
    'Ductman': 'HVAC',
    'Chw-Pipe-Fitter': 'HVAC',
    'Gi Duct Fabricator': 'HVAC',
    'Insulator': 'HVAC',
    'Welder': 'Welder',
    'Asst Plumber': 'PLU',
    'Plumber': 'PLU',
    'Fire Alarm-Helper': 'FA',
    'Fire Alarm & Emergency Technician': 'FA',
    'Fire Alarm Technician': 'FA',
    'Fire Fighting Technician-Helper': 'FF',
    'Fire Fighting - Pipe Fitter': 'FF',
    'Fire Fighting Technicans': 'FF',
    'Fire Sealant Technician': 'F/S',
    'Elv Technician': 'ELV',
    'Lpg Technician-Pipe Fitter': 'LPG Technician/Pipe Fitter',
    'Lpg  Helper': 'LPG Helper',
    'Welder-Cs-Lpg-Technician': 'LPG Welder'
}

# --------------------- STAGE 1 ---------------------
st.header("Stage 1: Upload Original Attendance Sheet")
uploaded_file_stage1 = st.file_uploader("📥 Upload Daily Attendance Excel File", type=["xlsx"], key="stage1")

if uploaded_file_stage1:
    try:
        # Read Excel file
        xls = pd.ExcelFile(uploaded_file_stage1)
        sheet_names = xls.sheet_names
        
        if len(sheet_names) > 1:
            sheet = st.selectbox("Select Sheet", sheet_names)
        else:
            sheet = sheet_names[0]
            
        df = xls.parse(sheet)
        
        st.write(f"📋 Data loaded: {len(df)} rows")
        
        # Show column names for debugging
        st.write("🔍 Available columns:", list(df.columns))
        
        # Check required columns exist
        required_cols = ['Working as', 'Building No', 'Status']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ Missing required columns: {missing_cols}")
            st.write("Available columns:", list(df.columns))
        else:
            # Clean data - remove rows with missing values in key columns
            df_cleaned = df[required_cols].copy()
            df_cleaned = df_cleaned.dropna()
            
            # Convert Status to string and clean
            df_cleaned['Status'] = df_cleaned['Status'].astype(str).str.strip()
            
            st.write(f"✅ Clean data: {len(df_cleaned)} rows")
            
            # Show unique status values for debugging
            unique_statuses = df_cleaned['Status'].unique()
            st.write("🔍 Unique Status values:", unique_statuses)
            
            # Filter day and night data - more flexible matching
            df_day = df_cleaned[df_cleaned['Status'].str.contains('DAY PRESENT', case=False, na=False)]
            df_night = df_cleaned[df_cleaned['Status'].str.contains('NIGHT PRESENT', case=False, na=False)]
            
            st.write(f"☀️ Day Present: {len(df_day)} records")
            st.write(f"🌙 Night Present: {len(df_night)} records")
            
            if len(df_day) == 0 and len(df_night) == 0:
                st.warning("⚠️ No records found with 'DAY PRESENT' or 'NIGHT PRESENT' status")
                st.write("Status values found:", unique_statuses)
            else:
                # Create pivot tables
                def create_pivot(df_data, title):
                    if len(df_data) == 0:
                        st.write(f"No data for {title}")
                        return pd.DataFrame()
                    
                    pivot = pd.pivot_table(
                        df_data,
                        index='Working as',
                        columns='Building No',
                        aggfunc='size',
                        fill_value=0
                    )
                    
                    # Add totals
                    pivot.loc['Total'] = pivot.sum()
                    pivot['Total'] = pivot.sum(axis=1)
                    
                    return pivot
                
                pivot_day = create_pivot(df_day, "Day Present")
                pivot_night = create_pivot(df_night, "Night Present")
                
                # Display pivot tables
                if not pivot_day.empty:
                    st.subheader("📊 Day Present Pivot Table")
                    st.dataframe(pivot_day, use_container_width=True)
                
                if not pivot_night.empty:
                    st.subheader("🌙 Night Present Pivot Table")
                    st.dataframe(pivot_night, use_container_width=True)
                
                # Store data in session state for stage 2
                st.session_state['pivot_day'] = pivot_day
                st.session_state['pivot_night'] = pivot_night
                st.session_state['df_day'] = df_day
                st.session_state['df_night'] = df_night
                
                # Create downloadable Excel
                output1 = io.BytesIO()
                with pd.ExcelWriter(output1, engine='xlsxwriter') as writer:
                    if not pivot_day.empty:
                        pivot_day.to_excel(writer, sheet_name='Day_Present')
                    if not pivot_night.empty:
                        pivot_night.to_excel(writer, sheet_name='Night_Present')
                output1.seek(0)
                
                st.download_button(
                    label="⬇️ Download Stage 1 Pivot Excel",
                    data=output1,
                    file_name="Manpower_Report-Day&Night.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
        st.write("Please check your Excel file format and column names")
else:
    st.info("👆 Upload attendance file to generate Day/Night pivot tables.")

# --------------------- STAGE 2 ---------------------
st.markdown("---")
st.header("Stage 2: Group-wise Summary")

if 'pivot_day' in st.session_state or 'pivot_night' in st.session_state:
    try:
        def process_group_summary(pivot_df, df_original):
            if pivot_df.empty or len(df_original) == 0:
                return pd.DataFrame()
            
            # Remove total row and column if exists
            pivot_clean = pivot_df.copy()
            if 'Total' in pivot_clean.index:
                pivot_clean = pivot_clean.drop('Total')
            if 'Total' in pivot_clean.columns:
                pivot_clean = pivot_clean.drop('Total', axis=1)
            
            # Get building columns (these will become our rows in final table)
            building_columns = list(pivot_clean.columns)
            
            # Get all unique groups (these will become our columns in final table)
            all_groups = set(sub_trade_to_group.values())
            
            # Initialize building data - Buildings as rows, Groups as columns
            building_data = {}
            
            # Initialize all buildings with zero values for each group
            for building in building_columns:
                building_data[building] = {group: 0 for group in all_groups}
                # Also initialize for unmapped trades
                for trade in pivot_clean.index:
                    if trade not in sub_trade_to_group:
                        building_data[building][trade] = 0
            
            # Process each trade (row in original pivot)
            for trade in pivot_clean.index:
                if trade in sub_trade_to_group:
                    # Map trade to group
                    group = sub_trade_to_group[trade]
                    # Add values for each building
                    for building in building_columns:
                        if building in pivot_clean.columns:
                            building_data[building][group] += pivot_clean.loc[trade, building]
                else:
                    # Handle unmapped trades - treat as separate column
                    for building in building_columns:
                        if building in pivot_clean.columns:
                            building_data[building][trade] = pivot_clean.loc[trade, building]
            
            # Convert to DataFrame - Buildings as index (rows), Groups as columns
            group_df = pd.DataFrame.from_dict(building_data, orient='index')
            
            # Add totals if DataFrame is not empty
            if not group_df.empty:
                # Add total column (sum of each row - total workers per building)
                group_df['Total'] = group_df.sum(axis=1)
                # Add total row (sum of each column - total workers per group across all buildings)
                group_df.loc['Total'] = group_df.sum()
            
            return group_df
        
        # Process both day and night
        group_day = pd.DataFrame()
        group_night = pd.DataFrame()
        
        if 'pivot_day' in st.session_state:
            group_day = process_group_summary(st.session_state['pivot_day'], st.session_state['df_day'])
        
        if 'pivot_night' in st.session_state:
            group_night = process_group_summary(st.session_state['pivot_night'], st.session_state['df_night'])
        
        # Display group summaries
        if not group_day.empty:
            st.subheader("📊 Group-wise Building Count (Day Present)")
            st.dataframe(group_day, use_container_width=True)
        
        if not group_night.empty:
            st.subheader("🌙 Group-wise Building Count (Night Present)")
            st.dataframe(group_night, use_container_width=True)
        
        # Create final Excel report
        if not group_day.empty or not group_night.empty:
            output2 = io.BytesIO()
            with pd.ExcelWriter(output2, engine='xlsxwriter') as writer:
                # Original pivot tables
                if 'pivot_day' in st.session_state and not st.session_state['pivot_day'].empty:
                    st.session_state['pivot_day'].to_excel(writer, sheet_name="Day_Present")
                if 'pivot_night' in st.session_state and not st.session_state['pivot_night'].empty:
                    st.session_state['pivot_night'].to_excel(writer, sheet_name="Night_Present")
                
                # Group summaries
                if not group_day.empty:
                    group_day.to_excel(writer, sheet_name="Day_Groupwise")
                if not group_night.empty:
                    group_night.to_excel(writer, sheet_name="Night_Groupwise")
            
            output2.seek(0)
            
            st.download_button(
                label="📥 Download Complete Excel Report with Groups",
                data=output2,
                file_name="Manpower_Report_Complete.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
    except Exception as e:
        st.error(f"❌ Error creating group summary: {e}")
        st.write("Debug info:")
        if 'pivot_day' in st.session_state:
            st.write("Day pivot shape:", st.session_state['pivot_day'].shape)
        if 'pivot_night' in st.session_state:
            st.write("Night pivot shape:", st.session_state['pivot_night'].shape)
else:
    st.info("👆 Complete Stage 1 first to enable group-wise processing.")

# --------------------- Footer ---------------------
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; font-size: 14px;'>
        Developed by: <strong>Viraj Niroshan Gunarathna</strong><br>
        This application is maintained under the authority and custody of Mr. Viraj Niroshan Gunarathna.
    </div>
    <div style='text-align: center; font-size: 12px; margin-top: 10px;'>
        For support or feedback, contact: <a href='mailto:Viraj.se@gmail.com'>Viraj.se@gmail.com</a> | 📞 0586804392
    </div>
    <div style='text-align: center; font-size: 11px; margin-top: 10px; color: gray;'>
        &copy; 2025 Viraj Niroshan Gunarathna. All rights reserved.
    </div>
    """,
    unsafe_allow_html=True
)