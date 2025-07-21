import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Daily Manpower Report", layout="wide")

st.title("📊 Daily Manpower Report - FREESIA")

uploaded_file = st.file_uploader("Upload Daily Attendance Excel File", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        sheet = st.selectbox("Select Sheet", sheet_names)
        df = xls.parse(sheet)

        # Clean necessary columns
        df_cleaned = df[['Working as', 'Building No', 'Status']].copy()
        df_cleaned.dropna(subset=['Working as', 'Building No', 'Status'], inplace=True)

        # Split into Day Present and Night Present
        df_day = df_cleaned[df_cleaned['Status'].str.lower() == 'day present']
        df_night = df_cleaned[df_cleaned['Status'].str.lower() == 'night present']

        # Create pivot tables
        pivot_day = pd.pivot_table(
            df_day,
            index='Working as',
            columns='Building No',
            aggfunc='size',
            fill_value=0
        )

        pivot_night = pd.pivot_table(
            df_night,
            index='Working as',
            columns='Building No',
            aggfunc='size',
            fill_value=0
        )

        # Display pivot tables
        st.subheader("📊 Day Present Pivot Table")
        st.dataframe(pivot_day, use_container_width=True)

        st.subheader("🌙 Night Present Pivot Table")
        st.dataframe(pivot_night, use_container_width=True)

        # Create Excel file with both tables
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            pivot_day.to_excel(writer, sheet_name='Day_Present')
            pivot_night.to_excel(writer, sheet_name='Night_Present')
        output.seek(0)

        # Download button
        st.download_button(
            label="📥 Download Excel Report",
            data=output,
            file_name="Manpower_Report-Day&Night.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("👆 Please upload your Excel attendance sheet.")



    # Sub trade to group mapping
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

# File uploader
uploaded_file = st.file_uploader("Upload Daily Attendance Excel File", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_day = xls.parse("Day_Present")
        df_night = xls.parse("Night_Present")

        def process_group_building(df):
            df = df.copy()
            df['Main Group'] = df['Working as'].map(sub_trade_to_group)
            df = df.dropna(subset=['Main Group'])  # Remove rows not in the mapping
            building_cols = df.columns[1:-1] if 'Main Group' in df.columns else df.columns[1:]
            df_grouped = df.groupby('Main Group')[building_cols].sum()
            df_grouped.loc['Total'] = df_grouped.sum()
            return df_grouped.reset_index()

        # Process Day and Night data
        group_building_day = process_group_building(df_day)
        group_building_night = process_group_building(df_night)

        # Show in app
        st.subheader("📊 Group-wise Building Count (Day Present)")
        st.dataframe(group_building_day, use_container_width=True)

        st.subheader("🌙 Group-wise Building Count (Night Present)")
        st.dataframe(group_building_night, use_container_width=True)

        # Create download Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_day.to_excel(writer, sheet_name="Raw_Day_Present", index=False)
            df_night.to_excel(writer, sheet_name="Raw_Night_Present", index=False)
            group_building_day.to_excel(writer, sheet_name="Day_Groupwise", index=False)
            group_building_night.to_excel(writer, sheet_name="Night_Groupwise", index=False)
        output.seek(0)

        # Download button
        st.download_button(
            label="📥 Download Excel Report with Group Totals",
            data=output,
            file_name="Manpower_Report_Full.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")

else:
    st.info("👆 Please upload your Excel attendance sheet.")


# --- Footer ---
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; font-size: 14px;'>
        Developed by: <strong>Viraj Niroshan Gunarathna</strong><br>
        This application is maintained under the authority and custody of Mr. Viraj Niroshan Gunarathna.
    </div>
    <div style='text-align: center; font-size: 12px; margin-top: 10px;'>
        For support or feedback, please contact: <a href='mailto:Viraj.se@gmail.com'>Viraj.se@gmail.com</a> | 📞 0586804392
    </div>
    <div style='text-align: center; font-size: 11px; margin-top: 10px; color: gray;'>
        &copy; 2025 Viraj Niroshan Gunarathna. All rights reserved.
    </div>
    """,
    unsafe_allow_html=True
)
