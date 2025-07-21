import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Daily Manpower Report", layout="wide")

st.title("📊 Daily Manpower Report - FREESIA")

# --------------------- STAGE 1 ---------------------
st.header("Stage 1: Upload Original Attendance Sheet")
uploaded_file_stage1 = st.file_uploader("📥 Upload Daily Attendance Excel File", type=["xlsx"], key="stage1")

if uploaded_file_stage1:
    try:
        xls = pd.ExcelFile(uploaded_file_stage1)
        sheet_names = xls.sheet_names
        sheet = st.selectbox("Select Sheet", sheet_names)
        df = xls.parse(sheet)

        # Clean necessary columns
        df_cleaned = df[['Working as', 'Building No', 'Status']].copy()
        df_cleaned.dropna(subset=['Working as', 'Building No', 'Status'], inplace=True)

        # Split into Day Present and Night Present
        df_day = df_cleaned[df_cleaned['Status'].str.lower() == 'DAY PRESENT']
        df_night = df_cleaned[df_cleaned['Status'].str.lower() == 'SUN NIGHT PRESENT']

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
        output1 = io.BytesIO()
        with pd.ExcelWriter(output1, engine='xlsxwriter') as writer:
            pivot_day.to_excel(writer, sheet_name='Day_Present')
            pivot_night.to_excel(writer, sheet_name='Night_Present')
        output1.seek(0)

        st.download_button(
            label="⬇️ Download Pivot Excel",
            data=output1,
            file_name="Manpower_Report-Day&Night.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("👆 Upload attendance file to generate Day/Night pivot tables.")


# --------------------- STAGE 2 ---------------------
st.markdown("---")
st.header("Stage 2: Upload Pivot File for Group-Wise Summary")
uploaded_file_stage2 = st.file_uploader("📥 Upload Pivot Excel File (from Stage 1)", type=["xlsx"], key="stage2")

if uploaded_file_stage2:
    try:
        # Grouping map
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

        def process_group_building(df):
            df = df.reset_index()
            df['Main Group'] = df['Working as'].map(sub_trade_to_group)
            df = df.dropna(subset=['Main Group'])
            building_cols = df.columns[1:-1]
            df_grouped = df.groupby('Main Group')[building_cols].sum()
            df_grouped.loc['Total'] = df_grouped.sum()
            return df_grouped.reset_index()

        xls = pd.ExcelFile(uploaded_file_stage2)
        df_day = xls.parse("Day_Present")
        df_night = xls.parse("Night_Present")

        group_building_day = process_group_building(df_day)
        group_building_night = process_group_building(df_night)

        st.subheader("📊 Group-wise Building Count (Day Present)")
        st.dataframe(group_building_day, use_container_width=True)

        st.subheader("🌙 Group-wise Building Count (Night Present)")
        st.dataframe(group_building_night, use_container_width=True)

        # Download final summary
        output2 = io.BytesIO()
        with pd.ExcelWriter(output2, engine='xlsxwriter') as writer:
            df_day.to_excel(writer, sheet_name="Raw_Day_Present", index=False)
            df_night.to_excel(writer, sheet_name="Raw_Night_Present", index=False)
            group_building_day.to_excel(writer, sheet_name="Day_Groupwise", index=False)
            group_building_night.to_excel(writer, sheet_name="Night_Groupwise", index=False)
        output2.seek(0)

        st.download_button(
            label="📥 Download Excel Report with Group Totals",
            data=output2,
            file_name="Manpower_Report_Full.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing group-wise summary: {e}")
else:
    st.info("👆 Upload the Excel file generated in Stage 1 to get group-wise building totals.")


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
