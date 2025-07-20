import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Daily Manpower Report", layout="wide")

st.title("📊 Daily Manpower Report- FREESIA")

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

        # Create pivot
        pivot_table = pd.pivot_table(
            df_cleaned,
            index='Working as',
            columns=['Building No', 'Status'],
            aggfunc='size',
            fill_value=0
        )

        st.success("✅ Pivot table generated!")
        st.dataframe(pivot_table, use_container_width=True)

        # Download button
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            pivot_table.to_excel(writer, sheet_name='Manpower Report')
        st.download_button("📥 Download Excel Report", output.getvalue(), "Manpower_Report.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("👆 Please upload your Excel attendance sheet.")
    # --- Footer ---
st.markdown("---")
st.markdown(
    "<div style='text-align: center; font-size: 14px;'>"
    "Developed by: <strong>Viraj Niroshan Gunarathna</strong><br>"
    "This application is maintained under the authority and custody of Mr. Viraj Niroshan Gunarathna."
"</div>"
"</div>"
    "For support or feedback, please contact: Viraj.se@gmail.com or 0586804392"
    "</div>",
    unsafe_allow_html=True
)

