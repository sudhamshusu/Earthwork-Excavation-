import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import io
import os

st.set_page_config(layout="wide")
st.title("📊 Earthwork Cross-Section Plot Generator")

# 📘 Instructions for Public Users
st.markdown("""
### 📌 How to Use This Tool:
1. **Download the Templates** using the buttons below.
2. **Fill out the Input Template (`sample.xlsx`)** with chainage, widths, height, slope, etc.
3. **Fill out the Summary Template (`summary.xlsx`)** with project-specific details.
4. **Upload both filled files** using the file uploaders.
5. Wait for the app to generate:
   - Project summary 📋
   - Cross-section plots 📊
   - Volume summary 📦
6. **Preview plots** and **Download final reports** as PDF.

💡 This tool helps automate earthwork quantity and visualization reporting for roads, embankments, and cut/fill projects.
""")

# 📄 Upload your Filled Templates
data_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])
summary_file = st.file_uploader("Upload Editable Summary Excel (Optional)", type=["xlsx"])

# Debug output of file availability
st.markdown("---")
st.subheader("📋 Or Enter Summary Information Manually")
use_manual_summary = st.checkbox("Use Manual Summary Input")
manual_summary = {}
if use_manual_summary:
    manual_summary["Contract Identification No"] = st.text_input("Contract Identification No")
    manual_summary["Project Name"] = st.text_input("Project Name")
    manual_summary["Client"] = st.text_input("Client")
    manual_summary["Contractor"] = st.text_input("Contractor")
    manual_summary["Agreement Date"] = st.date_input("Agreement Date")
    manual_summary["Completion Date"] = st.date_input("Expected Completion Date")

st.markdown("### 📁 Download Templates")
col1, col2 = st.columns(2)
with col1:
    with open("Sample.xlsx", "rb") as sample_file:
        st.download_button("📥 Download Input Template (Sample.xlsx)", sample_file, file_name="Sample.xlsx")

with col2:
    with open("Summary.xlsx", "rb") as summary_file_download:
        st.download_button("📥 Download Summary Template (Summary.xlsx)", summary_file_download, file_name="Summary.xlsx")

if data_file:
    raw = pd.read_excel(data_file, header=None, nrows=50)
    header_row_index = raw[raw.apply(lambda row: row.astype(str).str.contains("Chainage", case=False).any(), axis=1)].index[0]
    data = pd.read_excel(data_file, skiprows=header_row_index)
    data.dropna(subset=[data.columns[0]], inplace=True)

    expected_columns = [
        "S.No", "Chainage", "Finished Roadway Width", "Finished Vertical Height",
        "Original Roadway Width", "Area Coefficient", "Cutting slope"
    ]
    data = data.iloc[:, :7]
    data.columns = expected_columns

    data['Area (m²)'] = data.apply(lambda row: row['Area Coefficient'] * (row['Finished Roadway Width'] - row['Original Roadway Width']) * row['Finished Vertical Height'], axis=1)
    data['Volume (m³)'] = 0.0
    chainage_values = data['Chainage'].astype(str).str.replace("+", "").astype(float).values
    for i in range(1, len(data)):
        delta_ch = chainage_values[i] - chainage_values[i - 1]
        data.at[i, 'Volume (m³)'] = delta_ch * data.at[i, 'Area (m²)']
    total_volume = data['Volume (m³)'].sum()
    data.loc[len(data.index)] = ['Total', '', '', '', '', '', '', '', total_volume]

    contract_no = "Contract Identification No: "
    if summary_file:
        summary_df = pd.read_excel(summary_file, header=None, nrows=10)
        for row in summary_df.itertuples(index=False):
            for i, val in enumerate(row):
                if isinstance(val, str) and "Contract Identification No" in val:
                    contract_no += str(row[i + 1]) if i + 1 < len(row) else ""
                    break

    # Prepare PDF buffer
    pdf_buffer = io.BytesIO()
    pdf = PdfPages(pdf_buffer)

    # Generate Project Summary PDF
    if summary_file or use_manual_summary:
        fig_summary, ax_summary = plt.subplots(figsize=(11.7, 8.3))
        ax_summary.axis('off')

        if summary_file:
            table_data = summary_df.dropna(how='all').dropna(axis=1, how='all')
            cell_text = table_data.astype(str).values.tolist()
        elif use_manual_summary:
            cell_text = [[k, str(v)] for k, v in manual_summary.items()]

        table = ax_summary.table(cellText=cell_text, colLabels=None, loc='center', cellLoc='left')
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.scale(1, 1.5)
        fig_summary.suptitle("Project Summary", fontsize=12)

        pdf.savefig(fig_summary)
        plt.close(fig_summary)

    st.success("✅ Project summary PDF generated and ready to continue with plotting and volume summary.")
