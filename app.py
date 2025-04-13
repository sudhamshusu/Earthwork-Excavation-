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

st.markdown("### 📁 Download Templates")
col1, col2 = st.columns(2)
with col1:
    with open("Sample.xlsx", "rb") as sample_file:
        st.download_button("📥 Download Input Template (Sample.xlsx)", sample_file, file_name="Sample.xlsx", key="download_sample_main")

with col2:
    with open("Summary.xlsx", "rb") as summary_file_download:
        st.download_button("📥 Download Summary Template (Summary.xlsx)", summary_file_download, file_name="Summary.xlsx", key="download_summary_main")

# 📄 Upload your Filled Templates
data_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])
summary_file = st.file_uploader("Upload Editable Summary Excel (Optional)", type=["xlsx"])

# Debug output of file availability
st.markdown("---")
st.subheader("📋 Or Enter Summary Information Manually")
use_manual_summary = st.checkbox("Use Manual Summary Input")
contract_no_input = st.text_input("Contract Identification No (for Plot Title)")
manual_summary = {}
if use_manual_summary:
    manual_summary["Contract Identification No"] = st.text_input("Contract Identification No")
    manual_summary["Project Name"] = st.text_input("Project Name")
    manual_summary["Client"] = st.text_input("Client")
    manual_summary["Contractor"] = st.text_input("Contractor")
    manual_summary["Agreement Date"] = st.date_input("Agreement Date")
    manual_summary["Completion Date"] = st.date_input("Expected Completion Date")

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

    contract_no = contract_no_input if contract_no_input else ""
    if summary_file:
        summary_df = pd.read_excel(summary_file, header=None, nrows=10)
        for row in summary_df.itertuples(index=False):
            for i, val in enumerate(row):
                if isinstance(val, str) and "contract" in val.lower() and "identification" in val.lower():
                    contract_no = str(row[i + 1]) if i + 1 < len(row) else contract_no
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

    # Plot Cross Sections and Save to PDF
    rows, cols = 2, 2
    figsize = (11.7, 8.3)
    plot_count = 0
    fig, axs = plt.subplots(rows, cols, figsize=figsize)
    fig.suptitle(contract_no, fontsize=10, x=0.5, y=0.98)
    axs = axs.flatten()
    preview_imgs = []

    def plot_chainage_subplot(entry, ax):
        chainage = entry['Chainage']
        fw = entry['Finished Roadway Width']
        fh = entry['Finished Vertical Height']
        ow = entry['Original Roadway Width']
        ac = entry['Area Coefficient']
        angle_deg = entry['Cutting slope']

        try:
            angle_deg = float(angle_deg) if pd.notna(angle_deg) else 75
            if angle_deg == 0:
                raise ValueError("Angle cannot be zero.")
            slope = 1 / np.tan(np.radians(angle_deg))
        except Exception as e:
            st.warning(f"⚠️ Skipping Chainage {chainage}: Invalid Cutting slope value '{angle_deg}'")
            return
        h1_coef = (ac - 0.5) * 2
        area = ac * (fw - ow) * fh

        x1 = -fw / 2
        x3 = fw / 2
        x4 = fw / 2 + slope * fh
        x6 = x1 + ow
        x7 = x6 + h1_coef * fh * slope

        fg_x = [x1, 0, x3, x4]
        fg_y = [0, 0, 0, fh]
        og_x = [x1, x6, x7, x4]
        og_y = [0, 0, h1_coef * fh, fh]

        ax.plot(fg_x, fg_y, color="green", linewidth=2)
        ax.plot(og_x, og_y, color="red", linestyle="--", linewidth=2)
        ax.fill(og_x + fg_x[::-1], og_y + fg_y[::-1], facecolor='gray', alpha=0.3, hatch='//', edgecolor='black')

        ax.text((x3 + x7)/2, (fh + h1_coef * fh)/2, f"Area = {area:.2f} m²", fontsize=8, ha='center')
        ax.text(min(min(fg_x), min(og_x)), max(max(fg_y), max(og_y)) - 0.05, f"■ Finished Roadway (FR): {fw} m", ha='left', fontsize=7, color='green')
        ax.text(min(min(fg_x), min(og_x)), max(max(fg_y), max(og_y)) - 0.2, f"■ Original Roadway (OR): {ow} m", ha='left', fontsize=7, color='red')
        ax.text(min(min(fg_x), min(og_x)), max(max(fg_y), max(og_y)) - 0.35, f"■ Height: {fh} m", ha='left', fontsize=7, color='black')

        ax.axhline(0, color='black', linestyle=':')
        try:
            ch_int = int(float(chainage))
            km = ch_int // 1000
            m = ch_int % 1000
            chainage_str = f"{km}+{m:03d}"
        except:
            chainage_str = str(chainage)
        ax.set_title(f"Chainage {chainage_str}", fontsize=9)
        ax.set_xlabel("Roadway Width", fontsize=6)
        ax.set_ylabel("Height", fontsize=6)
        ax.grid(True, linestyle='--', linewidth=0.3)

    for idx, row in data.iterrows():
        if row.notna().all():
            ax = axs[plot_count % (rows * cols)]
            plot_chainage_subplot(row, ax)

            if plot_count < 40:
                buf = io.BytesIO()
                fig_preview, ax_preview = plt.subplots()
                plot_chainage_subplot(row, ax_preview)
                fig_preview.savefig(buf, format='png')
                plt.close(fig_preview)
                preview_imgs.append(buf.getvalue())

            plot_count += 1

            if plot_count % (rows * cols) == 0:
                pdf.savefig(fig)
                plt.close(fig)
                fig, axs = plt.subplots(rows, cols, figsize=figsize)
                fig.suptitle(contract_no, fontsize=10, x=0.5, y=0.98)
                axs = axs.flatten()

    if plot_count % (rows * cols) != 0:
        for i in range(plot_count % (rows * cols), rows * cols):
            fig.delaxes(axs[i])
        pdf.savefig(fig)
        plt.close(fig)

    pdf.close()

    if preview_imgs:
        st.subheader("🔍 Preview of Plots (up to 40 shown)")
        num_preview = min(40, len(preview_imgs))
        cols = st.columns(4)
        for i in range(num_preview):
            with cols[i % 4]:
                st.image(preview_imgs[i], use_column_width=True)

    st.download_button(
        label="📥 Download PDF with Plots",
        data=pdf_buffer.getvalue(),
        file_name="CrossSection_Plots.pdf",
        mime="application/pdf"
    )

    st.success("✅ Plots generated and ready for download.")
