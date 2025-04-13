import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import io

st.set_page_config(layout="wide")
st.title("📐 Earthwork Cross-Section Plot Generator")

# 📘 Instructions
st.markdown("""
### 📌 How to Use This Tool:
1. **Download the Templates** using the buttons below.
2. **Fill out the Input Template (`Sample.xlsx`)** with chainage, widths, height, slope, etc.
3. **Optionally fill the Summary Template (`Summary.xlsx`)** with project details.
4. **Upload your filled files below**.
5. Wait for the app to generate:
   - Cross-section plots 📊
   - Volume summary 📦
6. **Preview plots** and **Download final PDF reports**.

💡 Use this for earthwork estimation, excavation, and road section visualization.
""")

# 📂 Template Download
st.markdown("### 📁 Download Templates")
col1, col2 = st.columns(2)
with col1:
    with open("Sample.xlsx", "rb") as sample_file:
        st.download_button("📥 Download Sample.xlsx", sample_file, file_name="Sample.xlsx")
with col2:
    with open("Summary.xlsx", "rb") as summary_file:
        st.download_button("📥 Download Summary.xlsx", summary_file, file_name="Summary.xlsx")

# 📤 Upload Input
st.markdown("---")
data_file = st.file_uploader("Upload Filled Sample.xlsx", type=["xlsx"])

# Initialize preview storage and plotting logic
all_imgs = []
pdf_buffer = io.BytesIO()
pdf = PdfPages(pdf_buffer)

if data_file:
    raw = pd.read_excel(data_file, header=None, nrows=50)
    header_row_index = raw[raw.apply(lambda row: row.astype(str).str.contains("Chainage", case=False).any(), axis=1)].index[0]
    data = pd.read_excel(data_file, skiprows=header_row_index)
    data.dropna(subset=[data.columns[0]], inplace=True)

    expected_columns = ["S.No", "Chainage", "Finished Roadway Width", "Finished Vertical Height",
                        "Original Roadway Width", "Area Coefficient", "Cutting slope"]
    data = data.iloc[:, :7]
    data.columns = expected_columns

    # Calculate Area and Volume
    data['Area (m²)'] = data.apply(lambda row: row['Area Coefficient'] * (row['Finished Roadway Width'] - row['Original Roadway Width']) * row['Finished Vertical Height'], axis=1)
    data['Volume (m³)'] = 0.0
    chainage_values = data['Chainage'].astype(str).str.replace("+", "").astype(float).values
    for i in range(1, len(data)):
        delta_ch = chainage_values[i] - chainage_values[i - 1]
        data.at[i, 'Volume (m³)'] = delta_ch * data.at[i, 'Area (m²)']
    total_volume = data['Volume (m³)'].sum()
    data.loc[len(data.index)] = ['Total', '', '', '', '', '', '', '', total_volume]

    rows, cols = 2, 2
    figsize = (11.7, 8.3)
    plot_count = 0
    fig, axs = plt.subplots(rows, cols, figsize=figsize)
    axs = axs.flatten()

    def plot_chainage_subplot(entry, ax):
        chainage = entry['Chainage']
        fw = entry['Finished Roadway Width']
        fh = entry['Finished Vertical Height']
        ow = entry['Original Roadway Width']
        ac = entry['Area Coefficient']
        angle_deg = entry['Cutting slope']

        if pd.isna(angle_deg):
        st.error(f"❌ Missing Cutting slope at Chainage {chainage}. Please correct this value in your file.")
        st.stop()
    try:
            angle_deg = float(angle_deg)
            slope = 1 / np.tan(np.radians(angle_deg))
        except Exception:
        st.error(f"❌ Invalid Cutting slope value at Chainage {chainage}. Please correct this value in your file.")
        st.stop()
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
        ax.text(min(fg_x), fh + 0.1, f"FR: {fw} m", fontsize=7, color='green')
        ax.text(min(fg_x), fh + 0.25, f"OR: {ow} m", fontsize=7, color='red')
        ax.text(min(fg_x), fh + 0.4, f"Height: {fh} m", fontsize=7, color='black')

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

    st.info("⏳ Generating plots... please wait")
    progress_bar = st.progress(0)
    total_steps = len(data)
    current_step = 0

    for idx, row in data.iterrows():
        if row.notna().all():
            ax = axs[plot_count % (rows * cols)]
            plot_chainage_subplot(row, ax)

            # Save preview of all plots
            buf = io.BytesIO()
            fig_preview, ax_preview = plt.subplots()
            plot_chainage_subplot(row, ax_preview)
            fig_preview.savefig(buf, format='png')
            plt.close(fig_preview)
            all_imgs.append(buf.getvalue())

            plot_count += 1
            current_step += 1
            progress_bar.progress(min(current_step / total_steps, 1.0))

            if plot_count % (rows * cols) == 0:
                pdf.savefig(fig)
                plt.close(fig)
                fig, axs = plt.subplots(rows, cols, figsize=figsize)
                axs = axs.flatten()

    if plot_count % (rows * cols) != 0:
        for i in range(plot_count % (rows * cols), rows * cols):
            fig.delaxes(axs[i])
        pdf.savefig(fig)
        plt.close(fig)

    pdf.close()

    st.success("✅ Plots generated successfully!")
    if all_imgs:
        st.subheader("📑 Preview All Plots")
        for i, img in enumerate(all_imgs):
            st.image(img, caption=f"Chainage Plot {i+1}", use_column_width=True)

    st.download_button(
        label="📥 Download Plotted Cross-Sections PDF",
        data=pdf_buffer.getvalue(),
        file_name="CrossSection_Plots.pdf",
        mime="application/pdf"
    )
