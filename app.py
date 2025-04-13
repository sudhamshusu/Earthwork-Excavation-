import streamlit as st 
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import io

st.set_page_config(layout="wide")
st.title("📊 Earthwork Excavation Cross-Section Plotter")

# 📘 Instructions 
st.markdown("""
### 📌 How to Use This Tool:
Use the following area coefficients based on the type of cutting:
- Fresh Cutting – use 0.5
- Back Cutting – use 0.67
- Box Cutting – use 1.0.

These values are based on standard thumb rules to estimate earthwork cross-sectional areas.

1. **Download the Templates** using the buttons below.
2. **Fill out the Input Template (`sample.xlsx`)** with chainage, widths, height, slope, etc.
3. **Fill out the Summary Template (`summary.xlsx`)** with project-specific details.
4. **Upload both filled files without changing file name** using the file uploaders.
5. **Fill out Contract Identification No:** below to personalize your plots.
6. Click the **Generate Cross Section Plots** button.
7. **Preview plots** and **Download final report** as PDF and Excel summary.
""")

# 📁 Download Template Files
st.markdown("### 📁 Download Templates")
col1, col2 = st.columns(2)
with col1:
    with open("Sample.xlsx", "rb") as sample_file:
        st.download_button("📥 Download Input Template (Sample.xlsx)", sample_file, file_name="Sample.xlsx")
with col2:
    with open("Summary.xlsx", "rb") as summary_file:
        st.download_button("📥 Download Summary Template (Summary.xlsx)", summary_file, file_name="Summary.xlsx")

# 📤 Upload Section
data_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])
summary_file = st.file_uploader("Upload Editable Summary Excel (Optional)", type=["xlsx"])

# 📋 Optional Manual Summary
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

# 🖼 Plotting function
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
    except:
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

# ✅ Main Generate Button
if data_file and st.button("📊 Generate Cross Section Plots"):
    try:
        # Step 1: Read data
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

        # Step 2: Area and Volume Calculation
        data['Area (m²)'] = data.apply(
            lambda row: row['Area Coefficient'] * (row['Finished Roadway Width'] - row['Original Roadway Width']) * row['Finished Vertical Height'], axis=1
        )
        chainage_values = data['Chainage'].astype(str).str.replace("+", "").astype(float).values

        volume_list = [0.0]
        for i in range(1, len(data)):
            ch1, ch2 = chainage_values[i - 1], chainage_values[i]
            a1, a2 = data.at[i - 1, 'Area (m²)'], data.at[i, 'Area (m²)']
            vol = ((a1 + a2) / 2) * (ch2 - ch1)
            volume_list.append(vol)

        data['Volume (m³)'] = volume_list

        # Step 3: Contract Info
        contract_no = contract_no_input
        if summary_file:
            summary_df = pd.read_excel(summary_file, header=None, nrows=10)
            for row in summary_df.itertuples(index=False):
                for i, val in enumerate(row):
                    if isinstance(val, str) and "contract" in val.lower() and "identification" in val.lower():
                        contract_no = str(row[i + 1]) if i + 1 < len(row) else contract_no
                        break

        # Step 4: PDF Plot Generation
        pdf_buffer = io.BytesIO()
        pdf = PdfPages(pdf_buffer)

        if summary_file or use_manual_summary:
            fig_summary, ax_summary = plt.subplots(figsize=(11.7, 8.3))
            ax_summary.axis('off')
            if summary_file:
                table_data = summary_df.dropna(how='all').dropna(axis=1, how='all')
                cell_text = table_data.astype(str).values.tolist()
            else:
                cell_text = [[k, str(v)] for k, v in manual_summary.items()]
            table = ax_summary.table(cellText=cell_text, colLabels=None, loc='center', cellLoc='left')
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.scale(1, 1.5)
            fig_summary.suptitle("Project Summary", fontsize=12)
            pdf.savefig(fig_summary)
            plt.close(fig_summary)

        rows, cols = 2, 2
        figsize = (11.7, 8.3)
        fig, axs = plt.subplots(rows, cols, figsize=figsize)
        axs = axs.flatten()
        fig.suptitle(contract_no, fontsize=10, x=0.5, y=0.98)

        plot_count = 0
        preview_imgs = []
        progress = st.progress(0, text="Generating plots...")
        step = 1 / max(len(data), 1)

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
                progress.progress(min(int(plot_count * step * 100), 100), text="Processing...")

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
        progress.progress(100, text="Done!")

        st.success("✅ Plots generated successfully!")
        st.download_button(
            label="📥 Download PDF with Plots",
            data=pdf_buffer.getvalue(),
            file_name="CrossSection_Plots.pdf",
            mime="application/pdf"
        )

        if preview_imgs:
            st.subheader("🔍 Preview of Plots (up to 40 shown)")
            cols = st.columns(4)
            for i, img in enumerate(preview_imgs):
                with cols[i % 4]:
                    st.image(img, use_column_width=True)

        # Step 5: Generate Volume Calculation Sheet (Excel)
        minimal_df = pd.DataFrame({
            "S.No": data["S.No"],
            "Chainage": data["Chainage"],
            "Area": data["Area (m²)"].round(3),
            "Volume": data["Volume (m³)"].round(3)
        })

        total_volume = minimal_df["Volume"].sum()
        total_row = {
            "S.No": "Total",
            "Chainage": "",
            "Area": "",
            "Volume": round(total_volume, 3)
        }
        minimal_df = pd.concat([minimal_df, pd.DataFrame([total_row])], ignore_index=True)

        excel_buffer_min = io.BytesIO()
        with pd.ExcelWriter(excel_buffer_min) as writer:
            minimal_df.to_excel(writer, index=False, sheet_name='Volume Calculation Sheet')

        st.download_button(
            label="📥 Download Volume Calculation Sheet (S.No, Chainage, Area, Volume)",
            data=excel_buffer_min.getvalue(),
            file_name="Volume_Calculation_Sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
