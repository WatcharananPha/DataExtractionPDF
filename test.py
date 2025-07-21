import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
import os
from dotenv import load_dotenv
import tempfile

st.set_page_config(page_title="PDF Extractor", layout="centered")

st.title("PDF Extractor")

with st.expander("Instructions"):
    st.markdown("""
    - Upload one or more PDF files using the uploader below.
    - Click **Extract Data** to process the uploaded PDFs.
    - The extracted data will be merged and available for download as an Excel file.
    """)

load_dotenv()
API_KEY = os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=API_KEY)

output_excel_path = "level-1.xlsx"

prompt = """
You are a highly precise data extraction AI. Your mission is to convert the provided PDF quotation into a specific CSV format by treating the document as two distinct sections and then merging them.

**Execution Plan:**

1.  **Section 1: Header Information Extraction**
    *   First, identify the recipient's details at the top: 'เรียน' (Company), 'ที่อยู่' (Address), 'ATTN' (Attention), and 'โครงการ' (Project).
    *   You will create a separate CSV row for each of these pieces of information.

2.  **Section 2: Line Items Table Extraction**
    *   Next, identify the main table containing services with columns: "ลำดับ", "รายการ", "จำนวน", "หน่วย", "ราคา/หน่วย", "ราคารวม".
    *   Extract every row from this table, including rows that act as sub-headers (e.g., "Dynamic Load Test").
    *   **CRUCIAL: The "จำนวน" column must contain only numeric values (integers or decimals). If the value is not a number, leave it as an empty quoted field (""). Do not include any non-numeric text in this column.**

3.  **Section 3: Summary Extraction**
    *   At the end of the document, extract summary rows such as "ราคารวม", "ภาษีมูลค่าเพิ่ม 7%", and "ยอดรวมทั้งสิ้น" with their corresponding values.
    *   For these summary rows, place the summary label in the 'รายการ' column and the value in the 'ราคารวม' column. All other columns must be empty quoted fields ("").

4.  **Data Validation (CRUCIAL):**
    *   Before outputting the CSV, carefully check the context of the PDF to ensure that all extracted data is correct and matches the original document.
    *   If you detect any inconsistencies or errors in the extracted data, correct them before generating the CSV.

5.  **CSV Construction Rules (Crucial):**
    *   The final output MUST be a single CSV text block.
    *   The header must be exactly: "รายชื่อบริษัทและการติดต่อ","ลำดับ","รายการ","จำนวน","หน่วย","ราคาต่อหน่วย","ราคารวม"
    *   **CRUCIAL QUOTING RULE**: Every single field in every row **MUST** be enclosed in double quotes (""). This is mandatory to handle commas within text fields. Example: "Value with, a comma","Value 2","","Value 4"
    *   **For Section 1 Rows (Header Info):**
        *   Place the extracted text It is the name of the company that created the quotation that is on the first letterhead. (e.g., "บริษัท เอบีไอ เทสติ้ง เอ็นจิเนียริ่ง จำกัด ABI  TESTING  ENGINEERING  CO.,LTD สำนักงานใหญ่เลขที่ 9/317 หมู่ที่ 3 ตำบลบางขนุน อำเภอบางกรวย จังหวัดนนทบุรี 11130 มือถือ 096 991 5545
Email : abitestingeng@gmail.com เลขประจำตัวผู้เสียภาษี  0 1255 66040 07 1") in the first column (`รายชื่อบริษัทและการติดต่อ`).
        *   All other six columns for these rows MUST be empty fields (represented as "").
    *   **For Section 2 Rows (Line Items):**
        *   The first column (`รายชื่อบริษัทและการติดต่อ`) MUST be an empty field ("").
        *   Populate columns 2 through 7 with the data extracted from the table. If a cell is empty, it must be represented as an empty quoted field ("").
    *   **For Section 3 Rows (Summary):**
        *   The first, second, fourth, fifth, and sixth columns MUST be empty fields ("").
        *   Place the summary label (e.g., "ราคารวม", "ภาษีมูลค่าเพิ่ม 7%", "ยอดรวมทั้งสิ้น") in the third column (`รายการ`).
        *   Place the summary value in the seventh column (`ราคารวม`).

6.  **Final Output Constraints**:
    *   Your entire response must contain **ONLY the raw CSV data** and nothing else.
    *   Do not include any explanations, summaries, or markdown formatting like ```csv.
"""

uploaded_files = st.file_uploader(
    "Upload one or more PDF files", 
    type=["pdf"], 
    accept_multiple_files=True
)

if st.button("Extract Data") and uploaded_files and output_excel_path.strip():
    all_dfs = []
    progress = st.progress(0)
    for idx, uploaded_file in enumerate(uploaded_files):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_file_path = tmp_file.name
        with st.spinner(f"Processing {uploaded_file.name}..."):
            gemini_file = genai.upload_file(path=tmp_file_path, display_name=f"PDF for Extraction: {uploaded_file.name}")
            response = genai.GenerativeModel(model_name="gemini-2.5-flash").generate_content([prompt, gemini_file])
            csv_data_string = response.text.strip()
            if csv_data_string.startswith("```csv"):
                csv_data_string = csv_data_string[len("```csv"):].strip()
            if csv_data_string.endswith("```"):
                csv_data_string = csv_data_string[:-len("```")].strip()
            string_data_io = io.StringIO(csv_data_string)
            df = pd.read_csv(string_data_io, quotechar='"')
            all_dfs.append(df)
        progress.progress((idx + 1) / len(uploaded_files))
        os.unlink(tmp_file_path)
    if all_dfs:
        merged_df = pd.concat(all_dfs, ignore_index=True)
        merged_df.to_excel(output_excel_path, index=False)
        st.success(f"Successfully extracted and merged data from all PDFs to {output_excel_path}")
        st.dataframe(merged_df)
        with open(output_excel_path, "rb") as f:
            st.download_button(
                label="Download Extracted Excel File",
                data=f,
                file_name=output_excel_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("No PDF files found or extracted.")