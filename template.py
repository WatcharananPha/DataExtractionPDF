import google.generativeai as genai
import pandas as pd
import io
import os
import streamlit as st
from dotenv import load_dotenv

st.title("Quotation Comparison Table Generator")

load_dotenv()
API_KEY = os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=API_KEY)

uploaded_file = st.file_uploader("อัปโหลดไฟล์ใบเสนอราคา", type=["excel", "xlsx", "txt"])

if uploaded_file:
    try:
        input_quotation_data = uploaded_file.read().decode("utf-8")
    except UnicodeDecodeError:
        try:
            uploaded_file.seek(0)
            input_quotation_data = uploaded_file.read().decode("cp874")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            input_quotation_data = uploaded_file.read().decode("latin1")

    output_excel_path = "Quotation_Comparison_Result.xlsx"

    prompt = f"""
You are an expert procurement and data extraction assistant. Your task is to analyze the provided raw text containing multiple quotation documents from different suppliers and consolidate the pricing for similar items into a single, standardized comparison table.

**Input Data Description:**
The input text contains exactly 4 distinct quotation blocks. Each block typically includes:
-   Supplier Company Information (e.g., "บริษัท เอบีไอ เทสติ้ง เอ็นจิเนียริ่ง จำกัด")
-   Recipient Information (e.g., "เรียน:", "ที่อยู่:", "ATTN:", "โครงการ:")
-   Itemized List (with columns like "ลำดับ", "รายการ", "จำนวน", "หน่วย", "ราคา/หน่วย", "ราคารวม")
-   Summary totals (e.g., "ราคารวม", "ภาษีมูลค่าเพิ่ม 7%", "ยอดรวมทั้งสิ้น")

**Your Task (Output Structure - Tab-Separated Table):**

1.  **Identify Each Quote & Supplier:**
    *   Parse the input text to identify each of the 4 distinct quotation blocks.
    *   Extract the primary supplier company name for each block to use as column headers (e.g., "เอบีไอ เทสติ้ง", "บี.บี.เค ไพล์เทสติ้ง", "BOULTER STEWART", "S.K.E CONSULTANS").

2.  **Normalize Item Names:**
    *   Identify all unique items/services across all 4 quotations.
    *   Normalize similar item names to a consistent, concise representation (e.g., "Dynamic Load Test" and "Dynamic Load Test Service" should become "Dynamic Load Test").

3.  **Construct the Item Comparison Table:**
    *   The table must be tab-separated.
    *   **Header Row:**
        "ลำดับ\tรายการ\tจำนวน\tหน่วย\t{{Supplier1_ShortName}}\t{{Supplier2_ShortName}}\t{{Supplier3_ShortName}}\t{{Supplier4_ShortName}}"
        (Replace {{SupplierX_ShortName}} with the extracted, concise supplier names from step 1).
    *   **Item Rows:**
        *   `ลำดับ`: Running number (1, 2, 3...) for each unique item.
        *   `รายการ`: Normalized item name.
        *   `จำนวน`: Quantity (take from the most common quantity or the first supplier where the item appears).
        *   `หน่วย`: Unit (take from the most common unit or the first supplier where the item appears).
        *   `{{SupplierX_ShortName}}` columns: The 'ราคารวม' (total price for that specific item) from Supplier X. If an item is not found in a supplier's quote, leave the cell empty (""). If the value is "รวมแล้ว", convert it to "". Ensure all numeric values are clean (no commas).

4.  **Summary Rows:** After all item rows, include these summary rows. The "ราคารวม", "ภาษีมูลค่าเพิ่ม 7%", "ยอดรวมทั้งสิ้น" values should be the exact values extracted from each respective supplier's quote, not summed from the comparison table items (as some items might be "รวมแล้ว" or "Discount").
    *   `ราคารวม\t\t\t\t[Extracted ราคารวม from Supplier 1]\t[Extracted ราคารวม from Supplier 2]\t[Extracted ราคารวม from Supplier 3]\t[Extracted ราคารวม from Supplier 4]`
    *   `ภาษีมูลค่าเพิ่ม 7%\t\t\t\t[Extracted ภาษีมูลค่าเพิ่ม 7% from Supplier 1]\t[Extracted ภาษีมูลค่าเพิ่ม 7% from Supplier 2]\t[Extracted ภาษีมูลค่าเพิ่ม 7% from Supplier 3]\t[Extracted ภาษีมูลค่าเพิ่ม 7% from Supplier 4]`
    *   `ยอดรวมทั้งสิ้น\t\t\t\t[Extracted ยอดรวมทั้งสิ้น from Supplier 1]\t[Extracted ยอดรวมทั้งสิ้น from Supplier 2]\t[Extracted ยอดรวมทั้งสิ้น from Supplier 3]\t[Extracted ยอดรวมทั้งสิ้น from Supplier 4]`
    *   **Important:** Ensure all numeric values are clean (no commas or currency symbols) for easy parsing.

5.  **Supplier Conditions Section:** After the summary table, add a new section for supplier conditions. For the given input, these conditions are not explicitly present, so state "N/A" for each supplier.
    *   `เงื่อนไขการของ SUPPLIER\t\t\t\t{{Supplier1_ShortName}}\t{{Supplier2_ShortName}}\t{{Supplier3_ShortName}}\t{{Supplier4_ShortName}}` (This will be the header for this section, with supplier names repeated)
    *   `1\tกำหนดยืนราคา\t\t\t\tN/A\tN/A\tN/A\tN/A`
    *   `2\tระยะเวลาส่งมอบสินค้าหลังจากได้รับ PO\t\t\t\tN/A\tN/A\tN/A\tN/A`
    *   `3\tการชำระเงิน\t\t\t\tN/A\tN/A\tN/A\tN/A`
    *   `4\tอื่น ๆ\t\t\t\tN/A\tN/A\tN/A\tN/A`

**Final Output Constraints:**
Your entire response must be a single block of plain tab-separated text, ready to be read by a spreadsheet program.
Do not include any explanations, summaries, or markdown formatting (like ```csv).

---
**Input Quotation Data:**
{input_quotation_data}
"""

    model = genai.GenerativeModel(model_name="gemini-2.5-flash")
    response = model.generate_content(prompt)
    output_text = response.text.strip()

    if output_text.startswith("```"):
        output_text = output_text.split('\n', 1)[1]
    if output_text.endswith("```"):
        output_text = output_text.rsplit('\n', 1)[0]

    lines = [line for line in output_text.splitlines() if line.strip()]
    split_lines = [line.rstrip('\n').split('\t') for line in lines]
    max_cols = max(len(row) for row in split_lines)
    normalized_lines = [row + [''] * (max_cols - len(row)) for row in split_lines]

    df_raw = pd.DataFrame(normalized_lines)

    summary_start_row_idx = df_raw[df_raw.iloc[:, 0].str.contains('ราคารวม', na=False)].index.tolist()
    conditions_header_row_idx = df_raw[df_raw.iloc[:, 0].str.contains('เงื่อนไขการของ SUPPLIER', na=False)].index.tolist()

    header_row_values = df_raw.iloc[0].tolist()
    main_table_data = df_raw.iloc[1:summary_start_row_idx[0]]
    summary_data = df_raw.iloc[summary_start_row_idx[0]:conditions_header_row_idx[0]]
    conditions_header = df_raw.iloc[conditions_header_row_idx[0]].tolist()
    conditions_data = df_raw.iloc[conditions_header_row_idx[0] + 1:]

    df_main_table = pd.DataFrame(main_table_data.values, columns=header_row_values)

    numeric_cols = ['จำนวน'] + header_row_values[4:]
    for col in numeric_cols:
        if col in df_main_table.columns:
            df_main_table[col] = pd.to_numeric(df_main_table[col], errors='coerce').fillna('')

    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Quotation Comparison')

        for col_idx, value in enumerate(header_row_values):
            worksheet.write(0, col_idx, value)

        for row_idx, row_data in enumerate(df_main_table.values):
            for col_idx, value in enumerate(row_data):
                worksheet.write(row_idx + 1, col_idx, value)

        current_row = len(df_main_table) + 1

        for row_idx, row_data in enumerate(summary_data.values):
            for col_idx, value in enumerate(row_data):
                worksheet.write(current_row + row_idx, col_idx, value)

        current_row += len(summary_data) + 1

        for col_idx, value in enumerate(conditions_header):
            worksheet.write(current_row, col_idx, value)
        current_row += 1

        for row_idx, row_data in enumerate(conditions_data.values):
            for col_idx, value in enumerate(row_data):
                worksheet.write(current_row + row_idx, col_idx, value)

    st.success(f"สร้างไฟล์เปรียบเทียบใบเสนอราคาเสร็จสมบูรณ์: {output_excel_path}")
    with open(output_excel_path, "rb") as f:
        st.download_button("ดาวน์โหลดไฟล์ Excel", f, file_name=output_excel_path)