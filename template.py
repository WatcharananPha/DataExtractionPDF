import streamlit as st
import google.generativeai as genai
import json
import tempfile
import os
import gspread
import re
from difflib import SequenceMatcher
from google.oauth2.service_account import Credentials
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv

st.set_page_config(page_title="PDF Extractor", layout="centered")
st.title("PDF Extractor")

with st.expander("Instructions"):
    st.markdown("""
    - Upload one or more PDF files using the uploader below.
    - Enter your Google Sheet ID or use the default one.
    - Click **Extract and Update** to process the uploaded PDFs and update Google Sheet.
    - The extracted data will also be available for download as a JSON file.
    """)

load_dotenv()
API_KEY = os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=API_KEY)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file"
]
CREDS_FILE = 'glassy-keyword-466817-t8-8bdc0c7acadc.json'
DEFAULT_SHEET_ID = '17tMHStXQYXaIQHQIA4jdUyHaYt_tuoNCEEuJCstWEuw'

COMPANY_NAME_ROW = 1
CONTACT_INFO_ROW = 2
HEADER_ROW = 3
ITEM_MASTER_LIST_COL = 2
COLUMNS_PER_SUPPLIER = 4
SIMILARITY_THRESHOLD = 0.7

prompt = """# System Message for Product List Extraction (PDF/Text Table Processing)

## Input Format
Provide PDF files (or images/tables with text extraction) containing product information (receipts, invoices, product lists, etc.):
* Upload PDF files or images directly
* Text will be extracted automatically
* Source should contain clear product details in readable text

## Task
Analyze the extracted text from documents and extract ALL product information to create a comprehensive product summary list.
Do not skip any products – ensure complete extraction of every line item, including physical consumables, product accessories, and items such as "ค่าเจาะรูกระจก" if they are charged per unit.

## CRITICAL: Complete Product Collection Rule
Extract ALL products as individual entries - quotations typically contain unique, detailed specifications.
- Each line item in the quotation represents a distinct product with specific details
- DO NOT consolidate or merge products - preserve all individual entries exactly as listed
- Products in quotations are already unique due to detailed specifications (sizes, locations, materials, etc.)
- Extract every single row/line item that has quantity, unit price, and total price
- Maintain the exact product descriptions including all specifications and location details

## Output Format (JSON only)
You must return ONLY this JSON structure:
{
  "company": "company name or first name + last name (NEVER null)",
  "vat": true,
  "name": "customer name or null",
  "contact": "phone number or email or null",
  "priceGuaranteeDay": 0,
  "products": [
    {
      "name": "full product description in Thai including all specifications AND size (EXCLUDE quantity/unit/price)",
      "quantity": 1,
      "unit": "match the unit shown in the pricePerUnit column (e.g., แผ่น, ตร.ม., ชิ้น, ตัว, เมตร, ชุด)",
      "pricePerUnit": 0,
      "totalPrice": 0
    }
  ],
  "totalPrice": 0,
  "totalVat": 0,
  "totalPriceIncludeVat": 0
}

## Field Extraction Guidelines

### name (Product Description)
* Include: material, model/type, size/dimensions, technical specs, finish, location details
* Include: ALL distinguishing characteristics that make each product unique
* Exclude: quantity, unit, price
* Preserve detailed specifications exactly as shown in quotation (sizes, locations, installation details)
* Example: "งานกระจกบานใส กระจกเทมเปอร์ เปรย์ เกร์ 10 มม. ฝังตัวยูเหล็ก สีเทา บันได ชั้น 1-2 ขนาด 0.975x4.672 ม."
* For items like "ค่าเจาะรูกระจก กว้าง 16มม.", include all distinguishing attributes in name

### unit and quantity (DIRECT EXTRACTION RULE)
* Extract unit and quantity DIRECTLY from each line item as shown
* Use the exact unit shown in the quotation (ชุด, แผ่น, ตร.ม., ชิ้น, ตัว, เมตร, etc.)
* Each line item represents a separate product entry - extract as individual entries
* If an item is priced per unit and has physical/product characteristics, treat as a product and include

### pricePerUnit and totalPrice
* If the document shows both product/material cost (ค่าวัสดุ) and labor/service cost (ค่าแรง) in the same row, always add the two (pricePerUnit = product per unit + service per unit)
* totalPrice must be calculated as: quantity × pricePerUnit
* Use numeric values only (no currency symbols)
* Extract cleanly from pricing fields as shown in each line item
* Allow minor rounding errors if visible in the document
* Extract each line item separately - do not combine or consolidate pricing
* If the quote has discount items, select only the discounted price

## Complete Product Extraction Algorithm
1. Identify all line items with product descriptions, quantities, and prices
2. Extract each line item as a separate product - do not group or consolidate
3. Preserve all specifications and distinguishing details in the product name
4. Include location/installation details that make each product unique
5. Maintain individual quantities and pricing exactly as shown in quotation

## Inclusion/Exclusion Rules
* Extract ALL physical products and ALL items with price per unit, including consumables, parts, and accessories
* Include "ค่าเจาะรูกระจก" and similar items IF they are presented as a per-unit/consumable/physical item in the product table
* Exclude service/labor cost lines without a per-unit count (e.g., lump sum services)
* Each line item with specifications = separate product entry - preserve all individual entries
* Include all location-specific variations (different floors, units, areas) as separate products

## Quality Assurance Checklist
- [ ] ALL line items with products, quantities, and prices extracted
- [ ] Each product entry preserves detailed specifications and location details
- [ ] Product names include all distinguishing characteristics
- [ ] Quantities and prices match exactly what's shown in quotation
- [ ] No line items missed - complete extraction achieved

## Final Notes
* PRIMARY GOAL: Extract every single product line item - quotations contain unique, detailed specifications
* Extract all line items meeting the above rules - preserve each as individual product entry
* Include all specs, sizes, and location details in name for complete product identification
* Match unit and quantity with pricePerUnit exactly as shown in quotation
* Maintain complete fidelity to the original quotation structure and details
* If the quote has discount items, select only the discounted price
"""

def extract_json_from_text(text):
    start_idx = text.find('{')
    end_idx = text.rfind('}') + 1
    if start_idx >= 0 and end_idx > start_idx:
        return json.loads(text[start_idx:end_idx])
    return None

def authenticate_and_open_sheet(sheet_id):
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)
    return spreadsheet.get_worksheet(0)

def ensure_first_three_rows_exist(worksheet):
    payloads = []
    for i in range(1, 4):
        payloads.append({
            'range': f"A{i}:B{i}",
            'values': [["", ""]]
        })
    if payloads:
        worksheet.batch_update(payloads, value_input_option='USER_ENTERED')

def extract_key_product_features(product_name):
    # Extract material type
    material_types = ["กระจก", "อลูมิเนียม", "เหล็ก", "ไม้", "พลาสติก", "สแตนเลส"]
    material = next((m for m in material_types if m in product_name), "")
    
    # Extract dimensions if present
    dimensions = re.findall(r'\d+(?:\.\d+)?(?:\s*[xX×]\s*\d+(?:\.\d+)?)+(?:\s*(?:ม\.|มม\.|cm|MM|mm))?', product_name)
    dims = dimensions[0] if dimensions else ""
    
    # Extract thickness if present
    thickness = re.findall(r'(\d+(?:\.\d+)?)\s*(?:มม\.|mm)', product_name)
    thick = thickness[0] if thickness else ""
    
    # Extract location information if present
    locations = ["ชั้น", "บันได", "ห้อง", "ประตู", "หน้าต่าง"]
    location = next((loc for loc in locations if loc in product_name), "")
    
    # Extract product type or function
    product_types = ["บานเลื่อน", "บานสวิง", "บานพับ", "บานกระทุ้ง", "บานเปิด", "บานเฟี้ยม", "กระจกเปลือย"]
    prod_type = next((pt for pt in product_types if pt in product_name), "")
    
    return {
        "material": material,
        "dimensions": dims,
        "thickness": thick,
        "location": location,
        "product_type": prod_type,
        "full_name": product_name
    }

def calculate_product_similarity(product1, product2):
    # Extract key features
    features1 = extract_key_product_features(product1)
    features2 = extract_key_product_features(product2)
    
    # Calculate basic text similarity
    name_similarity = SequenceMatcher(None, product1, product2).ratio()
    
    # Calculate feature-specific similarities
    material_match = 1.0 if features1["material"] and features1["material"] == features2["material"] else 0.0
    thickness_match = 1.0 if features1["thickness"] and features1["thickness"] == features2["thickness"] else 0.0
    product_type_match = 1.0 if features1["product_type"] and features1["product_type"] == features2["product_type"] else 0.0
    
    # Weighted similarity score (adjust weights as needed)
    weighted_score = (name_similarity * 0.4) + (material_match * 0.25) + (thickness_match * 0.2) + (product_type_match * 0.15)
    
    return weighted_score

def find_matching_products(new_product, existing_products):
    scores = []
    for i, existing in enumerate(existing_products):
        similarity = calculate_product_similarity(new_product, existing)
        if similarity >= SIMILARITY_THRESHOLD:
            scores.append((i, similarity))
    
    # Return best match if any
    return max(scores, key=lambda x: x[1])[0] if scores else None

def update_google_sheet_with_multiple_files(worksheet, all_json_data):
    ensure_first_three_rows_exist(worksheet)
    
    # Get existing product items from column B (after header rows)
    existing_items = worksheet.col_values(ITEM_MASTER_LIST_COL)[3:]
    start_row_index = 4  # Start after header rows
    
    payloads = []
    
    # First supplier's data is added directly
    if all_json_data:
        first_json_data = all_json_data[0]
        supplier_col_start = ITEM_MASTER_LIST_COL + 1
        
        # Add company name and contact info
        payloads.append({
            'range': f"{get_column_letter(supplier_col_start)}{COMPANY_NAME_ROW}",
            'values': [[first_json_data["company"]]]
        })
        
        payloads.append({
            'range': f"{get_column_letter(supplier_col_start)}{CONTACT_INFO_ROW}",
            'values': [[first_json_data["contact"] or ""]]
        })
        
        # Add header row
        payloads.append({
            'range': f"{get_column_letter(supplier_col_start)}{HEADER_ROW}:{get_column_letter(supplier_col_start + 3)}{HEADER_ROW}",
            'values': [["ปริมาณ", "หน่วย", "ราคาต่อหน่วย", "รวมเป็นเงิน"]]
        })
        
        # Add all products from first supplier
        for product in first_json_data["products"]:
            product_name = product["name"]
            existing_items.append(product_name)
            row_index = len(existing_items) + start_row_index - 1
            
            payloads.append({
                'range': f"B{row_index}",
                'values': [[product_name]]
            })
            
            payloads.append({
                'range': f"{get_column_letter(supplier_col_start)}{row_index}:{get_column_letter(supplier_col_start + 3)}{row_index}",
                'values': [[product["quantity"], product["unit"], product["pricePerUnit"], product["totalPrice"]]]
            })
        
        # Add summary rows for first supplier
        summary_start_row = len(existing_items) + start_row_index + 1
        
        payloads.append({
            'range': f"B{summary_start_row}",
            'values': [["รวมเป็นเงิน"]]
        })
        payloads.append({
            'range': f"{get_column_letter(supplier_col_start + 3)}{summary_start_row}",
            'values': [[first_json_data["totalPrice"]]]
        })
        
        summary_start_row += 1
        payloads.append({
            'range': f"B{summary_start_row}",
            'values': [["ภาษีมูลค่าเพิ่ม 7%"]]
        })
        payloads.append({
            'range': f"{get_column_letter(supplier_col_start + 3)}{summary_start_row}",
            'values': [[first_json_data["totalVat"]]]
        })
        
        summary_start_row += 1
        payloads.append({
            'range': f"B{summary_start_row}",
            'values': [["ยอดรวมทั้งสิ้น"]]
        })
        payloads.append({
            'range': f"{get_column_letter(supplier_col_start + 3)}{summary_start_row}",
            'values': [[first_json_data["totalPriceIncludeVat"]]]
        })
        
        # Process additional suppliers with matching logic
        for idx, json_data in enumerate(all_json_data[1:], start=1):
            supplier_col_start = ITEM_MASTER_LIST_COL + 1 + (idx * COLUMNS_PER_SUPPLIER)
            
            # Add company name and contact info
            payloads.append({
                'range': f"{get_column_letter(supplier_col_start)}{COMPANY_NAME_ROW}",
                'values': [[json_data["company"]]]
            })
            
            payloads.append({
                'range': f"{get_column_letter(supplier_col_start)}{CONTACT_INFO_ROW}",
                'values': [[json_data["contact"] or ""]]
            })
            
            # Add header row
            payloads.append({
                'range': f"{get_column_letter(supplier_col_start)}{HEADER_ROW}:{get_column_letter(supplier_col_start + 3)}{HEADER_ROW}",
                'values': [["ปริมาณ", "หน่วย", "ราคาต่อหน่วย", "รวมเป็นเงิน"]]
            })
            
            matched_products = []
            new_products = []
            
            # Find matches for each product or add as new
            for product in json_data["products"]:
                product_name = product["name"]
                match_idx = find_matching_products(product_name, existing_items)
                
                if match_idx is not None:
                    # Match found
                    row_index = match_idx + start_row_index
                    matched_products.append((product, row_index))
                else:
                    # No match, add as new product
                    new_products.append(product)
            
            # Update matched products
            for product, row_index in matched_products:
                payloads.append({
                    'range': f"{get_column_letter(supplier_col_start)}{row_index}:{get_column_letter(supplier_col_start + 3)}{row_index}",
                    'values': [[product["quantity"], product["unit"], product["pricePerUnit"], product["totalPrice"]]]
                })
            
            # Add new products
            for product in new_products:
                product_name = product["name"]
                existing_items.append(product_name)
                row_index = len(existing_items) + start_row_index - 1
                
                payloads.append({
                    'range': f"B{row_index}",
                    'values': [[product_name]]
                })
                
                payloads.append({
                    'range': f"{get_column_letter(supplier_col_start)}{row_index}:{get_column_letter(supplier_col_start + 3)}{row_index}",
                    'values': [[product["quantity"], product["unit"], product["pricePerUnit"], product["totalPrice"]]]
                })
            
            # Add summary rows
            summary_start_row = len(existing_items) + start_row_index + 1
            
            payloads.append({
                'range': f"B{summary_start_row}",
                'values': [["รวมเป็นเงิน"]]
            })
            payloads.append({
                'range': f"{get_column_letter(supplier_col_start + 3)}{summary_start_row}",
                'values': [[json_data["totalPrice"]]]
            })
            
            summary_start_row += 1
            payloads.append({
                'range': f"B{summary_start_row}",
                'values': [["ภาษีมูลค่าเพิ่ม 7%"]]
            })
            payloads.append({
                'range': f"{get_column_letter(supplier_col_start + 3)}{summary_start_row}",
                'values': [[json_data["totalVat"]]]
            })
            
            summary_start_row += 1
            payloads.append({
                'range': f"B{summary_start_row}",
                'values': [["ยอดรวมทั้งสิ้น"]]
            })
            payloads.append({
                'range': f"{get_column_letter(supplier_col_start + 3)}{summary_start_row}",
                'values': [[json_data["totalPriceIncludeVat"]]]
            })
    
    # Execute all updates in one batch
    if payloads:
        worksheet.batch_update(payloads, value_input_option='USER_ENTERED')
    
    return len(all_json_data)

uploaded_files = st.file_uploader(
    "Upload one or more PDF files", 
    type=["pdf"], 
    accept_multiple_files=True
)

sheet_id = st.text_input("Google Sheet ID", value=DEFAULT_SHEET_ID)
similarity_threshold = st.slider("Product Matching Similarity Threshold", 0.5, 0.95, 0.7, 0.05)
SIMILARITY_THRESHOLD = similarity_threshold

if st.button("Extract and Update") and uploaded_files and sheet_id:
    all_data = []
    progress = st.progress(0)
    
    worksheet = authenticate_and_open_sheet(sheet_id)
    st.info("Connected to Google Sheet successfully.")
    
    for idx, uploaded_file in enumerate(uploaded_files):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_file_path = tmp_file.name
            
        with st.spinner(f"Processing {uploaded_file.name}..."):
            gemini_file = genai.upload_file(path=tmp_file_path, display_name=f"PDF for Extraction: {uploaded_file.name}")
            response = genai.GenerativeModel(model_name="gemini-2.5-flash").generate_content([prompt, gemini_file])
            
            json_data = extract_json_from_text(response.text)
            if json_data:
                all_data.append(json_data)
                st.success(f"Extracted data from {uploaded_file.name}")
            else:
                st.warning(f"Failed to extract data from {uploaded_file.name}")
                
        os.unlink(tmp_file_path)
        progress.progress((idx + 1) / len(uploaded_files))
    
    # Update sheet with all extracted data at once
    if all_data:
        with st.spinner("Updating Google Sheet with intelligent product matching..."):
            suppliers_updated = update_google_sheet_with_multiple_files(worksheet, all_data)
            st.success(f"Updated Google Sheet with data from {suppliers_updated} supplier(s)")
            
            match_stats = {
                "total_products": sum(len(json_data["products"]) for json_data in all_data[1:]) if len(all_data) > 1 else 0,
                "matched_products": 0,
                "new_products": 0
            }
            
            if len(all_data) > 1:
                for data in all_data[1:]:
                    for product in data["products"]:
                        if st.session_state.get("matched_products", set()):
                            match_stats["matched_products"] += 1
                        else:
                            match_stats["new_products"] += 1
                
                st.info(f"Product matching statistics: {match_stats['matched_products']} products matched, {match_stats['new_products']} new products added")
        
        merged_data = {"extractions": all_data}
        json_str = json.dumps(merged_data, indent=2, ensure_ascii=False)
        
        st.download_button(
            label="Download Extracted JSON File",
            data=json_str,
            file_name="extracted_data.json",
            mime="application/json"
        )
        
        st.subheader("Extracted Data")
        st.json(json_str)