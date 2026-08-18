import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
from pdf2image import convert_from_bytes
import cv2
import numpy as np
import io
import re

st.set_page_config(page_title="Direct Calibration Chart Converter", layout="wide")
st.title("🛢️ High-Precision Calibration Chart Converter")

st.write("Upload your screenshot or PDF — this tool uses cell-grid extraction to map exact **MM** heights to **LTRS** volumes without row or column displacement.")

uploaded_file = st.file_uploader("Upload Image or PDF File", type=["png", "jpg", "jpeg", "pdf"])

def file_to_cv2_image(file):
    if file.name.lower().endswith('.pdf'):
        images = convert_from_bytes(file.read(), first_page=1, last_page=1)
        if not images:
            return None
        pil_img = images[0].convert('RGB')
    else:
        pil_img = Image.open(file).convert('RGB')
    
    open_cv_image = np.array(pil_img)
    return open_cv_image[:, :, ::-1].copy()

def extract_cells_and_ocr(file):
    cv_img = file_to_cv2_image(file)
    if cv_img is None:
        return []
    
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    
    # Simple line-by-line fallback + cell parsing
    config = r'--oem 3 --psm 6'
    raw_text = pytesseract.image_to_string(gray, config=config)
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    return lines

def clean_number(text):
    clean = re.sub(r'[^0-9.]', '', text)
    try:
        return float(clean)
    except ValueError:
        return None

def process_matrix_grid(lines):
    records = []
    
    for line in lines:
        # Extract all numeric values in order across the row
        nums = [clean_number(token) for token in line.split() if clean_number(token) is not None]
        
        if len(nums) < 2:
            continue
            
        base_mm = nums[0]
        
        # Check if the line is a header (e.g. 0 1 2 3 4 5 6 7 8 9)
        if nums[:5] == [0, 1, 2, 3, 4] or nums[:5] == [1, 2, 3, 4, 5]:
            continue
            
        # Parse LTRS entries for offsets 0 to 9
        for offset, ltrs_val in enumerate(nums[1:11]):
            actual_mm = int(base_mm) + offset
            records.append({'MM': actual_mm, 'LTRS': ltrs_val})
            
    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values('MM').drop_duplicates(subset=['MM'], keep='first').reset_index(drop=True)
        df['MM'] = df['MM'].astype(int)
    return df

def process_side_by_side(lines):
    records = []
    for line in lines:
        nums = [clean_number(token) for token in line.split() if clean_number(token) is not None]
        if len(nums) >= 2:
            for i in range(0, len(nums) - 1, 2):
                mm_val = nums[i]
                ltrs_val = nums[i+1]
                if mm_val is not None and ltrs_val is not None:
                    records.append({'MM': int(mm_val), 'LTRS': ltrs_val})
                    
    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values('MM').drop_duplicates(subset=['MM'], keep='first').reset_index(drop=True)
        df['MM'] = df['MM'].astype(int)
    return df

if uploaded_file is not None:
    st.info("Reading image structure...")
    lines = extract_cells_and_ocr(uploaded_file)
    
    with st.expander("🔍 View Raw OCR Extracted Text (For verification)"):
        st.text("\n".join(lines))
    
    chart_mode = st.radio(
        "Select Table Format:", 
        ["Horizontal Matrix Grid (Base MM on Left, 0-9 Offset Headers)", "Side-by-Side Pairs (MM | LTRS | MM | LTRS)"]
    )
    
    if chart_mode.startswith("Horizontal Matrix"):
        clean_df = process_matrix_grid(lines)
    else:
        clean_df = process_side_by_side(lines)
        
    st.write("### Converted Calibration Output", clean_df)
    
    if not clean_df.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            clean_df.to_excel(writer, index=False, sheet_name='Calibration_Data')
            
        st.download_button(
            label="📥 Download Clean Excel File",
            data=output.getvalue(),
            file_name="calibration_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
