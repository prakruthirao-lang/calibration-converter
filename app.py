import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
from pdf2image import convert_from_bytes
import io
import re

st.set_page_config(page_title="Direct Image/PDF to Excel Converter", layout="wide")
st.title("📄 Direct Image/PDF Calibration Table Converter")

st.write("Upload your screenshot or PDF directly to extract calibration tables into standard sequential **MM** (0, 1, 2, 3...) and **LTRS** columns.")

uploaded_file = st.file_uploader("Upload Image or PDF File", type=["png", "jpg", "jpeg", "pdf"])

def extract_text_from_file(file):
    text_lines = []
    if file.name.lower().endswith('.pdf'):
        images = convert_from_bytes(file.read())
        for img in images:
            text = pytesseract.image_to_string(img)
            text_lines.extend(text.splitlines())
    else:
        img = Image.open(file)
        text = pytesseract.image_to_string(img)
        text_lines.extend(text.splitlines())
    return text_lines

def parse_ocr_text_pairs(lines):
    records = []
    for line in lines:
        numbers = re.findall(r"[-+]?\d*\.\d+|\d+", line)
        if len(numbers) >= 2:
            for i in range(0, len(numbers) - 1, 2):
                try:
                    mm = float(numbers[i])
                    ltrs = float(numbers[i+1])
                    records.append({'MM': int(mm), 'LTRS': ltrs})
                except ValueError:
                    continue
    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values('MM').drop_duplicates(subset=['MM']).reset_index(drop=True)
        df['MM'] = df['MM'].astype(int)
    return df

def parse_ocr_text_grid(lines):
    records = []
    for line in lines:
        numbers = re.findall(r"[-+]?\d*\.\d+|\d+", line)
        # Needs base MM + at least one value
        if len(numbers) >= 2:
            try:
                base_mm = float(numbers[0])
                # Columns 0 to 9 offset (up to 10 entries after base MM)
                for offset, val in enumerate(numbers[1:11]):
                    ltrs = float(val)
                    records.append({'MM': int(base_mm) + offset, 'LTRS': ltrs})
            except ValueError:
                continue
    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values('MM').drop_duplicates(subset=['MM']).reset_index(drop=True)
        df['MM'] = df['MM'].astype(int)
    return df

if uploaded_file is not None:
    st.info("Extracting text via OCR... Please wait.")
    lines = extract_text_from_file(uploaded_file)
    
    chart_mode = st.radio(
        "Select Table Format:", 
        ["Horizontal Matrix Grid (MM in Left Col, 0-9 Headers)", "Side-by-Side Pairs (MM | LTRS)"]
    )
    
    if chart_mode.startswith("Horizontal Matrix"):
        clean_df = parse_ocr_text_grid(lines)
    else:
        clean_df = parse_ocr_text_pairs(lines)
        
    st.write("### Converted Data Output", clean_df)
    
    if not clean_df.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            clean_df.to_excel(writer, index=False, sheet_name='Calibration_Data')
            
        st.download_button(
            label="📥 Download Standardized Excel File",
            data=output.getvalue(),
            file_name="converted_calibration.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
