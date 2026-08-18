cat << 'EOF' > app.py
import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
from pdf2image import convert_from_bytes
import io
import re

st.set_page_config(page_title="Direct Image/PDF to Excel Converter", layout="wide")
st.title("📄 Direct Image/PDF Calibration Table Converter")

st.write("Upload your screenshot or PDF directly — the app will extract table data via OCR and format it into standard **MM** and **LTRS** columns.")

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
        # Match number pairs across lines
        numbers = re.findall(r"[-+]?\d*\.\d+|\d+", line)
        if len(numbers) >= 2:
            # Group into consecutive pairs (MM, LTRS)
            for i in range(0, len(numbers) - 1, 2):
                try:
                    mm = float(numbers[i])
                    ltrs = float(numbers[i+1])
                    records.append({'MM': mm, 'LTRS': ltrs})
                except ValueError:
                    continue
    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values('MM').drop_duplicates(subset=['MM']).reset_index(drop=True)
    return df

def parse_ocr_text_grid(lines):
    records = []
    for line in lines:
        numbers = re.findall(r"[-+]?\d*\.\d+|\d+", line)
        if len(numbers) >= 2:
            base_mm = float(numbers[0])
            for idx, val in enumerate(numbers[1:10]):
                try:
                    ltrs = float(val)
                    records.append({'MM': int(base_mm) + idx, 'LTRS': ltrs})
                except ValueError:
                    continue
    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values('MM').drop_duplicates(subset=['MM']).reset_index(drop=True)
    return df

if uploaded_file is not None:
    st.info("Extracting text via OCR... Please wait.")
    lines = extract_text_from_file(uploaded_file)
    
    chart_mode = st.radio("Select Table Format:", ["Side-by-Side Pairs (MM | LTRS)", "Horizontal Matrix Grid (MM in Left Col, 1-9 Headers)"])
    
    if chart_mode.startswith("Side-by-Side"):
        clean_df = parse_ocr_text_pairs(lines)
    else:
        clean_df = parse_ocr_text_grid(lines)
        
    st.write("### Converted Data Output", clean_df)
    
    if not clean_df.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            clean_df.to_excel(writer, index=False, sheet_name='Calibration_Data')
            
        st.download_button(
            label="📥 Download Converted Excel File",
            data=output.getvalue(),
            file_name="converted_calibration.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
EOF
