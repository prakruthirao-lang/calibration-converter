import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
from pdf2image import convert_from_bytes
import cv2
import numpy as np
import io
import re

st.set_page_config(page_title="High-Precision Calibration Converter", layout="wide")
st.title("🛢️ High-Precision Calibration Chart Converter")

uploaded_file = st.file_uploader("Upload Image or PDF File", type=["png", "jpg", "jpeg", "pdf"])

def file_to_pil_image(file):
    if file.name.lower().endswith('.pdf'):
        images = convert_from_bytes(file.read(), first_page=1, last_page=1)
        if not images:
            return None
        return images[0].convert('RGB')
    else:
        return Image.open(file).convert('RGB')

def clean_num(val_str):
    if not val_str:
        return None
    cleaned = re.sub(r'[^0-9.]', '', str(val_str))
    try:
        return float(cleaned)
    except ValueError:
        return None

def process_spatial_grid(pil_img):
    # OCR with detailed bounding box metadata
    ocr_df = pytesseract.image_to_data(pil_img, output_type=pytesseract.Output.DATAFRAME)
    ocr_df = ocr_df[ocr_df.text.notnull()].copy()
    ocr_df['text'] = ocr_df['text'].astype(str).str.strip()
    ocr_df = ocr_df[ocr_df['text'] != ''].copy()

    if ocr_df.empty:
        return pd.DataFrame(columns=['mm', 'Ltrs'])

    # Group words into physical rows using vertical 'top' coordinates
    ocr_df = ocr_df.sort_values('top').reset_index(drop=True)
    rows = []
    current_row = []
    last_top = None

    for _, row in ocr_df.iterrows():
        if last_top is None or abs(row['top'] - last_top) <= 12:
            current_row.append(row)
            if last_top is None:
                last_top = row['top']
        else:
            rows.append(pd.DataFrame(current_row).sort_values('left'))
            current_row = [row]
            last_top = row['top']
            
    if current_row:
        rows.append(pd.DataFrame(current_row).sort_values('left'))

    records = []
    for r_df in rows:
        tokens = [clean_num(t) for t in r_df['text'] if clean_num(t) is not None]
        if len(tokens) < 2:
            continue
            
        base_mm = tokens[0]
        # Ignore top header offset rows (0..9 or 1..10)
        if tokens[:5] == [0, 1, 2, 3, 4] or tokens[:5] == [1, 2, 3, 4, 5]:
            continue
            
        for offset, ltrs_val in enumerate(tokens[1:11]):
            records.append({'mm': int(base_mm) + offset, 'Ltrs': ltrs_val})

    df = pd.DataFrame(records)
    if df.empty:
        return pd.DataFrame(columns=['mm', 'Ltrs'])

    df = df.sort_values('mm').drop_duplicates(subset=['mm'], keep='first').reset_index(drop=True)
    return df

def process_spatial_pairs(pil_img):
    ocr_df = pytesseract.image_to_data(pil_img, output_type=pytesseract.Output.DATAFRAME)
    ocr_df = ocr_df[ocr_df.text.notnull()].copy()
    ocr_df['text'] = ocr_df['text'].astype(str).str.strip()
    ocr_df = ocr_df[ocr_df['text'] != ''].copy()

    if ocr_df.empty:
        return pd.DataFrame(columns=['mm', 'Ltrs'])

    ocr_df = ocr_df.sort_values('top').reset_index(drop=True)
    rows = []
    current_row = []
    last_top = None

    for _, row in ocr_df.iterrows():
        if last_top is None or abs(row['top'] - last_top) <= 12:
            current_row.append(row)
            if last_top is None:
                last_top = row['top']
        else:
            rows.append(pd.DataFrame(current_row).sort_values('left'))
            current_row = [row]
            last_top = row['top']
            
    if current_row:
        rows.append(pd.DataFrame(current_row).sort_values('left'))

    records = []
    for r_df in rows:
        tokens = [clean_num(t) for t in r_df['text'] if clean_num(t) is not None]
        if len(tokens) >= 2:
            for i in range(0, len(tokens) - 1, 2):
                records.append({'mm': int(tokens[i]), 'Ltrs': tokens[i+1]})

    df = pd.DataFrame(records)
    if df.empty:
        return pd.DataFrame(columns=['mm', 'Ltrs'])

    df = df.sort_values('mm').drop_duplicates(subset=['mm'], keep='first').reset_index(drop=True)
    return df

def fill_template_0_1750(extracted_df, default_max=1750):
    if extracted_df.empty:
        max_mm = default_max
    else:
        max_mm = max(default_max, int(extracted_df['mm'].max()))

    full_df = pd.DataFrame({'mm': range(0, max_mm + 1)})

    if not extracted_df.empty:
        full_df = full_df.merge(extracted_df, on='mm', how='left')
        
        # Smoothly interpolate single missing OCR cells within the active reading range
        min_detected = extracted_df['mm'].min()
        max_detected = extracted_df['mm'].max()
        
        active_mask = (full_df['mm'] >= min_detected) & (full_df['mm'] <= max_detected)
        full_df.loc[active_mask, 'Ltrs'] = full_df.loc[active_mask, 'Ltrs'].interpolate(method='linear')
    else:
        full_df['Ltrs'] = None

    return full_df

if uploaded_file is not None:
    st.info("Analyzing spatial layout...")
    pil_image = file_to_pil_image(uploaded_file)
    
    chart_mode = st.radio(
        "Select Table Format:", 
        ["Horizontal Matrix Grid (Base MM on Left, 0-9 Offset Headers)", "Side-by-Side Pairs (MM | LTRS)"]
    )
    
    if chart_mode.startswith("Horizontal Matrix"):
        extracted_df = process_spatial_grid(pil_image)
    else:
        extracted_df = process_spatial_pairs(pil_image)
        
    final_df = fill_template_0_1750(extracted_df, default_max=1750)
    
    st.write("### Standardized Output (0–1750 mm)", final_df)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Calibration_Data')
        
    st.download_button(
        label="📥 Download Excel File",
        data=output.getvalue(),
        file_name="converted_calibration.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
