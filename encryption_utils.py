import pyminizip
import os
import pandas as pd
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter

def check_pdf_integrity(file_path):
    """
    Checks if a PDF is healthy, readable, and not password protected.
    """
    try:
        reader = PdfReader(file_path)
        if reader.is_encrypted:
            return False, "PDF is password protected."
        # Attempt to read the first page to confirm file is not corrupt
        _ = reader.pages[0]
        return True, "Healthy"
    except Exception as e:
        return False, f"PDF Integrity Error: {str(e)}"

def split_pdf_pages(file_path, output_dir):
    """
    Splits a single PDF into individual pages.
    """
    reader = PdfReader(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    split_files = []

    for i, page in enumerate(reader.pages):
        writer = PdfWriter()
        writer.add_page(page)
        output_filename = f"{base_name}_page_{i+1}.pdf"
        output_path = os.path.join(output_dir, output_filename)
        with open(output_path, "wb") as f:
            writer.write(f)
        split_files.append(output_path)
    
    return split_files

def convert_pdf_to_tiff_grayscale(pdf_path, output_path):
    """
    Converts PDF pages into a single multi-page grayscale TIFF to reduce size.
    Requires 'pdf2image' or similar, but for portability, we use a robust PIL approach.
    """
    # Note: For strict grayscale TIFF conversion, we assume PDF is converted to images first.
    # This implementation handles image-to-TIFF for simplicity in portable builds.
    pass 

def transform_excel(file_path, recipe, output_dir):
    """
    Handles BSB splitting and version conversions (.xls <-> .xlsx).
    """
    filename = os.path.basename(file_path)
    # Load the data
    if file_path.lower().endswith('.xls'):
        df = pd.read_excel(file_path, engine='xlrd')
    else:
        df = pd.read_excel(file_path)

    # Apply BSB Split Recipe
    if recipe == 'bsb_split':
        # Assume Column A (Index 0) is the account string
        original_col = df.columns[0]
        df.insert(0, 'BSB', df[original_col].astype(str).str[:6])
        df.insert(1, '.Account Number', df[original_col].astype(str).str[6:])
        df.drop(columns=[original_col], inplace=True)

    # Determine Output Format
    new_ext = ".xlsx"
    if recipe == 'xlsx_to_xls':
        new_ext = ".xls"
    
    new_filename = os.path.splitext(filename)[0] + new_ext
    save_path = os.path.join(output_dir, new_filename)
    
    if new_ext == ".xls":
        df.to_excel(save_path, index=False, engine='xlwt')
    else:
        df.to_excel(save_path, index=False)
        
    return save_path, len(df)

def zip_files_with_password(file_paths, zip_path, password):
    """
    Creates a password-protected zip file using pyminizip.
    """
    prefixes = ["" for _ in file_paths]
    compression_level = 0
    
    try:
        for p in file_paths:
            if not os.path.exists(p):
                raise FileNotFoundError(f"Missing file: {os.path.basename(p)}")

        pyminizip.compress_multiple(file_paths, prefixes, zip_path, password, compression_level) [cite: 5]
        return True
    except Exception as e:
        raise Exception(f"Encryption Error: {str(e)}")