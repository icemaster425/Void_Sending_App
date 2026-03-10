import pyminizip
import os
import time
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF for high-quality, standalone PDF to Image conversion

def check_file_integrity(file_path):
    """
    Validates that files aren't corrupted, locked, or still copying over the network.
    Returns: (bool, message)
    """
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        return False, "File is missing or 0 bytes."

    # 1. Stable Size Check (Network Copy Protection)
    initial_size = os.path.getsize(file_path)
    time.sleep(2)  # Wait 2 seconds
    if os.path.getsize(file_path) != initial_size:
        return False, "File is still being copied into the folder. Wait."

    ext = os.path.splitext(file_path)[1].lower()

    # 2. OS-Level Lock Check 
    try:
        with open(file_path, 'a'):
            pass
    except PermissionError:
        return False, "File is locked. Is it currently open in another program?"
    except Exception as e:
        return False, f"OS Lock Check Failed: {str(e)}"

    # 3. Deep PDF Health Check
    if ext == '.pdf':
        try:
            reader = PdfReader(file_path)
            if reader.is_encrypted:
                return False, "PDF is password protected. V.O.I.D. cannot process locked files."
            _ = reader.pages[0]
            return True, "Healthy"
        except Exception as e:
            return False, f"PDF File Corrupted: {str(e)}"

    # 4. Deep Excel Health Check
    if ext in ['.xls', '.xlsx']:
        try:
            engine = 'xlrd' if ext == '.xls' else 'openpyxl'
            pd.read_excel(file_path, engine=engine, nrows=2)
            return True, "Healthy"
        except Exception as e:
            return False, f"Excel File Corrupted: {str(e)}"
            
    return True, "Healthy (No deep scan for this extension)"


def split_pdf_pages(file_path, output_dir):
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


def convert_pdf_to_tiff(file_path, output_dir):
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    tiff_files = []
    
    doc = fitz.open(file_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        zoom_matrix = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=zoom_matrix, alpha=False)
        
        output_filename = f"{base_name}_img_{page_num+1}.tiff"
        output_path = os.path.join(output_dir, output_filename)
        
        pix.pil_save(output_path, format="TIFF")
        tiff_files.append(output_path)
        
    doc.close()
    return tiff_files


def transform_excel(file_path, recipes, output_dir):
    filename = os.path.basename(file_path)
    orig_ext = os.path.splitext(file_path)[1].lower()
    
    engine = 'xlrd' if orig_ext == '.xls' else 'openpyxl'
    df = pd.read_excel(file_path, engine=engine)
    record_count = len(df)

    if 'bsb_split' in recipes:
        original_col = df.columns[0]
        col_data = df[original_col].astype(str)
        df.insert(0, 'BSB', col_data.str[:6])
        df.insert(1, '.Account Number', col_data.str[6:])
        df.drop(columns=[original_col], inplace=True)

    if 'farm_split' in recipes:
        original_col = df.columns[0]
        col_data = df[original_col].astype(str)
        df.insert(0, 'Farm Number', col_data.str[:5])
        df.insert(1, 'Party Number', col_data.str[5:])
        df.drop(columns=[original_col], inplace=True)

    new_ext = orig_ext
    if 'xls_to_xlsx' in recipes:
        new_ext = '.xlsx'
    elif 'xlsx_to_xls' in recipes:
        new_ext = '.xls'
    elif 'xlsx_to_csv' in recipes or 'xls_to_csv' in recipes:
        new_ext = '.csv'
        
    new_filename = os.path.splitext(filename)[0] + "_processed" + new_ext
    save_path = os.path.join(output_dir, new_filename)
    
    # Safe Save Block specifically handling PyInstaller ghosting
    try:
        if new_ext == ".xls":
            import xlwt  # Forced local import
            df.to_excel(save_path, index=False, engine='xlwt')
        elif new_ext == ".csv":
            df.to_csv(save_path, index=False)
        else:
            df.to_excel(save_path, index=False, engine='openpyxl')
    except ImportError as e:
        df.to_excel(save_path, index=False) # Fallback
        
    return save_path, record_count


def zip_files_with_password(file_paths, zip_path, password, batch_name=""):
    prefixes = ["" for _ in file_paths]
    compression_level = 0
    
    try:
        for p in file_paths:
            if not os.path.exists(p):
                raise FileNotFoundError(f"Missing file for zipping: {os.path.basename(p)}")

        pyminizip.compress_multiple(file_paths, prefixes, zip_path, password, compression_level)
        return True
    except Exception as e:
        raise Exception(f"Encryption Error: {str(e)}")