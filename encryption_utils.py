import pyminizip
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF for high-quality, standalone PDF to Image conversion

def check_file_integrity(file_path):
    """
    STEP ZERO: Validates that files aren't corrupted or locked by another user.
    Returns: (bool, message)
    """
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        return False, "File is missing or 0 bytes."

    ext = os.path.splitext(file_path)[1].lower()

    # 1. OS-Level Lock Check (Checks if the file is open in Excel/Acrobat)
    try:
        # Trying to open in append mode temporarily. Fails if locked by another user.
        with open(file_path, 'a'):
            pass
    except PermissionError:
        return False, "File is locked. Is it currently open in another program?"
    except Exception as e:
        return False, f"OS Lock Check Failed: {str(e)}"

    # 2. Deep PDF Health Check
    if ext == '.pdf':
        try:
            reader = PdfReader(file_path)
            if reader.is_encrypted:
                return False, "PDF is password protected. V.O.I.D. cannot process locked files."
            
            # Attempt to read the first page to confirm the file is not deeply corrupted
            _ = reader.pages[0]
            return True, "Healthy"
        except Exception as e:
            return False, f"PDF File Corrupted: {str(e)}"

    # 3. Deep Excel Health Check
    if ext in ['.xls', '.xlsx']:
        try:
            engine = 'xlrd' if ext == '.xls' else 'openpyxl'
            # Just read the first 2 rows to ensure the headers aren't corrupted
            pd.read_excel(file_path, engine=engine, nrows=2)
            return True, "Healthy"
        except Exception as e:
            return False, f"Excel File Corrupted: {str(e)}"
            
    return True, "Healthy (No deep scan for this extension)"


def split_pdf_pages(file_path, output_dir):
    """
    Recipe: Breaks a multi-page PDF into individual single-page files.
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


def convert_pdf_to_tiff(file_path, output_dir):
    """
    Recipe: Renders a PDF into high-resolution TIFF images for OCR banking systems.
    """
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    tiff_files = []
    
    # Open the document using PyMuPDF
    doc = fitz.open(file_path)
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        
        # Matrix to increase resolution (Zoom x2 is roughly 144 DPI)
        zoom_matrix = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=zoom_matrix, alpha=False)
        
        output_filename = f"{base_name}_img_{page_num+1}.tiff"
        output_path = os.path.join(output_dir, output_filename)
        
        # Save directly as image
        pix.pil_save(output_path, format="TIFF")
        tiff_files.append(output_path)
        
    doc.close()
    return tiff_files


def transform_excel(file_path, recipes, output_dir):
    """
    Recipe Engine: Handles BSB/Farm splitting, CSV exports, and format swaps.
    Returns: (new_file_path, record_count)
    """
    filename = os.path.basename(file_path)
    orig_ext = os.path.splitext(file_path)[1].lower()
    
    # 1. Load Data
    engine = 'xlrd' if orig_ext == '.xls' else 'openpyxl'
    df = pd.read_excel(file_path, engine=engine)

    # Calculate exact rows (excluding headers)
    record_count = len(df)

    # 2. Apply Column Splits if requested
    if 'bsb_split' in recipes:
        original_col = df.columns[0]
        col_data = df[original_col].astype(str)
        
        df.insert(0, 'BSB', col_data.str[:6])
        df.insert(1, '.Account Number', col_data.str[6:])
        df.drop(columns=[original_col], inplace=True)

    if 'farm_split' in recipes:
        original_col = df.columns[0]
        col_data = df[original_col].astype(str)
        
        # Slices the first 5 digits for the Farm, the rest for the Party
        df.insert(0, 'Farm Number', col_data.str[:5])
        df.insert(1, 'Party Number', col_data.str[5:])
        df.drop(columns=[original_col], inplace=True)

    # 3. Determine Output Format
    new_ext = orig_ext  # Default to original unless asked to swap
    
    if 'xls_to_xlsx' in recipes:
        new_ext = '.xlsx'
    elif 'xlsx_to_xls' in recipes:
        new_ext = '.xls'
    elif 'xlsx_to_csv' in recipes or 'xls_to_csv' in recipes:
        new_ext = '.csv'
        
    new_filename = os.path.splitext(filename)[0] + "_processed" + new_ext
    save_path = os.path.join(output_dir, new_filename)
    
    # 4. Save Transformed File
    if new_ext == ".xls":
        df.to_excel(save_path, index=False, engine='xlwt')
    elif new_ext == ".csv":
        df.to_csv(save_path, index=False)
    else:
        df.to_excel(save_path, index=False, engine='openpyxl')
        
    return save_path, record_count


def zip_files_with_password(file_paths, zip_path, password, batch_name=""):
    """
    The Vault: Compresses and encrypts files.
    """
    prefixes = ["" for _ in file_paths]
    compression_level = 0  # Standard compression
    
    try:
        for p in file_paths:
            if not os.path.exists(p):
                raise FileNotFoundError(f"Missing file for zipping: {os.path.basename(p)}")

        pyminizip.compress_multiple(file_paths, prefixes, zip_path, password, compression_level)
        return True
    except Exception as e:
        raise Exception(f"Encryption Error: {str(e)}")