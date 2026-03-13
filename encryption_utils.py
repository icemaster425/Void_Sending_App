import pyminizip
import os
import time
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
import win32com.client as win32
import pythoncom
from PIL import Image

def check_file_integrity(file_path):
    file_path = str(file_path)
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        return False, "File is missing or 0 bytes."
    initial_size = os.path.getsize(file_path)
    time.sleep(2)
    if os.path.getsize(file_path) != initial_size:
        return False, "File is still being copied."
    try:
        with open(file_path, 'a'): pass
    except PermissionError:
        return False, "File is locked."
    return True, "Success"

def split_pdf_pages(file_path, output_dir):
    file_path, output_dir = str(file_path), str(output_dir)
    reader = PdfReader(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    generated_files = []
    for i, page in enumerate(reader.pages):
        writer = PdfWriter()
        writer.add_page(page)
        output_filename = f"{base_name}_Page_{i+1}.pdf"
        output_path = os.path.join(output_dir, output_filename)
        with open(output_path, "wb") as f:
            writer.write(f)
        generated_files.append(str(output_path))
    return generated_files

def convert_pdf_to_tiff(file_path, output_dir):
    file_path, output_dir = str(file_path), str(output_dir)
    doc = fitz.open(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.tiff")
    
    image_frames = []
    
    # Extract all pages into memory
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        mode = "RGBA" if pix.alpha else "RGB"
        
        # THE FIX: Pillow strictly requires a tuple (), not a list [] for size parameters.
        img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
        
        # Force uniform RGB mode to prevent Pillow from aborting the stitch
        img = img.convert("RGB")
        image_frames.append(img)
        
    # Stitch and save as a single multi-page TIFF
    if image_frames:
        image_frames[0].save(
            output_path, 
            format="TIFF", 
            save_all=True, 
            append_images=image_frames[1:]
        )
        
    doc.close()
    return [str(output_path)]

def transform_excel(file_path, output_dir, recipes):
    file_path, output_dir = str(file_path), str(output_dir)
    filename = os.path.basename(file_path)
    orig_ext = os.path.splitext(filename)[1].lower()
    
    # Load data with strict string typing to prevent leading zeroes from dropping
    if orig_ext == '.csv':
        df = pd.read_csv(file_path, dtype=str)
    elif orig_ext == '.xls':
        df = pd.read_excel(file_path, engine='xlrd', dtype=str)
    else:
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)

    record_count = len(df)

    # --- DATA RECIPES ---
    if 'trim_whitespace' in recipes:
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        
    if 'remove_duplicates' in recipes:
        df.drop_duplicates(inplace=True)
        
    if 'mask_id' in recipes:
        for col in [c for c in df.columns if 'id' in c.lower()]:
            df[col] = df[col].astype(str).apply(lambda x: x[:2] + '*' * (len(x)-4) + x[-2:] if len(x) > 4 else "****")

    # Column A Splitters (BSB and Farm)
    if 'bsb_split' in recipes or 'farm_split' in recipes:
        orig_col = df.columns[0] # Target Column A
        
        # Pop the column completely out of the dataframe first to kill naming collisions
        orig_data = df.pop(orig_col).astype(str).str.replace(r'\.0$', '', regex=True)
        
        if 'bsb_split' in recipes:
            if 'BSB' in df.columns: df.drop(columns=['BSB'], inplace=True)
            if 'Account Number' in df.columns: df.drop(columns=['Account Number'], inplace=True)
            
            df.insert(0, 'BSB', orig_data.str[:6])
            df.insert(1, 'Account Number', orig_data.str[6:])
            
        elif 'farm_split' in recipes:
            if 'Farm Number' in df.columns: df.drop(columns=['Farm Number'], inplace=True)
            if 'Party Number' in df.columns: df.drop(columns=['Party Number'], inplace=True)
            
            df.insert(0, 'Farm Number', orig_data.str[:5])
            df.insert(1, 'Party Number', orig_data.str[5:])

    # --- EXTENSION ROUTING ---
    new_ext = orig_ext
    if 'xls_to_xlsx' in recipes: new_ext = '.xlsx'
    elif 'xlsx_to_xls' in recipes: new_ext = '.xls'
    elif 'xlsx_to_csv' in recipes or 'xls_to_csv' in recipes: new_ext = '.csv'
        
    new_filename = os.path.splitext(filename)[0] + new_ext
    save_path = os.path.join(output_dir, new_filename)
    
    try:
        if new_ext == ".xls":
            temp_xlsx = save_path + "x"
            df.to_excel(temp_xlsx, index=False, engine='openpyxl')
            
            pythoncom.CoInitialize()
            try:
                excel = win32.Dispatch('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(os.path.abspath(temp_xlsx))
                wb.SaveAs(os.path.abspath(save_path), FileFormat=56) 
                wb.Close()
            finally:
                if 'excel' in locals():
                    excel.Quit()
                pythoncom.CoUninitialize()
            
            if os.path.exists(temp_xlsx):
                os.remove(temp_xlsx)
        elif new_ext == ".csv":
            df.to_csv(save_path, index=False)
        else:
            df.to_excel(save_path, index=False, engine='openpyxl')
            
    except Exception as e:
        fallback_path = os.path.splitext(save_path)[0] + ".xlsx"
        df.to_excel(fallback_path, index=False, engine='openpyxl')
        return str(fallback_path), record_count
        
    return str(save_path), record_count

def zip_files_with_password(file_paths, zip_path, password, batch_name=""):
    if not file_paths: return zip_path # Fail-safe against empty lists crashing the C-extension
    
    # Bulletproof list flattening
    flat_paths = []
    for f in file_paths:
        if isinstance(f, list): flat_paths.extend([str(x) for x in f])
        else: flat_paths.append(str(f))
        
    zip_path, password = str(zip_path), str(password)
    prefixes = ["" for _ in flat_paths]
    pyminizip.compress_multiple(flat_paths, prefixes, zip_path, password, 5)
    return zip_path