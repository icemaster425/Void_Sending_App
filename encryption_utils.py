import pyminizip
import os
import time
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
import win32com.client as win32
import pythoncom

def check_file_integrity(file_path):
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
        generated_files.append(output_path)
    return generated_files

def convert_pdf_to_tiff(file_path, output_dir):
    doc = fitz.open(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.tiff")
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
    pix.save(output_path)
    doc.close()
    return [output_path]

def transform_excel(file_path, output_dir, recipes):
    filename = os.path.basename(file_path)
    orig_ext = os.path.splitext(filename)[1].lower()
    
    if orig_ext == '.csv':
        df = pd.read_csv(file_path)
    elif orig_ext == '.xls':
        df = pd.read_excel(file_path, engine='xlrd')
    else:
        df = pd.read_excel(file_path, engine='openpyxl')

    record_count = len(df)

    # Apply Transformation Recipes
    if 'trim_whitespace' in recipes:
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    if 'remove_duplicates' in recipes:
        df.drop_duplicates(inplace=True)
    if 'mask_id' in recipes:
        for col in [c for c in df.columns if 'id' in c.lower()]:
            df[col] = df[col].astype(str).apply(lambda x: x[:2] + '*' * (len(x)-4) + x[-2:] if len(x) > 4 else "****")

    # Hard-lock the default output to .xls
    new_ext = '.xls'
    
    # Only override if a specific recipe demands an alternative format
    if 'xls_to_xlsx' in recipes: new_ext = '.xlsx'
    elif 'xlsx_to_csv' in recipes or 'xls_to_csv' in recipes: new_ext = '.csv'
        
    # Clean, identical filename without the _processed tag
    new_filename = os.path.splitext(filename)[0] + new_ext
    save_path = os.path.join(output_dir, new_filename)
    
    # THE OUTSIDE-THE-BOX FIX: Excel Handshake
    try:
        if new_ext == ".xls":
            # Save as modern first
            temp_xlsx = save_path + "x"
            df.to_excel(temp_xlsx, index=False, engine='openpyxl')
            
            # Use System Excel to convert
            pythoncom.CoInitialize()
            try:
                excel = win32.client.Dispatch('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(os.path.abspath(temp_xlsx))
                wb.SaveAs(os.path.abspath(save_path), FileFormat=56) # 56 = Legacy XLS
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
        # Fallback to standard save
        df.to_excel(save_path, index=False)
        
    return save_path, record_count

def zip_files_with_password(file_paths, zip_path, password, batch_name=""):
    prefixes = ["" for _ in file_paths]
    pyminizip.compress_multiple(file_paths, prefixes, zip_path, password, 5)
    return zip_path