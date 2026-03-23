import os
import time
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
import win32com.client as win32
import pythoncom
from PIL import Image
import pyzipper

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

def remove_first_and_split_pdf(file_path, output_dir):
    """Specific Recipe for Rabo: Skip page 1, split the rest."""
    file_path, output_dir = str(file_path), str(output_dir)
    reader = PdfReader(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    generated_files = []
    
    if len(reader.pages) <= 1:
        return [] 
        
    for i in range(1, len(reader.pages)):
        writer = PdfWriter()
        writer.add_page(reader.pages[i])
        output_filename = f"{base_name}_Page_{i+1}.pdf"
        output_path = os.path.join(output_dir, output_filename)
        with open(output_path, "wb") as f:
            writer.write(f)
        generated_files.append(str(output_path))
    return generated_files

def has_myob_id(file_path):
    """Rapid UI scanner to verify MYOB_ID exists without loading full data."""
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv':
            df = pd.read_csv(file_path, nrows=0)
        elif ext == '.xls':
            df = pd.read_excel(file_path, engine='xlrd', nrows=0)
        else:
            df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
            
        cols = [str(c).strip().upper() for c in df.columns]
        return 'MYOB_ID' in cols
    except Exception:
        return False

def convert_pdf_to_tiff(file_path, output_dir):
    file_path, output_dir = str(file_path), str(output_dir)
    doc = fitz.open(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.tiff")
    
    temp_files = []
    
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        temp_path = os.path.join(output_dir, f"{base_name}_temp_{i}.png")
        pix.save(temp_path)
        temp_files.append(temp_path)
        
    doc.close()
    
    if temp_files:
        images = []
        for t_file in temp_files:
            img = Image.open(t_file)
            img = img.convert("L")
            images.append(img)
            
        images[0].save(
            output_path, 
            format="TIFF", 
            save_all=True, 
            append_images=images[1:],
            compression="tiff_lzw" 
        )
        
        for img in images: img.close()
        for t_file in temp_files:
            if os.path.exists(t_file): os.remove(t_file)
                
    return [str(output_path)]

def transform_excel(file_path, output_dir, recipes):
    file_path, output_dir = str(file_path), str(output_dir)
    filename = os.path.basename(file_path)
    orig_ext = os.path.splitext(filename)[1].lower()
    
    if orig_ext == '.csv':
        df = pd.read_csv(file_path, dtype=str)
    elif orig_ext == '.xls':
        df = pd.read_excel(file_path, engine='xlrd', dtype=str)
    else:
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)

    record_count = len(df)

    if 'trim_whitespace' in recipes:
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        
    if 'remove_duplicates' in recipes:
        df.drop_duplicates(inplace=True)
        
    if 'mask_id' in recipes:
        for col in [c for c in df.columns if 'id' in c.lower()]:
            df[col] = df[col].astype(str).apply(lambda x: x[:2] + '*' * (len(x)-4) + x[-2:] if len(x) > 4 else "****")

    if 'bsb_split' in recipes or 'farm_split' in recipes:
        orig_col = df.columns[0] 
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
                if 'excel' in locals(): excel.Quit()
                pythoncom.CoUninitialize()
            
            if os.path.exists(temp_xlsx): os.remove(temp_xlsx)
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
    if not file_paths: return zip_path
    
    flat_paths = []
    for f in file_paths:
        if isinstance(f, list): flat_paths.extend([str(x) for x in f])
        else: flat_paths.append(str(f))
        
    zip_path = str(zip_path)
    password = str(password).strip() if password else ""
    
    if not password or password.lower() == "none":
        with pyzipper.AESZipFile(zip_path, 'w', compression=pyzipper.ZIP_DEFLATED) as zf:
            for f in flat_paths:
                if os.path.exists(f):
                    zf.write(f, os.path.basename(f))
    else:
        with pyzipper.AESZipFile(zip_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_ZIPCRYPTO) as zf:
            zf.setpassword(password.encode('utf-8'))
            for f in flat_paths:
                if os.path.exists(f):
                    zf.write(f, os.path.basename(f))
                    
    return zip_path