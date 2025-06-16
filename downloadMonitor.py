import os
import time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from openpyxl import load_workbook
import zipfile
import xml.etree.ElementTree as ET
import shutil
import time


class DownloadHandler(FileSystemEventHandler):
    def on_created(self, event):
        MODEL_CONFIGS = {
            "GOVERNANCE MODEL": {"prefix": "gov_"},
            "TAXONOMY APP MODEL": {"prefix": "tax_"},
            "WORKFLOW APP MODEL": {"prefix": "wfm_"},
            "INSTANCE DATA MODEL": {"prefix": "ins_"},
            "AUTHORIZATION MODEL": {"prefix": "auth_"},
            "KNOWLEDGE DATA MODEL": {"prefix": "kbm_"},
            "RS EXCEL": {"prefix": "data_"},
            "thing": {"prefix": "thg_"},
            "referenceData": {"prefix": "ref_"},
            "UOMData": {"prefix": "uom_"},
            "digitalAsset": {"prefix": "dam_"},
        }
        try:
            # Only process files, not directories
            if not event.is_directory:
                file_path = Path(event.src_path)
                file_extension = file_path.suffix.lower()
                already_has_prefix=has_known_prefix(file_path.name,MODEL_CONFIGS)
                if  file_extension in [".xlsx",".xlsm"] and not already_has_prefix:
                    wb=load_workbook(file_path)
                    ws=wb.active
                    template_name=str(ws['b4'].value)
                    domain_name=str(ws['b8'].value)
                    model_name=""
                    prefix=""
                    new_file_path=""
                    base_model=False

                    if template_name in MODEL_CONFIGS:
                        model_name=template_name
                    elif domain_name in MODEL_CONFIGS:
                        model_name=domain_name
                        base_model=True

                    print(f"New file detected: {file_path.name}")

                    if model_name!="":
                        prefix=MODEL_CONFIGS[model_name]["prefix"]
                        new_file_path=file_path.parent / (prefix+str(file_path.name ))
                        if new_file_path.exists():
                            new_file_path=add_suffix_to_filename(new_file_path, str(time.time())[-6:])
                        file_path.rename(new_file_path)                        
                        print(f"Template name: {template_name}")
                        if base_model:
                            print(f"Domain name: {domain_name}")
                            # disable_window_protection_in_sheetview(new_file_path)
                        print(f"Prefix: {prefix}")
                        print(f"New filename: {new_file_path.name}")
                        

                    print("-" * 50)    

        except Exception as e:
            print("some file caused some error")
            print(e)

def disable_window_protection_in_sheetview(xlsx_path):
    xlsx_path = Path(xlsx_path)
    temp_dir = Path("temp_unzip")
    if temp_dir.exists():
        shutil.rmtree(temp_dir)

    # Register namespace globally
    ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    # Step 1: Unzip workbook
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Step 2: Modify sheet XML files
    sheet_dir = temp_dir / 'xl' / 'worksheets'
    for fpath in sheet_dir.glob('sheet*.xml'):
        tree = ET.parse(fpath)
        root = tree.getroot()

        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        sheet_view = root.find('./main:sheetViews/main:sheetView', ns)
        if sheet_view is not None and 'windowProtection' in sheet_view.attrib:
            sheet_view.set('windowProtection', '0')
            print(f"✅ Removed windowProtection in {fpath.name}")
            tree.write(fpath, encoding='utf-8', xml_declaration=True)

    # Step 3: Repackage into a new .xlsx file
    new_file = xlsx_path.with_name(f"{xlsx_path.stem}_unprotected.xlsx")
    with zipfile.ZipFile(new_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for foldername, _, filenames in os.walk(temp_dir):
            for filename in filenames:
                full_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(full_path, temp_dir)
                zf.write(full_path, arcname)

    # Step 4: Cleanup
    shutil.rmtree(temp_dir)  
    print(f"🎉 Window protection removed. File saved as: {new_file}")



def has_known_prefix(file_name, model_configs) -> bool:
    """
    Returns True if the file name contains any prefix from the model_configs dictionary.
    
    Parameters:
    - file_name (str): The name of the file (e.g., file_path.name).
    - model_configs (dict): Dictionary with model names and their prefix settings.
    
    Returns:
    - bool: True if any prefix is found in the file name, False otherwise.
    """
    file_name_string=str(file_name)
    return any(cfg["prefix"] in file_name_string for cfg in model_configs.values())

def get_downloads_directory():
    """Get the user's downloads directory"""
    # Common download directory paths
    downloads_paths = [
        Path.home() / "Downloads",
        Path.home() / "Download",
        Path.home() / "downloads"
    ]
    
    for path in downloads_paths:
        if path.exists():
            return str(path)
    
    # Fallback: ask user to specify
    custom_path = input("Please enter your downloads directory path: ")
    return custom_path

def add_suffix_to_filename(path: Path, suffix: str = "1") -> Path:
    return path.with_name(f"{path.stem}{suffix}{path.suffix}")

def main():
    downloads_dir = get_downloads_directory()
    
    if not os.path.exists(downloads_dir):
        print(f"Directory {downloads_dir} does not exist!")
        return
    
    print(f"Monitoring downloads directory: {downloads_dir}")
    print("Press Ctrl+C to stop monitoring...")
    
    # Create event handler and observer
    event_handler = DownloadHandler()
    observer = Observer()
    observer.schedule(event_handler, downloads_dir, recursive=False)
    
    # Start monitoring
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping monitor...")
        observer.stop()

    observer.join()
    print("Monitor stopped.")

if __name__ == "__main__":
    main()