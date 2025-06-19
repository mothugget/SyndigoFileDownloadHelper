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
from lxml import etree as ET
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("Warning: python-dotenv not installed. Install with: pip install python-dotenv")


class DownloadHandler(FileSystemEventHandler):
    def on_created(self, event):
        print(f"🔔 File system event detected: {event.src_path}")
        print(f"   Event type: {'Directory' if event.is_directory else 'File'}")
        print(f"   Path exists: {os.path.exists(event.src_path)}")
        new_filename = self.process_file(event)
        if new_filename:
            print(f"📝 File renamed via watchdog: {event.src_path} -> {new_filename}")
        
    def on_modified(self, event):
        print(f"📝 File modified: {event.src_path}")
        
    def on_moved(self, event):
        print(f"📦 File moved: {event.src_path} -> {event.dest_path}")
        # Process the destination file if it's a complete download
        if not event.is_directory and event.dest_path.endswith('.xlsx'):
            # Create a mock event for the destination file
            class MockEvent:
                def __init__(self, path):
                    self.src_path = path
                    self.is_directory = False
            
            new_filename = self.process_file(MockEvent(event.dest_path))
            if new_filename:
                print(f"📝 File renamed via move event: {event.dest_path} -> {new_filename}")
        
    def process_file(self, event):
        """Process a file and return the new filename if renamed, None otherwise"""
        MODEL_CONFIGS = {
            "GOVERNANCE MODEL": {
                "prefix": "gov_", 
                "filename": os.getenv('GOVERNANCE_MODEL_FILENAME', 'governance_model')
            },
            "TAXONOMY APP MODEL": {
                "prefix": "tax_", 
                "filename": os.getenv('TAXONOMY_APP_MODEL_FILENAME', 'taxonomy_model')
            },
            "WORKFLOW APP MODEL": {
                "prefix": "wfm_", 
                "filename": os.getenv('WORKFLOW_APP_MODEL_FILENAME', 'workflow_model')
            },
            "INSTANCE DATA MODEL": {
                "prefix": "ins_", 
                "filename": os.getenv('INSTANCE_DATA_MODEL_FILENAME', 'instance_model')
            },
            "AUTHORIZATION MODEL": {
                "prefix": "auth_", 
                "filename": os.getenv('AUTHORIZATION_MODEL_FILENAME', 'authorization_model')
            },
            "KNOWLEDGE DATA MODEL": {
                "prefix": "kbm_", 
                "filename": os.getenv('KNOWLEDGE_DATA_MODEL_FILENAME', 'knowledge_model')
            },
            "RS EXCEL": {
                "prefix": "data_", 
                "filename": os.getenv('RS_EXCEL_FILENAME', 'rs_excel_data')
            },
            "thing": {
                "prefix": "thg_", 
                "filename": os.getenv('THING_MODEL_FILENAME', 'thing_model')
            },
            "referenceData": {
                "prefix": "ref_", 
                "filename": os.getenv('REFERENCEDATA_FILENAME', 'reference_data')
            },
            "UOMData": {
                "prefix": "uom_", 
                "filename": os.getenv('UOMDATA_FILENAME', 'uom_data')
            },
            "digitalAsset": {
                "prefix": "dam_", 
                "filename": os.getenv('DIGITALASSET_FILENAME', 'digital_asset')
            },
        }
        try:
            # Only process files, not directories
            if not event.is_directory:
                file_path = Path(event.src_path)
                file_extension = file_path.suffix.lower()
                already_has_prefix=has_known_prefix(file_path.name,MODEL_CONFIGS)
                
                print(f"📄 Processing file: {file_path.name}")
                print(f"   Extension: {file_extension}")
                print(f"   Has prefix: {already_has_prefix}")
                
                if  file_extension in [".xlsx",".xlsm"] and not already_has_prefix:
                    print("   ✅ File matches criteria, processing...")
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
                        replace_filename = os.getenv('REPLACE_FILENAME', 'false').lower() == 'true'
                        prefix = MODEL_CONFIGS[model_name]["prefix"]
                        postfix = os.getenv('FILENAME_POSTFIX', '')
                        
                        if replace_filename:
                            # Replace entire filename with custom name (no prefix)
                            custom_filename = MODEL_CONFIGS[model_name]["filename"]
                            new_file_path = file_path.parent / (custom_filename + postfix + file_path.suffix)
                        else:
                            # Add prefix to existing filename
                            base_name = file_path.stem  # filename without extension
                            new_filename = prefix + base_name + postfix + file_path.suffix
                            new_file_path = file_path.parent / new_filename
                        
                        if new_file_path.exists():
                            new_file_path = add_suffix_to_filename(new_file_path, str(time.time())[-6:])
                        
                        file_path.rename(new_file_path)                        
                        print(f"Template name: {template_name}")
                        if base_model:
                            print(f"Domain name: {domain_name}")
                            disable_window_protection_in_sheetview(new_file_path)
                        print(f"Prefix: {prefix}")
                        print(f"Postfix: {postfix}")
                        print(f"New filename: {new_file_path.name}")
                        print(f"Replace mode: {replace_filename}")
                        return str(new_file_path)  # Return new filename
                else:
                    print("   ❌ File doesn't match criteria (wrong extension or has prefix)")

                print("-" * 50)    
                return None  # No file was renamed

        except Exception as e:
            print("some file caused some error")
            print(e)
            return None

def disable_window_protection_in_sheetview(xlsx_path):
    xlsx_path = Path(xlsx_path)
    temp_dir = Path("temp_unzip")
    if temp_dir.exists():
        shutil.rmtree(temp_dir)

    # Step 1: Unzip workbook
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Step 2: Modify sheet XML files
    sheet_dir = temp_dir / 'xl' / 'worksheets'
    for fpath in sheet_dir.glob('sheet*.xml'):
        parser = ET.XMLParser(remove_blank_text=False)
        tree = ET.parse(str(fpath), parser)
        root = tree.getroot()

        ns = {
            'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        }
        sheet_view = root.find('.//main:sheetView', namespaces=ns)
        if sheet_view is not None and 'windowProtection' in sheet_view.attrib:
            del sheet_view.attrib['windowProtection']
            print(f"✅ Removed windowProtection in {fpath.name}")
            tree.write(str(fpath), encoding='utf-8', xml_declaration=True, pretty_print=False)

    # Step 3: Repackage - replace original file
    temp_new_file = xlsx_path.with_name(f"{xlsx_path.stem}_temp.xlsx")
    with zipfile.ZipFile(temp_new_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(temp_dir.rglob("*")):
            if file_path.is_file():
                arcname = file_path.relative_to(temp_dir)
                zf.write(file_path, arcname)

    # Step 4: Replace original file
    xlsx_path.unlink()  # Remove original file
    temp_new_file.rename(xlsx_path)  # Rename temp file to original name
    
    # Step 5: Cleanup
    shutil.rmtree(temp_dir)
    print(f"🎉 Window protection removed from: {xlsx_path.name}")



def has_known_prefix(file_name, model_configs) -> bool:
    """
    Returns True if the file name contains any prefix from the model_configs dictionary
    or matches any of the custom filenames.
    
    Parameters:
    - file_name (str): The name of the file (e.g., file_path.name).
    - model_configs (dict): Dictionary with model names and their prefix settings.
    
    Returns:
    - bool: True if any prefix or custom filename is found in the file name, False otherwise.
    """
    file_name_string = str(file_name)
    
    # Check for prefixes
    if any(cfg["prefix"] in file_name_string for cfg in model_configs.values()):
        return True
    
    # Check for custom filenames (when REPLACE_FILENAME=true)
    postfix = os.getenv('FILENAME_POSTFIX', '')
    for cfg in model_configs.values():
        custom_filename = cfg["filename"]
        # Remove file extension for comparison
        name_without_ext = Path(file_name_string).stem
        # Check both with and without postfix
        if name_without_ext == custom_filename or name_without_ext == (custom_filename + postfix):
            return True
    
    return False

def get_downloads_directory():
    """Get the user's downloads directory"""
    # First check .env file
    env_downloads_dir = os.getenv('DOWNLOADS_DIR')
    if env_downloads_dir:
        return env_downloads_dir
    
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

def poll_directory(downloads_dir):
    """Polling-based file monitoring for WSL compatibility"""
    print(f"🔄 Using polling method for WSL compatibility")
    known_files = set()
    handler = DownloadHandler()
    
    # Initialize with existing files
    try:
        for item in Path(downloads_dir).iterdir():
            if item.is_file():  
                known_files.add(str(item))
        print(f"📊 Initial scan found {len(known_files)} files")
    except Exception as e:
        print(f"Error scanning directory: {e}")
        return
    
    try:
        while True:
            current_files = set()
            try:
                for item in Path(downloads_dir).iterdir():
                    if item.is_file():
                        current_files.add(str(item))
                
                # Check for new files
                new_files = current_files - known_files
                if new_files:
                    print(f"🔍 Found {len(new_files)} new files to process")
                
                files_to_add = set()
                files_to_remove = set()
                
                for new_file in new_files:
                    print(f"🆕 New file detected via polling: {new_file}")
                    # Create a mock event object
                    class MockEvent:
                        def __init__(self, path):
                            self.src_path = path
                            self.is_directory = False
                    
                    # Process the file and get the new filename if renamed
                    new_filename = handler.process_file(MockEvent(new_file))
                    
                    if new_filename:
                        # File was renamed, track both old and new names
                        files_to_remove.add(new_file)
                        files_to_add.add(new_filename)
                        print(f"📝 File renamed: {new_file} -> {new_filename}")
                    else:
                        # File was not processed/renamed, just track it normally
                        files_to_add.add(new_file)
                
                # Update known_files properly after processing
                known_files = current_files
                
                # If we processed any files, wait a bit longer before next poll
                if new_files:
                    print("⏳ Waiting extra time after processing files...")
                    time.sleep(3)
                
            except Exception as e:
                print(f"Error during polling: {e}")
            
            time.sleep(2)  # Poll every 2 seconds
            
    except KeyboardInterrupt:
        print("\n🛑 Stopping polling monitor...")

def main():
    downloads_dir = get_downloads_directory().strip('"\'')
    
    print(f"Testing path: {downloads_dir}")
    print(f"Path exists (os.path.exists): {os.path.exists(downloads_dir)}")
    print(f"Path exists (Path.exists): {Path(downloads_dir).exists()}")
    print(f"Is directory: {os.path.isdir(downloads_dir)}")
    
    if not os.path.exists(downloads_dir):
        print(f"Directory {downloads_dir} does not exist!")
        return
    
    print(f"✅ Successfully found downloads directory: {downloads_dir}")
    
    # Check if we're in WSL - handle Windows compatibility
    is_wsl = False
    is_windows_path = False
    
    try:
        # Try Unix-style OS detection (works on Linux/macOS/WSL)
        uname_info = os.uname().release.lower()
        is_wsl = "microsoft" in uname_info or "wsl" in uname_info
        is_windows_path = downloads_dir.startswith("/mnt/")
    except AttributeError:
        # Windows doesn't have os.uname(), detect Windows paths instead
        is_windows_path = len(downloads_dir) > 1 and downloads_dir[1] == ':'
    
    if is_wsl and is_windows_path:
        print("🐧 WSL + Windows path detected, using polling method...")
        poll_directory(downloads_dir)
    else:
        print("🔍 Using watchdog file monitoring...")
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
