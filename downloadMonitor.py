import os
import time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class DownloadHandler(FileSystemEventHandler):
    def on_created(self, event):
        # Only process files, not directories
        if not event.is_directory:
            file_path = Path(event.src_path)
            file_extension = file_path.suffix.lower()
            
            print(f"New file detected: {file_path.name}")
            print(f"Extension: {file_extension if file_extension else 'No extension'}")
            print(f"Full path: {file_path}")
            print("-" * 50)
            
            # Optional: Log to file
            self.log_to_file(file_path.name, file_extension)
    
    def log_to_file(self, filename, extension):
        """Optional: Save file info to a log file"""
        log_path = Path.home() / "download_monitor.log"
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"{timestamp} - {filename} - {extension}\n")

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