# Syndigo File Download Helper

## Quick Start (Recommended)

**Windows:**
1. Double-click `run.bat`. Or `.\run.bat` in the terminal.
2. Edit `.env` file when prompted with your downloads directory path
3. Run `run.bat` again
4. exit with `ctrl+c`

**Mac/Linux/WSL:**
1. Run `./run.sh` in terminal
2. Edit `.env` file when prompted with your downloads directory path  
3. Run `./run.sh` again
4. exit with `ctrl+c` (Linux/WSL) or `cmd+c` (Mac)

## Manual Setup (Advanced Users)

1. Create virtual environment:
   ```bash
   python -m venv venv
   ```

2. Activate virtual environment:
   
   **Windows:**
   ```cmd
   venv\Scripts\activate
   ```
   
   **Mac/Linux/WSL:**
   ```bash
   source venv/bin/activate
   ```

3. Copy environment file:
   ```bash
   cp .env.example .env
   ```

4. Install requirements:
   ```bash
   pip install -r requirements.txt
   ```

5. Edit `.env` and set your downloads directory path

6. Run manually:

   **Windows:**
   ```cmd
   python downloadMonitor.py
   ```

   **Mac/Linux/WSL:**
   ```bash
   python3 downloadMonitor.py
   ```

## File Versioning

When a file with the same name already exists, the system uses smart versioning:

- **New file keeps the original name** (e.g., `gov_model.xlsx`)
- **Existing files get numbered versions** (e.g., `gov_model_oldv1.xlsx`, `gov_model_oldv2.xlsx`)
- **Old versions keep their names** - no shifting or renaming of previous versions

**Example sequence:**
1. First download: `model.xlsx`
2. Second download: `model_oldv1.xlsx`, `model.xlsx` (newest)
3. Third download: `model_oldv1.xlsx`, `model_oldv2.xlsx`, `model.xlsx` (newest)

**File Lock Handling:**
If a file is open in Excel or another application:
- System detects the lock and shows a message
- Waits up to 5 minutes for you to save and close the file
- Automatically proceeds once the file is available
- If timeout occurs, processing is skipped with a helpful message

## Environment Override Files

You can use different environment configurations without modifying your main `.env` file by using override files.

### Creating Override Files

1. Create a new file (e.g., `.env.overwrite`, `.env.client1`, `.env.testing`)
2. Add only the variables you want to change:
   ```bash
   # Example .env.overwrite
   FILENAME_POSTFIX=_CUSTOM
   PROCESSED_FILES_DIR="/path/to/different/folder"
   ```

### Using Override Files

**Quick Start Scripts:**
- Windows: `run.bat --override-env .env.overwrite`
- Mac/Linux/WSL: `./run.sh --override-env .env.overwrite`

**Manual Run:**
```bash
python downloadMonitor.py --override-env .env.overwrite
```

**How it works:**
1. Loads all settings from your main `.env` file
2. Overwrites only the variables specified in the override file
3. All other settings remain unchanged

**Common use cases:**
- `.env.overwrite` - Custom postfix and directory
- `.env.client1` - Client-specific configurations  
- `.env.testing` - Testing with different paths

**Note:** Override files are automatically ignored by git, so your custom configurations stay private.

## Requirements

- Python 3.x
- Required packages will be installed automatically
