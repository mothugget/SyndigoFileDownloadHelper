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

## Requirements

- Python 3.x
- Required packages will be installed automatically
