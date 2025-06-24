@echo off
echo Starting Syndigo File Download Helper...

REM Check if .env exists
if not exist .env (
    echo Copying .env.example to .env...
    copy .env.example .env
    echo Please edit .env file with your downloads directory path before running again.
    pause
    exit /b 1
)

REM Check if virtual environment exists
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo Error: Failed to create virtual environment. Make sure Python is installed.
        pause
        exit /b 1
    )
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Check if requirements are installed by trying to import a key package
python -c "import watchdog" 2>nul
if errorlevel 1 (
    echo Installing requirements...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo Error: Failed to install requirements.
        pause
        exit /b 1
    )
)

REM Run the program
echo Starting download monitor...
python downloadMonitor.py

pause