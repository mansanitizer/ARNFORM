@echo off
echo ========================================
echo ARN Change Form Generator - Starting...
echo ========================================
echo.

:: Check if virtual environment exists
if not exist "venv" (
    echo ERROR: Virtual environment not found!
    echo Please run setup.bat first to install the application.
    echo.
    pause
    exit /b 1
)

:: Check if required template file exists
if not exist "Request for Change of Broker.docx" (
    echo ERROR: Template file "Request for Change of Broker.docx" not found!
    echo Please ensure the Word template is in the same directory.
    echo.
    pause
    exit /b 1
)

:: Activate virtual environment
echo [1/3] Activating virtual environment...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment
    echo Please run setup.bat to reinstall the application.
    pause
    exit /b 1
)

echo Virtual environment activated successfully!
echo.

:: Start the Flask application in background
echo [2/3] Starting ARN Change Form Generator...
echo.
echo Server will start at: http://localhost:8000
echo.
echo Instructions:
echo 1. The web browser will open automatically
echo 2. Drag and drop your Excel file onto the webpage
echo 3. Click "Generate ARN Form" to process
echo 4. The populated Word document will download automatically
echo.
echo Press Ctrl+C to stop the server when done.
echo.

:: Open browser after a short delay
timeout /t 3 /nobreak >nul
start http://localhost:8000

:: Start the Flask application (this will keep running)
python app.py

:: This will only execute if the Python app exits
echo.
echo Application stopped.
pause