@echo off
echo ========================================
echo ARN Change Form Generator - Setup
echo ========================================
echo.

:: Check if Python is installed
echo [1/5] Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

python --version
echo Python found successfully!
echo.

:: Check if pip is available
echo [2/5] Checking pip availability...
python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: pip is not available
    echo Please reinstall Python with pip included
    pause
    exit /b 1
)

echo pip found successfully!
echo.

:: Create virtual environment
echo [3/5] Creating virtual environment...
if exist "venv" (
    echo Virtual environment already exists, skipping creation...
) else (
    python -m venv venv
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
    echo Virtual environment created successfully!
)
echo.

:: Activate virtual environment and install requirements
echo [4/5] Installing required packages...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment
    pause
    exit /b 1
)

:: Upgrade pip first
python -m pip install --upgrade pip

:: Install required packages
python -m pip install flask openpyxl python-docx
if %errorlevel% neq 0 (
    echo ERROR: Failed to install required packages
    pause
    exit /b 1
)

echo Required packages installed successfully!
echo.

:: Verify installation
echo [5/5] Verifying installation...
python -c "import flask, openpyxl, docx; print('All packages imported successfully!')"
if %errorlevel% neq 0 (
    echo ERROR: Package verification failed
    pause
    exit /b 1
)

echo.
echo ========================================
echo Setup completed successfully!
echo ========================================
echo.
echo You can now run the application using: run.bat
echo.
pause