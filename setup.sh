#!/bin/bash

echo "========================================"
echo "ARN Change Form Generator - Setup"
echo "========================================"
echo

# Check if Python is installed
echo "[1/5] Checking Python installation..."
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed"
    echo
    echo "Please install Python 3 from: https://www.python.org/downloads/"
    echo "Or use your system package manager:"
    echo "  Ubuntu/Debian: sudo apt update && sudo apt install python3 python3-pip python3-venv"
    echo "  macOS: brew install python3"
    echo "  CentOS/RHEL: sudo yum install python3 python3-pip"
    echo
    exit 1
fi

python3 --version
echo "Python found successfully!"
echo

# Check if pip is available
echo "[2/5] Checking pip availability..."
if ! python3 -m pip --version &> /dev/null; then
    echo "ERROR: pip is not available"
    echo "Please install pip for Python 3"
    exit 1
fi

echo "pip found successfully!"
echo

# Create virtual environment
echo "[3/5] Creating virtual environment..."
if [ -d "venv" ]; then
    echo "Virtual environment already exists, skipping creation..."
else
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to create virtual environment"
        exit 1
    fi
    echo "Virtual environment created successfully!"
fi
echo

# Activate virtual environment and install requirements
echo "[4/5] Installing required packages..."
source venv/bin/activate
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to activate virtual environment"
    exit 1
fi

# Upgrade pip first
python -m pip install --upgrade pip

# Install required packages
python -m pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install required packages"
    exit 1
fi

echo "Required packages installed successfully!"
echo

# Verify installation
echo "[5/5] Verifying installation..."
python -c "import flask, openpyxl, docx; print('All packages imported successfully!')"
if [ $? -ne 0 ]; then
    echo "ERROR: Package verification failed"
    exit 1
fi

echo
echo "========================================"
echo "Setup completed successfully!"
echo "========================================"
echo
echo "You can now run the application using: ./run.sh"
echo