#!/bin/bash

echo "========================================"
echo "ARN Change Form Generator - Starting..."
echo "========================================"
echo

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "ERROR: Virtual environment not found!"
    echo "Please run ./setup.sh first to install the application."
    echo
    exit 1
fi

# Check if required template file exists
if [ ! -f "Request for Change of Broker.docx" ]; then
    echo "ERROR: Template file 'Request for Change of Broker.docx' not found!"
    echo "Please ensure the Word template is in the same directory."
    echo
    exit 1
fi

# Activate virtual environment
echo "[1/3] Activating virtual environment..."
source venv/bin/activate
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to activate virtual environment"
    echo "Please run ./setup.sh to reinstall the application."
    exit 1
fi

echo "Virtual environment activated successfully!"
echo

# Start the Flask application
echo "[2/3] Starting ARN Change Form Generator..."
echo
echo "Server will start at: http://localhost:8000"
echo
echo "Instructions:"
echo "1. Open http://localhost:8000 in your web browser"
echo "2. Drag and drop your Excel file onto the webpage"
echo "3. Click 'Generate ARN Form' to process"
echo "4. The populated Word document will download automatically"
echo
echo "Press Ctrl+C to stop the server when done."
echo

# Try to open browser automatically (works on macOS and most Linux desktops)
sleep 3
if command -v open &> /dev/null; then
    # macOS
    open http://localhost:8000
elif command -v xdg-open &> /dev/null; then
    # Linux
    xdg-open http://localhost:8000
elif command -v start &> /dev/null; then
    # Windows (if running in WSL)
    start http://localhost:8000
fi

# Start the Flask application (this will keep running)
python app.py