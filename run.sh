#!/bin/bash
echo "================================================"
echo "  HOHEPA AUCKLAND ROSTER SYSTEM"
echo "================================================"
echo ""
echo "Installing packages if needed..."
pip install flask openpyxl pandas xlsxwriter --quiet 2>/dev/null || pip3 install flask openpyxl pandas xlsxwriter --quiet
echo ""
echo "Starting app..."
echo "Open your browser: http://localhost:5050"
echo "Press Ctrl+C to stop"
echo "================================================"
python3 app.py || python app.py
