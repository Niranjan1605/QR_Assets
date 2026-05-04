@echo off
REM ================================
REM QR Sticker Generator - SharePoint
REM ================================

echo Running Dynamic QR Code Generator...
cd /d "C:\QRAssets"

REM Optional: Activate your Python environment if needed
REM call C:\path\to\your\venv\Scripts\activate

python generate_dynamic_qr_sharepoint.py

echo.
echo -------------------------------------
echo ✅ QR Code generation complete!
echo Output saved in: C:\QRAssets\qr_codes
echo -------------------------------------
pause
