@echo off
echo ======================================================
echo          YouTube Transcript Downloader Builder
echo ======================================================

REM Check if Python is installed and in PATH
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Please install Python and add it to your PATH.
    pause
    exit /b
)

REM Check if pip is available
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] pip not found. Please ensure your Python installation includes pip.
    pause
    exit /b
)

echo.
echo [STEP 1] Installing required libraries (pyinstaller and others)...
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

if %errorlevel% neq 0 (
    echo [ERROR] Failed to install required libraries. Please check your internet connection and try again.
    pause
    exit /b
)

echo.
echo [STEP 2] Building the executable with PyInstaller...
echo This may take a few minutes.

pyinstaller --name myc --onefile --windowed --icon=NONE myc_gui.py

if %errorlevel% neq 0 (
    echo [ERROR] PyInstaller failed to build the executable.
    pause
    exit /b
)

echo.
echo ======================================================
echo      BUILD COMPLETE!
echo ======================================================
echo.
echo - The executable 'myc.exe' can be found in the 'dist' folder.
echo - You can now move 'myc.exe' to any location and run it.
echo.

REM Clean up build files
rmdir /s /q build
del myc.spec

pause
