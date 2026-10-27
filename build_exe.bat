@echo off
rem ============================================================
rem  Build Screen-Capture-to-PPT .exe  (error-visible version)
rem  Run on Windows by double-clicking this file.
rem  Output: dist\ScreenCapturePPT.exe
rem ============================================================
setlocal
pushd "%~dp0"

echo ============================================================
echo  Build Screen-Capture-to-PPT .exe
echo ============================================================

echo [1/4] Locate a usable Python ...
set "PYTHON_CMD=python"
where py >nul 2>&1 && set "PYTHON_CMD=py -3"
%PYTHON_CMD% --version
if errorlevel 1 (
    echo ERROR: no Python found. Install Python 3.11+ from python.org
    echo        and tick "Add Python to PATH", then re-run this file.
    pause
    exit /b 1
)
%PYTHON_CMD% -c "import sys;print('  Python at:',sys.executable)"
echo.

echo [2/4] Check pip ...
%PYTHON_CMD% -m pip --version
if errorlevel 1 (
    echo.
    echo ERROR: THIS Python has NO pip and cannot install packages.
    echo        This is typical for bundled runtimes (e.g. ST software).
    echo.
    echo SOLUTION: install OFFICIAL Python 3.11+ from
    echo        https://www.python.org/downloads/
    echo        and tick "Add Python to PATH", then close and reopen
    echo        the cmd/PowerShell window and re-run this file.
    pause
    exit /b 1
)
echo.

echo [3/4] Install dependencies ... (all output shown, may take minutes)
%PYTHON_CMD% -m pip install --upgrade pip
%PYTHON_CMD% -m pip install mss opencv-python numpy pillow python-pptx pyinstaller
if errorlevel 1 (
    echo ERROR: dependency install FAILED. Read the error text above.
    pause
    exit /b 1
)
echo.

echo [4/4] Package with PyInstaller ...
%PYTHON_CMD% -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "ScreenCapturePPT" ^
    --collect-data pptx ^
    --collect-all PIL ^
    --collect-all cv2 ^
    --hidden-import mss ^
    --hidden-import PIL._tkinter_finder ^
    "%~dp0scripts\screen_capture_gui.py"
if errorlevel 1 (
    echo ERROR: PyInstaller failed. Read the error text above.
    pause
    exit /b 1
)

echo.
echo SUCCESS! The .exe is at: "%~dp0dist\ScreenCapturePPT.exe"
pause
endlocal