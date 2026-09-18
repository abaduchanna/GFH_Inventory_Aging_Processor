@echo off
setlocal enabledelayedexpansion
title Build GFH_Inventory_Aging_Processor

set "SRCDIR=C:\Users\AbadUmairChanna\Downloads\GitHub\gfh-inventory-aging-processor"
set "OUTDIR=C:\Users\AbadUmairChanna\Downloads\GitHub"
set "WORKBASE=%TEMP%\pyi_build\GFH_Inventory_Aging_Processor"

echo.
echo  ============================================================
echo   Building: GFH_Inventory_Aging_Processor.exe
echo  ============================================================
echo.

REM Check prerequisites
python --version >nul 2>&1
if errorlevel 1 (
    echo    ERROR: Python not found in PATH.
    pause
    exit /b 1
)
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo    PyInstaller not found. Installing...
    python -m pip install --upgrade pyinstaller
)
git --version >nul 2>&1
if errorlevel 1 (
    echo    ERROR: Git not found in PATH.
    pause
    exit /b 1
)
echo    Prerequisites OK
echo.

REM -- Step 1: Sync code from GitHub -------------------------------------------
REM  NOTE: a plain "git pull" breaks silently when GitHub history gets
REM  rewritten (force-push), and the build then uses stale code forever.
REM  This block fetches origin/main and hard-resets the clone onto it, so
REM  the build ALWAYS matches GitHub. Any local edits are stashed first.
echo  Step 1: Syncing code from GitHub...
if exist "%SRCDIR%\.git" (
    cd /d "%SRCDIR%"
    git remote set-url origin "https://github.com/abaduchanna/gfh-inventory-aging-processor.git" >nul 2>&1
    git fetch origin main
    if errorlevel 1 (
        echo    WARNING: Could not reach GitHub - building the local code as-is.
    ) else (
        git stash push -m "auto-pre-build" >nul 2>&1
        git reset --hard origin/main
        if errorlevel 1 (
            echo    WARNING: Repo update failed - building the local code as-is.
        ) else (
            echo    Updated to latest GitHub code:
            git log -1 --oneline
        )
    )
) else if exist "%SRCDIR%" (
    echo    Folder exists but is not a git clone - moving it aside and re-cloning...
    if exist "%SRCDIR%.old" rmdir /s /q "%SRCDIR%.old"
    move "%SRCDIR%" "%SRCDIR%.old" >nul
    git clone "https://github.com/abaduchanna/gfh-inventory-aging-processor.git" "%SRCDIR%"
    cd /d "%SRCDIR%"
) else (
    echo    Cloning gfh-inventory-aging-processor...
    git clone "https://github.com/abaduchanna/gfh-inventory-aging-processor.git" "%SRCDIR%"
    cd /d "%SRCDIR%"
)
echo.

REM -- Step 2: Verify the Exclusions panel is in the source ---------------------
REM  Guards against the "stale EXE" problem: if this marker is missing the
REM  source is outdated and the built EXE will lack the Exclusions button.
echo  Step 2: Verifying source is current...
findstr /c:"class ExclusionsDialog" "GFH_Inventory_Aging_Processor.py" >nul 2>&1
if errorlevel 1 (
    echo    WARNING: Exclusions panel NOT found in source - code is STALE!
    echo    The EXE will NOT have the Exclusions button. Check Step 1 above.
) else (
    echo    Exclusions panel: PRESENT in source.
)
echo.

REM -- Step 3: Clean previous build ---------------------------------------------
echo  Step 3: Cleaning previous build...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "__pycache__" rmdir /s /q "__pycache__"
del /s /q *.pyc 2>nul

REM -- Step 4: Redirect PyInstaller workpath to TEMP ----------------------------
REM  Avoids base_library.zip crashes when OneDrive syncs the build folder.
if exist "%WORKBASE%" rmdir /s /q "%WORKBASE%"
mkdir "%WORKBASE%" 2>nul

REM -- Step 5: Install dependencies ---------------------------------------------
echo  Step 5: Installing requirements...
if exist "requirements.txt" (
    python -m pip install -r requirements.txt --quiet 2>nul
)

REM -- Step 6: Build -------------------------------------------------------------
echo    Building GFH_Inventory_Aging_Processor.spec...
python -m PyInstaller "GFH_Inventory_Aging_Processor.spec" --noconfirm --clean --workpath "%WORKBASE%" 2>&1

if errorlevel 1 (
    echo    FAILED: GFH_Inventory_Aging_Processor
    pause
    exit /b 1
)

echo    SUCCESS: GFH_Inventory_Aging_Processor

REM -- Step 7: Copy .exe to output ------------------------------------------------
if exist "dist\GFH_Inventory_Aging_Processor.exe" (
    if not exist "%OUTDIR%" mkdir "%OUTDIR%"
    copy /Y "dist\GFH_Inventory_Aging_Processor.exe" "%OUTDIR%\GFH_Inventory_Aging_Processor.exe" >nul
    echo    Collected: %OUTDIR%\GFH_Inventory_Aging_Processor.exe
) else (
    echo    WARNING: dist\GFH_Inventory_Aging_Processor.exe not found
)

echo.
echo  ============================================================
echo   Done: GFH_Inventory_Aging_Processor.exe
echo  ============================================================
echo.
pause
endlocal
exit /b 0
