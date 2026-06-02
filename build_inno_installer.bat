@echo off
setlocal

REM Build a standard Windows installer with Inno Setup.
REM This includes the bundled Python GUI app, but does not include Origin itself.
REM Target computers must already have Origin 2023b installed and licensed.

if not exist "dist\OriginScatterGUI\OriginScatterGUI.exe" (
    call build_exe.bat
)

set "ISCC=C:\Users\20894\AppData\Local\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
    set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
)
if not exist "%ISCC%" (
    echo Inno Setup was not found. Please install Inno Setup 6 first.
    exit /b 1
)

if not exist "release" mkdir "release"
"%ISCC%" "installer_assets\OriginScatterGUI.iss"

endlocal
