@echo off
setlocal

set "APP_NAME=OriginScatterGUI"
set "INSTALL_DIR=%LOCALAPPDATA%\%APP_NAME%"
set "PAYLOAD=%~dp0OriginScatterGUI.zip"

if not exist "%PAYLOAD%" (
    echo Install package is incomplete: OriginScatterGUI.zip was not found.
    pause
    exit /b 1
)

echo Installing %APP_NAME% to:
echo %INSTALL_DIR%

if exist "%INSTALL_DIR%" (
    rmdir /s /q "%INSTALL_DIR%"
)
mkdir "%INSTALL_DIR%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath '%PAYLOAD%' -DestinationPath '%INSTALL_DIR%' -Force"
if errorlevel 1 (
    echo Failed to extract application files.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Origin 二维散点图批量绘制工具.lnk'); $s.TargetPath='%INSTALL_DIR%\OriginScatterGUI.exe'; $s.WorkingDirectory='%INSTALL_DIR%'; $s.Save()"

echo.
echo Installation finished.
echo Origin 2023b must already be installed on this computer.
echo.
start "" "%INSTALL_DIR%\OriginScatterGUI.exe"

endlocal
