@echo off
setlocal

REM Build a user-level Windows installer with the built-in IExpress tool.
REM The installer includes this GUI app and Python dependencies, but does not include Origin itself.
REM Target computers must already have Origin 2023b installed and usable through originpro.

if not exist "dist\OriginScatterGUI\OriginScatterGUI.exe" (
    call build_exe.bat
)

if not exist "installer_assets" mkdir "installer_assets"
if not exist "release" mkdir "release"
if exist "release\OriginScatterGUI_Setup.exe" del /q "release\OriginScatterGUI_Setup.exe"
if exist "E:\iwen-codex\OriginScatterGUI_Setup.exe" del /q "E:\iwen-codex\OriginScatterGUI_Setup.exe"

powershell -NoProfile -ExecutionPolicy Bypass -Command "if (Test-Path 'installer_assets\OriginScatterGUI.zip') { Remove-Item 'installer_assets\OriginScatterGUI.zip' -Force }; Compress-Archive -Path 'dist\OriginScatterGUI\*' -DestinationPath 'installer_assets\OriginScatterGUI.zip' -Force"

set "TMP_INSTALLER=E:\iwen-codex\OriginScatterInstallerTmp"
if exist "%TMP_INSTALLER%" rmdir /s /q "%TMP_INSTALLER%"
mkdir "%TMP_INSTALLER%"
copy /y "installer_assets\install.bat" "%TMP_INSTALLER%\install.bat" >nul
copy /y "installer_assets\OriginScatterGUI.zip" "%TMP_INSTALLER%\OriginScatterGUI.zip" >nul

powershell -NoProfile -ExecutionPolicy Bypass -Command "$sed='[Version]'+[Environment]::NewLine+'Class=IEXPRESS'+[Environment]::NewLine+'SEDVersion=3'+[Environment]::NewLine+[Environment]::NewLine+'[Options]'+[Environment]::NewLine+'PackagePurpose=InstallApp'+[Environment]::NewLine+'ShowInstallProgramWindow=0'+[Environment]::NewLine+'HideExtractAnimation=1'+[Environment]::NewLine+'UseLongFileName=1'+[Environment]::NewLine+'InsideCompressed=0'+[Environment]::NewLine+'CAB_FixedSize=0'+[Environment]::NewLine+'CAB_ResvCodeSigning=0'+[Environment]::NewLine+'RebootMode=N'+[Environment]::NewLine+'InstallPrompt='+[Environment]::NewLine+'DisplayLicense='+[Environment]::NewLine+'FinishMessage='+[Environment]::NewLine+'TargetName=E:\iwen-codex\OriginScatterGUI_Setup.exe'+[Environment]::NewLine+'FriendlyName=OriginScatterGUI Installer'+[Environment]::NewLine+'AppLaunched=install.bat'+[Environment]::NewLine+'PostInstallCmd=<None>'+[Environment]::NewLine+'AdminQuietInstCmd=' + [Environment]::NewLine + 'UserQuietInstCmd=' + [Environment]::NewLine + 'SourceFiles=SourceFiles' + [Environment]::NewLine + [Environment]::NewLine + '[Strings]' + [Environment]::NewLine + 'FILE0=install.bat' + [Environment]::NewLine + 'FILE1=OriginScatterGUI.zip' + [Environment]::NewLine + [Environment]::NewLine + '[SourceFiles]' + [Environment]::NewLine + 'SourceFiles0=E:\iwen-codex\OriginScatterInstallerTmp\' + [Environment]::NewLine + [Environment]::NewLine + '[SourceFiles0]' + [Environment]::NewLine + '%%FILE0%%=' + [Environment]::NewLine + '%%FILE1%%='; Set-Content -LiteralPath '%TMP_INSTALLER%\setup.sed' -Value $sed -Encoding ASCII"
iexpress /N /Q "%TMP_INSTALLER%\setup.sed"
copy /y "E:\iwen-codex\OriginScatterGUI_Setup.exe" "release\OriginScatterGUI_Setup.exe" >nul

endlocal
