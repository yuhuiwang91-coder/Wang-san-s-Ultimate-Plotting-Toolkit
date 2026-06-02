@echo off
setlocal

REM Build the Tkinter GUI as a Windows executable.
REM Origin itself is not bundled. The target machine must have Origin 2023b installed.
py -3.10 -m PyInstaller --noconfirm --onedir --windowed --name OriginScatterGUI --icon assets\app_icon.ico --add-data "assets\app_icon.ico;assets" --add-data "assets\app_background.gif;assets" --hidden-import originpro app.py

endlocal
