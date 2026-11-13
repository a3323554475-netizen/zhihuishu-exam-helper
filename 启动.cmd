@echo off
rem Double-click launcher: runs launch.vbs, which starts the helper with no console window.
start "" wscript.exe //nologo "%~dp0launch.vbs"
exit /b 0
