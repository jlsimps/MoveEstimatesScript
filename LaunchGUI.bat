@echo off
set PowerShellScript=LaunchMoveEstimatesGUI.ps1

REM Launch the PowerShell GUI script
powershell -NoProfile -ExecutionPolicy Bypass -File "%PowerShellScript%"
pause