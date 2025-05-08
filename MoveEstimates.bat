@echo off
SETLOCAL

:: Ask user for directory containing PowerShell_Template.xlsx
set /p ExcelFolder="Enter the folder path where PowerShell_Template.xlsx is saved: "

:: Optional: ask if user wants dry run
set /p DryRunChoice="Do a dry run (Y/N)? "

:: Build dry run flag
set DryRunFlag=
if /I "%DryRunChoice%"=="Y" set DryRunFlag=-DryRun

:: Testing the command
echo powershell -NoProfile -ExecutionPolicy Bypass -File "MoveEstimates.ps1" -TemplateDirectory "%ExcelFolder%" %DryRunFlag%

:: Run the PowerShell script (wraps directory variable in escaped quotes to handle spaces in file path)
powershell -NoProfile -ExecutionPolicy Bypass -File "MoveEstimates.ps1" -TemplateDirectory "%ExcelFolder%" %DryRunFlag%



pause
ENDLOCAL