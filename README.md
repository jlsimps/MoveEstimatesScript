# Estimate Migration Script - README

## Overview

This project automates the process of moving estimate directories when converting from the Standard to the Enterprise version of the estimating software. It uses PowerShell scripts and a simple GUI to help users populate a template, assign estimates to divisions, and move folders accordingly.

## Folder Structure

```
Root Directory
├── EST\                        # Original location of all estimates (Standard version)
├── Division1\
│   └── EST\                    # New location for estimates belonging to Division1
├── Division2\
│   └── EST\                    # New location for estimates belonging to Division2
├── PowerShell_Template.xlsx    # Excel file used to map estimates to divisions
├── LaunchGUI.bat               # Launches the GUI for the process
├── LaunchMoveEstimatesGUI.ps1  # PowerShell GUI
├── MoveEstimates.ps1           # Script that performs the move
├── UndoMove.ps1                # Script to undo the move
└── PopulateTemplate.ps1        # Generates the Excel template
```

## Prerequisites

- PowerShell 5.1 or later
- ImportExcel module installed (`Install-Module -Name ImportExcel`)

## How to Use the Script

### 1. Launch the GUI

- Double-click `LaunchGUI.bat` to open the script GUI.

### 2. Populate the Template

- In the GUI, click "Populate Template".
- This runs `PopulateTemplate.ps1` and generates an Excel file (`PowerShell_Template.xlsx`) listing all current estimate folders.

### 3. Assign Divisions in Excel

- Open the generated Excel file.
- Fill in the **Target Division** column with the name of the division each estimate should be moved to.

  Example:

  | Estimate | Path               | Target Division |
  |----------|--------------------|-----------------|
  | EST001   | C:\Root\EST\EST001 | Paving        |
  | EST002   | C:\Root\EST\EST002 | Earthwork     |

- Save and close the Excel file.

### 4. (Optional but Recommended) Perform a Dry Run

- In the GUI, click "Dry Run".
- This will simulate the move process without actually relocating any folders.
- The GUI will display the list of actions that would be taken, allowing you to verify the intended moves.
- Review the output carefully to ensure estimates will be moved to the correct locations.

### 5. Move the Estimates

- Back in the GUI, click "Move Estimates".
- This runs `MoveEstimates.ps1`, which reads the Excel file and moves each estimate folder to the correct division's `EST` folder.

### 6. Undo (if needed)

- If an error occurs or the moves were incorrect, run `UndoMove.ps1`.
- This will revert the estimates to their original locations using the move log.

## Tips

- Ensure all target division folders and their `EST` subfolders exist before running the move.
- Avoid duplicate column headers in the Excel file.
- Always keep a backup of the original `EST` folder just in case.

## Support

For issues or enhancements, please contact the support team or project maintainer.

## Notes

- This project was developed internally to streamline estimate folder reorganization and reduce manual errors during conversions to the Enterprise version.
- The included GUI makes it easy for less technical users to use the scripts effectively.
