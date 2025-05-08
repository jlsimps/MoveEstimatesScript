Import-Module ImportExcel

# Path to the Excel file
$excelFilePath = "C:\Users\jake.simpson\Downloads\PowerShell_Template.xlsx"

# Read the Excel file (skip the first row)
$data = Import-Excel -Path $excelFilePath | Select-Object -Skip 1

foreach ($row in $data) {
    $folderName = $row."Estimate Code"  # Column A: Folder name
    $destinationPath = $row."Old Pathing"  # Column C: Original source (now destination)
    $sourcePath = $row."New Pathing"       # Column D: Original destination (now source)

    # Stop processing if the row values are blank
    if ([string]::IsNullOrWhiteSpace($folderName) -and [string]::IsNullOrWhiteSpace($sourcePath) -and [string]::IsNullOrWhiteSpace($destinationPath)) {
        break
    }

    # Ensure source and destination paths are valid
    if ((Test-Path -Path $sourcePath) -and ($folderName -ne "") -and ($destinationPath -ne "")) {
        $fullSourcePath = Join-Path -Path $sourcePath -ChildPath $folderName
        $fullDestinationPath = Join-Path -Path $destinationPath -ChildPath $folderName

        try {
            # Create the destination directory if it doesn't exist
            if (!(Test-Path -Path $destinationPath)) {
                New-Item -ItemType Directory -Path $destinationPath | Out-Null
            }

            # Move the directory back
            Move-Item -Path $fullSourcePath -Destination $fullDestinationPath -Force
            Write-Host "Successfully moved $folderName back to $fullDestinationPath"
        } catch {
            Write-Host "Failed to move $folderName back. Error: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Invalid source or destination path for $folderName. Skipping..." -ForegroundColor Yellow
    }
}