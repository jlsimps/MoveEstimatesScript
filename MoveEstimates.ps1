param(
    [Parameter(Mandatory = $true)]
    [string]$TemplateDirectory,

    [switch]$DryRun
)

# Ensure ImportExcel is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    try {
        Write-Host "[INFO] Installing required module 'ImportExcel'..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "[INFO] 'ImportExcel' module installed successfully."
    }
    catch {
        Write-Host "[ERROR] Failed to install the 'ImportExcel' module. Please install it manually with:"
        Write-Host "        Install-Module -Name ImportExcel -Scope CurrentUser"
        exit 1
    }
}

# Import the module
Import-Module ImportExcel -ErrorAction Stop

# Define path to Excel file
$excelFile = Join-Path -Path $TemplateDirectory -ChildPath "PowerShell_Template.xlsx"

# Validate the file path
if (-not (Test-Path $excelFile)) {
    Write-Host "[ERROR] Excel file not found at: $excelFile"
    exit 1
}

# Read Excel file and skip header
$data = Import-Excel -Path $excelFile | Select-Object -Skip 1

foreach ($row in $data) {
    $folderName      = $row."Estimate Code"
    $sourcePath      = $row."Old Pathing"
    $destinationPath = $row."New Pathing"

    # Skip empty rows
    if ([string]::IsNullOrWhiteSpace($folderName) -and [string]::IsNullOrWhiteSpace($sourcePath) -and [string]::IsNullOrWhiteSpace($destinationPath)) {
        break
    }

    # Validate source/destination
    if ((Test-Path -Path $sourcePath) -and ($folderName -ne "") -and ($destinationPath -ne "")) {
        $fullSourcePath = Join-Path -Path $sourcePath -ChildPath $folderName
        $fullDestinationPath = Join-Path -Path $destinationPath -ChildPath $folderName

        try {
            # Create destination path if needed
            if (-not (Test-Path -Path $destinationPath)) {
                if (-not $DryRun) {
                    New-Item -ItemType Directory -Path $destinationPath | Out-Null
                }
                Write-Host "[INFO] Created directory: $destinationPath"
            }

            # Move folder
            if ($DryRun) {
                Write-Host "[DryRun] Would move: $fullSourcePath --> $fullDestinationPath"
            } else {
                Move-Item -Path $fullSourcePath -Destination $fullDestinationPath -Force
                Write-Host "[OK] Moved $folderName to $fullDestinationPath"
            }
        }
        catch {
            Write-Host "[ERROR] Failed to move $folderName. Error: $_"
        }
    } else {
        Write-Host "[SKIP] Invalid source or destination path for '$folderName'. Skipping..."
    }
}