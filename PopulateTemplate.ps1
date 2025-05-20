param (
    [Parameter(Mandatory = $true)]
    [string]$TemplateDirectory,

    [Parameter(Mandatory = $true)]
    [string]$EstDirectory
)

# Path to EST folder and the Excel template
$estDir = $EstDirectory
$templatePath = Join-Path -Path $TemplateDirectory -ChildPath "PowerShell_Template.xlsx"
$sheetName = "Sheet1"

# Validate paths
if (-not (Test-Path $templatePath)) {
    Write-Host "[ERROR] Excel file not found at: $templatePath"
    exit 1
}

# Get list of estimate folders
$folders = Get-ChildItem -Path $estDir -Directory

# Open template for editing
$excelPackage = Open-ExcelPackage -Path $templatePath
$ws = $excelPackage.Workbook.Worksheets[$sheetName]

# Clear old data in columns A, B, and C starting from row 2
$rowCount = $ws.Dimension.Rows
for ($row = 2; $row -le $rowCount; $row++) {
    $ws.Cells["A$row"].Clear()
    $ws.Cells["B$row"].Clear()
    $ws.Cells["C$row"].Clear()
}

# Write new folder data starting from row 2
$row = 2
foreach ($folder in $folders) {
    $ws.Cells["A$row"].Value = $folder.Name
    $ws.Cells["C$row"].Value = $EstDirectory
    $row++
}

# Save and close the file
Close-ExcelPackage $excelPackage

Write-Host "Template updated: data cleared and new estimates written to $templatePath"