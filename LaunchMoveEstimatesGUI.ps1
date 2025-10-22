Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Move Estimates Launcher"
$form.Size = New-Object System.Drawing.Size(550, 600)
$form.StartPosition = "CenterScreen"

# === TEMPLATE SECTION ===

# Template directory label
$labelTemplate = New-Object System.Windows.Forms.Label
$labelTemplate.Text = "Select Directory where PowerShell_Template.xlsx is Located:"
$labelTemplate.Location = New-Object System.Drawing.Point(10, 20)
$labelTemplate.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($labelTemplate)

# Template directory textbox
$textBoxTemplate = New-Object System.Windows.Forms.TextBox
$textBoxTemplate.Location = New-Object System.Drawing.Point(10, 45)
$textBoxTemplate.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($textBoxTemplate)

# Browse button for template directory
$browseButtonTemplate = New-Object System.Windows.Forms.Button
$browseButtonTemplate.Text = "Browse"
$browseButtonTemplate.Location = New-Object System.Drawing.Point(370, 43)
$browseButtonTemplate.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $textBoxTemplate.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($browseButtonTemplate)

# EST directory label
$labelEST = New-Object System.Windows.Forms.Label
$labelEST.Text = "Select Original EST Directory:"
$labelEST.Location = New-Object System.Drawing.Point(10, 80)
$labelEST.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($labelEST)

# EST directory textbox
$textBoxEST = New-Object System.Windows.Forms.TextBox
$textBoxEST.Location = New-Object System.Drawing.Point(10, 105)
$textBoxEST.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($textBoxEST)

# Browse button for EST directory
$browseButtonEST = New-Object System.Windows.Forms.Button
$browseButtonEST.Text = "Browse"
$browseButtonEST.Location = New-Object System.Drawing.Point(370, 103)
$browseButtonEST.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $textBoxEST.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($browseButtonEST)

# Generate Template button
$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "Populate Template"
$generateButton.Location = New-Object System.Drawing.Point(10, 140)
$generateButton.Size = New-Object System.Drawing.Size(150, 30)
$generateButton.Add_Click({
    $templateDir = $textBoxTemplate.Text
    $estDir = $textBoxEST.Text

    if (-not (Test-Path $templateDir) -or -not (Test-Path $estDir)) {
        [System.Windows.Forms.MessageBox]::Show("Please select valid paths for both the template and EST directory.", "Error", "OK", "Error")
        return
    }

    # Run script and wait
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\PopulateTemplate.ps1`" -TemplateDirectory `"$templateDir`" -EstDirectory `"$estDir`"" -Wait -NoNewWindow

    # Open generated Excel file
    $excelFilePath = Join-Path $templateDir "PowerShell_Template.xlsx"
    if (Test-Path $excelFilePath) {
        Start-Process -FilePath $excelFilePath
    } else {
        [System.Windows.Forms.MessageBox]::Show("The Excel template was not found at: $excelFilePath", "Error", "OK", "Error")
    }
})
$form.Controls.Add($generateButton)

# Divider
$divider = New-Object System.Windows.Forms.Label
$divider.BorderStyle = 'Fixed3D'
$divider.AutoSize = $false
$divider.Height = 2
$divider.Width = 510
$divider.Location = New-Object System.Drawing.Point(10, 185)
$form.Controls.Add($divider)

# Dry Run button
$dryRunButton = New-Object System.Windows.Forms.Button
$dryRunButton.Text = "Perform Test Move"
$dryRunButton.Location = New-Object System.Drawing.Point(10, 200)
$dryRunButton.Size = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($dryRunButton)

# Label for results
$labelResults = New-Object System.Windows.Forms.Label
$labelResults.Text = "Test Move Results:"
$labelResults.Location = New-Object System.Drawing.Point(10, 250)
$labelResults.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($labelResults)

# Dry Run Results ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 275)
$listView.Size = New-Object System.Drawing.Size(510, 220)
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Columns.Add("Estimate", 100) | Out-Null
$listView.Columns.Add("Destination Division", 200) | Out-Null
$listView.Columns.Add("Destination Path", 200) | Out-Null
$form.Controls.Add($listView)

# Move Estimate Button
$moveButton = New-Object System.Windows.Forms.Button
$moveButton.Text = "Move Estimates"
$moveButton.Location = New-Object System.Drawing.Point(10, 520)
$moveButton.Size = New-Object System.Drawing.Size(150, 30)
$moveButton.Add_Click({
    $templateDir = $textBoxTemplate.Text
    $dryRun = $dryRunCheckbox.Checked

    if (-not (Test-Path $templateDir)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid template directory.", "Error", "OK", "Error")
        return
    }

    $paramArgs = "-TemplateDirectory `"$templateDir`""
    if ($dryRun) {
        $paramArgs += " -DryRun"
    }

    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\MoveEstimates.ps1`" $paramArgs" -Wait -NoNewWindow

    # Pop-up message after completion
    [System.Windows.Forms.MessageBox]::Show("Estimates have been successfully moved.", "Move Complete", "OK", "Information")
})
$form.Controls.Add($moveButton)

# Dry Run click event
$dryRunButton.Add_Click({
    $listView.Items.Clear()

    # Ensure file exists
    $templatePath = Join-Path $textBoxTemplate.Text "PowerShell_Template.xlsx"
    if (-Not (Test-Path $templatePath)) {
        [System.Windows.Forms.MessageBox]::Show("Template file not found at: $templatePath")
        return
    }

    try {
        $data = Import-Excel -Path $templatePath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error reading Excel file. Make sure it is not open and has valid headers.")
        return
    }

    foreach ($row in $data) {
        $estimate = $row."Estimate Code"
        $division = $row."Enterprise Code"
        $destination = $row."New Pathing"

        if ($estimate -and $division -and $destination) {
            $item = New-Object System.Windows.Forms.ListViewItem($estimate)
            $item.SubItems.Add($division)
            $item.SubItems.Add($destination)
            $listView.Items.Add($item) | Out-Null
        }
    }

    $listView.AutoResizeColumns("Header")
})

# Show the form
[void]$form.ShowDialog()