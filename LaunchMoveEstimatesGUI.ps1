Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Move Estimates Launcher"
$form.Size = New-Object System.Drawing.Size(550, 330)
$form.StartPosition = "CenterScreen"

# === TEMPLATE SECTION ===

# Template directory label
$labelTemplate = New-Object System.Windows.Forms.Label
$labelTemplate.Text = "Select Directory where PowerShell_Template.xlsx is located:"
$labelTemplate.Location = New-Object System.Drawing.Point(10, 20)
$labelTemplate.Size = New-Object System.Drawing.Size(200, 20)
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
$labelEST.Text = "Select EST Directory:"
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
$generateButton.Text = "Generate Template"
$generateButton.Location = New-Object System.Drawing.Point(10, 140)
$generateButton.Size = New-Object System.Drawing.Size(150, 30)
$generateButton.Add_Click({
    $templateDir = $textBoxTemplate.Text
    $estDir = $textBoxEST.Text

    if (-not (Test-Path $templateDir) -or -not (Test-Path $estDir)) {
        [System.Windows.Forms.MessageBox]::Show("Please select valid paths for both the template and EST directory.", "Error", "OK", "Error")
        return
    }

    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\PopulateTemplate.ps1`" -TemplateDirectory `"$templateDir`" -EstDirectory `"$estDir`"" -Wait -NoNewWindow
})
$form.Controls.Add($generateButton)

# === MOVE SECTION ===

# Dry Run checkbox
$dryRunCheckbox = New-Object System.Windows.Forms.CheckBox
$dryRunCheckbox.Text = "Dry Run (for testing - no files moved)"
$dryRunCheckbox.Location = New-Object System.Drawing.Point(10, 190)
$dryRunCheckbox.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($dryRunCheckbox)

# Move Estimates button
$moveButton = New-Object System.Windows.Forms.Button
$moveButton.Text = "Move Estimates"
$moveButton.Location = New-Object System.Drawing.Point(10, 220)
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
})
$form.Controls.Add($moveButton)

# Show the form
$form.ShowDialog()