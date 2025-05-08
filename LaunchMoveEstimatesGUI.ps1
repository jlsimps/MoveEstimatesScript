Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Move Estimates Launcher"
$form.Size = New-Object System.Drawing.Size(550,220)
$form.StartPosition = "CenterScreen"

# --- Label for folder path
$label = New-Object System.Windows.Forms.Label
$label.Text = "Select Template Directory:"
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($label)

# --- Textbox for folder path
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,45)
$textBox.Size = New-Object System.Drawing.Size(350,20)
$form.Controls.Add($textBox)

# --- Browse button
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse"
$browseButton.Location = New-Object System.Drawing.Point(370,43)
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $textBox.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($browseButton)

# --- Dry Run checkbox
$dryRunCheckbox = New-Object System.Windows.Forms.CheckBox
$dryRunCheckbox.Text = "Dry Run (for testing purposes - no files will be moved)"
$dryRunCheckbox.Location = New-Object System.Drawing.Point(10,80)
$dryRunCheckbox.Size = New-Object System.Drawing.Size(500, 20)
$form.Controls.Add($dryRunCheckbox)

# --- Run button
$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run Script"
$runButton.Location = New-Object System.Drawing.Point(10,120)
$runButton.Add_Click({
    $templateDir = $textBox.Text
    $dryRun = $dryRunCheckbox.Checked

    if (-not (Test-Path $templateDir)) {
        [System.Windows.Forms.MessageBox]::Show("Invalid directory path.", "Error", "OK", "Error")
        return
    }

    $paramArgs = "-TemplateDirectory `"$templateDir`""
    if ($dryRun) {
        $paramArgs += " -DryRun"
    }

    # Call the main script
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\MoveEstimates.ps1`" $paramArgs" -Wait -NoNewWindow
    $form.Close()
})
$form.Controls.Add($runButton)

# --- Show the form
$form.ShowDialog()