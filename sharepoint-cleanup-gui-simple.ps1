# SharePoint Empty Folder Cleanup GUI - Simple Direct Call Version
# This GUI directly launches the CLI script in a new PowerShell window

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint Cleanup Tool - Simple GUI"
$form.Size = New-Object System.Drawing.Size(550, 400)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Title label
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Location = New-Object System.Drawing.Point(20, 20)
$lblTitle.Size = New-Object System.Drawing.Size(500, 30)
$lblTitle.Text = "SharePoint Empty Folder Cleanup Tool"
$lblTitle.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::Navy
$form.Controls.Add($lblTitle)

# Instructions
$lblInstructions = New-Object System.Windows.Forms.Label
$lblInstructions.Location = New-Object System.Drawing.Point(20, 60)
$lblInstructions.Size = New-Object System.Drawing.Size(500, 40)
$lblInstructions.Text = "Fill in the details below and click 'Launch Cleanup Tool'.`nA new PowerShell window will open to handle authentication and cleanup."
$form.Controls.Add($lblInstructions)

# Site URL
$lblSite = New-Object System.Windows.Forms.Label
$lblSite.Location = New-Object System.Drawing.Point(20, 110)
$lblSite.Size = New-Object System.Drawing.Size(100, 20)
$lblSite.Text = "Site URL:"
$form.Controls.Add($lblSite)

$txtSite = New-Object System.Windows.Forms.TextBox
$txtSite.Location = New-Object System.Drawing.Point(130, 110)
$txtSite.Size = New-Object System.Drawing.Size(380, 20)
$txtSite.Text = "https://yourtenant.sharepoint.com/sites/yoursite"
$form.Controls.Add($txtSite)

# Library name
$lblLibrary = New-Object System.Windows.Forms.Label
$lblLibrary.Location = New-Object System.Drawing.Point(20, 140)
$lblLibrary.Size = New-Object System.Drawing.Size(100, 20)
$lblLibrary.Text = "Library Name:"
$form.Controls.Add($lblLibrary)

$txtLibrary = New-Object System.Windows.Forms.TextBox
$txtLibrary.Location = New-Object System.Drawing.Point(130, 140)
$txtLibrary.Size = New-Object System.Drawing.Size(200, 20)
$txtLibrary.Text = "Documents"
$form.Controls.Add($txtLibrary)

# Date picker
$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Location = New-Object System.Drawing.Point(20, 170)
$lblDate.Size = New-Object System.Drawing.Size(100, 20)
$lblDate.Text = "Modified Date:"
$form.Controls.Add($lblDate)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = New-Object System.Drawing.Point(130, 170)
$datePicker.Size = New-Object System.Drawing.Size(200, 20)
$datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$form.Controls.Add($datePicker)

# Preview checkbox
$chkPreview = New-Object System.Windows.Forms.CheckBox
$chkPreview.Location = New-Object System.Drawing.Point(130, 200)
$chkPreview.Size = New-Object System.Drawing.Size(300, 20)
$chkPreview.Text = "Preview Mode (scan only, don't delete)"
$chkPreview.Checked = $true
$chkPreview.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$chkPreview.ForeColor = [System.Drawing.Color]::DarkGreen
$form.Controls.Add($chkPreview)

# Info box
$grpInfo = New-Object System.Windows.Forms.GroupBox
$grpInfo.Location = New-Object System.Drawing.Point(20, 230)
$grpInfo.Size = New-Object System.Drawing.Size(490, 80)
$grpInfo.Text = "Information"
$form.Controls.Add($grpInfo)

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Location = New-Object System.Drawing.Point(10, 20)
$lblInfo.Size = New-Object System.Drawing.Size(470, 50)
$lblInfo.Text = "• A new PowerShell window will open when you click Launch`n• You'll be prompted to choose authentication method`n• Follow the prompts in the PowerShell window"
$grpInfo.Controls.Add($lblInfo)

# Launch button
$btnLaunch = New-Object System.Windows.Forms.Button
$btnLaunch.Location = New-Object System.Drawing.Point(130, 325)
$btnLaunch.Size = New-Object System.Drawing.Size(150, 35)
$btnLaunch.Text = "Launch Cleanup Tool"
$btnLaunch.BackColor = [System.Drawing.Color]::LightGreen
$btnLaunch.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnLaunch)

# Close button
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(290, 325)
$btnClose.Size = New-Object System.Drawing.Size(100, 35)
$btnClose.Text = "Close"
$form.Controls.Add($btnClose)

# Launch button click
$btnLaunch.Add_Click({
    # Validate inputs
    $siteUrl = $txtSite.Text.Trim()
    $libraryName = $txtLibrary.Text.Trim()
    $modifiedDate = $datePicker.Value.ToString("yyyy-MM-dd")
    
    if ([string]::IsNullOrWhiteSpace($siteUrl) -or $siteUrl -eq "https://yourtenant.sharepoint.com/sites/yoursite") {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid SharePoint site URL", 
            "Missing Information", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($libraryName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a library name", 
            "Missing Information", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # Build the command
    $scriptPath = Join-Path $PWD "sharepoint-cleanup-script.ps1"
    
    # Check if script exists
    if (-not (Test-Path $scriptPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "sharepoint-cleanup-script.ps1 not found in current directory!", 
            "Script Not Found", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    
    # Build the PowerShell command
    $whatIfParam = if ($chkPreview.Checked) { "-WhatIf" } else { "-WhatIf:`$false" }
    
    # Create the command to run
    $command = "& '$scriptPath' -SiteUrl '$siteUrl' -LibraryName '$libraryName' -ModifiedDate '$modifiedDate' $whatIfParam"
    
    # Show what will be executed
    $result = [System.Windows.Forms.MessageBox]::Show(
        "The following command will be executed in a new PowerShell window:`n`n$command`n`nProceed?", 
        "Confirm Execution", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Launch in a new PowerShell window
        Start-Process powershell -ArgumentList "-NoExit", "-ExecutionPolicy", "Bypass", "-Command", $command
        
        [System.Windows.Forms.MessageBox]::Show(
            "PowerShell window launched!`n`nFollow the prompts in the new window for authentication.", 
            "Launched", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# Close button click
$btnClose.Add_Click({
    $form.Close()
})

# Show the form
$form.ShowDialog()