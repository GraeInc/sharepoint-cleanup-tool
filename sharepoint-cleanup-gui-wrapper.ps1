# SharePoint Empty Folder Cleanup GUI - CLI Wrapper Version
# This GUI wraps the working CLI script for authentication and operations

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Global variables
$script:IsConnected = $false
$script:CurrentJob = $null
$script:EmptyFolders = @()

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint Cleanup Tool - GUI Wrapper"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Site URL
$lblSite = New-Object System.Windows.Forms.Label
$lblSite.Location = New-Object System.Drawing.Point(20, 20)
$lblSite.Size = New-Object System.Drawing.Size(80, 20)
$lblSite.Text = "Site URL:"
$form.Controls.Add($lblSite)

$txtSite = New-Object System.Windows.Forms.TextBox
$txtSite.Location = New-Object System.Drawing.Point(110, 20)
$txtSite.Size = New-Object System.Drawing.Size(450, 20)
$txtSite.Text = "https://yourtenant.sharepoint.com/sites/yoursite"
$form.Controls.Add($txtSite)

# Library name
$lblLibrary = New-Object System.Windows.Forms.Label
$lblLibrary.Location = New-Object System.Drawing.Point(20, 50)
$lblLibrary.Size = New-Object System.Drawing.Size(80, 20)
$lblLibrary.Text = "Library:"
$form.Controls.Add($lblLibrary)

$txtLibrary = New-Object System.Windows.Forms.TextBox
$txtLibrary.Location = New-Object System.Drawing.Point(110, 50)
$txtLibrary.Size = New-Object System.Drawing.Size(200, 20)
$txtLibrary.Text = "Documents"
$form.Controls.Add($txtLibrary)

# Date picker
$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Location = New-Object System.Drawing.Point(20, 80)
$lblDate.Size = New-Object System.Drawing.Size(80, 20)
$lblDate.Text = "Modified Date:"
$form.Controls.Add($lblDate)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = New-Object System.Drawing.Point(110, 80)
$datePicker.Size = New-Object System.Drawing.Size(200, 20)
$datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$form.Controls.Add($datePicker)

# Preview checkbox
$chkPreview = New-Object System.Windows.Forms.CheckBox
$chkPreview.Location = New-Object System.Drawing.Point(330, 80)
$chkPreview.Size = New-Object System.Drawing.Size(200, 20)
$chkPreview.Text = "Preview Mode (No Deletion)"
$chkPreview.Checked = $true
$form.Controls.Add($chkPreview)

# Auth method
$lblAuth = New-Object System.Windows.Forms.Label
$lblAuth.Location = New-Object System.Drawing.Point(20, 110)
$lblAuth.Size = New-Object System.Drawing.Size(80, 20)
$lblAuth.Text = "Auth Method:"
$form.Controls.Add($lblAuth)

$cmbAuth = New-Object System.Windows.Forms.ComboBox
$cmbAuth.Location = New-Object System.Drawing.Point(110, 110)
$cmbAuth.Size = New-Object System.Drawing.Size(200, 20)
$cmbAuth.DropDownStyle = "DropDownList"
$cmbAuth.Items.AddRange(@("Browser Login", "Credentials"))
$cmbAuth.SelectedIndex = 0
$form.Controls.Add($cmbAuth)

# Username/Password (hidden by default)
$lblUsername = New-Object System.Windows.Forms.Label
$lblUsername.Location = New-Object System.Drawing.Point(20, 140)
$lblUsername.Size = New-Object System.Drawing.Size(80, 20)
$lblUsername.Text = "Username:"
$lblUsername.Visible = $false
$form.Controls.Add($lblUsername)

$txtUsername = New-Object System.Windows.Forms.TextBox
$txtUsername.Location = New-Object System.Drawing.Point(110, 140)
$txtUsername.Size = New-Object System.Drawing.Size(200, 20)
$txtUsername.Visible = $false
$form.Controls.Add($txtUsername)

$lblPassword = New-Object System.Windows.Forms.Label
$lblPassword.Location = New-Object System.Drawing.Point(320, 140)
$lblPassword.Size = New-Object System.Drawing.Size(60, 20)
$lblPassword.Text = "Password:"
$lblPassword.Visible = $false
$form.Controls.Add($lblPassword)

$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Location = New-Object System.Drawing.Point(380, 140)
$txtPassword.Size = New-Object System.Drawing.Size(180, 20)
$txtPassword.PasswordChar = "*"
$txtPassword.Visible = $false
$form.Controls.Add($txtPassword)

# Run button
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Location = New-Object System.Drawing.Point(110, 170)
$btnRun.Size = New-Object System.Drawing.Size(120, 35)
$btnRun.Text = "Run Scan"
$btnRun.BackColor = [System.Drawing.Color]::LightGreen
$btnRun.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

# Cancel button
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Location = New-Object System.Drawing.Point(240, 170)
$btnCancel.Size = New-Object System.Drawing.Size(100, 35)
$btnCancel.Text = "Cancel"
$btnCancel.BackColor = [System.Drawing.Color]::LightCoral
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

# Clear button
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(350, 170)
$btnClear.Size = New-Object System.Drawing.Size(100, 35)
$btnClear.Text = "Clear Output"
$form.Controls.Add($btnClear)

# Output text box
$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(20, 220)
$txtOutput.Size = New-Object System.Drawing.Size(540, 200)
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = "Vertical"
$txtOutput.ReadOnly = $true
$txtOutput.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOutput.BackColor = [System.Drawing.Color]::Black
$txtOutput.ForeColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($txtOutput)

# Status label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 430)
$lblStatus.Size = New-Object System.Drawing.Size(540, 20)
$lblStatus.Text = "Ready"
$lblStatus.ForeColor = [System.Drawing.Color]::Blue
$form.Controls.Add($lblStatus)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 455)
$progressBar.Size = New-Object System.Drawing.Size(540, 20)
$progressBar.Style = "Marquee"
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Auth method change event
$cmbAuth.Add_SelectedIndexChanged({
    if ($cmbAuth.SelectedItem -eq "Credentials") {
        $lblUsername.Visible = $true
        $txtUsername.Visible = $true
        $lblPassword.Visible = $true
        $txtPassword.Visible = $true
    } else {
        $lblUsername.Visible = $false
        $txtUsername.Visible = $false
        $lblPassword.Visible = $false
        $txtPassword.Visible = $false
    }
})

# Clear button click
$btnClear.Add_Click({
    $txtOutput.Clear()
    $lblStatus.Text = "Output cleared"
})

# Cancel button click
$btnCancel.Add_Click({
    if ($script:CurrentJob) {
        Stop-Job -Job $script:CurrentJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:CurrentJob -Force -ErrorAction SilentlyContinue
        $script:CurrentJob = $null
        $txtOutput.AppendText("`r`n`r`n*** CANCELLED BY USER ***`r`n")
        $lblStatus.Text = "Operation cancelled"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        $progressBar.Visible = $false
        $btnRun.Enabled = $true
        $btnCancel.Enabled = $false
    }
})

# Run button click
$btnRun.Add_Click({
    # Validate inputs
    $siteUrl = $txtSite.Text.Trim()
    $libraryName = $txtLibrary.Text.Trim()
    $modifiedDate = $datePicker.Value.ToString("yyyy-MM-dd")
    
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a SharePoint site URL", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($libraryName)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a library name", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Prepare WhatIf parameter
    $whatIfParam = if ($chkPreview.Checked) { '$true' } else { '$false' }
    
    # Clear output
    $txtOutput.Clear()
    $txtOutput.AppendText("========================================`r`n")
    $txtOutput.AppendText("SharePoint Empty Folder Cleanup`r`n")
    $txtOutput.AppendText("========================================`r`n")
    $txtOutput.AppendText("Site: $siteUrl`r`n")
    $txtOutput.AppendText("Library: $libraryName`r`n")
    $txtOutput.AppendText("Date: $modifiedDate`r`n")
    $txtOutput.AppendText("Preview Mode: $($chkPreview.Checked)`r`n")
    $txtOutput.AppendText("========================================`r`n`r`n")
    
    # Update UI
    $btnRun.Enabled = $false
    $btnCancel.Enabled = $true
    $progressBar.Visible = $true
    $lblStatus.Text = "Running scan..."
    $lblStatus.ForeColor = [System.Drawing.Color]::Blue
    $form.Refresh()
    
    # Build the PowerShell command to run the CLI script
    $scriptPath = Join-Path $PWD "sharepoint-cleanup-script.ps1"
    
    # Check if script exists
    if (-not (Test-Path $scriptPath)) {
        $txtOutput.AppendText("ERROR: sharepoint-cleanup-script.ps1 not found!`r`n")
        $txtOutput.AppendText("Expected at: $scriptPath`r`n")
        $lblStatus.Text = "Script not found"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        $progressBar.Visible = $false
        $btnRun.Enabled = $true
        $btnCancel.Enabled = $false
        return
    }
    
    # Build the command based on authentication method
    if ($cmbAuth.SelectedItem -eq "Credentials" -and $txtUsername.Text -and $txtPassword.Text) {
        # Create a script block that will handle credential input
        $scriptBlock = {
            param($ScriptPath, $SiteUrl, $LibraryName, $ModifiedDate, $WhatIf, $Username, $Password)
            
            # Create a temporary script that feeds the credentials
            $tempScript = @"
`$username = '$Username'
`$password = ConvertTo-SecureString '$Password' -AsPlainText -Force
Write-Output 'y'
Write-Output `$username
Write-Output `$password | ConvertFrom-SecureString
& '$ScriptPath' -SiteUrl '$SiteUrl' -LibraryName '$LibraryName' -ModifiedDate '$ModifiedDate' -WhatIf:`$$WhatIf
"@
            Invoke-Expression $tempScript
        }
        
        $script:CurrentJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList $scriptPath, $siteUrl, $libraryName, $modifiedDate, $whatIfParam, $txtUsername.Text, $txtPassword.Text
    }
    else {
        # Use browser authentication (default)
        $txtOutput.AppendText("Using browser authentication...`r`n")
        $txtOutput.AppendText("A browser window will open for login.`r`n")
        $txtOutput.AppendText("Please complete the authentication in the browser.`r`n`r`n")
        $form.Refresh()
        
        # Create script block for browser auth
        $scriptBlock = {
            param($ScriptPath, $SiteUrl, $LibraryName, $ModifiedDate, $WhatIfBool)
            
            # Import the module first
            Import-Module PnP.PowerShell -ErrorAction SilentlyContinue
            
            # Create a temp file with the input
            $tempInput = [System.IO.Path]::GetTempFileName()
            "n" | Out-File -FilePath $tempInput -Encoding ASCII
            
            # Build the command with proper boolean
            if ($WhatIfBool -eq "true") {
                Get-Content $tempInput | & $ScriptPath -SiteUrl $SiteUrl -LibraryName $LibraryName -ModifiedDate $ModifiedDate -WhatIf:$true
            } else {
                Get-Content $tempInput | & $ScriptPath -SiteUrl $SiteUrl -LibraryName $LibraryName -ModifiedDate $ModifiedDate -WhatIf:$false
            }
            
            # Clean up temp file
            Remove-Item $tempInput -Force -ErrorAction SilentlyContinue
        }
        
        $whatIfBool = if ($chkPreview.Checked) { "true" } else { "false" }
        $script:CurrentJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList $scriptPath, $siteUrl, $libraryName, $modifiedDate, $whatIfBool
    }
    
    # Create a timer to check job status
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $timer.Add_Tick({
        if ($script:CurrentJob) {
            # Get any new output
            $output = Receive-Job -Job $script:CurrentJob -Keep
            if ($output) {
                foreach ($line in $output) {
                    $txtOutput.AppendText("$line`r`n")
                }
                $txtOutput.SelectionStart = $txtOutput.Text.Length
                $txtOutput.ScrollToCaret()
            }
            
            # Check if job is complete
            if ($script:CurrentJob.State -eq "Completed") {
                # Get final output
                $finalOutput = Receive-Job -Job $script:CurrentJob
                foreach ($line in $finalOutput) {
                    $txtOutput.AppendText("$line`r`n")
                }
                
                Remove-Job -Job $script:CurrentJob
                $script:CurrentJob = $null
                $timer.Stop()
                
                $txtOutput.AppendText("`r`n========================================`r`n")
                $txtOutput.AppendText("Operation completed successfully!`r`n")
                $txtOutput.AppendText("========================================`r`n")
                
                $lblStatus.Text = "Operation completed"
                $lblStatus.ForeColor = [System.Drawing.Color]::Green
                $progressBar.Visible = $false
                $btnRun.Enabled = $true
                $btnCancel.Enabled = $false
            }
            elseif ($script:CurrentJob.State -eq "Failed") {
                $error = $script:CurrentJob.ChildJobs[0].JobStateInfo.Reason.Message
                $txtOutput.AppendText("`r`nERROR: $error`r`n")
                
                Remove-Job -Job $script:CurrentJob
                $script:CurrentJob = $null
                $timer.Stop()
                
                $lblStatus.Text = "Operation failed"
                $lblStatus.ForeColor = [System.Drawing.Color]::Red
                $progressBar.Visible = $false
                $btnRun.Enabled = $true
                $btnCancel.Enabled = $false
            }
        }
        else {
            $timer.Stop()
        }
    })
    $timer.Start()
})

# Form closing event
$form.Add_FormClosing({
    if ($script:CurrentJob) {
        Stop-Job -Job $script:CurrentJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:CurrentJob -Force -ErrorAction SilentlyContinue
    }
})

# Show the form
$form.ShowDialog()