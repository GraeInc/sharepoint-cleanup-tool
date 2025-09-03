# SharePoint Empty Folder Cleanup GUI - Async Version
# Uses background runspace to prevent freezing during authentication

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Import module
Import-Module PnP.PowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# Global variables
$script:IsConnected = $false
$script:SiteUrl = ""
$script:RunspacePool = $null
$script:PowerShell = $null

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint Cleanup Tool - Async"
$form.Size = New-Object System.Drawing.Size(500, 400)
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
$txtSite.Size = New-Object System.Drawing.Size(350, 20)
$txtSite.Text = "https://yourtenant.sharepoint.com/sites/yoursite"
$form.Controls.Add($txtSite)

# Connect button
$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(110, 50)
$btnConnect.Size = New-Object System.Drawing.Size(100, 30)
$btnConnect.Text = "Connect"
$btnConnect.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($btnConnect)

# Status label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 90)
$lblStatus.Size = New-Object System.Drawing.Size(440, 40)
$lblStatus.Text = "Not connected"
$lblStatus.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblStatus)

# Library name
$lblLibrary = New-Object System.Windows.Forms.Label
$lblLibrary.Location = New-Object System.Drawing.Point(20, 140)
$lblLibrary.Size = New-Object System.Drawing.Size(80, 20)
$lblLibrary.Text = "Library:"
$form.Controls.Add($lblLibrary)

$txtLibrary = New-Object System.Windows.Forms.TextBox
$txtLibrary.Location = New-Object System.Drawing.Point(110, 140)
$txtLibrary.Size = New-Object System.Drawing.Size(200, 20)
$txtLibrary.Text = "Documents"
$txtLibrary.Enabled = $false
$form.Controls.Add($txtLibrary)

# Date picker
$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Location = New-Object System.Drawing.Point(20, 170)
$lblDate.Size = New-Object System.Drawing.Size(80, 20)
$lblDate.Text = "Modified Date:"
$form.Controls.Add($lblDate)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = New-Object System.Drawing.Point(110, 170)
$datePicker.Size = New-Object System.Drawing.Size(200, 20)
$datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$datePicker.Enabled = $false
$form.Controls.Add($datePicker)

# Scan button
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Location = New-Object System.Drawing.Point(110, 210)
$btnScan.Size = New-Object System.Drawing.Size(100, 30)
$btnScan.Text = "Scan"
$btnScan.BackColor = [System.Drawing.Color]::LightBlue
$btnScan.Enabled = $false
$form.Controls.Add($btnScan)

# Results text box
$txtResults = New-Object System.Windows.Forms.TextBox
$txtResults.Location = New-Object System.Drawing.Point(20, 250)
$txtResults.Size = New-Object System.Drawing.Size(440, 100)
$txtResults.Multiline = $true
$txtResults.ScrollBars = "Vertical"
$txtResults.ReadOnly = $true
$form.Controls.Add($txtResults)

# Connect button click - using simpler approach
$btnConnect.Add_Click({
    $siteUrl = $txtSite.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a site URL", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $btnConnect.Enabled = $false
    $lblStatus.Text = "Connecting... Browser window will open for authentication"
    $lblStatus.ForeColor = [System.Drawing.Color]::Orange
    $form.Refresh()
    
    # Run connection in background to prevent GUI freeze
    $script:SiteUrl = $siteUrl
    
    # Create a timer to check connection status
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 100
    
    # Start authentication in background
    $connectionJob = Start-Job -ScriptBlock {
        param($url)
        Import-Module PnP.PowerShell
        
        try {
            # Try different methods in order
            try {
                Connect-PnPOnline -Url $url -Interactive -ErrorAction Stop
                return @{Success=$true; Method="Interactive"; Title=(Get-PnPWeb).Title}
            }
            catch {
                try {
                    Connect-PnPOnline -Url $url -UseWebLogin -ErrorAction Stop
                    return @{Success=$true; Method="UseWebLogin"; Title=(Get-PnPWeb).Title}
                }
                catch {
                    try {
                        Connect-PnPOnline -Url $url -LaunchBrowser -ErrorAction Stop
                        return @{Success=$true; Method="LaunchBrowser"; Title=(Get-PnPWeb).Title}
                    }
                    catch {
                        return @{Success=$false; Error=$_.Exception.Message}
                    }
                }
            }
        }
        catch {
            return @{Success=$false; Error=$_.Exception.Message}
        }
    } -ArgumentList $siteUrl
    
    $timer.Add_Tick({
        if ($connectionJob.State -eq "Completed") {
            $result = Receive-Job -Job $connectionJob
            Remove-Job -Job $connectionJob
            $timer.Stop()
            
            if ($result.Success) {
                $script:IsConnected = $true
                $lblStatus.Text = "Connected to: $($result.Title) (Method: $($result.Method))"
                $lblStatus.ForeColor = [System.Drawing.Color]::Green
                $txtLibrary.Enabled = $true
                $datePicker.Enabled = $true
                $btnScan.Enabled = $true
                $btnConnect.Text = "Connected"
                $btnConnect.BackColor = [System.Drawing.Color]::LightGray
                $txtResults.AppendText("Successfully connected using $($result.Method)`r`n")
            }
            else {
                $lblStatus.Text = "Connection failed"
                $lblStatus.ForeColor = [System.Drawing.Color]::Red
                $btnConnect.Enabled = $true
                $txtResults.AppendText("Connection failed: $($result.Error)`r`n")
                [System.Windows.Forms.MessageBox]::Show("Failed to connect: $($result.Error)", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        elseif ($connectionJob.State -eq "Failed") {
            $timer.Stop()
            $lblStatus.Text = "Connection failed"
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $btnConnect.Enabled = $true
            Remove-Job -Job $connectionJob
        }
    })
    
    $timer.Start()
})

# Scan button click
$btnScan.Add_Click({
    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect first", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $libraryName = $txtLibrary.Text
    $modifiedDate = $datePicker.Value
    
    $txtResults.Clear()
    $txtResults.AppendText("Scanning for empty folders...`r`n")
    $txtResults.AppendText("Library: $libraryName`r`n")
    $txtResults.AppendText("Date: $($modifiedDate.ToShortDateString())`r`n")
    $txtResults.AppendText("---------------------------------`r`n")
    
    # Here you would add the actual scanning logic
    # For now, just show a message
    $txtResults.AppendText("Scan functionality would go here`r`n")
})

# Show form
$form.ShowDialog()