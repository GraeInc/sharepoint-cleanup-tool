# SharePoint Empty Folder Cleanup GUI - Browser Auth Version
# This version opens the browser immediately without freezing

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Import module
Import-Module PnP.PowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# Global variables
$script:IsConnected = $false

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint Cleanup Tool"
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

# Connect button click - spawns separate PowerShell for auth
$btnConnect.Add_Click({
    $siteUrl = $txtSite.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a site URL", "Error")
        return
    }
    
    $btnConnect.Enabled = $false
    $lblStatus.Text = "Connecting... Check the PowerShell window that opens"
    $lblStatus.ForeColor = [System.Drawing.Color]::Orange
    $form.Refresh()
    
    # Create a temporary script that will handle authentication
    $tempScript = [System.IO.Path]::GetTempFileName() + ".ps1"
    @"
Import-Module PnP.PowerShell
Write-Host 'Connecting to SharePoint...' -ForegroundColor Yellow
Write-Host 'A browser window will open for authentication.' -ForegroundColor Cyan

try {
    # Try the method that works in your CLI
    Connect-PnPOnline -Url '$siteUrl' -UseWebLogin
    `$web = Get-PnPWeb
    Write-Host "Successfully connected to: `$(`$web.Title)" -ForegroundColor Green
    
    # Save connection status to temp file
    `$result = @{
        Success = `$true
        Title = `$web.Title
        Url = `$web.Url
    }
    `$result | Export-CliXml -Path '$env:TEMP\sp_connection.xml'
}
catch {
    Write-Host "Connection failed: `$_" -ForegroundColor Red
    `$result = @{
        Success = `$false
        Error = `$_.Exception.Message
    }
    `$result | Export-CliXml -Path '$env:TEMP\sp_connection.xml'
}

Write-Host ""
Write-Host "You can close this window and return to the GUI" -ForegroundColor Yellow
Start-Sleep -Seconds 3
"@ | Out-File -FilePath $tempScript -Encoding UTF8
    
    # Start authentication in separate window
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$tempScript`"" -Wait
    
    # Check result
    if (Test-Path "$env:TEMP\sp_connection.xml") {
        $result = Import-CliXml -Path "$env:TEMP\sp_connection.xml"
        Remove-Item "$env:TEMP\sp_connection.xml" -Force
        
        if ($result.Success) {
            $script:IsConnected = $true
            $lblStatus.Text = "Connected to: $($result.Title)"
            $lblStatus.ForeColor = [System.Drawing.Color]::Green
            $txtLibrary.Enabled = $true
            $datePicker.Enabled = $true
            $btnScan.Enabled = $true
            $btnConnect.Text = "Connected"
            $btnConnect.BackColor = [System.Drawing.Color]::LightGray
            $txtResults.AppendText("Successfully connected!`r`n")
            
            # Now connect in this session too - MUST reconnect for this session
            $txtResults.AppendText("Reconnecting in this session...`r`n")
            try {
                # Open connection in THIS PowerShell session
                Connect-PnPOnline -Url $siteUrl -UseWebLogin
                $txtResults.AppendText("Session connected successfully!`r`n")
            } catch {
                $txtResults.AppendText("Warning: Could not establish connection in this session: $_`r`n")
                $script:IsConnected = $false
                $lblStatus.Text = "Connection failed - please try again"
                $lblStatus.ForeColor = [System.Drawing.Color]::Red
                $btnConnect.Enabled = $true
            }
        }
        else {
            $lblStatus.Text = "Connection failed"
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $btnConnect.Enabled = $true
            $txtResults.AppendText("Connection failed: $($result.Error)`r`n")
        }
    }
    else {
        $lblStatus.Text = "Connection cancelled or failed"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        $btnConnect.Enabled = $true
    }
    
    # Clean up temp script
    if (Test-Path $tempScript) {
        Remove-Item $tempScript -Force
    }
})

# Scan button click
$btnScan.Add_Click({
    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect first", "Not Connected")
        return
    }
    
    $libraryName = $txtLibrary.Text
    $modifiedDate = $datePicker.Value
    
    $txtResults.Clear()
    $txtResults.AppendText("Scanning for empty folders...`r`n")
    $txtResults.AppendText("Library: $libraryName`r`n")
    $txtResults.AppendText("Date: $($modifiedDate.ToShortDateString())`r`n")
    $txtResults.AppendText("---------------------------------`r`n")
    
    try {
        # Build query for folders modified on specific date
        $camlQuery = @"
<View>
    <Query>
        <Where>
            <And>
                <Eq>
                    <FieldRef Name='FSObjType'/>
                    <Value Type='Integer'>1</Value>
                </Eq>
                <Eq>
                    <FieldRef Name='Modified'/>
                    <Value Type='DateTime'>$($modifiedDate.ToString('yyyy-MM-dd'))</Value>
                </Eq>
            </And>
        </Where>
    </Query>
</View>
"@
        
        $folders = Get-PnPListItem -List $libraryName -Query $camlQuery
        $txtResults.AppendText("Found $($folders.Count) folders modified on this date`r`n")
        
        $emptyCount = 0
        foreach ($folder in $folders) {
            $folderPath = $folder.FieldValues.FileRef
            $folderName = $folder.FieldValues.FileLeafRef
            
            $folderContents = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType All -ErrorAction SilentlyContinue
            if ($folderContents.Count -eq 0) {
                $emptyCount++
                $txtResults.AppendText("Empty: $folderName`r`n")
            }
        }
        
        $txtResults.AppendText("`r`nTotal empty folders: $emptyCount`r`n")
    }
    catch {
        $txtResults.AppendText("Error: $_`r`n")
    }
})

# Show form
$form.ShowDialog()