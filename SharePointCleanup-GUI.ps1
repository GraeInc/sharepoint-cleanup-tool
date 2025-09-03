# SharePointCleanup-GUI.ps1
# Simplified GUI for SharePoint Cleanup Tool

param()

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Script variables
$script:BasePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:IsConnected = $false
$script:EmptyFolders = @()
$script:SPConnection = $null

# Initialize PnP Module
try {
    Import-Module PnP.PowerShell -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "PnP.PowerShell module is required.`nPlease run Install.bat first.",
        "Module Missing",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint Cleanup Tool v2.0"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Create tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(770, 510)

# Tab 1: Connection
$tabConnection = New-Object System.Windows.Forms.TabPage("Connection")

$lblSite = New-Object System.Windows.Forms.Label
$lblSite.Location = New-Object System.Drawing.Point(20, 30)
$lblSite.Size = New-Object System.Drawing.Size(80, 20)
$lblSite.Text = "Site URL:"
$tabConnection.Controls.Add($lblSite)

$txtSiteUrl = New-Object System.Windows.Forms.TextBox
$txtSiteUrl.Location = New-Object System.Drawing.Point(110, 30)
$txtSiteUrl.Size = New-Object System.Drawing.Size(500, 20)
$txtSiteUrl.Text = "https://yourtenant.sharepoint.com/sites/yoursite"
$tabConnection.Controls.Add($txtSiteUrl)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(110, 60)
$btnConnect.Size = New-Object System.Drawing.Size(100, 30)
$btnConnect.Text = "Connect"
$btnConnect.BackColor = [System.Drawing.Color]::LightGreen
$tabConnection.Controls.Add($btnConnect)

$lblConnStatus = New-Object System.Windows.Forms.Label
$lblConnStatus.Location = New-Object System.Drawing.Point(20, 100)
$lblConnStatus.Size = New-Object System.Drawing.Size(700, 60)
$lblConnStatus.Text = "Not connected"
$lblConnStatus.ForeColor = [System.Drawing.Color]::Red
$lblConnStatus.BorderStyle = "FixedSingle"
$lblConnStatus.Padding = New-Object System.Windows.Forms.Padding(5)
$tabConnection.Controls.Add($lblConnStatus)

$tabControl.TabPages.Add($tabConnection)

# Tab 2: Scan & Clean
$tabScan = New-Object System.Windows.Forms.TabPage("Scan & Clean")

$lblLibrary = New-Object System.Windows.Forms.Label
$lblLibrary.Location = New-Object System.Drawing.Point(20, 30)
$lblLibrary.Size = New-Object System.Drawing.Size(80, 20)
$lblLibrary.Text = "Library:"
$tabScan.Controls.Add($lblLibrary)

$txtLibrary = New-Object System.Windows.Forms.TextBox
$txtLibrary.Location = New-Object System.Drawing.Point(110, 30)
$txtLibrary.Size = New-Object System.Drawing.Size(200, 20)
$txtLibrary.Text = "Documents"
$txtLibrary.Enabled = $false
$tabScan.Controls.Add($txtLibrary)

$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Location = New-Object System.Drawing.Point(330, 30)
$lblDate.Size = New-Object System.Drawing.Size(100, 20)
$lblDate.Text = "Modified Date:"
$tabScan.Controls.Add($lblDate)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = New-Object System.Drawing.Point(440, 30)
$datePicker.Size = New-Object System.Drawing.Size(200, 20)
$datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$datePicker.Enabled = $false
$tabScan.Controls.Add($datePicker)

$chkPreview = New-Object System.Windows.Forms.CheckBox
$chkPreview.Location = New-Object System.Drawing.Point(110, 60)
$chkPreview.Size = New-Object System.Drawing.Size(200, 20)
$chkPreview.Text = "Preview Mode (Safe)"
$chkPreview.Checked = $true
$chkPreview.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$chkPreview.ForeColor = [System.Drawing.Color]::DarkGreen
$tabScan.Controls.Add($chkPreview)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Location = New-Object System.Drawing.Point(110, 90)
$btnScan.Size = New-Object System.Drawing.Size(100, 30)
$btnScan.Text = "Scan"
$btnScan.BackColor = [System.Drawing.Color]::LightBlue
$btnScan.Enabled = $false
$tabScan.Controls.Add($btnScan)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Location = New-Object System.Drawing.Point(220, 90)
$btnDelete.Size = New-Object System.Drawing.Size(100, 30)
$btnDelete.Text = "Delete Empty"
$btnDelete.BackColor = [System.Drawing.Color]::LightCoral
$btnDelete.Enabled = $false
$tabScan.Controls.Add($btnDelete)

$dgvResults = New-Object System.Windows.Forms.DataGridView
$dgvResults.Location = New-Object System.Drawing.Point(20, 130)
$dgvResults.Size = New-Object System.Drawing.Size(720, 300)
$dgvResults.AllowUserToAddRows = $false
$dgvResults.AllowUserToDeleteRows = $false
$dgvResults.SelectionMode = 'FullRowSelect'
$dgvResults.MultiSelect = $true
$dgvResults.AutoSizeColumnsMode = 'Fill'
$tabScan.Controls.Add($dgvResults)

$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Location = New-Object System.Drawing.Point(20, 440)
$lblSummary.Size = New-Object System.Drawing.Size(720, 30)
$lblSummary.Text = "No scan performed"
$tabScan.Controls.Add($lblSummary)

$tabControl.TabPages.Add($tabScan)

# Tab 3: Logs
$tabLogs = New-Object System.Windows.Forms.TabPage("Activity Log")

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 20)
$txtLog.Size = New-Object System.Drawing.Size(720, 420)
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = [System.Drawing.Color]::Black
$txtLog.ForeColor = [System.Drawing.Color]::LightGreen
$tabLogs.Controls.Add($txtLog)

$btnClearLog = New-Object System.Windows.Forms.Button
$btnClearLog.Location = New-Object System.Drawing.Point(20, 450)
$btnClearLog.Size = New-Object System.Drawing.Size(100, 25)
$btnClearLog.Text = "Clear Log"
$tabLogs.Controls.Add($btnClearLog)

$tabControl.TabPages.Add($tabLogs)

$form.Controls.Add($tabControl)

# Status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel("Ready")
$statusBar.Items.Add($statusLabel)
$form.Controls.Add($statusBar)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 530)
$progressBar.Size = New-Object System.Drawing.Size(770, 20)
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Helper Functions
function Write-Log {
    param($Level, $Message)
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message`r`n"
    
    $color = switch ($Level) {
        "ERROR" { [System.Drawing.Color]::Red }
        "WARNING" { [System.Drawing.Color]::Yellow }
        "SUCCESS" { [System.Drawing.Color]::LightGreen }
        "INFO" { [System.Drawing.Color]::Cyan }
        default { [System.Drawing.Color]::White }
    }
    
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.SelectionColor = $color
    $txtLog.AppendText($logEntry)
    $txtLog.ScrollToCaret()
    
    # Also log to file
    $logDir = Join-Path $script:BasePath "Logs"
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $logFile = Join-Path $logDir "sharepoint-cleanup-$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
}

# Event Handlers
$btnConnect.Add_Click({
    $siteUrl = $txtSiteUrl.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($siteUrl) -or $siteUrl -eq "https://yourtenant.sharepoint.com/sites/yoursite") {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid SharePoint site URL", "Invalid URL")
        return
    }
    
    $btnConnect.Enabled = $false
    $statusLabel.Text = "Connecting..."
    Write-Log "INFO" "Connecting to $siteUrl"
    $form.Refresh()
    
    try {
        # Try different authentication methods
        $connected = $false
        $authMethods = @("DeviceLogin", "Interactive", "UseWebLogin")
        
        foreach ($method in $authMethods) {
            if (-not $connected) {
                try {
                    Write-Log "INFO" "Trying $method authentication..."
                    
                    switch ($method) {
                        "DeviceLogin" { 
                            Connect-PnPOnline -Url $siteUrl -DeviceLogin
                        }
                        "Interactive" { 
                            Connect-PnPOnline -Url $siteUrl -Interactive
                        }
                        "UseWebLogin" { 
                            Connect-PnPOnline -Url $siteUrl -UseWebLogin
                        }
                    }
                    
                    $web = Get-PnPWeb
                    $connected = $true
                    Write-Log "SUCCESS" "Connected using $method"
                    break
                }
                catch {
                    Write-Log "WARNING" "$method failed: $_"
                }
            }
        }
        
        if ($connected) {
            $script:IsConnected = $true
            $script:SPConnection = Get-PnPConnection
            
            $lblConnStatus.Text = "Connected to: $($web.Title)`nURL: $($web.Url)`nAuthentication successful"
            $lblConnStatus.ForeColor = [System.Drawing.Color]::Green
            
            $txtLibrary.Enabled = $true
            $datePicker.Enabled = $true
            $btnScan.Enabled = $true
            $btnConnect.Text = "Reconnect"
            
            $statusLabel.Text = "Connected"
            $tabControl.SelectedIndex = 1
            
            Write-Log "SUCCESS" "Successfully connected to SharePoint"
        }
        else {
            throw "All authentication methods failed"
        }
    }
    catch {
        Write-Log "ERROR" "Connection failed: $_"
        [System.Windows.Forms.MessageBox]::Show("Connection failed: $_", "Error")
        $statusLabel.Text = "Connection failed"
    }
    finally {
        $btnConnect.Enabled = $true
    }
})

$btnScan.Add_Click({
    $libraryName = $txtLibrary.Text.Trim()
    $modifiedDate = $datePicker.Value.Date
    
    if ([string]::IsNullOrWhiteSpace($libraryName)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a library name", "Missing Information")
        return
    }
    
    $btnScan.Enabled = $false
    $progressBar.Visible = $true
    $progressBar.Style = 'Marquee'
    $statusLabel.Text = "Scanning..."
    Write-Log "INFO" "Scanning library: $libraryName for folders modified on $($modifiedDate.ToShortDateString())"
    $form.Refresh()
    
    try {
        # Build CAML query
        $dateStr = $modifiedDate.ToString('yyyy-MM-dd')
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
                    <Value Type='DateTime'>$dateStr</Value>
                </Eq>
            </And>
        </Where>
    </Query>
</View>
"@
        
        # Get folders
        $folders = Get-PnPListItem -List $libraryName -Query $camlQuery
        Write-Log "INFO" "Found $($folders.Count) folders modified on specified date"
        
        # Find empty folders
        $script:EmptyFolders = @()
        $progressBar.Style = 'Blocks'
        $progressBar.Maximum = $folders.Count
        $progressBar.Value = 0
        
        foreach ($folder in $folders) {
            $folderPath = $folder.FieldValues.FileRef
            $folderName = $folder.FieldValues.FileLeafRef
            
            try {
                $contents = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType All -ErrorAction SilentlyContinue
                if ($null -eq $contents -or $contents.Count -eq 0) {
                    $script:EmptyFolders += [PSCustomObject]@{
                        Select = $true
                        Id = $folder.Id
                        Name = $folderName
                        Path = $folderPath
                        Modified = $folder.FieldValues.Modified
                        CreatedBy = $folder.FieldValues.Author.LookupValue
                    }
                }
            }
            catch {
                Write-Log "WARNING" "Could not check folder: $folderName"
            }
            
            $progressBar.Value++
            if ($progressBar.Value % 10 -eq 0) {
                $form.Refresh()
            }
        }
        
        # Display results
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("Select", [bool])
        $dt.Columns.Add("Name", [string])
        $dt.Columns.Add("Path", [string])
        $dt.Columns.Add("Modified", [datetime])
        $dt.Columns.Add("CreatedBy", [string])
        $dt.Columns.Add("Id", [int])
        
        foreach ($folder in $script:EmptyFolders) {
            $row = $dt.NewRow()
            $row["Select"] = $folder.Select
            $row["Name"] = $folder.Name
            $row["Path"] = $folder.Path
            $row["Modified"] = $folder.Modified
            $row["CreatedBy"] = $folder.CreatedBy
            $row["Id"] = $folder.Id
            $dt.Rows.Add($row)
        }
        
        $dgvResults.DataSource = $dt
        $dgvResults.Columns["Id"].Visible = $false
        $dgvResults.Columns["Select"].Width = 50
        
        $lblSummary.Text = "Found $($script:EmptyFolders.Count) empty folders"
        if ($script:EmptyFolders.Count -gt 0) {
            $lblSummary.Text += " - " + $(if ($chkPreview.Checked) { "PREVIEW MODE" } else { "LIVE MODE" })
            $btnDelete.Enabled = $true
        }
        
        Write-Log "SUCCESS" "Scan complete - Found $($script:EmptyFolders.Count) empty folders"
        $statusLabel.Text = "Scan complete"
    }
    catch {
        Write-Log "ERROR" "Scan failed: $_"
        [System.Windows.Forms.MessageBox]::Show("Scan failed: $_", "Error")
        $statusLabel.Text = "Scan failed"
    }
    finally {
        $progressBar.Visible = $false
        $btnScan.Enabled = $true
    }
})

$btnDelete.Add_Click({
    $dt = $dgvResults.DataSource
    if ($null -eq $dt) { return }
    
    $selectedRows = $dt.Select("Select = true")
    
    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No folders selected", "Nothing to Delete")
        return
    }
    
    $message = "Delete $($selectedRows.Count) folders?"
    if ($chkPreview.Checked) {
        $message = "PREVIEW: Would delete $($selectedRows.Count) folders (no actual deletion)"
    }
    else {
        $message = "WARNING: Permanently delete $($selectedRows.Count) folders?"
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        "Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }
    
    $btnDelete.Enabled = $false
    $progressBar.Visible = $true
    $progressBar.Maximum = $selectedRows.Count
    $progressBar.Value = 0
    
    $successCount = 0
    $failCount = 0
    
    foreach ($row in $selectedRows) {
        $folderName = $row["Name"]
        $folderId = $row["Id"]
        
        if ($chkPreview.Checked) {
            Write-Log "INFO" "[PREVIEW] Would delete: $folderName"
            $successCount++
        }
        else {
            try {
                Remove-PnPListItem -List $txtLibrary.Text -Identity $folderId -Force
                Write-Log "SUCCESS" "Deleted: $folderName"
                $successCount++
            }
            catch {
                Write-Log "ERROR" "Failed to delete: $folderName - $_"
                $failCount++
            }
        }
        
        $progressBar.Value++
        $form.Refresh()
    }
    
    $statusLabel.Text = "Complete - Success: $successCount, Failed: $failCount"
    Write-Log "INFO" "Deletion batch complete - Success: $successCount, Failed: $failCount"
    
    if (-not $chkPreview.Checked -and $successCount -gt 0) {
        $btnScan.PerformClick()
    }
    
    $progressBar.Visible = $false
    $btnDelete.Enabled = $true
})

$btnClearLog.Add_Click({
    $txtLog.Clear()
})

# Initialize
Write-Log "INFO" "SharePoint Cleanup Tool v2.0 started"

# Show form
$form.ShowDialog()