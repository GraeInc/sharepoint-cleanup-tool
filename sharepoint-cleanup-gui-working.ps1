# SharePoint Empty Folder Cleanup GUI Tool
# Version: 3.0 - Full GUI with working authentication
# Requires PnP.PowerShell module

[CmdletBinding()]
param()

#region Assembly Loading
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region Module Check
try {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        [System.Windows.Forms.MessageBox]::Show(
            "PnP.PowerShell module is not installed.`n`nPlease run the installer first.", 
            "Missing Module", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        exit 1
    }
    
    # Import module but continue even if there are warnings
    Import-Module PnP.PowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Warning loading PnP module: $($_.Exception.Message)`n`nContinuing anyway...", 
        "Module Warning", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}
#endregion

#region Global Variables
$script:IsConnected = $false
$script:EmptyFolders = @()
#endregion

#region Form Creation
# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint Empty Folder Cleanup Tool v3.0"
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.Icon = [System.Drawing.SystemIcons]::Information

# Create TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(670, 500)
$form.Controls.Add($tabControl)

# Tab 1: Connection
$tabConnection = New-Object System.Windows.Forms.TabPage
$tabConnection.Text = "Connection"
$tabConnection.UseVisualStyleBackColor = $true
$tabControl.TabPages.Add($tabConnection)

# Connection controls
$lblSite = New-Object System.Windows.Forms.Label
$lblSite.Location = New-Object System.Drawing.Point(20, 30)
$lblSite.Size = New-Object System.Drawing.Size(100, 20)
$lblSite.Text = "Site URL:"
$tabConnection.Controls.Add($lblSite)

$txtSite = New-Object System.Windows.Forms.TextBox
$txtSite.Location = New-Object System.Drawing.Point(130, 30)
$txtSite.Size = New-Object System.Drawing.Size(400, 20)
$txtSite.Text = "https://yourtenant.sharepoint.com/sites/yoursite"
$tabConnection.Controls.Add($txtSite)

$lblAuthMethod = New-Object System.Windows.Forms.Label
$lblAuthMethod.Location = New-Object System.Drawing.Point(20, 70)
$lblAuthMethod.Size = New-Object System.Drawing.Size(100, 20)
$lblAuthMethod.Text = "Auth Method:"
$tabConnection.Controls.Add($lblAuthMethod)

$cmbAuthMethod = New-Object System.Windows.Forms.ComboBox
$cmbAuthMethod.Location = New-Object System.Drawing.Point(130, 70)
$cmbAuthMethod.Size = New-Object System.Drawing.Size(200, 20)
$cmbAuthMethod.DropDownStyle = "DropDownList"
$cmbAuthMethod.Items.AddRange(@("Web Login (Browser)", "Credentials"))
$cmbAuthMethod.SelectedIndex = 0
$tabConnection.Controls.Add($cmbAuthMethod)

$lblUsername = New-Object System.Windows.Forms.Label
$lblUsername.Location = New-Object System.Drawing.Point(20, 110)
$lblUsername.Size = New-Object System.Drawing.Size(100, 20)
$lblUsername.Text = "Username:"
$lblUsername.Visible = $false
$tabConnection.Controls.Add($lblUsername)

$txtUsername = New-Object System.Windows.Forms.TextBox
$txtUsername.Location = New-Object System.Drawing.Point(130, 110)
$txtUsername.Size = New-Object System.Drawing.Size(250, 20)
$txtUsername.Visible = $false
$tabConnection.Controls.Add($txtUsername)

$lblPassword = New-Object System.Windows.Forms.Label
$lblPassword.Location = New-Object System.Drawing.Point(20, 140)
$lblPassword.Size = New-Object System.Drawing.Size(100, 20)
$lblPassword.Text = "Password:"
$lblPassword.Visible = $false
$tabConnection.Controls.Add($lblPassword)

$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Location = New-Object System.Drawing.Point(130, 140)
$txtPassword.Size = New-Object System.Drawing.Size(250, 20)
$txtPassword.PasswordChar = "*"
$txtPassword.Visible = $false
$tabConnection.Controls.Add($txtPassword)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(130, 180)
$btnConnect.Size = New-Object System.Drawing.Size(100, 30)
$btnConnect.Text = "Connect"
$btnConnect.BackColor = [System.Drawing.Color]::LightGreen
$tabConnection.Controls.Add($btnConnect)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Location = New-Object System.Drawing.Point(240, 180)
$btnDisconnect.Size = New-Object System.Drawing.Size(100, 30)
$btnDisconnect.Text = "Disconnect"
$btnDisconnect.BackColor = [System.Drawing.Color]::LightCoral
$btnDisconnect.Enabled = $false
$tabConnection.Controls.Add($btnDisconnect)

$lblConnectionStatus = New-Object System.Windows.Forms.Label
$lblConnectionStatus.Location = New-Object System.Drawing.Point(20, 230)
$lblConnectionStatus.Size = New-Object System.Drawing.Size(600, 40)
$lblConnectionStatus.Text = "Status: Not connected"
$lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
$lblConnectionStatus.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$tabConnection.Controls.Add($lblConnectionStatus)

# Tab 2: Scan Settings
$tabScan = New-Object System.Windows.Forms.TabPage
$tabScan.Text = "Scan Settings"
$tabScan.UseVisualStyleBackColor = $true
$tabControl.TabPages.Add($tabScan)

$lblLibrary = New-Object System.Windows.Forms.Label
$lblLibrary.Location = New-Object System.Drawing.Point(20, 30)
$lblLibrary.Size = New-Object System.Drawing.Size(100, 20)
$lblLibrary.Text = "Library Name:"
$tabScan.Controls.Add($lblLibrary)

$txtLibrary = New-Object System.Windows.Forms.TextBox
$txtLibrary.Location = New-Object System.Drawing.Point(130, 30)
$txtLibrary.Size = New-Object System.Drawing.Size(200, 20)
$txtLibrary.Text = "Documents"
$txtLibrary.Enabled = $false
$tabScan.Controls.Add($txtLibrary)

$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Location = New-Object System.Drawing.Point(20, 70)
$lblDate.Size = New-Object System.Drawing.Size(100, 20)
$lblDate.Text = "Modified Date:"
$tabScan.Controls.Add($lblDate)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = New-Object System.Drawing.Point(130, 70)
$datePicker.Size = New-Object System.Drawing.Size(200, 20)
$datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$datePicker.Enabled = $false
$tabScan.Controls.Add($datePicker)

$chkPreview = New-Object System.Windows.Forms.CheckBox
$chkPreview.Location = New-Object System.Drawing.Point(130, 110)
$chkPreview.Size = New-Object System.Drawing.Size(200, 20)
$chkPreview.Text = "Preview Mode (No Deletion)"
$chkPreview.Checked = $true
$chkPreview.Enabled = $false
$tabScan.Controls.Add($chkPreview)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Location = New-Object System.Drawing.Point(130, 150)
$btnScan.Size = New-Object System.Drawing.Size(100, 30)
$btnScan.Text = "Scan Folders"
$btnScan.BackColor = [System.Drawing.Color]::LightBlue
$btnScan.Enabled = $false
$tabScan.Controls.Add($btnScan)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 200)
$progressBar.Size = New-Object System.Drawing.Size(620, 20)
$progressBar.Visible = $false
$tabScan.Controls.Add($progressBar)

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(20, 230)
$lblProgress.Size = New-Object System.Drawing.Size(620, 20)
$lblProgress.Text = ""
$tabScan.Controls.Add($lblProgress)

# Tab 3: Results
$tabResults = New-Object System.Windows.Forms.TabPage
$tabResults.Text = "Results"
$tabResults.UseVisualStyleBackColor = $true
$tabControl.TabPages.Add($tabResults)

$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(20, 20)
$dataGrid.Size = New-Object System.Drawing.Size(620, 350)
$dataGrid.AllowUserToAddRows = $false
$dataGrid.AllowUserToDeleteRows = $false
$dataGrid.SelectionMode = 'FullRowSelect'
$dataGrid.MultiSelect = $true
$dataGrid.AutoSizeColumnsMode = 'Fill'
$tabResults.Controls.Add($dataGrid)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Location = New-Object System.Drawing.Point(20, 380)
$btnDelete.Size = New-Object System.Drawing.Size(120, 30)
$btnDelete.Text = "Delete Selected"
$btnDelete.BackColor = [System.Drawing.Color]::Salmon
$btnDelete.Enabled = $false
$tabResults.Controls.Add($btnDelete)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Location = New-Object System.Drawing.Point(150, 380)
$btnExport.Size = New-Object System.Drawing.Size(120, 30)
$btnExport.Text = "Export to CSV"
$btnExport.Enabled = $false
$tabResults.Controls.Add($btnExport)

$lblResultsSummary = New-Object System.Windows.Forms.Label
$lblResultsSummary.Location = New-Object System.Drawing.Point(20, 420)
$lblResultsSummary.Size = New-Object System.Drawing.Size(620, 40)
$lblResultsSummary.Text = "No scan performed yet"
$tabResults.Controls.Add($lblResultsSummary)

# Status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusBar.Items.Add($statusLabel)
$form.Controls.Add($statusBar)

# Exit button
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = New-Object System.Drawing.Point(580, 520)
$btnExit.Size = New-Object System.Drawing.Size(100, 30)
$btnExit.Text = "Exit"
$form.Controls.Add($btnExit)
#endregion

#region Event Handlers

# Auth method changed
$cmbAuthMethod.Add_SelectedIndexChanged({
    if ($cmbAuthMethod.SelectedItem -eq "Credentials") {
        $lblUsername.Visible = $true
        $txtUsername.Visible = $true
        $lblPassword.Visible = $true
        $txtPassword.Visible = $true
    }
    else {
        $lblUsername.Visible = $false
        $txtUsername.Visible = $false
        $lblPassword.Visible = $false
        $txtPassword.Visible = $false
    }
})

# Connect button
$btnConnect.Add_Click({
    $siteUrl = $txtSite.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid SharePoint site URL.", 
            "Invalid Input", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # Disable button during connection
    $btnConnect.Enabled = $false
    $statusLabel.Text = "Connecting..."
    $lblConnectionStatus.Text = "Status: Connecting..."
    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Orange
    $form.Refresh()
    
    try {
        if ($cmbAuthMethod.SelectedItem -eq "Credentials") {
            # Credential authentication
            if ([string]::IsNullOrWhiteSpace($txtUsername.Text) -or [string]::IsNullOrWhiteSpace($txtPassword.Text)) {
                throw "Username and password are required"
            }
            
            $securePassword = ConvertTo-SecureString $txtPassword.Text -AsPlainText -Force
            $credentials = New-Object System.Management.Automation.PSCredential($txtUsername.Text, $securePassword)
            
            try {
                Connect-PnPOnline -Url $siteUrl -Credentials $credentials
            }
            catch {
                # If credentials fail, fallback to web login
                [System.Windows.Forms.MessageBox]::Show(
                    "Credential authentication failed. Trying web login...", 
                    "Authentication", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                Connect-PnPOnline -Url $siteUrl -UseWebLogin
            }
        }
        else {
            # Web Login (Browser) - Using Interactive authentication for GUI
            [System.Windows.Forms.MessageBox]::Show(
                "A browser window will open for authentication.`nPlease sign in and complete the authentication.", 
                "Web Login", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            # Try Interactive first (works better in GUI context)
            try {
                Connect-PnPOnline -Url $siteUrl -Interactive
            }
            catch {
                # Fallback to UseWebLogin if Interactive fails
                try {
                    Connect-PnPOnline -Url $siteUrl -UseWebLogin
                }
                catch {
                    # Last resort - try LaunchBrowser
                    Connect-PnPOnline -Url $siteUrl -LaunchBrowser
                }
            }
        }
        
        # Test connection
        $web = Get-PnPWeb
        
        $script:IsConnected = $true
        $lblConnectionStatus.Text = "Status: Connected to $($web.Title)"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Green
        $statusLabel.Text = "Connected"
        
        # Enable controls
        $btnConnect.Enabled = $false
        $btnDisconnect.Enabled = $true
        $txtLibrary.Enabled = $true
        $datePicker.Enabled = $true
        $chkPreview.Enabled = $true
        $btnScan.Enabled = $true
        
        # Switch to scan tab
        $tabControl.SelectedTab = $tabScan
    }
    catch {
        $lblConnectionStatus.Text = "Status: Connection failed - $($_.Exception.Message)"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
        $statusLabel.Text = "Connection failed"
        
        # Re-enable connect button on failure
        $btnConnect.Enabled = $true
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to SharePoint:`n`n$($_.Exception.Message)", 
            "Connection Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# Disconnect button
$btnDisconnect.Add_Click({
    try {
        Disconnect-PnPOnline
        $script:IsConnected = $false
        
        $lblConnectionStatus.Text = "Status: Not connected"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
        $statusLabel.Text = "Disconnected"
        
        # Disable controls
        $btnConnect.Enabled = $true
        $btnDisconnect.Enabled = $false
        $txtLibrary.Enabled = $false
        $datePicker.Enabled = $false
        $chkPreview.Enabled = $false
        $btnScan.Enabled = $false
        $btnDelete.Enabled = $false
        $btnExport.Enabled = $false
        
        # Clear results
        $dataGrid.DataSource = $null
        $script:EmptyFolders = @()
    }
    catch {
        # Already disconnected
    }
})

# Scan button
$btnScan.Add_Click({
    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please connect to SharePoint first.", 
            "Not Connected", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $libraryName = $txtLibrary.Text.Trim()
    $modifiedDate = $datePicker.Value.Date
    
    if ([string]::IsNullOrWhiteSpace($libraryName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a library name.", 
            "Invalid Input", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # Clear previous results
    $dataGrid.DataSource = $null
    $script:EmptyFolders = @()
    $btnDelete.Enabled = $false
    $btnExport.Enabled = $false
    
    # Show progress
    $progressBar.Visible = $true
    $progressBar.Style = 'Marquee'
    $lblProgress.Text = "Getting folders from library..."
    $statusLabel.Text = "Scanning..."
    $btnScan.Enabled = $false
    $form.Refresh()
    
    try {
        # Get folders
        $folders = Get-PnPListItem -List $libraryName -Query "
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
            </View>"
        
        $lblProgress.Text = "Found $($folders.Count) folders. Checking for empty folders..."
        $progressBar.Style = 'Blocks'
        $progressBar.Maximum = $folders.Count
        $progressBar.Value = 0
        $form.Refresh()
        
        $emptyFolders = @()
        $processedCount = 0
        
        foreach ($folder in $folders) {
            $processedCount++
            $progressBar.Value = $processedCount
            $lblProgress.Text = "Checking folder $processedCount of $($folders.Count): $($folder.FieldValues.FileLeafRef)"
            $form.Refresh()
            
            try {
                $folderContents = Get-PnPFolderItem -FolderSiteRelativeUrl $folder.FieldValues.FileRef -ItemType All -ErrorAction SilentlyContinue
                
                if ($null -eq $folderContents -or $folderContents.Count -eq 0) {
                    $emptyFolders += [PSCustomObject]@{
                        Selected = $true
                        Name = $folder.FieldValues.FileLeafRef
                        Path = $folder.FieldValues.FileRef
                        Modified = $folder.FieldValues.Modified
                        Id = $folder.Id
                    }
                }
            }
            catch {
                # Skip folders we can't access
            }
        }
        
        $script:EmptyFolders = $emptyFolders
        
        # Update UI
        $progressBar.Visible = $false
        $lblProgress.Text = ""
        $statusLabel.Text = "Scan complete"
        
        # Create DataTable
        $dataTable = New-Object System.Data.DataTable
        [void]$dataTable.Columns.Add("Selected", [System.Boolean])
        [void]$dataTable.Columns.Add("Name", [System.String])
        [void]$dataTable.Columns.Add("Modified", [System.DateTime])
        [void]$dataTable.Columns.Add("Path", [System.String])
        
        foreach ($folder in $emptyFolders) {
            [void]$dataTable.Rows.Add($folder.Selected, $folder.Name, $folder.Modified, $folder.Path)
        }
        
        $dataGrid.DataSource = $dataTable
        
        # Update summary
        $lblResultsSummary.Text = "Total folders checked: $($folders.Count)`nEmpty folders found: $($emptyFolders.Count)"
        
        if ($emptyFolders.Count -gt 0) {
            $btnExport.Enabled = $true
            if (-not $chkPreview.Checked) {
                $btnDelete.Enabled = $true
            }
        }
        
        # Switch to results tab
        $tabControl.SelectedTab = $tabResults
    }
    catch {
        $progressBar.Visible = $false
        $lblProgress.Text = ""
        $statusLabel.Text = "Scan failed"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to scan folders:`n`n$($_.Exception.Message)", 
            "Scan Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        $btnScan.Enabled = $true
    }
})

# Delete button
$btnDelete.Add_Click({
    if ($script:EmptyFolders.Count -eq 0) {
        return
    }
    
    # Get selected folders
    $selectedFolders = @()
    for ($i = 0; $i -lt $dataGrid.Rows.Count; $i++) {
        if ($dataGrid.Rows[$i].Cells["Selected"].Value -eq $true) {
            $selectedFolders += $script:EmptyFolders[$i]
        }
    }
    
    if ($selectedFolders.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one folder to delete.", 
            "No Selection", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to delete $($selectedFolders.Count) empty folder(s)?`n`nThis action cannot be undone!", 
        "Confirm Deletion", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $progressBar.Visible = $true
        $progressBar.Maximum = $selectedFolders.Count
        $progressBar.Value = 0
        $statusLabel.Text = "Deleting folders..."
        $btnDelete.Enabled = $false
        $form.Refresh()
        
        $deletedCount = 0
        $failedCount = 0
        
        foreach ($folder in $selectedFolders) {
            try {
                $lblProgress.Text = "Deleting: $($folder.Name)"
                $form.Refresh()
                
                Remove-PnPListItem -List $txtLibrary.Text -Identity $folder.Id -Force
                $deletedCount++
            }
            catch {
                $failedCount++
            }
            $progressBar.Value++
        }
        
        $progressBar.Visible = $false
        $lblProgress.Text = ""
        $statusLabel.Text = "Deletion complete"
        $btnDelete.Enabled = $true
        
        [System.Windows.Forms.MessageBox]::Show(
            "Deletion complete!`n`nDeleted: $deletedCount`nFailed: $failedCount", 
            "Deletion Complete", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        # Refresh scan
        $btnScan.PerformClick()
    }
})

# Export button
$btnExport.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveDialog.FileName = "empty_folders_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:EmptyFolders | Select-Object Name, Path, Modified | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
        
        [System.Windows.Forms.MessageBox]::Show(
            "Results exported to:`n$($saveDialog.FileName)", 
            "Export Complete", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# Preview mode changed
$chkPreview.Add_CheckedChanged({
    if ($chkPreview.Checked) {
        $btnDelete.Enabled = $false
    }
    else {
        if ($script:EmptyFolders.Count -gt 0) {
            $btnDelete.Enabled = $true
        }
    }
})

# Exit button
$btnExit.Add_Click({
    if ($script:IsConnected) {
        try {
            Disconnect-PnPOnline
        }
        catch {}
    }
    $form.Close()
})

# Form closing
$form.Add_FormClosing({
    if ($script:IsConnected) {
        try {
            Disconnect-PnPOnline
        }
        catch {}
    }
})
#endregion

# Show form
[void]$form.ShowDialog()