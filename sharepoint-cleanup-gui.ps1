# SharePoint Empty Folder Cleanup GUI Tool
# Requires PnP.PowerShell module
# Version: 2.1
# Author: SharePoint Cleanup Tool

[CmdletBinding()]
param()

#region Assembly Loading
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region Module Check
# Check if PnP.PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    [System.Windows.Forms.MessageBox]::Show(
        "PnP.PowerShell module is not installed.`n`nPlease run:`nInstall-Module -Name PnP.PowerShell", 
        "Missing Module", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    exit 1
}

# Import module
Import-Module PnP.PowerShell -ErrorAction Stop
#endregion

#region Global Variables
$script:Connection = $null
$script:EmptyFolders = @()
$script:IsConnected = $false
#endregion

#region Form Creation
# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint Empty Folder Cleanup Tool"
$form.Size = New-Object System.Drawing.Size(650, 550)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Site URL
$labelSite = New-Object System.Windows.Forms.Label
$labelSite.Location = New-Object System.Drawing.Point(10, 20)
$labelSite.Size = New-Object System.Drawing.Size(100, 20)
$labelSite.Text = "Site URL:"
$form.Controls.Add($labelSite)

$textBoxSite = New-Object System.Windows.Forms.TextBox
$textBoxSite.Location = New-Object System.Drawing.Point(120, 20)
$textBoxSite.Size = New-Object System.Drawing.Size(400, 20)
$textBoxSite.Text = "https://yourtenant.sharepoint.com/sites/yoursite"
$form.Controls.Add($textBoxSite)

# Connect Button
$buttonConnect = New-Object System.Windows.Forms.Button
$buttonConnect.Location = New-Object System.Drawing.Point(530, 18)
$buttonConnect.Size = New-Object System.Drawing.Size(80, 25)
$buttonConnect.Text = "Connect"
$buttonConnect.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($buttonConnect)

# Connection Status
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(120, 45)
$labelStatus.Size = New-Object System.Drawing.Size(400, 20)
$labelStatus.Text = "Not connected"
$labelStatus.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($labelStatus)

# Library Name
$labelLibrary = New-Object System.Windows.Forms.Label
$labelLibrary.Location = New-Object System.Drawing.Point(10, 70)
$labelLibrary.Size = New-Object System.Drawing.Size(100, 20)
$labelLibrary.Text = "Library Name:"
$form.Controls.Add($labelLibrary)

$textBoxLibrary = New-Object System.Windows.Forms.TextBox
$textBoxLibrary.Location = New-Object System.Drawing.Point(120, 70)
$textBoxLibrary.Size = New-Object System.Drawing.Size(200, 20)
$textBoxLibrary.Text = "Documents"
$textBoxLibrary.Enabled = $false
$form.Controls.Add($textBoxLibrary)

# Modified Date
$labelDate = New-Object System.Windows.Forms.Label
$labelDate.Location = New-Object System.Drawing.Point(10, 100)
$labelDate.Size = New-Object System.Drawing.Size(100, 20)
$labelDate.Text = "Modified Date:"
$form.Controls.Add($labelDate)

$dateTimePicker = New-Object System.Windows.Forms.DateTimePicker
$dateTimePicker.Location = New-Object System.Drawing.Point(120, 100)
$dateTimePicker.Size = New-Object System.Drawing.Size(200, 20)
$dateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$dateTimePicker.Enabled = $false
$form.Controls.Add($dateTimePicker)

# Preview Mode Checkbox
$checkBoxPreview = New-Object System.Windows.Forms.CheckBox
$checkBoxPreview.Location = New-Object System.Drawing.Point(120, 130)
$checkBoxPreview.Size = New-Object System.Drawing.Size(200, 20)
$checkBoxPreview.Text = "Preview Mode (No Deletion)"
$checkBoxPreview.Checked = $true
$checkBoxPreview.Enabled = $false
$form.Controls.Add($checkBoxPreview)

# Scan Button
$buttonScan = New-Object System.Windows.Forms.Button
$buttonScan.Location = New-Object System.Drawing.Point(120, 160)
$buttonScan.Size = New-Object System.Drawing.Size(100, 30)
$buttonScan.Text = "Scan Folders"
$buttonScan.BackColor = [System.Drawing.Color]::LightBlue
$buttonScan.Enabled = $false
$form.Controls.Add($buttonScan)

# Delete Button
$buttonDelete = New-Object System.Windows.Forms.Button
$buttonDelete.Location = New-Object System.Drawing.Point(230, 160)
$buttonDelete.Size = New-Object System.Drawing.Size(100, 30)
$buttonDelete.Text = "Delete Selected"
$buttonDelete.BackColor = [System.Drawing.Color]::Salmon
$buttonDelete.Enabled = $false
$form.Controls.Add($buttonDelete)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 200)
$progressBar.Size = New-Object System.Drawing.Size(600, 20)
$progressBar.Style = 'Marquee'
$progressBar.MarqueeAnimationSpeed = 30
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Progress Label
$labelProgress = New-Object System.Windows.Forms.Label
$labelProgress.Location = New-Object System.Drawing.Point(10, 225)
$labelProgress.Size = New-Object System.Drawing.Size(600, 20)
$labelProgress.Text = ""
$form.Controls.Add($labelProgress)

# Results DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 250)
$dataGridView.Size = New-Object System.Drawing.Size(600, 200)
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.SelectionMode = 'FullRowSelect'
$dataGridView.MultiSelect = $true
$dataGridView.AutoSizeColumnsMode = 'Fill'
$form.Controls.Add($dataGridView)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 460)
$statusLabel.Size = New-Object System.Drawing.Size(600, 40)
$statusLabel.Text = "Ready to connect to SharePoint"
$form.Controls.Add($statusLabel)

# Exit Button
$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Location = New-Object System.Drawing.Point(510, 505)
$buttonExit.Size = New-Object System.Drawing.Size(100, 30)
$buttonExit.Text = "Exit"
$form.Controls.Add($buttonExit)
#endregion

#region Event Handlers

# Connect Button Click
$buttonConnect.Add_Click({
    $siteUrl = $textBoxSite.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid SharePoint site URL.", 
            "Invalid Input", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    try {
        $labelStatus.Text = "Connecting... Please check your browser for authentication"
        $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        $form.Refresh()
        
        # Try to connect with device login
        $script:Connection = Connect-PnPOnline -Url $siteUrl -DeviceLogin -ReturnConnection -ErrorAction Stop
        
        $script:IsConnected = $true
        $labelStatus.Text = "Connected successfully"
        $labelStatus.ForeColor = [System.Drawing.Color]::Green
        
        # Enable controls
        $textBoxLibrary.Enabled = $true
        $dateTimePicker.Enabled = $true
        $checkBoxPreview.Enabled = $true
        $buttonScan.Enabled = $true
        
        $statusLabel.Text = "Connected to SharePoint. Ready to scan for empty folders."
    }
    catch {
        $labelStatus.Text = "Connection failed"
        $labelStatus.ForeColor = [System.Drawing.Color]::Red
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to SharePoint:`n`n$($_.Exception.Message)", 
            "Connection Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# Scan Button Click
$buttonScan.Add_Click({
    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please connect to SharePoint first.", 
            "Not Connected", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $libraryName = $textBoxLibrary.Text.Trim()
    $modifiedDate = $dateTimePicker.Value.Date
    
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
    $dataGridView.DataSource = $null
    $script:EmptyFolders = @()
    
    # Show progress
    $progressBar.Visible = $true
    $labelProgress.Text = "Scanning folders..."
    $buttonScan.Enabled = $false
    $buttonDelete.Enabled = $false
    $form.Refresh()
    
    try {
        # Get folders
        $labelProgress.Text = "Getting folders from library..."
        $form.Refresh()
        
        $folders = Get-PnPListItem -List $libraryName -Connection $script:Connection -Query "
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
            </View>" -ErrorAction Stop
        
        $labelProgress.Text = "Found $($folders.Count) folders. Checking for empty folders..."
        $form.Refresh()
        
        $emptyFolders = @()
        $processedCount = 0
        
        foreach ($folder in $folders) {
            $processedCount++
            $labelProgress.Text = "Checking folder $processedCount of $($folders.Count)"
            $form.Refresh()
            
            try {
                $folderContents = Get-PnPFolderItem -FolderSiteRelativeUrl $folder.FieldValues.FileRef -ItemType All -Connection $script:Connection -ErrorAction SilentlyContinue
                
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
        $labelProgress.Text = "Scan complete. Found $($emptyFolders.Count) empty folders."
        
        if ($emptyFolders.Count -gt 0) {
            # Create DataTable for better display
            $dataTable = New-Object System.Data.DataTable
            [void]$dataTable.Columns.Add("Selected", [System.Boolean])
            [void]$dataTable.Columns.Add("Name", [System.String])
            [void]$dataTable.Columns.Add("Modified", [System.DateTime])
            [void]$dataTable.Columns.Add("Path", [System.String])
            
            foreach ($folder in $emptyFolders) {
                [void]$dataTable.Rows.Add($folder.Selected, $folder.Name, $folder.Modified, $folder.Path)
            }
            
            $dataGridView.DataSource = $dataTable
            $dataGridView.Columns["Selected"].Width = 60
            $dataGridView.Columns["Name"].Width = 200
            $dataGridView.Columns["Modified"].Width = 100
            
            $buttonDelete.Enabled = -not $checkBoxPreview.Checked
            $statusLabel.Text = "Found $($emptyFolders.Count) empty folders. Select folders and click 'Delete Selected' to remove them."
        }
        else {
            $statusLabel.Text = "No empty folders found matching the criteria."
        }
    }
    catch {
        $progressBar.Visible = $false
        $labelProgress.Text = "Scan failed"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to scan folders:`n`n$($_.Exception.Message)", 
            "Scan Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        $buttonScan.Enabled = $true
    }
})

# Delete Button Click
$buttonDelete.Add_Click({
    if ($script:EmptyFolders.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No folders to delete.", 
            "No Selection", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    # Get selected folders
    $selectedFolders = @()
    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
        if ($dataGridView.Rows[$i].Cells["Selected"].Value -eq $true) {
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
        $labelProgress.Text = "Deleting folders..."
        $buttonDelete.Enabled = $false
        $buttonScan.Enabled = $false
        $form.Refresh()
        
        $deletedCount = 0
        $failedCount = 0
        
        foreach ($folder in $selectedFolders) {
            try {
                $labelProgress.Text = "Deleting: $($folder.Name)"
                $form.Refresh()
                
                Remove-PnPListItem -List $textBoxLibrary.Text -Identity $folder.Id -Connection $script:Connection -Force
                $deletedCount++
            }
            catch {
                $failedCount++
            }
        }
        
        $progressBar.Visible = $false
        $buttonDelete.Enabled = $true
        $buttonScan.Enabled = $true
        
        $statusLabel.Text = "Deletion complete. Deleted: $deletedCount, Failed: $failedCount"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Deletion complete!`n`nDeleted: $deletedCount`nFailed: $failedCount", 
            "Deletion Complete", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        # Refresh the scan
        $buttonScan.PerformClick()
    }
})

# Preview Mode Changed
$checkBoxPreview.Add_CheckedChanged({
    if ($checkBoxPreview.Checked) {
        $buttonDelete.Enabled = $false
        $statusLabel.Text = "Preview mode enabled - no folders will be deleted"
    }
    else {
        if ($script:EmptyFolders.Count -gt 0) {
            $buttonDelete.Enabled = $true
            $statusLabel.Text = "Preview mode disabled - folders can be deleted"
        }
    }
})

# Exit Button Click
$buttonExit.Add_Click({
    if ($script:IsConnected) {
        Disconnect-PnPOnline
    }
    $form.Close()
})

# Form Closing
$form.Add_FormClosing({
    if ($script:IsConnected) {
        Disconnect-PnPOnline
    }
})
#endregion

# Show form
[void]$form.ShowDialog()