# SharePoint Empty Folder Cleanup GUI - Final Version
# Handles authentication directly in the same session

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
$form.Size = New-Object System.Drawing.Size(500, 450)
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

# Preview checkbox
$chkPreview = New-Object System.Windows.Forms.CheckBox
$chkPreview.Location = New-Object System.Drawing.Point(110, 200)
$chkPreview.Size = New-Object System.Drawing.Size(200, 20)
$chkPreview.Text = "Preview Mode (No Deletion)"
$chkPreview.Checked = $true
$chkPreview.Enabled = $false
$form.Controls.Add($chkPreview)

# Scan button
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Location = New-Object System.Drawing.Point(110, 230)
$btnScan.Size = New-Object System.Drawing.Size(100, 30)
$btnScan.Text = "Scan"
$btnScan.BackColor = [System.Drawing.Color]::LightBlue
$btnScan.Enabled = $false
$form.Controls.Add($btnScan)

# Delete button
$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Location = New-Object System.Drawing.Point(220, 230)
$btnDelete.Size = New-Object System.Drawing.Size(100, 30)
$btnDelete.Text = "Delete Empty"
$btnDelete.BackColor = [System.Drawing.Color]::Salmon
$btnDelete.Enabled = $false
$btnDelete.Visible = $false
$form.Controls.Add($btnDelete)

# Results text box
$txtResults = New-Object System.Windows.Forms.TextBox
$txtResults.Location = New-Object System.Drawing.Point(20, 270)
$txtResults.Size = New-Object System.Drawing.Size(440, 120)
$txtResults.Multiline = $true
$txtResults.ScrollBars = "Vertical"
$txtResults.ReadOnly = $true
$txtResults.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtResults)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 400)
$progressBar.Size = New-Object System.Drawing.Size(440, 20)
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Store empty folders found
$script:EmptyFolders = @()

# Connect button click - direct connection in same session
$btnConnect.Add_Click({
    $siteUrl = $txtSite.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a site URL", "Error")
        return
    }
    
    $btnConnect.Enabled = $false
    $lblStatus.Text = "Connecting... Browser will open for authentication"
    $lblStatus.ForeColor = [System.Drawing.Color]::Orange
    $txtResults.Clear()
    $txtResults.AppendText("Connecting to SharePoint...`r`n")
    $form.Refresh()
    
    try {
        # Show message about browser
        $txtResults.AppendText("Opening browser for authentication...`r`n")
        $form.Refresh()
        
        # Try different authentication methods
        $connected = $false
        $methods = @("Interactive", "UseWebLogin", "LaunchBrowser")
        
        foreach ($method in $methods) {
            if (-not $connected) {
                try {
                    $txtResults.AppendText("Trying $method authentication...`r`n")
                    $form.Refresh()
                    
                    switch ($method) {
                        "Interactive" { Connect-PnPOnline -Url $siteUrl -Interactive }
                        "UseWebLogin" { Connect-PnPOnline -Url $siteUrl -UseWebLogin }
                        "LaunchBrowser" { Connect-PnPOnline -Url $siteUrl -LaunchBrowser }
                    }
                    
                    # Test connection
                    $web = Get-PnPWeb
                    $connected = $true
                    $txtResults.AppendText("SUCCESS: Connected using $method!`r`n")
                    $txtResults.AppendText("Site: $($web.Title)`r`n")
                    break
                }
                catch {
                    $txtResults.AppendText("$method failed: $($_.Exception.Message.Split("`n")[0])`r`n")
                }
            }
        }
        
        if ($connected) {
            $script:IsConnected = $true
            $lblStatus.Text = "Connected to: $($web.Title)"
            $lblStatus.ForeColor = [System.Drawing.Color]::Green
            
            # Enable controls
            $txtLibrary.Enabled = $true
            $datePicker.Enabled = $true
            $chkPreview.Enabled = $true
            $btnScan.Enabled = $true
            $btnConnect.Text = "Connected"
            $btnConnect.BackColor = [System.Drawing.Color]::LightGray
        }
        else {
            throw "All authentication methods failed"
        }
    }
    catch {
        $lblStatus.Text = "Connection failed"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        $btnConnect.Enabled = $true
        $txtResults.AppendText("`r`nERROR: Failed to connect`r`n")
        [System.Windows.Forms.MessageBox]::Show("Failed to connect: $_", "Error")
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
    $previewMode = $chkPreview.Checked
    
    $txtResults.Clear()
    $txtResults.AppendText("Scanning for empty folders...`r`n")
    $txtResults.AppendText("Library: $libraryName`r`n")
    $txtResults.AppendText("Date: $($modifiedDate.ToShortDateString())`r`n")
    if ($previewMode) {
        $txtResults.AppendText("Mode: PREVIEW (no deletion)`r`n")
    }
    $txtResults.AppendText("---------------------------------`r`n")
    $form.Refresh()
    
    $progressBar.Visible = $true
    $btnScan.Enabled = $false
    $script:EmptyFolders = @()
    
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
        
        $txtResults.AppendText("Querying folders...`r`n")
        $folders = Get-PnPListItem -List $libraryName -Query $camlQuery
        $txtResults.AppendText("Found $($folders.Count) folders modified on this date`r`n`r`n")
        
        if ($folders.Count -eq 0) {
            $txtResults.AppendText("No folders found for the specified date.`r`n")
            $progressBar.Visible = $false
            $btnScan.Enabled = $true
            return
        }
        
        $progressBar.Maximum = $folders.Count
        $progressBar.Value = 0
        
        $emptyCount = 0
        $processedCount = 0
        
        foreach ($folder in $folders) {
            $processedCount++
            $progressBar.Value = $processedCount
            
            $folderPath = $folder.FieldValues.FileRef
            $folderName = $folder.FieldValues.FileLeafRef
            
            try {
                $folderContents = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType All -ErrorAction SilentlyContinue
                if ($folderContents.Count -eq 0) {
                    $emptyCount++
                    $script:EmptyFolders += [PSCustomObject]@{
                        Name = $folderName
                        Path = $folderPath
                        Modified = $folder.FieldValues.Modified
                        Id = $folder.Id
                    }
                    $txtResults.AppendText("EMPTY: $folderName`r`n")
                }
            }
            catch {
                $txtResults.AppendText("ERROR checking: $folderName`r`n")
            }
            
            # Update form periodically
            if ($processedCount % 10 -eq 0) {
                $form.Refresh()
            }
        }
        
        $txtResults.AppendText("`r`n---------------------------------`r`n")
        $txtResults.AppendText("Scan complete!`r`n")
        $txtResults.AppendText("Total folders checked: $($folders.Count)`r`n")
        $txtResults.AppendText("Empty folders found: $emptyCount`r`n")
        
        if ($emptyCount -gt 0 -and -not $previewMode) {
            $btnDelete.Visible = $true
            $btnDelete.Enabled = $true
            $txtResults.AppendText("`r`nClick 'Delete Empty' to remove these folders`r`n")
        }
    }
    catch {
        $txtResults.AppendText("`r`nERROR during scan: $_`r`n")
        [System.Windows.Forms.MessageBox]::Show("Error during scan: $_", "Scan Error")
    }
    finally {
        $progressBar.Visible = $false
        $btnScan.Enabled = $true
    }
})

# Delete button click
$btnDelete.Add_Click({
    if ($script:EmptyFolders.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No empty folders to delete", "Nothing to Delete")
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to delete $($script:EmptyFolders.Count) empty folders?`n`nThis action cannot be undone!", 
        "Confirm Deletion", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $txtResults.AppendText("`r`nDeleting empty folders...`r`n")
        $progressBar.Visible = $true
        $progressBar.Maximum = $script:EmptyFolders.Count
        $progressBar.Value = 0
        $btnDelete.Enabled = $false
        
        $deletedCount = 0
        foreach ($folder in $script:EmptyFolders) {
            $progressBar.Value++
            try {
                Remove-PnPListItem -List $txtLibrary.Text -Identity $folder.Id -Force
                $txtResults.AppendText("Deleted: $($folder.Name)`r`n")
                $deletedCount++
            }
            catch {
                $txtResults.AppendText("Failed to delete: $($folder.Name) - $_`r`n")
            }
            $form.Refresh()
        }
        
        $txtResults.AppendText("`r`nDeletion complete!`r`n")
        $txtResults.AppendText("Deleted $deletedCount of $($script:EmptyFolders.Count) folders`r`n")
        $progressBar.Visible = $false
        $btnDelete.Visible = $false
        $script:EmptyFolders = @()
    }
})

# Show form
$form.ShowDialog()