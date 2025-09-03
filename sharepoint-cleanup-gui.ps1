# SharePoint Empty Folder Cleanup GUI Tool
# Requires PnP.PowerShell module

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check if PnP.PowerShell is installed
if (!(Get-Module -ListAvailable -Name PnP.PowerShell)) {
    [System.Windows.Forms.MessageBox]::Show("PnP.PowerShell module is not installed.`n`nPlease run:`nInstall-Module -Name PnP.PowerShell", "Missing Module", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    exit
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint Empty Folder Cleanup Tool"
$form.Size = New-Object System.Drawing.Size(600, 500)
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
$textBoxSite.Size = New-Object System.Drawing.Size(450, 20)
$textBoxSite.Text = "https://yourtenant.sharepoint.com/sites/yoursite"
$form.Controls.Add($textBoxSite)

# Library Name
$labelLibrary = New-Object System.Windows.Forms.Label
$labelLibrary.Location = New-Object System.Drawing.Point(10, 50)
$labelLibrary.Size = New-Object System.Drawing.Size(100, 20)
$labelLibrary.Text = "Library Name:"
$form.Controls.Add($labelLibrary)

$textBoxLibrary = New-Object System.Windows.Forms.TextBox
$textBoxLibrary.Location = New-Object System.Drawing.Point(120, 50)
$textBoxLibrary.Size = New-Object System.Drawing.Size(200, 20)
$textBoxLibrary.Text = "Documents"
$form.Controls.Add($textBoxLibrary)

# Modified Date
$labelDate = New-Object System.Windows.Forms.Label
$labelDate.Location = New-Object System.Drawing.Point(10, 80)
$labelDate.Size = New-Object System.Drawing.Size(100, 20)
$labelDate.Text = "Modified Date:"
$form.Controls.Add($labelDate)

$dateTimePicker = New-Object System.Windows.Forms.DateTimePicker
$dateTimePicker.Location = New-Object System.Drawing.Point(120, 80)
$dateTimePicker.Size = New-Object System.Drawing.Size(200, 20)
$dateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$form.Controls.Add($dateTimePicker)

# Preview Mode Checkbox
$checkBoxPreview = New-Object System.Windows.Forms.CheckBox
$checkBoxPreview.Location = New-Object System.Drawing.Point(120, 110)
$checkBoxPreview.Size = New-Object System.Drawing.Size(200, 20)
$checkBoxPreview.Text = "Preview Mode (No Deletion)"
$checkBoxPreview.Checked = $true
$form.Controls.Add($checkBoxPreview)

# Scan Button
$buttonScan = New-Object System.Windows.Forms.Button
$buttonScan.Location = New-Object System.Drawing.Point(120, 140)
$buttonScan.Size = New-Object System.Drawing.Size(100, 30)
$buttonScan.Text = "Scan Folders"
$buttonScan.BackColor = [System.Drawing.Color]::LightBlue
$form.Controls.Add($buttonScan)

# Delete Button
$buttonDelete = New-Object System.Windows.Forms.Button
$buttonDelete.Location = New-Object System.Drawing.Point(230, 140)
$buttonDelete.Size = New-Object System.Drawing.Size(100, 30)
$buttonDelete.Text = "Delete Selected"
$buttonDelete.BackColor = [System.Drawing.Color]::LightCoral
$buttonDelete.Enabled = $false
$form.Controls.Add($buttonDelete)

# Results ListBox
$labelResults = New-Object System.Windows.Forms.Label
$labelResults.Location = New-Object System.Drawing.Point(10, 180)
$labelResults.Size = New-Object System.Drawing.Size(200, 20)
$labelResults.Text = "Empty Folders Found:"
$form.Controls.Add($labelResults)

$listBoxResults = New-Object System.Windows.Forms.CheckedListBox
$listBoxResults.Location = New-Object System.Drawing.Point(10, 200)
$listBoxResults.Size = New-Object System.Drawing.Size(560, 200)
$listBoxResults.CheckOnClick = $true
$form.Controls.Add($listBoxResults)

# Status Label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(10, 410)
$labelStatus.Size = New-Object System.Drawing.Size(560, 20)
$labelStatus.Text = "Ready to scan..."
$labelStatus.ForeColor = [System.Drawing.Color]::Blue
$form.Controls.Add($labelStatus)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 430)
$progressBar.Size = New-Object System.Drawing.Size(560, 20)
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Global variables for folder data
$script:emptyFolders = @()

# Scan Button Click Event
$buttonScan.Add_Click({
    $siteUrl = $textBoxSite.Text.Trim()
    $libraryName = $textBoxLibrary.Text.Trim()
    $modifiedDate = $dateTimePicker.Value.Date
    
    if ([string]::IsNullOrEmpty($siteUrl) -or [string]::IsNullOrEmpty($libraryName)) {
        [System.Windows.Forms.MessageBox]::Show("Please fill in Site URL and Library Name", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $buttonScan.Enabled = $false
    $buttonDelete.Enabled = $false
    $listBoxResults.Items.Clear()
    $progressBar.Visible = $true
    $labelStatus.Text = "Connecting to SharePoint..."
    $labelStatus.ForeColor = [System.Drawing.Color]::Orange
    
    try {
        # Connect to SharePoint
        Connect-PnPOnline -Url $siteUrl -Interactive
        $labelStatus.Text = "Connected. Scanning folders..."
        
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
        
        $script:emptyFolders = @()
        $progressBar.Maximum = $folders.Count
        
        for ($i = 0; $i -lt $folders.Count; $i++) {
            $folder = $folders[$i]
            $progressBar.Value = $i + 1
            $labelStatus.Text = "Checking folder $($i + 1) of $($folders.Count): $($folder.FieldValues.FileLeafRef)"
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                $folderContents = Get-PnPFolderItem -FolderSiteRelativeUrl $folder.FieldValues.FileRef -ItemType All -ErrorAction SilentlyContinue
                
                if ($folderContents.Count -eq 0) {
                    $emptyFolderInfo = [PSCustomObject]@{
                        Name = $folder.FieldValues.FileLeafRef
                        Path = $folder.FieldValues.FileRef
                        Modified = $folder.FieldValues.Modified
                        Id = $folder.Id
                    }
                    $script:emptyFolders += $emptyFolderInfo
                    $listBoxResults.Items.Add("$($emptyFolderInfo.Name) - $($emptyFolderInfo.Modified.ToString('yyyy-MM-dd HH:mm'))", $true)
                }
            }
            catch {
                # Skip folders we can't access
            }
        }
        
        $labelStatus.Text = "Scan complete. Found $($script:emptyFolders.Count) empty folders out of $($folders.Count) total folders."
        $labelStatus.ForeColor = [System.Drawing.Color]::Green
        
        if ($script:emptyFolders.Count -gt 0) {
            $buttonDelete.Enabled = $true
        }
        
    }
    catch {
        $labelStatus.Text = "Error: $($_.Exception.Message)"
        $labelStatus.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show("Error occurred during scan: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $buttonScan.Enabled = $true
        $progressBar.Visible = $false
    }
})

# Delete Button Click Event
$buttonDelete.Add_Click({
    $selectedIndices = $listBoxResults.CheckedIndices
    
    if ($selectedIndices.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select folders to delete", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ($checkBoxPreview.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Preview Mode: $($selectedIndices.Count) folders would be deleted.`n`nUncheck 'Preview Mode' to actually delete folders.", "Preview Mode", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete $($selectedIndices.Count) selected empty folders?`n`nThis action cannot be undone.", "Confirm Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $buttonDelete.Enabled = $false
        $progressBar.Visible = $true
        $progressBar.Maximum = $selectedIndices.Count
        $deletedCount = 0
        
        for ($i = 0; $i -lt $selectedIndices.Count; $i++) {
            $index = $selectedIndices[$i]
            $folderToDelete = $script:emptyFolders[$index]
            $progressBar.Value = $i + 1
            $labelStatus.Text = "Deleting folder $($i + 1) of $($selectedIndices.Count): $($folderToDelete.Name)"
            $labelStatus.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                Remove-PnPListItem -List $textBoxLibrary.Text -Identity $folderToDelete.Id -Force
                $listBoxResults.Items.RemoveAt($index - $deletedCount)
                $deletedCount++
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to delete folder '$($folderToDelete.Name)': $($_.Exception.Message)", "Deletion Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        
        $labelStatus.Text = "Deletion complete. Successfully deleted $deletedCount folders."
        $labelStatus.ForeColor = [System.Drawing.Color]::Green
        $progressBar.Visible = $false
        $buttonDelete.Enabled = ($listBoxResults.Items.Count -gt 0)
    }
})

# Show form
$form.ShowDialog()