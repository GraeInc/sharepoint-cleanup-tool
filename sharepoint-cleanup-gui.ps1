# SharePoint Empty Folder Cleanup GUI Tool
# Requires PnP.PowerShell module
# Version: 2.0
# Author: SharePoint Cleanup Tool

[CmdletBinding()]
param()

#region Assembly Loading
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
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
#endregion

#region Form Creation
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

# Cancel Button
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(230, 140)
$buttonCancel.Size = New-Object System.Drawing.Size(100, 30)
$buttonCancel.Text = "Cancel"
$buttonCancel.BackColor = [System.Drawing.Color]::LightYellow
$buttonCancel.Enabled = $false
$form.Controls.Add($buttonCancel)

# Delete Button
$buttonDelete = New-Object System.Windows.Forms.Button
$buttonDelete.Location = New-Object System.Drawing.Point(340, 140)
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
#endregion

#region Global Variables
# Global variables for folder data and cancellation
$script:EmptyFolders = @()
$script:CancellationToken = $false
$script:ScanJob = $null
#endregion

#region Helper Functions
# Function to validate URL format
function Test-UrlFormat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )
    
    try {
        $uri = [System.Uri]::new($Url)
        return ($uri.Scheme -eq "https" -and $uri.Host -like "*.sharepoint.com")
    }
    catch {
        Write-Verbose "Invalid URL format: $Url"
        return $false
    }
}

# Function to run scan with timeout
function Start-ScanWithTimeout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$LibraryName,
        
        [Parameter(Mandatory = $true)]
        [datetime]$ModifiedDate
    )
    
    $job = Start-Job -ScriptBlock {
        param($SiteUrl, $LibraryName, $ModifiedDate)
        
        try {
            # Import module in job context
            Import-Module PnP.PowerShell -ErrorAction Stop
            
            # Connect with timeout
            $connection = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection -ErrorAction Stop
            
            # Get folders with timeout
            $folders = Get-PnPListItem -List $LibraryName -Connection $connection -Query "
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
                                    <Value Type='DateTime'>$($ModifiedDate.ToString('yyyy-MM-dd'))</Value>
                                </Eq>
                            </And>
                        </Where>
                    </Query>
                </View>" -ErrorAction Stop
            
            $emptyFolders = @()
            
            foreach ($folder in $folders) {
                try {
                    $folderContents = Get-PnPFolderItem -FolderSiteRelativeUrl $folder.FieldValues.FileRef -ItemType All -Connection $connection -ErrorAction SilentlyContinue
                    
                    if ($folderContents.Count -eq 0) {
                        $emptyFolderInfo = [PSCustomObject]@{
                            Name = $folder.FieldValues.FileLeafRef
                            Path = $folder.FieldValues.FileRef
                            Modified = $folder.FieldValues.Modified
                            Id = $folder.Id
                        }
                        $emptyFolders += $emptyFolderInfo
                    }
                }
                catch {
                    # Skip folders we can't access
                    Write-Verbose "Cannot access folder: $($folder.FieldValues.FileLeafRef)"
                }
            }
            
            return @{
                Success = $true
                EmptyFolders = $emptyFolders
                TotalFolders = $folders.Count
                Error = $null
            }
        }
        catch {
            return @{
                Success = $false
                EmptyFolders = @()
                TotalFolders = 0
                Error = $_.Exception.Message
            }
        }
    } -ArgumentList $SiteUrl, $LibraryName, $ModifiedDate
    
    return $job
}
#endregion

#region Event Handlers
# Scan Button Click Event 
$buttonScan.Add_Click({
    $siteUrl = $textBoxSite.Text.Trim()
    $libraryName = $textBoxLibrary.Text.Trim()
    $modifiedDate = $dateTimePicker.Value.Date
    
    if ([string]::IsNullOrEmpty($siteUrl) -or [string]::IsNullOrEmpty($libraryName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please fill in Site URL and Library Name", 
            "Input Required", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # Validate URL format
    if (-not (Test-UrlFormat -Url $siteUrl)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid SharePoint Online URL (https://*.sharepoint.com/sites/*)", 
            "Invalid URL", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $buttonScan.Enabled = $false
    $buttonDelete.Enabled = $false
    $buttonCancel.Enabled = $true
    $listBoxResults.Items.Clear()
    $progressBar.Visible = $true
    $labelStatus.Text = "Connecting to SharePoint..."
    $labelStatus.ForeColor = [System.Drawing.Color]::Orange
    $script:CancellationToken = $false
    
    try {
        # Start scan job with timeout
        $script:ScanJob = Start-ScanWithTimeout -SiteUrl $siteUrl -LibraryName $libraryName -ModifiedDate $modifiedDate
        
        # Monitor job with timeout
        $timeout = 120  # 2 minutes timeout
        $startTime = Get-Date
        
        while ($script:ScanJob.State -eq "Running" -and -not $script:CancellationToken) {
            if (((Get-Date) - $startTime).TotalSeconds -gt $timeout) {
                Stop-Job $script:ScanJob -ErrorAction SilentlyContinue
                Remove-Job $script:ScanJob -ErrorAction SilentlyContinue
                throw "Operation timed out after $timeout seconds. Please check your connection and try again."
            }
            
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 100
        }
        
        if ($script:CancellationToken) {
            $labelStatus.Text = "Scan cancelled by user."
            $labelStatus.ForeColor = [System.Drawing.Color]::Orange
            return
        }
        
        # Get results
        $result = Receive-Job $script:ScanJob
        Remove-Job $script:ScanJob -ErrorAction SilentlyContinue
        
        if (-not $result.Success) {
            throw $result.Error
        }
        
        $script:EmptyFolders = $result.EmptyFolders
        
        if ($script:EmptyFolders.Count -gt 0) {
            foreach ($folder in $script:EmptyFolders) {
                $listBoxResults.Items.Add("$($folder.Name) - $($folder.Modified.ToString('yyyy-MM-dd HH:mm'))", $true)
            }
            $buttonDelete.Enabled = $true
        }
        
        $labelStatus.Text = "Scan complete. Found $($script:EmptyFolders.Count) empty folders out of $($result.TotalFolders) total folders."
        $labelStatus.ForeColor = [System.Drawing.Color]::Green
        
    }
    catch {
        $labelStatus.Text = "Error: $($_.Exception.Message)"
        $labelStatus.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show(
            "Error occurred during scan: $($_.Exception.Message)", 
            "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        if ($script:ScanJob -and $script:ScanJob.State -eq "Running") {
            Stop-Job $script:ScanJob -ErrorAction SilentlyContinue
            Remove-Job $script:ScanJob -ErrorAction SilentlyContinue
        }
        $buttonScan.Enabled = $true
        $buttonCancel.Enabled = $false
        $progressBar.Visible = $false
        $script:ScanJob = $null
    }
})

# Cancel Button Click Event
$buttonCancel.Add_Click({
    $script:CancellationToken = $true
    if ($script:ScanJob -and $script:ScanJob.State -eq "Running") {
        Stop-Job $script:ScanJob -ErrorAction SilentlyContinue
        Remove-Job $script:ScanJob -ErrorAction SilentlyContinue
    }
    $labelStatus.Text = "Scan cancelled by user."
    $labelStatus.ForeColor = [System.Drawing.Color]::Orange
    $buttonScan.Enabled = $true
    $buttonCancel.Enabled = $false
    $progressBar.Visible = $false
})

# Delete Button Click Event
$buttonDelete.Add_Click({
    $selectedIndices = $listBoxResults.CheckedIndices
    
    if ($selectedIndices.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select folders to delete", 
            "No Selection", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    if ($checkBoxPreview.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            "Preview Mode: $($selectedIndices.Count) folders would be deleted.`n`nUncheck 'Preview Mode' to actually delete folders.", 
            "Preview Mode", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to delete $($selectedIndices.Count) selected empty folders?`n`nThis action cannot be undone.", 
        "Confirm Deletion", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $buttonDelete.Enabled = $false
        $progressBar.Visible = $true
        $progressBar.Maximum = $selectedIndices.Count
        $deletedCount = 0
        
        for ($i = 0; $i -lt $selectedIndices.Count; $i++) {
            $index = $selectedIndices[$i]
            $folderToDelete = $script:EmptyFolders[$index]
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
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to delete folder '$($folderToDelete.Name)': $($_.Exception.Message)", 
                    "Deletion Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
        
        $labelStatus.Text = "Deletion complete. Successfully deleted $deletedCount folders."
        $labelStatus.ForeColor = [System.Drawing.Color]::Green
        $progressBar.Visible = $false
        $buttonDelete.Enabled = ($listBoxResults.Items.Count -gt 0)
    }
})
#endregion

#region Form Events
# Form closing event to clean up jobs
$form.Add_FormClosing({
    if ($script:ScanJob -and $script:ScanJob.State -eq "Running") {
        Stop-Job $script:ScanJob -ErrorAction SilentlyContinue
        Remove-Job $script:ScanJob -ErrorAction SilentlyContinue
    }
})
#endregion

#region Main Execution
# Show form
$form.ShowDialog()
#endregion