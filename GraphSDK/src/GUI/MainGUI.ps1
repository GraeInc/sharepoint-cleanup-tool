# MainGUI.ps1
# Windows Forms GUI for SharePoint Cleanup Tool

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import core modules
$corePath = Join-Path $PSScriptRoot "..\Core"
. (Join-Path $corePath "GraphAuth.ps1")
. (Join-Path $corePath "FolderOps.ps1")
. (Join-Path $corePath "Logger.ps1")

# Script-level variables
$Script:Connection = $null
$Script:CurrentSite = $null
$Script:CurrentLibrary = $null
$Script:EmptyFolders = @()

function Show-CleanupGUI {
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "SharePoint Cleanup Tool - Graph SDK"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # Create tab control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(770, 500)
    
    # Tab 1: Connection
    $tabConnect = New-Object System.Windows.Forms.TabPage
    $tabConnect.Text = "Connection"
    $tabConnect.BackColor = [System.Drawing.Color]::White
    
    # Site URL input
    $lblSiteUrl = New-Object System.Windows.Forms.Label
    $lblSiteUrl.Text = "SharePoint Site URL:"
    $lblSiteUrl.Location = New-Object System.Drawing.Point(20, 30)
    $lblSiteUrl.Size = New-Object System.Drawing.Size(150, 20)
    $tabConnect.Controls.Add($lblSiteUrl)
    
    $txtSiteUrl = New-Object System.Windows.Forms.TextBox
    $txtSiteUrl.Location = New-Object System.Drawing.Point(20, 55)
    $txtSiteUrl.Size = New-Object System.Drawing.Size(500, 25)
    $txtSiteUrl.Text = "https://contoso.sharepoint.com/sites/TeamSite"
    $tabConnect.Controls.Add($txtSiteUrl)
    
    # Connect button
    $btnConnect = New-Object System.Windows.Forms.Button
    $btnConnect.Text = "Connect"
    $btnConnect.Location = New-Object System.Drawing.Point(530, 53)
    $btnConnect.Size = New-Object System.Drawing.Size(100, 27)
    $btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnConnect.ForeColor = [System.Drawing.Color]::White
    $btnConnect.FlatStyle = "Flat"
    $tabConnect.Controls.Add($btnConnect)
    
    # Connection status
    $lblConnStatus = New-Object System.Windows.Forms.Label
    $lblConnStatus.Text = "Not connected"
    $lblConnStatus.Location = New-Object System.Drawing.Point(20, 100)
    $lblConnStatus.Size = New-Object System.Drawing.Size(700, 60)
    $lblConnStatus.ForeColor = [System.Drawing.Color]::Gray
    $tabConnect.Controls.Add($lblConnStatus)
    
    # Instructions
    $lblInstructions = New-Object System.Windows.Forms.Label
    $lblInstructions.Text = @"
Instructions:
1. Enter your SharePoint site URL above
2. Click Connect to authenticate with Microsoft Graph
3. A browser window will open for authentication (supports MFA)
4. After connecting, go to the 'Scan & Clean' tab
5. Select a document library and date to find empty folders

Note: This tool uses Microsoft Graph SDK and does not require app registration.
"@
    $lblInstructions.Location = New-Object System.Drawing.Point(20, 180)
    $lblInstructions.Size = New-Object System.Drawing.Size(700, 200)
    $tabConnect.Controls.Add($lblInstructions)
    
    $tabControl.TabPages.Add($tabConnect)
    
    # Tab 2: Scan & Clean
    $tabScan = New-Object System.Windows.Forms.TabPage
    $tabScan.Text = "Scan & Clean"
    $tabScan.BackColor = [System.Drawing.Color]::White
    
    # Library selection
    $lblLibrary = New-Object System.Windows.Forms.Label
    $lblLibrary.Text = "Document Library:"
    $lblLibrary.Location = New-Object System.Drawing.Point(20, 20)
    $lblLibrary.Size = New-Object System.Drawing.Size(150, 20)
    $tabScan.Controls.Add($lblLibrary)
    
    $cmbLibrary = New-Object System.Windows.Forms.ComboBox
    $cmbLibrary.Location = New-Object System.Drawing.Point(20, 45)
    $cmbLibrary.Size = New-Object System.Drawing.Size(300, 25)
    $cmbLibrary.DropDownStyle = "DropDownList"
    $cmbLibrary.Enabled = $false
    $tabScan.Controls.Add($cmbLibrary)
    
    # Date filter
    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Text = "Modified Date:"
    $lblDate.Location = New-Object System.Drawing.Point(340, 20)
    $lblDate.Size = New-Object System.Drawing.Size(150, 20)
    $tabScan.Controls.Add($lblDate)
    
    $datePicker = New-Object System.Windows.Forms.DateTimePicker
    $datePicker.Location = New-Object System.Drawing.Point(340, 45)
    $datePicker.Size = New-Object System.Drawing.Size(200, 25)
    $datePicker.Format = "Short"
    $datePicker.Enabled = $false
    $tabScan.Controls.Add($datePicker)
    
    # Preview mode checkbox
    $chkPreview = New-Object System.Windows.Forms.CheckBox
    $chkPreview.Text = "Preview Mode (don't delete)"
    $chkPreview.Location = New-Object System.Drawing.Point(560, 45)
    $chkPreview.Size = New-Object System.Drawing.Size(180, 25)
    $chkPreview.Checked = $true
    $tabScan.Controls.Add($chkPreview)
    
    # Scan button
    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Scan for Empty Folders"
    $btnScan.Location = New-Object System.Drawing.Point(20, 80)
    $btnScan.Size = New-Object System.Drawing.Size(150, 30)
    $btnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnScan.ForeColor = [System.Drawing.Color]::White
    $btnScan.FlatStyle = "Flat"
    $btnScan.Enabled = $false
    $tabScan.Controls.Add($btnScan)
    
    # Results grid
    $dgvResults = New-Object System.Windows.Forms.DataGridView
    $dgvResults.Location = New-Object System.Drawing.Point(20, 120)
    $dgvResults.Size = New-Object System.Drawing.Size(730, 250)
    $dgvResults.AllowUserToAddRows = $false
    $dgvResults.AllowUserToDeleteRows = $false
    $dgvResults.SelectionMode = "FullRowSelect"
    $dgvResults.MultiSelect = $true
    $dgvResults.AutoSizeColumnsMode = "Fill"
    $tabScan.Controls.Add($dgvResults)
    
    # Action buttons
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Delete Selected"
    $btnDelete.Location = New-Object System.Drawing.Point(20, 380)
    $btnDelete.Size = New-Object System.Drawing.Size(120, 30)
    $btnDelete.BackColor = [System.Drawing.Color]::FromArgb(217, 83, 79)
    $btnDelete.ForeColor = [System.Drawing.Color]::White
    $btnDelete.FlatStyle = "Flat"
    $btnDelete.Enabled = $false
    $tabScan.Controls.Add($btnDelete)
    
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "Export to CSV"
    $btnExport.Location = New-Object System.Drawing.Point(150, 380)
    $btnExport.Size = New-Object System.Drawing.Size(120, 30)
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(92, 184, 92)
    $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.FlatStyle = "Flat"
    $btnExport.Enabled = $false
    $tabScan.Controls.Add($btnExport)
    
    # Summary label
    $lblSummary = New-Object System.Windows.Forms.Label
    $lblSummary.Text = "No scan performed"
    $lblSummary.Location = New-Object System.Drawing.Point(20, 420)
    $lblSummary.Size = New-Object System.Drawing.Size(730, 30)
    $tabScan.Controls.Add($lblSummary)
    
    $tabControl.TabPages.Add($tabScan)
    
    # Tab 3: Logs
    $tabLogs = New-Object System.Windows.Forms.TabPage
    $tabLogs.Text = "Activity Logs"
    $tabLogs.BackColor = [System.Drawing.Color]::White
    
    $txtLogs = New-Object System.Windows.Forms.RichTextBox
    $txtLogs.Location = New-Object System.Drawing.Point(20, 20)
    $txtLogs.Size = New-Object System.Drawing.Size(730, 400)
    $txtLogs.ReadOnly = $true
    $txtLogs.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtLogs.BackColor = [System.Drawing.Color]::Black
    $txtLogs.ForeColor = [System.Drawing.Color]::LightGreen
    $tabLogs.Controls.Add($txtLogs)
    
    $btnRefreshLogs = New-Object System.Windows.Forms.Button
    $btnRefreshLogs.Text = "Refresh Logs"
    $btnRefreshLogs.Location = New-Object System.Drawing.Point(20, 430)
    $btnRefreshLogs.Size = New-Object System.Drawing.Size(100, 25)
    $tabLogs.Controls.Add($btnRefreshLogs)
    
    $tabControl.TabPages.Add($tabLogs)
    
    $form.Controls.Add($tabControl)
    
    # Status bar
    $statusBar = New-Object System.Windows.Forms.StatusStrip
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = "Ready"
    $statusBar.Items.Add($statusLabel) | Out-Null
    $form.Controls.Add($statusBar)
    
    # Event handlers
    $btnConnect.Add_Click({
        $statusLabel.Text = "Connecting..."
        $form.Refresh()
        
        $siteUrl = $txtSiteUrl.Text.Trim()
        if (-not $siteUrl) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a SharePoint site URL", "Error", "OK", "Error")
            return
        }
        
        Write-Log "INFO" "Attempting to connect to: $siteUrl"
        
        # Connect to SharePoint via Graph
        $Script:Connection = Connect-SharePointGraph -SiteUrl $siteUrl
        
        if ($Script:Connection) {
            $lblConnStatus.Text = "Connected to: $($Script:Connection.Site.DisplayName)`nSite ID: $($Script:Connection.SiteId)"
            $lblConnStatus.ForeColor = [System.Drawing.Color]::Green
            
            # Load libraries
            $libraries = Get-SharePointLibraries -SiteId $Script:Connection.SiteId
            $cmbLibrary.Items.Clear()
            foreach ($lib in $libraries) {
                $cmbLibrary.Items.Add($lib.Name) | Out-Null
            }
            if ($cmbLibrary.Items.Count -gt 0) {
                $cmbLibrary.SelectedIndex = 0
            }
            
            # Enable controls
            $cmbLibrary.Enabled = $true
            $datePicker.Enabled = $true
            $btnScan.Enabled = $true
            $btnConnect.Text = "Reconnect"
            
            $statusLabel.Text = "Connected"
            $tabControl.SelectedIndex = 1
            
            Write-Log "SUCCESS" "Connected to SharePoint successfully"
        } else {
            $lblConnStatus.Text = "Connection failed. Check the URL and try again."
            $lblConnStatus.ForeColor = [System.Drawing.Color]::Red
            $statusLabel.Text = "Connection failed"
            
            Write-Log "ERROR" "Failed to connect to SharePoint"
        }
    })
    
    $btnScan.Add_Click({
        if (-not $Script:Connection) {
            [System.Windows.Forms.MessageBox]::Show("Not connected to SharePoint", "Error", "OK", "Error")
            return
        }
        
        $statusLabel.Text = "Scanning..."
        $form.Refresh()
        
        $libraryName = $cmbLibrary.SelectedItem
        $modifiedDate = $datePicker.Value.Date
        
        Write-Log "INFO" "Scanning library: $libraryName for folders modified on $($modifiedDate.ToString('yyyy-MM-dd'))"
        
        # Get library ID
        $libraries = Get-SharePointLibraries -SiteId $Script:Connection.SiteId
        $library = $libraries | Where-Object { $_.Name -eq $libraryName } | Select-Object -First 1
        
        if ($library) {
            # Find empty folders
            $Script:EmptyFolders = Find-EmptyFolders -SiteId $Script:Connection.SiteId -LibraryId $library.Id -ModifiedDate $modifiedDate
            
            # Display results
            $dgvResults.Columns.Clear()
            $dgvResults.Columns.Add("Name", "Folder Name") | Out-Null
            $dgvResults.Columns.Add("Path", "Path") | Out-Null
            $dgvResults.Columns.Add("Modified", "Modified") | Out-Null
            $dgvResults.Columns.Add("ModifiedBy", "Modified By") | Out-Null
            
            $dgvResults.Rows.Clear()
            foreach ($folder in $Script:EmptyFolders) {
                $dgvResults.Rows.Add($folder.Name, $folder.Path, $folder.Modified, $folder.ModifiedBy) | Out-Null
            }
            
            $lblSummary.Text = "Found $($Script:EmptyFolders.Count) empty folders"
            
            if ($Script:EmptyFolders.Count -gt 0) {
                $btnDelete.Enabled = $true
                $btnExport.Enabled = $true
            } else {
                $btnDelete.Enabled = $false
                $btnExport.Enabled = $false
            }
            
            $statusLabel.Text = "Scan complete"
            Write-Log "SUCCESS" "Scan complete - found $($Script:EmptyFolders.Count) empty folders"
        }
    })
    
    $btnDelete.Add_Click({
        if ($Script:EmptyFolders.Count -eq 0) {
            return
        }
        
        $selectedCount = $dgvResults.SelectedRows.Count
        if ($selectedCount -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select folders to delete", "Info", "OK", "Information")
            return
        }
        
        $message = if ($chkPreview.Checked) {
            "Preview Mode: Would delete $selectedCount folder(s)`n`nNo actual deletion will occur."
        } else {
            "Are you sure you want to delete $selectedCount folder(s)?`n`nThis action cannot be undone!"
        }
        
        $result = [System.Windows.Forms.MessageBox]::Show($message, "Confirm", "YesNo", "Warning")
        
        if ($result -eq "Yes") {
            $statusLabel.Text = "Deleting..."
            $form.Refresh()
            
            $deleted = 0
            $failed = 0
            
            foreach ($row in $dgvResults.SelectedRows) {
                $folderName = $row.Cells["Name"].Value
                $folder = $Script:EmptyFolders | Where-Object { $_.Name -eq $folderName } | Select-Object -First 1
                
                if ($folder) {
                    if ($chkPreview.Checked) {
                        Write-Log "INFO" "[PREVIEW] Would delete: $($folder.Name)"
                        $deleted++
                    } else {
                        $libraryName = $cmbLibrary.SelectedItem
                        $libraries = Get-SharePointLibraries -SiteId $Script:Connection.SiteId
                        $library = $libraries | Where-Object { $_.Name -eq $libraryName } | Select-Object -First 1
                        
                        if (Remove-EmptyFolder -SiteId $Script:Connection.SiteId -LibraryId $library.Id -FolderId $folder.Id) {
                            Write-Log "DELETE-SUCCESS" "Deleted: $($folder.Name)"
                            $deleted++
                        } else {
                            Write-Log "DELETE-FAIL" "Failed to delete: $($folder.Name)"
                            $failed++
                        }
                    }
                }
            }
            
            if ($chkPreview.Checked) {
                [System.Windows.Forms.MessageBox]::Show("Preview complete. Would have deleted $deleted folder(s)", "Preview", "OK", "Information")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Deleted $deleted folder(s). Failed: $failed", "Complete", "OK", "Information")
                # Rescan after deletion
                $btnScan.PerformClick()
            }
            
            $statusLabel.Text = "Ready"
        }
    })
    
    $btnExport.Add_Click({
        if ($Script:EmptyFolders.Count -eq 0) {
            return
        }
        
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
        $saveDialog.FileName = "EmptyFolders_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            Export-FolderReport -Folders $Script:EmptyFolders -Path $saveDialog.FileName
            [System.Windows.Forms.MessageBox]::Show("Report exported successfully", "Success", "OK", "Information")
            Write-Log "INFO" "Exported report to: $($saveDialog.FileName)"
        }
    })
    
    $btnRefreshLogs.Add_Click({
        $logContent = Get-LogContent
        $txtLogs.Text = $logContent -join "`n"
        $txtLogs.SelectionStart = $txtLogs.Text.Length
        $txtLogs.ScrollToCaret()
    })
    
    # Show form
    [void]$form.ShowDialog()
    
    # Cleanup on close
    if ($Script:Connection) {
        Disconnect-SharePointGraph
    }
}

# Export function
Export-ModuleMember -Function Show-CleanupGUI