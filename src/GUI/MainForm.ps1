# MainForm.ps1
# Main GUI application for SharePoint Cleanup Tool

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Load core modules
$scriptPath = Split-Path -Parent $PSScriptRoot
. (Join-Path $scriptPath "Core\SharePointManager.ps1")
. (Join-Path $scriptPath "Core\Logger.ps1")
. (Join-Path $scriptPath "Core\ConfigManager.ps1")

class MainForm {
    [System.Windows.Forms.Form]$Form
    [SharePointManager]$SPManager
    [Logger]$Logger
    [ConfigManager]$Config
    [array]$EmptyFolders = @()
    
    # Controls
    [System.Windows.Forms.TextBox]$txtSiteUrl
    [System.Windows.Forms.ComboBox]$cmbLibrary
    [System.Windows.Forms.DateTimePicker]$datePicker
    [System.Windows.Forms.DataGridView]$dgvResults
    [System.Windows.Forms.CheckBox]$chkPreviewMode
    [System.Windows.Forms.Label]$lblStatus
    [System.Windows.Forms.ProgressBar]$progressBar
    [System.Windows.Forms.RichTextBox]$txtLog
    [System.Windows.Forms.Button]$btnConnect
    [System.Windows.Forms.Button]$btnScan
    [System.Windows.Forms.Button]$btnDelete
    [System.Windows.Forms.TabControl]$tabControl
    
    MainForm() {
        $basePath = Split-Path -Parent $PSScriptRoot
        $this.Config = [ConfigManager]::new($basePath)
        $this.Logger = [Logger]::new($basePath)
        $this.Logger.SetConsoleOutput($false)
        $this.InitializeForm()
    }
    
    [void] InitializeForm() {
        # Create main form
        $this.Form = New-Object System.Windows.Forms.Form
        $this.Form.Text = "SharePoint Cleanup Tool v2.0"
        $this.Form.Size = New-Object System.Drawing.Size(900, 700)
        $this.Form.StartPosition = "CenterScreen"
        $this.Form.FormBorderStyle = "FixedDialog"
        $this.Form.MaximizeBox = $false
        $this.Form.Icon = [System.Drawing.SystemIcons]::Application
        
        # Create menu bar
        $menuBar = New-Object System.Windows.Forms.MenuStrip
        $fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("&File")
        $helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem("&Help")
        
        $exportMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Export Results to CSV")
        $exportMenuItem.Add_Click({ $this.ExportResults() })
        $fileMenu.DropDownItems.Add($exportMenuItem)
        
        $exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
        $exitMenuItem.Add_Click({ $this.Form.Close() })
        $fileMenu.DropDownItems.Add($exitMenuItem)
        
        $aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("About")
        $aboutMenuItem.Add_Click({ $this.ShowAbout() })
        $helpMenu.DropDownItems.Add($aboutMenuItem)
        
        $menuBar.Items.Add($fileMenu)
        $menuBar.Items.Add($helpMenu)
        $this.Form.MainMenuStrip = $menuBar
        $this.Form.Controls.Add($menuBar)
        
        # Create tab control
        $this.tabControl = New-Object System.Windows.Forms.TabControl
        $this.tabControl.Location = New-Object System.Drawing.Point(10, 30)
        $this.tabControl.Size = New-Object System.Drawing.Size(870, 580)
        
        # Connection Tab
        $tabConnection = New-Object System.Windows.Forms.TabPage("Connection")
        $this.CreateConnectionTab($tabConnection)
        $this.tabControl.TabPages.Add($tabConnection)
        
        # Scan Tab
        $tabScan = New-Object System.Windows.Forms.TabPage("Scan & Clean")
        $this.CreateScanTab($tabScan)
        $this.tabControl.TabPages.Add($tabScan)
        
        # Logs Tab
        $tabLogs = New-Object System.Windows.Forms.TabPage("Activity Log")
        $this.CreateLogsTab($tabLogs)
        $this.tabControl.TabPages.Add($tabLogs)
        
        $this.Form.Controls.Add($this.tabControl)
        
        # Status bar
        $statusStrip = New-Object System.Windows.Forms.StatusStrip
        $this.lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel("Ready")
        $statusStrip.Items.Add($this.lblStatus)
        $this.Form.Controls.Add($statusStrip)
        
        # Progress bar at bottom
        $this.progressBar = New-Object System.Windows.Forms.ProgressBar
        $this.progressBar.Location = New-Object System.Drawing.Point(10, 620)
        $this.progressBar.Size = New-Object System.Drawing.Size(870, 20)
        $this.progressBar.Visible = $false
        $this.Form.Controls.Add($this.progressBar)
        
        # Load saved settings
        $this.LoadSettings()
    }
    
    [void] CreateConnectionTab([System.Windows.Forms.TabPage]$tab) {
        # Site URL
        $lblSite = New-Object System.Windows.Forms.Label
        $lblSite.Location = New-Object System.Drawing.Point(20, 30)
        $lblSite.Size = New-Object System.Drawing.Size(100, 20)
        $lblSite.Text = "Site URL:"
        $tab.Controls.Add($lblSite)
        
        $this.txtSiteUrl = New-Object System.Windows.Forms.TextBox
        $this.txtSiteUrl.Location = New-Object System.Drawing.Point(130, 30)
        $this.txtSiteUrl.Size = New-Object System.Drawing.Size(500, 20)
        $tab.Controls.Add($this.txtSiteUrl)
        
        # Recent sites dropdown
        $btnRecent = New-Object System.Windows.Forms.Button
        $btnRecent.Location = New-Object System.Drawing.Point(640, 29)
        $btnRecent.Size = New-Object System.Drawing.Size(100, 22)
        $btnRecent.Text = "Recent Sites"
        $btnRecent.Add_Click({ $this.ShowRecentSites() })
        $tab.Controls.Add($btnRecent)
        
        # Connect button
        $this.btnConnect = New-Object System.Windows.Forms.Button
        $this.btnConnect.Location = New-Object System.Drawing.Point(130, 70)
        $this.btnConnect.Size = New-Object System.Drawing.Size(120, 35)
        $this.btnConnect.Text = "Connect"
        $this.btnConnect.BackColor = [System.Drawing.Color]::LightGreen
        $this.btnConnect.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $this.btnConnect.Add_Click({ $this.ConnectToSharePoint() })
        $tab.Controls.Add($this.btnConnect)
        
        # Connection status
        $grpStatus = New-Object System.Windows.Forms.GroupBox
        $grpStatus.Location = New-Object System.Drawing.Point(20, 120)
        $grpStatus.Size = New-Object System.Drawing.Size(820, 100)
        $grpStatus.Text = "Connection Status"
        $tab.Controls.Add($grpStatus)
        
        $lblConnStatus = New-Object System.Windows.Forms.Label
        $lblConnStatus.Location = New-Object System.Drawing.Point(10, 25)
        $lblConnStatus.Size = New-Object System.Drawing.Size(800, 60)
        $lblConnStatus.Text = "Not connected"
        $lblConnStatus.ForeColor = [System.Drawing.Color]::Red
        $grpStatus.Controls.Add($lblConnStatus)
        
        # Instructions
        $grpInstructions = New-Object System.Windows.Forms.GroupBox
        $grpInstructions.Location = New-Object System.Drawing.Point(20, 240)
        $grpInstructions.Size = New-Object System.Drawing.Size(820, 250)
        $grpInstructions.Text = "Instructions"
        $tab.Controls.Add($grpInstructions)
        
        $txtInstructions = New-Object System.Windows.Forms.RichTextBox
        $txtInstructions.Location = New-Object System.Drawing.Point(10, 25)
        $txtInstructions.Size = New-Object System.Drawing.Size(800, 210)
        $txtInstructions.ReadOnly = $true
        $txtInstructions.Text = @"
AUTHENTICATION:
• Click Connect to authenticate with SharePoint
• A browser window or device code will appear
• Sign in with your organizational account
• Multi-factor authentication is supported

REQUIREMENTS:
• SharePoint site member permissions (minimum)
• Delete permissions on the target library
• PnP.PowerShell module (automatically installed)

WORKFLOW:
1. Connect to your SharePoint site
2. Switch to 'Scan & Clean' tab
3. Select library and date
4. Scan for empty folders
5. Review results
6. Delete selected folders (if preview mode is off)

SAFETY:
• Preview mode is enabled by default
• All actions are logged for audit
• Only completely empty folders can be deleted
"@
        $grpInstructions.Controls.Add($txtInstructions)
    }
    
    [void] CreateScanTab([System.Windows.Forms.TabPage]$tab) {
        # Library selection
        $lblLibrary = New-Object System.Windows.Forms.Label
        $lblLibrary.Location = New-Object System.Drawing.Point(20, 30)
        $lblLibrary.Size = New-Object System.Drawing.Size(100, 20)
        $lblLibrary.Text = "Library:"
        $tab.Controls.Add($lblLibrary)
        
        $this.cmbLibrary = New-Object System.Windows.Forms.ComboBox
        $this.cmbLibrary.Location = New-Object System.Drawing.Point(130, 30)
        $this.cmbLibrary.Size = New-Object System.Drawing.Size(200, 20)
        $this.cmbLibrary.DropDownStyle = "DropDown"
        $this.cmbLibrary.Items.Add("Documents")
        $this.cmbLibrary.Text = "Documents"
        $this.cmbLibrary.Enabled = $false
        $tab.Controls.Add($this.cmbLibrary)
        
        # Date picker
        $lblDate = New-Object System.Windows.Forms.Label
        $lblDate.Location = New-Object System.Drawing.Point(350, 30)
        $lblDate.Size = New-Object System.Drawing.Size(100, 20)
        $lblDate.Text = "Modified Date:"
        $tab.Controls.Add($lblDate)
        
        $this.datePicker = New-Object System.Windows.Forms.DateTimePicker
        $this.datePicker.Location = New-Object System.Drawing.Point(460, 30)
        $this.datePicker.Size = New-Object System.Drawing.Size(200, 20)
        $this.datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
        $this.datePicker.Enabled = $false
        $tab.Controls.Add($this.datePicker)
        
        # Preview mode checkbox
        $this.chkPreviewMode = New-Object System.Windows.Forms.CheckBox
        $this.chkPreviewMode.Location = New-Object System.Drawing.Point(680, 30)
        $this.chkPreviewMode.Size = New-Object System.Drawing.Size(150, 20)
        $this.chkPreviewMode.Text = "Preview Mode"
        $this.chkPreviewMode.Checked = $true
        $this.chkPreviewMode.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
        $this.chkPreviewMode.ForeColor = [System.Drawing.Color]::DarkGreen
        $tab.Controls.Add($this.chkPreviewMode)
        
        # Scan button
        $this.btnScan = New-Object System.Windows.Forms.Button
        $this.btnScan.Location = New-Object System.Drawing.Point(130, 70)
        $this.btnScan.Size = New-Object System.Drawing.Size(120, 35)
        $this.btnScan.Text = "Scan Folders"
        $this.btnScan.BackColor = [System.Drawing.Color]::LightBlue
        $this.btnScan.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $this.btnScan.Enabled = $false
        $this.btnScan.Add_Click({ $this.ScanFolders() })
        $tab.Controls.Add($this.btnScan)
        
        # Delete button
        $this.btnDelete = New-Object System.Windows.Forms.Button
        $this.btnDelete.Location = New-Object System.Drawing.Point(260, 70)
        $this.btnDelete.Size = New-Object System.Drawing.Size(120, 35)
        $this.btnDelete.Text = "Delete Selected"
        $this.btnDelete.BackColor = [System.Drawing.Color]::LightCoral
        $this.btnDelete.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $this.btnDelete.Enabled = $false
        $this.btnDelete.Add_Click({ $this.DeleteFolders() })
        $tab.Controls.Add($this.btnDelete)
        
        # Results grid
        $lblResults = New-Object System.Windows.Forms.Label
        $lblResults.Location = New-Object System.Drawing.Point(20, 120)
        $lblResults.Size = New-Object System.Drawing.Size(100, 20)
        $lblResults.Text = "Results:"
        $lblResults.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $tab.Controls.Add($lblResults)
        
        $this.dgvResults = New-Object System.Windows.Forms.DataGridView
        $this.dgvResults.Location = New-Object System.Drawing.Point(20, 145)
        $this.dgvResults.Size = New-Object System.Drawing.Size(820, 350)
        $this.dgvResults.AllowUserToAddRows = $false
        $this.dgvResults.AllowUserToDeleteRows = $false
        $this.dgvResults.SelectionMode = 'FullRowSelect'
        $this.dgvResults.MultiSelect = $true
        $this.dgvResults.AutoSizeColumnsMode = 'Fill'
        $tab.Controls.Add($this.dgvResults)
        
        # Summary label
        $lblSummary = New-Object System.Windows.Forms.Label
        $lblSummary.Location = New-Object System.Drawing.Point(20, 500)
        $lblSummary.Size = New-Object System.Drawing.Size(820, 30)
        $lblSummary.Text = "No scan performed"
        $lblSummary.Font = New-Object System.Drawing.Font("Arial", 9)
        $tab.Controls.Add($lblSummary)
    }
    
    [void] CreateLogsTab([System.Windows.Forms.TabPage]$tab) {
        $this.txtLog = New-Object System.Windows.Forms.RichTextBox
        $this.txtLog.Location = New-Object System.Drawing.Point(20, 20)
        $this.txtLog.Size = New-Object System.Drawing.Size(820, 450)
        $this.txtLog.ReadOnly = $true
        $this.txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
        $this.txtLog.BackColor = [System.Drawing.Color]::Black
        $this.txtLog.ForeColor = [System.Drawing.Color]::LightGreen
        $tab.Controls.Add($this.txtLog)
        
        # Buttons
        $btnClearLog = New-Object System.Windows.Forms.Button
        $btnClearLog.Location = New-Object System.Drawing.Point(20, 480)
        $btnClearLog.Size = New-Object System.Drawing.Size(100, 30)
        $btnClearLog.Text = "Clear Log"
        $btnClearLog.Add_Click({ $this.txtLog.Clear() })
        $tab.Controls.Add($btnClearLog)
        
        $btnOpenLogFile = New-Object System.Windows.Forms.Button
        $btnOpenLogFile.Location = New-Object System.Drawing.Point(130, 480)
        $btnOpenLogFile.Size = New-Object System.Drawing.Size(100, 30)
        $btnOpenLogFile.Text = "Open Log File"
        $btnOpenLogFile.Add_Click({ 
            if ($this.Logger) {
                Start-Process notepad $this.Logger.GetLogPath()
            }
        })
        $tab.Controls.Add($btnOpenLogFile)
    }
    
    [void] LoadSettings() {
        $lastSite = $this.Config.Get("LastUsedSite")
        if ($lastSite) {
            $this.txtSiteUrl.Text = $lastSite
        }
        
        $this.chkPreviewMode.Checked = $this.Config.Get("PreviewMode")
        $defaultLib = $this.Config.Get("DefaultLibrary")
        if ($defaultLib) {
            $this.cmbLibrary.Text = $defaultLib
        }
    }
    
    [void] ShowRecentSites() {
        $recentSites = $this.Config.GetRecentSites()
        if ($recentSites.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No recent sites", "Recent Sites")
            return
        }
        
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Select Recent Site"
        $form.Size = New-Object System.Drawing.Size(500, 200)
        $form.StartPosition = "CenterScreen"
        
        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point(10, 10)
        $listBox.Size = New-Object System.Drawing.Size(460, 100)
        foreach ($site in $recentSites) {
            $listBox.Items.Add($site)
        }
        $form.Controls.Add($listBox)
        
        $btnSelect = New-Object System.Windows.Forms.Button
        $btnSelect.Location = New-Object System.Drawing.Point(150, 120)
        $btnSelect.Size = New-Object System.Drawing.Size(100, 30)
        $btnSelect.Text = "Select"
        $btnSelect.Add_Click({
            if ($listBox.SelectedItem) {
                $this.txtSiteUrl.Text = $listBox.SelectedItem
                $form.Close()
            }
        })
        $form.Controls.Add($btnSelect)
        
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(260, 120)
        $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
        $btnCancel.Text = "Cancel"
        $btnCancel.Add_Click({ $form.Close() })
        $form.Controls.Add($btnCancel)
        
        $form.ShowDialog()
    }
    
    [void] ConnectToSharePoint() {
        $siteUrl = $this.txtSiteUrl.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a SharePoint site URL", "Connection Error")
            return
        }
        
        $this.btnConnect.Enabled = $false
        $this.lblStatus.Text = "Connecting to SharePoint..."
        $this.LogToUI("INFO", "Connecting to $siteUrl...")
        $this.Form.Refresh()
        
        try {
            $this.SPManager = [SharePointManager]::new($siteUrl)
            
            if ($this.SPManager.Connect()) {
                $this.Config.SaveRecentSite($siteUrl)
                $this.lblStatus.Text = "Connected to SharePoint"
                $this.LogToUI("SUCCESS", "Connected successfully to $siteUrl")
                $this.Logger.LogInfo("Connected to SharePoint: $siteUrl")
                
                # Enable controls
                $this.cmbLibrary.Enabled = $true
                $this.datePicker.Enabled = $true
                $this.btnScan.Enabled = $true
                $this.btnConnect.Text = "Reconnect"
                
                # Update status in Connection tab
                $connStatus = $this.tabControl.TabPages[0].Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] -and $_.Text -eq "Connection Status" }
                if ($connStatus) {
                    $statusLabel = $connStatus[0].Controls[0]
                    $statusLabel.Text = "Connected to: $siteUrl`nAuthentication successful`nReady to scan folders"
                    $statusLabel.ForeColor = [System.Drawing.Color]::Green
                }
                
                # Switch to Scan tab
                $this.tabControl.SelectedIndex = 1
            }
            else {
                throw "Authentication failed"
            }
        }
        catch {
            $this.lblStatus.Text = "Connection failed"
            $this.LogToUI("ERROR", "Connection failed: $_")
            $this.Logger.LogError("Connection failed: $_")
            [System.Windows.Forms.MessageBox]::Show("Failed to connect: $_", "Connection Error")
        }
        finally {
            $this.btnConnect.Enabled = $true
        }
    }
    
    [void] ScanFolders() {
        $libraryName = $this.cmbLibrary.Text.Trim()
        $modifiedDate = $this.datePicker.Value.Date
        
        if ([string]::IsNullOrWhiteSpace($libraryName)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a library name", "Scan Error")
            return
        }
        
        $this.btnScan.Enabled = $false
        $this.progressBar.Visible = $true
        $this.progressBar.Style = 'Marquee'
        $this.lblStatus.Text = "Scanning folders..."
        $this.LogToUI("INFO", "Starting scan - Library: $libraryName, Date: $($modifiedDate.ToShortDateString())")
        $this.Form.Refresh()
        
        try {
            $this.EmptyFolders = $this.SPManager.FindEmptyFolders($libraryName, $modifiedDate)
            
            # Create DataTable for results
            $dt = New-Object System.Data.DataTable
            $dt.Columns.Add("Select", [bool])
            $dt.Columns.Add("Name", [string])
            $dt.Columns.Add("Path", [string])
            $dt.Columns.Add("Modified", [datetime])
            $dt.Columns.Add("CreatedBy", [string])
            $dt.Columns.Add("Id", [int])
            
            foreach ($folder in $this.EmptyFolders) {
                $row = $dt.NewRow()
                $row["Select"] = $true
                $row["Name"] = $folder.Name
                $row["Path"] = $folder.Path
                $row["Modified"] = $folder.Modified
                $row["CreatedBy"] = $folder.CreatedBy
                $row["Id"] = $folder.Id
                $dt.Rows.Add($row)
            }
            
            $this.dgvResults.DataSource = $dt
            $this.dgvResults.Columns["Id"].Visible = $false
            $this.dgvResults.Columns["Select"].Width = 50
            $this.dgvResults.Columns["Name"].Width = 200
            
            $summaryLabel = $this.tabControl.TabPages[1].Controls | Where-Object { $_.Location.Y -eq 500 }
            if ($summaryLabel) {
                $summaryLabel[0].Text = "Found $($this.EmptyFolders.Count) empty folders"
                if ($this.EmptyFolders.Count -gt 0) {
                    $summaryLabel[0].Text += " - " + $(if ($this.chkPreviewMode.Checked) { "PREVIEW MODE" } else { "LIVE MODE - Deletions will be permanent!" })
                }
            }
            
            $this.lblStatus.Text = "Scan complete - Found $($this.EmptyFolders.Count) empty folders"
            $this.LogToUI("SUCCESS", "Scan complete - Found $($this.EmptyFolders.Count) empty folders")
            $this.Logger.LogInfo("Scan complete - Library: $libraryName, Date: $($modifiedDate.ToShortDateString()), Empty folders: $($this.EmptyFolders.Count)")
            
            if ($this.EmptyFolders.Count -gt 0) {
                $this.btnDelete.Enabled = $true
            }
        }
        catch {
            $this.lblStatus.Text = "Scan failed"
            $this.LogToUI("ERROR", "Scan failed: $_")
            $this.Logger.LogError("Scan failed: $_")
            [System.Windows.Forms.MessageBox]::Show("Scan failed: $_", "Scan Error")
        }
        finally {
            $this.btnScan.Enabled = $true
            $this.progressBar.Visible = $false
        }
    }
    
    [void] DeleteFolders() {
        $dt = $this.dgvResults.DataSource
        $selectedRows = $dt.Select("Select = true")
        
        if ($selectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No folders selected for deletion", "Delete Error")
            return
        }
        
        $message = "You are about to delete $($selectedRows.Count) folders."
        if ($this.chkPreviewMode.Checked) {
            $message += "`n`nPREVIEW MODE: No actual deletions will occur."
        }
        else {
            $message += "`n`nWARNING: This action cannot be undone!"
        }
        $message += "`n`nContinue?"
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            $message,
            "Confirm Deletion",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
        
        $this.btnDelete.Enabled = $false
        $this.progressBar.Visible = $true
        $this.progressBar.Style = 'Blocks'
        $this.progressBar.Maximum = $selectedRows.Count
        $this.progressBar.Value = 0
        
        $successCount = 0
        $failCount = 0
        
        foreach ($row in $selectedRows) {
            $folderName = $row["Name"]
            $folderId = $row["Id"]
            
            $this.lblStatus.Text = "Processing: $folderName"
            $this.Form.Refresh()
            
            if ($this.chkPreviewMode.Checked) {
                $this.LogToUI("PREVIEW", "Would delete: $folderName")
                $this.Logger.LogInfo("[PREVIEW] Would delete folder: $folderName")
                $successCount++
            }
            else {
                try {
                    if ($this.SPManager.DeleteFolder($this.cmbLibrary.Text, $folderId)) {
                        $this.LogToUI("SUCCESS", "Deleted: $folderName")
                        $this.Logger.LogDeletion($folderName, $row["Path"], $true)
                        $successCount++
                        $row["Select"] = $false
                    }
                    else {
                        throw "Delete operation returned false"
                    }
                }
                catch {
                    $this.LogToUI("ERROR", "Failed to delete: $folderName - $_")
                    $this.Logger.LogDeletion($folderName, $row["Path"], $false)
                    $failCount++
                }
            }
            
            $this.progressBar.Value++
        }
        
        $this.lblStatus.Text = "Deletion complete - Success: $successCount, Failed: $failCount"
        $this.LogToUI("INFO", "Deletion complete - Success: $successCount, Failed: $failCount")
        $this.Logger.LogInfo("Deletion batch complete - Success: $successCount, Failed: $failCount")
        
        if (-not $this.chkPreviewMode.Checked -and $successCount -gt 0) {
            # Refresh the grid
            $this.ScanFolders()
        }
        
        $this.progressBar.Visible = $false
        $this.btnDelete.Enabled = $true
    }
    
    [void] ExportResults() {
        if ($this.EmptyFolders.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No results to export", "Export Error")
            return
        }
        
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Files|*.csv"
        $saveDialog.FileName = "SharePoint_Empty_Folders_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $this.EmptyFolders | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
                [System.Windows.Forms.MessageBox]::Show("Results exported successfully", "Export Complete")
                $this.Logger.LogInfo("Exported results to: $($saveDialog.FileName)")
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Export failed: $_", "Export Error")
            }
        }
    }
    
    [void] ShowAbout() {
        $about = @"
SharePoint Cleanup Tool v2.0

A comprehensive tool for managing empty folders in SharePoint Online.

Features:
• Windows integrated authentication
• Batch folder deletion
• Preview mode for safety
• Full audit logging
• Export capabilities

© 2025 - Enterprise Administration Tool
"@
        [System.Windows.Forms.MessageBox]::Show($about, "About", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    
    [void] LogToUI([string]$level, [string]$message) {
        if ($null -eq $this.txtLog) { return }
        
        $timestamp = Get-Date -Format "HH:mm:ss"
        $logEntry = "[$timestamp] [$level] $message`r`n"
        
        # Color coding
        $color = switch ($level) {
            "ERROR" { [System.Drawing.Color]::Red }
            "WARNING" { [System.Drawing.Color]::Yellow }
            "SUCCESS" { [System.Drawing.Color]::LightGreen }
            "INFO" { [System.Drawing.Color]::Cyan }
            "PREVIEW" { [System.Drawing.Color]::Orange }
            default { [System.Drawing.Color]::White }
        }
        
        $this.txtLog.SelectionStart = $this.txtLog.TextLength
        $this.txtLog.SelectionColor = $color
        $this.txtLog.AppendText($logEntry)
        $this.txtLog.ScrollToCaret()
    }
    
    [void] Run() {
        $this.Form.ShowDialog()
    }
}

# Run the application
$app = [MainForm]::new()
$app.Run()