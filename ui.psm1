<# ui.psm1
    .SYNOPSIS
        Constructs, manages, and defines the logic for the entire graphical user interface (GUI) of the Entropia Dashboard except for the Ftool. 
		It uses Windows Forms and custom-drawn controls to create a dark-themed, interactive frontend for the application's features.

    .DESCRIPTION
        This PowerShell module is exclusively responsible for the application's user interface. It dynamically builds all visual components, registers event handlers for user interactions, and manages the synchronization of settings between the UI and the application's configuration state. The UI is built using Windows Forms, but heavily customized with owner-drawn controls from `classes.psm1` to create a consistent dark-mode aesthetic.
#>

#region Helper Functions

#region Function: Show-InputBox
function Show-InputBox
{
    param(
        [string]$Title,
        [string]$Prompt,
        [string]$DefaultText
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(300, 150)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $form.ForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    # FIX: Ensure this dialog stays on top of the TopMost SettingsForm
    $form.TopMost = $true

    # Try to set icon if available
    if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon)) {
        try { $form.Icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon) } catch {}
    }

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(260, 20)
    $label.Text = $Prompt
    $label.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 40)
    $textBox.Size = New-Object System.Drawing.Size(260, 25)
    $textBox.Text = $DefaultText
    $textBox.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $textBox.ForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $textBox.BorderStyle = 'FixedSingle'
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(110, 80)
    $okButton.Size = New-Object System.Drawing.Size(75, 25)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.FlatStyle = 'Flat'
    $okButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60,60,60)
    $okButton.BackColor = [System.Drawing.Color]::FromArgb(40,40,40)
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(195, 80)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 25)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.FlatStyle = 'Flat'
    $cancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60,60,60)
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(40,40,40)
    $form.Controls.Add($cancelButton)

    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    # FIX: Pass the SettingsForm as the owner if it exists, ensuring correct Z-order
    if ($global:DashboardConfig.UI.SettingsForm) {
        $result = $form.ShowDialog($global:DashboardConfig.UI.SettingsForm)
    } else {
        $result = $form.ShowDialog()
    }

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        return $null
    }
}
#endregion Function: Show-InputBox

#region Function: Sync-UIToConfig
function Sync-UIToConfig
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try
    {
        Write-Verbose '  UI: Syncing UI to config' -ForegroundColor Cyan
        $UI = $global:DashboardConfig.UI
        if (-not ($UI -and $global:DashboardConfig.Config)) { return $false }

        # Ensure Sections
        @('LauncherPath', 'ProcessName', 'MaxClients', 'Login', 'LoginConfig', 'Options', 'Paths', 'Profiles') |
        ForEach-Object { 
            if (-not $global:DashboardConfig.Config.Contains($_)) {
                $global:DashboardConfig.Config[$_] = [ordered]@{}
            }
        }

        # General
        $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] = $UI.InputLauncher.Text
        $global:DashboardConfig.Config['ProcessName']['ProcessName'] = $UI.InputProcess.Text
        $global:DashboardConfig.Config['MaxClients']['MaxClients'] = $UI.InputMax.Text
        $global:DashboardConfig.Config['Paths']['JunctionTarget'] = $UI.InputJunction.Text

        # Checkboxes
        $global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin'] = if ($UI.NeverRestartingCollectorLogin.Checked) { '1' } else { '0' }
        $global:DashboardConfig.Config['Options']['HideMinimizedWindows'] = if ($UI.HideMinimizedWindows.Checked) { '1' } else { '0' } 
        $global:DashboardConfig.Config['Options']['AutoReconnect'] = if ($UI.AutoReconnect.Checked) { '1' } else { '0' }

        # Profiles (New Grid)
        $global:DashboardConfig.Config['Profiles'].Clear()
        foreach ($row in $UI.ProfileGrid.Rows) {
            if ($row.Cells[0].Value -and $row.Cells[1].Value) {
                # Key: Profile Name, Value: Path
                $key = $row.Cells[0].Value.ToString()
                $val = $row.Cells[1].Value.ToString()
                $global:DashboardConfig.Config['Profiles'][$key] = $val
            }
        }
        
        # Save Selected Profile to Options so it persists
        if ($UI.ProfileGrid.SelectedRows.Count -gt 0) {
            $val = $UI.ProfileGrid.SelectedRows[0].Cells[0].Value
            if ($val) {
                $global:DashboardConfig.Config['Options']['SelectedProfile'] = $val.ToString()
            } else {
                $global:DashboardConfig.Config['Options']['SelectedProfile'] = ""
            }
        } else {
            $global:DashboardConfig.Config['Options']['SelectedProfile'] = ""
        }

        # Login Config (Coords)
        foreach ($key in $UI.LoginPickers.Keys) {
            $global:DashboardConfig.Config['LoginConfig']["${key}Coords"] = $UI.LoginPickers[$key].Text.Text
        }
        $global:DashboardConfig.Config['LoginConfig']['PostLoginDelay'] = $UI.InputPostLoginDelay.Text
        
        # Login Config (Grid)
        $grid = $UI.LoginConfigGrid
        foreach ($row in $grid.Rows) {
            $clientNum = $row.Cells[0].Value
            $s = if ($row.Cells[1].Value) { $row.Cells[1].Value.ToString() } else { "1" }
            $c = if ($row.Cells[2].Value) { $row.Cells[2].Value.ToString() } else { "1" }
            $char = if ($row.Cells[3].Value) { $row.Cells[3].Value.ToString() } else { "1" }
            $coll = if ($row.Cells[4].Value) { $row.Cells[4].Value.ToString() } else { "No" }
            
            $val = "$s,$c,$char,$coll"
            $global:DashboardConfig.Config['LoginConfig']["Client${clientNum}_Settings"] = $val
        }

        Write-Verbose '  UI: UI synced to config' -ForegroundColor Green
        return $true
    }
    catch
    {
        Write-Verbose "  UI: Failed to sync UI to config: $_" -ForegroundColor Red
        return $false
    }
}


#endregion Function: Sync-UIToConfig

#region Function: Sync-ConfigToUI
function Sync-ConfigToUI
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try
    {
     
        Write-Verbose '  UI: Syncing config to UI' -ForegroundColor Cyan
        $UI = $global:DashboardConfig.UI
        if (-not ($UI -and $global:DashboardConfig.Config)) { return $false }

        # General
        if ($global:DashboardConfig.Config['LauncherPath']['LauncherPath']) { $UI.InputLauncher.Text = $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] }
        if ($global:DashboardConfig.Config['ProcessName']['ProcessName']) { $UI.InputProcess.Text = $global:DashboardConfig.Config['ProcessName']['ProcessName'] }
        if ($global:DashboardConfig.Config['MaxClients']['MaxClients']) { $UI.InputMax.Text = $global:DashboardConfig.Config['MaxClients']['MaxClients'] }
        if ($global:DashboardConfig.Config['Paths']['JunctionTarget']) { $UI.InputJunction.Text = $global:DashboardConfig.Config['Paths']['JunctionTarget'] }

        # Checkboxes
        if ($global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin']) { $UI.NeverRestartingCollectorLogin.Checked = ([int]$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin']) -eq 1 }
        if ($global:DashboardConfig.Config['Options']['HideMinimizedWindows']) { $UI.HideMinimizedWindows.Checked = ([int]$global:DashboardConfig.Config['Options']['HideMinimizedWindows']) -eq 1 }
        if ($global:DashboardConfig.Config['Options']['AutoReconnect']) { $UI.AutoReconnect.Checked = ([int]$global:DashboardConfig.Config['Options']['AutoReconnect']) -eq 1 }

        # Profiles
        $UI.ProfileGrid.Rows.Clear()
        if ($global:DashboardConfig.Config['Profiles']) {
            $profiles = $global:DashboardConfig.Config['Profiles']
            foreach ($key in $profiles.Keys) {
                $path = $profiles[$key]
                if (Test-Path $path) {
                    $UI.ProfileGrid.Rows.Add($key, $path) | Out-Null
                }
            }
        }
        
        # Restore Selected Profile
        $selectedProfileName = $null
        if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['SelectedProfile']) {
            $selectedProfileName = $global:DashboardConfig.Config['Options']['SelectedProfile'].ToString()
        }
        
        if (-not [string]::IsNullOrEmpty($selectedProfileName)) {
            $UI.ProfileGrid.ClearSelection()
            $found = $false
            foreach ($row in $UI.ProfileGrid.Rows) {
                if ($row.Cells[0].Value.ToString() -eq $selectedProfileName) {
                    $row.Selected = $true
                    $UI.ProfileGrid.CurrentCell = $row.Cells[0]
                    $found = $true
                    break
                }
            }
            if (-not $found) { $UI.ProfileGrid.ClearSelection() }
        } else {
             $UI.ProfileGrid.ClearSelection()
        }

        # Login Config
        if ($global:DashboardConfig.Config['LoginConfig']) {
            $lc = $global:DashboardConfig.Config['LoginConfig']
            
            # Coords
            foreach ($key in $UI.LoginPickers.Keys) {
                $coordKey = "${key}Coords"
                if ($lc[$coordKey]) {
                    $UI.LoginPickers[$key].Text.Text = $lc[$coordKey]
                }
            }
            if ($lc['PostLoginDelay']) { $UI.InputPostLoginDelay.Text = $lc['PostLoginDelay'] }
            
            # Grid
            $grid = $UI.LoginConfigGrid
            foreach ($row in $grid.Rows) {
                $clientNum = $row.Cells[0].Value
                $settingKey = "Client${clientNum}_Settings"
                
                $s="1"; $c="1"; $char="1"; $coll="No"
                
                if ($lc[$settingKey]) {
                    $parts = $lc[$settingKey] -split ','
                    if ($parts.Count -eq 4) {
                        if ($parts[0] -in @("1","2")) { $s = $parts[0] }
                        if ($parts[1] -in @("1","2")) { $c = $parts[1] }
                        if ($parts[2] -in @("1","2","3")) { $char = $parts[2] }
                        if ($parts[3] -in @("Yes","No")) { $coll = $parts[3] }
                    }
                }
                
                $row.Cells[1].Value = $s
                $row.Cells[2].Value = $c
                $row.Cells[3].Value = $char
                $row.Cells[4].Value = $coll
            }
        }

        Write-Verbose '  UI: Config synced to UI' -ForegroundColor Green
        return $true
    }
    catch
    {
        Write-Verbose "  UI: Failed to sync config to UI: $_" -ForegroundColor Red
        return $false
    }
}
#endregion Function: Sync-ConfigToUI

#endregion Helper Functions

#region Core UI Functions

#region Function: Initialize-UI
function Initialize-UI
{
    [CmdletBinding()]
    param()

    Write-Verbose '  UI: Initializing UI...' -ForegroundColor Cyan

    $uiPropertiesToAdd = @{}

    #region Step: Create Main UI Elements
    $p = @{ type='Form'; visible=$false; width=470; height=440; bg=@(30, 30, 30); id='MainForm'; text='Entropia Dashboard'; startPosition='CenterScreen'; formBorderStyle=[System.Windows.Forms.FormBorderStyle]::None }
    $mainForm = Set-UIElement @p

    if (-not $global:DashboardConfig.UI) { $global:DashboardConfig | Add-Member -MemberType NoteProperty -Name UI -Value ([PSCustomObject]@{}) -Force }

    $toolTipMain = New-Object System.Windows.Forms.ToolTip
    $toolTipMain.AutoPopDelay = 5000 
    $toolTipMain.InitialDelay = 100 
    $toolTipMain.ReshowDelay = 10
    $toolTipMain.ShowAlways = $true
    
    $global:DashboardConfig.UI | Add-Member -MemberType NoteProperty -Name ToolTip -Value $toolTipMain -Force

    #region Step: Settings Form
    $p = @{ type='Form'; width=600; height=550; bg=@(30, 30, 30); id='SettingsForm'; text='Settings'; startPosition='CenterScreen'; formBorderStyle=[System.Windows.Forms.FormBorderStyle]::None; topMost=$true; opacity=0.0 }
    $settingsForm = Set-UIElement @p
    #endregion

    if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon)) {
        try {
            $icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon)
            $mainForm.Icon = $icon
            $settingsForm.Icon = $icon
        } catch {}
    }

    $p = @{ type='Panel'; width=470; height=30; bg=@(20, 20, 20); id='TopBar' }
    $topBar = Set-UIElement @p
    $p = @{ type='Label'; width=140; height=12; top=5; left=10; fg=@(240, 240, 240); id='TitleLabel'; text='Entropia Dashboard'; font=(New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)) }
    $titleLabelForm = Set-UIElement @p
    $p = @{ type='Label'; width=140; height=10; top=16; left=10; fg=@(230, 230, 230); id='CopyrightLabel'; text=[char]0x00A9 + ' Immortal / Divine 2025 - v1.4'; font=(New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Italic)) }
    $copyrightLabelForm = Set-UIElement @p
    $p = @{ type='Button'; width=30; height=30; left=410; bg=@(40, 40, 40); fg=@(240, 240, 240); id='MinForm'; text='_'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)); tooltip='Minimize' }
    $btnMinimizeForm = Set-UIElement @p
    $p = @{ type='Button'; width=30; height=30; left=440; bg=@(150, 20, 20); fg=@(240, 240, 240); id='CloseForm'; text=[char]0x166D; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)); tooltip='Exit' }
    $btnCloseForm = Set-UIElement @p
    #endregion

    #region Step: Main Form Buttons & Grids
    $p = @{ type='Button'; width=125; height=30; top=40; left=15; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Launch'; text='Launch'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Start Launch Process / Right Click for single Client (check settings)' }
    $btnLaunch = Set-UIElement @p
    
    # --- CONTEXT MENU CREATION (FIXED) ---
    $LaunchContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $LaunchContextMenu.Name = 'LaunchContextMenu'
    # Add dummy item so menu is not empty on first click
    $LaunchContextMenu.Items.Add("Loading...") | Out-Null
    $btnLaunch.ContextMenuStrip = $LaunchContextMenu
    # -------------------------------------

    $p = @{ type='Button'; width=125; height=30; top=40; left=150; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Login'; text='Login'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Login selected Clients (check settings) / Nickname List with 10 nicknames and 1024x768 mandatory' }
    $btnLogin = Set-UIElement @p
    $p = @{ type='Button'; width=80; height=30; top=40; left=285; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Settings'; text='Settings'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Edit Dashboard Settings' }
    $btnSettings = Set-UIElement @p
    $p = @{ type='Button'; width=80; height=30; top=40; left=375; bg=@(150, 20, 20); fg=@(240, 240, 240); id='Terminate'; text='Terminate'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Closes all selected Clients' }
    $btnStop = Set-UIElement @p
    $p = @{ type='Button'; width=440; height=30; top=75; left=15; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Ftool'; text='Ftool'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Start Ftool for selected Clients' }
    $btnFtool = Set-UIElement @p
    
    $p = @{ type='DataGridView'; visible=$false; width=155; height=320; top=115; left=5; bg=@(40, 40, 40); fg=@(240, 240, 240); id='DataGridMain'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $DataGridMain = Set-UIElement @p
    
    $p = @{ type='DataGridView'; width=450; height=300; top=115; left=10; bg=@(40, 40, 40); fg=@(240, 240, 240); id='DataGridFiller'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $DataGridFiller = Set-UIElement @p
    
    # --- ADD MAIN GRID COLUMNS ---
    $mainGridCols = @(
        (New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Index'; HeaderText = '#'; FillWeight = 8; SortMode = 'NotSortable';}),
        (New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Titel'; HeaderText = 'Titel'; SortMode = 'NotSortable';}),
        (New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'ID'; HeaderText = 'ID'; FillWeight = 20; SortMode = 'NotSortable';}),
        (New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'State'; HeaderText = 'State'; FillWeight = 40; SortMode = 'NotSortable';})
    )
    $DataGridFiller.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$mainGridCols)
    
    # USE CUSTOM TEXT PROGRESS BAR
    $p = @{ type='TextProgressBar'; width=450; height=18; top=415; left=10; id='GlobalProgressBar'; style='Continuous'; visible=$false }
    $GlobalProgressBar = Set-UIElement @p
    #endregion

    #region Step: Settings Form with Tabs
    
    # Create a completely separate ToolTip object for the Settings Form.
    $toolTipSettings = New-Object System.Windows.Forms.ToolTip
    $toolTipSettings.AutoPopDelay = 5000 
    $toolTipSettings.InitialDelay = 100 
    $toolTipSettings.ReshowDelay = 10
    $toolTipSettings.ShowAlways = $true
    
    # Update the global UI object to point to this new tooltip. 
    $global:DashboardConfig.UI.ToolTip = $toolTipSettings
    # -----------------------------------------------------

    # USE CUSTOM DARK TAB CONTROL
    $settingsTabs = New-Object Custom.DarkTabControl
    $settingsTabs.Dock = 'Top'
    $settingsTabs.Height = 505
    
    $tabGeneral = New-Object System.Windows.Forms.TabPage "General"
    $tabGeneral.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    
    $tabLoginSettings = New-Object System.Windows.Forms.TabPage "Login Settings"
    $tabLoginSettings.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $tabLoginSettings.AutoScroll = $true 
    
    $settingsTabs.TabPages.Add($tabGeneral)
    $settingsTabs.TabPages.Add($tabLoginSettings)
    $settingsForm.Controls.Add($settingsTabs)

    # --- GENERAL TAB CONTROLS (Left Side) ---
    $p = @{ type='Label'; width=125; height=20; top=25; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelLauncher'; text='Select Main Launcher Path:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Select your Main Launcher, recommended as Collector Client' }
    $lblLauncher = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=50; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputLauncher'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Selected folder of your Main Launcher' }
    $txtLauncher = Set-UIElement @p
    $p = @{ type='Button'; width=55; height=25; top=20; left=150; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Browse'; text='Browse'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Select your Main Launcher' }
    $btnBrowseLauncher = Set-UIElement @p
    
    $p = @{ type='Label'; width=85; height=20; top=95; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelProcess'; text='Process Name:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Enter Process Name' }
    $lblProcessName = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=120; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputProcess'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Enter Process Name' }
    $txtProcessName = Set-UIElement @p
    
    $p = @{ type='Label'; width=250; height=20; top=165; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelMax'; text='Max Total Clients For Default Launch:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Set maximum total amount of clients, that the client is allowed to launch' }
    $lblMaxClients = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=190; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputMax'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Set maximum total amount of clients, that the client is allowed to launch' }
    $txtMaxClients = Set-UIElement @p

    # --- CLIENT JUNCTION CONTROLS ---
    $p = @{ type='Label'; width=125; height=20; top=235; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelJunction'; text='Select Profiles Folder:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Folder to create Junctions/Copy' }
    $lblJunction = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=260; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputJunction'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Selected folder to create Junctions/Copy' }
    $txtJunction = Set-UIElement @p
    
    $p = @{ type='Button'; width=55; height=25; top=230; left=145; bg=@(40, 40, 40); fg=@(240, 240, 240); id='BrowseJunction'; text='Browse'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Select Target Folder' }
    $btnBrowseJunction = Set-UIElement @p
    
    $p = @{ type='Button'; width=55; height=25; top=230; left=215; bg=@(35, 175, 75); fg=@(240, 240, 240); id='StartJunction'; text='Create'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Create a lightweight copy of your main client with separate Client settings ' }
    $btnStartJunction = Set-UIElement @p
    
    # Checkboxes
    $p = @{ type='CheckBox'; width=250; height=20; top=300; left=0; bg=@(30, 30, 30); fg=@(240, 240, 240); id='HideMinimizedWindows'; text='Hide minimized clients'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Hides from Taskbar and Alt+Tab View.' }
    $chkHideMinimizedWindows = Set-UIElement @p
    $p = @{ type='CheckBox'; width=200; height=20; top=325; left=0; bg=@(30, 30, 30); fg=@(240, 240, 240); id='NeverRestartingCollectorLogin'; text='Collector Double Click'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='In rare cases the Collector Button has to be clicked twice, tick this checkbox to fix this' }
    $chkNeverRestartingLogin = Set-UIElement @p
    $p = @{ type='CheckBox'; width=320; height=20; top=350; left=0; bg=@(30, 30, 30); fg=@(240, 240, 240); id='AutoReconnect'; text='Enable Auto Reconnect (requires auditing) - WIP'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Run the login sequence when a client disconnects. Requires Windows Auditing.' }
    $chkAutoReconnect = Set-UIElement @p

    # --- GENERAL TAB (Right Side) - Profiles Grid ---
    $p = @{ type='Label'; width=220; height=20; top=25; left=300; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelProfiles'; text='Select Client Profile for Launch:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $lblProfiles = Set-UIElement @p

    $p = @{ type='DataGridView'; width=260; height=230; top=50; left=300; bg=@(40, 40, 40); fg=@(240, 240, 240); id='ProfileGrid'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='List of available Client Profiles (Folders)' }
    $ProfileGrid = Set-UIElement @p
    $ProfileGrid.AllowUserToAddRows = $false
    $ProfileGrid.RowHeadersVisible = $false
    $ProfileGrid.EditMode = 'EditProgrammatically'
    $ProfileGrid.SelectionMode = 'FullRowSelect'
    $ProfileGrid.AutoSizeColumnsMode = 'Fill'
    $ProfileGrid.ColumnHeadersHeight = 30
    $ProfileGrid.RowTemplate.Height = 25
    $ProfileGrid.MultiSelect = $false

    $colProfName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colProfName.HeaderText = "Name"; $colProfName.FillWeight = 35; $colProfName.ReadOnly = $true
    $colProfPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colProfPath.HeaderText = "Path"; $colProfPath.FillWeight = 65; $colProfPath.ReadOnly = $true
    
    $ProfileGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colProfName, $colProfPath))

    $p = @{ type='Button'; width=80; height=25; top=290; left=300; bg=@(40, 40, 40); fg=@(240, 240, 240); id='AddProfile'; text='Add'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Manually add an existing folder as a profile' }
    $btnAddProfile = Set-UIElement @p
    $p = @{ type='Button'; width=80; height=25; top=290; left=390; bg=@(40, 40, 40); fg=@(240, 240, 240); id='RenameProfile'; text='Rename'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Rename selected profile' }
    $btnRenameProfile = Set-UIElement @p
    $p = @{ type='Button'; width=80; height=25; top=290; left=480; bg=@(40, 40, 40); fg=@(240, 240, 240); id='RemoveProfile'; text='Remove'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Remove selected profile from list' }
    $btnRemoveProfile = Set-UIElement @p

    $tabGeneral.Controls.AddRange(@($lblLauncher, $txtLauncher, $btnBrowseLauncher, $lblProcessName, $txtProcessName, $lblMaxClients, $txtMaxClients, $lblJunction, $txtJunction, $btnBrowseJunction, $btnStartJunction, $chkHideMinimizedWindows, $chkNeverRestartingLogin, $chkAutoReconnect, $lblProfiles, $ProfileGrid, $btnAddProfile, $btnRenameProfile, $btnRemoveProfile))

    # --- LOGIN SETTINGS TAB CONTROLS ---
    $Pickers = @{}
    $rowY = 10
    
    $AddPickerRow = {
        param($LabelText, $KeyName, $Top, $Col)
        $leftOffset = if($Col -eq 2) { 300 } else { 10 }
        
        $p = @{ type='Label'; width=100; height=20; top=$Top; left=$leftOffset; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); text=$LabelText; font=(New-Object System.Drawing.Font('Segoe UI', 8)) }
        $l = Set-UIElement @p
        $p = @{ type='TextBox'; width=80; height=20; top=$Top; left=($leftOffset+100); bg=@(40, 40, 40); fg=@(240, 240, 240); id="txt$KeyName"; text='0,0'; font=(New-Object System.Drawing.Font('Segoe UI', 8)) }
        $t = Set-UIElement @p
        $p = @{ type='Button'; width=40; height=20; top=$Top; left=($leftOffset+185); bg=@(60, 60, 100); fg=@(240, 240, 240); id="btnPick$KeyName"; text='Set'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 7)) }
        $b = Set-UIElement @p
        return @{ Label=$l; Text=$t; Button=$b }
    }

    # Left Column: Servers & Channels
    $p = &$AddPickerRow "Server 1:" "Server1" $rowY 1; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Server1"] = $p
    $p = &$AddPickerRow "Server 2:" "Server2" ($rowY+25) 1; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Server2"] = $p
    $p = &$AddPickerRow "Channel 1:" "Channel1" ($rowY+50) 1; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Channel1"] = $p
    $p = &$AddPickerRow "Channel 2:" "Channel2" ($rowY+75) 1; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Channel2"] = $p
    
    # Right Column: Character Slots
    $p = &$AddPickerRow "Char Slot 1:" "Char1" $rowY 2; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Char1"] = $p
    $p = &$AddPickerRow "Char Slot 2:" "Char2" ($rowY+25) 2; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Char2"] = $p
    $p = &$AddPickerRow "Char Slot 3:" "Char3" ($rowY+50) 2; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Char3"] = $p

    $p = &$AddPickerRow "Collector Start:" "CollectorStart" ($rowY+75) 2; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["CollectorStart"] = $p 

    # Disconnect OK
    $p = &$AddPickerRow "Disconnect OK:" "DisconnectOK" ($rowY+110) 2; 
    $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["DisconnectOK"] = $p

    # Post-Login Delay Input
    $p = @{ type='Label'; width=150; height=20; top=($rowY+135); left=10;
    bg=@(30, 30, 30, 0); fg=@(240, 240, 240); text="Post-Login Delay (s):"; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $lblPostLoginDelay = Set-UIElement @p
    $p = @{ type='TextBox'; width=80; height=20; top=($rowY+135); left=160; bg=@(40, 40, 40); fg=@(240, 240, 240); id='txtPostLoginDelayInput'; text='5'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Post login delay in seconds' }
    $txtPostLoginDelay = Set-UIElement @p
    $tabLoginSettings.Controls.AddRange(@($lblPostLoginDelay, $txtPostLoginDelay))

    # 2. Client Configuration Grid
    $gridTop = $rowY + 165
    $gridHeight = 285
    
    $p = @{ type='DataGridView'; width=560; height=$gridHeight; top=$gridTop; left=10; bg=@(40, 40, 40); fg=@(240, 240, 240); id='LoginConfigGrid'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $LoginConfigGrid = Set-UIElement @p
    $LoginConfigGrid.AllowUserToAddRows = $false
    $LoginConfigGrid.RowHeadersVisible = $false
    $LoginConfigGrid.EditMode = 'EditProgrammatically'
    $LoginConfigGrid.SelectionMode = 'CellSelect'
    $LoginConfigGrid.AutoSizeColumnsMode = 'Fill'
    $LoginConfigGrid.ColumnHeadersHeight = 30
    $LoginConfigGrid.RowTemplate.Height = 25
    
    # Define Columns
    $colClient = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colClient.HeaderText = "Client #"; $colClient.FillWeight = 15; $colClient.ReadOnly = $true
    
    $colSrv = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSrv.HeaderText = "Server"; $colSrv.FillWeight = 20; $colSrv.ReadOnly = $true
    
    $colCh = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colCh.HeaderText = "Channel"; $colCh.FillWeight = 20; $colCh.ReadOnly = $true

    $colChar = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colChar.HeaderText = "Character"; $colChar.FillWeight = 25; $colChar.ReadOnly = $true
    
    $colColl = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colColl.HeaderText = "Collecting"; $colColl.FillWeight = 20; $colColl.ReadOnly = $true
    
    $cols = @($colClient, $colSrv, $colCh, $colChar, $colColl)
    $LoginConfigGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$cols)
    
    # Populate Rows (1-10) with defaults
    for ($i=1; $i -le 10; $i++) {
        $LoginConfigGrid.Rows.Add($i, "1", "1", "1", "No") | Out-Null
    }
    
    $tabLoginSettings.Controls.Add($LoginConfigGrid)

    # --- BOTTOM BUTTONS (Outside Tabs) ---
    $p = @{ type='Button'; width=120; height=40; top=500; left=20; bg=@(35, 175, 75); fg=@(240, 240, 240); id='Save'; text='Save'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Save all settings' }
    $btnSave = Set-UIElement @p
    
    $p = @{ type='Button'; width=120; height=40; top=500; left=150; bg=@(210, 45, 45); fg=@(240, 240, 240); id='Cancel'; text='Cancel'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Close and do not save' }
    $btnCancel = Set-UIElement @p
    
    $settingsForm.Controls.AddRange(@($btnSave, $btnCancel))
    #endregion

    #region Step: Set Up Control Hierarchy
    $mainForm.Controls.AddRange(@($topBar, $btnLogin, $btnFtool, $btnLaunch, $btnSettings, $btnStop, $DataGridMain, $DataGridFiller, $GlobalProgressBar))
    $topBar.Controls.AddRange(@($titleLabelForm, $copyrightLabelForm, $btnMinimizeForm, $btnCloseForm))
    
    # Grid Context Menu
    $ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $itmFront = New-Object System.Windows.Forms.ToolStripMenuItem('Show')
    $itmBack = New-Object System.Windows.Forms.ToolStripMenuItem('Minimize')
    $itmResizeCenter = New-Object System.Windows.Forms.ToolStripMenuItem('Resize')
    $itmRelog = New-Object System.Windows.Forms.ToolStripMenuItem('Relog after Disconnect (wip)')
    $ctxMenu.Items.AddRange(@($itmFront, $itmBack, $itmResizeCenter, $itmRelog))
    $DataGridMain.ContextMenuStrip = $ctxMenu
    $DataGridFiller.ContextMenuStrip = $ctxMenu

    #endregion

    #region Step: Global UI Object
    $global:DashboardConfig.UI = [PSCustomObject]@{
        MainForm = $mainForm
        SettingsForm = $settingsForm
        TopBar = $topBar
        CloseForm = $btnCloseForm
        MinForm = $btnMinimizeForm
        DataGridMain = $DataGridMain
        DataGridFiller = $DataGridFiller
        GlobalProgressBar = $GlobalProgressBar
        LoginButton = $btnLogin
        Ftool = $btnFtool
        Settings = $btnSettings
        Exit = $btnStop
        Launch = $btnLaunch
        LaunchContextMenu = $LaunchContextMenu # ADDED: Required for Event Handler
        ToolTip = $toolTipSettings 
        
        # General Tab Inputs
        InputLauncher = $txtLauncher
        InputJunction = $txtJunction
        StartJunction = $btnStartJunction
        InputProcess = $txtProcessName
        InputMax = $txtMaxClients
        Browse = $btnBrowseLauncher
        BrowseJunction = $btnBrowseJunction
        NeverRestartingCollectorLogin = $chkNeverRestartingLogin
        HideMinimizedWindows = $chkHideMinimizedWindows
        AutoReconnect = $chkAutoReconnect
        ProfileGrid = $ProfileGrid
        AddProfile = $btnAddProfile
        RenameProfile = $btnRenameProfile
        RemoveProfile = $btnRemoveProfile
        
        # Login Settings Tab Inputs
        LoginConfigGrid = $LoginConfigGrid
        LoginPickers = $Pickers
        InputPostLoginDelay = $txtPostLoginDelay
        
        Save = $btnSave
        Cancel = $btnCancel
        ContextMenuFront = $itmFront
        ContextMenuBack = $itmBack
        ContextMenuResizeAndCenter = $itmResizeCenter
        Relog = $itmRelog
    }
    
    if ($null -ne $uiPropertiesToAdd) {
        $uiPropertiesToAdd.GetEnumerator() | ForEach-Object {
            $global:DashboardConfig.UI | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value -Force
        }
    }
    #endregion

    Register-UIEventHandlers
    return $true
}


#endregion Function: Initialize-UI

#region Function: Register-UIEventHandlers
function Register-UIEventHandlers
	{
		[CmdletBinding()]
		param()

		if ($null -eq $global:DashboardConfig.UI) { return }

		$eventMappings = @{
			MainForm = @{
				Load = {
							if ($global:DashboardConfig.Paths.Ini)
							{
								$iniExists = Test-Path -Path $global:DashboardConfig.Paths.Ini
								if ($iniExists)
								{
									$iniSettings = Get-IniFileContent -Ini $global:DashboardConfig.Paths.Ini
									if ($iniSettings.Count -gt 0)
									{
										$global:DashboardConfig.Config = $iniSettings
									}
								}
								Sync-ConfigToUI
							}

							$script:initialControlProps = @{}
							$script:initialFormWidth = $global:DashboardConfig.UI.MainForm.Width
							$script:initialFormHeight = $global:DashboardConfig.UI.MainForm.Height

							$controlsToScale = @('TopBar', 'Login', 'Ftool', 'Settings', 'Exit', 'Launch', 'DataGridMain', 'DataGridFiller', 'MinForm', 'CloseForm')

							foreach ($controlName in $controlsToScale)
							{
								$control = $global:DashboardConfig.UI.$controlName
								if ($control)
								{
									$script:initialControlProps[$controlName] = @{
										Left             = $control.Left
										Top              = $control.Top
										Width            = $control.Width
										Height           = $control.Height
										IsScalableBottom = ($controlName -eq 'DataGridFiller' -or $controlName -eq 'DataGridMain') 
									}
								}
							}
						}
					Shown       = {
							if ($global:DashboardConfig.UI.DataGridFiller)
							{
								try { Start-DataGridUpdateTimer } catch {}
							}
						}
					FormClosing = {
							param($src, $e)
                            if (Get-Command Stop-Dashboard -ErrorAction SilentlyContinue) { Stop-Dashboard }
						}
					Resize      = {
							if (-not $script:initialControlProps -or -not $global:DashboardConfig.UI) { return }

							$currentFormWidth = $global:DashboardConfig.UI.MainForm.ClientSize.Width
							$currentFormHeight = $global:DashboardConfig.UI.MainForm.ClientSize.Height
							$scaleW = $currentFormWidth / $script:initialFormWidth

							$fixedTopHeight = 125
							$bottomMargin = 10

							foreach ($controlName in $script:initialControlProps.Keys)
							{
								$control = $global:DashboardConfig.UI.$controlName
								if ($control)
								{
									$initialProps = $script:initialControlProps[$controlName]
									$newLeft = [int]($initialProps.Left * $scaleW)
									$newWidth = [int]($initialProps.Width * $scaleW)

									if ($initialProps.IsScalableBottom)
									{
										$control.Top = $fixedTopHeight
										$control.Height = [Math]::Max(100, $currentFormHeight - $fixedTopHeight - $bottomMargin)
									}
									else
									{
										$control.Top = $initialProps.Top
										$control.Height = $initialProps.Height
									}
									$control.Left = $newLeft
									$control.Width = $newWidth
								}
							}
						}
				}
                # CLICK HANDLER FOR SETTINGS GRID
                LoginConfigGrid = @{
                    CellClick = {
                        param($s, $e)
                        if ($e.RowIndex -lt 0) { return }
                        
                        $row = $s.Rows[$e.RowIndex]
                        $colIndex = $e.ColumnIndex
                        
                        # Col 1: Server
                        if ($colIndex -eq 1) { $row.Cells[1].Value = if ($row.Cells[1].Value -eq "1") { "2" } else { "1" } }
                        # Col 2: Channel
                        elseif ($colIndex -eq 2) { $row.Cells[2].Value = if ($row.Cells[2].Value -eq "1") { "2" } else { "1" } }
                        # Col 3: Char
                        elseif ($colIndex -eq 3) { $row.Cells[3].Value = switch ($row.Cells[3].Value) { "1" {"2"}; "2" {"3"}; "3" {"1"}; default {"1"} } }
                        # Col 4: Collector? (Yes/No)
                        elseif ($colIndex -eq 4) { $row.Cells[4].Value = if ($row.Cells[4].Value -eq "Yes") { "No" } else { "Yes" } }
                    }
                }
				LaunchContextMenu = @{
					Opening = {
						param($sender, $e)
						$sender.Items.Clear()
						
						# Header
						$header = $sender.Items.Add("Quick Launch (1 Client)")
						$header.Enabled = $false
						$sender.Items.Add("-")
						
						# 1. Add Default Profile Option
						$defaultItem = $sender.Items.Add("Default / Selected")
						# FIX: Removed '$true'
						$defaultItem.add_Click({ 
							Start-ClientLaunch -OneClientOnly 
						})
						
						# 2. Add Specific Profiles
						if ($global:DashboardConfig.Config['Profiles']) {
							$profiles = $global:DashboardConfig.Config['Profiles']
							if ($profiles.Count -gt 0) {
								$sender.Items.Add("-")
								foreach ($key in $profiles.Keys) {
									$item = $sender.Items.Add($key)
									$item.Tag = $key 
									# FIX: Removed '$true'
									$item.add_Click({ 
										param($s, $ev)
										$profName = $s.Tag
										Start-ClientLaunch -ProfileNameOverride $profName -OneClientOnly
									})
								}
							}
						}
					}
				}

                SettingsForm = @{ Load = { Sync-ConfigToUI } }
				MinForm = @{ Click = { $global:DashboardConfig.UI.MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized } }
				CloseForm = @{ Click = { $global:DashboardConfig.UI.MainForm.Close() } }
				TopBar = @{ MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.MainForm.Handle, 0xA1, 0x2, 0) } }
				Settings = @{ Click = { Show-SettingsForm } }
				Save = @{ Click = { Sync-UIToConfig; Write-Config; Hide-SettingsForm } }
				Cancel = @{ Click = { Hide-SettingsForm } }
				Browse = @{ Click = { $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Executable Files (*.exe)|*.exe|All Files (*.*)|*.*'; if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $global:DashboardConfig.UI.InputLauncher.Text = $d.FileName } } }
                BrowseJunction = @{ 
                    Click = { 
                        $d = New-Object System.Windows.Forms.OpenFileDialog
                        $d.Title = "Select the destination folder for the Client Copy"
                        $d.ValidateNames = $false
                        $d.CheckFileExists = $false
                        $d.CheckPathExists = $true
                        $d.FileName = "Select Folder"
                        $d.Filter = "Folders|`n"
                        if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
                            $global:DashboardConfig.UI.InputJunction.Text = [System.IO.Path]::GetDirectoryName($d.FileName)
                        } 
                    } 
                }
                StartJunction = @{
                    Click = {
                        $srcExe = $global:DashboardConfig.UI.InputLauncher.Text
                        $baseParentDir = $global:DashboardConfig.UI.InputJunction.Text

                        if ([string]::IsNullOrWhiteSpace($srcExe) -or -not (Test-Path $srcExe)) {
                            [System.Windows.Forms.MessageBox]::Show("Please select a valid Launcher executable first.", "Error", "OK", "Error")
                            return
                        }
                        if ([string]::IsNullOrWhiteSpace($baseParentDir)) {
                            [System.Windows.Forms.MessageBox]::Show("Please select a destination folder.", "Error", "OK", "Error")
                            return
                        }

                        $sourceDir = Split-Path $srcExe -Parent
                        
                        # Calculate folder name from source
                        $folderName = Split-Path $sourceDir -Leaf
                        if ($folderName -in @('bin32', 'bin64')) {
                             # Go up one level to find the actual client folder name
                             $parentOfBin = Split-Path $sourceDir -Parent
                             $folderName = Split-Path $parentOfBin -Leaf
                        }
                        
                        if ([string]::IsNullOrWhiteSpace($folderName)) { $folderName = "Client" }

                        # Check for uniqueness and append number if necessary
                        $baseName = "${folderName}_Copy"
                        $destDir = Join-Path $baseParentDir $baseName
                        
                        $counter = 1
                        while (Test-Path $destDir) {
                            $destDir = Join-Path $baseParentDir "${baseName}_${counter}"
                            $counter++
                        }

                        $confirm = [System.Windows.Forms.MessageBox]::Show("This will create Junctions for 'bin32', 'bin64', 'Data', 'Effect' and copy all other files from the launcher parent directory to:`n$destDir`n`nProceed?", "Confirm Junction & Copy", "YesNo", "Question")
                        
                        if ($confirm -eq 'Yes') {
                            try {
                                if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }

                                $junctionFolders = @('bin32', 'bin64', 'Data', 'Effect')
                                
                                # Disable UI briefly
                                $global:DashboardConfig.UI.SettingsForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

                                Get-ChildItem -Path $sourceDir | ForEach-Object {
                                    $itemName = $_.Name
                                    $sourcePath = $_.FullName
                                    $targetPath = Join-Path $destDir $itemName

                                    if ($itemName -in $junctionFolders) {
                                        # Create Junction (mklink /J)
                                        # Note: mklink is a CMD internal command, requires shell execution
                                        if (Test-Path $targetPath) {
                                            # Skip if exists
                                        } else {
                                            $cmdArgs = "/c mklink /J `"$targetPath`" `"$sourcePath`""
                                            Start-Process -FilePath "cmd.exe" -ArgumentList $cmdArgs -WindowStyle Hidden -Wait
                                        }
                                    } else {
                                        # Normal Copy
                                        Copy-Item -Path $sourcePath -Destination $targetPath -Recurse -Force
                                    }
                                }
                                
                                $global:DashboardConfig.UI.SettingsForm.Cursor = [System.Windows.Forms.Cursors]::Default
                                
                                # ADD TO PROFILE GRID AUTOMATICALLY WITH NAME PROMPT
                                $defaultName = Split-Path $destDir -Leaf
                                $profName = Show-InputBox -Title "Profile Name" -Prompt "Enter a name for the new profile:" -DefaultText $defaultName
                                
                                if ([string]::IsNullOrWhiteSpace($profName)) {
                                    $profName = $defaultName
                                }
                                
                                $global:DashboardConfig.UI.ProfileGrid.Rows.Add($profName, $destDir) | Out-Null
                                [System.Windows.Forms.MessageBox]::Show("Junctions created and Profile '$profName' added successfully.", "Success", "OK", "Information")

                            } catch {
                                $global:DashboardConfig.UI.SettingsForm.Cursor = [System.Windows.Forms.Cursors]::Default
                                [System.Windows.Forms.MessageBox]::Show("An error occurred:`n$($_.Exception.Message)`n`nNote: Creating Junctions may require running the Dashboard as Administrator.", "Error", "OK", "Error")
                            }
                        }
                    }
                }
                AddProfile = @{
                    Click = {
                        $d = New-Object System.Windows.Forms.OpenFileDialog
                        $d.Title = "Select an existing Client folder"
                        $d.ValidateNames = $false
                        $d.CheckFileExists = $false
                        $d.CheckPathExists = $true
                        $d.FileName = "Select Folder"
                        $d.Filter = "Folders|`n"
                        
                        if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                            $path = [System.IO.Path]::GetDirectoryName($d.FileName)
                            $defaultName = Split-Path $path -Leaf
                            
                            $profName = Show-InputBox -Title "Profile Name" -Prompt "Enter a name for this profile:" -DefaultText $defaultName
                            
                            if (-not [string]::IsNullOrWhiteSpace($profName)) {
                                $global:DashboardConfig.UI.ProfileGrid.Rows.Add($profName, $path) | Out-Null
                            }
                        }
                    }
                }
                RenameProfile = @{
                    Click = {
                        if ($global:DashboardConfig.UI.ProfileGrid.SelectedRows.Count -gt 0) {
                            $row = $global:DashboardConfig.UI.ProfileGrid.SelectedRows[0]
                            $oldName = $row.Cells[0].Value
                            
                            $newName = Show-InputBox -Title "Rename Profile" -Prompt "Enter new profile name:" -DefaultText $oldName
                            
                            if (-not [string]::IsNullOrWhiteSpace($newName)) {
                                $row.Cells[0].Value = $newName
                            }
                        } else {
                            [System.Windows.Forms.MessageBox]::Show("Please select a profile to rename.", "Rename Profile", "OK", "Warning")
                        }
                    }
                }
                RemoveProfile = @{
                    Click = {
                        $rows = $global:DashboardConfig.UI.ProfileGrid.SelectedRows
                        foreach ($row in $rows) {
                            $global:DashboardConfig.UI.ProfileGrid.Rows.Remove($row)
                        }
                    }
                }
				DataGridFiller = @{
                    DoubleClick = { param($s,$e); $h=$s.HitTest($e.X,$e.Y); if($h.RowIndex -ge 0 -and $s.Rows[$h.RowIndex].Tag) { [Custom.Native]::BringToFront($s.Rows[$h.RowIndex].Tag.MainWindowHandle) } }
                    MouseDown = {
                        param($s,$e)
                        $h=$s.HitTest($e.X,$e.Y)
                        if ($e.Button -eq 'Right' -and $h.RowIndex -ge 0) {
                            $clickedRow = $s.Rows[$h.RowIndex]
                            # Only clear selection and select the clicked row if it's NOT already part of the selection
                            if (-not $clickedRow.Selected) {
                                $s.ClearSelection()
                                $clickedRow.Selected=$true
                            }
                        }
                    }
                }
                ContextMenuFront = @{ Click = { $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { [Custom.Native]::BringToFront($_.Tag.MainWindowHandle) } } }
                ContextMenuBack = @{ Click = { $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { [Custom.Native]::SendToBack($_.Tag.MainWindowHandle) } } }
                ContextMenuResizeAndCenter = @{ Click = { $scr=[System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea; $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { [Custom.Native]::PositionWindow($_.Tag.MainWindowHandle, [Custom.Native]::TopWindowHandle, [int](($scr.Width-1040)/2), [int](($scr.Height-807)/2), 1040, 807, 0x0010) } } }
                # LAUNCH BUTTON LOGIC
                Launch = @{ 
                    Click = { 
                        if ($global:DashboardConfig.State.LaunchActive) {
                            # Abort logic
                            try { Stop-ClientLaunch } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message) }
                        } else {
                            # Start logic
                            try { Start-ClientLaunch } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message) } 
                        }
                    } 
                }
                
                # LOGIN BUTTON LOGIC
                LoginButton = @{
                    Click = {
                        try {
                            # Dynamically find the command to ensure it's available before calling
                            $loginCommand = Get-Command LoginSelectedRow -ErrorAction Stop
                            
                            # Construct the log file path
                            $logFilePath = Join-Path -Path (Split-Path -Path $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] -Parent) -ChildPath "Log\network_$(Get-Date -f 'yyyyMMdd').log"
                            
                            # Execute the login command
                            & $loginCommand -LogFilePath $logFilePath
                        }
                        catch {
                            # Provide detailed feedback if the command fails
                            $errorMessage = "Login action failed.`n`nCould not find or execute the 'LoginSelectedRow' function. The 'login.psm1' module may have failed to load correctly.`n`nTechnical Details: $($_.Exception.Message)"
                            
                            # Attempt to use the application's standard error dialog
                            try {
                                Show-ErrorDialog -Message $errorMessage
                            }
                            # Fallback to a standard Windows Forms MessageBox if the custom dialog is also unavailable
                            catch {
                                [System.Windows.Forms.MessageBox]::Show($errorMessage, 'Login Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            }
                        }
                    }
                }

            	Ftool = @{ Click = { if(Get-Command FtoolSelectedRow -EA 0){ $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { FtoolSelectedRow $_ } } } }
            Exit = @{ Click = { if([System.Windows.Forms.MessageBox]::Show("Terminate selected?","Confirm",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes'){ $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { Stop-Process -Id $_.Tag.Id -Force -EA 0 } } } }
		}

        # DYNAMIC EVENT REGISTRATION FOR PICKERS
        $pickers = $global:DashboardConfig.UI.LoginPickers
        if ($pickers) {
            foreach ($key in $pickers.Keys) {
                $btn = $pickers[$key].Button
                $txt = $pickers[$key].Text
                
                $action = {
                    param($s, $e)
                    $targetTxt = $s.Tag 
                    $global:DashboardConfig.UI.SettingsForm.Visible = $false
                    [System.Windows.Forms.MessageBox]::Show("1. Focus client.`n2. Hover target.`n3. Wait 3s.`n`nClick OK.", "Picker")
                    Start-Sleep -Seconds 3
                    $cursorPos = [System.Windows.Forms.Cursor]::Position
                    $hWnd = [Custom.Native]::GetForegroundWindow()
                    $rect = New-Object Custom.Native+RECT
                    if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect)) {
                        $relX = $cursorPos.X - $rect.Left
                        $relY = $cursorPos.Y - $rect.Top
                        $targetTxt.Text = "$relX,$relY"
                    } else {
                        $targetTxt.Text = "Error"
                    }
                    $global:DashboardConfig.UI.SettingsForm.Visible = $true
                    $global:DashboardConfig.UI.SettingsForm.BringToFront()
                }
                
                $btn.Tag = $txt
                $sourceIdentifier = "EntropiaDashboard.Picker.$key"
                Get-EventSubscriber -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue | Unregister-Event
                Register-ObjectEvent -InputObject $btn -EventName Click -Action $action -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue
            }
        }

		foreach ($elementName in $eventMappings.Keys)
		{
			$element = $global:DashboardConfig.UI.$elementName
			if ($element)
			{
				foreach ($e in $eventMappings[$elementName].Keys)
				{
					$sourceIdentifier = "EntropiaDashboard.$elementName.$e"
					Get-EventSubscriber -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue | Unregister-Event
					Register-ObjectEvent -InputObject $element -EventName $e -Action $eventMappings[$elementName][$e] -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue
				}
			}
		}

		$global:DashboardConfig.State.UIInitialized = $true
	}



#endregion Function: Register-UIEventHandlers

#region Function: Show-SettingsForm
	function Show-SettingsForm
	{
		<#
		.SYNOPSIS
			Shows the settings form with a fade-in animation effect.
		#>
		[CmdletBinding()]
		param()

		#region Step: Prevent Concurrent Animations
			if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
				($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
			{
				return # Exit if an animation is active
			}
		#endregion Step: Prevent Concurrent Animations

		#region Step: Validate UI Objects
			if (-not ($global:DashboardConfig.UI -and $global:DashboardConfig.UI.SettingsForm -and $global:DashboardConfig.UI.MainForm))
			{
				return
			}
		#endregion Step: Validate UI Objects

		$settingsForm = $global:DashboardConfig.UI.SettingsForm

		#region Step: Position and Show Settings Form
			if ($settingsForm.Opacity -lt 0.95)
			{
				$settingsForm.Visible = $true
				$mainFormLocation = $global:DashboardConfig.UI.MainForm.Location
				$settingsFormWidth = $settingsForm.Width
				$settingsFormHeight = $settingsForm.Height
				$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
				$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

				$x = $mainFormLocation.X + (($global:DashboardConfig.UI.MainForm.Width - $settingsFormWidth) / 2)
				$y = $mainFormLocation.Y + (($global:DashboardConfig.UI.MainForm.Height - $settingsFormHeight) / 2)

				$margin = 0
				$x = [Math]::Max($margin, [Math]::Min($x, $screenWidth - $settingsFormWidth - $margin))
				$y = [Math]::Max($margin, [Math]::Min($y, $screenHeight - $settingsFormHeight - $margin))

				$settingsForm.Location = New-Object System.Drawing.Point($x, $y)
				$settingsForm.BringToFront()
				$settingsForm.Activate()
			}
		#endregion Step: Position and Show Settings Form

		#region Step: Create and Start Fade-In Animation Timer
			if ($script:fadeInTimer) { $script:fadeInTimer.Dispose() }
			$script:fadeInTimer = New-Object System.Windows.Forms.Timer
			$script:fadeInTimer.Interval = 15 
			$script:fadeInTimer.Add_Tick({
					if (-not $global:DashboardConfig.UI.SettingsForm -or $global:DashboardConfig.UI.SettingsForm.IsDisposed)
					{
						$script:fadeInTimer.Stop()
						$script:fadeInTimer.Dispose()
						$script:fadeInTimer = $null
						return
					}
					if ($global:DashboardConfig.UI.SettingsForm.Opacity -lt 1)
					{
						$global:DashboardConfig.UI.SettingsForm.Opacity += 0.1
					}
					else
					{
						$global:DashboardConfig.UI.SettingsForm.Opacity = 1 
						$script:fadeInTimer.Stop()
						$script:fadeInTimer.Dispose()
						$script:fadeInTimer = $null
					}
				})
			$script:fadeInTimer.Start()
			$global:DashboardConfig.Resources.Timers['fadeInTimer'] = $script:fadeInTimer
		#endregion Step: Create and Start Fade-In Animation Timer
	}
#endregion Function: Show-SettingsForm

#region Function: Hide-SettingsForm
	function Hide-SettingsForm
	{
		<#
		.SYNOPSIS
			Hides the settings form with a fade-out animation effect.
		#>
		[CmdletBinding()]
		param()

		#region Step: Prevent Concurrent Animations
			if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
				($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
			{
				return 
			}
		#endregion Step: Prevent Concurrent Animations

		#region Step: Validate UI Objects
			if (-not ($global:DashboardConfig.UI -and $global:DashboardConfig.UI.SettingsForm))
			{
				return
			}
		#endregion Step: Validate UI Objects

		#region Step: Create and Start Fade-Out Animation Timer
			if ($global:fadeOutTimer) { $global:fadeOutTimer.Dispose() }
			$global:fadeOutTimer = New-Object System.Windows.Forms.Timer
			$global:fadeOutTimer.Interval = 15
			$global:fadeOutTimer.Add_Tick({
					if (-not $global:DashboardConfig.UI.SettingsForm -or $global:DashboardConfig.UI.SettingsForm.IsDisposed)
					{
						$global:fadeOutTimer.Stop()
						$global:fadeOutTimer.Dispose()
						$global:fadeOutTimer = $null
						return
					}

					if ($global:DashboardConfig.UI.SettingsForm.Opacity -gt 0)
					{
						$global:DashboardConfig.UI.SettingsForm.Opacity -= 0.1
					}
					else
					{
						$global:DashboardConfig.UI.SettingsForm.Opacity = 0 
						$global:fadeOutTimer.Stop()
						$global:fadeOutTimer.Dispose()
						$global:fadeOutTimer = $null
						$global:DashboardConfig.UI.SettingsForm.Hide()
					}
				})
			$global:fadeOutTimer.Start()
			$global:DashboardConfig.Resources.Timers['fadeOutTimer'] = $global:fadeOutTimer
		#endregion Step: Create and Start Fade-Out Animation Timer
	}
#endregion Function: Hide-SettingsForm

#region Function: Set-UIElement
function Set-UIElement
	{
		[CmdletBinding()]
		param(
			[Parameter(Mandatory=$true)]
			[ValidateSet('Form', 'Panel', 'Button', 'Label', 'DataGridView', 'TextBox', 'ComboBox', 'CheckBox', 'Toggle', 'ProgressBar', 'TextProgressBar')]
			[string]$type,
			[bool]$visible,
			[int]$width,
			[int]$height,
			[int]$top,
			[int]$left,
			[array]$bg,
			[array]$fg,
			[string]$id,
			[string]$text,
			[System.Windows.Forms.FlatStyle]$fs,
			[System.Drawing.Font]$font,
			[string]$startPosition,
			[int]$formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None,
			[double]$opacity = 1.0,
			[bool]$topMost,
            [bool]$checked,
			[switch]$multiline,
			[switch]$readOnly,
			[switch]$scrollBars,
			[ValidateSet('Simple', 'DropDown', 'DropDownList')]
			[string]$dropDownStyle = 'DropDownList',
            [ValidateSet('Blocks', 'Continuous', 'Marquee')]
            [string]$style = 'Continuous',
            [string]$tooltip
		)

		#region Step: Create UI Element Based on Type
			$el = switch ($type)
			{
				'Form'            { New-Object System.Windows.Forms.Form }
				'Panel'           { New-Object System.Windows.Forms.Panel }
				'Button'          { New-Object System.Windows.Forms.Button }
				'Label'           { New-Object System.Windows.Forms.Label }
				'DataGridView'    { New-Object System.Windows.Forms.DataGridView }
				'TextBox'         { New-Object System.Windows.Forms.TextBox }
				'ComboBox'        { New-Object System.Windows.Forms.ComboBox }
				'CheckBox'        { New-Object System.Windows.Forms.CheckBox }
				'Toggle'          { New-Object Custom.Toggle }
                'ProgressBar'     { New-Object System.Windows.Forms.ProgressBar }
                'TextProgressBar' { New-Object Custom.TextProgressBar }
				default           { throw "Invalid element type specified: $type" }
			}
		#endregion Step: Create UI Element Based on Type

		#region Step: Configure DataGridView Specific Properties
			if ($type -eq 'DataGridView')
			{
				$el.AllowUserToAddRows = $false
				$el.ReadOnly = $false
				$el.AllowUserToOrderColumns = $true
				$el.AllowUserToResizeColumns  = $false
				$el.AllowUserToResizeRows = $false
				$el.RowHeadersVisible = $false
				$el.MultiSelect = $true
				$el.SelectionMode = 'FullRowSelect'
				$el.AutoSizeColumnsMode = 'Fill'
				$el.BorderStyle = 'FixedSingle'
				$el.EnableHeadersVisualStyles = $false
				$el.CellBorderStyle = 'SingleHorizontal'
				$el.ColumnHeadersBorderStyle = 'Single'
				$el.EditMode = 'EditProgrammatically'
				$el.ColumnHeadersHeightSizeMode = 'DisableResizing'
				$el.RowHeadersWidthSizeMode = 'DisableResizing'
				$el.DefaultCellStyle.Alignment = 'MiddleCenter'
				$el.ColumnHeadersDefaultCellStyle.Alignment = 'MiddleCenter'

				$el.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
				$el.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 37)
				$el.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
				$el.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
				$el.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
				$el.GridColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
				$el.BackgroundColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
				$el.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(60, 80, 180)
				$el.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
			}
		#endregion

		#region Step: Apply Common Control Properties
			if ($el -is [System.Windows.Forms.Control])
			{
				if ($PSBoundParameters.ContainsKey('visible')) { $el.Visible = $visible }
				if ($PSBoundParameters.ContainsKey('width'))   { $el.Width = $width }
				if ($PSBoundParameters.ContainsKey('height'))  { $el.Height = $height }
				if ($PSBoundParameters.ContainsKey('top'))     { $el.Top = $top }
				if ($PSBoundParameters.ContainsKey('left'))    { $el.Left = $left }

				if ($bg -is [array] -and $bg.Count -ge 3)
				{
					$el.BackColor = if ($bg.Count -eq 4) { [System.Drawing.Color]::FromArgb($bg[0], $bg[1], $bg[2], $bg[3]) }
									else                 { [System.Drawing.Color]::FromArgb($bg[0], $bg[1], $bg[2]) }
				}

				if ($fg -is [array] -and $fg.Count -ge 3)
				{
					$el.ForeColor = [System.Drawing.Color]::FromArgb($fg[0], $fg[1], $fg[2])
				}

				if ($PSBoundParameters.ContainsKey('font')) { $el.Font = $font }
			}
		#endregion

		#region Step: Apply Type-Specific Properties
			switch ($type)
			{
				'Form' {
					if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
					if ($PSBoundParameters.ContainsKey('startPosition')) { try { $el.StartPosition = [System.Windows.Forms.FormStartPosition]::$startPosition } catch {} }
					if ($PSBoundParameters.ContainsKey('formBorderStyle')) { try { $el.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]$formBorderStyle } catch {} }
					if ($PSBoundParameters.ContainsKey('opacity')) { $el.Opacity = [double]$opacity }
					if ($PSBoundParameters.ContainsKey('topMost')) { $el.TopMost = $topMost }
					if ($PSBoundParameters.ContainsKey('icon')) { $el.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSCommandPath) }
				}
				'Button' {
					if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
					if ($PSBoundParameters.ContainsKey('fs')) {
						$el.FlatStyle = $fs
						$el.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
						$el.FlatAppearance.BorderSize = 1
						$el.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
						$el.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)

						$el.Add_Paint({
							param($src, $e)
							if ($src.FlatStyle -eq [System.Windows.Forms.FlatStyle]::Flat)
							{
								$bgBrush = [System.Drawing.SolidBrush]::new($src.BackColor)
								$e.Graphics.FillRectangle($bgBrush, 0, 0, $src.Width, $src.Height)
								$textBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(240, 240, 240))
								$textFormat = [System.Drawing.StringFormat]::new()
								$textFormat.Alignment = [System.Drawing.StringAlignment]::Center
								$textFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
								$e.Graphics.DrawString($src.Text, $src.Font, $textBrush, [System.Drawing.RectangleF]::new(0, 0, $src.Width, $src.Height), $textFormat)
								$borderPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(60, 60, 60))
								$e.Graphics.DrawRectangle($borderPen, 0, 0, $src.Width, $src.Height)
								$bgBrush.Dispose(); $textBrush.Dispose(); $borderPen.Dispose(); $textFormat.Dispose()
							}
						})
					}
				}
				'Label' {
					if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
				}
				'TextBox' {
					if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
					if ($PSBoundParameters.ContainsKey('multiline')) { $el.Multiline = $multiline }
					if ($PSBoundParameters.ContainsKey('readOnly')) { $el.ReadOnly = $readOnly }
					if ($PSBoundParameters.ContainsKey('scrollBars')) { $el.ScrollBars = if ($scrollBars -and $multiline) { [System.Windows.Forms.ScrollBars]::Vertical } else { [System.Windows.Forms.ScrollBars]::None } }
					$el.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
					$el.TextAlign = "Center"
					$el.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
					$el.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
				}
				'ComboBox' {
					if ($null -ne $dropDownStyle) { try { $el.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::$dropDownStyle } catch {} }
					if ($null -ne $fs) {
						$el.FlatStyle = $fs; $el.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed; $el.IntegralHeight = $false
						$customComboBox = New-Object Custom.DarkComboBox
						$customComboBox.Location = $el.Location; $customComboBox.Size = $el.Size; $customComboBox.Width = $el.Width - 20
						$customComboBox.DropDownStyle = $el.DropDownStyle; $customComboBox.FlatStyle = $el.FlatStyle
						$customComboBox.DrawMode = $el.DrawMode; $customComboBox.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
						$customComboBox.ForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240); $customComboBox.Font = $el.Font
						$customComboBox.IntegralHeight = $false; $customComboBox.TabIndex = $el.TabIndex; $customComboBox.Name = $el.Name
						foreach ($item in $el.Items) { $customComboBox.Items.Add($item) }
						$el = $customComboBox
					}
				}
                'CheckBox' {
                    if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
                    if ($PSBoundParameters.ContainsKey('checked')) { $el.Checked = $checked }
                    $el.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $el.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60, 60, 60); $el.FlatAppearance.BorderSize = 1
                    $el.FlatAppearance.CheckedBackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
                    $el.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
                    $el.UseVisualStyleBackColor = $false; $el.CheckAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $el.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft; $el.Padding = [System.Windows.Forms.Padding]::new(20, 0, 0, 0)
                }
                'Toggle' {
                    if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
                    if ($PSBoundParameters.ContainsKey('checked')) { $el.Checked = $checked }
                }
                'ProgressBar' {
                    if ($PSBoundParameters.ContainsKey('style')) { $el.Style = $style }
                }
                'TextProgressBar' {
                    if ($PSBoundParameters.ContainsKey('style')) { $el.Style = $style }
                }
			}
		#endregion

        # If a tooltip was provided, apply it using the global ToolTip object
        if ($PSBoundParameters.ContainsKey('tooltip') -and $tooltip -ne $null -and $global:DashboardConfig.UI.ToolTip) {
            $global:DashboardConfig.UI.ToolTip.SetToolTip($el, $tooltip)
        }

		return $el
	}

#endregion Function: Set-UIElement


#endregion Core UI Functions

#region Module Exports
#region Step: Export Public Functions
	# Export the functions intended for use by other modules or the main script.
	Export-ModuleMember -Function Initialize-UI, Set-UIElement, Show-SettingsForm, Hide-SettingsForm, Sync-ConfigToUI, Sync-UIToConfig, Register-UIEventHandlers, Show-InputBox
#endregion Step: Export Public Functions
#endregion Module Exports