<# ui.psm1
    .SYNOPSIS
        Constructs, manages, and defines the logic for the entire graphical user interface (GUI) of the Entropia Dashboard except for the Ftool. 
		It uses Windows Forms and custom-drawn controls to create a dark-themed, interactive frontend for the application's features.

    .DESCRIPTION
        This PowerShell module is exclusively responsible for the application's user interface. It dynamically builds all visual components, registers event handlers for user interactions, and manages the synchronization of settings between the UI and the application's configuration state. The UI is built using Windows Forms, but heavily customized with owner-drawn controls from `classes.psm1` to create a consistent dark-mode aesthetic.

        The module's architecture and key functions include:

        1.  **UI Construction (`Initialize-UI` & `Set-UIElement`):**
            *   **`Initialize-UI`:** The main function that orchestrates the creation of all windows and controls. It builds the `MainForm` (the primary dashboard) and a `SettingsForm` for configuration.
            *   **`Set-UIElement`:** A centralized factory function that creates and applies a standard dark theme to all UI components, from basic buttons and labels to complex `DataGridViews` and the custom controls (`DarkTabControl`, `Toggle`, etc.). This ensures a consistent look and feel.
            *   All created UI elements are stored in the `$global:DashboardConfig.UI` object for easy access from other modules.

        2.  **Event and Interaction Logic (`Register-UIEventHandlers`):**
            *   This function brings the UI to life by attaching script blocks to user events (e.g., clicks, double-clicks, form loading).
            *   It wires up the main dashboard buttons (`Launch`, `Login`, `Ftool`, `Terminate`) to trigger core application logic defined in other modules (e.g., `launch.psm1`, `login.psm1`).
            *   It implements the custom draggable title bar and the logic for the coordinate picker tools in the settings menu.

        3.  **Configuration and Data Binding (`Sync-UIToConfig` & `Sync-ConfigToUI`):**
            *   These two functions provide two-way data binding between the UI controls and the in-memory configuration (`$global:DashboardConfig.Config`).
            *   `Sync-ConfigToUI` populates the form with saved settings when it is displayed.
            *   `Sync-UIToConfig` gathers the user's changes from the form fields before they are saved to the `config.ini` file.

        4.  **Animated Transitions (`Show-SettingsForm` & `Hide-SettingsForm`):**
            *   These functions enhance the user experience by providing a simple fade-in/fade-out animation for the settings dialog, which is achieved by manipulating the form's opacity with a timer.
#>

#region Helper Functions
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
        @('LauncherPath', 'ProcessName', 'MaxClients', 'Login', 'LoginConfig') | ForEach-Object {
            if (-not $global:DashboardConfig.Config.Contains($_)) {
                $global:DashboardConfig.Config[$_] = [ordered]@{}
            }
        }

        # General
        $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] = $UI.InputLauncher.Text
        $global:DashboardConfig.Config['ProcessName']['ProcessName'] = $UI.InputProcess.Text
        $global:DashboardConfig.Config['MaxClients']['MaxClients'] = $UI.InputMax.Text
        
        # Checkboxes
        $global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin'] = if ($UI.NeverRestartingCollectorLogin.Checked) { '1' } else { '0' }
        $global:DashboardConfig.Config['Options']['HideMinimizedWindows'] = if ($UI.HideMinimizedWindows.Checked) { '0' } else { '0' }

        # Login Config (Coords)
        foreach ($key in $UI.LoginPickers.Keys) {
            $global:DashboardConfig.Config['LoginConfig']["${key}Coords"] = $UI.LoginPickers[$key].Text.Text
        }
        $global:DashboardConfig.Config['LoginConfig']['PostLoginDelay'] = $UI.InputPostLoginDelay.Text
        
        # Login Config (Grid) - CRITICAL FIX: Ensure values are strings
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

        # Checkboxes
        if ($global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin']) { $UI.NeverRestartingCollectorLogin.Checked = ([int]$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin']) -eq 1 }
        if ($global:DashboardConfig.Config['Options']['HideMinimizedWindows']) { $UI.HideMinimizedWindows.Checked = ([int]$global:DashboardConfig.Config['Options']['HideMinimizedWindows']) -eq 1 }

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
            
            # Grid - CRITICAL FIX: Robust parsing
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

    #region Step: Create Main UI Elements
    $p = @{ type='Form'; visible=$false; width=470; height=440; bg=@(30, 30, 30); id='MainForm'; text='Entropia Dashboard'; startPosition='CenterScreen'; formBorderStyle=[System.Windows.Forms.FormBorderStyle]::None }
    $mainForm = Set-UIElement @p

    $p = @{ type='Form'; visible=$false; width=600; height=550; bg=@(30, 30, 30); id='SettingsForm'; text='Settings'; startPosition='Manual'; formBorderStyle=[System.Windows.Forms.FormBorderStyle]::None; opacity=0; topMost=$true }
    $settingsForm = Set-UIElement @p

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
    $p = @{ type='Label'; width=140; height=10; top=16; left=10; fg=@(230, 230, 230); id='CopyrightLabel'; text=[char]0x00A9 + ' Immortal / Divine 2025 - v1.3'; font=(New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Italic)) }
    $copyrightLabelForm = Set-UIElement @p
    $p = @{ type='Button'; width=30; height=30; left=410; bg=@(40, 40, 40); fg=@(240, 240, 240); id='MinForm'; text='_'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)) }
    $btnMinimizeForm = Set-UIElement @p
    $p = @{ type='Button'; width=30; height=30; left=440; bg=@(150, 20, 20); fg=@(240, 240, 240); id='CloseForm'; text=[char]0x166D; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)) }
    $btnCloseForm = Set-UIElement @p
    #endregion

    #region Step: Main Form Buttons & Grids
    $p = @{ type='Button'; width=125; height=30; top=40; left=15; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Launch'; text='Launch'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $btnLaunch = Set-UIElement @p
    $p = @{ type='Button'; width=125; height=30; top=40; left=150; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Login'; text='Login'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $btnLogin = Set-UIElement @p
    $p = @{ type='Button'; width=80; height=30; top=40; left=285; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Settings'; text='Settings'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $btnSettings = Set-UIElement @p
    $p = @{ type='Button'; width=80; height=30; top=40; left=375; bg=@(150, 20, 20); fg=@(240, 240, 240); id='Terminate'; text='Terminate'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $btnStop = Set-UIElement @p
    $p = @{ type='Button'; width=440; height=30; top=75; left=15; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Ftool'; text='Ftool'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
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
    # USE CUSTOM DARK TAB CONTROL
    $settingsTabs = New-Object Custom.DarkTabControl
    $settingsTabs.Dock = 'Top'
    $settingsTabs.Height = 480
    # Custom class handles appearance
    
    $tabGeneral = New-Object System.Windows.Forms.TabPage "General"
    $tabGeneral.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    
    $tabLoginSettings = New-Object System.Windows.Forms.TabPage "Login Settings"
    $tabLoginSettings.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $tabLoginSettings.AutoScroll = $true 
    
    $settingsTabs.TabPages.Add($tabGeneral)
    $settingsTabs.TabPages.Add($tabLoginSettings)
    $settingsForm.Controls.Add($settingsTabs)

    # --- GENERAL TAB CONTROLS ---
    $p = @{ type='Label'; width=85; height=20; top=25; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelLauncher'; text='Launcher Path:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $lblLauncher = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=50; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputLauncher'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $txtLauncher = Set-UIElement @p
    $p = @{ type='Button'; width=55; height=25; top=20; left=110; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Browse'; text='Browse'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $btnBrowseLauncher = Set-UIElement @p
    
    $p = @{ type='Label'; width=85; height=20; top=95; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelProcess'; text='Process Name:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $lblProcessName = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=120; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputProcess'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $txtProcessName = Set-UIElement @p
    
    $p = @{ type='Label'; width=85; height=20; top=165; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelMax'; text='Max Clients:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $lblMaxClients = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=190; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputMax'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $txtMaxClients = Set-UIElement @p

    $p = @{ type='CheckBox'; width=0; height=0; top=230; left=20; bg=@(30, 30, 30); fg=@(240, 240, 240); id='HideMinimizedWindows'; text='Hide minimized clients from Taskbar and Tab View'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $chkHideMinimizedWindows = Set-UIElement @p
    $p = @{ type='CheckBox'; width=200; height=20; top=255; left=20; bg=@(30, 30, 30); fg=@(240, 240, 240); id='NeverRestartingCollectorLogin'; text='Collector Double Click'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $chkNeverRestartingLogin = Set-UIElement @p
    
    $tabGeneral.Controls.AddRange(@($lblLauncher, $txtLauncher, $btnBrowseLauncher, $lblProcessName, $txtProcessName, $lblMaxClients, $txtMaxClients, $chkHideMinimizedWindows, $chkNeverRestartingLogin))

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

    # Collector Start (Right Column)
    $p = &$AddPickerRow "Collector Start:" "CollectorStart" ($rowY+75) 2; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["CollectorStart"] = $p
    
    # Post-Login Delay Input
    $p = @{ type='Label'; width=150; height=20; top=($rowY+110); left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); text="Post-Login Delay (s):"; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $lblDelay = Set-UIElement @p
    $p = @{ type='TextBox'; width=60; height=25; top=($rowY+110); left=180; bg=@(40, 40, 40); fg=@(240, 240, 240); id="InputPostLoginDelay"; text='1'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $txtDelay = Set-UIElement @p
    $tabLoginSettings.Controls.AddRange(@($lblDelay, $txtDelay))

    # 2. Client Configuration Grid
    $gridTop = $rowY + 140
    $gridHeight = 290
    
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
    $p = @{ type='Button'; width=120; height=40; top=480; left=20; bg=@(35, 175, 75); fg=@(240, 240, 240); id='Save'; text='Save'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $btnSave = Set-UIElement @p
    
    $p = @{ type='Button'; width=120; height=40; top=480; left=150; bg=@(210, 45, 45); fg=@(240, 240, 240); id='Cancel'; text='Cancel'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
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
    $ctxMenu.Items.AddRange(@($itmFront, $itmBack, $itmResizeCenter))
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
        
        # General Tab Inputs
        InputLauncher = $txtLauncher
        InputProcess = $txtProcessName
        InputMax = $txtMaxClients
        Browse = $btnBrowseLauncher
        NeverRestartingCollectorLogin = $chkNeverRestartingLogin
        HideMinimizedWindows = $chkHideMinimizedWindows
        
        # Login Settings Tab Inputs
        LoginConfigGrid = $LoginConfigGrid
        LoginPickers = $Pickers
        InputPostLoginDelay = $txtDelay
        
        Save = $btnSave
        Cancel = $btnCancel
        ContextMenuFront = $itmFront
        ContextMenuBack = $itmBack
        ContextMenuResizeAndCenter = $itmResizeCenter
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
                            # Restore window styles before closing to prevent orphaned windows
                            if (Get-Command Restore-WindowStyles -ErrorAction SilentlyContinue) { Restore-WindowStyles }
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

                SettingsForm = @{ Load = { Sync-ConfigToUI } }
				MinForm = @{ Click = { $global:DashboardConfig.UI.MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized } }
				CloseForm = @{ Click = { $global:DashboardConfig.UI.MainForm.Close() } }
				TopBar = @{ MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.MainForm.Handle, 0xA1, 0x2, 0) } }
				Settings = @{ Click = { Show-SettingsForm } }
				Save = @{ Click = { Sync-UIToConfig; Write-Config; Hide-SettingsForm } }
				Cancel = @{ Click = { Hide-SettingsForm } }
				Browse = @{ Click = { $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Executable Files (*.exe)|*.exe|All Files (*.*)|*.*'; if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $global:DashboardConfig.UI.InputLauncher.Text = $d.FileName } } }
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
                            $logFilePath = ($global:DashboardConfig.Config['LauncherPath']['LauncherPath'] -replace '\\Launcher\.exe$','') + "\Log\network_$(Get-Date -f 'yyyyMMdd').log"
                            
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
            [string]$style = 'Continuous'
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

		return $el
	}

#endregion Function: Set-UIElement

#endregion Core UI Functions

#region Module Exports
#region Step: Export Public Functions
	# Export the functions intended for use by other modules or the main script.
	Export-ModuleMember -Function Initialize-UI, Set-UIElement, Show-SettingsForm, Hide-SettingsForm, Sync-ConfigToUI, Sync-UIToConfig, Register-UIEventHandlers
#endregion Step: Export Public Functions
#endregion Module Exports