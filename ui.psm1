<# ui.psm1
#>

#region Helper Functions

function Get-NeuzResolution {
    param([string]$ProfilePath)
    if ([string]::IsNullOrWhiteSpace($ProfilePath)) { return $null }
    $neuzIni = Join-Path $ProfilePath "neuz.ini"
    if (Test-Path $neuzIni) {
        try {
            $content = Get-Content $neuzIni -Raw
            if ($content -match '(?m)^\s*resolution\s+(\d+)\s+(\d+)') {
                return "$($Matches[1]) x $($Matches[2])"
            }
        } catch {}
    }
    return $null
}

function Set-NeuzResolution {
    param([string]$ProfilePath, [string]$ResolutionString)
    if ([string]::IsNullOrWhiteSpace($ProfilePath) -or [string]::IsNullOrWhiteSpace($ResolutionString)) { return }
    $neuzIni = Join-Path $ProfilePath "neuz.ini"
    if (-not (Test-Path $neuzIni)) { return }

    try {
        $parts = $ResolutionString -split ' x '
        if ($parts.Count -eq 2) {
            $w = $parts[0].Trim()
            $h = $parts[1].Trim()
            $content = Get-Content $neuzIni -Raw
            if ($content -match '(?m)^\s*resolution\s+\d+\s+\d+') {
                $content = $content -replace '(?m)(^\s*resolution\s+)\d+\s+\d+', "`${1}$w $h"
            } else {
                $content += "`r`nresolution $w $h"
            }
            Set-Content -Path $neuzIni -Value $content -Encoding ASCII -Force
        }
    } catch {
        Write-Verbose "Failed to set resolution in $neuzIni $_"
    }
}

function Sync-ProfilesToConfig
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try
    {
        Write-Verbose '  UI: Syncing ProfileGrid to config profiles...' -ForegroundColor Cyan
        $UI = $global:DashboardConfig.UI
        if (-not ($UI -and $UI.ProfileGrid -and $global:DashboardConfig.Config)) { return $false }

        if (-not $global:DashboardConfig.Config.Contains('Profiles') -or $null -eq $global:DashboardConfig.Config['Profiles']) {
            $global:DashboardConfig.Config['Profiles'] = [ordered]@{}
        }
        if (-not $global:DashboardConfig.Config.Contains('ReconnectProfiles') -or $null -eq $global:DashboardConfig.Config['ReconnectProfiles']) {
            $global:DashboardConfig.Config['ReconnectProfiles'] = @{}
        }
        if (-not $global:DashboardConfig.Config.Contains('HideProfiles') -or $null -eq $global:DashboardConfig.Config['HideProfiles']) {
            $global:DashboardConfig.Config['HideProfiles'] = @{}
        }

        $global:DashboardConfig.Config['Profiles'].Clear()
        $global:DashboardConfig.Config['ReconnectProfiles'].Clear()
        $global:DashboardConfig.Config['HideProfiles'].Clear()

        foreach ($row in $UI.ProfileGrid.Rows) {
            if ($row.Cells[0].Value -and $row.Cells[1].Value) {
                $key = $row.Cells[0].Value.ToString()
                $val = $row.Cells[1].Value.ToString()

                $global:DashboardConfig.Config['Profiles'][$key] = $val

                if ($row.Cells[2].Value -eq $true -or $row.Cells[2].Value -eq 1) {
                    $global:DashboardConfig.Config['ReconnectProfiles'][$key] = "1"
                }

                if ($row.Cells[3].Value -eq $true -or $row.Cells[3].Value -eq 1) {
                    $global:DashboardConfig.Config['HideProfiles'][$key] = "1"
                }
            }
        }

        Write-Verbose '  UI: ProfileGrid synced to config profiles.' -ForegroundColor Green
        return $true
    }
    catch
    {
        Write-Verbose "  UI: Failed to sync ProfileGrid to config profiles: $_" -ForegroundColor Red
        return $false
    }
}

function RefreshLoginProfileSelector
{
    [CmdletBinding()]
    param()

    try
    {
        Write-Verbose '  UI: Refreshing LoginProfileSelector...' -ForegroundColor Cyan
        $UI = $global:DashboardConfig.UI
        if (-not $UI -or -not $UI.LoginProfileSelector) { return }

        $selectedItem = $UI.LoginProfileSelector.SelectedItem
        $UI.LoginProfileSelector.Items.Clear()
        $UI.LoginProfileSelector.Items.Add("Default") | Out-Null

        if ($global:DashboardConfig.Config['Profiles']) {
            foreach ($key in $global:DashboardConfig.Config['Profiles'].Keys) {
                $UI.LoginProfileSelector.Items.Add($key) | Out-Null
            }
        }

        if ($selectedItem -and $UI.LoginProfileSelector.Items.Contains($selectedItem)) {
            $UI.LoginProfileSelector.SelectedItem = $selectedItem
        } else {
            $UI.LoginProfileSelector.SelectedIndex = 0
        }

        Write-Verbose '  UI: LoginProfileSelector refreshed.' -ForegroundColor Green
    }
    catch
    {
        Write-Verbose "  UI: Failed to refresh LoginProfileSelector: $_" -ForegroundColor Red
    }
}

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
    $form.TopMost = $true

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

function Sync-UIToConfig
{
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$ProfileToSave = $null
    )

    try
    {
        Write-Verbose '  UI: Syncing UI to config' -ForegroundColor Cyan
        $UI = $global:DashboardConfig.UI
        if (-not ($UI -and $global:DashboardConfig.Config)) { return $false }

        @('LauncherPath', 'ProcessName', 'MaxClients', 'Login', 'LoginConfig', 'Options', 'Paths', 'Profiles', 'ReconnectProfiles', 'HideProfiles') |
        ForEach-Object {
            if (-not $global:DashboardConfig.Config.Contains($_)) {
                $global:DashboardConfig.Config[$_] = [ordered]@{}
            }
        }

        $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] = $UI.InputLauncher.Text
        $global:DashboardConfig.Config['ProcessName']['ProcessName'] = $UI.InputProcess.Text
        $global:DashboardConfig.Config['MaxClients']['MaxClients'] = $UI.InputMax.Text
        $global:DashboardConfig.Config['Paths']['JunctionTarget'] = $UI.InputJunction.Text

        $global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin'] = if ($UI.NeverRestartingCollectorLogin.Checked) { '1' } else { '0' }

        $global:DashboardConfig.Config['Profiles'] = [ordered]@{}
        $global:DashboardConfig.Config['ReconnectProfiles'] = @{}
        $global:DashboardConfig.Config['HideProfiles'] = @{}

        foreach ($row in $UI.ProfileGrid.Rows) {
            if ($row.Cells[0].Value -and $row.Cells[1].Value) {
                $key = $row.Cells[0].Value.ToString()
                $val = $row.Cells[1].Value.ToString()
                $global:DashboardConfig.Config['Profiles'][$key] = $val

                if ($row.Cells[2].Value -eq $true -or $row.Cells[2].Value -eq 1) {
                    $global:DashboardConfig.Config['ReconnectProfiles'][$key] = "1"
                }

                if ($row.Cells[3].Value -eq $true -or $row.Cells[3].Value -eq 1) {
                    $global:DashboardConfig.Config['HideProfiles'][$key] = "1"
                }
            }
        }

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

        $targetProfile = $ProfileToSave

        if (-not $targetProfile) {
            $targetProfile = $UI.LoginProfileSelector.SelectedItem
            if (-not $targetProfile) { $targetProfile = "Default" }
        }

        $profPath = $null
        if ($targetProfile -eq 'Default') {
             if ($global:DashboardConfig.Config['LauncherPath']['LauncherPath']) {
                 $profPath = Split-Path -Parent $global:DashboardConfig.Config['LauncherPath']['LauncherPath']
             }
        } elseif ($global:DashboardConfig.Config['Profiles'].Contains($targetProfile)) {
             $profPath = $global:DashboardConfig.Config['Profiles'][$targetProfile]
        }
        
        if ($profPath) {
            $resSelection = $UI.ResolutionSelector.SelectedItem
            if ($resSelection) {
                Set-NeuzResolution -ProfilePath $profPath -ResolutionString $resSelection
            }
        }

        Write-Verbose "  UI: Sync-UIToConfig - Saving Login Settings for: '$targetProfile'" -ForegroundColor DarkMagenta

        if (-not ($global:DashboardConfig.Config['LoginConfig'].Contains($targetProfile))) {
            $global:DashboardConfig.Config['LoginConfig'][$targetProfile] = [ordered]@{}
        }

        $currentProfileLoginConfig = $global:DashboardConfig.Config['LoginConfig'][$targetProfile]

        $null = $coordKey

        foreach ($key in $UI.LoginPickers.Keys) {
            $coordKey = "${key}Coords"
            $currentProfileLoginConfig["${key}Coords"] = $UI.LoginPickers[$key].Text.Text
        }

        $currentProfileLoginConfig['PostLoginDelay'] = $UI.InputPostLoginDelay.Text

        $grid = $UI.LoginConfigGrid
        for ($i=0; $i -le 9; $i++) {
            $row = $grid.Rows[$i]
            $clientNum = $row.Cells[0].Value
            $s = if ($row.Cells[1].Value) { $row.Cells[1].Value.ToString() } else { "1" }
            $c = if ($row.Cells[2].Value) { $row.Cells[2].Value.ToString() } else { "1" }
            $char = if ($row.Cells[3].Value) { $row.Cells[3].Value.ToString() } else { "1" }
            $coll = if ($row.Cells[4].Value) { $row.Cells[4].Value.ToString() } else { "No" }

            $val = "$s,$c,$char,$coll"
            $currentProfileLoginConfig["Client${clientNum}_Settings"] = $val
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

        if ($global:DashboardConfig.Config['LauncherPath']['LauncherPath']) { $UI.InputLauncher.Text = $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] }
        if ($global:DashboardConfig.Config['ProcessName']['ProcessName']) { $UI.InputProcess.Text = $global:DashboardConfig.Config['ProcessName']['ProcessName'] }
        if ($global:DashboardConfig.Config['MaxClients']['MaxClients']) { $UI.InputMax.Text = $global:DashboardConfig.Config['MaxClients']['MaxClients'] }
        if ($global:DashboardConfig.Config['Paths']['JunctionTarget']) { $UI.InputJunction.Text = $global:DashboardConfig.Config['Paths']['JunctionTarget'] }
        if ($global:DashboardConfig.Config.Contains('LauncherPath') -and $global:DashboardConfig.Config['LauncherPath'].Contains('LauncherPath')) { $UI.InputLauncher.Text = $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] }
        if ($global:DashboardConfig.Config.Contains('ProcessName') -and $global:DashboardConfig.Config['ProcessName'].Contains('ProcessName')) { $UI.InputProcess.Text = $global:DashboardConfig.Config['ProcessName']['ProcessName'] }
        if ($global:DashboardConfig.Config.Contains('MaxClients') -and $global:DashboardConfig.Config['MaxClients'].Contains('MaxClients')) { $UI.InputMax.Text = $global:DashboardConfig.Config['MaxClients']['MaxClients'] }
        if ($global:DashboardConfig.Config.Contains('Paths') -and $global:DashboardConfig.Config['Paths'].Contains('JunctionTarget')) { $UI.InputJunction.Text = $global:DashboardConfig.Config['Paths']['JunctionTarget'] }

        if ($global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin']) { $UI.NeverRestartingCollectorLogin.Checked = ([int]$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin']) -eq 1 }
        if ($global:DashboardConfig.Config.Contains('Login') -and $global:DashboardConfig.Config['Login'].Contains('NeverRestartingCollectorLogin')) { $UI.NeverRestartingCollectorLogin.Checked = ([int]$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin']) -eq 1 }


        $UI.ProfileGrid.Rows.Clear()

        $recProfiles = if ($global:DashboardConfig.Config['ReconnectProfiles']) { $global:DashboardConfig.Config['ReconnectProfiles'] } else { @{} }
        $hideProfiles = if ($global:DashboardConfig.Config['HideProfiles']) { $global:DashboardConfig.Config['HideProfiles'] } else { @{} }

        if ($global:DashboardConfig.Config['Profiles']) {
            $profiles = $global:DashboardConfig.Config['Profiles']
            foreach ($key in $profiles.Keys) {
                $path = $profiles[$key]

                $isRecChecked = $recProfiles.Contains($key)
                $isHideChecked = $hideProfiles.Contains($key)

                $UI.ProfileGrid.Rows.Add($key, $path, $isRecChecked, $isHideChecked) | Out-Null
            }
        }

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
                    $UI.ProfileGrid.Tag = $row.Index
                    $found = $true
                    break
                }
            }
            if (-not $found) {
                $UI.ProfileGrid.ClearSelection()
                $UI.ProfileGrid.Tag = -1
            }
        } else {
             $UI.ProfileGrid.ClearSelection()
             $UI.ProfileGrid.Tag = -1
        }

        $selectedLoginProfile = $null
        if ($UI.LoginProfileSelector.SelectedItem) {
            $selectedLoginProfile = $UI.LoginProfileSelector.SelectedItem.ToString()
        }
        if (-not $selectedLoginProfile) {
            $selectedLoginProfile = "Default"
        }

        $UI.LoginProfileSelector.Tag = $selectedLoginProfile

        $profPath = $null
        if ($selectedLoginProfile -eq 'Default') {
             if ($global:DashboardConfig.Config['LauncherPath']['LauncherPath']) {
                 $profPath = Split-Path -Parent $global:DashboardConfig.Config['LauncherPath']['LauncherPath']
             }
        } elseif ($global:DashboardConfig.Config['Profiles'].Contains($selectedLoginProfile)) {
             $profPath = $global:DashboardConfig.Config['Profiles'][$selectedLoginProfile]
        }
        
        if ($profPath) {
            $currentRes = Get-NeuzResolution -ProfilePath $profPath
            if ($currentRes -and $UI.ResolutionSelector.Items.Contains($currentRes)) {
                $UI.ResolutionSelector.SelectedItem = $currentRes
            } else {
                $UI.ResolutionSelector.SelectedItem = $null
            }
        } else {
            $UI.ResolutionSelector.SelectedItem = $null
        }

        Write-Verbose "  UI: Sync-ConfigToUI - Loading settings for profile '$selectedLoginProfile'." -ForegroundColor DarkMagenta

        if ($global:DashboardConfig.Config['LoginConfig'] -and $global:DashboardConfig.Config['LoginConfig'][$selectedLoginProfile]) {
            $lc = $global:DashboardConfig.Config['LoginConfig'][$selectedLoginProfile]
        } else {
            $lc = [ordered]@{}
        }

        foreach ($key in $UI.LoginPickers.Keys) {
            $coordKey = "${key}Coords"
            if ($lc[$coordKey]) {
                $UI.LoginPickers[$key].Text.Text = $lc[$coordKey]
            } else {
                $UI.LoginPickers[$key].Text.Text = "0,0"
            }
        }

        if ($lc['PostLoginDelay']) { $UI.InputPostLoginDelay.Text = $lc['PostLoginDelay'] }
        else { $UI.InputPostLoginDelay.Text = "1" }

        $grid = $UI.LoginConfigGrid
        for ($i=0; $i -le 9; $i++) {
            $row = $grid.Rows[$i]
            $clientNum = $i + 1
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

        Write-Verbose '  UI: Config synced to UI' -ForegroundColor Green
        return $true
    }
    catch
    {
        Write-Verbose "  UI: Failed to sync config to UI: $_" -ForegroundColor Red
        return $false
    }
}

#endregion

#region Core Functions

function Initialize-UI
{
    [CmdletBinding()]
    param()

    if (-not ([System.Management.Automation.PSTypeName]'FolderBrowser').Type) {
        Add-Type -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Reflection;

public class FolderBrowser
{
    public string SelectedPath { get; set; }
    public string Description { get; set; }
    
    public DialogResult ShowDialog(IWin32Window owner)
    {
        var ofd = new OpenFileDialog();
        ofd.FileName = "Select Folder";
		ofd.Filter = "Folders|";
        ofd.CheckFileExists = false;
        ofd.CheckPathExists = true;
        ofd.Title = Description;
        if (!string.IsNullOrEmpty(SelectedPath)) ofd.InitialDirectory = SelectedPath;

        var type = ofd.GetType();
        var createVistaDialog = type.GetMethod("CreateVistaDialog", BindingFlags.Instance | BindingFlags.NonPublic);
        
        if (createVistaDialog != null)
        {
            var dialog = createVistaDialog.Invoke(ofd, null);
            var dialogType = dialog.GetType();
            var setOptions = dialogType.GetMethod("SetOptions");
            var getOptions = dialogType.GetMethod("GetOptions");
            
            if (setOptions != null && getOptions != null)
            {
                uint options = (uint)getOptions.Invoke(dialog, null);
                setOptions.Invoke(dialog, new object[] { options | 0x20 }); // FOS_PICKFOLDERS
                
                var show = dialogType.GetMethod("Show");
                int result = (int)show.Invoke(dialog, new object[] { owner.Handle });
                
                if (result == 0) // S_OK
                {
                    var getResult = dialogType.GetMethod("GetResult");
                    var item = getResult.Invoke(dialog, null);
                    var getDisplayName = item.GetType().GetMethod("GetDisplayName");
                    SelectedPath = (string)getDisplayName.Invoke(item, new object[] { 0x80058000 }); // SIGDN_FILESYSPATH
                    return DialogResult.OK;
                }
                return DialogResult.Cancel;
            }
        }
        
        // Fallback to standard FolderBrowserDialog if Vista style fails
        using (var fbd = new FolderBrowserDialog())
        {
            fbd.Description = Description;
            fbd.SelectedPath = SelectedPath;
            var res = fbd.ShowDialog(owner);
            if (res == DialogResult.OK) SelectedPath = fbd.SelectedPath;
            return res;
        }
    }
}
"@ -ReferencedAssemblies System.Windows.Forms
    }

    Write-Verbose '  UI: Initializing UI...' -ForegroundColor Cyan

    $uiPropertiesToAdd = @{}

    $p = @{ type='Form'; visible=$false; width=470; height=440; bg=@(30, 30, 30); id='MainForm'; text='Entropia Dashboard'; startPosition='CenterScreen'; formBorderStyle=[System.Windows.Forms.FormBorderStyle]::None }
    $mainForm = Set-UIElement @p


    if (-not $global:DashboardConfig.UI) { $global:DashboardConfig | Add-Member -MemberType NoteProperty -Name UI -Value ([PSCustomObject]@{}) -Force }

    $toolTipMain = New-Object System.Windows.Forms.ToolTip
    $toolTipMain.AutoPopDelay = 5000
    $toolTipMain.InitialDelay = 100
    $toolTipMain.ReshowDelay = 10
    $toolTipMain.ShowAlways = $true

    $global:DashboardConfig.UI | Add-Member -MemberType NoteProperty -Name ToolTip -Value $toolTipMain -Force

    $p = @{ type='Form'; width=600; height=655; bg=@(30, 30, 30); id='SettingsForm'; text='Settings'; startPosition='CenterScreen'; formBorderStyle=[System.Windows.Forms.FormBorderStyle]::None; topMost=$false; opacity=0.0 }
    $settingsForm = Set-UIElement @p
    $settingsForm.Owner = $mainForm


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
    $p = @{ type='Label'; width=140; height=10; top=16; left=10; fg=@(230, 230, 230); id='CopyrightLabel'; text=[char]0x00A9 + ' Immortal / Divine 2026 - v2.1'; font=(New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Italic)) }
    $copyrightLabelForm = Set-UIElement @p
    $p = @{ type='Button'; width=30; height=30; left=410; bg=@(40, 40, 40); fg=@(240, 240, 240); id='MinForm'; text='_'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)); tooltip='Minimize' }
    $btnMinimizeForm = Set-UIElement @p
    $p = @{ type='Button'; width=30; height=30; left=440; bg=@(150, 20, 20); fg=@(240, 240, 240); id='CloseForm'; text=[char]0x166D; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)); tooltip='Exit' }
    $btnCloseForm = Set-UIElement @p


    $p = @{ type='Button'; width=125; height=30; top=40; left=15; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Launch'; text='Launch ' + [char]0x25BE; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Start Launch Process defined in Settings / Right Click for manual Start.' }
    $p = @{ type='Button'; width=125; height=30; top=40; left=15; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Launch'; text='Launch ' + [char]0x25BE; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Open Launch Menu / Left Click to Cancel (if active).' }
    $btnLaunch = Set-UIElement @p

    $LaunchContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $LaunchContextMenu.Name = 'LaunchContextMenu'
    $LaunchContextMenu.Items.Add("Loading...") | Out-Null
    $btnLaunch.ContextMenuStrip = $LaunchContextMenu

    $p = @{ type='Button'; width=125; height=30; top=40; left=150; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Login'; text='Login'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Login selected Clients with Default or Profile Settings / Nickname List with 10 nicknames and 1024x768 mandatory.' }
    $btnLogin = Set-UIElement @p
    $p = @{ type='Button'; width=80; height=30; top=40; left=285; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Settings'; text='Settings'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Edit Dashboard Settings.' }
    $btnSettings = Set-UIElement @p
    $p = @{ type='Button'; width=80; height=30; top=40; left=375; bg=@(150, 20, 20); fg=@(240, 240, 240); id='Terminate'; text='Terminate'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Closes all selected Clients instantly.' }
    $btnStop = Set-UIElement @p
    $p = @{ type='Button'; width=440; height=30; top=75; left=15; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Ftool'; text='Ftool'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Start Ftool for selected Clients.' }
    $btnFtool = Set-UIElement @p

    $p = @{ type='DataGridView'; visible=$false; width=155; height=320; top=115; left=5; bg=@(40, 40, 40); fg=@(240, 240, 240); id='DataGridMain'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $DataGridMain = Set-UIElement @p

    $p = @{ type='DataGridView'; width=450; height=300; top=115; left=10; bg=@(40, 40, 40); fg=@(240, 240, 240); id='DataGridFiller'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $DataGridFiller = Set-UIElement @p

    $mainGridCols = @(
        (New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Index'; HeaderText = '#'; FillWeight = 10; SortMode = 'NotSortable';}),
        (New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Titel'; HeaderText = 'Titel'; SortMode = 'NotSortable';}),
        (New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'ID'; HeaderText = 'PID'; FillWeight = 20; SortMode = 'NotSortable';}),
        (New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'State'; HeaderText = 'State'; FillWeight = 30; SortMode = 'NotSortable';})    )
    $DataGridFiller.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$mainGridCols)

    $p = @{ type='TextProgressBar'; width=450; height=18; top=415; left=10; id='GlobalProgressBar'; style='Continuous'; visible=$false }
    $GlobalProgressBar = Set-UIElement @p

    $toolTipSettings = New-Object System.Windows.Forms.ToolTip
    $toolTipSettings.AutoPopDelay = 5000
    $toolTipSettings.InitialDelay = 100
    $toolTipSettings.ReshowDelay = 10
    $toolTipSettings.ShowAlways = $true

    $global:DashboardConfig.UI.ToolTip = $toolTipSettings

    $settingsTabs = New-Object Custom.DarkTabControl
    $settingsTabs.Dock = 'Top'
    $settingsTabs.Height = 575

    $tabGeneral = New-Object System.Windows.Forms.TabPage "General"
    $tabGeneral.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)


    $tabLoginSettings = New-Object System.Windows.Forms.TabPage "Login Settings"
    $tabLoginSettings.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $tabLoginSettings.AutoScroll = $true

    $settingsTabs.TabPages.Add($tabGeneral)
    $settingsTabs.TabPages.Add($tabLoginSettings)

    $settingsForm.Controls.Add($settingsTabs)

    $p = @{ type='Label'; width=127; height=20; top=25; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelLauncher'; text='Select Main Launcher:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Select your Main Launcher, this will be used as Default Launcher Path for all features.' }
    $lblLauncher = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=50; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputLauncher'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Select your Main Launcher, this will be used as Default Launcher Path for all features.' }
    $txtLauncher = Set-UIElement @p
    $p = @{ type='Button'; width=55; height=25; top=20; left=150; bg=@(40, 40, 40); fg=@(240, 240, 240); id='Browse'; text='Browse'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Select your Main Launcher, this will be used as Default Launcher Path for all features.' }
    $btnBrowseLauncher = Set-UIElement @p

    $p = @{ type='Label'; width=85; height=20; top=95; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelProcess'; text='Process Name:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Enter Process Name, by default neuz.' }
    $lblProcessName = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=120; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputProcess'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Enter Process Name, by default neuz.' }
    $txtProcessName = Set-UIElement @p

    $p = @{ type='Label'; width=250; height=20; top=165; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelMax'; text='Max Total Clients For Selection:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Set maximum total amount of profile clients, that the dashboard is allowed to launch up.' }
    $lblMaxClients = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=190; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputMax'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Set maximum total amount of clients, that the dashboard is allowed to launch up to.' }
    $txtMaxClients = Set-UIElement @p

	$p = @{ type='Button'; width=120; height=25; top=295; left=150; bg=@(60, 60, 100); fg=@(240, 240, 240); id='SaveLaunchState'; text='Save One-Click Setup'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 8)); tooltip='Save the current running clients (Profiles & Count) as the desired Launch and Login Configuration.' }
    $btnSaveLaunchState = Set-UIElement @p

    $p = @{ type='Label'; width=125; height=20; top=235; left=20; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelJunction'; text='Select Profiles Folder:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='This is the default folder where all your Profiles will be.' }
    $lblJunction = Set-UIElement @p
    $p = @{ type='TextBox'; width=250; height=30; top=260; left=20; bg=@(40, 40, 40); fg=@(240, 240, 240); id='InputJunction'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='This is the default folder where all your Profiles will be.' }
    $txtJunction = Set-UIElement @p

    $p = @{ type='Button'; width=55; height=25; top=230; left=145; bg=@(40, 40, 40); fg=@(240, 240, 240); id='BrowseJunction'; text='Browse'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Select the default folder where all your Profiles will be.' }
    $btnBrowseJunction = Set-UIElement @p

    $p = @{ type='Button'; width=55; height=25; top=230; left=215; bg=@(35, 175, 75); fg=@(240, 240, 240); id='StartJunction'; text='Create'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Create a ~300MB copy of your main client as new Profile with separate Client settings.' }
    $btnStartJunction = Set-UIElement @p

    $p = @{ type='Button'; width=120; height=25; top=295; left=20; bg=@(60, 60, 100); fg=@(240, 240, 240); id='CopyNeuz'; text='Copy Neuz.exe'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Copy neuz.exe from main client to all profiles. This is the final step to complete the patching process of the Main Launcher for your Profiles.' }
    $btnCopyNeuz = Set-UIElement @p

    $p = @{ type='CheckBox'; width=200; height=20; top=360; left=0; bg=@(30, 30, 30); fg=@(240, 240, 240); id='NeverRestartingCollectorLogin'; text='Collector Double Click'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='In rare cases the Collector Button has to be clicked twice to be started, tick this checkbox to fix this.' }
    $chkNeverRestartingLogin = Set-UIElement @p

	$p = @{ type='Label'; width=220; height=20; top=25; left=300; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); id='LabelProfiles'; text='Select Client Profile for Default Launch:'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Selected Profile will be used for the Default Launch sequence.' }
    $lblProfiles = Set-UIElement @p

    $p = @{ type='DataGridView'; width=260; height=300; top=50; left=300; bg=@(40, 40, 40); fg=@(240, 240, 240); id='ProfileGrid'; text=''; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
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
    $colProfName.HeaderText = "Name"; $colProfName.FillWeight = 20; $colProfName.ReadOnly = $true
	$colProfName.ToolTipText = "Name of the profile."

    $colProfPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colProfPath.HeaderText = "Path"; $colProfPath.FillWeight = 25; $colProfPath.ReadOnly = $true
	$colProfPath.ToolTipText = "Path of the profile."

    $colReconnect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colReconnect.HeaderText = "Reconnect?";
    $colReconnect.FillWeight = 25;
    $colReconnect.ReadOnly = $false
    $colReconnect.ToolTipText = "Enables Auto-Reconnect for this profile."

    $colHide = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colHide.HeaderText = "Hide?";
    $colHide.FillWeight = 15;
    $colHide.ReadOnly = $false
    $colHide.ToolTipText = "Hides from Taskbar and Alt+Tab View when minimized."

    $ProfileGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colProfName, $colProfPath, $colReconnect, $colHide))

	$p = @{ type='Button'; width=60; height=25; top=360; left=300; bg=@(40, 40, 40); fg=@(240, 240, 240); id='AddProfile'; text='Add'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 8)); tooltip='Manually add an existing folder as a profile.' }
    $btnAddProfile = Set-UIElement @p
    $p = @{ type='Button'; width=60; height=25; top=360; left=365; bg=@(40, 40, 40); fg=@(240, 240, 240); id='RenameProfile'; text='Rename'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 8)); tooltip='Rename selected profile.' }
    $btnRenameProfile = Set-UIElement @p
    $p = @{ type='Button'; width=60; height=25; top=360; left=430; bg=@(40, 40, 40); fg=@(240, 240, 240); id='RemoveProfile'; text='Remove'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 8)); tooltip='Remove selected profile from list.' }
    $btnRemoveProfile = Set-UIElement @p
    $p = @{ type='Button'; width=60; height=25; top=360; left=495; bg=@(150, 20, 20); fg=@(240, 240, 240); id='DeleteProfile'; text='Delete'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 8)); tooltip='Permanently delete profile folder from disk.' }
    $btnDeleteProfile = Set-UIElement @p

	$tabGeneral.Controls.AddRange(@($lblLauncher, $txtLauncher, $btnBrowseLauncher, $lblProcessName, $txtProcessName, $lblMaxClients, $txtMaxClients, $btnSaveLaunchState, $lblJunction, $txtJunction, $btnBrowseJunction, $btnStartJunction, $btnCopyNeuz, $chkNeverRestartingLogin, $lblProfiles, $ProfileGrid, $btnAddProfile, $btnRenameProfile, $btnRemoveProfile, $btnDeleteProfile))

	$p = @{ type='Label'; width=80; height=20; top=10; left=10; bg=@(30,30,30,0); fg=@(240,240,240); text="Edit Profile:"; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $lblProfSel = Set-UIElement @p

    $p = @{ type='Custom.DarkComboBox'; width=430; height=25; top=7; left=100; bg=@(40,40,40); fg=@(240,240,240); id='LoginProfileSelector'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); dropDownStyle='DropDownList' }
    $ddlLoginProfile = Set-UIElement @p
    $ddlLoginProfile.Items.Add("Default") | Out-Null

    $ddlLoginProfile.Tag = "Default"

    if ($global:DashboardConfig.Config['Profiles']) {
        foreach ($key in $global:DashboardConfig.Config['Profiles'].Keys) {
            $ddlLoginProfile.Items.Add($key) | Out-Null
        }
    }
    $ddlLoginProfile.SelectedIndex = 0

    $tabLoginSettings.Controls.AddRange(@($lblProfSel, $ddlLoginProfile))

    $p = @{ type='Label'; width=80; height=20; top=40; left=10; bg=@(30,30,30,0); fg=@(240,240,240); text="Resolution:"; font=(New-Object System.Drawing.Font('Segoe UI', 9)) }
    $lblResolution = Set-UIElement @p

    $p = @{ type='Custom.DarkComboBox'; width=430; height=25; top=37; left=100; bg=@(40,40,40); fg=@(240,240,240); id='ResolutionSelector'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); dropDownStyle='DropDownList' }
    $ddlResolution = Set-UIElement @p
    
    $resolutions = @(
        "1024 x 768", "1280 x 720", "1280 x 768", "1280 x 800", "1280 x 1024",
        "1360 x 768", "1440 x 900", "1600 x 815", "1600 x 900", "1600 x 1200",
        "1680 x 1050", "1920 x 1080", "1920 x 1440", "2560 x 1080", "2560 x 1440",
        "3440 x 1440", "3840 x 2160", "4096 x 3072", "5120 x 1440"
    )
    foreach ($res in $resolutions) { $ddlResolution.Items.Add($res) | Out-Null }

    $tabLoginSettings.Controls.AddRange(@($lblResolution, $ddlResolution))

    $Pickers = @{}
    $rowY = 45

    $AddPickerRow = {
        param($LabelText, $KeyName, $Top, $Col, $Tooltip)
        $leftOffset = if($Col -eq 2) { 300 } else { 10 }

        $p = @{ type='Label'; width=100; height=20; top=$Top; left=$leftOffset; bg=@(30, 30, 30, 0); fg=@(240, 240, 240); text=$LabelText; font=(New-Object System.Drawing.Font('Segoe UI', 8)); tooltip=$Tooltip }
        $l = Set-UIElement @p
        $p = @{ type='TextBox'; width=80; height=20; top=$Top; left=($leftOffset+100); bg=@(40, 40, 40); fg=@(240, 240, 240); id="txt$KeyName"; text='0,0'; font=(New-Object System.Drawing.Font('Segoe UI', 8)); tooltip=$Tooltip }
        $t = Set-UIElement @p
        $p = @{ type='Button'; width=40; height=20; top=$Top; left=($leftOffset+185); bg=@(60, 60, 100); fg=@(240, 240, 240); id="btnPick$KeyName"; text='Set'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 7)); tooltip=$Tooltip }
        $b = Set-UIElement @p
        return @{ Label=$l; Text=$t; Button=$b }
    }

    $p = &$AddPickerRow "Server 1:" "Server1" ($rowY+30) 1 'Click the first server on the server selection screen.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Server1"] = $p
    $p = &$AddPickerRow "Server 2:" "Server2" ($rowY+55) 1 'Click the second server on the server selection screen.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Server2"] = $p
    $p = &$AddPickerRow "Channel 1:" "Channel1" ($rowY+80) 1 'Click the first channel on the channel selection screen.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Channel1"] = $p
    $p = &$AddPickerRow "Channel 2:" "Channel2" ($rowY+105) 1 'Click the second channel on the channel selection screen.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Channel2"] = $p
    $p = &$AddPickerRow "First Nickname:" "FirstNick" ($rowY+130) 1 'Click the first character nickname in the account selection list.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["FirstNick"] = $p
    $p = &$AddPickerRow "Scroll Down Arrow" "ScrollDown"($rowY+155) 1 'Click the scroll-down arrow in the account selection list.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["ScrollDown"] = $p

    $p = &$AddPickerRow "Char Slot 1:" "Char1" ($rowY+30) 2 'Click the first character slot.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Char1"] = $p
    $p = &$AddPickerRow "Char Slot 2:" "Char2" ($rowY+55) 2 'Click the second character slot.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Char2"] = $p
    $p = &$AddPickerRow "Char Slot 3:" "Char3" ($rowY+80) 2 'Click the third character slot.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["Char3"] = $p

    $p = &$AddPickerRow "Collector Start:" "CollectorStart" ($rowY+105) 2 'Click the ''Start'' button for the collector/bot.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["CollectorStart"] = $p
    $p = &$AddPickerRow "Disconnect OK:" "DisconnectOK" ($rowY+130) 2 'Click the ''OK'' button on the disconnect confirmation dialog.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["DisconnectOK"] = $p
    $p = &$AddPickerRow "Login Wrong OK:" "LoginDetailsOK" ($rowY+155) 2 'Click the ''OK'' button on the ''wrong login details'' dialog.'; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers["LoginDetailsOK"] = $p

    $p = @{ type='Label'; width=150; height=20; top=($rowY+180); left=10;
    bg=@(30, 30, 30, 0); fg=@(240, 240, 240); text="Post-Login Delay (s):"; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Time in seconds to wait after logging in before starting the collector or minimizing the window.' }
    $lblPostLoginDelay = Set-UIElement @p
    $p = @{ type='TextBox'; width=80; height=20; top=($rowY+180); left=160; bg=@(40, 40, 40); fg=@(240, 240, 240); id='txtPostLoginDelayInput'; text='5'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Time in seconds to wait after logging in before starting the collector or minimizing the window.' }
    $txtPostLoginDelay = Set-UIElement @p
    $tabLoginSettings.Controls.AddRange(@($lblPostLoginDelay, $txtPostLoginDelay))

    $gridTop = $rowY + 205
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
    $LoginConfigGrid.MultiSelect = $false

    $colClient = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colClient.HeaderText = "Client #"; $colClient.FillWeight = 15; $colClient.ReadOnly = $true
	$colClient.ToolTipText = "Position in Datagrid and # Nickname to login."

    $colSrv = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSrv.HeaderText = "Server"; $colSrv.FillWeight = 20; $colSrv.ReadOnly = $true
	$colSrv.ToolTipText = "Select which sever to collect to."

    $colCh = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colCh.HeaderText = "Channel"; $colCh.FillWeight = 20; $colCh.ReadOnly = $true
	$colCh.ToolTipText = "Select which channel to connect to."

    $colChar = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colChar.HeaderText = "Character"; $colChar.FillWeight = 25; $colChar.ReadOnly = $true
	$colChar.ToolTipText = "Select the character to login."

    $colColl = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colColl.HeaderText = "Collecting?"; $colColl.FillWeight = 20; $colColl.ReadOnly = $true
	$colColl.ToolTipText = "If the character should start collecting. Be sure to have the Collector item equipped."

    $cols = @($colClient, $colSrv, $colCh, $colChar, $colColl)
    $LoginConfigGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$cols)

    for ($i=1; $i -le 10; $i++) {
        $LoginConfigGrid.Rows.Add($i, "1", "1", "1", "No") | Out-Null
    }

    $tabLoginSettings.Controls.Add($LoginConfigGrid)

    $p = @{ type='Button'; width=120; height=40; top=585; left=20; bg=@(35, 175, 75); fg=@(240, 240, 240); id='Save'; text='Save'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Save all settings' }
    $btnSave = Set-UIElement @p

    $p = @{ type='Button'; width=120; height=40; top=585; left=150; bg=@(210, 45, 45); fg=@(240, 240, 240); id='Cancel'; text='Cancel'; fs='Flat'; font=(New-Object System.Drawing.Font('Segoe UI', 9)); tooltip='Close and do not save' }
    $btnCancel = Set-UIElement @p

    $settingsForm.Controls.AddRange(@($btnSave, $btnCancel))


    $mainForm.Controls.AddRange(@($topBar, $btnLogin, $btnFtool, $btnLaunch, $btnSettings, $btnStop, $DataGridMain, $DataGridFiller, $GlobalProgressBar))
    $topBar.Controls.AddRange(@($titleLabelForm, $copyrightLabelForm, $btnMinimizeForm, $btnCloseForm))

    $ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $itmFront = New-Object System.Windows.Forms.ToolStripMenuItem('Show')
    $itmBack = New-Object System.Windows.Forms.ToolStripMenuItem('Minimize')
    $itmResizeCenter = New-Object System.Windows.Forms.ToolStripMenuItem('Resize')
    $itmRelog = New-Object System.Windows.Forms.ToolStripMenuItem('Relog after Disconnect (wip)')
    $ctxMenu.Items.AddRange(@($itmFront, $itmBack, $itmResizeCenter, $itmRelog))
    $DataGridMain.ContextMenuStrip = $ctxMenu
    $DataGridFiller.ContextMenuStrip = $ctxMenu


    $global:DashboardConfig.UI = [PSCustomObject]@{
        MainForm = $mainForm
        SettingsForm = $settingsForm
        SettingsTabs = $settingsTabs
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
        LaunchContextMenu = $LaunchContextMenu
        ToolTip = $toolTipSettings

        InputLauncher = $txtLauncher
        InputJunction = $txtJunction
        StartJunction = $btnStartJunction
        InputProcess = $txtProcessName
		SaveLaunchState = $btnSaveLaunchState
        InputMax = $txtMaxClients
        Browse = $btnBrowseLauncher
        BrowseJunction = $btnBrowseJunction
        NeverRestartingCollectorLogin = $chkNeverRestartingLogin
        ProfileGrid = $ProfileGrid
        AddProfile = $btnAddProfile
        RenameProfile = $btnRenameProfile
        RemoveProfile = $btnRemoveProfile
        DeleteProfile = $btnDeleteProfile
        CopyNeuz = $btnCopyNeuz

        LoginConfigGrid = $LoginConfigGrid
        LoginPickers = $Pickers
        InputPostLoginDelay = $txtPostLoginDelay
        LoginProfileSelector = $ddlLoginProfile
        ResolutionSelector = $ddlResolution

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

    Register-UIEventHandlers
    return $true
}

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

                            if (Get-Command Update-ReconnectSupervisor -ErrorAction SilentlyContinue) {
                                Update-ReconnectSupervisor -Enabled $true
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

                LoginConfigGrid = @{
                    CellClick = {
                        param($s, $e)
                        if ($e.RowIndex -lt 0) { return }

                        $row = $s.Rows[$e.RowIndex]
                        $colIndex = $e.ColumnIndex

                        if ($colIndex -eq 1) { $row.Cells[1].Value = if ($row.Cells[1].Value -eq "1") { "2" } else { "1" } }
                        elseif ($colIndex -eq 2) { $row.Cells[2].Value = if ($row.Cells[2].Value -eq "1") { "2" } else { "1" } }
                        elseif ($colIndex -eq 3) { $row.Cells[3].Value = switch ($row.Cells[3].Value) { "1" {"2"}; "2" {"3"}; "3" {"1"}; default {"1"} } }
                        elseif ($colIndex -eq 4) { $row.Cells[4].Value = if ($row.Cells[4].Value -eq "Yes") { "No" } else { "Yes" } }
                    }
                }
                LoginProfileSelector = @{
                    SelectedIndexChanged = {
                        param($s, $e)
                        $oldProfile = $s.Tag
                        if (-not $oldProfile) { $oldProfile = "Default" }

                        Write-Verbose "  UI: Saving settings for '$oldProfile' before switching..." -ForegroundColor Cyan
                        Sync-UIToConfig -ProfileToSave $oldProfile
                        if (Get-Command Write-Config -ErrorAction SilentlyContinue) { Write-Config }

                        Write-Verbose "  UI: Loading settings for new selection..." -ForegroundColor Cyan
                        Sync-ConfigToUI
                    }
                }

				SaveLaunchState = @{
                    Click = {
                        # 1. Capture State
                        $grid = $global:DashboardConfig.UI.DataGridFiller
                        if ($grid.Rows.Count -eq 0) {
                            [System.Windows.Forms.MessageBox]::Show("The client list is empty.", "Save State", "OK", "Information")
                            return
                        }

                        $detailedStateList = [System.Collections.Generic.List[PSObject]]::new()
                        $profileCounters = @{}

                        # Sort visually to keep order
                        $sortedRows = $grid.Rows | Sort-Object Index

                        foreach ($row in $sortedRows) {
                            $gridPosition = $row.Index + 1
                            $fullTitle = $row.Cells[1].Value.ToString()
                            $profileName = "Default"
                            $cleanTitle = $fullTitle

                            # Extract [Profile]
                            if ($fullTitle -match '^\[([^\]]+)\](.*)') {
                                $profileName = $matches[1]
                                $cleanTitle = $matches[2].Trim()
                            }

                            # --- INTELLIGENT DETECTION ---
                            # If title contains " - ", we assume it is "Server - CharName" -> Login Required
                            # If NO " - ", it is just "Server" -> No Login (Launcher/Server Select)
                            $isLoggedIn = $false
                            $serverName = $cleanTitle
                            $charName = ""

                            if ($cleanTitle -match '^(.*?) - (.*)$') {
                                $serverName = $matches[1]
                                $charName = $matches[2]
                                $isLoggedIn = $true
                            }

                            # Track profile index (Mage #1, Mage #2)
                            if (-not $profileCounters.ContainsKey($profileName)) {
                                $profileCounters[$profileName] = 0
                            }
                            $profileCounters[$profileName]++
                            $profileIndex = $profileCounters[$profileName]

                            $detailedStateList.Add([PSCustomObject]@{
                                GridPosition  = $gridPosition
                                Profile       = $profileName
                                ProfileIndex  = $profileIndex
                                FullTitle     = $fullTitle
                                ServerName    = $serverName
                                Character     = $charName
                                IsLoggedIn    = $isLoggedIn
                            })
                        }

                        # 2. Save to Config
                        # CSV Format: GridPos, Profile, ProfileIndex, Title, LoginFlag(1/0)
                        $configSection = [ordered]@{}
                        for ($i = 0; $i -lt $detailedStateList.Count; $i++) {
                            $client = $detailedStateList[$i]
                            $loginFlag = if ($client.IsLoggedIn) { 1 } else { 0 }

                            # Sanitize title (remove commas) just in case
                            $safeTitle = $client.FullTitle -replace ',', ''

                            $valueString = "$($client.GridPosition),$($client.Profile),$($client.ProfileIndex),$safeTitle,$loginFlag"
                            $configSection["Client$i"] = $valueString
                        }

                        $global:DashboardConfig.Config['SavedLaunchConfig'] = $configSection
                        if (Get-Command Write-Config -ErrorAction SilentlyContinue) { Write-Config }

                        # 3. Detailed Summary Message
                        $sb = [System.Text.StringBuilder]::new()
                        $sb.AppendLine("Launch Configuration Saved Successfully!")
                        $sb.AppendLine("==============================")

                        foreach($client in $detailedStateList) {
                            $typeStr = if ($client.IsLoggedIn) { "AUTO-LOGIN -> $($client.Character)" } else { "NO LOGIN" }

                            $sb.AppendLine(
                                "[Pos:$($client.GridPosition.ToString().PadRight(2))] " +
                                "[$($client.Profile)] " +
                                "$($client.ServerName) : $typeStr"
                            )
                        }

                        $sb.AppendLine("==============================")
                        $sb.AppendLine("Total: $($detailedStateList.Count) clients.")
                        $sb.Append("Use 'Launch -> Start and Login saved configuration' to restore.")

                        [System.Windows.Forms.MessageBox]::Show($sb.ToString(), "Configuration Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                }
                LaunchContextMenu = @{
					Opening = {
						param($s, $e)
						$s.Items.Clear()
						
						$selectedProfileName = "Default"
						if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['SelectedProfile'] -and -not [string]::IsNullOrWhiteSpace($global:DashboardConfig.Config['Options']['SelectedProfile'])) {
							$selectedProfileName = $global:DashboardConfig.Config['Options']['SelectedProfile']
						}

						$maxClients = 1
						if ($global:DashboardConfig.Config['MaxClients'] -and $global:DashboardConfig.Config['MaxClients']['MaxClients']) {
							if (-not ([int]::TryParse($global:DashboardConfig.Config['MaxClients']['MaxClients'], [ref]$maxClients))) {
								$maxClients = 1
							}
						}

						$processName = "neuz"
						if ($global:DashboardConfig.Config['ProcessName'] -and $global:DashboardConfig.Config['ProcessName']['ProcessName']) {
							$processName = $global:DashboardConfig.Config['ProcessName']['ProcessName']
						}
						
                        $profileClientCount = 0
                        if (Get-Command Get-ProcessProfile -ErrorAction SilentlyContinue) {
                            $allClients = Get-Process -Name $processName -ErrorAction SilentlyContinue
                            foreach ($client in $allClients) {
                                $clientProfile = Get-ProcessProfile -Process $client
                                if (([string]::IsNullOrEmpty($clientProfile) -and $selectedProfileName -eq 'Default') -or $clientProfile -eq $selectedProfileName) {
                                    $profileClientCount++
                                }
                            }
                        } else {
                            $profileClientCount = @(Get-Process -Name $processName -ErrorAction SilentlyContinue).Count
                        }
						$howManyUntilMax = [Math]::Max(0, $maxClients - $profileClientCount)

						$header = $s.Items.Add("Launch $($howManyUntilMax)x '$($selectedProfileName)'")
						$header.Enabled = $true
						$header.add_Click({
							Start-ClientLaunch
						})


						$s.Items.Add("-")

						$defaultItem = $s.Items.Add("Default / Selected")


						1..10 | ForEach-Object {
							$count = $_
							$subItem = $defaultItem.DropDownItems.Add("Start $($count)x")
							$subItem.Tag = $count

							$subItem.add_Click({
								param($s, $ev)
								$launchCount = $s.Tag
								Start-ClientLaunch -ClientAddCount $launchCount
							})
						}

						if ($global:DashboardConfig.Config['Profiles']) {
							$profiles = $global:DashboardConfig.Config['Profiles']
							if ($profiles.Count -gt 0) {
								$s.Items.Add("-")
								foreach ($key in $profiles.Keys) {
									$profileItem = $s.Items.Add($key)

									1..10 | ForEach-Object {
										$count = $_
										$subItem = $profileItem.DropDownItems.Add("Start $($count)x")

										$subItem.Tag = @{
											ProfileName = $key
											Count = $count
										}

										$subItem.add_Click({
											param($s, $ev)
											$data = $s.Tag
											Start-ClientLaunch -ProfileNameOverride $data.ProfileName -ClientAddCount $data.Count
										})
									}
								}
							}
						}

						$s.Items.Add("-")

						$savedLaunch = $s.Items.Add("Start and Login saved One-Click configuration")
						$savedLaunch.add_Click({
							Start-ClientLaunch -SavedLaunchLoginConfig
						})
					}
				}

                SettingsForm = @{
                    Load = { Sync-ConfigToUI }
                    Move = {
                        $sf = $global:DashboardConfig.UI.SettingsForm
                        $mf = $global:DashboardConfig.UI.MainForm
                        if ($sf -and $mf -and -not $mf.IsDisposed) {
                            $x = $sf.Left + [int](($sf.Width - $mf.Width) / 2)
                            $y = $sf.Top + [int](($sf.Height - $mf.Height) / 2)
                            $mf.Location = New-Object System.Drawing.Point($x, $y)
                        }
                    }
                }
				settingsTabs = @{ MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.SettingsForm.Handle, 0xA1, 0x2, 0) } }

                MinForm = @{ Click = { $global:DashboardConfig.UI.MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized } }
                CloseForm = @{ Click = { $global:DashboardConfig.UI.MainForm.Close() } }
                TopBar = @{ MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.MainForm.Handle, 0xA1, 0x2, 0) } }
                Settings = @{ Click = { Show-SettingsForm } }
                Save = @{
                    Click = {
                        Sync-UIToConfig
                        Write-Config

                        if (Get-Command Update-ReconnectSupervisor -ErrorAction SilentlyContinue) {
                                Update-ReconnectSupervisor -Enabled $true
                            }

                        Hide-SettingsForm
                    }
                }

                Cancel = @{ Click = { Hide-SettingsForm } }
                Browse = @{ Click = { $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Executable Files (*.exe)|*.exe|All Files (*.*)|*.*'; if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $global:DashboardConfig.UI.InputLauncher.Text = $d.FileName } } }
                BrowseJunction = @{
                    Click = {
                        $fb = New-Object FolderBrowser
                        $fb.Description = "Select the destination folder for the Client Copy"
                        if (-not [string]::IsNullOrWhiteSpace($global:DashboardConfig.UI.InputJunction.Text)) {
                            $fb.SelectedPath = $global:DashboardConfig.UI.InputJunction.Text
                        }
                        if ($fb.ShowDialog($global:DashboardConfig.UI.SettingsForm) -eq [System.Windows.Forms.DialogResult]::OK) {
                            $global:DashboardConfig.UI.InputJunction.Text = $fb.SelectedPath
                        }
                    }
                }
                InputJunction = @{
                    Leave = {
                        if ([string]::IsNullOrWhiteSpace($global:DashboardConfig.UI.InputJunction.Text)) {
                            $global:DashboardConfig.UI.InputJunction.Text = $global:DashboardConfig.Paths.Profiles
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
                        $folderName = Split-Path $sourceDir -Leaf
                        if ($folderName -in @('bin32', 'bin64')) {
                             $parentOfBin = Split-Path $sourceDir -Parent
                             $folderName = Split-Path $parentOfBin -Leaf
                        }
                        if ([string]::IsNullOrWhiteSpace($folderName)) { $folderName = "Client" }

                        $userProvidedProfileName = Show-InputBox -Title "New Profile Name" -Prompt "Enter a name for the new client profile folder:" -DefaultText "${folderName}_NewProfile"

                        [string]$finalProfileName = $null
                        [string]$finalDestDir = $null

                        if (-not [string]::IsNullOrWhiteSpace($userProvidedProfileName)) {
                            $userProvidedProfileName = $userProvidedProfileName -replace '[\\\/:*?"<>|\x00-\x1F]', '_'
                            $potentialDestDir = Join-Path $baseParentDir $userProvidedProfileName

                            if (-not (Test-Path $potentialDestDir)) {
                                $finalProfileName = $userProvidedProfileName
                                $finalDestDir = $potentialDestDir
                            } else {
                                [System.Windows.Forms.MessageBox]::Show("The profile name '$userProvidedProfileName' already exists. Falling back to default naming scheme.", "Name Exists", "OK", "Warning")
                                $baseNameFallback = "${folderName}_Copy"
                                $counter = 1
                                $destDirFallback = Join-Path $baseParentDir $baseNameFallback
                                while (Test-Path $destDirFallback) {
                                    $destDirFallback = Join-Path $baseParentDir "${baseNameFallback}_${counter}"
                                    $counter++
                                }
                                $finalProfileName = Split-Path $destDirFallback -Leaf
                                $finalDestDir = $destDirFallback
                            }
                        } else {
                            $baseNameFallback = "${folderName}_Copy"
                            $counter = 1
                            $destDirFallback = Join-Path $baseParentDir $baseNameFallback
                            while (Test-Path $destDirFallback) {
                                $destDirFallback = Join-Path $baseParentDir "${baseNameFallback}_${counter}"
                                $counter++
                            }
                            $finalProfileName = Split-Path $destDirFallback -Leaf
                            $finalDestDir = $destDirFallback
                        }

                        if ([string]::IsNullOrWhiteSpace($finalDestDir) -or [string]::IsNullOrWhiteSpace($finalProfileName)) {
                            [System.Windows.Forms.MessageBox]::Show("Failed to determine a valid profile name or destination directory.", "Error", "OK", "Error")
                            return
                        }

                        try {
                            $sourceFull = (Get-Item $sourceDir).FullName.TrimEnd('\')
                            $destFull = $finalDestDir.TrimEnd('\')
                            if ($destFull -eq $sourceFull -or $destFull.StartsWith($sourceFull + '\', [StringComparison]::OrdinalIgnoreCase)) {
                                [System.Windows.Forms.MessageBox]::Show("The destination folder cannot be inside the source game folder.`nThis would create an endless copy loop.", "Invalid Destination", "OK", "Error")
                                return
                            }
                        } catch {}

                        $confirm = [System.Windows.Forms.MessageBox]::Show("This will create Junctions for 'Data' and 'Effect'. Also a copy of all other files from the $folderName directory to:`n$finalDestDir`n`nProceed?", "Confirm Junction & Copy", "YesNo", "Question")

                        if ($confirm -eq 'Yes') {
                            try {
                                if (-not (Test-Path $finalDestDir)) { New-Item -ItemType Directory -Path $finalDestDir -Force | Out-Null }

                                $junctionFolders = @('Data', 'Effect')

                                $global:DashboardConfig.UI.SettingsForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

                                Get-ChildItem -Path $sourceDir | ForEach-Object {
                                    $itemName = $_.Name
                                    $sourcePath = $_.FullName
                                    $targetPath = Join-Path $finalDestDir $itemName

                                    if ($itemName -in $junctionFolders) {
                                        if (Test-Path $targetPath) {
                                        } else {
                                            $cmdArgs = "/c mklink /J `"$targetPath`" `"$sourcePath`""
                                            Start-Process -FilePath "cmd.exe" -ArgumentList $cmdArgs -WindowStyle Hidden -Wait
                                        }
                                    } else {
                                        Copy-Item -Path $sourcePath -Destination $targetPath -Recurse -Force
                                    }
                                }

                                $global:DashboardConfig.UI.SettingsForm.Cursor = [System.Windows.Forms.Cursors]::Default

                                $global:DashboardConfig.UI.ProfileGrid.Rows.Add($finalProfileName, $finalDestDir) | Out-Null
                                [System.Windows.Forms.MessageBox]::Show("Junctions created and Profile '$finalProfileName' added successfully.`nPlease remember that 'neuz.exe' must be manually patched for each profile on new Flyff update.", "Success", "OK", "Information")
                                Sync-ProfilesToConfig
                                RefreshLoginProfileSelector

                            } catch {
                                $global:DashboardConfig.UI.SettingsForm.Cursor = [System.Windows.Forms.Cursors]::Default
                                [System.Windows.Forms.MessageBox]::Show("An error occurred:`n$($_.Exception.Message)`n`nNote: Creating Junctions may require running the Dashboard as Administrator.", "Error", "OK", "Error")
                            }
                        }
                    }
                }
                AddProfile = @{
                    Click = {
                        $fb = New-Object FolderBrowser
                        $fb.Description = "Select an existing Client folder"
                        if ($fb.ShowDialog($global:DashboardConfig.UI.SettingsForm) -eq [System.Windows.Forms.DialogResult]::OK) {
                            $path = $fb.SelectedPath
                            $defaultName = Split-Path $path -Leaf

                            $profName = Show-InputBox -Title "Profile Name" -Prompt "Enter a name for this profile:" -DefaultText $defaultName

                            if (-not [string]::IsNullOrWhiteSpace($profName)) {
                                $global:DashboardConfig.UI.ProfileGrid.Rows.Add($profName, $path) | Out-Null
                                Sync-ProfilesToConfig
                                RefreshLoginProfileSelector
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

                            if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne $oldName) {
                                $row.Cells[0].Value = $newName

                                if ($global:DashboardConfig.Config['LoginConfig'].Contains($oldName)) {
                                    $data = $global:DashboardConfig.Config['LoginConfig'][$oldName]
                                    $global:DashboardConfig.Config['LoginConfig'].Remove($oldName)
                                    $global:DashboardConfig.Config['LoginConfig'][$newName] = $data
                                }

                                Sync-ProfilesToConfig
                                RefreshLoginProfileSelector
                                Write-Config
                            }
                        } else {
                            [System.Windows.Forms.MessageBox]::Show("Please select a profile to rename.", "Rename Profile", "OK", "Warning")
                        }
                    }
                }
                RemoveProfile = @{
                    Click = {
                        $rowsToRemove = $global:DashboardConfig.UI.ProfileGrid.SelectedRows | ForEach-Object { $_ }
                        foreach ($row in $rowsToRemove) {
                            $profileName = $row.Cells[0].Value.ToString()
                            $profilePath = $row.Cells[1].Value.ToString()
                            $logDir = Join-Path $profilePath "Log"

                            if (Get-Command "Global:Stop-WatcherForLogDir" -ErrorAction SilentlyContinue) {
                                try { Global:Stop-WatcherForLogDir -LogDir $logDir } catch {}
                            }

                            if ($global:DashboardConfig.Config['LoginConfig'].Contains($profileName)) {
                                $global:DashboardConfig.Config['LoginConfig'].Remove($profileName)
                            }

                            $global:DashboardConfig.UI.ProfileGrid.Rows.Remove($row)
                        }
                        Sync-ProfilesToConfig
                        RefreshLoginProfileSelector
                        Write-Config
                    }
                }
                DeleteProfile = @{
                    Click = {
                        $rowsToDelete = $global:DashboardConfig.UI.ProfileGrid.SelectedRows | ForEach-Object { $_ }
                        if ($rowsToDelete.Count -eq 0) {
                             [System.Windows.Forms.MessageBox]::Show("Please select a profile to delete.", "Delete Profile", "OK", "Warning")
                             return
                        }

                        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to PERMANENTLY DELETE the selected profile(s) from the hard drive?`n`nThis action cannot be undone.", "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

                        if ($confirm -eq 'Yes') {
                            foreach ($row in $rowsToDelete) {
                                $profileName = $row.Cells[0].Value.ToString()
                                $profilePath = $row.Cells[1].Value.ToString()
                                $logDir = Join-Path $profilePath "Log"

                                try {
                                    if (Test-Path $profilePath) { Remove-Item -Path $profilePath -Recurse -Force -ErrorAction Stop }
                                    if (Get-Command "Global:Stop-WatcherForLogDir" -ErrorAction SilentlyContinue) { try { Global:Stop-WatcherForLogDir -LogDir $logDir } catch {} }
                                    if ($global:DashboardConfig.Config['LoginConfig'].Contains($profileName)) { $global:DashboardConfig.Config['LoginConfig'].Remove($profileName) }
                                    $global:DashboardConfig.UI.ProfileGrid.Rows.Remove($row)
                                } catch {
                                    [System.Windows.Forms.MessageBox]::Show("Failed to delete profile '$profileName' at '$profilePath'.`nError: $($_.Exception.Message)", "Delete Error", "OK", "Error")
                                }
                            }
                            Sync-ProfilesToConfig
                            RefreshLoginProfileSelector
                            Write-Config
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

                            if (-not $clickedRow.Selected) {
                                $s.ClearSelection()
                                $clickedRow.Selected=$true
                            }
                        }
                    }
                }
				ProfileGrid = @{
                    CellClick = {
                        param($s, $e)
                        if ($e.RowIndex -lt 0) { return }


                        if ($e.ColumnIndex -eq 0 -or $e.ColumnIndex -eq 1) {
                            $s.Tag = $e.RowIndex
                        }
                        elseif ($e.ColumnIndex -eq 2 -or $e.ColumnIndex -eq 3) {
                            $row = $s.Rows[$e.RowIndex]

                            $currentVal = $false
                            if ($row.Cells[$e.ColumnIndex].Value -eq $true -or $row.Cells[$e.ColumnIndex].Value -eq 1) {
                                $currentVal = $true
                            }
                            $row.Cells[$e.ColumnIndex].Value = -not $currentVal
                            $s.EndEdit()

                            $prevIndex = $s.Tag

                            if ($null -ne $prevIndex -and $prevIndex -gt -1 -and $prevIndex -lt $s.Rows.Count) {
                                $s.ClearSelection()
                                $s.Rows[$prevIndex].Selected = $true
                                try { $s.CurrentCell = $s.Rows[$prevIndex].Cells[0] } catch {}
                            } else {
                                $s.ClearSelection()
                            }
                        }
                    }
                }
                ContextMenuFront = @{ Click = { $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { [Custom.Native]::BringToFront($_.Tag.MainWindowHandle) } } }
                ContextMenuBack = @{ Click = { $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { [Custom.Native]::SendToBack($_.Tag.MainWindowHandle) } } }
                ContextMenuResizeAndCenter = @{ Click = { $scr=[System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea; $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { [Custom.Native]::PositionWindow($_.Tag.MainWindowHandle, [Custom.Native]::TopWindowHandle, [int](($scr.Width-1040)/2), [int](($scr.Height-807)/2), 1040, 807, 0x0010) } } }

                Launch = @{
                    Click = {
                        if ($global:DashboardConfig.State.LaunchActive) {
                            try { Stop-ClientLaunch } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message) }
                        } else {
                            $this.ContextMenuStrip.Show($this, 0, $this.Height)
                        }
                    }
                }

                LoginButton = @{
                    Click = {
                        try {
                            $loginCommand = Get-Command LoginSelectedRow -ErrorAction Stop
                            & $loginCommand
                        }
                        catch {
                            $errorMessage = "Login action failed.`n`nCould not find or execute the 'LoginSelectedRow' function. The 'login.psm1' module may have failed to load correctly.`n`nTechnical Details: $($_.Exception.Message)"
                            try { Show-ErrorDialog -Message $errorMessage }
                            catch { [System.Windows.Forms.MessageBox]::Show($errorMessage, 'Login Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
                        }
                    }
                }

                Ftool = @{ Click = { if(Get-Command FtoolSelectedRow -EA 0){ $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { FtoolSelectedRow $_ } } } }
                Exit = @{ Click = { if([System.Windows.Forms.MessageBox]::Show("Terminate selected?","Confirm",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes'){ $global:DashboardConfig.UI.DataGridFiller.SelectedRows | ForEach-Object { Stop-Process -Id $_.Tag.Id -Force -EA 0 } } } }
                CopyNeuz = @{ Click = {
                    $confirm = [System.Windows.Forms.MessageBox]::Show("This will copy neuz.exe from your main client's bin32 and bin64 folders to all your defined profiles. Profiles that use a custom or different neuz.exe might break. Proceed?", "Confirm Neuz.exe Copy", "YesNo", "Warning")

                    if ($confirm -eq 'Yes') {
                        $mainLauncherPath = $global:DashboardConfig.UI.InputLauncher.Text
                        if ([string]::IsNullOrWhiteSpace($mainLauncherPath) -or -not (Test-Path $mainLauncherPath)) {
                            [System.Windows.Forms.MessageBox]::Show("Please select a valid Main Launcher Path first in the settings.", "Error", "OK", "Error")
                            return
                        }

                        $mainClientDir = Split-Path $mainLauncherPath -Parent
                        $sourceBin32 = Join-Path $mainClientDir "bin32\neuz.exe"
                        $sourceBin64 = Join-Path $mainClientDir "bin64\neuz.exe"

                        $profiles = $global:DashboardConfig.Config['Profiles']
                        if (-not $profiles -or $profiles.Count -eq 0) {
                            [System.Windows.Forms.MessageBox]::Show("No profiles defined. Please create or add profiles first.", "Information", "OK", "Information")
                            return
                        }

                        $copiedCount = 0
                        $failedCopies = @()

                        foreach ($profileName in $profiles.Keys) {
                            $profilePath = $profiles[$profileName]

                            $targetBin32 = Join-Path $profilePath "bin32\neuz.exe"
                            $targetBin64 = Join-Path $profilePath "bin64\neuz.exe"

                            try {
                                if (Test-Path $sourceBin32) {
                                    if (-not (Test-Path (Split-Path $targetBin32 -Parent))) {
                                        New-Item -ItemType Directory -Path (Split-Path $targetBin32 -Parent) -Force | Out-Null
                                    }
                                    Copy-Item -Path $sourceBin32 -Destination $targetBin32 -Force -ErrorAction Stop
                                    $copiedCount++
                                }
                                if (Test-Path $sourceBin64) {
                                    if (-not (Test-Path (Split-Path $targetBin64 -Parent))) {
                                        New-Item -ItemType Directory -Path (Split-Path $targetBin64 -Parent) -Force | Out-Null
                                    }
                                    Copy-Item -Path $sourceBin64 -Destination $targetBin64 -Force -ErrorAction Stop
                                    $copiedCount++
                                }
                            }
                            catch {
                                $failedCopies += "Profile '$profileName': $($_.Exception.Message)"
                            }
                        }

                        if ($failedCopies.Count -eq 0) {
                            [System.Windows.Forms.MessageBox]::Show("Successfully copied neuz.exe to $copiedCount locations.", "Success", "OK", "Information")
                        } else {
                            [System.Windows.Forms.MessageBox]::Show("Copied neuz.exe to $($copiedCount - $failedCopies.Count) locations. Failed to copy to the following profiles:`n" + ($failedCopies -join "`n"), "Partial Success/Error", "OK", "Warning")
                        }
                    }
                }
            }
		}

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

function Show-SettingsForm
    {

        [CmdletBinding()]
        param()

            if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
                ($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
            {
                return
            }

            if (-not ($global:DashboardConfig.UI -and $global:DashboardConfig.UI.SettingsForm -and $global:DashboardConfig.UI.MainForm))
            {
                return
            }

        $settingsForm = $global:DashboardConfig.UI.SettingsForm

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
    }

function Hide-SettingsForm
    {

        [CmdletBinding()]
        param()

            if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
                ($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
            {
                return
            }

            if (-not ($global:DashboardConfig.UI -and $global:DashboardConfig.UI.SettingsForm))
            {
                return
            }

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
    }

function Set-UIElement
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [ValidateSet('Form', 'Panel', 'Button', 'Label', 'DataGridView', 'TextBox', 'ComboBox', 'Custom.DarkComboBox', 'CheckBox', 'Toggle', 'ProgressBar', 'TextProgressBar')]
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

            $el = switch ($type)
            {
                'Form'            { New-Object System.Windows.Forms.Form }
                'Panel'           { New-Object System.Windows.Forms.Panel }
                'Button'          { New-Object System.Windows.Forms.Button }
                'Label'           { New-Object System.Windows.Forms.Label }
                'DataGridView'    { New-Object System.Windows.Forms.DataGridView }
                'TextBox'         { New-Object System.Windows.Forms.TextBox }
                'ComboBox'        { New-Object System.Windows.Forms.ComboBox }
                'Custom.DarkComboBox' { New-Object Custom.DarkComboBox }
                'CheckBox'        { New-Object System.Windows.Forms.CheckBox }
                'Toggle'          { New-Object Custom.Toggle }
                'ProgressBar'     { New-Object System.Windows.Forms.ProgressBar }
                'TextProgressBar' { New-Object Custom.TextProgressBar }
                default           { throw "Invalid element type specified: $type" }
            }

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
                        $el.FlatStyle = $fs
                    }
                }
                'Custom.DarkComboBox' {
                    if ($null -ne $dropDownStyle) { try { $el.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::$dropDownStyle } catch {} }
                    if ($null -ne $fs) {
                        $el.FlatStyle = $fs
                    }
                    $el.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
                    $el.IntegralHeight = $false
                    $el.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
                    $el.ForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
                    if ($PSBoundParameters.ContainsKey('font')) { $el.Font = $font }
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


        if ($PSBoundParameters.ContainsKey('tooltip') -and $tooltip -ne $null -and $global:DashboardConfig.UI.ToolTip) {
            $global:DashboardConfig.UI.ToolTip.SetToolTip($el, $tooltip)
        }

        return $el
    }


#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion