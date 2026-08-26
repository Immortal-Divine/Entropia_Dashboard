<# ui.psm1 #>

#region Helper Functions

function Show-DarkMessageBox
{
    [CmdletBinding(DefaultParameterSetName='ImplicitOwner')]
    param(
        [Parameter(ParameterSetName='ExplicitOwner', Mandatory=$true)]
        [System.Windows.Forms.IWin32Window]$Owner,

        [Parameter(ParameterSetName='ImplicitOwner', Position=0, Mandatory=$true)]
        [Parameter(ParameterSetName='ExplicitOwner', Position=0, Mandatory=$true)]
        [string]$Text,

        [Parameter(ParameterSetName='ImplicitOwner', Position=1)]
        [Parameter(ParameterSetName='ExplicitOwner', Position=1)]
        [string]$Caption = "Notification",

        [Parameter(ParameterSetName='ImplicitOwner', Position=2)]
        [Parameter(ParameterSetName='ExplicitOwner', Position=2)]
        [System.Windows.Forms.MessageBoxButtons]$Buttons = 'OK',

        [Parameter(ParameterSetName='ImplicitOwner', Position=3)]
        [Parameter(ParameterSetName='ExplicitOwner', Position=3)]
        [System.Windows.Forms.MessageBoxIcon]$Icon = 'Information',

        [Parameter(ParameterSetName='ImplicitOwner', Position=4)]
        [Parameter(ParameterSetName='ExplicitOwner', Position=4)]
        [string]$Type,

        [Parameter(ParameterSetName='ImplicitOwner', Position=5)]
        [Parameter(ParameterSetName='ExplicitOwner', Position=5)]
        [switch]$TopMost
    )

    
    $global:DashboardConfig.State.LoginActive = $true
    try
    {
        $isSuccess = $false
        if ($PSBoundParameters.ContainsKey('Type')) {
            $isSuccess = ($Type.ToLower() -eq "success")
        }

        $ownerForm = $null
        if ($PSCmdlet.ParameterSetName -eq 'ExplicitOwner') {
            $ownerForm = $PSBoundParameters['Owner']
        } else {
            $ownerForm = $global:DashboardConfig.UI.MainForm
        }
        
        
        $msgBox = New-Object Custom.DarkMessageBox($Text, $Caption, $Buttons, $Icon, $isSuccess)
        
        if ($TopMost) { $msgBox.TopMost = $true }

        if ($ownerForm -and -not $ownerForm.IsDisposed)
        {
            $msgBox.Owner = $ownerForm
            $msgBox.StartPosition = 'CenterScreen'
        }

        
        return Show-FormAsDialog -Form $msgBox
    }
    finally
    {
        $global:DashboardConfig.State.LoginActive = $false
    }
}

function Show-FormAsDialog
{
	param(
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.Form]$Form
	)

	$Form.Show()
	try
	{
		if (-not ('Custom.Native' -as [Type])) {
			throw "Custom.Native class not found. Cannot use P/Invoke message loop."
		}
				
		$msg = New-Object Custom.Native+MSG
		while ($Form.Visible)
		{
			$null = [Custom.Native]::AsyncExecution(0, [IntPtr[]]@(), $false, 50, [Custom.Native]::QS_ALLINPUT)
			while ([Custom.Native]::PeekMessage([ref]$msg, [IntPtr]::Zero, 0, 0, [Custom.Native]::PM_REMOVE))
			{
				[void][Custom.Native]::TranslateMessage([ref]$msg)
				[void][Custom.Native]::DispatchMessage([ref]$msg)
			}
		}
	}
	catch { while ($Form.Visible) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 20 } }
	return $Form.DialogResult
}
function ShowInputBox
{
	param(
		[string]$Title,
		[string]$Prompt,
		[string]$DefaultText
	)

	$form = New-Object Custom.DarkInputBox($Title, $Prompt, $DefaultText)

	if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon))
	{
		try { $form.Icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon) } catch {}
	}

	if ($global:DashboardConfig.UI.SettingsForm)
	{
		$form.Owner = $global:DashboardConfig.UI.SettingsForm
	}
	$result = Show-FormAsDialog -Form $form

	if ($result -eq [System.Windows.Forms.DialogResult]::OK)
	{
		return $form.ResultText
	}
 else
	{
		return $null
	}
}

function Register-ButtonFadeEffect {
	param(
		[System.Windows.Forms.Button]$Button,
		[System.Drawing.Color]$NormalStartColor,
		[System.Drawing.Color]$NormalEndColor,
		[System.Drawing.Color]$HoverStartColor,
		[System.Drawing.Color]$HoverEndColor,
		[System.Drawing.Color]$LeftPressedStartColor,
		[System.Drawing.Color]$LeftPressedEndColor,
		[System.Drawing.Color]$RightPressedStartColor,
		[System.Drawing.Color]$RightPressedEndColor
	)

	$state = [PSCustomObject]@{
		TargetStartColor = $NormalStartColor
		TargetEndColor   = $NormalEndColor
		CurrentStartR    = [double]$NormalStartColor.R
		CurrentStartG    = [double]$NormalStartColor.G
		CurrentStartB    = [double]$NormalStartColor.B
		CurrentEndR      = [double]$NormalEndColor.R
		CurrentEndG      = [double]$NormalEndColor.G
		CurrentEndB      = [double]$NormalEndColor.B
	}

	$timer = New-Object System.Windows.Forms.Timer
	$timer.Interval = 15

	$timer.Add_Tick({
			$targetStart = $state.TargetStartColor
			$targetEnd = $state.TargetEndColor

			$diffStartR = $targetStart.R - $state.CurrentStartR
			$diffStartG = $targetStart.G - $state.CurrentStartG
			$diffStartB = $targetStart.B - $state.CurrentStartB

			$diffEndR = $targetEnd.R - $state.CurrentEndR
			$diffEndG = $targetEnd.G - $state.CurrentEndG
			$diffEndB = $targetEnd.B - $state.CurrentEndB

			$step = 0.15
			$close = $true

			if ([Math]::Abs($diffStartR) -lt 1 -and [Math]::Abs($diffStartG) -lt 1 -and [Math]::Abs($diffStartB) -lt 1) {
				$state.CurrentStartR = $targetStart.R
				$state.CurrentStartG = $targetStart.G
				$state.CurrentStartB = $targetStart.B
			}
			else {
				$state.CurrentStartR += $diffStartR * $step
				$state.CurrentStartG += $diffStartG * $step
				$state.CurrentStartB += $diffStartB * $step
				$close = $false
			}

			if ([Math]::Abs($diffEndR) -lt 1 -and [Math]::Abs($diffEndG) -lt 1 -and [Math]::Abs($diffEndB) -lt 1) {
				$state.CurrentEndR = $targetEnd.R
				$state.CurrentEndG = $targetEnd.G
				$state.CurrentEndB = $targetEnd.B
			}
			else {
				$state.CurrentEndR += $diffEndR * $step
				$state.CurrentEndG += $diffEndG * $step
				$state.CurrentEndB += $diffEndB * $step
				$close = $false
			}

			$Button.Invalidate()
			if ($close) {
				$timer.Stop()
			}
		}.GetNewClosure())

	$Button.Add_MouseEnter({
			$state.TargetStartColor = $HoverStartColor
			$state.TargetEndColor = $HoverEndColor
			$timer.Start()
		}.GetNewClosure())

	$Button.Add_MouseLeave({
			$state.TargetStartColor = $NormalStartColor
			$state.TargetEndColor = $NormalEndColor
			$timer.Start()
		}.GetNewClosure())

	$Button.Add_MouseDown({
			if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
				$state.TargetStartColor = $LeftPressedStartColor
				$state.TargetEndColor = $LeftPressedEndColor
				$timer.Start()
			}
			elseif ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
				$state.TargetStartColor = $RightPressedStartColor
				$state.TargetEndColor = $RightPressedEndColor
				$timer.Start()
			}
		}.GetNewClosure())

	$Button.Add_MouseUp({
			if ($Button.ClientRectangle.Contains($Button.PointToClient([System.Windows.Forms.Control]::MousePosition))) {
				$state.TargetStartColor = $HoverStartColor
				$state.TargetEndColor = $HoverEndColor
			}
			else {
				$state.TargetStartColor = $NormalStartColor
				$state.TargetEndColor = $NormalEndColor
			}
			$timer.Start()
		}.GetNewClosure())

	$Button.Add_Paint({
			param($src, $e)
			$rect = [System.Drawing.Rectangle]::new(0, 0, $src.Width, $src.Height)
			$color1 = [System.Drawing.Color]::FromArgb([int]$state.CurrentStartR, [int]$state.CurrentStartG, [int]$state.CurrentStartB)
			$color2 = [System.Drawing.Color]::FromArgb([int]$state.CurrentEndR, [int]$state.CurrentEndG, [int]$state.CurrentEndB)
			
			$bgBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new($rect, $color1, $color2, [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
			$e.Graphics.FillRectangle($bgBrush, $rect)
			$bgBrush.Dispose()

			$textBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(240, 240, 240))
			$textFormat = [System.Drawing.StringFormat]::new()
			$textFormat.Alignment = [System.Drawing.StringAlignment]::Center
			$textFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
			$e.Graphics.DrawString($src.Text, $src.Font, $textBrush, [System.Drawing.RectangleF]::new(0, 0, $src.Width, $src.Height), $textFormat)
			$textBrush.Dispose()
			$textFormat.Dispose()

			$borderPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(60, 60, 60))
			$e.Graphics.DrawRectangle($borderPen, 0, 0, $src.Width - 1, $src.Height - 1)
			$borderPen.Dispose()
		}.GetNewClosure())
}

function ShowSettingsForm
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
		
		if ($settingsForm.StartPosition -ne [System.Windows.Forms.FormStartPosition]::Manual)
		{
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
			$settingsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
		}
		
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

function HideSettingsForm
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
				
				if ($global:DashboardConfig.UI.SettingsForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
					if (-not $global:DashboardConfig.Config.Contains('WindowPosition')) { $global:DashboardConfig.Config['WindowPosition'] = [ordered]@{} }
					$global:DashboardConfig.Config['WindowPosition']['SettingsFormX'] = $global:DashboardConfig.UI.SettingsForm.Location.X
					$global:DashboardConfig.Config['WindowPosition']['SettingsFormY'] = $global:DashboardConfig.UI.SettingsForm.Location.Y
					if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
				}

				$global:DashboardConfig.UI.SettingsForm.Hide()
				$global:DashboardConfig.UI.MainForm.Show()
			}
		})
	$global:fadeOutTimer.Start()
	$global:DashboardConfig.Resources.Timers['fadeOutTimer'] = $global:fadeOutTimer
}

function ShowExtraForm
{
	[CmdletBinding()]
	param()

	if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
		($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
	{
		return
	}

	if (-not ($global:DashboardConfig.UI -and $global:DashboardConfig.UI.ExtraForm -and $global:DashboardConfig.UI.MainForm))
	{
		return
	}

	$ExtraForm = $global:DashboardConfig.UI.ExtraForm

	if ($ExtraForm.Opacity -lt 0.95)
	{
		$ExtraForm.Visible = $true
		
		if ($ExtraForm.StartPosition -ne [System.Windows.Forms.FormStartPosition]::Manual)
		{
			$mainFormLocation = $global:DashboardConfig.UI.MainForm.Location
			$ExtraFormWidth = $ExtraForm.Width
			$ExtraFormHeight = $ExtraForm.Height
			$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
			$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

			$x = $mainFormLocation.X + (($global:DashboardConfig.UI.MainForm.Width - $ExtraFormWidth) / 2)
			$y = $mainFormLocation.Y + (($global:DashboardConfig.UI.MainForm.Height - $ExtraFormHeight) / 2)

			$margin = 0
			$x = [Math]::Max($margin, [Math]::Min($x, $screenWidth - $ExtraFormWidth - $margin))
			$y = [Math]::Max($margin, [Math]::Min($y, $screenHeight - $ExtraFormHeight - $margin))

			$ExtraForm.Location = New-Object System.Drawing.Point($x, $y)
			$ExtraForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
		}
		
		$ExtraForm.BringToFront()
		$ExtraForm.Activate()
	}

	if ($script:fadeInTimer) { $script:fadeInTimer.Dispose() }
	$script:fadeInTimer = New-Object System.Windows.Forms.Timer
	$script:fadeInTimer.Interval = 15
	$script:fadeInTimer.Add_Tick({
			if (-not $global:DashboardConfig.UI.ExtraForm -or $global:DashboardConfig.UI.ExtraForm.IsDisposed)
			{
				$script:fadeInTimer.Stop()
				$script:fadeInTimer.Dispose()
				$script:fadeInTimer = $null
				return
			}
			if ($global:DashboardConfig.UI.ExtraForm.Opacity -lt 1)
			{
				$global:DashboardConfig.UI.ExtraForm.Opacity += 0.1
			}
			else
			{
				$global:DashboardConfig.UI.ExtraForm.Opacity = 1
				$script:fadeInTimer.Stop()
				$script:fadeInTimer.Dispose()
				$script:fadeInTimer = $null
			}
		})
	$script:fadeInTimer.Start()
	$global:DashboardConfig.Resources.Timers['fadeInTimer'] = $script:fadeInTimer
}

function HideExtraForm
{

	[CmdletBinding()]
	param()

	if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
		($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
	{
		return
	}

	if (-not ($global:DashboardConfig.UI -and $global:DashboardConfig.UI.ExtraForm))
	{
		return
	}

	if ($global:DashboardConfig.State['ChestHotkeyId'])
	{
		try
		{
			if (Get-Command Stop-ChestOpener -ErrorAction SilentlyContinue -Verbose:$False) { Stop-ChestOpener }
		}
		catch {}
		try
		{
			if (Get-Command UnregisterHotkeyInstance -ErrorAction SilentlyContinue -Verbose:$False)
			{
				UnregisterHotkeyInstance -Id $global:DashboardConfig.State['ChestHotkeyId'] -OwnerKey 'ChestOpener'
			}
		}
		catch {}
		$global:DashboardConfig.State['ChestHotkeyId'] = $null
	}

	if ($global:fadeOutTimer) { $global:fadeOutTimer.Dispose() }
	$global:fadeOutTimer = New-Object System.Windows.Forms.Timer
	$global:fadeOutTimer.Interval = 15
	$global:fadeOutTimer.Add_Tick({
			if (-not $global:DashboardConfig.UI.ExtraForm -or $global:DashboardConfig.UI.ExtraForm.IsDisposed)
			{
				$global:fadeOutTimer.Stop()
				$global:fadeOutTimer.Dispose()
				$global:fadeOutTimer = $null
				return
			}

			if ($global:DashboardConfig.UI.ExtraForm.Opacity -gt 0)
			{
				$global:DashboardConfig.UI.ExtraForm.Opacity -= 0.1
			}
			else
			{
				$global:DashboardConfig.UI.ExtraForm.Opacity = 0
				$global:fadeOutTimer.Stop()
				$global:fadeOutTimer.Dispose()
				$global:fadeOutTimer = $null
				
				if ($global:DashboardConfig.UI.ExtraForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
					if (-not $global:DashboardConfig.Config.Contains('WindowPosition')) { $global:DashboardConfig.Config['WindowPosition'] = [ordered]@{} }
					$global:DashboardConfig.Config['WindowPosition']['ExtraFormX'] = $global:DashboardConfig.UI.ExtraForm.Location.X
					$global:DashboardConfig.Config['WindowPosition']['ExtraFormY'] = $global:DashboardConfig.UI.ExtraForm.Location.Y
					if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
				}

				$global:DashboardConfig.UI.ExtraForm.Hide()
				$global:DashboardConfig.UI.MainForm.Show()
			}
		})
	$global:fadeOutTimer.Start()
	$global:DashboardConfig.Resources.Timers['fadeOutTimer'] = $global:fadeOutTimer

}

function RefreshLoginProfileSelector
{
	[CmdletBinding()]
	param()

	try
	{
		Write-Verbose '  UI: Refreshing LoginProfileSelector...'
		$UI = $global:DashboardConfig.UI
		if (-not $UI -or -not $UI.LoginProfileSelector) { return }

		$selectedItem = $UI.LoginProfileSelector.SelectedItem
		$UI.LoginProfileSelector.Items.Clear()
		$UI.LoginProfileSelector.Items.Add('Default') | Out-Null

		if ($global:DashboardConfig.Config['Profiles'])
		{
			foreach ($key in $global:DashboardConfig.Config['Profiles'].Keys)
			{
				$UI.LoginProfileSelector.Items.Add($key) | Out-Null
			}
		}

		if ($selectedItem -and $UI.LoginProfileSelector.Items.Contains($selectedItem))
		{
			$UI.LoginProfileSelector.SelectedItem = $selectedItem
		}
		else
		{
			$UI.LoginProfileSelector.SelectedIndex = 0
		}

		Write-Verbose '  UI: LoginProfileSelector refreshed.'
	}
	catch
	{
		Write-Verbose "  UI: Failed to refresh LoginProfileSelector: $_"
	}
}


function RegisterConfiguredHotkeys
{
	[CmdletBinding()]
	param()

	Write-Verbose 'HOTKEYS: Registering all configured hotkeys (groups and windows)...'
	
	if (-not (Get-Command SetHotkey -ErrorAction SilentlyContinue -Verbose:$False))
	{
		Write-Verbose '  HOTKEYS: SetHotkey command not found. Skipping hotkey registration.'
		return
	}
	
	$config = $global:DashboardConfig.Config
	if (-not $config.Contains('Hotkeys')) { return }

	$hotkeys = $config['Hotkeys']
    $groups = if ($config.Contains('HotkeyGroups')) { $config['HotkeyGroups'] } else { @{} }

	foreach ($hotkeyName in $hotkeys.Keys)
	{
		try
		{
			$keyCombo = $hotkeys[$hotkeyName]

			if ([string]::IsNullOrWhiteSpace($keyCombo) -or $keyCombo -eq 'none')
			{
				continue
			}

			
			if ($groups.Contains($hotkeyName))
			{
				
				$memberString = $groups[$hotkeyName]
				if ([string]::IsNullOrWhiteSpace($memberString))
				{
					continue
				}

				$memberList = $memberString -split ','
				$ownerKey = "GroupHotkey_$hotkeyName"

				SetHotkey -KeyCombinationString $keyCombo -OwnerKey $ownerKey -Action ({
						$targetMembers = $memberList
						$grid = $global:DashboardConfig.UI.DataGridMain
						if (-not $grid) { return }

						foreach ($row in $grid.Rows)
						{
							$identity = Get-RowIdentity -Row $row
							if ($targetMembers -contains $identity)
							{
								if ($row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
								{
									$h = $row.Tag.MainWindowHandle
									SetWindowToolStyle -hWnd $h -Hide $false
									[Custom.Native]::BringToFront($h)
									SetWindowToolStyle -hWnd $h -Hide $false
								}
							}
						}
					}.GetNewClosure())
				Write-Verbose "  HOTKEYS: Registered GROUP hotkey '$keyCombo' for group '$hotkeyName'."
			}
			else
			{
				$windowTitle = $hotkeyName
				$ownerKey = $windowTitle
				
				$actionScript = "[Custom.Native]::BringToFront((Get-Process | Where-Object { `$_.MainWindowTitle -eq '$($windowTitle.Replace("'", "''"))' } | Select-Object -First 1).MainWindowHandle)"
				$action = [scriptblock]::Create($actionScript)
				SetHotkey -KeyCombinationString $keyCombo -Action $action -OwnerKey $ownerKey
				Write-Verbose "  HOTKEYS: Registered WINDOW hotkey '$keyCombo' for window '$windowTitle'."
			}
		}
		catch
		{
			Write-Verbose "  HOTKEYS: Failed to register hotkey '$keyCombo' for '$hotkeyName'. Error: $($_.Exception.Message)"
		}
	}

	if ($config.Contains('HotkeysHide'))
	{
		$hotkeysHide = $config['HotkeysHide']
		$groupsHide = if ($config.Contains('HotkeyGroupsHide')) { $config['HotkeyGroupsHide'] } else { @{} }
		foreach ($hotkeyName in $hotkeysHide.Keys)
		{
			try
			{
				$keyCombo = $hotkeysHide[$hotkeyName]
				if ([string]::IsNullOrWhiteSpace($keyCombo) -or $keyCombo -eq 'none') { continue }

				if ($groupsHide.Contains($hotkeyName))
				{
					$memberString = $groupsHide[$hotkeyName]
					if ([string]::IsNullOrWhiteSpace($memberString)) { continue }

					$memberList = $memberString -split ','
					$ownerKey = "GroupHotkeyHide_$hotkeyName"

					SetHotkey -KeyCombinationString $keyCombo -OwnerKey $ownerKey -Action ({
						$targetMembers = $memberList
						$grid = $global:DashboardConfig.UI.DataGridMain
						if (-not $grid) { return }

						foreach ($row in $grid.Rows)
						{
							$identity = Get-RowIdentity -Row $row
							if ($targetMembers -contains $identity)
							{
								if ($row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
								{
									$h = $row.Tag.MainWindowHandle
									try { [Custom.Native]::ShowWindow($h, [Custom.Native]::SW_MINIMIZE); [Custom.Native]::EmptyWorkingSet($row.Tag.Id) } catch {}
								}
							}
						}
					}.GetNewClosure())
					Write-Verbose "  HOTKEYS: Registered GROUP hide hotkey '$keyCombo' for group '$hotkeyName'."
				}
			}
			catch
			{
				Write-Verbose "  HOTKEYS: Failed to register hide hotkey '$keyCombo' for '$hotkeyName'. Error: $($_.Exception.Message)"
			}
		}
	}

	if ($config.Contains('HotkeysLoadPos'))
	{
		$hotkeysLoad = $config['HotkeysLoadPos']
		$groupsLoad = if ($config.Contains('HotkeyGroupsLoadPos')) { $config['HotkeyGroupsLoadPos'] } else { @{} }
		foreach ($hotkeyName in $hotkeysLoad.Keys)
		{
			try
			{
				$keyCombo = $hotkeysLoad[$hotkeyName]
				if ([string]::IsNullOrWhiteSpace($keyCombo) -or $keyCombo -eq 'none') { continue }

				if ($groupsLoad.Contains($hotkeyName))
				{
					$memberString = $groupsLoad[$hotkeyName]
					if ([string]::IsNullOrWhiteSpace($memberString)) { continue }

					$memberList = $memberString -split ','
					$ownerKey = "GroupHotkeyLoadPos_$hotkeyName"

					SetHotkey -KeyCombinationString $keyCombo -OwnerKey $ownerKey -Action ({
						$targetMembers = $memberList
						$grid = $global:DashboardConfig.UI.DataGridMain
						if (-not $grid) { return }

						foreach ($row in $grid.Rows)
						{
							$identity = Get-RowIdentity -Row $row
							if ($targetMembers -contains $identity)
							{
								if ($global:DashboardConfig.Config.Contains('SavedWindowPositions') -and $global:DashboardConfig.Config['SavedWindowPositions'].Contains($identity) -and $row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
								{
									$pos = $global:DashboardConfig.Config['SavedWindowPositions'][$identity] -split ','
									if ($pos.Count -eq 4)
									{
										if (-not [Custom.Native]::IsMinimized($row.Tag.MainWindowHandle))
										{
											[Custom.Native]::PositionWindow($row.Tag.MainWindowHandle, [IntPtr]::Zero, [int]$pos[0], [int]$pos[1], [int]$pos[2], [int]$pos[3], 0x0014)
										}
									}
								}
							}
						}
					}.GetNewClosure())
					Write-Verbose "  HOTKEYS: Registered GROUP load-pos hotkey '$keyCombo' for group '$hotkeyName'."
				}
			}
			catch
			{
				Write-Verbose "  HOTKEYS: Failed to register load-pos hotkey '$keyCombo' for '$hotkeyName'. Error: $($_.Exception.Message)"
			}
		}
	}

	if ($config.Contains('GlobalHotkeys'))
	{
		$globalHks = $config['GlobalHotkeys']
		
		if (-not $globalHks) { return }

		$masterPauseKey = if ($globalHks.Contains('MasterPauseHotkey'))  { $globalHks['MasterPauseHotkey']  } else { $null }
		$masterResumeKey = if ($globalHks.Contains('MasterResumeHotkey')) { $globalHks['MasterResumeHotkey'] } else { $null }

		if (-not [string]::IsNullOrWhiteSpace($masterPauseKey) -and $masterPauseKey -ne 'none')
		{
			try
			{
				SetHotkey -KeyCombinationString $masterPauseKey -OwnerKey 'master_pause_hotkey' -OwnerLabel 'Pause All Ftool/Macro Hotkeys' -Action { if (Get-Command PauseMasterHotkeys -ErrorAction SilentlyContinue -Verbose:$False) { PauseMasterHotkeys } }
				Write-Verbose "  HOTKEYS: Registered MASTER PAUSE hotkey '$masterPauseKey'."
			}
			catch { Write-Verbose "  HOTKEYS: Failed to register master pause hotkey: $($_.Exception.Message)" }
		}

		if (-not [string]::IsNullOrWhiteSpace($masterResumeKey) -and $masterResumeKey -ne 'none')
		{
			try
			{
				SetHotkey -KeyCombinationString $masterResumeKey -OwnerKey 'master_resume_hotkey' -OwnerLabel 'Resume All Ftool/Macro Hotkeys' -Action { if (Get-Command ResumeMasterHotkeys -ErrorAction SilentlyContinue -Verbose:$False) { ResumeMasterHotkeys } }
				Write-Verbose "  HOTKEYS: Registered MASTER RESUME hotkey '$masterResumeKey'."
			}
			catch { Write-Verbose "  HOTKEYS: Failed to register master resume hotkey: $($_.Exception.Message)" }
		}
	}
}

function RefreshNotificationGrid
{
	if ($global:DashboardConfig.UI.NotificationGrid)
	{
		$grid = $global:DashboardConfig.UI.NotificationGrid
		if ($grid.InvokeRequired)
		{
			$grid.Invoke([Action]{ RefreshNotificationGrid })
			return
		}
		$grid.SuspendLayout()
		try
		{
			$grid.Rows.Clear()
			if ($global:DashboardConfig.State.NotificationHistory)
			{
				foreach ($entry in $global:DashboardConfig.State.NotificationHistory)
				{
					$timeStr = $entry.Timestamp.ToString('HH:mm:ss')
					$details = "$($entry.Title): $($entry.Message -replace "`r`n", " ")"
					$idx = $grid.Rows.Add($timeStr, $entry.Type, $details)
					if ($idx -ge 0)
					{
						$row = $grid.Rows[$idx]
						$row.Tag = $entry
						$color = switch ($entry.Type) { 'Warning' { [System.Drawing.Color]::Orange } 'Error' { [System.Drawing.Color]::IndianRed } default { [System.Drawing.Color]::CornflowerBlue } }
						$row.Cells[1].Style.ForeColor = $color
					}
				}
			}
		}
		catch { Write-Verbose "RefreshNotificationGrid Error: $_" }
		finally { $grid.ResumeLayout() }
	}
}

function RefreshHotkeysList
{
	[CmdletBinding()]
	param()

	try
	{
		$grid = $global:DashboardConfig.UI.HotkeysGrid
		if ($grid -and $grid.InvokeRequired)
		{
			$grid.Invoke([Action]{ RefreshHotkeysList })
			return
		}
		if (-not $grid) { return }
		$grid.Rows.Clear()

		if ($grid.Columns.Count -lt 3) {
			Write-Verbose "UI: RefreshHotkeysList aborted because HotkeysGrid has $($grid.Columns.Count) columns."
			return
		}

		$AddHotkeysGridRow = {
			param($rowValues)
			if (-not $rowValues -or $rowValues.Count -gt $grid.Columns.Count) { return -1 }
			$rowIndex = $grid.Rows.Add()
			if ($rowIndex -lt 0) { return -1 }
			for ($i = 0; $i -lt $rowValues.Count; $i++) {
				try { $grid.Rows[$rowIndex].Cells[$i].Value = $rowValues[$i] } catch {}
			}
			return $rowIndex
		}

		
		$allHotkeys = @{}
		if ($global:RegisteredHotkeys)
		{
			foreach ($id in $global:RegisteredHotkeys.Keys)
			{
				$allHotkeys[$id] = $global:RegisteredHotkeys[$id]
			}
		}
		if ($global:PausedRegisteredHotkeys)
		{
			foreach ($id in $global:PausedRegisteredHotkeys.Keys)
			{
				
				$allHotkeys[$id] = $global:PausedRegisteredHotkeys[$id]
			}
		}

		if ($allHotkeys.Count -gt 0)
		{



			$GetDecoratedTitle = {
					param($pidStr)
					$title = ''
					$g = $global:DashboardConfig.UI.DataGridMain
					if ($g)
					{
						$r = $g.Rows | Where-Object { $null -ne $_.Cells -and $_.Cells.Count -gt 2 -and $null -ne $_.Cells[2].Value -and $_.Cells[2].Value.ToString() -eq $pidStr } | Select-Object -First 1
						if ($r -and $null -ne $r.Cells -and $r.Cells.Count -gt 1) { $title = $r.Cells[1].Value }
					}
					return $title
				}

				$groups = [ordered]@{
					'Master' = @()
					'Ftool' = @()
					'Macro' = @()
					'Windows' = @()
					'Groups' = @()
					'QuickCommands' = @()
				}

			$sortedKeys = $allHotkeys.Keys | Sort-Object { $allHotkeys[$_].KeyString }


			foreach ($id in $sortedKeys)
				{
					$meta = $allHotkeys[$id]
					$keyString = $meta.KeyString
					$isPaused = $id -lt 0
					if ($meta.Owners)
					{
						foreach ($ownerKey in $meta.Owners.Keys)
						{
							if (-not $ownerKey) { continue }
							
							$displayOwner = $ownerKey; $displayAction = ''
							if ($ownerKey -eq 'master_pause_hotkey' -or $ownerKey -eq 'master_resume_hotkey')
							{
								$displayOwner = "Master Hotkey"
								$displayAction = if ($ownerKey -eq 'master_pause_hotkey') { 'Pause All Ftool/Macro Hotkeys' } else { 'Resume All Ftool/Macro Hotkeys' }
								$groupName = 'Master'
							}
							elseif ($ownerKey -match '^global_toggle_(.+)')
							{
								$instanceId = $Matches[1]
								$displayOwner = "Ftool Global Toggle: $instanceId"
								$displayAction = 'Enable/Disable hotkeys'
								$decoratedTitle = &$GetDecoratedTitle $instanceId
								if (-not [string]::IsNullOrEmpty($decoratedTitle)) { $displayOwner = "Ftool Global Toggle: $decoratedTitle" }
								$groupName = 'Ftool'
							}
							elseif ($ownerKey -match '^ext_(.+)_\d+')
							{
								$instanceId = $Matches[1]
								$extKey = $ownerKey
								$displayOwner = "Ftool Extension: $instanceId"
								$displayAction = 'Start/Stop'
								if ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
								{
									$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
									if ($form -and $form.Tag)
									{
										$decoratedTitle = &$GetDecoratedTitle $instanceId
										if ($extKey -and $global:DashboardConfig.Resources.ExtensionData -and $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
										{
											$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
											$extName = if ($extData.Name -and $extData.Name.Text) { $extData.Name.Text } else { "Ext $($extData.ExtNum)" }
											$displayOwner = "Ftool $decoratedTitle - $extName"
											$spammedKey = if ($extData.BtnKeySelect -and $extData.BtnKeySelect.Text) { $extData.BtnKeySelect.Text } else { 'none' }
											$displayAction = "Ftool Key: $spammedKey"
										}
									}
								$groupName = 'Ftool'
								}
							}
							elseif ($ownerKey -match '^GroupHotkey(?:Hide|LoadPos)?_(.+)')
							{
								$group = $Matches[1]
								$displayOwner = "Window Group: $group"
								if ($ownerKey -match '^GroupHotkeyHide_') { $displayAction = 'Hide' }
								elseif ($ownerKey -match '^GroupHotkeyLoadPos_') { $displayAction = 'Position' }
								else { $displayAction = 'Show' }
								$groupName = 'Groups'
							}
							elseif ($ownerKey -match '^quickcmd_(.+)_(.+)')
							{
								$instanceId = $Matches[1]; $cmdId = $Matches[2]
								$pidPart = ($instanceId -split '_')[-1]
								$decoratedTitle = &$GetDecoratedTitle $pidPart
								if (-not [string]::IsNullOrEmpty($decoratedTitle)) { $displayOwner = "QuickCommand: $decoratedTitle" } else { $displayOwner = "QuickCommand: $instanceId" }
								$displayAction = "Invoke Command: $cmdId"
								$ownerEntry = $meta.Owners[$ownerKey]
								if ($ownerEntry -and $ownerEntry.PSObject.Properties['OwnerLabel'] -and -not [string]::IsNullOrEmpty($ownerEntry.OwnerLabel)) { $displayAction = $ownerEntry.OwnerLabel }
								$ownerEntry = $meta.Owners[$ownerKey]
								if ($ownerEntry -and $ownerEntry.PSObject.Properties['OwnerLabel'] -and -not [string]::IsNullOrEmpty($ownerEntry.OwnerLabel)) { $displayAction = $ownerEntry.OwnerLabel }
								if ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
								{
									$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
									if ($form -and $form.Tag)
									{
										$data = $form.Tag
										$cmd = $null
										if ($data.SavedCommands) { $cmd = $data.SavedCommands | Where-Object { $_.Id -eq $cmdId } | Select-Object -First 1 }
										if (-not $cmd -and $data.GlobalCommandList) { $cmd = $data.GlobalCommandList | Where-Object { $_.Id -eq $cmdId } | Select-Object -First 1 }
										if ($cmd)
										{
											if ($cmd.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($cmd.Name)) { $displayAction = $cmd.Name }
											elseif ($cmd.PSObject.Properties['CommandText'] -and -not [string]::IsNullOrEmpty($cmd.CommandText)) { $displayAction = $cmd.CommandText }
											else { $displayAction = "Invoke Command: $cmdId" }
										}
									}
								}
								$groupName = 'QuickCommands'
							}
							elseif ($ownerKey -match '^macroinst_(.+)')
							{
								$instanceId = $Matches[1]
								$pidPart = ($instanceId -split '_')[-1]
								$decoratedTitle = &$GetDecoratedTitle $pidPart
								$displayOwner = if (-not [string]::IsNullOrEmpty($decoratedTitle)) { "Macro Instance: $decoratedTitle" } else { "Macro Instance: $instanceId" }
								$displayAction = 'Toggle Macros'
								$groupName = 'Macro'
							}
							else
							{
								if ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Contains($ownerKey))
								{
									$instanceId = $ownerKey
									$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
									if ($form -and $form.Tag -and $form.Tag.Type -eq 'Macro')
									{
										$pidPart = ($instanceId -split '_')[-1]
										$decoratedTitle = &$GetDecoratedTitle $pidPart
										$displayOwner = if (-not [string]::IsNullOrEmpty($decoratedTitle)) { "Macro Instance: $decoratedTitle" } else { "Macro Instance: $instanceId" }
										$displayAction = 'Toggle Macros'
										$groupName = 'Macro'
									}
									else
									{
										$displayOwner = "Ftool Instance: $instanceId"
										$displayAction = 'Start/Stop'
										if ($form -and $form.Tag)
										{
											$decoratedTitle = &$GetDecoratedTitle $instanceId
											$ftoolName = if ($form.Tag.Name -and $form.Tag.Name.Text) { $form.Tag.Name.Text } else { 'Main' }
											$displayOwner = "Ftool $decoratedTitle - $ftoolName"
											$spammedKey = if ($form.Tag.BtnKeySelect -and $form.Tag.BtnKeySelect.Text) { $form.Tag.BtnKeySelect.Text } else { 'none' }
											$displayAction = "Ftool Key: $spammedKey"
										}
										$groupName = 'Ftool'
									}
								}
								else
								{
									if ($ownerKey -match '^Macro_.+')
									{
										$instanceId = $ownerKey
										$displayOwner = "Macro Instance: $instanceId"
										$decoratedTitle = &$GetDecoratedTitle $instanceId
										if (-not [string]::IsNullOrEmpty($decoratedTitle)) { $displayOwner = "Macro Instance: $decoratedTitle" }
										$displayAction = 'Toggle Macros'
										$groupName = 'Macro'
									}
									elseif ($ownerKey -match '^[0-9]+$')
									{
										$instanceId = $ownerKey
										$displayOwner = "Ftool Instance: $instanceId"
										$decoratedTitle = &$GetDecoratedTitle $instanceId
										if (-not [string]::IsNullOrEmpty($decoratedTitle)) { $displayOwner = "Ftool $decoratedTitle - Main" }
										$displayAction = 'Start/Stop'
										$groupName = 'Ftool'
									}
									elseif ($ownerKey -match '^[A-Za-z]+_\d+')
									{
										$instanceId = $ownerKey
										$displayOwner = "Ftool Instance: $instanceId"
										$displayAction = 'Start/Stop'
										try { $decoratedTitle = &$GetDecoratedTitle $instanceId } catch { $decoratedTitle = '' }
										if (-not [string]::IsNullOrEmpty($decoratedTitle)) { $displayOwner = "Ftool $decoratedTitle - Main" }
										$groupName = 'Ftool'
									}
									else
									{
										$groupName = 'Unknown'
									}
								}
							}
							$finalKeyString = if ($isPaused) { "$keyString (Paused)" } else { $keyString }
							$entry = [ordered]@{ Key = $finalKeyString; Owner = $displayOwner; Action = $displayAction; Id = $id; OwnerKey = $ownerKey }
							$groups[$groupName] += $entry
						}
					}
				}

				$renderOrder = @('Master','Groups','Ftool','Macro','QuickCommands','Unknown')
				foreach ($gName in $renderOrder)
				{
					$items = $groups[$gName]
					if ($items -and $items.Count -gt 0)
					{
						$headerText = "--- $gName Hotkeys ---"
						$rowIndex = & $AddHotkeysGridRow @('',$headerText,'')
						if ($rowIndex -ge 0) {
							$grid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(60,60,60)
							$grid.Rows[$rowIndex].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(200,200,200)
							$grid.Rows[$rowIndex].DefaultCellStyle.Font = (New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Bold))
							$grid.Rows[$rowIndex].Tag = @{ Header = $true }
							$grid.Rows[$rowIndex].ReadOnly = $true
							$grid.Rows[$rowIndex].DefaultCellStyle.SelectionBackColor = $grid.Rows[$rowIndex].DefaultCellStyle.BackColor
						}
						$sortedItems = $items | Sort-Object Key
						foreach ($entry in $sortedItems)
						{
							$rowIndex = & $AddHotkeysGridRow @($entry.Key, $entry.Owner, $entry.Action)
							if ($rowIndex -ge 0) {
								$grid.Rows[$rowIndex].Tag = @{ Id = $entry.Id; OwnerKey = $entry.OwnerKey }
							}
						}
					}
				}
			$grid.ClearSelection()
		}
	}
	catch { Write-Verbose "UI: Failed to refresh hotkeys list: $_" }
}

function Update-BossButtonImage
{
	param($btn)
	Add-Type -AssemblyName System.Drawing

	if (-not $btn -or $btn.IsDisposed) { return }

	$showImages = $true
	if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options'].Contains('ShowBossImages')) {
		$showImages = ([int]$global:DashboardConfig.Config['Options']['ShowBossImages']) -eq 1
	}

	$bName = $btn.Text
	$type = 'Normal'
	if ($btn.Tag -is [System.Collections.IDictionary]) {
		$bName = $btn.Tag['BossName']
		$type = $btn.Tag['Type']
	} elseif ($btn.Tag -is [string]) {
		$type = $btn.Tag
	}

	if ($showImages)
	{
		$bData = $global:DashboardConfig.Resources.BossData[$bName]
		if ($bData -and $bData.url) {
			$fileName = $bData.url.Split('/')[-1]
			$localPath = Join-Path $global:DashboardConfig.Paths.Bosses $fileName

			if (Test-Path $localPath)
			{
				$fileStream = $null
				$img = $null
				try {
					$fileStream = New-Object System.IO.FileStream($localPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
					$originalImage = [System.Drawing.Image]::FromStream($fileStream)
					$img = $originalImage.Clone()
					$originalImage.Dispose()
				}
				catch {
					Write-Verbose "Error drawing boss image for $bName : $_"
					$btn.BackgroundImage = $null
					$btn.Text = $bName
				}
				finally {
					if ($fileStream) { $fileStream.Dispose() }
				}

				if ($img) {
					$btn.BackgroundImage = $img
					$btn.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
					$btn.BackColor = [System.Drawing.Color]::Transparent
					$btn.Text = ''
				}
			}
			else
			{
				$btn.BackgroundImage = $null
				$btn.Text = $bName

				if (-not $global:DashboardConfig.Resources.ActiveDownloads) { $global:DashboardConfig.Resources.ActiveDownloads = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new()) }
				$tempPath = "$localPath.$([Guid]::NewGuid()).tmp"

				try {
					[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
					$wc = New-Object System.Net.WebClient
					$wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
					$wc.Add_DownloadFileCompleted({
						param($s, $e)
						try {
							if ($global:DashboardConfig.Resources.ActiveDownloads.Contains($s)) { $global:DashboardConfig.Resources.ActiveDownloads.Remove($s) }
							if (-not $e.Error -and -not $e.Cancelled) {
								try {
									if (Test-Path $localPath) { Remove-Item -LiteralPath $localPath -Force }
									Move-Item -LiteralPath $tempPath -Destination $localPath -Force
								} catch {
									Write-Verbose "Error moving downloaded file '$tempPath' to '$localPath': $_"
									return
								}

								if ($btn -and -not $btn.IsDisposed) {
									$invoker = if ($btn.IsHandleCreated) { $btn } elseif ($global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.IsHandleCreated) { $global:DashboardConfig.UI.MainForm } else { $null }
									if ($invoker) {
										try { $invoker.Invoke([Action]{ Update-BossButtonImage $btn }) } catch {}
									}
								}
							}
						} finally {
							if (Test-Path $tempPath) { try { Remove-Item -LiteralPath $tempPath -Force } catch {} }
							$s.Dispose()
						}
					}.GetNewClosure())
					$global:DashboardConfig.Resources.ActiveDownloads.Add($wc) | Out-Null
					$wc.DownloadFileAsync([Uri]$bData.url, $tempPath)
				} catch {}
			}
		}
	}
	else
	{
		$btn.BackgroundImage = $null
		$btn.Text = $bName
		$btn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
		$btn.Font = (New-Object System.Drawing.Font('Segoe UI Semibold', 9))
		$btn.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
		$btn.BackColor = [System.Drawing.Color]::FromArgb(70, 36, 36)
	}
}

function Get-RowIdentity
{
	param($Row)
	if (-not $Row -or -not $Row.Cells -or -not $Row.Cells[1].Value) { return '' }
	$fullTitle = $Row.Cells[1].Value.ToString()
	$p = 'Default'
	$title = $fullTitle

	
	if ($fullTitle -match '^\[([^\]]+)\]\s*(.*)')
	{
		$p = $Matches[1]
		$title = $Matches[2].Trim()
	}

	return "[$p]$title"
}

function GetNeuzResolution
{
	param([string]$ProfilePath)
	if ([string]::IsNullOrWhiteSpace($ProfilePath)) { return $null }
	$neuzIni = Join-Path $ProfilePath 'neuz.ini'
	if (Test-Path $neuzIni)
	{
		try
		{
			$content = Get-Content $neuzIni -Raw
			if ($content -match '(?m)^\s*resolution\s+(\d+)\s+(\d+)')
			{
				return "$($Matches[1]) x $($Matches[2])"
			}
		}
		catch {}
	}
	return $null
}

function SetNeuzResolution
{
	param([string]$ProfilePath, [string]$ResolutionString)
	if ([string]::IsNullOrWhiteSpace($ProfilePath) -or [string]::IsNullOrWhiteSpace($ResolutionString)) { return }
	$neuzIni = Join-Path $ProfilePath 'neuz.ini'
	if (-not (Test-Path $neuzIni)) { return }

	try
	{
		$parts = $ResolutionString -split ' x '
		if ($parts.Count -eq 2)
		{
			$w = $parts[0].Trim()
			$h = $parts[1].Trim()
			$content = Get-Content $neuzIni -Raw
			if ($content -match '(?m)^\s*resolution\s+\d+\s+\d+')
			{
				$content = $content -replace '(?m)(^\s*resolution\s+)\d+\s+\d+', "`${1}$w $h"
			}
			else
			{
				$content += "`r`nresolution $w $h"
			}
			Set-Content -Path $neuzIni -Value $content -Encoding ASCII -Force
		}
	}
 catch
	{
		Write-Verbose "Failed to set resolution in $neuzIni $_"
	}
}

function GetNeuzGraphicsAPI
{
	param([string]$ProfilePath)
	if ([string]::IsNullOrWhiteSpace($ProfilePath)) { return $null }
	$neuzIni = Join-Path $ProfilePath 'neuz.ini'
	if (Test-Path $neuzIni)
	{
		try
		{
			$content = Get-Content $neuzIni -Raw
			if ($content -match '(?m)^\s*GraphicsAPI\s+(\d+)')
			{
				if ($Matches[1] -eq '1') { return 'DirectX 10' }
				else { return 'DirectX 9' }
			}
		}
		catch {}
	}
	return $null
}

function SetNeuzGraphicsAPI
{
	param([string]$ProfilePath, [string]$DirectXString)
	if ([string]::IsNullOrWhiteSpace($ProfilePath) -or [string]::IsNullOrWhiteSpace($DirectXString)) { return }
	$neuzIni = Join-Path $ProfilePath 'neuz.ini'
	if (-not (Test-Path $neuzIni)) { return }

	try
	{
		$apiValue = if ($DirectXString -match '10') { '1' } else { '0' }
		$content = Get-Content $neuzIni -Raw
		if ($content -match '(?m)^\s*GraphicsAPI\s+\d+')
		{
			$content = $content -replace '(?m)(^\s*GraphicsAPI\s+)\d+', "`${1}$apiValue"
		}
		else
		{
			$content += "`r`nGraphicsAPI $apiValue"
		}
		Set-Content -Path $neuzIni -Value $content -Encoding ASCII -Force
	}
	catch
	{
		Write-Verbose "Failed to set GraphicsAPI in $neuzIni $_"
	}
}

function SyncProfilesToConfig
{
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	try
	{
		Write-Verbose '  UI: Syncing ProfileGrid to config profiles...'
		$UI = $global:DashboardConfig.UI
		if (-not ($UI -and $UI.ProfileGrid -and $global:DashboardConfig.Config)) { return $false }

		if (-not $global:DashboardConfig.Config.Contains('Profiles') -or $null -eq $global:DashboardConfig.Config['Profiles'])
		{
			$global:DashboardConfig.Config['Profiles'] = [ordered]@{}
		}
		if (-not $global:DashboardConfig.Config.Contains('ReconnectProfiles') -or $null -eq $global:DashboardConfig.Config['ReconnectProfiles'])
		{
			$global:DashboardConfig.Config['ReconnectProfiles'] = [ordered]@{}
		}
		if (-not $global:DashboardConfig.Config.Contains('HideProfiles') -or $null -eq $global:DashboardConfig.Config['HideProfiles'])
		{
			$global:DashboardConfig.Config['HideProfiles'] = [ordered]@{}
		}

		$global:DashboardConfig.Config['Profiles'].Clear()
		$global:DashboardConfig.Config['ReconnectProfiles'].Clear()
		$global:DashboardConfig.Config['HideProfiles'].Clear()

		foreach ($row in $UI.ProfileGrid.Rows)
		{
			if ($null -ne $row.Cells -and $row.Cells.Count -gt 1 -and $null -ne $row.Cells[0].Value -and $null -ne $row.Cells[1].Value)
			{
				$key = $row.Cells[0].Value.ToString()
				$val = $row.Cells[1].Value.ToString()

				$global:DashboardConfig.Config['Profiles'][$key] = $val

				if ($row.Cells.Count -gt 2 -and ($row.Cells[2].Value -eq $true -or $row.Cells[2].Value -eq 1))
				{
					$global:DashboardConfig.Config['ReconnectProfiles'][$key] = '1'
				}

				if ($row.Cells.Count -gt 3 -and ($row.Cells[3].Value -eq $true -or $row.Cells[3].Value -eq 1))
				{
					$global:DashboardConfig.Config['HideProfiles'][$key] = '1'
				}
			}
		}

		Write-Verbose '  UI: ProfileGrid synced to config profiles.'
		return $true
	}
	catch
	{
		Write-Verbose "  UI: Failed to sync ProfileGrid to config profiles: $_"
		return $false
	}
}

function SyncUIToConfig
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[string]$ProfileToSave = $null
	)

	try
	{
		Write-Verbose '  UI: Syncing UI to config'
		$UI = $global:DashboardConfig.UI
		if (-not ($UI -and $global:DashboardConfig.Config)) { return $false }

		if (-not $global:DashboardConfig.Config.Contains('LauncherPath')) { $global:DashboardConfig.Config['LauncherPath'] = [ordered]@{} }
		if (-not $global:DashboardConfig.Config.Contains('ProcessName')) { $global:DashboardConfig.Config['ProcessName'] = [ordered]@{} }
		if (-not $global:DashboardConfig.Config.Contains('MaxClients')) { $global:DashboardConfig.Config['MaxClients'] = [ordered]@{} }
		if (-not $global:DashboardConfig.Config.Contains('Paths')) { $global:DashboardConfig.Config['Paths'] = [ordered]@{} }
		if (-not $global:DashboardConfig.Config.Contains('Login')) { $global:DashboardConfig.Config['Login'] = [ordered]@{} }
		if (-not $global:DashboardConfig.Config.Contains('Options')) { $global:DashboardConfig.Config['Options'] = [ordered]@{} }

		$global:DashboardConfig.Config['LauncherPath']['LauncherPath'] = $UI.InputLauncher.Text
		$global:DashboardConfig.Config['ProcessName']['ProcessName']   = $UI.InputProcess.Text
		$global:DashboardConfig.Config['MaxClients']['MaxClients']     = $UI.InputMax.Text
		$global:DashboardConfig.Config['Paths']['JunctionTarget']      = $UI.InputJunction.Text
		
		$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin'] = if ($UI.NeverRestartingCollectorLogin.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['WorldBossListener']           = if ($UI.WorldBossListener.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['ShowBossImages']              = if ($UI.ShowBossImages.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['AutoStartWithWindows']        = if ($UI.AutoStartWithWindows.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['AutoStartSetup']              = if ($UI.AutoStartSetup.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['NotificationPosition']        = $UI.NotificationPositionSelector.SelectedItem

		$runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
		$appName = "EntropiaDashboard"

		$taskName = $appName

		if ($UI.AutoStartWithWindows.Checked) {
			
			if ($global:DashboardConfig.State.IsRunningAsExe) {
				$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
				$action = New-ScheduledTaskAction -Execute $exePath
			} else {
				$startPath = Join-Path $global:DashboardConfig.Paths.Source 'start.ps1'
				$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$startPath`""
			}

			$trigger = New-ScheduledTaskTrigger -AtLogOn

			$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest

			try {
				Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
				Write-Verbose "Auto-start task registered successfully."
			} catch {
				Write-Verbose "Failed to register scheduled task. Ensure you are running this setup script as Administrator."
			}

		} else {
			
			if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
				Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
			}
			
		}

		$global:DashboardConfig.Config['Profiles'] = [ordered]@{}
		$global:DashboardConfig.Config['ReconnectProfiles'] = [ordered]@{}
		$global:DashboardConfig.Config['HideProfiles'] = [ordered]@{}

		foreach ($row in $UI.ProfileGrid.Rows)
		{
			if ($null -ne $row.Cells -and $row.Cells.Count -gt 1 -and $null -ne $row.Cells[0].Value -and $null -ne $row.Cells[1].Value)
			{
				$key = $row.Cells[0].Value.ToString()
				$val = $row.Cells[1].Value.ToString()
				$global:DashboardConfig.Config['Profiles'][$key] = $val

				if ($row.Cells.Count -gt 2 -and ($row.Cells[2].Value -eq $true -or $row.Cells[2].Value -eq 1))
				{
					$global:DashboardConfig.Config['ReconnectProfiles'][$key] = '1'
				}

				if ($row.Cells.Count -gt 3 -and ($row.Cells[3].Value -eq $true -or $row.Cells[3].Value -eq 1))
				{
					$global:DashboardConfig.Config['HideProfiles'][$key] = '1'
				}
			}
		}

		if ($UI.ProfileGrid.SelectedRows.Count -gt 0)
		{
			$row = $UI.ProfileGrid.SelectedRows[0]
			$global:DashboardConfig.Config['Options']['LauncherTimeout'] = $UI.InputLauncherTimeout.Text
			if ($null -ne $row.Cells -and $row.Cells.Count -gt 0 -and $null -ne $row.Cells[0].Value)
			{
				$global:DashboardConfig.Config['Options']['SelectedProfile'] = $row.Cells[0].Value.ToString()
			}
			else
			{
				$global:DashboardConfig.Config['Options']['SelectedProfile'] = ''
			}
		}
		else
		{
			$global:DashboardConfig.Config['Options']['LauncherTimeout'] = $UI.InputLauncherTimeout.Text
			$global:DashboardConfig.Config['Options']['SelectedProfile'] = ''
		}

		if ($UI.SetupGrid.SelectedRows.Count -gt 0)
		{
			$row = $UI.SetupGrid.SelectedRows[0]
			if ($null -ne $row.Cells -and $row.Cells.Count -gt 0 -and $null -ne $row.Cells[0].Value)
			{
				$global:DashboardConfig.Config['Options']['SelectedSetup'] = $row.Cells[0].Value.ToString()
			}
		}

		$targetProfile = $ProfileToSave

		if (-not $targetProfile)
		{
			$targetProfile = $UI.LoginProfileSelector.SelectedItem
			if (-not $targetProfile) { $targetProfile = 'Default' }
		}

		$profPath = $null
		if ($targetProfile -eq 'Default')
		{
			if ($global:DashboardConfig.Config['LauncherPath']['LauncherPath'])
			{
				$profPath = Split-Path -Parent $global:DashboardConfig.Config['LauncherPath']['LauncherPath']
			}
		}
		elseif ($global:DashboardConfig.Config['Profiles'].Contains($targetProfile))
		{
			$profPath = $global:DashboardConfig.Config['Profiles'][$targetProfile]
		}
        
		if ($profPath)
		{
			$resSelection = $UI.ResolutionSelector.SelectedItem
			if ($resSelection)
			{
				SetNeuzResolution -ProfilePath $profPath -ResolutionString $resSelection
			}
			$dxSelection = $UI.DirectXSelector.SelectedItem
			if ($dxSelection)
			{
				SetNeuzGraphicsAPI -ProfilePath $profPath -DirectXString $dxSelection
			}
		}

		Write-Verbose "  UI: SyncUIToConfig - Saving Login Settings for: '$targetProfile'"

		if (-not ($global:DashboardConfig.Config['LoginConfig'].Contains($targetProfile)))
		{
			$global:DashboardConfig.Config['LoginConfig'][$targetProfile] = [hashtable]@{}
		}

		$currentProfileLoginConfig = $global:DashboardConfig.Config['LoginConfig'][$targetProfile]

		$null = $coordKey

		foreach ($key in $UI.LoginPickers.Keys)
		{
			$coordKey = "${key}Coords"
			$currentProfileLoginConfig["${key}Coords"] = $UI.LoginPickers[$key].Text.Text
		}

		$currentProfileLoginConfig['PostLoginDelay'] = $UI.InputPostLoginDelay.Text

		$grid = $UI.LoginConfigGrid
		for ($i = 0; $i -le 9; $i++)
		{
			$row = $grid.Rows[$i]
			if ($null -ne $row.Cells -and $row.Cells.Count -gt 4)
			{
				$clientNum = $row.Cells[0].Value
				$s = if ($row.Cells[1].Value) { $row.Cells[1].Value.ToString() } else { '1' }
				$c = if ($row.Cells[2].Value) { $row.Cells[2].Value.ToString() } else { '1' }
				$char = if ($row.Cells[3].Value) { $row.Cells[3].Value.ToString() } else { '1' }
				$coll = if ($row.Cells[4].Value) { $row.Cells[4].Value.ToString() } else { 'No' }

				$val = "$s,$c,$char,$coll"
				$currentProfileLoginConfig["Client${clientNum}_Settings"] = $val
			}
		}

		if ($UI.MasterPauseHotkeyBtn -or $UI.MasterResumeHotkeyBtn)
		{
			if (-not $global:DashboardConfig.Config.Contains('GlobalHotkeys')) { $global:DashboardConfig.Config['GlobalHotkeys'] = [ordered]@{} }
			if ($UI.MasterPauseHotkeyBtn) { $global:DashboardConfig.Config['GlobalHotkeys']['MasterPauseHotkey'] = $UI.MasterPauseHotkeyBtn.Text }
			if ($UI.MasterResumeHotkeyBtn) { $global:DashboardConfig.Config['GlobalHotkeys']['MasterResumeHotkey'] = $UI.MasterResumeHotkeyBtn.Text }
		}

		Write-Verbose '  UI: UI synced to config'
		return $true
	}
	catch
	{
		Write-Verbose "  UI: Failed to sync UI to config: $_"
		return $false
	}
}

function SyncConfigToUI
{
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	try
	{
     
		Write-Verbose '  UI: Syncing config to UI'
		$UI = $global:DashboardConfig.UI
		if (-not ($UI -and $global:DashboardConfig.Config)) { return $false }

		if ($global:DashboardConfig.Config.Contains('LauncherPath') -and $global:DashboardConfig.Config['LauncherPath'].Contains('LauncherPath')) { $UI.InputLauncher.Text = $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] }
		if ($global:DashboardConfig.Config.Contains('ProcessName') -and $global:DashboardConfig.Config['ProcessName'].Contains('ProcessName')) { $UI.InputProcess.Text = $global:DashboardConfig.Config['ProcessName']['ProcessName'] }
		if ($global:DashboardConfig.Config.Contains('MaxClients') -and $global:DashboardConfig.Config['MaxClients'].Contains('MaxClients')) { $UI.InputMax.Text = $global:DashboardConfig.Config['MaxClients']['MaxClients'] }
		if ($global:DashboardConfig.Config.Contains('Options') -and $global:DashboardConfig.Config['Options'].Contains('LauncherTimeout')) { $UI.InputLauncherTimeout.Text = $global:DashboardConfig.Config['Options']['LauncherTimeout'] } else { $UI.InputLauncherTimeout.Text = '60' }
		if ($global:DashboardConfig.Config.Contains('Paths') -and $global:DashboardConfig.Config['Paths'].Contains('JunctionTarget')) { $UI.InputJunction.Text = $global:DashboardConfig.Config['Paths']['JunctionTarget'] }

		if ($global:DashboardConfig.Config.Contains('Login') -and $global:DashboardConfig.Config['Login'].Contains('NeverRestartingCollectorLogin')) { $UI.NeverRestartingCollectorLogin.Checked = ([int]$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin']) -eq 1 }
		if ($global:DashboardConfig.Config.Contains('Options') -and $global:DashboardConfig.Config['Options'].Contains('WorldBossListener')) { $UI.WorldBossListener.Checked = ([int]$global:DashboardConfig.Config['Options']['WorldBossListener']) -eq 1 }
		if ($global:DashboardConfig.Config.Contains('Options') -and $global:DashboardConfig.Config['Options'].Contains('ShowBossImages')) { $UI.ShowBossImages.Checked = ([int]$global:DashboardConfig.Config['Options']['ShowBossImages']) -eq 1 } else { $UI.ShowBossImages.Checked = $true }
		
		if ($global:DashboardConfig.Config.Contains('Options') -and $global:DashboardConfig.Config['Options'].Contains('AutoStartWithWindows')) { $UI.AutoStartWithWindows.Checked = ([int]$global:DashboardConfig.Config['Options']['AutoStartWithWindows']) -eq 1 }
		if ($global:DashboardConfig.Config.Contains('Options') -and $global:DashboardConfig.Config['Options'].Contains('AutoStartSetup')) { $UI.AutoStartSetup.Checked = ([int]$global:DashboardConfig.Config['Options']['AutoStartSetup']) -eq 1 }
		if ($global:DashboardConfig.Config.Contains('Options') -and $global:DashboardConfig.Config['Options'].Contains('NotificationPosition')) { 
			$pos = $global:DashboardConfig.Config['Options']['NotificationPosition']
			if ($UI.NotificationPositionSelector.Items.Contains($pos)) { $UI.NotificationPositionSelector.SelectedItem = $pos }
		} else { $UI.NotificationPositionSelector.SelectedIndex = 0 }


		try
		{
			$UI.ProfileGrid.Rows.Clear()

			$recProfiles = if ($global:DashboardConfig.Config['ReconnectProfiles']) { $global:DashboardConfig.Config['ReconnectProfiles'] } else { [ordered]@{} }
			$hideProfiles = if ($global:DashboardConfig.Config['HideProfiles']) { $global:DashboardConfig.Config['HideProfiles'] } else { [ordered]@{} }

			if ($global:DashboardConfig.Config['Profiles'])
			{
				$profiles = $global:DashboardConfig.Config['Profiles']
				foreach ($key in $profiles.Keys)
				{
					$path = $profiles[$key]

					$isRecChecked = $recProfiles.Contains($key)
					$isHideChecked = $hideProfiles.Contains($key)

					$UI.ProfileGrid.Rows.Add($key, $path, $isRecChecked, $isHideChecked) | Out-Null
				}
			}
		}
		catch { Write-Verbose "  UI: Error syncing ProfileGrid: $_" }

		try
		{
			if (-not $global:DashboardConfig.State.SetupEditMode)
			{
				$UI.SetupGrid.Rows.Clear()
				$configKeys = $global:DashboardConfig.Config.Keys
				foreach ($key in $configKeys)
				{
					if ($key -match '^Setup_(.+)$')
					{
						$setupName = $Matches[1]
						$clientCount = $global:DashboardConfig.Config[$key].Count
						$UI.SetupGrid.Rows.Add($setupName, $clientCount) | Out-Null
					}
				}
				$UI.SetupGrid.ClearSelection()
			}
		}
		catch { Write-Verbose "  UI: Error syncing SetupGrid: $_" }

		$selectedProfileName = $null
		if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['SelectedProfile'])
		{
			$selectedProfileName = $global:DashboardConfig.Config['Options']['SelectedProfile'].ToString()
		}

		if (-not [string]::IsNullOrEmpty($selectedProfileName))
		{
			$UI.ProfileGrid.ClearSelection()
			$found = $false
			foreach ($row in $UI.ProfileGrid.Rows)
			{
				if ($null -ne $row.Cells -and $row.Cells.Count -gt 0 -and $null -ne $row.Cells[0].Value -and $row.Cells[0].Value.ToString() -eq $selectedProfileName)
				{
					$row.Selected = $true
					$UI.ProfileGrid.CurrentCell = $row.Cells[0]
					$UI.ProfileGrid.Tag = $row.Index
					$found = $true
					break
				}
			}
			if (-not $found)
			{
				$UI.ProfileGrid.ClearSelection()
				$UI.ProfileGrid.Tag = -1
			}
		}
		else
		{
			$UI.ProfileGrid.ClearSelection()
			$UI.ProfileGrid.Tag = -1
		}

		$selectedSetupName = $null
		if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['SelectedSetup'])
		{
			$selectedSetupName = $global:DashboardConfig.Config['Options']['SelectedSetup'].ToString()
		}

		if (-not [string]::IsNullOrEmpty($selectedSetupName))
		{
			$UI.SetupGrid.ClearSelection()
			foreach ($row in $UI.SetupGrid.Rows)
			{
				if ($row.Cells[0].Value -eq $selectedSetupName)
				{
					$row.Selected = $true
					break
				}
			}
		}

		$selectedLoginProfile = $null
		try
		{
			if ($UI.LoginProfileSelector.SelectedItem)
			{
				$selectedLoginProfile = $UI.LoginProfileSelector.SelectedItem.ToString()
			}
		}
		catch { Write-Verbose "  UI: Error getting selected login profile: $_" }

		if (-not $selectedLoginProfile)
		{
			$selectedLoginProfile = 'Default'
		}

		$UI.LoginProfileSelector.Tag = $selectedLoginProfile

		$profPath = $null
		if ($selectedLoginProfile -eq 'Default')
		{
			if ($global:DashboardConfig.Config['LauncherPath']['LauncherPath'])
			{
				$profPath = Split-Path -Parent $global:DashboardConfig.Config['LauncherPath']['LauncherPath']
			}
		}
		elseif ($global:DashboardConfig.Config['Profiles'].Contains($selectedLoginProfile))
		{
			$profPath = $global:DashboardConfig.Config['Profiles'][$selectedLoginProfile]
		}
        
		if ($profPath)
		{
			$currentRes = GetNeuzResolution -ProfilePath $profPath
			if ($currentRes -and $UI.ResolutionSelector.Items.Contains($currentRes))
			{
				$UI.ResolutionSelector.SelectedItem = $currentRes
			}
			else
			{
				$UI.ResolutionSelector.SelectedItem = $null
			}

			$currentDx = GetNeuzGraphicsAPI -ProfilePath $profPath
			if ($currentDx -and $UI.DirectXSelector.Items.Contains($currentDx))
			{
				$UI.DirectXSelector.SelectedItem = $currentDx
			}
			else
			{
				$UI.DirectXSelector.SelectedItem = $null
			}
		}
		else
		{
			$UI.ResolutionSelector.SelectedItem = $null
			$UI.DirectXSelector.SelectedItem = $null
		}

		Write-Verbose "  UI: SyncConfigToUI - Loading settings for profile '$selectedLoginProfile'."

		if ($global:DashboardConfig.Config['LoginConfig'] -and $global:DashboardConfig.Config['LoginConfig'][$selectedLoginProfile])
		{
			$lc = $global:DashboardConfig.Config['LoginConfig'][$selectedLoginProfile]
		}
		else
		{
			$lc = [hashtable]@{}
		}

		foreach ($key in $UI.LoginPickers.Keys)
		{
			$coordKey = "${key}Coords"
			if ($lc[$coordKey])
			{
				$UI.LoginPickers[$key].Text.Text = $lc[$coordKey]
			}
			else
			{
				$UI.LoginPickers[$key].Text.Text = '0,0'
			}
		}

		if ($lc['PostLoginDelay']) { $UI.InputPostLoginDelay.Text = $lc['PostLoginDelay'] }
		else { $UI.InputPostLoginDelay.Text = '1' }

		$grid = $UI.LoginConfigGrid
		for ($i = 0; $i -le 9; $i++)
		{
			$row = $grid.Rows[$i]
			$clientNum = $i + 1
			$settingKey = "Client${clientNum}_Settings"

			$s = '1'; $c = '1'; $char = '1'; $coll = 'No'

			if ($lc[$settingKey])
			{
				$parts = $lc[$settingKey] -split ','
				if ($parts.Count -eq 4)
				{
					if ($parts[0] -in @('1','2')) { $s = $parts[0] }
					if ($parts[1] -in @('1','2')) { $c = $parts[1] }
					if ($parts[2] -in @('1','2','3','4','5','6')) { $char = $parts[2] }
					if ($parts[3] -in @('Yes','No')) { $coll = $parts[3] }
				}
			}

			if ($null -ne $row.Cells -and $row.Cells.Count -gt 4)
			{
				$row.Cells[1].Value = $s
				$row.Cells[2].Value = $c
				$row.Cells[3].Value = $char
				$row.Cells[4].Value = $coll
			}
		}

		$masterPauseText = 'none'
		$masterResumeText = 'none'
		if ($global:DashboardConfig.Config.Contains('GlobalHotkeys'))
		{
			$globalHks = $global:DashboardConfig.Config['GlobalHotkeys']
			if ($globalHks.Contains('MasterPauseHotkey')) { $masterPauseText = $globalHks['MasterPauseHotkey'] }
			if ($globalHks.Contains('MasterResumeHotkey')) { $masterResumeText = $globalHks['MasterResumeHotkey'] }
		}
		if ($UI.MasterPauseHotkeyBtn) { $UI.MasterPauseHotkeyBtn.Text = $masterPauseText }
		if ($UI.MasterResumeHotkeyBtn) { $UI.MasterResumeHotkeyBtn.Text = $masterResumeText }

		if ($UI.MasterHotkeysStatus)
		{
			if ($global:DashboardConfig.State.MasterHotkeysPaused)
			{
				$UI.MasterHotkeysStatus.Text = ([char]0x23F8).ToString() + ' Paused'
				$UI.MasterHotkeysStatus.ForeColor = [System.Drawing.Color]::FromArgb(255, 165, 0)
			}
			else
			{
				$UI.MasterHotkeysStatus.Text = ([char]0x25CF).ToString() + ' Active'
				$UI.MasterHotkeysStatus.ForeColor = [System.Drawing.Color]::FromArgb(100, 200, 100)
			}
		}

		if (Get-Command RefreshNoteGrid -ErrorAction SilentlyContinue -Verbose:$False) { RefreshNoteGrid }

		Write-Verbose '  UI: Config synced to UI'
		return $true
	}
	catch
	{
		Write-Verbose "  UI: Failed to sync config to UI: $_"
		return $false
	}
}

#endregion

#region Core Functions

function InitializeUI
{
	[CmdletBinding()]
	param()

	Write-Verbose '  UI: Initializing UI...'

	$global:DashboardConfig.Resources.BossData = @{
		"Entropia King" = @{
			name = "Entropia King"
			url  = "https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Entropia-King.png"
			role = "1030608207460171837"
		}
		"Clockworks" = @{
			name = "Clockworks"
			url  = "https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Clockworks.png"
			role = "1030608298266869791"
		}
		"Royal Knight" = @{
			name = "Royal Knight"
			url  = "https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Royal-Knight.png"
			role = "1030608553741930536"
		}
		"Solarion" = @{
			name = "Solarion"
			url  = "https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Solarion.png"
			role = "1418869138016960512"
		}
		"C-A01" = @{
			name = "C-A01"
			url  = "https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/C-A01.png"
			role = "1210685566010658926"
		}
		"Great Venux Tree" = @{
			name = "Great Venux Tree"
			url  = "https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Great-Venux-Tree.png"
			role = "1210685569164640296"
		}
		"Grinch" = @{
			name = "Grinch"
			url  = "https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Grinch.png"
			role = "1444911271978729616"
		}
	}
	
	$global:DashboardConfig.Resources.UpgradeData = ConvertFrom-Json @'
{
  "Set": {
    "Rainbow Stone": [10.0, 10.0, 10.0, 10.0, 20.0, 20.0, 30.0, 30.0, 40.0, 40.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Fire/Ice Shard (1-5)": [700.0, 900.0, 1000.0, 1200.0, 1500.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Fire/Ice Core (6-10)": [0.0, 0.0, 0.0, 0.0, 0.0, 500.0, 700.0, 900.0, 1000.0, 1200.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "C.E.L. Smithing Stone": [100.0, 250.0, 300.0, 350.0, 500.0, 750.0, 1000.0, 1300.0, 1500.0, 2000.0, 2500.0, 3250.0, 4250.0, 5250.0, 6000.0],
    "Cardia Monster Coin": [300.0, 500.0, 700.0, 1000.0, 1300.0, 1600.0, 2000.0, 2200.0, 2500.0, 2750.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +1 / +2": [1000.0, 1200.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +3 / +4": [0.0, 0.0, 820.0, 1000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +5 / +6": [0.0, 0.0, 0.0, 0.0, 2800.0, 3000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +7 / +8": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 3800.0, 5000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +9 / +10": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 3000.0, 4000.0, 0.0, 0.0, 0.0, 0.0, 0.0],
	"Perin": [350.0, 500.0, 1200.0, 1800.0, 2100.0, 2200.0, 2500.0, 5500.0, 6200.0, 7000.0, 13000.0, 13500.0, 14500.0, 15000.0, 16000.0],
    "ECP": [100.0, 120.0, 150.0, 170.0, 200.0, 220.0, 250.0, 270.0, 300.0, 350.0, 1400.0, 1550.0, 1700.0, 1850.0, 2000.0],
    "Gleamcore": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 50.0, 60.0, 70.0, 80.0, 100.0],
    "Rune of Kogarashi": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 16000.0, 18500.0, 21000.0, 24000.0, 28000.0],
    "Kogarashi Orb": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 16000.0, 18500.0, 21000.0, 24000.0, 28000.0],
    "Kogarashi Claw": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 16000.0, 18500.0, 21000.0, 24000.0, 28000.0],
    "Smithing Stone +11/+12": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 18000.0, 22000.0, 0.0, 0.0, 0.0],
    "Smithing Stone +13/+14": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 16000.0, 20000.0, 0.0],
    "Smithing Stone +15": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 16000.0]
  },
  "Weapon": {
    "Eternal Stone": [20.0, 20.0, 20.0, 20.0, 40.0, 40.0, 60.0, 60.0, 80.0, 80.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Cardia Monster Coin": [3000.0, 4100.0, 5000.0, 5500.0, 6200.0, 7000.0, 8300.0, 9200.0, 10400.0, 11000.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Cardia Coin": [160.0, 0.0, 220.0, 300.0, 560.0, 600.0, 700.0, 800.0, 1000.0, 1200.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "C.E.L. Smithing Stone": [750.0, 1000.0, 1200.0, 1400.0, 1600.0, 3000.0, 3500.0, 4600.0, 7000.0, 9000.0, 11000.0, 11500.0, 12000.0, 12500.0, 13500.0],
    "Royal Knight Rune": [10.0, 10.0, 20.0, 20.0, 30.0, 30.0, 40.0, 40.0, 60.0, 60.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +1 / +2": [6800.0, 7400.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +3 / +4": [0.0, 0.0, 6000.0, 7000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +5 / +6": [0.0, 0.0, 0.0, 0.0, 10000.0, 15000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +7 / +8": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 25000.0, 30000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +9 / +10": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 10500.0, 12000.0, 0.0, 0.0, 0.0, 0.0, 0.0],
	"Perin": [2000.0, 4000.0, 7000.0, 9500.0, 11000.0, 13000.0, 15500.0, 18000.0, 20500.0, 22000.0, 25000.0, 35000.0, 45000.0, 60000.0, 75000.0],
    "ECP": [250.0, 400.0, 600.0, 800.0, 1000.0, 1500.0, 1700.0, 2080.0, 2700.0, 3200.0, 2500.0, 3500.0, 4500.0, 6500.0, 8000.0],
    "Bloomcore": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 90.0, 100.0, 110.0, 120.0, 150.0],
    "Heavens Force": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 30000.0, 35000.0, 40000.0, 45000.0, 50000.0],
    "Heavens Scale": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 30000.0, 35000.0, 40000.0, 45000.0, 50000.0],
    "Mystic Shard": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 50.0, 70.0, 90.0, 120.0, 120.0],
    "Smithing Stone +11/+12": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 34000.0, 40000.0, 0.0, 0.0, 0.0],
    "Smithing Stone +13/+14": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 36000.0, 44000.0, 0.0],
    "Smithing Stone +15": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 40000.0]
  },
  "Jewelry": {
    "Rainbow Stone": [10.0, 10.0, 10.0, 10.0, 20.0, 20.0, 30.0, 30.0, 40.0, 40.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Royal Rainbow Diamond": [500.0, 750.0, 1100.0, 1300.0, 1450.0, 1700.0, 2100.0, 2400.0, 2600.0, 3000.0, 3500.0, 3750.0, 4000.0, 4250.0, 4750.0],
    "Royal Power Stone": [500.0, 750.0, 1100.0, 1300.0, 1450.0, 1700.0, 2100.0, 2400.0, 2600.0, 3000.0, 3500.0, 3750.0, 4000.0, 4250.0, 4750.0],
    "C.E.L. Smithing Stone": [250.0, 320.0, 400.0, 600.0, 1000.0, 1300.0, 1700.0, 2200.0, 2800.0, 4000.0, 4000.0, 5000.0, 6000.0, 7000.0, 8000.0],
    "Cardia Coin": [80.0, 100.0, 120.0, 150.0, 200.0, 280.0, 320.0, 360.0, 400.0, 450.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +1 / +2": [10000.0, 15000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +3 / +4": [0.0, 0.0, 7500.0, 10000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +5 / +6": [0.0, 0.0, 0.0, 0.0, 5000.0, 6000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +7 / +8": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 6400.0, 7400.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    "Smithing Stone +9 / +10": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 5500.0, 7000.0, 0.0, 0.0, 0.0, 0.0, 0.0],
	"Perin": [1000.0, 1440.0, 2000.0, 2400.0, 3200.0, 3900.0, 4500.0, 6800.0, 7400.0, 9000.0, 13000.0, 13500.0, 14000.0, 14500.0, 15000.0],
    "ECP": [160.0, 200.0, 240.0, 300.0, 480.0, 560.0, 640.0, 720.0, 1000.0, 1500.0, 1700.0, 2000.0, 2200.0, 2500.0, 3000.0],
    "Echo of Eternity": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 10.0, 10.0, 20.0, 30.0, 30.0],
    "Kogarashi Blood": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 16000.0, 18500.0, 21000.0, 24000.0, 28000.0],
    "Smithing Stone +11/+12": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 15000.0, 20000.0, 0.0, 0.0, 0.0],
    "Smithing Stone +13/+14": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 12000.0, 15000.0, 0.0],
    "Smithing Stone +15": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 12000.0]
  },
  "Key": {
    "Obsidian Mirroring Stone": [20.0, 30.0, 40.0, 60.0, 80.0],
    "Devil Oni Soul": [8000.0, 11000.0, 15000.0, 18000.0, 21000.0],
    "Obsidian Oni Eye": [15000.0, 22000.0, 31000.0, 35000.0, 40000.0],
    "Enhanced C.E.L.": [12000.0, 18000.0, 24000.0, 28000.0, 34000.0],
    "Enhanced Smithing Stone +1/+2": [40000.0, 50000.0, 0.0, 0.0, 0.0],
    "Enhanced Smithing Stone +3/+4": [0.0, 0.0, 20000.0, 24000.0, 0.0],
    "Enhanced Smithing Stone +5/+6": [0.0, 0.0, 0.0, 0.0, 20000.0],
	"Kogarashi Sunmark Coin": [130.0, 180.0, 250.0, 340.0, 400.0],
	"Perin": [20000.0, 30000.0, 40000.0, 50000.0, 60000.0],
    "ECP": [3000.0, 4500.0, 6000.0, 7500.0, 9000.0]
  }
}
'@

	$uiPropertiesToAdd = @{}


	$p = @{ type = 'Form'; visible = $false; width = 470; height = 440; bg = @(30, 30, 30); id = 'MainForm'; text = 'Entropia Dashboard'; startPosition = 'Manual'; formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None }
	$mainForm = SetUIElement @p
	$p = @{ type = 'Form'; width = 600; height = 795; bg = @(30, 30, 30); id = 'SettingsForm'; text = 'Settings'; startPosition = 'CenterScreen'; formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None; topMost = $false; opacity = 0.0 }
	$settingsForm = SetUIElement @p
	$settingsForm.Owner = $mainForm
	$p = @{ type = 'Form'; width = 400; height = 600; bg = @(30, 30, 30); id = 'ExtraForm'; text = 'Extra'; startPosition = 'CenterScreen'; formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None; topMost = $false; opacity = 0.0 }
	$extraForm = SetUIElement @p
	$extraForm.Owner = $mainForm


	if ($wp = $global:DashboardConfig.Config['WindowPosition']) {
		try {
			if (($h = [int]$wp.MainFormHeight) -ge 440) {
				$mainForm.Height = $h; Write-Verbose "UI: Applied height: $h"
			}
			$x, $y = [int]$wp.MainFormX, [int]$wp.MainFormY
			$rect  = [System.Drawing.Rectangle]::new($x, $y, $mainForm.Width, $mainForm.Height)
			if ([System.Windows.Forms.Screen]::AllScreens.WorkingArea.IntersectsWith($rect) -contains $true) {
				$mainForm.Location = [System.Drawing.Point]::new($x, $y)
			}
			if ($wp.Contains('ExtraFormX') -and $wp.Contains('ExtraFormY')) {
				$ex, $ey = [int]$wp.ExtraFormX, [int]$wp.ExtraFormY
				$eRect = [System.Drawing.Rectangle]::new($ex, $ey, $extraForm.Width, $extraForm.Height)
				if ([System.Windows.Forms.Screen]::AllScreens.WorkingArea.IntersectsWith($eRect) -contains $true) {
					$extraForm.StartPosition = 'Manual'; $extraForm.Location = [System.Drawing.Point]::new($ex, $ey)
				}
			}
			if ($wp.Contains('SettingsFormX') -and $wp.Contains('SettingsFormY')) {
				$sx, $sy = [int]$wp.SettingsFormX, [int]$wp.SettingsFormY
				$sRect = [System.Drawing.Rectangle]::new($sx, $sy, $settingsForm.Width, $settingsForm.Height)
				if ([System.Windows.Forms.Screen]::AllScreens.WorkingArea.IntersectsWith($sRect) -contains $true) {
					$settingsForm.StartPosition = 'Manual'; $settingsForm.Location = [System.Drawing.Point]::new($sx, $sy)
				}
			}
		} catch { Write-Verbose "UI: Failed to apply saved window position: $_" }
	} else { $mainForm.StartPosition = 'CenterScreen' }


	if (-not $global:DashboardConfig.UI) { $global:DashboardConfig | Add-Member -MemberType NoteProperty -Name UI -Value ([PSCustomObject]@{}) -Force }

	if (-not $global:DashboardConfig.Resources.ActiveDownloads) { $global:DashboardConfig.Resources.ActiveDownloads = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new()) }

	$newTip = {
		$t = [System.Windows.Forms.ToolTip]@{AutoPopDelay=5000;InitialDelay=100;ReshowDelay=10;ShowAlways=$true;OwnerDraw=$true}
		$t | Add-Member -MemberType NoteProperty -Name 'TipFont' -Value (New-Object System.Drawing.Font('Segoe UI', 9))
		$t.Add_Draw({
			$g, $b, $c = $_.Graphics, $_.Bounds, [System.Drawing.Color]
			$g.FillRectangle((New-Object System.Drawing.SolidBrush $c::FromArgb(30,30,30)), $b)
			$g.DrawRectangle((New-Object System.Drawing.Pen $c::FromArgb(100,100,100)), $b.X, $b.Y, $b.Width-1, $b.Height-1)
			$g.DrawString($_.ToolTipText, $this.TipFont, (New-Object System.Drawing.SolidBrush $c::FromArgb(240,240,240)), 3, 3, [System.Drawing.StringFormat]::GenericTypographic)
		})
		$t.Add_Popup({
			$g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
			$s = $g.MeasureString($this.GetToolTip($_.AssociatedControl), $this.TipFont, [System.Drawing.PointF]::new(0,0), [System.Drawing.StringFormat]::GenericTypographic)
			$g.Dispose(); $_.ToolTipSize = [System.Drawing.Size]::new($s.Width+12, $s.Height+8)
		})
		return $t
	}

	$global:DashboardConfig.UI.ToolTipMain     = ($toolTipMain     = &$newTip)
	$global:DashboardConfig.UI.ToolTipSettings = ($toolTipSettings = &$newTip)
	$global:DashboardConfig.UI.ToolTipExtra    = ($toolTipExtra    = &$newTip)



	if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon))
	{
		try
		{
			$icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon)
			$mainForm.Icon = $icon
			$settingsForm.Icon = $icon
			$ExtraForm.Icon = $icon
		}
		catch {}
	}

	$script:CurrentBuilderToolTip = $toolTipMain

	#region MainForm

	$p = @{ type = 'Panel'; width = 470; height = 30; bg = @(20, 20, 20); id = 'TopBar' }
	$topBar = SetUIElement @p
	$p = @{ type = 'Label'; width = 140; height = 12; top = 5; left = 10; fg = @(240, 240, 240); id = 'TitleLabel'; text = 'Entropia Dashboard'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)) }
	$titleLabelForm = SetUIElement @p
	$p = @{ type = 'Label'; width = 140; height = 10; top = 16; left = 10; fg = @(230, 230, 230); id = 'CopyrightLabel'; text = [char]0x00A9 + ' Immortal / Divine 2026 - v2.8.2'; font = (New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Italic)) }
	$copyrightLabelForm = SetUIElement @p
	$p = @{ type = 'Button'; width = 40; height = 30; left = 370; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'InfoForm'; text = 'Guide'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 7)); tooltip = "Help / Info`nOpens the project website for documentation and support."; noCustomPaint = $true }
	$btnInfoForm = SetUIElement @p
	$p = @{ type = 'Button'; width = 30; height = 30; left = 410; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'MinForm'; text = '_'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)); tooltip = "Minimize`nMinimize the dashboard to the taskbar."; noCustomPaint = $true }
	$btnMinimizeForm = SetUIElement @p
	$p = @{ type = 'Button'; width = 30; height = 30; left = 440; bg = @(150, 20, 20); fg = @(240, 240, 240); id = 'CloseForm'; text = [char]0x166D; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)); tooltip = "Exit`nClose the Entropia Dashboard application."; noCustomPaint = $true }
	$btnCloseForm = SetUIElement @p


	$p = @{ type = 'Button'; width = 125; height = 30; top = 40; left = 15; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'Launch'; text = 'Launch ' + [char]0x25BE; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Launch`nRight Click: Open menu for launch options."; noCustomPaint = $true }
	$btnLaunch = SetUIElement @p

	$LaunchContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
	$LaunchContextMenu.Name = 'LaunchContextMenu'
	$LaunchContextMenu.Renderer = New-Object Custom.DarkRenderer
	$LaunchContextMenu.Items.Add('Loading...') | Out-Null
	$btnLaunch.ContextMenuStrip = $LaunchContextMenu

	$p = @{ type = 'Button'; width = 125; height = 30; top = 40; left = 150; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'Login'; text = 'Login'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Login`nLogs in selected clients using profile settings."; noCustomPaint = $true }
	$btnLogin = SetUIElement @p
	$p = @{ type = 'Button'; width = 80; height = 30; top = 40; left = 285; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'Settings'; text = 'Settings'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Settings`nOpen the configuration menu to adjust paths, profiles, and options."; noCustomPaint = $true }
	$btnSettings = SetUIElement @p
	$p = @{ type = 'Button'; width = 80; height = 30; top = 40; left = 375; bg = @(150, 20, 20); fg = @(240, 240, 240); id = 'Terminate'; text = 'Terminate ' + [char]0x25BE; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Terminate`nLeft Click: Instantly close selected clients.`nRight Click: Disconnect selected clients (TCP close)."; noCustomPaint = $true }
	$btnStop = SetUIElement @p
	$p = @{ type = 'Button'; width = 125; height = 30; top = 75; left = 15; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'Ftool'; text = 'Ftool ' + [char]0x25BE; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Ftool`nLeft Click: Open Ftool.`nRight Click: Save/Load template for selected client."; noCustomPaint = $true }
	$btnFtool = SetUIElement @p
	$btnFtool.Add_Click({
		if ($global:DashboardConfig.UI.DataGridMain.SelectedRows.Count -eq 0) { return }
		$row = $global:DashboardConfig.UI.DataGridMain.SelectedRows[0]
		FtoolSelectedRow $row
	})
	$p = @{ type = 'Button'; width = 125; height = 30; top = 75; left = 150; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'Macro'; text = 'Macro ' + [char]0x25BE; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Macro`nLeft Click: Open Macro.`nRight Click: Save/Load template for selected client."; noCustomPaint = $true }
	$btnMacro = SetUIElement @p
	$btnMacro.Add_Click({
		if ($global:DashboardConfig.UI.DataGridMain.SelectedRows.Count -eq 0) { return }
		$row = $global:DashboardConfig.UI.DataGridMain.SelectedRows[0]
		MacroSelectedRow $row
	})
	$p = @{ type = 'Button'; width = 80; height = 30; top = 75; left = 285; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'Extra'; text = 'Extra ' + [char]0x25BE; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Extra`nLeft Click: Open the Extra panel for World Boss timers, Notes, Tools, and Notifications.`nRight Click: Open QuickCommands."; noCustomPaint = $true }
	$btnExtra = SetUIElement @p
	$p = @{ type = 'Button'; width = 80; height = 30; top = 75; left = 375; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'Wiki'; text = 'Wiki'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Wiki`nOpen the Wiki panel for Entropia."; noCustomPaint = $true }
	$btnWiki = SetUIElement @p

	$grayNormalStart = [System.Drawing.Color]::FromArgb(50, 50, 50)
	$grayNormalEnd   = [System.Drawing.Color]::FromArgb(35, 35, 35)
	$grayHoverStart  = [System.Drawing.Color]::FromArgb(70, 70, 70)
	$grayHoverEnd    = [System.Drawing.Color]::FromArgb(50, 50, 50)
	$grayLeftStart   = [System.Drawing.Color]::FromArgb(90, 90, 90)
	$grayLeftEnd     = [System.Drawing.Color]::FromArgb(70, 70, 70)
	$grayRightStart  = [System.Drawing.Color]::FromArgb(80, 80, 110)
	$grayRightEnd    = [System.Drawing.Color]::FromArgb(55, 55, 80)

	$redNormalStart  = [System.Drawing.Color]::FromArgb(170, 30, 30)
	$redNormalEnd    = [System.Drawing.Color]::FromArgb(130, 15, 15)
	$redHoverStart   = [System.Drawing.Color]::FromArgb(200, 40, 40)
	$redHoverEnd     = [System.Drawing.Color]::FromArgb(155, 20, 20)
	$redLeftStart    = [System.Drawing.Color]::FromArgb(140, 15, 15)
	$redLeftEnd      = [System.Drawing.Color]::FromArgb(100, 10, 10)
	$redRightStart   = [System.Drawing.Color]::FromArgb(180, 30, 90)
	$redRightEnd     = [System.Drawing.Color]::FromArgb(145, 15, 65)

	$blueNormalStart = [System.Drawing.Color]::FromArgb(35, 90, 180)
	$blueNormalEnd   = [System.Drawing.Color]::FromArgb(20, 60, 130)
	$blueHoverStart  = [System.Drawing.Color]::FromArgb(50, 115, 220)
	$blueHoverEnd    = [System.Drawing.Color]::FromArgb(35, 90, 180)
	$blueLeftStart   = [System.Drawing.Color]::FromArgb(25, 70, 140)
	$blueLeftEnd     = [System.Drawing.Color]::FromArgb(15, 45, 100)
	$blueRightStart  = [System.Drawing.Color]::FromArgb(55, 100, 190)
	$blueRightEnd    = [System.Drawing.Color]::FromArgb(40, 75, 145)

	$greenNormalStart = [System.Drawing.Color]::FromArgb(40, 140, 70)
	$greenNormalEnd   = [System.Drawing.Color]::FromArgb(25, 95, 45)
	$greenHoverStart  = [System.Drawing.Color]::FromArgb(55, 175, 90)
	$greenHoverEnd    = [System.Drawing.Color]::FromArgb(40, 140, 70)
	$greenLeftStart   = [System.Drawing.Color]::FromArgb(30, 110, 55)
	$greenLeftEnd     = [System.Drawing.Color]::FromArgb(20, 75, 35)
	$greenRightStart  = [System.Drawing.Color]::FromArgb(60, 150, 85)
	$greenRightEnd    = [System.Drawing.Color]::FromArgb(45, 110, 60)


	Register-ButtonFadeEffect -Button $btnInfoForm -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd
	Register-ButtonFadeEffect -Button $btnMinimizeForm -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd
	Register-ButtonFadeEffect -Button $btnCloseForm -NormalStartColor $redNormalStart -NormalEndColor $redNormalEnd -HoverStartColor $redHoverStart -HoverEndColor $redHoverEnd -LeftPressedStartColor $redLeftStart -LeftPressedEndColor $redLeftEnd -RightPressedStartColor $redRightStart -RightPressedEndColor $redRightEnd
	Register-ButtonFadeEffect -Button $btnLaunch   -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd
	Register-ButtonFadeEffect -Button $btnLogin    -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd
	Register-ButtonFadeEffect -Button $btnSettings -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd
	Register-ButtonFadeEffect -Button $btnStop -NormalStartColor $redNormalStart -NormalEndColor $redNormalEnd -HoverStartColor $redHoverStart -HoverEndColor $redHoverEnd -LeftPressedStartColor $redLeftStart -LeftPressedEndColor $redLeftEnd -RightPressedStartColor $redRightStart -RightPressedEndColor $redRightEnd
	Register-ButtonFadeEffect -Button $btnFtool    -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd
	Register-ButtonFadeEffect -Button $btnMacro    -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd
	Register-ButtonFadeEffect -Button $btnExtra    -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd
	Register-ButtonFadeEffect -Button $btnWiki     -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd

	$p = @{ type = 'DataGridView'; width = 450; height = 300; top = 115; left = 10; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'DataGridMain'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$DataGridMain = SetUIElement @p

	$DataGridMain.Height = $mainForm.Height - 115 - 25

	$mainGridCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Index'; HeaderText = '#'; FillWeight = 12; SortMode = 'NotSortable';}),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Titel'; HeaderText = 'Titel'; SortMode = 'NotSortable';}),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'ID'; HeaderText = 'PID'; FillWeight = 20; SortMode = 'NotSortable';}),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'State'; HeaderText = 'State'; FillWeight = 30; SortMode = 'NotSortable';})    
	)
	$DataGridMain.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$mainGridCols)

	$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
	$ctxMenu.Renderer = New-Object Custom.DarkRenderer
	$itmFront = New-Object System.Windows.Forms.ToolStripMenuItem('Show')
	$itmBack = New-Object System.Windows.Forms.ToolStripMenuItem('Minimize')
	$itmResizeCenter = New-Object System.Windows.Forms.ToolStripMenuItem('Resize')
	$itmSavePos = New-Object System.Windows.Forms.ToolStripMenuItem('Save Position')
	$itmLoadPos = New-Object System.Windows.Forms.ToolStripMenuItem('Load Position')
	$itmRelog = New-Object System.Windows.Forms.ToolStripMenuItem('Relog after Disconnect')
	$itmSetHotkeyShow = New-Object System.Windows.Forms.ToolStripMenuItem('Set Hotkey (Show)')
	$itmSetHotkeyHide = New-Object System.Windows.Forms.ToolStripMenuItem('Set Hotkey (Hide)')
	$itmSetHotkeyLoadPos = New-Object System.Windows.Forms.ToolStripMenuItem('Set Hotkey (Load Pos)')
	$ctxMenu.Items.AddRange(@($itmFront, $itmBack, $itmResizeCenter, $itmSavePos, $itmLoadPos, $itmRelog, $itmSetHotkeyShow, $itmSetHotkeyHide, $itmSetHotkeyLoadPos))
	$DataGridMain.ContextMenuStrip = $ctxMenu

	$p = @{ type = 'Panel'; height = 8; dock = 'Bottom'; cursor = 'SizeNS'; id = 'ResizeGrip'; bg = @(20, 20, 20); tooltip = "Resize`nClick and drag to resize the dashboard window." }
	$resizeGrip = SetUIElement @p

	#endregion

	$script:CurrentBuilderToolTip = $toolTipSettings

	#region SettingsForm

	$settingsTabs = New-Object Custom.DarkTabControl
	$settingsTabs.Dock = 'Top'
	$settingsTabs.Height = 740

	$tabGeneral = New-Object System.Windows.Forms.TabPage 'General'
	$tabGeneral.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)


	$tabLoginSettings = New-Object System.Windows.Forms.TabPage 'Login Settings'
	$tabLoginSettings.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
	$tabLoginSettings.AutoScroll = $true

	$tabHotkeys = New-Object System.Windows.Forms.TabPage 'Hotkeys'
	$tabHotkeys.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

	$settingsTabs.TabPages.Add($tabGeneral)
	$settingsTabs.TabPages.Add($tabLoginSettings)
	$settingsTabs.TabPages.Add($tabHotkeys)

	$settingsForm.Controls.Add($settingsTabs)

	#region General

	$p = @{ type = 'Label'; width = 127; height = 20; top = 25; left = 20; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); id = 'LabelLauncher'; text = 'Select Main Launcher:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Main Launcher`nSelect the 'Launcher.exe' of your main game client.`nThis path is used as the base for all operations." }
	$lblLauncher = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 250; height = 30; top = 50; left = 20; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'InputLauncher'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Main Launcher`nSelect the 'Launcher.exe' of your main game client.`nThis path is used as the base for all operations." }
	$txtLauncher = SetUIElement @p
	$p = @{ type = 'Button'; width = 55; height = 25; top = 20; left = 150; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'Browse'; text = 'Browse'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Main Launcher`nSelect the 'Launcher.exe' of your main game client.`nThis path is used as the base for all operations."; noCustomPaint = $true }
	$btnBrowseLauncher = SetUIElement @p

	$p = @{ type = 'Label'; width = 85; height = 20; top = 95; left = 20; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); id = 'LabelProcess'; text = 'Process Name:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Process Name`nThe name of the game process (usually 'neuz').`nUsed to detect running clients." }
	$lblProcessName = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 250; height = 30; top = 120; left = 20; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'InputProcess'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Process Name`nThe name of the game process (usually 'neuz').`nUsed to detect running clients." }
	$txtProcessName = SetUIElement @p

	$p = @{ type = 'Label'; width = 250; height = 20; top = 165; left = 20; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); id = 'LabelMax'; text = 'Max Total Clients For Selection:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Max Clients`nLimit the total number of clients the dashboard can launch for the default launcher or selected profile.`nPrevents accidental mass-launching." }
	$lblMaxClients = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 250; height = 30; top = 190; left = 20; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'InputMax'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Max Clients`nLimit the total number of clients the dashboard can launch for the default launcher or selected profile`nPrevents accidental mass-launching." }
	$txtMaxClients = SetUIElement @p

	$p = @{ type = 'Label'; width = 250; height = 20; top = 235; left = 20; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); id = 'LabelLauncherTimeout'; text = 'Launcher Timeout (Seconds):'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Launcher Timeout`nSeconds to wait for the launcher to open the game client before retrying." }
	$lblLauncherTimeout = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 250; height = 30; top = 260; left = 20; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'InputLauncherTimeout'; text = '60'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Launcher Timeout`nSeconds to wait for the launcher to open the game client before retrying." }
	$txtLauncherTimeout = SetUIElement @p

	$p = @{ type = 'Label'; width = 125; height = 20; top = 305; left = 20; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); id = 'LabelJunction'; text = 'Select Profiles Folder:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Profiles Folder`nThe directory where your client profiles (copies) will be stored." }
	$lblJunction = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 250; height = 30; top = 330; left = 20; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'InputJunction'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Profiles Folder`nThe directory where your client profiles (copies) will be stored." }
	$txtJunction = SetUIElement @p

	$p = @{ type = 'Button'; width = 55; height = 25; top = 300; left = 145; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'BrowseJunction'; text = 'Browse'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Browse Folder`nChoose the directory for storing client profiles."; noCustomPaint = $true }
	$btnBrowseJunction = SetUIElement @p

	$p = @{ type = 'Button'; width = 55; height = 25; top = 300; left = 215; bg = @(35, 175, 75); fg = @(240, 240, 240); id = 'StartJunction'; text = 'Create'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Create Profile`nCreates a lightweight copy (Junction) of your client.`nAllows separate settings for each profile."; noCustomPaint = $true }
	$btnStartJunction = SetUIElement @p

	$p = @{ type = 'Button'; width = 120; height = 25; top = 365; left = 20; bg = @(60, 60, 100); fg = @(240, 240, 240); id = 'CopyNeuz'; text = 'Copy Neuz.exe'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Copy Neuz.exe`nCopies 'neuz.exe' from the main folder to all profiles.`nUse this after patching the main client to update profiles."; noCustomPaint = $true }
	$btnCopyNeuz = SetUIElement @p

	$p = @{ type = 'CheckBox'; width = 200; height = 20; top = 395; left = 0; bg = @(30, 30, 30); fg = @(240, 240, 240); id = 'AutoStartWithWindows'; text = 'Auto Start with Windows'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Auto Start`nEnable this to have the Entropia Dashboard start automatically when you log into Windows." }
	$chkAutoStartWindows = SetUIElement @p

	$p = @{ type = 'CheckBox'; width = 250; height = 20; top = 415; left = 0; bg = @(30, 30, 30); fg = @(240, 240, 240); id = 'AutoStartSetup'; text = 'Auto Start Selected Setup'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Auto Setup`nAutomatically launch the currently selected setup when the Dashboard starts." }
	$chkAutoStartSetup = SetUIElement @p

	$p = @{ type = 'CheckBox'; width = 200; height = 20; top = 435; left = 0; bg = @(30, 30, 30); fg = @(240, 240, 240); id = 'NeverRestartingCollectorLogin'; text = 'Collector Double Click'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Collector Fix`nEnable if the Collector start button needs a double-click to work." }
	$chkNeverRestartingLogin = SetUIElement @p
	
	$p = @{ type = 'Label'; width = 120; height = 20; top = 460; left = 20; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); id = 'LabelNotificationPosition'; text = 'Notification Position:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Notification Position`nSets the corner of the screen where notifications will appear." }
	$lblNotificationPosition = SetUIElement @p
	
	$p = @{ type = 'Custom.DarkComboBox'; width = 145; height = 25; top = 457; left = 145; bg = @(40,40,40); fg = @(240,240,240); id = 'NotificationPositionSelector'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); dropDownStyle = 'DropDownList' }
	$ddlNotificationPosition = SetUIElement @p
	$ddlNotificationPosition.Items.AddRange(@('Bottom Right', 'Bottom Left', 'Top Right', 'Top Left')) | Out-Null

	$p = @{ type = 'Label'; width = 220; height = 20; top = 25; left = 300; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); id = 'LabelProfiles'; text = 'Select Client Profile for Default Launch:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Default Profile`nSelect the profile to be used when clicking the main 'Launch' button." }
	$lblProfiles = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 260; height = 180; top = 50; left = 300; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'ProfileGrid'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$ProfileGrid = SetUIElement @p
	$ProfileGrid.AllowUserToAddRows = $false
	$ProfileGrid.RowHeadersVisible = $false
	$ProfileGrid.EditMode = 'EditProgrammatically'
	$ProfileGrid.SelectionMode = 'FullRowSelect'
	$ProfileGrid.AutoSizeColumnsMode = 'Fill'
	$ProfileGrid.ColumnHeadersHeight = 30
	$ProfileGrid.RowTemplate.Height = 25
	$ProfileGrid.MultiSelect = $false

	$profileGridCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Name'; FillWeight = 20; ReadOnly = $true; ToolTipText = "Profile Name`nThe unique name of this client profile." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Path'; FillWeight = 25; ReadOnly = $true; ToolTipText = "Path`nLocation on disk. Double-click to open folder." }),
		(New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ HeaderText = 'Reconnect?'; FillWeight = 25; ReadOnly = $false; ToolTipText = "Auto-Reconnect`nIf checked, the dashboard attempts to reconnect this profile if it disconnects." }),
		(New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ HeaderText = 'Hide?'; FillWeight = 15; ReadOnly = $false; ToolTipText = "Hide Window`nIf checked, the client window is hidden from Taskbar/Alt-Tab when minimized." })
	)
	$ProfileGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$profileGridCols)

	$p = @{ type = 'Button'; width = 60; height = 25; top = 235; left = 300; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'AddProfile'; text = 'Add'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Add Profile`nRegister an existing game folder as a profile manually."; noCustomPaint = $true }
	$btnAddProfile = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 235; left = 365; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'RenameProfile'; text = 'Rename'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Rename`nChange the display name of the selected profile."; noCustomPaint = $true }
	$btnRenameProfile = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 235; left = 430; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'RemoveProfile'; text = 'Remove'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Remove`nRemove the selected profile from the list (does not delete files)."; noCustomPaint = $true }
	$btnRemoveProfile = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 235; left = 495; bg = @(150, 20, 20); fg = @(240, 240, 240); id = 'DeleteProfile'; text = 'Delete'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Delete`nPermanently delete the selected profile and its files from the disk."; noCustomPaint = $true }
	$btnDeleteProfile = SetUIElement @p

	$p = @{ type = 'Label'; width = 220; height = 20; top = 270; left = 300; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); id = 'LabelSetups'; text = 'Saved Setups:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Saved Setups`nList of pre-configured client launch groups." }
	$lblSetups = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 260; height = 180; top = 295; left = 300; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'SetupGrid'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$SetupGrid = SetUIElement @p
	$SetupGrid.AllowUserToAddRows = $false
	$SetupGrid.RowHeadersVisible = $false
	$SetupGrid.EditMode = 'EditProgrammatically'
	$SetupGrid.SelectionMode = 'FullRowSelect'
	$SetupGrid.AutoSizeColumnsMode = 'Fill'
	$SetupGrid.ColumnHeadersHeight = 30
	$SetupGrid.RowTemplate.Height = 25
	$SetupGrid.MultiSelect = $false
	
	$setupGridCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Setup Name'; FillWeight = 60; ReadOnly = $true; ToolTipText = "Setup Name`nName of the configuration." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Clients'; FillWeight = 40; ReadOnly = $true; ToolTipText = "Client Count`nNumber of clients included in this setup." })
	)
	$SetupGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$setupGridCols)

	$p = @{ type = 'Button'; width = 60; height = 25; top = 480; left = 300; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'AddSetup'; text = 'Add'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Add Setup`nSave the currently selected clients in the main list as a new launch setup."; noCustomPaint = $true }
	$btnAddSetup = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 480; left = 365; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'EditSetup'; text = 'Edit'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Edit Setup`nModify the clients included in the selected setup."; noCustomPaint = $true }
	$btnEditSetup = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 480; left = 430; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'RenameSetup'; text = 'Rename'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Rename`nChange the name of the selected setup."; noCustomPaint = $true }
	$btnRenameSetup = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 480; left = 495; bg = @(150, 20, 20); fg = @(240, 240, 240); id = 'DeleteSetup'; text = 'Delete'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Delete`nRemove the selected setup configuration."; noCustomPaint = $true }
	$btnDeleteSetup = SetUIElement @p

	$tabGeneral.Controls.AddRange(@($lblLauncher, $txtLauncher, $btnBrowseLauncher, $lblProcessName, $txtProcessName, $lblMaxClients, $txtMaxClients, $lblLauncherTimeout, $txtLauncherTimeout, $lblJunction, $txtJunction, $btnBrowseJunction, $btnStartJunction, $btnCopyNeuz, $chkAutoStartWindows, $chkAutoStartSetup, $chkNeverRestartingLogin, $lblNotificationPosition, $ddlNotificationPosition, $lblProfiles, $ProfileGrid, $btnAddProfile, $btnRenameProfile, $btnRemoveProfile, $btnDeleteProfile, $lblSetups, $SetupGrid, $btnAddSetup, $btnEditSetup, $btnRenameSetup, $btnDeleteSetup))

	$uiPropertiesToAdd.NotificationPositionSelector = $ddlNotificationPosition
	#endregion

	#region Login Settings

	$p = @{ type = 'Label'; width = 80; height = 20; top = 10; left = 10; bg = @(30,30,30,0); fg = @(240,240,240); text = 'Edit Profile:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$lblProfSel = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 430; height = 25; top = 7; left = 100; bg = @(40,40,40); fg = @(240,240,240); id = 'LoginProfileSelector'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); dropDownStyle = 'DropDownList' }
	$ddlLoginProfile = SetUIElement @p
	$ddlLoginProfile.Items.Add('Default') | Out-Null

	$ddlLoginProfile.Tag = 'Default'

	if ($global:DashboardConfig.Config['Profiles'])
	{
		foreach ($key in $global:DashboardConfig.Config['Profiles'].Keys)
		{
			$ddlLoginProfile.Items.Add($key) | Out-Null
		}
	}
	$ddlLoginProfile.SelectedIndex = 0

	$tabLoginSettings.Controls.AddRange(@($lblProfSel, $ddlLoginProfile))

	$p = @{ type = 'Label'; width = 75; height = 20; top = 40; left = 10; bg = @(30,30,30,0); fg = @(240,240,240); text = 'Resolution:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$lblResolution = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 180; height = 25; top = 37; left = 85; bg = @(40,40,40); fg = @(240,240,240); id = 'ResolutionSelector'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); dropDownStyle = 'DropDownList' }
	$ddlResolution = SetUIElement @p
    
	$resolutions = @(
		'1024 x 768', '1280 x 720', '1280 x 768', '1280 x 800', '1280 x 1024',
		'1360 x 768', '1440 x 900', '1600 x 815', '1600 x 900', '1600 x 1200',
		'1680 x 1050', '1920 x 1080', '1920 x 1440', '2560 x 1080', '2560 x 1440',
		'3440 x 1440', '3840 x 2160', '4096 x 3072', '5120 x 1440'
	)
	foreach ($res in $resolutions) { $ddlResolution.Items.Add($res) | Out-Null }

	$p = @{ type = 'Label'; width = 60; height = 20; top = 40; left = 285; bg = @(30,30,30,0); fg = @(240,240,240); text = 'DirectX:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$lblDirectX = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 185; height = 25; top = 37; left = 345; bg = @(40,40,40); fg = @(240,240,240); id = 'DirectXSelector'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); dropDownStyle = 'DropDownList' }
	$ddlDirectX = SetUIElement @p
	$dxVersions = @('DirectX 9', 'DirectX 10')
	foreach ($dx in $dxVersions) { $ddlDirectX.Items.Add($dx) | Out-Null }

	$tabLoginSettings.Controls.AddRange(@($lblResolution, $ddlResolution, $lblDirectX, $ddlDirectX))

	$Pickers = @{}
	$rowY = 45

	$AddPickerRow = {
		param($LabelText, $KeyName, $Top, $Col, $Tooltip)
		$leftOffset = if ($Col -eq 2) { 300 } else { 10 }

		$p = @{ type = 'Label'; width = 100; height = 20; top = $Top; left = $leftOffset; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); text = $LabelText; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = $Tooltip }
		$l = SetUIElement @p
		$p = @{ type = 'TextBox'; width = 80; height = 20; top = $Top; left = ($leftOffset + 100); bg = @(40, 40, 40); fg = @(240, 240, 240); id = "txt$KeyName"; text = '0,0'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = $Tooltip }
		$t = SetUIElement @p
		$p = @{ type = 'Button'; width = 40; height = 20; top = $Top; left = ($leftOffset + 185); bg = @(60, 60, 100); fg = @(240, 240, 240); id = "btnPick$KeyName"; text = 'Set'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 7)); tooltip = $Tooltip; noCustomPaint = $true }
		$b = SetUIElement @p
		Register-ButtonFadeEffect -Button $b -NormalStartColor $blueNormalStart -NormalEndColor $blueNormalEnd -HoverStartColor $blueHoverStart -HoverEndColor $blueHoverEnd -LeftPressedStartColor $blueLeftStart -LeftPressedEndColor $blueLeftEnd -RightPressedStartColor $blueRightStart -RightPressedEndColor $blueRightEnd
		return @{ Label = $l; Text = $t; Button = $b }
	}

	$p = @{ type = 'Label'; width = 300; height = 20; top = ($rowY + 20); left = 15; bg = @(30, 30, 30, 0); fg = @(160, 160, 160); text = 'BASIC SETTINGS'; font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)) }
	$tabLoginSettings.Controls.Add((SetUIElement @p))

	$p = &$AddPickerRow 'Launcher Play:' 'LauncherPlay' ($rowY + 40) 1 "Set Coordinate`n1. Click your Launcher window.`n2. Hover over the Play/Start button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['LauncherPlay'] = $p
	$p = &$AddPickerRow 'First Nickname:' 'FirstNick' ($rowY + 40) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the 1st Nickname in the list.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['FirstNick'] = $p

	$p = @{ type = 'Label'; width = 300; height = 20; top = ($rowY + 75); left = 15; bg = @(30, 30, 30, 0); fg = @(160, 160, 160); text = 'ADVANCED SETTINGS'; font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)) }
	$tabLoginSettings.Controls.Add((SetUIElement @p))

	$p = &$AddPickerRow 'Scroll Down:' 'ScrollDown' ($rowY + 95) 1 "Set Coordinate`n1. Click your game window.`n2. Hover slightly above the Scroll Down arrow.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['ScrollDown'] = $p

	$p = &$AddPickerRow 'Server 1:' 'Server1' ($rowY + 120) 1 "Set Coordinate`n1. Click your game window.`n2. Hover over the Server button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Server1'] = $p
	$p = &$AddPickerRow 'Server 2:' 'Server2' ($rowY + 120) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the Server button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Server2'] = $p
	
	$p = &$AddPickerRow 'Channel 1:' 'Channel1' ($rowY + 145) 1 "Set Coordinate`n1. Click your game window.`n2. Hover over the Channel button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Channel1'] = $p
	$p = &$AddPickerRow 'Channel 2:' 'Channel2' ($rowY + 145) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the Channel button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Channel2'] = $p

	$p = &$AddPickerRow 'Char Slot 1:' 'Char1' ($rowY + 170) 1 "Set Coordinate`n1. Click your game window.`n2. Hover over the Character Slot.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Char1'] = $p
	$p = &$AddPickerRow 'Char Slot 2:' 'Char2' ($rowY + 170) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the Character Slot.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Char2'] = $p
	$p = &$AddPickerRow 'Char Slot 3:' 'Char3' ($rowY + 195) 1 "Set Coordinate`n1. Click your game window.`n2. Hover over the Character Slot.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Char3'] = $p
	$p = &$AddPickerRow 'Char Slot 4:' 'Char4' ($rowY + 195) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the Character Slot.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Char4'] = $p
	$p = &$AddPickerRow 'Char Slot 5:' 'Char5' ($rowY + 220) 1 "Set Coordinate`n1. Click your game window.`n2. Hover over the Character Slot.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Char5'] = $p
	$p = &$AddPickerRow 'Char Slot 6:' 'Char6' ($rowY + 220) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the Character Slot.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['Char6'] = $p

	$p = &$AddPickerRow 'Switch Left:' 'SwitchLeft' ($rowY + 245) 1 "Set Coordinate`n1. Click your game window.`n2. Hover over the left arrow in the character selection.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['SwitchLeft'] = $p
	$p = &$AddPickerRow 'Switch Right:' 'SwitchRight' ($rowY + 245) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the right arrow in the character selection.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['SwitchRight'] = $p

	$p = &$AddPickerRow 'Collector:' 'CollectorStart' ($rowY + 270) 1 "Set Coordinate`n1. Click your game window.`n2. Hover over the Collector Start button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['CollectorStart'] = $p

	$p = @{ type = 'Label'; width = 100; height = 20; top = ($rowY + 270); left = 300; bg = @(30, 30, 30, 0); fg = @(240, 240, 240); text = 'Login Delay (s):'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Post-Login Delay`nSeconds to wait after login before performing actions (e.g., Collector)." }
	$lblPostLoginDelay = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 80; height = 20; top = ($rowY + 270); left = 400; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'txtPostLoginDelayInput'; text = '5'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Post-Login Delay`nSeconds to wait after login before performing actions (e.g., Collector)." }
	$txtPostLoginDelay = SetUIElement @p
	$tabLoginSettings.Controls.AddRange(@($lblPostLoginDelay, $txtPostLoginDelay))

	$p = @{ type = 'Label'; width = 300; height = 20; top = ($rowY + 305); left = 15; bg = @(30, 30, 30, 0); fg = @(160, 160, 160); text = 'RECONNECT SETTINGS'; font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)) }
	$tabLoginSettings.Controls.Add((SetUIElement @p))

	$p = &$AddPickerRow 'Login Wrong:' 'LoginDetailsOK' ($rowY + 325) 1 "Set Coordinate`n1. Click your game window.`n2. Hover over the Wrong Password 'OK' button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['LoginDetailsOK'] = $p
	$p = &$AddPickerRow 'Disconnect:' 'DisconnectOK' ($rowY + 325) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the Disconnect 'OK' button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['DisconnectOK'] = $p

	$gridTop = $rowY + 370
	$gridHeight = 285

	$p = @{ type = 'DataGridView'; width = 560; height = $gridHeight; top = $gridTop; left = 10; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'LoginConfigGrid'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$LoginConfigGrid = SetUIElement @p
	$LoginConfigGrid.AllowUserToAddRows = $false
	$LoginConfigGrid.RowHeadersVisible = $false
	$LoginConfigGrid.EditMode = 'EditProgrammatically'
	$LoginConfigGrid.SelectionMode = 'CellSelect'
	$LoginConfigGrid.AutoSizeColumnsMode = 'Fill'
	$LoginConfigGrid.ColumnHeadersHeight = 30
	$LoginConfigGrid.RowTemplate.Height = 25
	$LoginConfigGrid.MultiSelect = $false

	$loginConfigCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Client #'; FillWeight = 15; ReadOnly = $true; ToolTipText = "Client #`nCorresponds to the row number in the main list and the nickname position." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Server'; FillWeight = 20; ReadOnly = $true; ToolTipText = "Server`nSelect the server index (1 or 2) for this client." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Channel'; FillWeight = 20; ReadOnly = $true; ToolTipText = "Channel`nSelect the channel index (1 or 2) for this client." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Character'; FillWeight = 25; ReadOnly = $true; ToolTipText = "Character`nSelect the character slot (1 to 6) to login." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Collecting?'; FillWeight = 20; ReadOnly = $true; ToolTipText = "Collector`nIf 'Yes', the bot will attempt to start the collector after login." })
	)
	$LoginConfigGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$loginConfigCols)

	for ($i = 1; $i -le 10; $i++)
	{
		$LoginConfigGrid.Rows.Add($i, '1', '1', '1', 'No') | Out-Null
	}

	$tabLoginSettings.Controls.Add($LoginConfigGrid)

	#endregion

	#region Hotkeys

	$p = @{ type = 'Panel'; width = 560; height = 50; top = 10; left = 10; bg = @(35, 35, 38); fg = @(240, 240, 240); id = 'MasterHotkeysPanel'; text = '' }
	$MasterHotkeysPanel = SetUIElement @p
	$tabHotkeys.Controls.Add($MasterHotkeysPanel)

	$p = @{ type = 'Label'; width = 115; height = 20; top = 15; left = 10; bg = @(35, 35, 38, 0); fg = @(200, 200, 200); id = 'MasterPauseLabel'; text = 'Master Pause:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$lblMasterPause = SetUIElement @p
	$lblMasterPause.TextAlign = 'MiddleRight'
	$MasterHotkeysPanel.Controls.Add($lblMasterPause)

	$p = @{ type = 'Button'; width = 80; height = 26; top = 12; left = 130; bg = @(45, 45, 48); fg = @(240, 240, 240); id = 'MasterPauseHotkeyBtn'; text = 'none'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Master Pause Hotkey`nAssign a global hotkey to pause all ftool/macro hotkeys at once."; noCustomPaint = $true }
	$btnMasterPauseHotkey = SetUIElement @p
	$MasterHotkeysPanel.Controls.Add($btnMasterPauseHotkey)

	$p = @{ type = 'Label'; width = 115; height = 20; top = 15; left = 220; bg = @(35, 35, 38, 0); fg = @(200, 200, 200); id = 'MasterResumeLabel'; text = 'Master Resume:'; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$lblMasterResume = SetUIElement @p
	$lblMasterResume.TextAlign = 'MiddleRight'
	$MasterHotkeysPanel.Controls.Add($lblMasterResume)

	$p = @{ type = 'Button'; width = 80; height = 26; top = 12; left = 340; bg = @(45, 45, 48); fg = @(240, 240, 240); id = 'MasterResumeHotkeyBtn'; text = 'none'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Master Resume Hotkey`nAssign a global hotkey to resume all ftool/macro hotkeys at once."; noCustomPaint = $true }
	$btnMasterResumeHotkey = SetUIElement @p
	$MasterHotkeysPanel.Controls.Add($btnMasterResumeHotkey)

	$p = @{ type = 'Label'; width = 110; height = 20; top = 15; left = 430; bg = @(35, 35, 38, 0); fg = @(100, 200, 100); id = 'MasterHotkeysStatus'; text = '● Active'; font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)) }
	$lblMasterHotkeysStatus = SetUIElement @p
	$lblMasterHotkeysStatus.TextAlign = 'MiddleCenter'
	$MasterHotkeysPanel.Controls.Add($lblMasterHotkeysStatus)

	$p = @{ type = 'DataGridView'; width = 560; height = 520; top = 70; left = 10; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'HotkeysGrid'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)) }
	$HotkeysGrid = SetUIElement @p
	$HotkeysGrid.AllowUserToAddRows = $false
	$HotkeysGrid.RowHeadersVisible = $false
	$HotkeysGrid.EditMode = 'EditProgrammatically'
	$HotkeysGrid.SelectionMode = 'FullRowSelect'
	$HotkeysGrid.AutoSizeColumnsMode = 'Fill'
	$HotkeysGrid.ColumnHeadersHeight = 35
	$HotkeysGrid.RowTemplate.Height = 25	
	$HotkeysGrid.MultiSelect = $true
	$hotkeysGridCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = "Hotkey`nEditable"; FillWeight = 25; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Assigned To'; FillWeight = 45; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Action'; FillWeight = 30; ReadOnly = $true })
	)
	$HotkeysGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$hotkeysGridCols)
	$tabHotkeys.Controls.Add($HotkeysGrid)

	$p = @{ type = 'Button'; width = 200; height = 40; top = 600; left = 195; bg = @(150, 50, 50); fg = @(240, 240, 240); id = 'UnregisterHotkey'; text = 'Unregister and delete selected Hotkeys permanently'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Unregister Hotkey`nDelete the selected hotkey assignment.`nDouble-click a row to edit the key instead."; noCustomPaint = $true }
	$btnUnregisterHotkey = SetUIElement @p
	$tabHotkeys.Controls.Add($btnUnregisterHotkey)

	#endregion

	$p = @{ type = 'Button'; width = 120; height = 40; top = 735; left = 20; bg = @(35, 175, 75); fg = @(240, 240, 240); id = 'Save'; text = 'Save'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Save`nApply and save all changes to the configuration file."; noCustomPaint = $true }
	$btnSave = SetUIElement @p

	$p = @{ type = 'Button'; width = 120; height = 40; top = 735; left = 150; bg = @(210, 45, 45); fg = @(240, 240, 240); id = 'Cancel'; text = 'Cancel'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Cancel`nDiscard changes and close the settings window."; noCustomPaint = $true }
	$btnCancel = SetUIElement @p


	$settingsForm.Controls.AddRange(@($btnSave, $btnCancel))

	$btnLaunch.Anchor = 'Top, Left'
	$btnLogin.Anchor = 'Top, Left'
	$btnSettings.Anchor = 'Top, Left'
	$btnStop.Anchor = 'Top, Right'
	$btnFtool.Anchor = 'Top, Left, Right'
	$btnExtra.Anchor = 'Top, Right'
	$btnWiki.Anchor = 'Top, Right'
	$DataGridMain.Anchor = 'Top, Bottom, Left, Right'
	$topBar.Anchor = 'Top, Left, Right'

	$mainForm.Controls.AddRange(@($topBar, $btnLogin, $btnFtool, $btnMacro, $btnExtra, $btnWiki, $btnLaunch, $btnSettings, $btnStop, $DataGridMain, $resizeGrip))
	$topBar.Controls.AddRange(@($titleLabelForm, $copyrightLabelForm, $btnInfoForm, $btnMinimizeForm, $btnCloseForm))

	#endregion

	$script:CurrentBuilderToolTip = $toolTipExtra

	#region ExtraForm

	$Theme = @{
		Background   = [System.Drawing.ColorTranslator]::FromHtml("#0f1219")
		PanelColor   = [System.Drawing.ColorTranslator]::FromHtml("#232838")
		InputBack    = [System.Drawing.ColorTranslator]::FromHtml("#161a26")
		AccentColor  = [System.Drawing.ColorTranslator]::FromHtml("#ff2e4c")
		TextColor    = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
		SubTextColor = [System.Drawing.ColorTranslator]::FromHtml("#a0a5b0")
		
		
		ButtonColor     = [System.Drawing.ColorTranslator]::FromHtml("#462424")
	}

	$ExtraForm.BackColor = $Theme.Background

	$p = @{ type = 'Label'; text = 'EXTRA CONTROL PANEL'; font = (New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)); fg = @(255, 46, 76); top = 10; left = 15; bg = @(15, 18, 25) }
	$lblTitle = SetUIElement @p
	$lblTitle.AutoSize = $true
	$lblTitle.Add_MouseDown({ param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.ExtraForm.Handle, 0xA1, 0x2, 0) })

	$p = @{ type = 'Label'; text = 'X'; font = (New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)); fg = @(160, 165, 176); top = 8; left = 370; cursor = 'Hand'; bg = @(15, 18, 25) }
	$btnClose = SetUIElement @p
	$btnClose.Add_Click({ HideExtraForm })

	$extraTabs = New-Object Custom.DarkTabControl
	$extraTabs.Location = New-Object System.Drawing.Point(0, 35)
	$extraTabs.Size = New-Object System.Drawing.Size(400, 560)
	$extraTabs.Anchor = 'Top, Bottom, Left, Right'
	$extraTabs.ItemSize = New-Object System.Drawing.Size(99, 30)

	$tabWorldBoss = New-Object System.Windows.Forms.TabPage 'World Boss'
	$tabWorldBoss.BackColor = $Theme.Background

	$tabNotes = New-Object System.Windows.Forms.TabPage 'Notes'
	$tabNotes.BackColor = $Theme.Background

	$tabNotifications = New-Object System.Windows.Forms.TabPage 'Notifications'
	$tabNotifications.BackColor = $Theme.Background

	$tabTools = New-Object System.Windows.Forms.TabPage 'Tools'
	$tabTools.BackColor = $Theme.Background

	$extraTabs.TabPages.Add($tabWorldBoss)
	$extraTabs.TabPages.Add($tabNotes)
	$extraTabs.TabPages.Add($tabTools)

	#region World Boss

	$BossList = @(
		"Entropia King",
		"Clockworks",
		"Royal Knight",
		"Solarion",
		"C-A01",
		"Great Venux Tree",
		"Grinch"
	)

	$p = @{ type = 'Label'; text = 'ACCESS CODE'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); fg = @(160, 165, 176); top = 10; left = 15; bg = @(15, 18, 25); tooltip = "Access Code`nEnter a valid code to send/receive World Boss alerts.`nTo get an acces code, you must be member of the Divine Discord.`nUse the Access Code TEST to test this feature out locally." }
	$lblCode = SetUIElement @p
	$lblCode.AutoSize = $true

	$p = @{ type = 'Label'; text = 'USERNAME'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); fg = @(160, 165, 176); top = 10; left = 200; bg = @(15, 18, 25); tooltip = "Username`nYour name will be shown in Discord alerts you trigger." }
	$lblUser = SetUIElement @p
	$lblUser.AutoSize = $true

	$p = @{ type = 'TextBox'; width = 175; height = 25; top = 30; left = 15; bg = @(22, 26, 38); fg = @(255, 255, 255); font = (New-Object System.Drawing.Font('Segoe UI', 10)); tooltip = "Access Code`nEnter a valid code to send/receive World Boss alerts.`nTo get an acces code, you must be member of the Divine Discord.`nUse the Access Code TEST to test this feature out locally." }
	$codeBox = SetUIElement @p
	$codeBox.PasswordChar = '*'
	$codeBox.TextAlign = 'Left'
	
	if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['AccessCode']) {
		$codeBox.Text = $global:DashboardConfig.Config['Options']['AccessCode']
	}
	$codeBox.Add_TextChanged({
		if (-not $global:DashboardConfig.Config.Contains('Options')) { $global:DashboardConfig.Config['Options'] = [ordered]@{} }
		$global:DashboardConfig.Config['Options']['AccessCode'] = $this.Text
	})
	$codeBox.Add_Leave({ if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig } })

	$p = @{ type = 'TextBox'; width = 185; height = 25; top = 30; left = 200; bg = @(22, 26, 38); fg = @(255, 255, 255); font = (New-Object System.Drawing.Font('Segoe UI', 10)); tooltip = "Username`nYour name will be shown in Discord alerts you trigger." }
	$userBox = SetUIElement @p
	$userBox.TextAlign = 'Left'

	if ($global:DashboardConfig.Config['User'] -and $global:DashboardConfig.Config['User']['Username']) {
		$userBox.Text = $global:DashboardConfig.Config['User']['Username']
	}

	$userBox.Add_TextChanged({
		if (-not $global:DashboardConfig.Config.Contains('User')) { $global:DashboardConfig.Config['User'] = [ordered]@{} }
		$global:DashboardConfig.Config['User']['Username'] = $this.Text
	})
	$userBox.Add_Leave({ if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig } })

	$p = @{ type = 'Label'; text = 'Show Images'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); fg = @(160, 165, 176); top = 60; left = 15; bg = @(15, 18, 25); tooltip = "Show Images`nToggle displaying boss images on buttons and in toast notifications." }
	$lblShowImages = SetUIElement @p
	$lblShowImages.AutoSize = $true

	$p = @{ type = 'Toggle'; width = 30; height = 17; top = 59; left = 90; bg = @(40, 80, 80); fg = @(255, 255, 255); cursor = 'Hand'; tooltip = "Show Images`nToggle displaying boss images on buttons and in toast notifications." }
	$chkShowImages = SetUIElement @p
	$chkShowImages.Name = 'ShowBossImages'

	$p = @{ type = 'Label'; text = 'World Boss Listener'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); fg = @(160, 165, 176); top = 60; left = 140; bg = @(15, 18, 25); tooltip = "Listener`nEnable to receive network alerts when a World Boss spawns." }
	$lblWorldBossListener = SetUIElement @p
	$lblWorldBossListener.AutoSize = $true

	$p = @{ type = 'Toggle'; width = 30; height = 17; top = 59; left = 255; bg = @(40, 80, 80); fg = @(255, 255, 255); cursor = 'Hand'; tooltip = "Listener`nEnable to receive network alerts when a World Boss spawns." }
	$chkWorldBossListener = SetUIElement @p
	$chkWorldBossListener.Name = 'WorldBossListener'

	$p = @{ type = 'Button'; width = 80; height = 20; top = 57; left = 295; bg = @(40, 40, 40); fg = @(255, 255, 255); id = 'RefreshBosses'; text = 'Update '+[char]0x21BB; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Refresh Schedule`nTrigger the daily schedule notification again with updated data."; noCustomPaint = $true }
	$btnRefreshBosses = SetUIElement @p

	$p = @{ type = 'Label'; text = 'KHELDOR'; font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)); top = 85; left = 60; width = 140 }
	$lblNormalHeader = SetUIElement @p
	$lblNormalHeader.ForeColor = $Theme.TextColor
	$lblNormalHeader.BackColor = $Theme.Background
	$lblNormalHeader.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter


	$p = @{ type = 'Label'; text = 'Listen'; font = (New-Object System.Drawing.Font('Segoe UI', 7)); top = 105; left = 15; width = 44; tooltip = "Listen`nToggle notifications for this specific boss." }
	$lblListenNormal = SetUIElement @p
	$lblListenNormal.ForeColor = $Theme.SubTextColor
	$lblListenNormal.BackColor = $Theme.Background
	$lblListenNormal.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

	$p = @{ type = 'Label'; text = 'CLICK TO REPORT'; font = (New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)); top = 105; left = 59; width = 140; tooltip = "Report`nClick the button below to report this boss as spawned (Notifies everyone!)." }
	$lblReportNormal = SetUIElement @p
	$lblReportNormal.ForeColor = $Theme.AccentColor
	$lblReportNormal.BackColor = $Theme.Background
	$lblReportNormal.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter


	$ButtonPanel = New-Object System.Windows.Forms.TableLayoutPanel
	$ButtonPanel.Location = New-Object System.Drawing.Point(15, 125) 
	$ButtonPanel.Size = New-Object System.Drawing.Size(370, 380) 
	$ButtonPanel.BackColor = [System.Drawing.Color]::Transparent
	$ButtonPanel.ColumnCount = 4
	$ButtonPanel.AutoScroll = $true 
	
	$ButtonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 12)))
	$ButtonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 38)))
	$ButtonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 12)))
	$ButtonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 38)))
	
	$percentage = 100 / $BossList.Count
	foreach ($i in 1..$BossList.Count) {
		$ButtonPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $percentage)))
	}

	if (-not $global:DashboardConfig.Resources.ActiveDownloads) { $global:DashboardConfig.Resources.ActiveDownloads = [System.Collections.ArrayList]::new() }

	foreach ($Boss in $BossList) {
		
		$bossNameNormal = $Boss
		$p = @{ type = 'Toggle'; width = 30; height = 17; bg = @(15, 18, 25); fg = @(255, 255, 255); cursor = 'Hand'; tooltip = "Toggle Listen`nEnable/Disable alerts for $bossNameNormal." }
		$tglNormal = SetUIElement @p
		$tglNormal.Anchor = [System.Windows.Forms.AnchorStyles]::None
		$tglNormal.Checked = $true
		if ($global:DashboardConfig.Config['BossFilter'] -and $global:DashboardConfig.Config['BossFilter'].Contains($bossNameNormal)) {
			$tglNormal.Checked = [bool]::Parse($global:DashboardConfig.Config['BossFilter'][$bossNameNormal])
		}
		$tglNormal.Add_Click({
			param($s, $e)
			if (-not $global:DashboardConfig.Config.Contains('BossFilter')) { $global:DashboardConfig.Config['BossFilter'] = [ordered]@{} }
			$global:DashboardConfig.Config['BossFilter'][$bossNameNormal] = $s.Checked
			if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
		}.GetNewClosure())

		$p = @{ type = 'Button'; text = $Boss; height = 50; bg = @(70, 36, 36); fg = @(255, 255, 255); fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI Semibold', 9)); cursor = 'Hand'; dock = 'Fill'; noCustomPaint = $true }
		$btnNormal = SetUIElement @p
		$btnNormal.Tag = 'Normal'
		$btnNormal.Margin = New-Object System.Windows.Forms.Padding(2)
		$btnNormal.FlatAppearance.BorderSize = 0
		$btnNormal.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Empty
		$btnNormal.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Empty

		$btnNormal.Add_Click({
			if ([string]::IsNullOrWhiteSpace($codeBox.Text)) {
				Show-DarkMessageBox 'Access Code Required' 'Wrong Access Code' 'Ok' 'Error'
				return
			}
			if ([string]::IsNullOrWhiteSpace($userBox.Text)) {
				Show-DarkMessageBox 'Username Required' 'Missing Username' 'Ok' 'Error'
				return
			}
			$bName = $this.Tag['BossName']
			if ((Show-DarkMessageBox "Are you sure you want to report '$bName'?`nThis will notify everyone!" "Confirm Report" "YesNo" "Question") -eq 'Yes') {
				Send-Message -code $codeBox.Text -bossName $bName
			}
		}.GetNewClosure())
		
		$btnNormal.Tag = @{ Type = 'Normal'; BossName = $bossNameNormal }
		Update-BossButtonImage $btnNormal
		
		$ButtonPanel.Controls.Add($tglNormal)
		$ButtonPanel.Controls.Add($btnNormal)
	}

	if (-not $global:DashboardConfig.Resources.BossSchedule) {
		$global:DashboardConfig.Resources.BossSchedule = @{
			"Entropia King" = [TimeSpan]::Parse("18:30")
			"Solarion"      = [TimeSpan]::Parse("19:45")
			"Clockworks"    = [TimeSpan]::Parse("20:00")
			"Royal Knight"  = [TimeSpan]::Parse("20:15")
		}
	}
	$BossSchedule = $global:DashboardConfig.Resources.BossSchedule

	try { $tz = [TimeZoneInfo]::FindSystemTimeZoneById("W. Europe Standard Time") } catch { $tz = [TimeZoneInfo]::Local }
	$fontTimerImage = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
	$fontTimerNoImage = New-Object System.Drawing.Font('Segoe UI Semibold', 7)

	$scheduledItems = [System.Collections.ArrayList]::new()
	if ($ButtonPanel) {
		foreach ($btn in $ButtonPanel.Controls) {
			if ($btn -is [System.Windows.Forms.Button]) {
				$bName = if ($btn.Tag -is [System.Collections.IDictionary]) { $btn.Tag['BossName'] } else { $btn.Text }
				if ($BossSchedule.ContainsKey($bName)) {
					$scheduledItems.Add(@{ Button = $btn; BossName = $bName }) | Out-Null
				}
			}
		}
	}

	if ($global:DashboardConfig.Resources.Timers['BossTimer']) {
		$global:DashboardConfig.Resources.Timers['BossTimer'].Stop()
		$global:DashboardConfig.Resources.Timers['BossTimer'].Dispose()
	}

	$bossTimer = New-Object System.Windows.Forms.Timer
	$bossTimerTick = {
		$nextInterval = 60000
		try {
			$this.Stop()
			if (-not $global:DashboardConfig.UI.MainForm -or $global:DashboardConfig.UI.MainForm.IsDisposed) { return }

			$serverTime = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)

			foreach ($item in $scheduledItems) {
				$btn = $item.Button
				if ($btn.IsDisposed) { continue }
				
				$bName = $item.BossName
				if (-not $BossSchedule.ContainsKey($bName)) { continue }
				
				$spawnTime = $BossSchedule[$bName]
				$todaySpawn = $serverTime.Date.Add($spawnTime)
				
				if ($serverTime -gt $todaySpawn) { $nextSpawn = $todaySpawn.AddDays(1) } 
				else { $nextSpawn = $todaySpawn }
				
				$diff = $nextSpawn - $serverTime
				$timeStr = $diff.ToString("hh\:mm")
				$fixedtime = $spawnTime.ToString("hh\:mm")
				
				$hasImage = ($null -ne $btn.BackgroundImage)
				if ($hasImage) {
					$newText = "Fixed Spawn in $($timeStr)h ($($fixedtime))"
					if ($btn.Text -ne $newText) {
						$btn.Text = $newText
						$btn.TextAlign = [System.Drawing.ContentAlignment]::BottomCenter
						$btn.ForeColor = [System.Drawing.Color]::White
						if ($btn.Font -ne $fontTimerImage) { $btn.Font = $fontTimerImage }
					}
				} else {
					$newText = "$bName`nFixed Spawn in $($timeStr)h`n($($fixedtime))"
					if ($btn.Text -ne $newText) {
						$btn.Text = $newText
						$btn.TextAlign = [System.Drawing.ContentAlignment]::BottomCenter
						$btn.ForeColor = [System.Drawing.Color]::White
						if ($btn.Font -ne $fontTimerNoImage) { $btn.Font = $fontTimerNoImage }
					}
				}
			}

			$now = [DateTime]::Now
			$msToNextMinute = (60 - $now.Second) * 1000 - $now.Millisecond
			if ($msToNextMinute -le 0) { $msToNextMinute += 60000 }
			
			$this.Interval = $msToNextMinute
			$nextInterval = $msToNextMinute
		} 
		catch { 
			Write-Verbose "Boss timer tick failed: $_" 
			$nextInterval = 15000
		} 
		finally {
			if ($this.Interval -ne $nextInterval) { $this.Interval = $nextInterval }
			$this.Start()
		}
	}
	$bossTimer.Add_Tick($bossTimerTick.GetNewClosure())
	$bossTimer.Interval = 2000
	$bossTimer.Start()
	$global:DashboardConfig.Resources.Timers['BossTimer'] = $bossTimer

	$tabWorldBoss.Controls.AddRange(@(
		$lblCode,
		$codeBox,
		$lblUser,
		$userBox,
		$lblShowImages,
		$chkShowImages,
		$lblWorldBossListener,
		$chkWorldBossListener,
		$btnRefreshBosses,
		$lblNormalHeader,
		$lblListenNormal,
		$lblReportNormal,
		$ButtonPanel
	))

	#endregion

	#region Notes

	$p = @{ type = 'DataGridView'; width = 370; height = 450; top = 10; left = 15; bg = @(22, 26, 38); fg = @(255, 255, 255); id = 'NoteGrid'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "Notes`nYour saved notes and timers.`nDouble-click a row to edit." }
	$NoteGrid = SetUIElement @p
	$NoteGrid.AllowUserToAddRows = $false
	$NoteGrid.RowHeadersVisible = $false
	$NoteGrid.SelectionMode = 'FullRowSelect'
	$NoteGrid.MultiSelect = $false
	$NoteGrid.AutoSizeColumnsMode = 'Fill'
	$NoteGrid.ColumnHeadersHeight = 30
	$NoteGrid.RowTemplate.Height = 25

	$noteCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Title'; FillWeight = 40; ReadOnly = $true; ToolTipText = "Title`nSummary of the note." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Type'; FillWeight = 20; ReadOnly = $true; ToolTipText = "Type`nNote, Timer, or Reminder." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Due / Content'; FillWeight = 40; ReadOnly = $true; ToolTipText = "Content`nNote text or due time." })
	)
	$NoteGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$noteCols)

	$p = @{ type = 'Button'; width = 80; height = 25; top = 470; left = 15; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'AddNote'; text = 'Add'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Add Note`nCreate a new note, timer, or reminder."; noCustomPaint = $true }
	$btnAddNote = SetUIElement @p

	$p = @{ type = 'Button'; width = 80; height = 25; top = 470; left = 105; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'EditNote'; text = 'Edit'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Edit Note`nModify the selected note."; noCustomPaint = $true }
	$btnEditNote = SetUIElement @p

	$p = @{ type = 'Button'; width = 80; height = 25; top = 470; left = 195; bg = @(150, 50, 50); fg = @(240, 240, 240); id = 'RemoveNote'; text = 'Remove'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Remove Note`nDelete the selected note permanently."; noCustomPaint = $true }
	$btnRemoveNote = SetUIElement @p

	$tabNotes.Controls.AddRange(@(
		$NoteGrid,
		$btnAddNote,
		$btnEditNote,
		$btnRemoveNote
	))

	#endregion

	#region Notifications

	$p = @{ type = 'DataGridView'; width = 370; height = 450; top = 10; left = 15; bg = @(22, 26, 38); fg = @(255, 255, 255); id = 'NotificationGrid'; text = ''; font = (New-Object System.Drawing.Font('Segoe UI', 9)); tooltip = "History`nLog of recent notifications.`nDouble-click to view details." }
	$NotificationGrid = SetUIElement @p
	$NotificationGrid.AllowUserToAddRows = $false
	$NotificationGrid.RowHeadersVisible = $false
	$NotificationGrid.SelectionMode = 'FullRowSelect'
	$NotificationGrid.MultiSelect = $false
	$NotificationGrid.AutoSizeColumnsMode = 'Fill'
	$NotificationGrid.ColumnHeadersHeight = 30
	$NotificationGrid.RowTemplate.Height = 25

	$notificationCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Time'; FillWeight = 25; ReadOnly = $true; ToolTipText = "Time`nTimestamp of the event." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Type'; FillWeight = 20; ReadOnly = $true; ToolTipText = "Type`nSeverity (Info, Warning, Error)." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Details'; FillWeight = 55; ReadOnly = $true; ToolTipText = "Details`nMessage content." })
	)
	$NotificationGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$notificationCols)

	$p = @{ type = 'Button'; width = 120; height = 25; top = 470; left = 15; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'ShowNotification'; text = 'Show Notification'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Show`nRe-display the selected notification toast."; noCustomPaint = $true }
	$btnShowNotification = SetUIElement @p

	$p = @{ type = 'Button'; width = 130; height = 25; top = 470; left = 145; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'HideAllNotifications'; text = 'Hide all notifications'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Hide All`nClose all currently visible notification toasts."; noCustomPaint = $true }
	$btnHideAllNotifications = SetUIElement @p

	$tabNotifications.Controls.AddRange(@($NotificationGrid, $btnShowNotification, $btnHideAllNotifications))
	$extraTabs.TabPages.Add($tabNotifications)

	#endregion

	#region Tools

	$toolsTabs = New-Object Custom.DarkTabControl
	$toolsTabs.Location = New-Object System.Drawing.Point(0, 0)
	$toolsTabs.Size = New-Object System.Drawing.Size(395, 525)
	$toolsTabs.Anchor = 'Top, Bottom, Left, Right'

	$tabUpgradeCalc = New-Object System.Windows.Forms.TabPage 'Upgrade Calc'
	$tabUpgradeCalc.BackColor = $Theme.Background

	$tabChestOpener = New-Object System.Windows.Forms.TabPage 'Chest Opener'
	$tabChestOpener.BackColor = $Theme.Background

	$toolsTabs.TabPages.Add($tabUpgradeCalc)
	$toolsTabs.TabPages.Add($tabChestOpener)
	$tabTools.Controls.Add($toolsTabs)

	$p = @{ type = 'Label'; width = 150; text = 'UPGRADE CALCULATOR'; font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)); top = 10; left = 15; fg = @(160, 165, 176); bg = @(15, 18, 25) }
	$lblCalcTitle = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 120; height = 25; top = 35; left = 15; bg = @(40,40,40); fg = @(240,240,240); id = 'UpgradeItemSelector'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Item Category`nSelect the type of item you want to upgrade." }
	$ddlUpgradeItem = SetUIElement @p
	foreach ($key in $global:DashboardConfig.Resources.UpgradeData.psobject.Properties.Name) { $ddlUpgradeItem.Items.Add($key) | Out-Null }

	$p = @{ type = 'Custom.DarkComboBox'; width = 60; height = 25; top = 35; left = 140; bg = @(40,40,40); fg = @(240,240,240); id = 'UpgradeTypeSelector'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Weapon Type`n1H or 2H (2H weapons cost double)." }
	$ddlUpgradeType = SetUIElement @p
	$ddlUpgradeType.Items.AddRange(@('1H', '2H')) | Out-Null

	$p = @{ type = 'Label'; text = 'Lvl:'; font = (New-Object System.Drawing.Font('Segoe UI', 7)); top = 38; left = 205; width = 20; fg = @(160, 165, 176) }
	$lblLvlStart = SetUIElement @p
	$p = @{ type = 'Custom.DarkComboBox'; width = 45; height = 25; top = 35; left = 225; bg = @(40,40,40); fg = @(240,240,240); id = 'UpgradeLvlStart'; font = (New-Object System.Drawing.Font('Segoe UI', 8)) }
	$ddlUpgradeLvlStart = SetUIElement @p

	$p = @{ type = 'Label'; text = 'To:'; font = (New-Object System.Drawing.Font('Segoe UI', 7)); top = 38; left = 275; width = 20; fg = @(160, 165, 176) }
	$lblLvlEnd = SetUIElement @p
	$p = @{ type = 'Custom.DarkComboBox'; width = 45; height = 25; top = 35; left = 295; bg = @(40,40,40); fg = @(240,240,240); id = 'UpgradeLvlEnd'; font = (New-Object System.Drawing.Font('Segoe UI', 8)) }
	$ddlUpgradeLvlEnd = SetUIElement @p

	$p = @{ type = 'Button'; width = 350; height = 25; top = 65; left = 15; bg = @(0, 120, 215); fg = @(255, 255, 255); id = 'AddUpgrade'; text = 'Add to List'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)); noCustomPaint = $true }
	$btnAddUpgrade = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 350; height = 120; top = 98; left = 15; bg = @(22, 26, 38); fg = @(255, 255, 255); id = 'UpgradeCartGrid'; font = (New-Object System.Drawing.Font('Segoe UI', 8)) }
	$UpgradeCartGrid = SetUIElement @p
	$UpgradeCartGrid.ColumnHeadersHeight = 25
	$UpgradeCartGrid.MultiSelect = $false
	$UpgradeCartGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect

	$cartCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = '#'; FillWeight = 10; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Item'; FillWeight = 40; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Range'; FillWeight = 50; ReadOnly = $true })
	)
	$UpgradeCartGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$cartCols)

	$p = @{ type = 'Button'; width = 170; height = 25; top = 225; left = 15; bg = @(150, 50, 50); fg = @(255, 255, 255); id = 'RemoveUpgrade'; text = 'Remove'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); noCustomPaint = $true }
	$btnRemoveUpgrade = SetUIElement @p
	$p = @{ type = 'Button'; width = 170; height = 25; top = 225; left = 195; bg = @(40, 40, 40); fg = @(255, 255, 255); id = 'ResetUpgrade'; text = 'Reset All'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); noCustomPaint = $true }
	$btnResetUpgrade = SetUIElement @p

	$p = @{ type = 'Label'; width = 150; text = 'TOTAL MATERIALS REQUIRED'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)); top = 258; left = 15; fg = @(160, 165, 176) }
	$lblMaterialsTitle = SetUIElement @p
	$p = @{ type = 'CheckBox'; width = 150; height = 20; top = 258; left = 225; bg = @(15, 18, 25); fg = @(160, 165, 176); id = 'UpgradeHideZero'; text = 'Hide Zeros'; font = (New-Object System.Drawing.Font('Segoe UI', 7)); checked = $true }
	$chkUpgradeHideZero = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 350; height = 150; top = 282; left = 15; bg = @(22, 26, 38); fg = @(255, 255, 255); id = 'UpgradeResultsGrid'; font = (New-Object System.Drawing.Font('Consolas', 9)) }
	$UpgradeResultsGrid = SetUIElement @p
	$UpgradeResultsGrid.ColumnHeadersVisible = $false
	$resCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Material'; FillWeight = 70; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Count'; FillWeight = 30; ReadOnly = $true })
	)
	$UpgradeResultsGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$resCols)
	$UpgradeResultsGrid.DefaultCellStyle.Alignment = 'MiddleLeft'

	$p = @{ type = 'Label'; width = 170; height = 20; top = 440; left = 15; bg = @(15, 18, 25); fg = @(97, 175, 239); id = 'UpgradeTotalPerin'; text = 'Total Perin: 0'; font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)) }
	$lblTotalPerin = SetUIElement @p
	$p = @{ type = 'Label'; width = 170; height = 20; top = 440; left = 195; bg = @(15, 18, 25); fg = @(152, 195, 121); id = 'UpgradeTotalEcp'; text = 'Total ECP: 0'; font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)) }
	$lblTotalEcp = SetUIElement @p

	$tabUpgradeCalc.Controls.AddRange(@(
		$lblCalcTitle,
		$ddlUpgradeItem, $ddlUpgradeType, $lblLvlStart, $ddlUpgradeLvlStart, $lblLvlEnd, $ddlUpgradeLvlEnd,
		$btnAddUpgrade,
		$UpgradeCartGrid,
		$btnRemoveUpgrade, $btnResetUpgrade,
		$lblMaterialsTitle, $chkUpgradeHideZero,
		$UpgradeResultsGrid,
		$lblTotalPerin, $lblTotalEcp
	))

	$p = @{ type = 'Label'; width = 180; text = 'CHEST OPENER'; font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)); top = 10; left = 15; fg = @(160, 165, 176); bg = @(15, 18, 25) }
	$lblChestTitle = SetUIElement @p

	$p = @{ type = 'Label'; text = 'TARGET WINDOW'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); height = 15; fg = @(160, 165, 176); top = 35; left = 15; bg = @(15, 18, 25); tooltip = "Target Window`nSelect the game client instance to execute chest clicks." }
	$lblChestWindow = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 350; height = 25; top = 52; left = 15; bg = @(40,40,40); fg = @(240,240,240); id = 'ChestWindowSelector'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); dropDownStyle = 'DropDownList'; tooltip = "Target Window`nSelect the game client instance to execute chest clicks." }
	$ddlChestWindow = SetUIElement @p

	$p = @{ type = 'Label'; text = 'CHEST PRESET'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); height = 15; fg = @(160, 165, 176); top = 85; left = 15; bg = @(15, 18, 25); tooltip = "Chest Preset`nSelect the chest type from database presets." }
	$lblChestPreset = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 350; height = 25; top = 102; left = 15; bg = @(40,40,40); fg = @(240,240,240); id = 'ChestPresetSelector'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); dropDownStyle = 'DropDownList'; tooltip = "Chest Preset`nSelect the chest type from database presets." }
	$ddlChestPreset = SetUIElement @p

	$p = @{ type = 'Panel'; width = 350; height = 90; top = 125; left = 15; bg = @(22, 26, 38); id = 'ChestRewardsPanel' }
	$pnlChestRewards = SetUIElement @p
	$pnlChestRewards.AutoScroll = $true

	$p = @{ type = 'Label'; text = 'CHEST KEY:'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)); width = 65; fg = @(160, 165, 176); top = 225; left = 15; bg = @(15, 18, 25); tooltip = "Chest Key`nKey pressed in-game to open the chest selection window before reward clicks." }
	$lblChestItemHotkey = SetUIElement @p

	$p = @{ type = 'Button'; width = 90; height = 24; top = 222; left = 85; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'ChestItemHotkeyBtn'; text = 'Hotkey'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)); tooltip = "Chest Key`nKey pressed in-game to open the chest selection window before reward clicks."; noCustomPaint = $true }
	$btnChestItemHotkey = SetUIElement @p

	$p = @{ type = 'Label'; text = 'START/STOP:'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)); width = 70; fg = @(160, 165, 176); top = 225; left = 185; bg = @(15, 18, 25); tooltip = "Toggle Hotkey`nKeyboard shortcut to start/stop opening chests." }
	$lblChestHotkey = SetUIElement @p

	$p = @{ type = 'Button'; width = 105; height = 24; top = 222; left = 260; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'ChestHotkeyBtn'; text = 'Hotkey'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)); tooltip = "Toggle Hotkey`nKeyboard shortcut to start/stop opening chests."; noCustomPaint = $true }
	$btnChestHotkey = SetUIElement @p

	$p = @{ type = 'Label'; text = 'AMOUNT:'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)); width = 55; fg = @(160, 165, 176); top = 255; left = 15; bg = @(15, 18, 25); tooltip = "Loop Count`nNumber of times to open the chest in sequence." }
	$lblChestCount = SetUIElement @p

	$p = @{ type = 'TextBox'; width = 50; height = 24; top = 252; left = 75; bg = @(40, 40, 40); fg = @(240, 240, 240); id = 'ChestCountInput'; text = '1'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Loop Count`nNumber of times to open the chest in sequence." }
	$txtChestCount = SetUIElement @p

	$p = @{ type = 'Label'; width = 230; height = 24; top = 252; left = 135; fg = @(100, 200, 100); bg = @(15, 18, 25); id = 'ChestStatusLabel'; text = 'Status: IDLE'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)) }
	$lblChestStatus = SetUIElement @p

	$p = @{ type = 'Button'; width = 165; height = 24; top = 282; left = 15; bg = @(0, 120, 215); fg = @(255, 255, 255); id = 'AddChestQueue'; text = '+ Add to Queue'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)); tooltip = "Add Chest to Queue`nAdd the current chest preset, reward choices, key, and amount to execution queue."; noCustomPaint = $true }
	$btnAddChestQueue = SetUIElement @p

	$p = @{ type = 'Button'; width = 85; height = 24; top = 282; left = 185; bg = @(150, 50, 50); fg = @(255, 255, 255); id = 'RemoveChestQueue'; text = 'Remove'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Remove selected chest from queue."; noCustomPaint = $true }
	$btnRemoveChestQueue = SetUIElement @p

	$p = @{ type = 'Button'; width = 90; height = 24; top = 282; left = 275; bg = @(40, 40, 40); fg = @(255, 255, 255); id = 'ResetChestQueue'; text = 'Clear All'; fs = 'Flat'; font = (New-Object System.Drawing.Font('Segoe UI', 8)); tooltip = "Clear all chests from queue."; noCustomPaint = $true }
	$btnResetChestQueue = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 350; height = 155; top = 312; left = 15; bg = @(22, 26, 38); fg = @(255, 255, 255); id = 'ChestQueueGrid'; font = (New-Object System.Drawing.Font('Segoe UI', 8)) }
	$ChestQueueGrid = SetUIElement @p
	$ChestQueueGrid.ColumnHeadersHeight = 25
	$ChestQueueGrid.MultiSelect = $false
	$ChestQueueGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect

	$queueCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = '#'; FillWeight = 10; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Chest'; FillWeight = 50; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Key'; FillWeight = 15; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Amount'; FillWeight = 15; ReadOnly = $true })
	)
	$ChestQueueGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$queueCols)

	$tabChestOpener.Controls.AddRange(@(
		$lblChestTitle,
		$lblChestWindow, $ddlChestWindow,
		$lblChestPreset, $ddlChestPreset,
		$pnlChestRewards,
		$lblChestItemHotkey, $btnChestItemHotkey,
		$lblChestHotkey, $btnChestHotkey,
		$lblChestCount, $txtChestCount,
		$lblChestStatus,
		$btnAddChestQueue, $btnRemoveChestQueue, $btnResetChestQueue,
		$ChestQueueGrid
	))

	$uiPropertiesToAdd.ChestWindowSelector  = $ddlChestWindow
	$uiPropertiesToAdd.ChestPresetSelector  = $ddlChestPreset
	$uiPropertiesToAdd.ChestRewardsPanel    = $pnlChestRewards
	$uiPropertiesToAdd.ChestCountInput      = $txtChestCount
	$uiPropertiesToAdd.ChestItemHotkeyBtn   = $btnChestItemHotkey
	$uiPropertiesToAdd.ChestHotkeyBtn       = $btnChestHotkey
	$uiPropertiesToAdd.ChestHotkeySelector  = $btnChestHotkey
	$uiPropertiesToAdd.ChestStatusLabel     = $lblChestStatus
	$uiPropertiesToAdd.ChestQueueGrid       = $ChestQueueGrid
	$uiPropertiesToAdd.AddChestQueue         = $btnAddChestQueue
	$uiPropertiesToAdd.RemoveChestQueue      = $btnRemoveChestQueue
	$uiPropertiesToAdd.ResetChestQueue       = $btnResetChestQueue
	if (-not $global:DashboardConfig.Resources.ChestQueue) { $global:DashboardConfig.Resources.ChestQueue = [System.Collections.ArrayList]::new() }
	$uiPropertiesToAdd.ChestQueue            = $global:DashboardConfig.Resources.ChestQueue

	$uiPropertiesToAdd.UpgradeItemSelector = $ddlUpgradeItem
	$uiPropertiesToAdd.UpgradeTypeSelector = $ddlUpgradeType
	$uiPropertiesToAdd.UpgradeLvlStart = $ddlUpgradeLvlStart
	$uiPropertiesToAdd.UpgradeLvlEnd = $ddlUpgradeLvlEnd
	$uiPropertiesToAdd.UpgradeCartGrid = $UpgradeCartGrid
	$uiPropertiesToAdd.UpgradeResultsGrid = $UpgradeResultsGrid
	$uiPropertiesToAdd.UpgradeTotalPerin = $lblTotalPerin
	$uiPropertiesToAdd.UpgradeTotalEcp = $lblTotalEcp
	$uiPropertiesToAdd.UpgradeHideZero = $chkUpgradeHideZero
	if (-not $global:DashboardConfig.Resources.UpgradeCart) { $global:DashboardConfig.Resources.UpgradeCart = [System.Collections.ArrayList]::new() }
	$uiPropertiesToAdd.UpgradeCart = $global:DashboardConfig.Resources.UpgradeCart
	$uiPropertiesToAdd.AddUpgrade = $btnAddUpgrade
	$uiPropertiesToAdd.RemoveUpgrade = $btnRemoveUpgrade
	$uiPropertiesToAdd.ResetUpgrade = $btnResetUpgrade

	$ddlUpgradeItem.Add_SelectedIndexChanged({
		UpdateUpgradeItemSelection
	}.GetNewClosure())

	$ddlUpgradeType.Add_SelectedIndexChanged({
		UpdateUpgradeTypeSelection
	}.GetNewClosure())

	$btnAddUpgrade.Add_Click({
		AddUpgradeItem
	}.GetNewClosure())

	$btnRemoveUpgrade.Add_Click({
		RemoveUpgradeItem
	}.GetNewClosure())

	$btnResetUpgrade.Add_Click({
		ResetUpgradeCalculator
	}.GetNewClosure())

	$chkUpgradeHideZero.Add_CheckedChanged({ 
		CalculateUpgradeTotalMaterials 
	}.GetNewClosure())

    $UpgradeCartGrid.Add_RowsAdded({ 
		HandleUpgradeCartRowsAdded
	}.GetNewClosure())

	$ddlChestPreset.Add_SelectedIndexChanged({
		if (Get-Command Update-ChestRewardControls -ErrorAction SilentlyContinue) {
			Update-ChestRewardControls
		}
	}.GetNewClosure())

	$ddlChestWindow.Add_DropDown({
		if (Get-Command Update-ChestOpenerWindowList -ErrorAction SilentlyContinue) {
			Update-ChestOpenerWindowList
		}
	}.GetNewClosure())

	$btnAddChestQueue.Add_Click({
		if (Get-Command Add-ChestToQueue -ErrorAction SilentlyContinue) {
			Add-ChestToQueue
		}
	}.GetNewClosure())

	$btnRemoveChestQueue.Add_Click({
		if (Get-Command Remove-ChestFromQueue -ErrorAction SilentlyContinue) {
			Remove-ChestFromQueue
		}
	}.GetNewClosure())

	$btnResetChestQueue.Add_Click({
		if (Get-Command Reset-ChestQueue -ErrorAction SilentlyContinue) {
			Reset-ChestQueue
		}
	}.GetNewClosure())

	$btnChestItemHotkey.Add_Click({
		$currentKey = $btnChestItemHotkey.Text
		if ($currentKey -eq 'Hotkey' -or $currentKey -eq 'none') { $currentKey = $null }

		$ownerForm = if ($global:DashboardConfig.UI.MainForm) { $global:DashboardConfig.UI.MainForm } else { $this.FindForm() }
		$newKey = Show-KeyCaptureDialog -currentKey $currentKey -OwnerForm $ownerForm

		if ($null -ne $newKey -and $newKey -ne $currentKey) {
			$btnChestItemHotkey.Text = $newKey
			if (-not $global:DashboardConfig.Config.Contains('ChestOpener')) {
				$global:DashboardConfig.Config['ChestOpener'] = [ordered]@{}
			}
			$global:DashboardConfig.Config['ChestOpener']['ItemHotkey'] = $newKey
			if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$false) { WriteConfig }
		}
		elseif (-not $newKey) {
			$btnChestItemHotkey.Text = 'Hotkey'
			if (-not $global:DashboardConfig.Config.Contains('ChestOpener')) {
				$global:DashboardConfig.Config['ChestOpener'] = [ordered]@{}
			}
			$global:DashboardConfig.Config['ChestOpener']['ItemHotkey'] = ''
			if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$false) { WriteConfig }
		}
	}.GetNewClosure())

	$btnChestHotkey.Add_Click({
		$currentKey = $btnChestHotkey.Text
		if ($currentKey -eq 'Hotkey' -or $currentKey -eq 'none') { $currentKey = $null }

		$ownerForm = if ($global:DashboardConfig.UI.MainForm) { $global:DashboardConfig.UI.MainForm } else { $this.FindForm() }
		$newKey = Show-KeyCaptureDialog -currentKey $currentKey -OwnerForm $ownerForm

		$oldHotkeyId = $global:DashboardConfig.State['ChestHotkeyId']
		if ($null -ne $newKey -and $newKey -ne $currentKey) {
			$btnChestHotkey.Text = $newKey

			$scriptBlock = {
				$fgHwnd = [Custom.Native]::GetForegroundWindow()

				$selectedWindow = $global:DashboardConfig.UI.ChestWindowSelector.SelectedItem

				if ($selectedWindow) {
					$targetHwnd = [IntPtr]::Zero

					if ($selectedWindow -is [IntPtr]) {
						$targetHwnd = $selectedWindow
					} elseif ($selectedWindow.HWnd) {
						$targetHwnd = [IntPtr]$selectedWindow.HWnd
					} elseif ($selectedWindow.Handle) {
						$targetHwnd = [IntPtr]$selectedWindow.Handle
					} elseif ($selectedWindow.MainWindowHandle) {
						$targetHwnd = [IntPtr]$selectedWindow.MainWindowHandle
					}

					if ($targetHwnd -ne [IntPtr]::Zero -and $fgHwnd -ne $targetHwnd) {
						return
					}
				}

				if ($script:SyncState.IsRunning) {
					Stop-ChestOpener
				} else {
					Start-ChestOpener
				}
			}

			try {
				$global:DashboardConfig.State['ChestHotkeyId'] = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey 'ChestOpener' -OwnerLabel 'Chest Opener' -OldHotkeyId $oldHotkeyId
			} catch {
				Write-Warning "ChestOpener: Failed to register hotkey $newKey : $_"
			}

			if (-not $global:DashboardConfig.Config.Contains('ChestOpener')) {
				$global:DashboardConfig.Config['ChestOpener'] = [ordered]@{}
			}
			$global:DashboardConfig.Config['ChestOpener']['Hotkey'] = $newKey
			if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$false) { WriteConfig }
		}
		elseif (-not $newKey -and $oldHotkeyId) {
			try { UnregisterHotkeyInstance -Id $oldHotkeyId -OwnerKey 'ChestOpener' } catch {}
			$global:DashboardConfig.State['ChestHotkeyId'] = $null
			$btnChestHotkey.Text = 'Hotkey'
			if (-not $global:DashboardConfig.Config.Contains('ChestOpener')) {
				$global:DashboardConfig.Config['ChestOpener'] = [ordered]@{}
			}
			$global:DashboardConfig.Config['ChestOpener']['Hotkey'] = ''
			if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$false) { WriteConfig }
		}
	}.GetNewClosure())

	@( $btnBrowseLauncher, $btnBrowseJunction, $btnAddProfile, $btnRenameProfile, $btnRemoveProfile, $btnAddSetup, $btnEditSetup, $btnRenameSetup, $btnMasterPauseHotkey, $btnMasterResumeHotkey, $btnRefreshBosses, $btnAddNote, $btnEditNote, $btnShowNotification, $btnHideAllNotifications, $btnResetUpgrade, $btnChestHotkey, $btnChestItemHotkey, $btnResetChestQueue ) | ForEach-Object {
		Register-ButtonFadeEffect -Button $_ -NormalStartColor $grayNormalStart -NormalEndColor $grayNormalEnd -HoverStartColor $grayHoverStart -HoverEndColor $grayHoverEnd -LeftPressedStartColor $grayLeftStart -LeftPressedEndColor $grayLeftEnd -RightPressedStartColor $grayRightStart -RightPressedEndColor $grayRightEnd
	}

	@( $btnDeleteProfile, $btnDeleteSetup, $btnUnregisterHotkey, $btnCancel, $btnRemoveNote, $btnRemoveUpgrade, $btnRemoveChestQueue ) | ForEach-Object {
		Register-ButtonFadeEffect -Button $_ -NormalStartColor $redNormalStart -NormalEndColor $redNormalEnd -HoverStartColor $redHoverStart -HoverEndColor $redHoverEnd -LeftPressedStartColor $redLeftStart -LeftPressedEndColor $redLeftEnd -RightPressedStartColor $redRightStart -RightPressedEndColor $redRightEnd
	}

	@( $btnStartJunction, $btnSave ) | ForEach-Object {
		Register-ButtonFadeEffect -Button $_ -NormalStartColor $greenNormalStart -NormalEndColor $greenNormalEnd -HoverStartColor $greenHoverStart -HoverEndColor $greenHoverEnd -LeftPressedStartColor $greenLeftStart -LeftPressedEndColor $greenLeftEnd -RightPressedStartColor $greenRightStart -RightPressedEndColor $greenRightEnd
	}

	@( $btnCopyNeuz, $btnAddUpgrade, $btnAddChestQueue ) | ForEach-Object {
		Register-ButtonFadeEffect -Button $_ -NormalStartColor $blueNormalStart -NormalEndColor $blueNormalEnd -HoverStartColor $blueHoverStart -HoverEndColor $blueHoverEnd -LeftPressedStartColor $blueLeftStart -LeftPressedEndColor $blueLeftEnd -RightPressedStartColor $blueRightStart -RightPressedEndColor $blueRightEnd
	}

	$ExtraForm.Controls.AddRange(@(
		$lblTitle,
		$btnClose,
		$extraTabs
	))

	#endregion
	#endregion

	$global:DashboardConfig.UI = [PSCustomObject]@{
		MainForm                           = $mainForm
		SettingsForm                       = $settingsForm
		ExtraForm                      	   = $ExtraForm
		SettingsTabs                       = $settingsTabs
		TopBar                             = $topBar
		CloseForm                          = $btnCloseForm
		MinForm                            = $btnMinimizeForm
		InfoForm                           = $btnInfoForm
		ResizeGrip                         = $resizeGrip
		DataGridMain                       = $DataGridMain
		LoginButton                        = $btnLogin
		Ftool                              = $btnFtool
		Macro                              = $btnMacro
		Extra	                           = $btnExtra
		Wiki	                           = $btnWiki
		Settings                           = $btnSettings
		Exit                               = $btnStop
		Launch                             = $btnLaunch
		LaunchContextMenu                  = $LaunchContextMenu
		ToolTipMain                        = $toolTipMain
		ToolTipSettings                    = $toolTipSettings
		ToolTipExtra                       = $toolTipExtra
		ToolTipFtool                       = $null
		InputLauncher                      = $txtLauncher
		InputJunction                      = $txtJunction
		InputUser                          = $userBox 
		StartJunction                      = $btnStartJunction
		InputProcess                       = $txtProcessName
		InputMax                           = $txtMaxClients
		Browse                             = $btnBrowseLauncher
		InputLauncherTimeout               = $txtLauncherTimeout
		BrowseJunction                     = $btnBrowseJunction
		NeverRestartingCollectorLogin      = $chkNeverRestartingLogin
		AutoStartWithWindows               = $chkAutoStartWindows
		AutoStartSetup                     = $chkAutoStartSetup
		WorldBossListener                  = $chkWorldBossListener
		RefreshBosses                      = $btnRefreshBosses
		ShowBossImages                     = $chkShowImages
		ProfileGrid                        = $ProfileGrid
		AddProfile                         = $btnAddProfile
		ButtonPanel                        = $ButtonPanel
		RenameProfile                      = $btnRenameProfile
		RemoveProfile                      = $btnRemoveProfile
		DeleteProfile                      = $btnDeleteProfile
		SetupGrid                          = $SetupGrid
		AddSetup                           = $btnAddSetup
		EditSetup                          = $btnEditSetup
		RenameSetup                        = $btnRenameSetup
		DeleteSetup                        = $btnDeleteSetup
		CopyNeuz                           = $btnCopyNeuz
		LoginConfigGrid                    = $LoginConfigGrid
		LoginPickers                       = $Pickers
		InputPostLoginDelay                = $txtPostLoginDelay
		LoginProfileSelector               = $ddlLoginProfile
		ResolutionSelector                 = $ddlResolution
		DirectXSelector                    = $ddlDirectX
		HotkeysGrid                        = $HotkeysGrid
		MasterPauseHotkeyBtn               = $btnMasterPauseHotkey
		MasterResumeHotkeyBtn              = $btnMasterResumeHotkey
		MasterHotkeysStatus                = $lblMasterHotkeysStatus
		UnregisterHotkey                   = $btnUnregisterHotkey
		Save                               = $btnSave
		Cancel                             = $btnCancel
		NoteGrid                           = $NoteGrid
		NotificationGrid                   = $NotificationGrid
		ShowNotification                   = $btnShowNotification
		HideAllNotifications               = $btnHideAllNotifications
		AddNote                            = $btnAddNote
		EditNote                           = $btnEditNote
		RemoveNote                         = $btnRemoveNote
		ContextMenuFront                   = $itmFront
		ContextMenuBack                    = $itmBack
		ContextMenuResizeAndCenter         = $itmResizeCenter
		ContextMenuSavePos                 = $itmSavePos
		ContextMenuLoadPos                 = $itmLoadPos
		SetHotkeyShow                      = $itmSetHotkeyShow
		SetHotkeyHide                      = $itmSetHotkeyHide
		SetHotkeyLoadPos                   = $itmSetHotkeyLoadPos
		Relog                              = $itmRelog
	}
	if ($null -ne $uiPropertiesToAdd)
	{
		$uiPropertiesToAdd.GetEnumerator() | ForEach-Object {
			$global:DashboardConfig.UI | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value -Force
		}
	}

	RegisterUIEventHandlers
	$script:CurrentBuilderToolTip = $null
	return $true
}

function RegisterUIEventHandlers
{
	[CmdletBinding()]
	param()

	if ($null -eq $global:DashboardConfig.UI) { return }

	$eventMappings = @{
		MainForm                   = @{
			Load        = {
				if ($global:DashboardConfig.Paths.Ini)
				{
					$iniExists = Test-Path -Path $global:DashboardConfig.Paths.Ini
					if ($iniExists)
					{
						$iniSettings = GetIniFileContent -Ini $global:DashboardConfig.Paths.Ini
						if ($iniSettings.Count -gt 0)
						{
							$global:DashboardConfig.Config = $iniSettings
						}
					}

					RefreshLoginProfileSelector
					SyncConfigToUI
					RegisterConfiguredHotkeys
					RefreshHotkeysList
					
				}
				else { $global:DashboardConfig.UI.MainForm.StartPosition = 'CenterScreen' }

			}
			Shown       = {
				SyncConfigToUI
				RegisterConfiguredHotkeys
				RefreshHotkeysList
			}
			FormClosing = {
				param($src, $e)
				try
				{
					if ($src.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal)
					{
						if (-not $global:DashboardConfig.Config.Contains('WindowPosition')) {
							$global:DashboardConfig.Config['WindowPosition'] = [ordered]@{}
						}
						$global:DashboardConfig.Config['WindowPosition']['MainFormX'] = $src.Location.X
						$global:DashboardConfig.Config['WindowPosition']['MainFormY'] = $src.Location.Y
						$global:DashboardConfig.Config['WindowPosition']['MainFormHeight'] = $src.Height
						if ($global:DashboardConfig.UI.DataGridMain) { $global:DashboardConfig.Config['WindowPosition']['DataGridHeight'] = $global:DashboardConfig.UI.DataGridMain.Height }
						
						if ($global:DashboardConfig.UI.ExtraForm -and -not $global:DashboardConfig.UI.ExtraForm.IsDisposed) {
							$ef = $global:DashboardConfig.UI.ExtraForm
							if ($ef.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
								$global:DashboardConfig.Config['WindowPosition']['ExtraFormX'] = $ef.Location.X; $global:DashboardConfig.Config['WindowPosition']['ExtraFormY'] = $ef.Location.Y
							}
						}
						
						if ($global:DashboardConfig.UI.SettingsForm -and -not $global:DashboardConfig.UI.SettingsForm.IsDisposed) {
							$sf = $global:DashboardConfig.UI.SettingsForm
							if ($sf.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
								$global:DashboardConfig.Config['WindowPosition']['SettingsFormX'] = $sf.Location.X; $global:DashboardConfig.Config['WindowPosition']['SettingsFormY'] = $sf.Location.Y
							}
						}
						Write-Verbose "Saving Window Position: X=$($src.Location.X), Y=$($src.Location.Y), H=$($src.Height)"
						if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
					}
				}
				catch
				{
				}
				if (Get-Command RemoveAllHotkeys -ErrorAction SilentlyContinue -Verbose:$False) { RemoveAllHotkeys }
				if (Get-Command StopDashboard -ErrorAction SilentlyContinue -Verbose:$False) { StopDashboard }
			}
		}
		SettingsForm               = @{
			Load = { SyncConfigToUI; RegisterConfiguredHotkeys; RefreshHotkeysList }
		}
		ExtraForm                  = @{
			MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.ExtraForm.Handle, 0xA1, 0x2, 0) }
		}
		#region MainForm
		TopBar                     = @{ MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.MainForm.Handle, 0xA1, 0x2, 0) } }
		InfoForm                   = @{ MouseDown = { 
				$cms = New-Object System.Windows.Forms.ContextMenuStrip
				if ('Custom.DarkRenderer' -as [Type]) { $cms.Renderer = New-Object Custom.DarkRenderer }
				$itmWeb = $cms.Items.Add("Open Website")
				$itmWeb.Image = if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon)) { 
					try {
						$icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon)
						$bmp = $icon.ToBitmap()
						$icon.Dispose()
						$bmp
					} catch { $null }
				} else { $null }
				$itmWeb.add_Click({ Start-Process "https://immortal-divine.github.io/Entropia_Dashboard/" })
				
				$itmGuide = $cms.Items.Add("Start Interactive Guide")
				try {
					$guideBmp = New-Object System.Drawing.Bitmap(16, 16)
					$g = [System.Drawing.Graphics]::FromImage($guideBmp)
					$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
					$g.FillEllipse([System.Drawing.Brushes]::MediumSeaGreen, 0, 0, 15, 15)
					$sf = New-Object System.Drawing.StringFormat
					$sf.Alignment = [System.Drawing.StringAlignment]::Center
					$sf.LineAlignment = [System.Drawing.StringAlignment]::Center
					$g.DrawString("?", (New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)), [System.Drawing.Brushes]::White, 8, 8, $sf)
					$g.Dispose()
					$itmGuide.Image = $guideBmp
				} catch {}
				$itmGuide.add_Click({ if (Get-Command Show-Guide -ErrorAction SilentlyContinue) { Show-Guide } })
				
				$itmKofi = $cms.Items.Add("Support me on Ko-fi")
				$kofiPath = Join-Path $global:DashboardConfig.Paths.App 'kofi.png'
				if (-not (Test-Path $kofiPath)) {
					try {
						[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
						$wc = New-Object System.Net.WebClient
						$wc.DownloadFile('https://storage.ko-fi.com/cdn/cup-border.png', $kofiPath)
						$wc.Dispose()
					} catch {}
				}
				if (Test-Path $kofiPath) {
					try { $itmKofi.Image = [System.Drawing.Image]::FromFile($kofiPath) } catch {}
				}
				$itmKofi.add_Click({ Start-Process "https://ko-fi.com/U7U61S0EGT" })
				
				$itmPayPal = $cms.Items.Add("Support me with PayPal")
				$paypalPath = Join-Path $global:DashboardConfig.Paths.App 'paypal.png'
				if (-not (Test-Path $paypalPath)) {
					try {
						[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
						$wc = New-Object System.Net.WebClient
						$wc.DownloadFile('https://www.paypalobjects.com/webstatic/icon/pp258.png', $paypalPath)
						$wc.Dispose()
					} catch {}
				}
				if (Test-Path $paypalPath) {
					try { $itmPayPal.Image = [System.Drawing.Image]::FromFile($paypalPath) } catch {}
				}
				$itmPayPal.add_Click({ Start-Process "https://www.paypal.me/immortaldivine" })
				
				$cms.Show($Sender, 0, $Sender.Height)
			} 
		}
		MinForm                    = @{ Click = { $global:DashboardConfig.UI.MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized } }
		CloseForm                  = @{ Click = { $global:DashboardConfig.UI.MainForm.Close() } }

		Launch                     = @{
			Click = {
				if ($global:DashboardConfig.State.LaunchActive)
				{
					try { StopClientLaunch } catch { Show-DarkMessageBox $_.Exception.Message }
				}
				else
				{
					$this.ContextMenuStrip.Show($this, 0, $this.Height)
				}
			}
		}
		LaunchContextMenu          = @{
			Opening = {
				param($s, $e)
				$s.Items.Clear()
						
				$selectedProfileName = 'Default'
				if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['SelectedProfile'] -and -not [string]::IsNullOrWhiteSpace($global:DashboardConfig.Config['Options']['SelectedProfile']))
				{
					$selectedProfileName = $global:DashboardConfig.Config['Options']['SelectedProfile']
				}

				$maxClients = 1
				if ($global:DashboardConfig.Config['MaxClients'] -and $global:DashboardConfig.Config['MaxClients']['MaxClients'])
				{
					if (-not ([int]::TryParse($global:DashboardConfig.Config['MaxClients']['MaxClients'], [ref]$maxClients)))
					{
						$maxClients = 1
					}
				}

				$processName = 'neuz'
				if ($global:DashboardConfig.Config['ProcessName'] -and $global:DashboardConfig.Config['ProcessName']['ProcessName'])
				{
					$processName = $global:DashboardConfig.Config['ProcessName']['ProcessName']
				}
						
				$profileClientCount = 0
				$profileClientCount = @(Get-Process -Name $processName -ErrorAction SilentlyContinue).Count
				$howManyUntilMax = [Math]::Max(0, $maxClients - $profileClientCount)

				$header = $s.Items.Add("Launch $($howManyUntilMax)x '$($selectedProfileName)'")
				$header.Enabled = $true
				$header.add_Click({
						StartClientLaunch
					})


				$s.Items.Add('-')
				$defaultItem = $s.Items.Add('Main Launcher')

				1..10 | ForEach-Object {
					$count = $_
					$subItem = $defaultItem.DropDownItems.Add("Start $($count)x")
					$subItem.Tag = $count

					$subItem.add_Click({
							param($s, $ev)
							$launchCount = $s.Tag
							$launchCount = StartClientLaunch -ClientAddCount $launchCount -ProfileNameOverride 'Default'
						})
				}

				if ($global:DashboardConfig.Config['Profiles'])
				{
					$profiles = $global:DashboardConfig.Config['Profiles']
					if ($profiles.Count -gt 0)
					{
						$s.Items.Add('-')
						foreach ($key in $profiles.Keys)
						{
							$profileItem = $s.Items.Add($key)

							1..10 | ForEach-Object {
								$count = $_
								$subItem = $profileItem.DropDownItems.Add("Start $($count)x")

								$subItem.Tag = @{
									ProfileName = $key
									Count       = $count
								}

								$subItem.add_Click({
										param($s, $ev)
										$data = $s.Tag
										StartClientLaunch -ProfileNameOverride $data.ProfileName -ClientAddCount $data.Count
									})
							}
						}
					}
				}

				$s.Items.Add('-')

				$configKeys = $global:DashboardConfig.Config.Keys
				foreach ($key in $configKeys)
				{
					if ($key -match '^Setup_(.+)$')
					{
						$setupName = $Matches[1]
						$setupItem = $s.Items.Add("Start Setup: $setupName")
						$setupItem.Tag = $setupName
						$setupItem.add_Click({
							param($s, $ev)
							$sn = $s.Tag
							StartClientLaunch -SavedLaunchLoginConfig -SetupName $sn
						})
					}
				}
			}
		}
		LoginButton                = @{
			Click = {
				try
				{
					$loginCommand = Get-Command LoginSelectedRow -ErrorAction Stop -Verbose:$False
					& $loginCommand
				}
				catch
				{
					$errorMessage = "Login action failed.`n`nCould not find or execute the 'LoginSelectedRow' function. The 'login.psm1' module may have failed to load correctly.`n`nTechnical Details: $($_.Exception.Message)"
					try { ShowErrorDialog -Message $errorMessage }
					catch { Show-DarkMessageBox $errorMessage 'Login Error' 'Ok' 'Error' }
				}
			}
		}
		Settings                   = @{ Click = { ShowSettingsForm; RegisterConfiguredHotkeys; RefreshHotkeysList } }
		Exit                       = @{
			MouseDown = {
				param($s, $e)
				if ($global:DashboardConfig.State.LoginActive) { return }
				if ($script:LastExitClick -and [DateTime]::Now -lt $script:LastExitClick.AddMilliseconds(500)) { return }
				$script:LastExitClick = [DateTime]::Now

				if ($e.Button -eq 'Left')
				{
					if ((Show-DarkMessageBox 'Are you sure you want to close the selected Clients?' 'Confirm' 'YesNo' 'Error') -eq 'Yes') { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { Stop-Process -Id $_.Tag.Id -Force -EA 0 } }
				}
				elseif ($e.Button -eq 'Right')
				{
					if ((Show-DarkMessageBox 'Are you sure you want to disconnect the selected Clients?' 'Confirm' 'YesNo' 'Warning') -eq 'Yes') { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { try { [Custom.Native]::CloseTcpConnectionsForPid($_.Tag.Id) } catch {} } }
				}
				$script:LastExitClick = [DateTime]::Now
			}
		}
		Ftool                      = @{ 
			MouseDown = {
				param($s, $e)
				if ($e.Button -eq 'Left') {
                    Write-Verbose "Ftool: Left Click"
					if (Get-Command FtoolSelectedRow -ErrorAction SilentlyContinue) {
                        $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { FtoolSelectedRow $_ }
                    } else {
                        Write-Error "Ftool: FtoolSelectedRow command not found!"
                    }
				} elseif ($e.Button -eq 'Right') {
                    Write-Verbose "Ftool: Right Click"
                    try {
                        $selectedRows = $global:DashboardConfig.UI.DataGridMain.SelectedRows
                        if ($selectedRows.Count -eq 0) { return }
                        $row = $selectedRows[0]
                        $rowIdentity = Get-RowIdentity -Row $row
                        $safeIdentity = $rowIdentity -replace '[^a-zA-Z0-9_\-]', '_'
                        $sessionKey = "Session_Ftool_$safeIdentity"

                        if ($null -eq $global:DashboardConfig.Config['Ftool'] -or $global:DashboardConfig.Config['Ftool'] -is [string]) { $global:DashboardConfig.Config['Ftool'] = [ordered]@{} }
                        if ($null -eq $global:DashboardConfig.Config['Macro'] -or $global:DashboardConfig.Config['Macro'] -is [string]) { $global:DashboardConfig.Config['Macro'] = [ordered]@{} }

                        $activeForms = [System.Collections.ArrayList]@()
                        if ($global:DashboardConfig.Resources.FtoolForms) {
                            foreach ($f in $global:DashboardConfig.Resources.FtoolForms.Values) {
                                $isFtool = $false
                                if ($f -and -not $f.IsDisposed -and $f.Tag -and $f.Tag -isnot [string]) {
                                    if ($null -eq $f.Tag.Type -or $f.Tag.Type -ne 'Macro') {
                                        if ($f.Tag.SelectedWindow -eq $row.Tag.MainWindowHandle) {
                                            $isFtool = $true
                                        }
                                    }
                                }
                                
                                if ($isFtool) {
                                    $activeForms.Add($f) | Out-Null
                                    Write-Verbose "Ftool Debug: Captured form $($f.Name)"
                                }
                            }
                        }

                        if ($activeForms.Count -gt 0) {
                            $global:DashboardConfig.Config['Ftool'][$sessionKey] = $activeForms.Count
                            for ($i = 0; $i -lt $activeForms.Count; $i++) {
                                $f = $activeForms[$i]
                                $data = $f.Tag
                                $configKey = "${sessionKey}_$i"
                                if (-not $global:DashboardConfig.Config['Ftool'].Contains($configKey) -or $global:DashboardConfig.Config['Ftool'][$configKey] -isnot [System.Collections.IDictionary]) { $global:DashboardConfig.Config['Ftool'][$configKey] = [ordered]@{} }
                                $config = $global:DashboardConfig.Config['Ftool'][$configKey]
                                $config['key'] = if ($data.BtnKeySelect) { $data.BtnKeySelect.Text } else { "" }
                                $config['interval'] = if ($data.Interval) { $data.Interval.Text } else { "" }
                                $config['name'] = if ($data.Name) { $data.Name.Text } else { "" }
                                
                                $instanceId = $data.InstanceId
                                $instanceKey = "instance_$instanceId"
                                if ($global:DashboardConfig.Resources.ExtensionTracking -and $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey)) {
                                    $activeExtensions = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions
                                    $config['ExtCount'] = $activeExtensions.Count
                                    for ($j = 0; $j -lt $activeExtensions.Count; $j++) {
                                        $extKey = $activeExtensions[$j]
                                        if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey)) {
                                            $extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
                                            $extConfig = [ordered]@{}
                                            $extConfig['key'] = if ($extData.BtnKeySelect) { $extData.BtnKeySelect.Text } else { "" }
                                            $extConfig['interval'] = if ($extData.Interval) { $extData.Interval.Text } else { "" }
                                            $extConfig['name'] = if ($extData.Name) { $extData.Name.Text } else { "" }
                                            $extConfig['hotkey'] = if ($extData.Hotkey) { $extData.Hotkey } else { "" }
                                            $config["Ext_$j"] = $extConfig
                                        }
                                    }
                                }
                            }
                            WriteConfig
                            
                            foreach ($f in $activeForms) {
                                $f.Close()
                            }
                            Show-DarkMessageBox "Saved $($activeForms.Count) active Ftool instances for PID $pid." "Session Saved" "OK" "Information" "success"
                        } else {
                            if ($global:DashboardConfig.Config['Ftool'].Contains($sessionKey)) {
                                $savedValue = $global:DashboardConfig.Config['Ftool'][$sessionKey]
                                $count = [int]$savedValue
                                for ($i = 0; $i -lt $count; $i++) {
                                    $configKey = "${sessionKey}_$i"
                                    $config = $null
                                    if ($global:DashboardConfig.Config['Ftool'].Contains($configKey)) {
                                        $configRaw = $global:DashboardConfig.Config['Ftool'][$configKey]
                                        if ($configRaw -is [string] -and $configRaw.StartsWith('{')) {
                                            $config = $configRaw | ConvertFrom-Json
                                        } else { $config = $configRaw }
                                    }
                                    FtoolSelectedRow $row -ForceNew -Config $config -InstanceIndex ($i + 1)
                                }
                                Show-DarkMessageBox "Restored $count Ftool instances for $($row.Cells[1].Value)." "Session Restored" "OK" "Information" "success"
                            } else {
                                Show-DarkMessageBox "No active Ftool instances to close and no saved session to restore." "No Action" "OK" "Information"
                            }
                        }
                    } catch {
                        Show-DarkMessageBox "Error handling Ftool session: $($_.Exception.Message)" "Session Error" "OK" "Error"
                    }
				}
			} 
		}
		Macro                      = @{ 
			MouseDown = {
				param($s, $e)
				if ($e.Button -eq 'Left') {
					if (Get-Command MacroSelectedRow -EA 0 -Verbose:$False) { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { MacroSelectedRow $_ } }
				} elseif ($e.Button -eq 'Right') {
		    		try {
						$selectedRows = $global:DashboardConfig.UI.DataGridMain.SelectedRows
						if ($selectedRows.Count -eq 0) { return }
						$row = $selectedRows[0]

						$rowIdentity = Get-RowIdentity -Row $row
						$safeIdentity = $rowIdentity -replace '[^a-zA-Z0-9_\-]', '_'
						$sessionKey = "Session_Macro_Macro_$safeIdentity"
						Write-Verbose "Macro Right-Click: Identity=$rowIdentity, SessionKey=$sessionKey"

						if ([string]::IsNullOrWhiteSpace($sessionKey) -or $rowIdentity -eq "") {
							Show-DarkMessageBox "Failed to generate a valid session key. Row data missing." "Session Error" "OK" "Error"
							return
						}

						if ($null -eq $global:DashboardConfig.Config['Ftool'] -or $global:DashboardConfig.Config['Ftool'] -is [string]) { $global:DashboardConfig.Config['Ftool'] = [ordered]@{} }
						if ($null -eq $global:DashboardConfig.Config['Macro'] -or $global:DashboardConfig.Config['Macro'] -is [string]) { $global:DashboardConfig.Config['Macro'] = [ordered]@{} }

						$activeForms = [System.Collections.ArrayList]@()
						if ($global:DashboardConfig.Resources.FtoolForms) {
							foreach ($k in $global:DashboardConfig.Resources.FtoolForms.Keys) {
								if ($k -eq "Macro_$safeIdentity" -or $k -match "^Macro_$([regex]::Escape($safeIdentity))_\d+$") {
									$f = $global:DashboardConfig.Resources.FtoolForms[$k]
									if ($f -and -not $f.IsDisposed -and $f.Tag -and $f.Tag -isnot [string]) {
										$type = if ($f.Tag.PSObject.Properties['Type']) { $f.Tag.Type } else { $null }
										if ($type -eq 'Macro') {
											$activeForms.Add($f) | Out-Null
										}
									}
								}
							}
						}

						if ($activeForms.Count -gt 0) {
							$global:DashboardConfig.Config['Macro'][$sessionKey] = $activeForms.Count
							WriteConfig

							foreach ($f in $activeForms) {
								$f.Close()
							}
							Show-DarkMessageBox "Saved $($activeForms.Count) active Macro instances for $rowIdentity." "Session Saved" "OK" "Information" "success"
						} else {
							if ($global:DashboardConfig.Config['Macro'].Contains($sessionKey)) {
								$count = [int]$global:DashboardConfig.Config['Macro'][$sessionKey]
								for ($i = 1; $i -le $count; $i++) {
									MacroSelectedRow $row -ForceNew -InstanceIndex $i
								}
								Show-DarkMessageBox "Restored $count Macro instances for $rowIdentity." "Session Restored" "OK" "Information" "success"
							} else {
								Show-DarkMessageBox "No active Macro instances to close and no saved session to restore." "No Action" "OK" "Information"
							}
						}
					} catch {
						Show-DarkMessageBox "Error handling Macro session: $($_.Exception.Message)" "Session Error" "OK" "Error"
					}
				}
			} 
		}
		Extra                      = @{ MouseDown = {
			param($s,$e)
				if ($e.Button -eq 'Left') {
					ShowExtraForm
					if (Get-Command RefreshNoteGrid -ErrorAction SilentlyContinue -Verbose:$False) { RefreshNoteGrid }
					if (Get-Command RefreshNotificationGrid -ErrorAction SilentlyContinue -Verbose:$False) { RefreshNotificationGrid }
					if (Get-Command Initialize-ChestOpener -ErrorAction SilentlyContinue) { Initialize-ChestOpener }
					if (Get-Command Update-ChestOpenerWindowList -ErrorAction SilentlyContinue) { Update-ChestOpenerWindowList }
				}
				if ($e.Button -eq 'Right') {
					if (Get-Command ChatCommanderSelectedRow -EA 0 -Verbose:$False) { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { ChatCommanderSelectedRow $_ } }
				}
			} 
		}
		Wiki                      = @{ Click = { 
			try
				{
					$wikiCommand = Get-Command Show-Wiki -ErrorAction Stop -Verbose:$False
					& $wikiCommand
				}
				catch
				{
					$errorMessage = "Wiki action failed.`n`nCould not find or execute the 'Show-Wiki' function. The 'wiki.psm1' module may have failed to load correctly.`n`nTechnical Details: $($_.Exception.Message)"
					try { ShowErrorDialog -Message $errorMessage }
					catch { Show-DarkMessageBox $errorMessage 'Wiki Error' 'Ok' 'Error' }
				}
			}
		}
		DataGridMain              = @{
			DoubleClick = { param($s,$e); $h = $s.HitTest($e.X,$e.Y); if ($h.RowIndex -ge 0 -and $s.Rows[$h.RowIndex].Tag) { SetWindowToolStyle -hWnd ($s.Rows[$h.RowIndex].Tag.MainWindowHandle) -Hide $false; [Custom.Native]::BringToFront($s.Rows[$h.RowIndex].Tag.MainWindowHandle); SetWindowToolStyle -hWnd ($s.Rows[$h.RowIndex].Tag.MainWindowHandle) -Hide $false } }
			MouseDown   = {
				param($s,$e)
				$h = $s.HitTest($e.X,$e.Y)
				if ($e.Button -eq 'Right' -and $h.RowIndex -ge 0)
				{
					$clickedRow = $s.Rows[$h.RowIndex]

					if (-not $clickedRow.Selected)
					{
						$s.ClearSelection()
						$clickedRow.Selected = $true
					}
				}
			}
		}
		ContextMenuFront           = @{ Click = { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { $h = $_.Tag.MainWindowHandle; SetWindowToolStyle -hWnd $h -Hide $false; [Custom.Native]::BringToFront($h); SetWindowToolStyle -hWnd $h -Hide $false } } }                
		ContextMenuBack            = @{ Click = { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { [Custom.Native]::SendToBack($_.Tag.MainWindowHandle); [Custom.Native]::EmptyWorkingSet($_.Tag.Id) } } }
		ContextMenuResizeAndCenter = @{ Click = { $scr = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea; $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { [Custom.Native]::PositionWindow($_.Tag.MainWindowHandle, [Custom.Native]::TopWindowHandle, [int](($scr.Width - 1040) / 2), [int](($scr.Height - 807) / 2), 1040, 807, 0x0010) } } }
		ContextMenuSavePos         = @{ Click = { if (Get-Command Save-WindowPositions -ErrorAction SilentlyContinue -Verbose:$False) { Save-WindowPositions } } }
		ContextMenuLoadPos         = @{ Click = { if (Get-Command Restore-WindowPositions -ErrorAction SilentlyContinue -Verbose:$False) { Restore-WindowPositions } } }
		Relog                      = @{
			Click = {
				$selected = $global:DashboardConfig.UI.DataGridMain.SelectedRows
				if ($selected)
				{
					if (-not $global:DashboardConfig.State.NotificationActionQueue) { $global:DashboardConfig.State.NotificationActionQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new()) }
					foreach ($row in $selected)
					{
						if ($row.Tag -and $row.Tag.Id)
						{
							$global:DashboardConfig.State.NotificationActionQueue.Enqueue(@{ Action = 'Reconnect'; Pid = $row.Tag.Id })
						}
					}
				}
			}
		}
		HotkeysGrid                = @{
			CellDoubleClick = {
				param($s, $e)
				if ($e.RowIndex -lt 0) { return }
				
				$row = $s.Rows[$e.RowIndex]
				if (-not $row.Tag -or ($row.Tag -is [System.Collections.IDictionary] -and $row.Tag.Contains('Header') -and $row.Tag.Header)) { return }
				if (-not $row.Tag.Id -or -not $row.Tag.OwnerKey) { return }
				
				$null = $id; $id = $row.Tag.Id
				$ownerKey = $row.Tag.OwnerKey
				$currentKeyString = $row.Cells[0].Value.ToString()
				if ($currentKeyString -match '\(Paused\)$') { $currentKeyString = $currentKeyString -replace ' \(Paused\)$', '' }

				$newKey = Show-KeyCaptureDialog -currentKey $currentKeyString -Owner $global:DashboardConfig.UI.MainForm
				
				if ($null -eq $newKey -or ($newKey -eq $currentKeyString)) { return }
				
				if ($ownerKey -match '^GroupHotkey_(.+)')
				{
					$groupName = $Matches[1]
					
					if (-not $global:DashboardConfig.Config.Contains('Hotkeys')) { $global:DashboardConfig.Config['Hotkeys'] = [ordered]@{} }
					$global:DashboardConfig.Config['Hotkeys'][$groupName] = $newKey
					
					$memberString = ''
					if ($global:DashboardConfig.Config.Contains('HotkeyGroups') -and $global:DashboardConfig.Config['HotkeyGroups'].Contains($groupName))
					{
						$memberString = $global:DashboardConfig.Config['HotkeyGroups'][$groupName]
					}
					
					if (-not [string]::IsNullOrWhiteSpace($memberString))
					{
						$memberList = $memberString -split ','
						
						SetHotkey -KeyCombinationString $newKey -OwnerKey $ownerKey -Action ({
								$targetMembers = $memberList
								$grid = $global:DashboardConfig.UI.DataGridMain
								if (-not $grid) { return }

								foreach ($row in $grid.Rows)
								{
									$identity = Get-RowIdentity -Row $row
									if ($targetMembers -contains $identity)
									{
										if ($row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
										{
											$h = $row.Tag.MainWindowHandle
											SetWindowToolStyle -hWnd $h -Hide $false
											[Custom.Native]::BringToFront($h)
											SetWindowToolStyle -hWnd $h -Hide $false
										}
									}
								}
							}.GetNewClosure())
					}
				}
				elseif ($ownerKey -match '^quickcmd_(.+)_(.+)')
				{
					$instanceId = $Matches[1]
					$cmdId = $Matches[2]
					if ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
					{
						$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
						if ($form -and -not $form.IsDisposed)
						{
							$data = $form.Tag
							$cmd = $null
							if ($data.SavedCommands) { $cmd = $data.SavedCommands | Where-Object { $_.Id -eq $cmdId } | Select-Object -First 1 }
							if (-not $cmd -and $data.GlobalCommandList) { $cmd = $data.GlobalCommandList | Where-Object { $_.Id -eq $cmdId } | Select-Object -First 1 }
							if ($cmd)
							{
								if (-not $global:DashboardConfig.Config.Contains('ChatCommander')) { $global:DashboardConfig.Config['ChatCommander'] = [ordered]@{} }
								$cmd.Hotkey = $newKey
								$cmd.HotkeyId = $null
								try { UnregisterHotkeyInstance -Id $id -OwnerKey "quickcmd_$($instanceId)_$($cmdId)" } catch {}
								try { UnregisterHotkeyInstance -Id $id -OwnerKey $instanceId } catch {}
								if ($newKey) {
									$scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($instanceId)' -CommandId '$($cmdId)'")
									$ownerKeyNew = "quickcmd_$($instanceId)_$($cmdId)"
									$ownerLabel = $null
									if ($cmd.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($cmd.Name)) { $ownerLabel = $cmd.Name }
									elseif ($cmd.PSObject.Properties['CommandText'] -and -not [string]::IsNullOrEmpty($cmd.CommandText)) { $ownerLabel = $cmd.CommandText }
									else { $ownerLabel = "Invoke Command: $cmdId" }
									try { $cmd.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKeyNew -OwnerLabel $ownerLabel -OldHotkeyId $id } catch {}
								}
								if (Get-Command UpdateChatCommanderSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateChatCommanderSettings $data -forceWrite }
								if ($form -and -not $form.IsDisposed) { RefreshSavedCommandsUI $form }
							}
						}
					}
				}
				elseif ($ownerKey -match '^macroinst_(.+)')
				{
					$instanceId = $Matches[1]
					if ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
					{
						$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
						if ($form -and -not $form.IsDisposed)
						{
							$data = $form.Tag
							$data.Hotkey = $newKey
							$data.HotkeyId = $null
							try { UnregisterHotkeyInstance -Id $id -OwnerKey "macroinst_$($instanceId)" } catch {}
							try { UnregisterHotkeyInstance -Id $id -OwnerKey $instanceId } catch {}
							if ($newKey) {
								$script = [scriptblock]::Create("ToggleMacroInstance -InstanceId '$($instanceId)'")
								try { $data.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $script -OwnerKey "macroinst_$($instanceId)" -OldHotkeyId $id } catch {}
							}
							if (Get-Command UpdateMacroSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateMacroSettings $data -forceWrite }
						}
					}
				}
				elseif ($ownerKey -match '^global_toggle_(.+)')
				{
					$instanceId = $Matches[1]
					if ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
					{
						$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
						if ($form -and -not $form.IsDisposed)
						{
							$data = $form.Tag
							$data.GlobalHotkey = $newKey
							
							$script = @"
Write-Verbose "FTOOL: Global-toggle hotkey triggered for instance '$($data.InstanceId)'"
if (`$global:DashboardConfig.Resources.FtoolForms.Contains('$($data.InstanceId)')) {
    `$f = `$global:DashboardConfig.Resources.FtoolForms['$($data.InstanceId)']
    if (`$f -and -not `$f.IsDisposed -and `$f.Tag) {
        `$toggle = `$f.Tag.BtnHotkeyToggle
        if (`$toggle) {
            try {
                if (`$toggle.InvokeRequired) {
                    `$action = [System.Action]{ `$toggle.Checked = -not `$toggle.Checked }
                    `$toggle.Invoke(`$action)
                } else {
                    `$toggle.Checked = -not `$toggle.Checked
                }
                ToggleInstanceHotkeys -InstanceId '$($data.InstanceId)' -ToggleState `$toggle.Checked
                Write-Verbose "FTOOL: Global toggle action completed for instance '$($data.InstanceId)'."
            } catch {
                Write-Warning "FTOOL: Error during global toggle action for instance '$($data.InstanceId)'. Error: `$_"
            }
        }
    }
}
"@
							$scriptBlock = [scriptblock]::Create($script)
							$oldId = $data.GlobalHotkeyId
							try {
								$data.GlobalHotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKey -OldHotkeyId $oldId
								if (Get-Command UpdateSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateSettings $data -forceWrite }
							} catch {
								Show-DarkMessageBox "Failed to update hotkey: $_" "Error" "OK" "Error"
							}
						}
					}
				}
				elseif ($ownerKey -match '^ext_(.+)_\d+')
				{
					if ($global:DashboardConfig.Resources.ExtensionData -and $global:DashboardConfig.Resources.ExtensionData.Contains($ownerKey))
					{
						$extData = $global:DashboardConfig.Resources.ExtensionData[$ownerKey]
						$extData.Hotkey = $newKey
						$extData.BtnHotKey.Text = $newKey
						
						$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($extData.InstanceId)' -ExtKey '$($ownerKey)'")
						$oldId = $extData.HotkeyId
						try {
							$extData.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKey -OldHotkeyId $oldId
							$form = $extData.Panel.FindForm()
							if ($form -and $form.Tag) { if (Get-Command UpdateSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateSettings $form.Tag $extData -forceWrite } }
						} catch {
							Show-DarkMessageBox "Failed to update hotkey: $_" "Error" "OK" "Error"
						}
					}
				}
				elseif ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Contains($ownerKey))
				{
					$instanceId = $ownerKey
					$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
					if ($form -and -not $form.IsDisposed)
					{
						$data = $form.Tag
						$data.Hotkey = $newKey
						$data.BtnHotKey.Text = $newKey
						
						$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($data.InstanceId)' -ExtKey `$null")
						$oldId = $data.HotkeyId
						try {
							$data.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKey -OldHotkeyId $oldId
							if (Get-Command UpdateSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateSettings $data -forceWrite }
						} catch {
							Show-DarkMessageBox "Failed to update hotkey: $_" "Error" "OK" "Error"
						}
					}
				}
				
				if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
				SyncConfigToUI
				RegisterConfiguredHotkeys
				RefreshHotkeysList
			}
		}
		SetHotkeyShow              = @{
			Click = {
				$grid = $global:DashboardConfig.UI.DataGridMain
				$selectedRows = $grid.SelectedRows
				if ($selectedRows.Count -eq 0) { return }

				$identities = @()
				foreach ($row in $selectedRows)
				{
					$id = Get-RowIdentity -Row $row
					if (-not [string]::IsNullOrEmpty($id)) { $identities += $id }
				}

				$defaultName = 'NewGroup'
				if ($identities.Count -eq 1)
				{ $defaultName = ($identities[0] -split ':')[0] }
				
				$groupName = ShowInputBox -Title 'Hotkey Group (Show)' -Prompt 'Enter a name for this selection group:' -DefaultText $defaultName
				if ([string]::IsNullOrWhiteSpace($groupName)) { return }

				$currentKey = 'none'
				if ($global:DashboardConfig.Config.Contains('Hotkeys'))
				{
					if ($global:DashboardConfig.Config['Hotkeys'].Contains($groupName))
					{
						$currentKey = $global:DashboardConfig.Config['Hotkeys'][$groupName]
					}
				}

				$newKey = Show-KeyCaptureDialog -currentKey $currentKey -Owner $global:DashboardConfig.UI.MainForm
				
				if ($null -eq $newKey -or ($newKey -eq $currentKey -and $currentKey -ne 'none')) { return }

				$global:DashboardConfig.Config['Hotkeys'][$groupName] = $newKey
				$global:DashboardConfig.Config['HotkeyGroups'][$groupName] = ($identities -join ',')

				if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }

				SyncConfigToUI
				RegisterConfiguredHotkeys
				RefreshHotkeysList
				
				Show-DarkMessageBox "Hotkey '$newKey' assigned as '$groupName' to:`n$($identities -join "`n")" 'Hotkeys Created' 'Ok' 'Information' 'success'
			}
		}
		SetHotkeyHide              = @{
			Click = {
				$grid = $global:DashboardConfig.UI.DataGridMain
				$selectedRows = $grid.SelectedRows
				if ($selectedRows.Count -eq 0) { return }

				$identities = @()
				foreach ($row in $selectedRows)
				{
					$id = Get-RowIdentity -Row $row
					if (-not [string]::IsNullOrEmpty($id)) { $identities += $id }
				}

				$defaultName = 'NewGroup'
				if ($identities.Count -eq 1)
				{ $defaultName = ($identities[0] -split ':')[0] }
				
				$groupName = ShowInputBox -Title 'Hotkey Group (Hide)' -Prompt 'Enter a name for this selection group:' -DefaultText $defaultName
				if ([string]::IsNullOrWhiteSpace($groupName)) { return }

				$currentKey = 'none'
				if ($global:DashboardConfig.Config.Contains('HotkeysHide'))
				{
					if ($global:DashboardConfig.Config['HotkeysHide'].Contains($groupName))
					{
						$currentKey = $global:DashboardConfig.Config['HotkeysHide'][$groupName]
					}
				}

				$newKey = Show-KeyCaptureDialog -currentKey $currentKey -Owner $global:DashboardConfig.UI.MainForm
				
				if ($null -eq $newKey -or ($newKey -eq $currentKey -and $currentKey -ne 'none')) { return }

				if (-not $global:DashboardConfig.Config.Contains('HotkeysHide')) {
					$global:DashboardConfig.Config['HotkeysHide'] = @{}
				}
				if (-not $global:DashboardConfig.Config.Contains('HotkeyGroupsHide')) {
					$global:DashboardConfig.Config['HotkeyGroupsHide'] = @{}
				}

				$global:DashboardConfig.Config['HotkeysHide'][$groupName] = $newKey
				$global:DashboardConfig.Config['HotkeyGroupsHide'][$groupName] = ($identities -join ',')

				if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }

				SyncConfigToUI
				RegisterConfiguredHotkeys
				RefreshHotkeysList
				
				Show-DarkMessageBox "Hotkey '$newKey' assigned as '$groupName' to:`n$($identities -join "`n")" 'Hotkeys (Hide) Created' 'Ok' 'Information' 'success'
			}
		}
		SetHotkeyLoadPos           = @{
			Click = {
				$grid = $global:DashboardConfig.UI.DataGridMain
				$selectedRows = $grid.SelectedRows
				if ($selectedRows.Count -eq 0) { return }

				$identities = @()
				foreach ($row in $selectedRows)
				{
					$id = Get-RowIdentity -Row $row
					if (-not [string]::IsNullOrEmpty($id)) { $identities += $id }
				}

				$defaultName = 'NewGroup'
				if ($identities.Count -eq 1)
				{ $defaultName = ($identities[0] -split ':')[0] }
				
				$groupName = ShowInputBox -Title 'Hotkey Group (Load Pos)' -Prompt 'Enter a name for this selection group:' -DefaultText $defaultName
				if ([string]::IsNullOrWhiteSpace($groupName)) { return }

				$currentKey = 'none'
				if ($global:DashboardConfig.Config.Contains('HotkeysLoadPos'))
				{
					if ($global:DashboardConfig.Config['HotkeysLoadPos'].Contains($groupName))
					{
						$currentKey = $global:DashboardConfig.Config['HotkeysLoadPos'][$groupName]
					}
				}

				$newKey = Show-KeyCaptureDialog -currentKey $currentKey -Owner $global:DashboardConfig.UI.MainForm
				
				if ($null -eq $newKey -or ($newKey -eq $currentKey -and $currentKey -ne 'none')) { return }

				if (-not $global:DashboardConfig.Config.Contains('HotkeysLoadPos')) {
					$global:DashboardConfig.Config['HotkeysLoadPos'] = @{}
				}
				if (-not $global:DashboardConfig.Config.Contains('HotkeyGroupsLoadPos')) {
					$global:DashboardConfig.Config['HotkeyGroupsLoadPos'] = @{}
				}

				$global:DashboardConfig.Config['HotkeysLoadPos'][$groupName] = $newKey
				$global:DashboardConfig.Config['HotkeyGroupsLoadPos'][$groupName] = ($identities -join ',')

				if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }

				SyncConfigToUI
				RegisterConfiguredHotkeys
				RefreshHotkeysList
				
				Show-DarkMessageBox "Hotkey '$newKey' assigned as '$groupName' to:`n$($identities -join "`n")" 'Hotkeys (Load Pos) Created' 'Ok' 'Information' 'success'
			}
		}
		ResizeGrip                 = @{
			MouseDown = {
				param($s, $e)
				if ($e.Button -eq 'Left') {
					[Custom.Native]::ReleaseCapture()
					[Custom.Native]::SendMessage($global:DashboardConfig.UI.MainForm.Handle, 0x112, 0xF006, 0)
				}
				if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
				RegisterConfiguredHotkeys
				RefreshHotkeysList
			}
		}
		#endregion
		#region SettingsForm
		settingsTabs               = @{ 
			MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.SettingsForm.Handle, 0xA1, 0x2, 0) } 
		}
		Browse                     = @{ 
			Click = { $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Executable Files (*.exe)|*.exe|All Files (*.*)|*.*'; if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $global:DashboardConfig.UI.InputLauncher.Text = $d.FileName } } 
		}
		BrowseJunction             = @{
			Click = {
				$fb = New-Object Custom.FolderBrowser
				$fb.Description = 'Select the destination folder for the Client Copy'
				if (-not [string]::IsNullOrWhiteSpace($global:DashboardConfig.UI.InputJunction.Text))
				{
					$fb.SelectedPath = $global:DashboardConfig.UI.InputJunction.Text
				}
				if ($fb.ShowDialog($global:DashboardConfig.UI.SettingsForm) -eq [System.Windows.Forms.DialogResult]::OK)
				{
					$global:DashboardConfig.UI.InputJunction.Text = $fb.SelectedPath
				}
			}
		}
		InputJunction              = @{
			Leave = {
				if ([string]::IsNullOrWhiteSpace($global:DashboardConfig.UI.InputJunction.Text))
				{
					$global:DashboardConfig.UI.InputJunction.Text = $global:DashboardConfig.Paths.Profiles
				}
			}
		}
		StartJunction              = @{
			Click = {
				$srcExe = $global:DashboardConfig.UI.InputLauncher.Text
				$baseParentDir = $global:DashboardConfig.UI.InputJunction.Text

				if ([string]::IsNullOrWhiteSpace($srcExe) -or -not (Test-Path $srcExe))
				{
					Show-DarkMessageBox "Please select a valid Launcher executable first with the first Browse button.`nYou must select the Main Launcher!" 'Wrong Launcher.exe' 'OK' 'Error'
					return
				}
				if ([string]::IsNullOrWhiteSpace($baseParentDir))
				{
					Show-DarkMessageBox 'Please select a destination folder.' 'Error' 'OK' 'Error'
					return
				}

				$sourceDir = Split-Path $srcExe -Parent
				$folderName = Split-Path $sourceDir -Leaf
				if ($folderName -in @('bin32', 'bin64'))
				{
					$parentOfBin = Split-Path $sourceDir -Parent
					$folderName = Split-Path $parentOfBin -Leaf
				}
				if ([string]::IsNullOrWhiteSpace($folderName)) { $folderName = 'Client' }

				$userProvidedProfileName = ShowInputBox -Title 'New Profile Name' -Prompt 'Enter a name for the new client profile folder:' -DefaultText "${folderName}_NewProfile"

				[string]$finalProfileName = $null
				[string]$finalDestDir = $null

				if (-not [string]::IsNullOrWhiteSpace($userProvidedProfileName))
				{
					$userProvidedProfileName = $userProvidedProfileName -replace '[\\\/:*?"<>|\x00-\x1F]', '_'
					$potentialDestDir = Join-Path $baseParentDir $userProvidedProfileName

					if (-not (Test-Path $potentialDestDir))
					{
						$finalProfileName = $userProvidedProfileName
						$finalDestDir = $potentialDestDir
					}
					else
					{
						Show-DarkMessageBox "The profile name '$userProvidedProfileName' already exists. Using default naming scheme." 'Profile already exists' 'OK' 'Warning'
						$baseNameFallback = "${folderName}_Copy"
						$counter = 1
						$destDirFallback = Join-Path $baseParentDir $baseNameFallback
						while (Test-Path $destDirFallback)
						{
							$destDirFallback = Join-Path $baseParentDir "${baseNameFallback}_${counter}"
							$counter++
						}
						$finalProfileName = Split-Path $destDirFallback -Leaf
						$finalDestDir = $destDirFallback
					}
				}
				else
				{
					$baseNameFallback = "${folderName}_Copy"
					$counter = 1
					$destDirFallback = Join-Path $baseParentDir $baseNameFallback
					while (Test-Path $destDirFallback)
					{
						$destDirFallback = Join-Path $baseParentDir "${baseNameFallback}_${counter}"
						$counter++
					}
					$finalProfileName = Split-Path $destDirFallback -Leaf
					$finalDestDir = $destDirFallback
				}

				if ([string]::IsNullOrWhiteSpace($finalDestDir) -or [string]::IsNullOrWhiteSpace($finalProfileName))
				{
					Show-DarkMessageBox 'Failed to determine a valid profile name or destination directory.' 'Error' 'OK' 'Error'
					return
				}

				try
				{
					$sourceFull = (Get-Item $sourceDir).FullName.TrimEnd('\')
					$destFull = $finalDestDir.TrimEnd('\')
					if ($destFull -eq $sourceFull -or $destFull.StartsWith($sourceFull + '\', [StringComparison]::OrdinalIgnoreCase))
					{
						Show-DarkMessageBox "The destination folder cannot be inside the source game folder.`nThis would create an endless copy loop.`n`nEither use a valid folder outside the source or use the default path by deleting the path in the textbox." 'Invalid Destination' 'OK' 'Error'
						return
					}
				}
				catch {}

				$confirm = Show-DarkMessageBox "This will create Junctions for the folders 'Data' and 'Effect'.`nAlso a copy of all other files from the $folderName directory to:`n$finalDestDir`n`nProceed?" 'Confirm Junction & Copy' 'YesNo' 'Question'

				if ($confirm -eq 'Yes')
				{
					try
					{
						if (-not (Test-Path $finalDestDir)) { New-Item -ItemType Directory -Path $finalDestDir -Force | Out-Null }

						$junctionFolders = @('Data', 'Effect')

						$global:DashboardConfig.UI.SettingsForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

						Get-ChildItem -Path $sourceDir | ForEach-Object {
							$itemName = $_.Name
							$sourcePath = $_.FullName
							$targetPath = Join-Path $finalDestDir $itemName

							if ($itemName -in $junctionFolders)
							{
								if (Test-Path $targetPath)
								{
								}
								else
								{
									$cmdArgs = "/c mklink /J `"$targetPath`" `"$sourcePath`""
									Start-Process -FilePath 'cmd.exe' -ArgumentList $cmdArgs -WindowStyle Hidden -Wait
								}
							}
							else
							{
								Copy-Item -Path $sourcePath -Destination $targetPath -Recurse -Force
							}
						}

						$global:DashboardConfig.UI.SettingsForm.Cursor = [System.Windows.Forms.Cursors]::Default

						$global:DashboardConfig.UI.ProfileGrid.Rows.Add($finalProfileName, $finalDestDir) | Out-Null
						Show-DarkMessageBox "Junctions created and Profile '$finalProfileName' added successfully.`nPlease remember that 'neuz.exe' must be manually patched for each profile on new Patches." 'Profile Created' 'OK' 'Information' 'success'
						SyncProfilesToConfig
						RefreshLoginProfileSelector
						WriteConfig

					}
					catch
					{
						$global:DashboardConfig.UI.SettingsForm.Cursor = [System.Windows.Forms.Cursors]::Default
						Show-DarkMessageBox "An error occurred:`n$($_.Exception.Message)`n`nNote: Creating Junctions may require running the Dashboard as Administrator.`nBe sure you used the correct paths and have enough disk space." 'Error' 'OK' 'Error'
					}
				}
			}
		}
		CopyNeuz                   = @{ 
			Click = {
				$confirm = Show-DarkMessageBox "Do you want to copy the 'neuz.exe' from your main client's bin32 and bin64 folders to all your profiles.`nProfiles that were not created through junctions or have a different neuz.exe might break.`n`nThis action is only working if you patched your Main Launcher folder first!`nProcced?" 'Confirm Neuz.exe Copy' 'YesNo' 'Warning'

				if ($confirm -eq 'Yes')
				{
					$mainLauncherPath = $global:DashboardConfig.UI.InputLauncher.Text
					if ([string]::IsNullOrWhiteSpace($mainLauncherPath) -or -not (Test-Path $mainLauncherPath))
					{
						Show-DarkMessageBox 'Please select a valid Main Launcher Path first.' 'Error' 'OK' 'Error'
						return
					}

					$mainClientDir = Split-Path $mainLauncherPath -Parent
					$sourceBin32 = Join-Path $mainClientDir 'bin32\neuz.exe'
					$sourceBin64 = Join-Path $mainClientDir 'bin64\neuz.exe'

					$profiles = $global:DashboardConfig.Config['Profiles']
					if (-not $profiles -or $profiles.Count -eq 0)
					{
						Show-DarkMessageBox 'No profiles found. Please create or add profiles first.' 'No Profiles found' 'OK' 'Information'
						return
					}

					$copiedCount = 0
					$failedCopies = @()

					foreach ($profileName in $profiles.Keys)
					{
						$profilePath = $profiles[$profileName]

						$targetBin32 = Join-Path $profilePath 'bin32\neuz.exe'
						$targetBin64 = Join-Path $profilePath 'bin64\neuz.exe'

						try
						{
							if (Test-Path $sourceBin32)
							{
								if (-not (Test-Path (Split-Path $targetBin32 -Parent)))
								{
									New-Item -ItemType Directory -Path (Split-Path $targetBin32 -Parent) -Force | Out-Null
								}
								Copy-Item -Path $sourceBin32 -Destination $targetBin32 -Force -ErrorAction Stop
								$copiedCount++
							}
							if (Test-Path $sourceBin64)
							{
								if (-not (Test-Path (Split-Path $targetBin64 -Parent)))
								{
									New-Item -ItemType Directory -Path (Split-Path $targetBin64 -Parent) -Force | Out-Null
								}
								Copy-Item -Path $sourceBin64 -Destination $targetBin64 -Force -ErrorAction Stop
								$copiedCount++
							}
						}
						catch
						{
							$failedCopies += "Profile '$profileName': $($_.Exception.Message)"
						}
					}

					if ($failedCopies.Count -eq 0)
					{
						Show-DarkMessageBox "Successfully copied neuz.exe to $copiedCount locations." 'Patch complete!' 'OK' 'Information' 'success'
					}
					else
					{
						Show-DarkMessageBox "Copied neuz.exe to $($copiedCount - $failedCopies.Count) locations.`n`nFailed to copy to the following profiles:`n" + ($failedCopies -join "`n") 'Patch Error' 'OK' 'Warning'
					}
				}
			}
		}
		ProfileGrid                = @{
			CellClick       = {
				param($s, $e)
				if ($e.RowIndex -lt 0) { return }


				if ($e.ColumnIndex -eq 0 -or $e.ColumnIndex -eq 1)
				{
					$s.Tag = $e.RowIndex
				}
				elseif ($e.ColumnIndex -eq 2 -or $e.ColumnIndex -eq 3)
				{
					$row = $s.Rows[$e.RowIndex]

					$currentVal = $false
					if ($row.Cells[$e.ColumnIndex].Value -eq $true -or $row.Cells[$e.ColumnIndex].Value -eq 1)
					{
						$currentVal = $true
					}
					$row.Cells[$e.ColumnIndex].Value = -not $currentVal
					$s.EndEdit()

					$prevIndex = $s.Tag

					if ($null -ne $prevIndex -and $prevIndex -gt -1 -and $prevIndex -lt $s.Rows.Count)
					{
						$s.ClearSelection()
						$s.Rows[$prevIndex].Selected = $true
						try { $s.CurrentCell = $s.Rows[$prevIndex].Cells[0] } catch {}
					}
					else
					{
						$s.ClearSelection()
					}
				}
			}
			CellDoubleClick = {
				param($s, $e)
				if ($e.RowIndex -lt 0) { return }

				if ($e.ColumnIndex -eq 1)
				{
					$path = $s.Rows[$e.RowIndex].Cells[1].Value
					if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path))
					{
						try { Invoke-Item $path } catch {}
					}
				}
			}
		}
		AddProfile                 = @{
			Click = {
				$fb = New-Object Custom.FolderBrowser
				$fb.Description = 'Select an existing Client folder'
				if ($fb.ShowDialog($global:DashboardConfig.UI.SettingsForm) -eq [System.Windows.Forms.DialogResult]::OK)
				{
					$path = $fb.SelectedPath
					$defaultName = Split-Path $path -Leaf

					$profName = ShowInputBox -Title 'Profile Name' -Prompt 'Enter a name for this profile:' -DefaultText $defaultName

					if (-not [string]::IsNullOrWhiteSpace($profName))
					{
						$global:DashboardConfig.UI.ProfileGrid.Rows.Add($profName, $path) | Out-Null
						SyncProfilesToConfig
						RefreshLoginProfileSelector
						WriteConfig
					}
				}
			}
		}
		RenameProfile              = @{
			Click = {
				if ($global:DashboardConfig.UI.ProfileGrid.SelectedRows.Count -gt 0)
				{
					$row = $global:DashboardConfig.UI.ProfileGrid.SelectedRows[0]
					$oldName = $row.Cells[0].Value

					$newName = ShowInputBox -Title 'Rename Profile' -Prompt 'Enter new profile name:' -DefaultText $oldName

					if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne $oldName)
					{
						$row.Cells[0].Value = $newName

						if ($global:DashboardConfig.Config['LoginConfig'].Contains($oldName))
						{
							$data = $global:DashboardConfig.Config['LoginConfig'][$oldName]
							$global:DashboardConfig.Config['LoginConfig'].Remove($oldName)
							$global:DashboardConfig.Config['LoginConfig'][$newName] = $data
						}

						if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['SelectedProfile'] -eq $oldName)
						{
							$global:DashboardConfig.Config['Options']['SelectedProfile'] = $newName
						}

						$configKeys = [System.Collections.ArrayList]@($global:DashboardConfig.Config.Keys)
						foreach ($key in $configKeys)
						{
							if ($key -match '^Setup_')
							{
								$setupData = $global:DashboardConfig.Config[$key]
								$clientKeys = [System.Collections.ArrayList]@($setupData.Keys)
								foreach ($clientKey in $clientKeys)
								{
									$val = $setupData[$clientKey]
									$parts = $val -split ','
									if ($parts.Count -ge 2 -and $parts[1] -eq $oldName)
									{
										$parts[1] = $newName
										if ($parts.Count -ge 4)
										{
											$title = $parts[3]
											if ($title -match "^\[$([regex]::Escape($oldName))\]")
											{
												$parts[3] = $title -replace "^\[$([regex]::Escape($oldName))\]", "[$newName]"
											}
										}
										$setupData[$clientKey] = $parts -join ','
									}
								}
							}
						}

						if ($global:DashboardConfig.Config.Contains('SavedWindowPositions'))
						{
							$swp = $global:DashboardConfig.Config['SavedWindowPositions']
							$swpKeys = [System.Collections.ArrayList]@($swp.Keys)
							foreach ($key in $swpKeys)
							{
								if ($key -match "^\[$([regex]::Escape($oldName))\](.*)")
								{
									$rest = $Matches[1]
									$newKey = "[$newName]$rest"
									$val = $swp[$key]
									$swp.Remove($key)
									$swp[$newKey] = $val
								}
							}
						}

						if ($global:DashboardConfig.Config.Contains('HotkeyGroups'))
						{
							$hkg = $global:DashboardConfig.Config['HotkeyGroups']
							$hkgKeys = [System.Collections.ArrayList]@($hkg.Keys)
							foreach ($groupName in $hkgKeys)
							{
								$memberString = $hkg[$groupName]
								if (-not [string]::IsNullOrWhiteSpace($memberString))
								{
									$members = $memberString -split ','
									$newMembers = @()
									$modified = $false
									foreach ($member in $members)
									{
										if ($member -match "^\[$([regex]::Escape($oldName))\](.*)")
										{
											$rest = $Matches[1]
											$newMembers += "[$newName]$rest"
											$modified = $true
										}
										else
										{
											$newMembers += $member
										}
									}
									if ($modified)
									{
										$hkg[$groupName] = ($newMembers -join ',')
									}
								}
							}
						}

						SyncProfilesToConfig
						RefreshLoginProfileSelector
						WriteConfig
					}
				}
				else
				{
					Show-DarkMessageBox 'Please select a profile first.' 'Select Profile' 'OK' 'Warning'
				}
			}
		}
		RemoveProfile              = @{
			Click = {
				$rowsToRemove = $global:DashboardConfig.UI.ProfileGrid.SelectedRows | ForEach-Object { $_ }
				foreach ($row in $rowsToRemove)
				{
					$profileName = $row.Cells[0].Value.ToString()
					$profilePath = $row.Cells[1].Value.ToString()
					$null = $logDir; $logDir = Join-Path $profilePath 'Log'


					if ($global:DashboardConfig.Config['LoginConfig'].Contains($profileName))
					{
						$global:DashboardConfig.Config['LoginConfig'].Remove($profileName)
					}

					$global:DashboardConfig.UI.ProfileGrid.Rows.Remove($row)
				}
				SyncProfilesToConfig
				RefreshLoginProfileSelector
				WriteConfig
			}
		}
		DeleteProfile              = @{
			Click = {
				$rowsToDelete = $global:DashboardConfig.UI.ProfileGrid.SelectedRows | ForEach-Object { $_ }
				if ($rowsToDelete.Count -eq 0)
				{
					Show-DarkMessageBox 'Please select a profile first.' 'Select Profile' 'OK' 'Warning'
					return
				}

				$confirm = Show-DarkMessageBox "Are you sure you want to PERMANENTLY DELETE the selected profile(s) from the hard drive?`n`nThis action cannot be undone." 'Are You Sure?' 'YesNo' 'Warning'

				if ($confirm -eq 'Yes')
				{
					foreach ($row in $rowsToDelete)
					{
						$profileName = $row.Cells[0].Value.ToString()
						$profilePath = $row.Cells[1].Value.ToString()
						$null = $logDir; $logDir = Join-Path $profilePath 'Log'

						try
						{
							if (Test-Path $profilePath) { Remove-Item -Path $profilePath -Recurse -Force -ErrorAction Stop }
							if ($global:DashboardConfig.Config['LoginConfig'].Contains($profileName)) { $global:DashboardConfig.Config['LoginConfig'].Remove($profileName) }
							$global:DashboardConfig.UI.ProfileGrid.Rows.Remove($row)
						}
						catch
						{
						}
					}
					SyncProfilesToConfig
					RefreshLoginProfileSelector
					WriteConfig
				}
			}
		}
		SetupGrid                  = @{
			CellClick = {
				param($s, $e)
				if ($global:DashboardConfig.State.SetupEditMode -and $e.RowIndex -ge 0)
				{
					if ($e.ColumnIndex -eq 2)
					{
						$row = $s.Rows[$e.RowIndex]
						$val = $row.Cells[2].Value
						$row.Cells[2].Value = -not ($val -eq $true -or $val -eq 1)
					}
					elseif ($e.ColumnIndex -eq 3)
					{
						$s.Rows.RemoveAt($e.RowIndex)
					}
				}
			}
		}
		EditSetup                  = @{
			Click = {
				$UI = $global:DashboardConfig.UI
				if ($global:DashboardConfig.State.SetupEditMode -and $UI.EditSetup.Text -eq 'Save')
				{
					$setupName = $global:DashboardConfig.State.EditingSetupName
					$setupKey = "Setup_$setupName"

					$newConfigSection = [ordered]@{}
					$clientIndex = 0
					$profileCounters = @{}

					foreach ($row in $UI.SetupGrid.Rows)
					{
						$p = $row.Cells[0].Value
						$title = $row.Cells[1].Value
						$loginFlag = if ($row.Cells[2].Value) { 1 } else { 0 }

						if (-not $profileCounters.ContainsKey($p)) { $profileCounters[$p] = 0 }
						$profileCounters[$p]++
						$profileIndex = $profileCounters[$p]

						$gridPos = $profileIndex
						$safeTitle = $title -replace ',', ''

						if ($safeTitle -notmatch "^\[$([regex]::Escape($p))\]")
						{
							$safeTitle = "[$p]$safeTitle"
						}

						$valueString = "$($gridPos),$($p),$($profileIndex),$safeTitle,$loginFlag"
						$newConfigSection["Client$clientIndex"] = $valueString
						$clientIndex++
					}

					$global:DashboardConfig.Config[$setupKey] = $newConfigSection
					if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }

					$global:DashboardConfig.State.SetupEditMode = $false
					$global:DashboardConfig.State.EditingSetupName = $null

					$UI.EditSetup.Text = 'Edit'
					$UI.DeleteSetup.Text = 'Delete'
					
					$UI.AddSetup.Visible = $true
					$UI.RenameSetup.Visible = $true
					$UI.AddSetup.Enabled = $true
					$UI.RenameSetup.Enabled = $true
					
					$UI.SetupGrid.SelectionMode = 'FullRowSelect'
					$UI.SetupGrid.EditMode = 'EditProgrammatically'

					$UI.SetupGrid.Columns.Clear()
					$colSetupName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSetupName.HeaderText = 'Setup Name'; $colSetupName.FillWeight = 60; $colSetupName.ReadOnly = $true
					$colSetupCount = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSetupCount.HeaderText = 'Clients'; $colSetupCount.FillWeight = 40; $colSetupCount.ReadOnly = $true
					$UI.SetupGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colSetupName, $colSetupCount))

					SyncConfigToUI
				}
				else
				{
					if ($UI.SetupGrid.SelectedRows.Count -gt 0)
					{
						$row = $UI.SetupGrid.SelectedRows[0]
						$setupName = $row.Cells[0].Value
						if (-not [string]::IsNullOrWhiteSpace($setupName))
						{
							$setupKey = "Setup_$setupName"
							if (-not $global:DashboardConfig.Config.Contains($setupKey)) { return }

							$global:DashboardConfig.State.SetupEditMode = $true
							$global:DashboardConfig.State.EditingSetupName = $setupName

							$UI.EditSetup.Text = 'Save'
							$UI.DeleteSetup.Text = 'Cancel'
							
							$UI.AddSetup.Visible = $false
							$UI.RenameSetup.Visible = $false
							$UI.AddSetup.Enabled = $false
							$UI.RenameSetup.Enabled = $false
							
							$UI.SetupGrid.SelectionMode = 'CellSelect'
							$UI.SetupGrid.EditMode = 'EditOnKeystrokeOrF2'

							$UI.SetupGrid.Columns.Clear()

							$colProfile = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
							$colProfile.HeaderText = 'Profile'
							$colProfile.ReadOnly = $false
							$colProfile.FillWeight = 14
							$colProfile.ToolTipText = 'The client profile for this entry.'

							$colTitle = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
							$colTitle.HeaderText = 'Title'
							$colTitle.ReadOnly = $false
							$colTitle.FillWeight = 18
							$colTitle.ToolTipText = 'The window title or identifier for this client.'

							$colAutoLogin = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
							$colAutoLogin.HeaderText = 'Login'
							$colAutoLogin.FillWeight = 10
							$colAutoLogin.ReadOnly = $true
							$colAutoLogin.ToolTipText = 'Check this if this client should be automatically logged in.'

							$colRemove = New-Object System.Windows.Forms.DataGridViewButtonColumn
							$colRemove.HeaderText = ''
							$colRemove.Text = 'X'
							$colRemove.UseColumnTextForButtonValue = $true
							$colRemove.FillWeight = 5
							$colRemove.ToolTipText = 'Click to remove this client from the setup.'
							$colRemove.FlatStyle = 'Flat'
							$colRemove.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
							$colRemove.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
							$colRemove.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
							$colRemove.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
							$colRemove.DefaultCellStyle.Font = (New-Object System.Drawing.Font('Segoe UI', 8))

							$colExtra = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
							$colExtra.HeaderText = ' '
							$colExtra.ReadOnly = $true
							$colExtra.FillWeight = 2

							$UI.SetupGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colProfile, $colTitle, $colAutoLogin, $colRemove, $colExtra))
							
							$UI.SetupGrid.Rows.Clear()

							$configSection = $global:DashboardConfig.Config[$setupKey]
							$clients = @($configSection.GetEnumerator()) | Sort-Object { if ($_.Key -match 'Client(\d+)') { [int]$Matches[1] } else { 9999 } }

							foreach ($clientEntry in $clients)
							{
								$parts = $clientEntry.Value -split ',', 5
								if ($parts.Count -ge 4)
								{
									$gridPos = $parts[0]
									$p = $parts[1]
									$null = $profIndex; $profIndex = $parts[2]
									$title = $parts[3]
									if ($title -match '^\[.*?\](.*)') { $title = $matches[1] }
									$loginFlag = if ($parts.Count -ge 5) { ($parts[4] -eq '1') } else { $false }

									$rowIndex = $UI.SetupGrid.Rows.Add($p, $title, $loginFlag)
									$UI.SetupGrid.Rows[$rowIndex].Tag = @{
										GridPosition = $gridPos
									}
								}
							}
							$UI.SetupGrid.ClearSelection()
						}
					}
				}
			}
		}
		AddSetup                   = @{
			Click = {
				
				$grid = $global:DashboardConfig.UI.DataGridMain
				$selectedRows = $grid.SelectedRows
				if ($selectedRows.Count -eq 0)
				{
					Show-DarkMessageBox "Please select the clients you want to include in the setup." 'No Selection' 'OK' 'Warning'
					return
				}

				$detailedStateList = [System.Collections.Generic.List[PSObject]]::new()
				$profileCounters = @{}

				
				$sortedRows = $selectedRows | Sort-Object Index

				foreach ($row in $sortedRows)
				{
					$gridPosition = $row.Index + 1
					$fullTitle = $row.Cells[1].Value.ToString()
					$profileName = 'Default'
					$cleanTitle = $fullTitle

					if ($fullTitle -match '^\[([^\]]+)\](.*)')
					{
						$profileName = $matches[1]
						$cleanTitle = $matches[2].Trim()
					}

					$isLoggedIn = $false
					$serverName = $cleanTitle
					$charName = ''

					if ($cleanTitle -match '^(.*?) - (.*)$')
					{
						$serverName = $matches[1]
						$charName = $matches[2]
						$isLoggedIn = $true
					}

					if (-not $profileCounters.ContainsKey($profileName))
					{
						$profileCounters[$profileName] = 0
					}
					$profileCounters[$profileName]++
					$profileIndex = $profileCounters[$profileName]

					$detailedStateList.Add([PSCustomObject]@{
							GridPosition = $profileIndex
							Profile      = $profileName
							ProfileIndex = $profileIndex
							FullTitle    = $fullTitle
							ServerName   = $serverName
							Character    = $charName
							IsLoggedIn   = $isLoggedIn
						})
				}

				$setupName = ShowInputBox -Title 'New Setup' -Prompt 'Enter a name for this setup:' -DefaultText 'MySetup'
				if ([string]::IsNullOrWhiteSpace($setupName)) { return }
				$setupName = $setupName -replace '[^a-zA-Z0-9_\-]', ''
				
				
				
				$configSection = [ordered]@{}
				for ($i = 0; $i -lt $detailedStateList.Count; $i++)
				{
					$client = $detailedStateList[$i]
					$loginFlag = if ($client.IsLoggedIn) { 1 } else { 0 }

					
					$safeTitle = $client.FullTitle -replace ',', ''

					$valueString = "$($client.GridPosition),$($client.Profile),$($client.ProfileIndex),$safeTitle,$loginFlag"
					$configSection["Client$i"] = $valueString
				}

				$global:DashboardConfig.Config["Setup_$setupName"] = $configSection
				if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
				SyncConfigToUI

				Show-DarkMessageBox "Setup '$setupName' saved successfully with $($detailedStateList.Count) clients." 'Setup Saved' 'OK' 'Information' 'success'
			}
		}
		RenameSetup                = @{
			Click = {
				if ($global:DashboardConfig.State.SetupEditMode) { return }
				if ($global:DashboardConfig.UI.SetupGrid.SelectedRows.Count -gt 0)
				{
					$row = $global:DashboardConfig.UI.SetupGrid.SelectedRows[0]
					$oldName = $row.Cells[0].Value
					$newName = ShowInputBox -Title 'Rename Setup' -Prompt 'Enter new setup name:' -DefaultText $oldName
					if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne $oldName)
					{
						$newName = $newName -replace '[^a-zA-Z0-9_\-]', ''
						$oldKey = "Setup_$oldName"
						$newKey = "Setup_$newName"
						if ($global:DashboardConfig.Config.Contains($oldKey))
						{
							$data = $global:DashboardConfig.Config[$oldKey]
							$global:DashboardConfig.Config.Remove($oldKey)
							$global:DashboardConfig.Config[$newKey] = $data
							if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
							SyncConfigToUI
						}
					}
				}
			}
		}
		DeleteSetup                = @{
			Click = {
				$UI = $global:DashboardConfig.UI
				if ($global:DashboardConfig.State.SetupEditMode)
				{	
					$editingSetupName = $global:DashboardConfig.State.EditingSetupName
					$global:DashboardConfig.State.SetupEditMode = $false
					$global:DashboardConfig.State.EditingSetupName = $null

					$UI.EditSetup.Text = 'Edit'
					$UI.DeleteSetup.Text = 'Delete'
					
					$UI.AddSetup.Visible = $true
					$UI.RenameSetup.Visible = $true
					$UI.AddSetup.Enabled = $true
					$UI.RenameSetup.Enabled = $true
					
					$UI.SetupGrid.SelectionMode = 'FullRowSelect'
					$UI.SetupGrid.EditMode = 'EditProgrammatically'

					$UI.SetupGrid.Columns.Clear()
					$colSetupName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSetupName.HeaderText = 'Setup Name'; $colSetupName.FillWeight = 60; $colSetupName.ReadOnly = $true
					$colSetupCount = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSetupCount.HeaderText = 'Clients'; $colSetupCount.FillWeight = 40; $colSetupCount.ReadOnly = $true
					$UI.SetupGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colSetupName, $colSetupCount))

					SyncConfigToUI

					if (-not [string]::IsNullOrEmpty($editingSetupName))
					{
						foreach ($r in $UI.SetupGrid.Rows)
						{
							if ($r.Cells[0].Value -eq $editingSetupName)
							{
								$r.Selected = $true
								break
							}
						}
					}
				}
				else
				{
					if ($UI.SetupGrid.SelectedRows.Count -gt 0)
					{
						$row = $UI.SetupGrid.SelectedRows[0]
						$name = $row.Cells[0].Value
						if ((Show-DarkMessageBox "Are you sure you want to delete setup '$name'?" 'Confirm Delete' 'YesNo' 'Warning') -eq 'Yes')
						{
							$key = "Setup_$name"
							if ($global:DashboardConfig.Config.Contains($key))
							{
								$global:DashboardConfig.Config.Remove($key)
								if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
								SyncConfigToUI
							}
						}
					}
				}
			}
		}
		LoginConfigGrid            = @{
			CellClick = {
				param($s, $e)
				if ($e.RowIndex -lt 0) { return }

				$row = $s.Rows[$e.RowIndex]
				$colIndex = $e.ColumnIndex

				if ($colIndex -eq 1) { $row.Cells[1].Value = if ($row.Cells[1].Value -eq '1') { '2' } else { '1' } }
				elseif ($colIndex -eq 2) { $row.Cells[2].Value = if ($row.Cells[2].Value -eq '1') { '2' } else { '1' } }
				elseif ($colIndex -eq 3) { $row.Cells[3].Value = switch ($row.Cells[3].Value) { '1' {'2'}; '2' {'3'}; '3' {'4'}; '4' {'5'}; '5' {'6'};'6' {'1'}; default {'1'} } }
				elseif ($colIndex -eq 4) { $row.Cells[4].Value = if ($row.Cells[4].Value -eq 'Yes') { 'No' } else { 'Yes' } }
			}
		}
		LoginProfileSelector       = @{
			SelectedIndexChanged = {
				param($s, $e)
				$oldProfile = $s.Tag
				if (-not $oldProfile) { $oldProfile = 'Default' }

				Write-Verbose "  UI: Saving settings for '$oldProfile' before switching..."
				SyncUIToConfig -ProfileToSave $oldProfile
				if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }

				Write-Verbose '  UI: Loading settings for new selection...'
				SyncConfigToUI
			}
		}
		MasterPauseHotkeyBtn       = @{
			Click = {
				$currentKey = $global:DashboardConfig.UI.MasterPauseHotkeyBtn.Text
				$newKey = Show-KeyCaptureDialog -currentKey $currentKey -OwnerForm $global:DashboardConfig.UI.SettingsForm
				if ($null -ne $newKey -and $newKey -ne $currentKey)
				{
					if (-not $global:DashboardConfig.Config.Contains('GlobalHotkeys')) { $global:DashboardConfig.Config['GlobalHotkeys'] = [ordered]@{} }
					$global:DashboardConfig.Config['GlobalHotkeys']['MasterPauseHotkey'] = $newKey
					$global:DashboardConfig.UI.MasterPauseHotkeyBtn.Text = $newKey

					if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }

					RegisterConfiguredHotkeys
					RefreshHotkeysList
				}
			}
		}
		MasterResumeHotkeyBtn      = @{
			Click = {
				$currentKey = $global:DashboardConfig.UI.MasterResumeHotkeyBtn.Text
				$newKey = Show-KeyCaptureDialog -currentKey $currentKey -OwnerForm $global:DashboardConfig.UI.SettingsForm
				if ($null -ne $newKey -and $newKey -ne $currentKey)
				{
					if (-not $global:DashboardConfig.Config.Contains('GlobalHotkeys')) { $global:DashboardConfig.Config['GlobalHotkeys'] = [ordered]@{} }
					$global:DashboardConfig.Config['GlobalHotkeys']['MasterResumeHotkey'] = $newKey
					$global:DashboardConfig.UI.MasterResumeHotkeyBtn.Text = $newKey

					if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }

					RegisterConfiguredHotkeys
					RefreshHotkeysList
				}
			}
		}
		Save                       = @{
			Click = {
				SyncUIToConfig
				WriteConfig
				HideSettingsForm
			}
		}
		Cancel                     = @{ 
			Click = { HideSettingsForm } 
		}
		UnregisterHotkey           = @{
			Click = {
				$grid = $global:DashboardConfig.UI.HotkeysGrid
				if (-not $grid) { return }
				
				$selectedRows = $grid.SelectedRows
				if ($selectedRows.Count -eq 0) { return }

				if ((Show-DarkMessageBox "Are you sure you want to unregister and delete $($selectedRows.Count) hotkey(s)?" 'Confirm Unregister' 'YesNo' 'Question') -ne 'Yes') { 
					return 
				}

				foreach ($row in $selectedRows)
				{
					if (-not $row.Tag -or -not $row.Tag.Id) { continue }
					if ($row.Tag)
					{
						$id = $row.Tag.Id
						$ownerKey = $row.Tag.OwnerKey
						
						if (Get-Command UnregisterHotkeyInstance -ErrorAction SilentlyContinue -Verbose:$False)
						{
							UnregisterHotkeyInstance -Id $id -OwnerKey $ownerKey
						}
						
						if ($ownerKey -match '^GroupHotkey_(.+)')
						{
							$groupName = $Matches[1]
							if ($global:DashboardConfig.Config.Contains('Hotkeys') -and $global:DashboardConfig.Config['Hotkeys'].Contains($groupName))
							{
								$global:DashboardConfig.Config['Hotkeys'].Remove($groupName)
							}
							if ($global:DashboardConfig.Config.Contains('HotkeyGroups') -and $global:DashboardConfig.Config['HotkeyGroups'].Contains($groupName))
							{
								$global:DashboardConfig.Config['HotkeyGroups'].Remove($groupName)
							}
						}
						elseif ($ownerKey -match '^GroupHotkeyHide_(.+)')
						{
							$groupName = $Matches[1]
							if ($global:DashboardConfig.Config.Contains('HotkeysHide') -and $global:DashboardConfig.Config['HotkeysHide'].Contains($groupName))
							{
								$global:DashboardConfig.Config['HotkeysHide'].Remove($groupName)
							}
							if ($global:DashboardConfig.Config.Contains('HotkeyGroupsHide') -and $global:DashboardConfig.Config['HotkeyGroupsHide'].Contains($groupName))
							{
								$global:DashboardConfig.Config['HotkeyGroupsHide'].Remove($groupName)
							}
						}
						elseif ($ownerKey -match '^GroupHotkeyLoadPos_(.+)')
						{
							$groupName = $Matches[1]
							if ($global:DashboardConfig.Config.Contains('HotkeysLoadPos') -and $global:DashboardConfig.Config['HotkeysLoadPos'].Contains($groupName))
							{
								$global:DashboardConfig.Config['HotkeysLoadPos'].Remove($groupName)
							}
							if ($global:DashboardConfig.Config.Contains('HotkeyGroupsLoadPos') -and $global:DashboardConfig.Config['HotkeyGroupsLoadPos'].Contains($groupName))
							{
								$global:DashboardConfig.Config['HotkeyGroupsLoadPos'].Remove($groupName)
							}
						}
						elseif ($ownerKey -eq 'master_pause_hotkey' -or $ownerKey -eq 'master_resume_hotkey')
						{
							if ($global:DashboardConfig.Config.Contains('GlobalHotkeys'))
							{
								$configName = if ($ownerKey -eq 'master_pause_hotkey') { 'MasterPauseHotkey' } else { 'MasterResumeHotkey' }
								if ($global:DashboardConfig.Config['GlobalHotkeys'].Contains($configName))
								{
									$global:DashboardConfig.Config['GlobalHotkeys'].Remove($configName)
								}
							}
						}
						
						if ($global:DashboardConfig.Resources.FtoolForms)
						{
							if ($ownerKey -match '^quickcmd_(.+)_(.+)')
							{
								$instanceId = $Matches[1]
								$cmdId = $Matches[2]
								if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
								{
									$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
									if ($form -and -not $form.IsDisposed)
									{
										$data = $form.Tag
										$cmd = $null
										if ($data.SavedCommands) { $cmd = $data.SavedCommands | Where-Object { $_.Id -eq $cmdId } | Select-Object -First 1 }
										if (-not $cmd -and $data.GlobalCommandList) { $cmd = $data.GlobalCommandList | Where-Object { $_.Id -eq $cmdId } | Select-Object -First 1 }
										if ($cmd) { $cmd.Hotkey = $null; $cmd.HotkeyId = $null }
										if (Get-Command UpdateChatCommanderSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateChatCommanderSettings $data -forceWrite }
										if ($form -and -not $form.IsDisposed) { RefreshSavedCommandsUI $form }
									}
								}
							}
							elseif ($ownerKey -match '^macroinst_(.+)')
							{
								$instanceId = $Matches[1]
								if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
								{
									$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
									if ($form -and -not $form.IsDisposed)
									{
										$data = $form.Tag; $data.Hotkey = $null; $data.HotkeyId = $null; $data.BtnHotKey.Text = 'Hotkey'
										if (Get-Command UpdateMacroSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateMacroSettings $data -forceWrite }
									}
								}
							}
							if ($ownerKey -match '^global_toggle_(.+)')
							{
								$instanceId = $Matches[1]
								if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
								{
									$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
									if ($form -and -not $form.IsDisposed)
									{
										$data = $form.Tag; $data.GlobalHotkey = $null; $data.GlobalHotkeyId = $null
										if (Get-Command UpdateSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateSettings $data -forceWrite }
									}
								}
							}
							elseif ($ownerKey -match '^ext_(.+)_\d+')
							{
								if ($global:DashboardConfig.Resources.ExtensionData.Contains($ownerKey))
								{
									$extData = $global:DashboardConfig.Resources.ExtensionData[$ownerKey]
									$extData.Hotkey = $null; $extData.HotkeyId = $null; $extData.BtnHotKey.Text = 'Hotkey'
									$form = $extData.Panel.FindForm()
									if ($form -and $form.Tag) { if (Get-Command UpdateSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateSettings $form.Tag $extData -forceWrite } }
								}
							}
							elseif ($global:DashboardConfig.Resources.FtoolForms.Contains($ownerKey))
							{
								$form = $global:DashboardConfig.Resources.FtoolForms[$ownerKey]
								if ($form -and -not $form.IsDisposed) { $data = $form.Tag; $data.Hotkey = $null; $data.HotkeyId = $null; $data.BtnHotKey.Text = 'Hotkey'; if (Get-Command UpdateSettings -ErrorAction SilentlyContinue -Verbose:$False) { UpdateSettings $data -forceWrite } }
							}
						}
					}
				}
				if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
				SyncConfigToUI
				RegisterConfiguredHotkeys
				RefreshHotkeysList
			}
		}

		#endregion
		#region ExtraForm
		WorldBossListener          = @{
			CheckedChanged = {
				param($s, $e)

				if (-not $global:DashboardConfig.Config.Contains('Options')) { $global:DashboardConfig.Config['Options'] = [ordered]@{} }
				$newValue = if ($s.Checked) { '1' } else { '0' }
				if ($global:DashboardConfig.Config['Options']['WorldBossListener'] -ne $newValue) {
					$global:DashboardConfig.Config['Options']['WorldBossListener'] = $newValue
					if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
				}

				if ($s.Checked) {
					if (Get-Command Start-WorldBossListener -ErrorAction SilentlyContinue -Verbose:$False) { Start-WorldBossListener }
				} else {
					if (Get-Command Stop-WorldBossListener -ErrorAction SilentlyContinue -Verbose:$False) { Stop-WorldBossListener }
				}
			}
		}
		RefreshBosses              = @{
			Click = {
				if ($global:DashboardConfig.State.WorldBossListener -and $global:DashboardConfig.State.WorldBossListener.Timer) {
					$global:DashboardConfig.State.WorldBossListener.IsFirstRun = $true
					$global:DashboardConfig.State.WorldBossListener.NextRunTime = [DateTime]::MinValue
				} else {
					if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "World Boss" -Message "Listener is not running." -Type "Warning" }
				}
			}
		}
		NotificationPositionSelector = @{
			SelectedIndexChanged = {
				param($s, $e)
				if (-not $global:DashboardConfig.Config.Contains('Options')) { $global:DashboardConfig.Config['Options'] = [ordered]@{} }
				$newValue = $s.SelectedItem
				if ($newValue -and $global:DashboardConfig.Config['Options']['NotificationPosition'] -ne $newValue) {
					$global:DashboardConfig.Config['Options']['NotificationPosition'] = $newValue
					if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
					if (Get-Command UpdateNotificationPositions -ErrorAction SilentlyContinue -Verbose:$False) { UpdateNotificationPositions }
				}
			}
		}
		ShowBossImages             = @{
			CheckedChanged = {
				param($s, $e)
				if (-not $global:DashboardConfig.Config.Contains('Options')) { $global:DashboardConfig.Config['Options'] = [ordered]@{} }
				$newValue = if ($s.Checked) { '1' } else { '0' }
				if ($global:DashboardConfig.Config['Options']['ShowBossImages'] -ne $newValue) {
					$global:DashboardConfig.Config['Options']['ShowBossImages'] = $newValue
					if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
				}

				if ($global:DashboardConfig.UI.ButtonPanel) {
					foreach ($ctrl in $global:DashboardConfig.UI.ButtonPanel.Controls) {
						if ($ctrl -is [System.Windows.Forms.Button]) {
							Update-BossButtonImage $ctrl
						}
					}
				}

				if ($global:DashboardConfig.Resources.Timers['BossTimer']) {
					$t = $global:DashboardConfig.Resources.Timers['BossTimer']
					$t.Stop()
					$t.Interval = 1000
					$t.Start()
				}
			}
		}
		NoteGrid                   = @{
			DoubleClick = { if (Get-Command EditNote -ErrorAction SilentlyContinue -Verbose:$False) { EditNote } }
		}
		NotificationGrid           = @{
			CellDoubleClick = {
				param($s, $e)
				if ($e.RowIndex -ge 0)
				{
					$row = $s.Rows[$e.RowIndex]
					if ($row.Tag)
					{
						$entry = $row.Tag
						if ($entry.ExtraData -and $entry.ExtraData.IsInteractive)
						{
							if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue -Verbose:$False) { ShowInteractiveNotification -Title $entry.Title -Message $entry.Message -Buttons $entry.ExtraData.Buttons -Type $entry.Type -Key $entry.Key -TimeoutSeconds $entry.TimeoutSeconds -Progress $entry.Progress -IgnoreCancellation:$entry.IgnoreCancellation }
						}
						else
						{
							if (Get-Command ShowToast -ErrorAction SilentlyContinue -Verbose:$False) { ShowToast -Title $entry.Title -Message $entry.Message -Type $entry.Type -Key $entry.Key -TimeoutSeconds $entry.TimeoutSeconds -Progress $entry.Progress -IgnoreCancellation:$entry.IgnoreCancellation }
						}
					}
				}
			}
		}
		ShowNotification           = @{
			Click = {
				$grid = $global:DashboardConfig.UI.NotificationGrid
				if ($grid.SelectedRows.Count -gt 0)
				{
					$row = $grid.SelectedRows[0]
					if ($row.Tag)
					{
						$entry = $row.Tag
						if ($entry.ExtraData -and $entry.ExtraData.IsInteractive)
						{
							if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue -Verbose:$False) { ShowInteractiveNotification -Title $entry.Title -Message $entry.Message -Buttons $entry.ExtraData.Buttons -Type $entry.Type -Key $entry.Key -TimeoutSeconds $entry.TimeoutSeconds -Progress $entry.Progress -IgnoreCancellation:$entry.IgnoreCancellation }
						}
						else
						{
							if (Get-Command ShowToast -ErrorAction SilentlyContinue -Verbose:$False) { ShowToast -Title $entry.Title -Message $entry.Message -Type $entry.Type -Key $entry.Key -TimeoutSeconds $entry.TimeoutSeconds -Progress $entry.Progress -IgnoreCancellation:$entry.IgnoreCancellation }
						}
					}
				}
			}
		}
		HideAllNotifications       = @{
			Click = {
				if ($global:LoginNotificationStack)
				{
					$stackCopy = [System.Collections.ArrayList]@($global:LoginNotificationStack)
					foreach ($form in $stackCopy)
					{
						if ($form -and -not $form.IsDisposed) { $form.Close() }
					}
				}
			}
		}
		AddNote                    = @{
			Click = { if (Get-Command AddNote -ErrorAction SilentlyContinue -Verbose:$False) { AddNote } }
		}
		EditNote                   = @{
			Click = { if (Get-Command EditNote -ErrorAction SilentlyContinue -Verbose:$False) { EditNote } }
		}
		RemoveNote                 = @{
			Click = { if (Get-Command RemoveNote -ErrorAction SilentlyContinue -Verbose:$False) { RemoveNote } }
		}
		#endregion
	}

	$pickers = $global:DashboardConfig.UI.LoginPickers
	if ($pickers)
	{
		foreach ($key in $pickers.Keys)
		{
			$btn = $pickers[$key].Button
			$txt = $pickers[$key].Text

			$action = {
				param($s, $e)
				$targetTxt = $s.Tag
				$global:DashboardConfig.UI.SettingsForm.Visible = $false
				Show-DarkMessageBox "1. Click on your client. It must be in focus!`n2. Hover with the mouse above the target button.`n3. Wait 3 seconds until the settings page opens again.`n`nClick OK to start the timer." "Set Profile Coordinates" "Ok" "Information"
				Start-Sleep -Seconds 3
				$cursorPos = [System.Windows.Forms.Cursor]::Position
				$hWnd = [Custom.Native]::GetForegroundWindow()
				$rect = New-Object Custom.Native+RECT
				if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect))
				{
					$relX = $cursorPos.X - $rect.Left
					$relY = $cursorPos.Y - $rect.Top
					$targetTxt.Text = "$relX,$relY"
				}
				else
				{
					$targetTxt.Text = 'Error'
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
				Get-EventSubscriber -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue | Unregister-Event -Force
				Register-ObjectEvent -InputObject $element -EventName $e -Action $eventMappings[$elementName][$e] -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue
			}
		}
	}

	$global:DashboardConfig.State.UIInitialized = $true
}

function SetUIElement
{
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
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
		[string]$tooltip,
		[string]$dock,
		[string]$cursor,
		[switch]$noCustomPaint
	)

	$el = switch ($type)
	{
		'Form' { New-Object System.Windows.Forms.Form }
		'Panel' { New-Object System.Windows.Forms.Panel }
		'Button' { New-Object System.Windows.Forms.Button }
		'Label' { New-Object System.Windows.Forms.Label }
		'DataGridView' { New-Object Custom.DarkDataGridView }
		'TextBox' { New-Object System.Windows.Forms.TextBox }
		'ComboBox' { New-Object System.Windows.Forms.ComboBox }
		'Custom.DarkComboBox' { New-Object Custom.DarkComboBox }
		'CheckBox' { New-Object System.Windows.Forms.CheckBox }
		'Toggle' { New-Object Custom.Toggle }
		'ProgressBar' { New-Object System.Windows.Forms.ProgressBar }
		'TextProgressBar' { New-Object Custom.TextProgressBar }
		default { throw "Invalid element type specified: $type" }
	}

	if ($type -eq 'DataGridView')
	{
		$el.AllowUserToAddRows = $false
		$el.ReadOnly = $false
		$el.AllowUserToOrderColumns = $true
		$el.AllowUserToResizeColumns = $false
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
		$el.ShowCellToolTips = $true
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
		if ($PSBoundParameters.ContainsKey('width')) { $el.Width = $width }
		if ($PSBoundParameters.ContainsKey('height')) { $el.Height = $height }
		if ($PSBoundParameters.ContainsKey('top')) { $el.Top = $top }
		if ($PSBoundParameters.ContainsKey('left')) { $el.Left = $left }
		if ($PSBoundParameters.ContainsKey('dock'))
		{
			try { $el.Dock = [System.Windows.Forms.DockStyle]::$dock }
			catch { Write-Verbose "Failed to set Dock property to '$dock' on element of type '$type'" }
		}
		if ($PSBoundParameters.ContainsKey('cursor'))
		{
			try { $el.Cursor = [System.Windows.Forms.Cursors]::$cursor }
			catch { Write-Verbose "Failed to set Cursor property to '$cursor' on element of type '$type'" }
		}

		if ($bg -is [array] -and $bg.Count -ge 3)
		{
			$el.BackColor = if ($bg.Count -eq 4) { [System.Drawing.Color]::FromArgb($bg[0], $bg[1], $bg[2], $bg[3]) }
			else { [System.Drawing.Color]::FromArgb($bg[0], $bg[1], $bg[2]) }
		}

		if ($fg -is [array] -and $fg.Count -ge 3)
		{
			$el.ForeColor = [System.Drawing.Color]::FromArgb($fg[0], $fg[1], $fg[2])
		}

		if ($PSBoundParameters.ContainsKey('font')) { $el.Font = $font }
	}

	switch ($type)
	{
		'Form'
		{
			if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
			if ($PSBoundParameters.ContainsKey('startPosition')) { try { $el.StartPosition = [System.Windows.Forms.FormStartPosition]::$startPosition } catch {} }
			if ($PSBoundParameters.ContainsKey('formBorderStyle')) { try { $el.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]$formBorderStyle } catch {} }
			if ($PSBoundParameters.ContainsKey('opacity')) { $el.Opacity = [double]$opacity }
			if ($PSBoundParameters.ContainsKey('topMost')) { $el.TopMost = $topMost }
			if ($PSBoundParameters.ContainsKey('icon')) { $el.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSCommandPath) }
		}
		'Button'
		{
			if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
			if ($PSBoundParameters.ContainsKey('fs'))
			{
				$el.FlatStyle = $fs
				$el.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
				$el.FlatAppearance.BorderSize = 1
				$el.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
				$el.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)

				if (-not $noCustomPaint) {
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
		}
		'Label'
		{
			if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
		}
		'TextBox'
		{
			if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
			if ($PSBoundParameters.ContainsKey('multiline')) { $el.Multiline = $multiline }
			if ($PSBoundParameters.ContainsKey('readOnly')) { $el.ReadOnly = $readOnly }
			if ($PSBoundParameters.ContainsKey('scrollBars')) { $el.ScrollBars = if ($scrollBars -and $multiline) { [System.Windows.Forms.ScrollBars]::Vertical } else { [System.Windows.Forms.ScrollBars]::None } }
			$el.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
			$el.TextAlign = 'Center'
			$el.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
			$el.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
		}
		'ComboBox'
		{
			if ($null -ne $dropDownStyle) { try { $el.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::$dropDownStyle } catch {} }
			if ($null -ne $fs)
			{
				$el.FlatStyle = $fs
			}
		}
		'Custom.DarkComboBox'
		{
			if ($null -ne $dropDownStyle) { try { $el.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::$dropDownStyle } catch {} }
			if ($null -ne $fs)
			{
				$el.FlatStyle = $fs
			}
			$el.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
			$el.IntegralHeight = $false
			$el.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
			$el.ForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
			if ($PSBoundParameters.ContainsKey('font')) { $el.Font = $font }
		}
		'CheckBox'
		{
			if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
			if ($PSBoundParameters.ContainsKey('checked')) { $el.Checked = $checked }
			$el.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
			$el.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60, 60, 60); $el.FlatAppearance.BorderSize = 1
			$el.FlatAppearance.CheckedBackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
			$el.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
			$el.UseVisualStyleBackColor = $false; $el.CheckAlign = [System.Drawing.ContentAlignment]::MiddleLeft
			$el.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft; $el.Padding = [System.Windows.Forms.Padding]::new(20, 0, 0, 0)
		}
		'Toggle'
		{
			if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
			if ($PSBoundParameters.ContainsKey('checked')) { $el.Checked = $checked }
		}
		'ProgressBar'
		{
			if ($PSBoundParameters.ContainsKey('style')) { $el.Style = $style }
		}
		'TextProgressBar'
		{
			if ($PSBoundParameters.ContainsKey('style')) { $el.Style = $style }
		}
	}


	$targetToolTip = $null
	if ($script:CurrentBuilderToolTip) { 
		$targetToolTip = $script:CurrentBuilderToolTip 
	} elseif ($global:DashboardConfig.UI.ToolTipFtool) {
		$targetToolTip = $global:DashboardConfig.UI.ToolTipFtool
	} elseif ($global:DashboardConfig.UI.ToolTip) { 
		$targetToolTip = $global:DashboardConfig.UI.ToolTip 
	} elseif ($global:DashboardConfig.UI.ToolTipMain) {
		$targetToolTip = $global:DashboardConfig.UI.ToolTipMain
	}

	if ($PSBoundParameters.ContainsKey('tooltip') -and $tooltip -ne $null -and $targetToolTip)
	{
		$targetToolTip.SetToolTip($el, $tooltip)
	}

	if ($PSBoundParameters.ContainsKey('id') -and -not [string]::IsNullOrEmpty($id)) {
		$el.Name = $id
	}

	return $el
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion