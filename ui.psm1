<# ui.psm1 #>

#region Helper Functions

$script:HotkeysListRefreshPending = $false
$script:HotkeysListRefreshPosted = $false
$script:NotificationGridRefreshPending = $false
$script:HotkeysGridEventDepth = 0
$script:HotkeysGridMouseActive           = $false
$script:NotificationGridMouseActive      = $false
$script:HotkeysGridSafetyTimer           = $null
$script:NotificationGridEventDepth = 0

function Get-UIFont
{
	param(
		[string]$Family = 'Segoe UI',
		[double]$Size = 9,
		[string]$Style = 'Regular'
	)

	$fontStyle = [System.Drawing.FontStyle]::Regular
	try { $fontStyle = [System.Drawing.FontStyle]$Style } catch { }

	$key = "$Family|$Size|$([int]$fontStyle)"
	$cache = $script:UIFontCache
	if ($null -eq $cache) { $cache = @{}; $script:UIFontCache = $cache }
	if ($cache.ContainsKey($key)) { return $cache[$key] }

	$font = New-Object System.Drawing.Font($Family, $Size, $fontStyle)

	try { [System.GC]::SuppressFinalize($font) } catch { }

	$cache[$key] = $font
	return $font
}

function Get-BoldUIFont
{
	param([System.Drawing.Font]$Font)

	if ($null -eq $Font) { return (Get-UIFont 'Segoe UI' 9 'Bold') }

	$key = "bold|$($Font.Name)|$($Font.Size)|$($Font.Unit)"
	$cache = $script:BoldFontCache
	if ($null -eq $cache) { $cache = @{}; $script:BoldFontCache = $cache }
	if ($cache.ContainsKey($key)) { return $cache[$key] }

	$bold = $null
	try { $bold = New-Object System.Drawing.Font($Font, [System.Drawing.FontStyle]::Bold) }
	catch { return (Get-UIFont 'Segoe UI' 9 'Bold') }
	try { [System.GC]::SuppressFinalize($bold) } catch { }

	$cache[$key] = $bold
	return $bold
}
function Show-DarkMessageBox
{
	[CmdletBinding(DefaultParameterSetName = 'ImplicitOwner')]
	param(
		[Parameter(ParameterSetName = 'ExplicitOwner', Mandatory = $true)]
		[System.Windows.Forms.IWin32Window]$Owner,

		[Parameter(ParameterSetName = 'ImplicitOwner', Position = 0, Mandatory = $true)]
		[Parameter(ParameterSetName = 'ExplicitOwner', Position = 0, Mandatory = $true)]
		[string]$Text,

		[Parameter(ParameterSetName = 'ImplicitOwner', Position = 1)]
		[Parameter(ParameterSetName = 'ExplicitOwner', Position = 1)]
		[string]$Caption = 'Notification',

		[Parameter(ParameterSetName = 'ImplicitOwner', Position = 2)]
		[Parameter(ParameterSetName = 'ExplicitOwner', Position = 2)]
		[System.Windows.Forms.MessageBoxButtons]$Buttons = 'OK',

		[Parameter(ParameterSetName = 'ImplicitOwner', Position = 3)]
		[Parameter(ParameterSetName = 'ExplicitOwner', Position = 3)]
		[System.Windows.Forms.MessageBoxIcon]$Icon = 'Information',

		[Parameter(ParameterSetName = 'ImplicitOwner', Position = 4)]
		[Parameter(ParameterSetName = 'ExplicitOwner', Position = 4)]
		[string]$Type,

		[Parameter(ParameterSetName = 'ImplicitOwner', Position = 5)]
		[Parameter(ParameterSetName = 'ExplicitOwner', Position = 5)]
		[switch]$TopMost
	)

    
	$global:DashboardConfig.State.LoginActive = $true
	try
	{
		$isSuccess = $false
		if ($PSBoundParameters.ContainsKey('Type'))
		{
			$isSuccess = ($Type.ToLower() -eq 'success')
		}

		$ownerForm = $null
		if ($PSCmdlet.ParameterSetName -eq 'ExplicitOwner')
		{
			$ownerForm = $PSBoundParameters['Owner']
		}
		else
		{
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
		if (-not ('Custom.Native' -as [Type]))
  {
			throw 'Custom.Native class not found. Cannot use P/Invoke message loop.'
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

function Register-ButtonFadeEffect
{
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
		NormalStartColor       = $NormalStartColor
		NormalEndColor         = $NormalEndColor
		HoverStartColor        = $HoverStartColor
		HoverEndColor          = $HoverEndColor
		LeftPressedStartColor  = $LeftPressedStartColor
		LeftPressedEndColor    = $LeftPressedEndColor
		RightPressedStartColor = $RightPressedStartColor
		RightPressedEndColor   = $RightPressedEndColor
		TargetStartColor       = $NormalStartColor
		TargetEndColor         = $NormalEndColor
		CurrentStartR          = [double]$NormalStartColor.R
		CurrentStartG          = [double]$NormalStartColor.G
		CurrentStartB          = [double]$NormalStartColor.B
		CurrentEndR            = [double]$NormalEndColor.R
		CurrentEndG            = [double]$NormalEndColor.G
		CurrentEndB            = [double]$NormalEndColor.B
	}

	$Button | Add-Member -MemberType NoteProperty -Name 'FadeState' -Value $state -Force

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

			if ([Math]::Abs($diffStartR) -lt 1 -and [Math]::Abs($diffStartG) -lt 1 -and [Math]::Abs($diffStartB) -lt 1)
			{
				$state.CurrentStartR = $targetStart.R
				$state.CurrentStartG = $targetStart.G
				$state.CurrentStartB = $targetStart.B
			}
			else
			{
				$state.CurrentStartR += $diffStartR * $step
				$state.CurrentStartG += $diffStartG * $step
				$state.CurrentStartB += $diffStartB * $step
				$close = $false
			}

			if ([Math]::Abs($diffEndR) -lt 1 -and [Math]::Abs($diffEndG) -lt 1 -and [Math]::Abs($diffEndB) -lt 1)
			{
				$state.CurrentEndR = $targetEnd.R
				$state.CurrentEndG = $targetEnd.G
				$state.CurrentEndB = $targetEnd.B
			}
			else
			{
				$state.CurrentEndR += $diffEndR * $step
				$state.CurrentEndG += $diffEndG * $step
				$state.CurrentEndB += $diffEndB * $step
				$close = $false
			}

			$Button.Invalidate()
			if ($close)
			{
				$timer.Stop()
			}
		}.GetNewClosure())

	$Button.Add_MouseEnter({
			$state.TargetStartColor = $state.HoverStartColor
			$state.TargetEndColor = $state.HoverEndColor
			$timer.Start()
		}.GetNewClosure())

	$Button.Add_MouseLeave({
			$state.TargetStartColor = $state.NormalStartColor
			$state.TargetEndColor = $state.NormalEndColor
			$timer.Start()
		}.GetNewClosure())

	$Button.Add_Paint({
			param($src, $e)
			$rect = [System.Drawing.Rectangle]::new(0, 0, $src.Width, $src.Height)
			$color1 = [System.Drawing.Color]::FromArgb([int]$state.CurrentStartR, [int]$state.CurrentStartG, [int]$state.CurrentStartB)
			$color2 = [System.Drawing.Color]::FromArgb([int]$state.CurrentEndR, [int]$state.CurrentEndG, [int]$state.CurrentEndB)
        
			$bgBrush = New-ColorGradient -Rectangle $rect -StartColor $color1 -EndColor $color2 -Mode Vertical
			$e.Graphics.FillRectangle($bgBrush, $rect)
			$bgBrush.Dispose()

			$textBrush = [System.Drawing.SolidBrush]::new($src.ForeColor)
			$textFormat = [System.Drawing.StringFormat]::new()
			$textFormat.Alignment = [System.Drawing.StringAlignment]::Center
			$textFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
			$img = $src.Image
			if ($null -ne $img) {
				$imgH = [int][Math]::Min([int]$img.Height, [Math]::Max(1, ($src.Height - 4)))
				$imgW = [int][Math]::Round([double]$img.Width * ($imgH / [Math]::Max(1.0, [double]$img.Height)))
				if ($imgW -lt 1) { $imgW = 1 }
				if ($imgH -lt 1) { $imgH = 1 }
				$imgX = [int](($src.Width - $imgW) / 2)
				$imgY = [int](($src.Height - $imgH) / 2)
				if (-not [string]::IsNullOrEmpty($src.Text)) {
					$imgX = 7
					$e.Graphics.DrawImage($img, $imgX, $imgY, $imgW, $imgH)
					$tx = [float]($imgX + $imgW + 2)
					$tw = [float]([Math]::Max(4, ($src.Width - $imgX - $imgW - 4)))
					$e.Graphics.DrawString($src.Text, $src.Font, $textBrush, [System.Drawing.RectangleF]::new($tx, [float]0, $tw, [float]$src.Height), $textFormat)
				} else {
					$e.Graphics.DrawImage($img, $imgX, $imgY, $imgW, $imgH)
				}
			} else {
				$e.Graphics.DrawString($src.Text, $src.Font, $textBrush, [System.Drawing.RectangleF]::new(0, 0, $src.Width, $src.Height), $textFormat)
			}
			$textBrush.Dispose()
			$textFormat.Dispose()

			$borderColor = if (Test-DashboardCommand 'Get-Theme') { Get-Theme 'BorderColor' } else { Get-Theme 'ButtonBg' }
			$borderPen = [System.Drawing.Pen]::new($borderColor)
			$e.Graphics.DrawRectangle($borderPen, 0, 0, $src.Width - 1, $src.Height - 1)
			$borderPen.Dispose()
		}.GetNewClosure())
}

function Show-FadeForm
{
	[CmdletBinding()]
	param(
		[System.Windows.Forms.Form]$Form
	)

	if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
		($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
	{
		return
	}

	if (-not ($Form -and $global:DashboardConfig.UI -and $global:DashboardConfig.UI.MainForm))
	{
		return
	}

	if ($Form.Opacity -lt 0.95)
	{
		$Form.Visible = $true

		if ($Form.StartPosition -ne [System.Windows.Forms.FormStartPosition]::Manual)
		{
			$mainFormLocation = $global:DashboardConfig.UI.MainForm.Location
			$formWidth = $Form.Width
			$formHeight = $Form.Height
			$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
			$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

			$x = $mainFormLocation.X + (($global:DashboardConfig.UI.MainForm.Width - $formWidth) / 2)
			$y = $mainFormLocation.Y + (($global:DashboardConfig.UI.MainForm.Height - $formHeight) / 2)

			$margin = 0
			$x = [Math]::Max($margin, [Math]::Min($x, $screenWidth - $formWidth - $margin))
			$y = [Math]::Max($margin, [Math]::Min($y, $screenHeight - $formHeight - $margin))

			$Form.Location = New-Object System.Drawing.Point($x, $y)
			$Form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
		}

		$Form.BringToFront()
		$Form.Activate()
	}

	if ($script:fadeInTimer) { $script:fadeInTimer.Dispose() }
	$script:fadeInTimer = New-Object System.Windows.Forms.Timer
	$script:fadeInTimer.Interval = 15
	$script:fadeInTimer.Add_Tick({
			$f = $script:fadeInTarget
			if (-not $f -or $f.IsDisposed)
			{
				$script:fadeInTimer.Stop()
				$script:fadeInTimer.Dispose()
				$script:fadeInTimer = $null
				$script:fadeInTarget = $null
				return
			}
			if ($f.Opacity -lt 1)
			{
				$f.Opacity += 0.1
			}
			else
			{
				$f.Opacity = 1
				$script:fadeInTimer.Stop()
				$script:fadeInTimer.Dispose()
				$script:fadeInTimer = $null
				$script:fadeInTarget = $null
			}
		})
	$script:fadeInTarget = $Form
	$script:fadeInTimer.Start()
	$global:DashboardConfig.Resources.Timers['fadeInTimer'] = $script:fadeInTimer
}

function Hide-FadeForm
{
	[CmdletBinding()]
	param(
		[System.Windows.Forms.Form]$Form,
		[string]$PositionKey = 'SettingsForm'
	)

	if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
		($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
	{
		return
	}

	if (-not ($Form -and $global:DashboardConfig.UI))
	{
		return
	}

	if ($global:fadeOutTimer) { $global:fadeOutTimer.Dispose() }
	$global:fadeOutTimer = New-Object System.Windows.Forms.Timer
	$global:fadeOutTimer.Interval = 15
	$global:fadeOutTimer.Add_Tick({
			$f = $script:fadeOutTarget
			if (-not $f -or $f.IsDisposed)
			{
				$global:fadeOutTimer.Stop()
				$global:fadeOutTimer.Dispose()
				$global:fadeOutTimer = $null
				$script:fadeOutTarget = $null
				return
			}

			if ($f.Opacity -gt 0)
			{
				$f.Opacity -= 0.1
			}
			else
			{
				$f.Opacity = 0
				$global:fadeOutTimer.Stop()
				$global:fadeOutTimer.Dispose()
				$global:fadeOutTimer = $null
				$script:fadeOutTarget = $null
				
				if ($f.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal)
				{
					if (-not $global:DashboardConfig.Config.Contains('WindowPosition')) { $global:DashboardConfig.Config['WindowPosition'] = [ordered]@{} }
					$global:DashboardConfig.Config['WindowPosition'][$script:fadeOutPositionKey + 'X'] = $f.Location.X
					$global:DashboardConfig.Config['WindowPosition'][$script:fadeOutPositionKey + 'Y'] = $f.Location.Y
					if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
				}

				$f.Hide()
				$global:DashboardConfig.UI.MainForm.Show()
			}
		})
	$script:fadeOutTarget = $Form
	$script:fadeOutPositionKey = $PositionKey
	$global:fadeOutTimer.Start()
	$global:DashboardConfig.Resources.Timers['fadeOutTimer'] = $global:fadeOutTimer
}

function ShowSettingsForm
{

	[CmdletBinding()]
	param()

	Show-FadeForm -Form $global:DashboardConfig.UI.SettingsForm
}

function HideSettingsForm
{

	[CmdletBinding()]
	param()

	Hide-FadeForm -Form $global:DashboardConfig.UI.SettingsForm -PositionKey 'SettingsForm'
}

function ShowExtraForm
{
	[CmdletBinding()]
	param()

	Show-FadeForm -Form $global:DashboardConfig.UI.ExtraForm
}

function HideExtraForm
{

	[CmdletBinding()]
	param()

	if ($global:DashboardConfig.State['ChestHotkeyId'])
	{
		try
		{
			if (Test-DashboardCommand 'Stop-ChestOpener') { Stop-ChestOpener }
		}
		catch {}
		try
		{
			if (Test-DashboardCommand 'UnregisterHotkeyInstance')
			{
				UnregisterHotkeyInstance -Id $global:DashboardConfig.State['ChestHotkeyId'] -OwnerKey 'ChestOpener'
			}
		}
		catch {}
		$global:DashboardConfig.State['ChestHotkeyId'] = $null
	}

	Hide-FadeForm -Form $global:DashboardConfig.UI.ExtraForm -PositionKey 'ExtraForm'
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
	
	if (-not (Test-DashboardCommand 'SetHotkey'))
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

		$masterPauseKey = if ($globalHks.Contains('MasterPauseHotkey')) { $globalHks['MasterPauseHotkey']  } else { $null }
		$masterResumeKey = if ($globalHks.Contains('MasterResumeHotkey')) { $globalHks['MasterResumeHotkey'] } else { $null }

		if (-not [string]::IsNullOrWhiteSpace($masterPauseKey) -and $masterPauseKey -ne 'none')
		{
			try
			{
				SetHotkey -KeyCombinationString $masterPauseKey -OwnerKey 'master_pause_hotkey' -OwnerLabel 'Pause All Ftool/Macro Hotkeys' -Action { if (Test-DashboardCommand 'PauseMasterHotkeys') { PauseMasterHotkeys } }
				Write-Verbose "  HOTKEYS: Registered MASTER PAUSE hotkey '$masterPauseKey'."
			}
			catch { Write-Verbose "  HOTKEYS: Failed to register master pause hotkey: $($_.Exception.Message)" }
		}

		if (-not [string]::IsNullOrWhiteSpace($masterResumeKey) -and $masterResumeKey -ne 'none')
		{
			try
			{
				SetHotkey -KeyCombinationString $masterResumeKey -OwnerKey 'master_resume_hotkey' -OwnerLabel 'Resume All Ftool/Macro Hotkeys' -Action { if (Test-DashboardCommand 'ResumeMasterHotkeys') { ResumeMasterHotkeys } }
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
			$grid.Invoke([Action] { RefreshNotificationGrid })
			return
		}

		if ($script:NotificationGridMouseActive) {
			if (-not $script:NotificationGridRefreshPending) {
				$script:NotificationGridRefreshPending = $true
				$rebuildDeferred = {
					$script:NotificationGridRefreshPending = $false
					RefreshNotificationGrid
				}
				try { $null = $grid.BeginInvoke([System.Action]$rebuildDeferred) } catch { }
			}
			return
		}

		if ($script:NotificationGridEventDepth -gt 0)
		{
			$rebuildDeferred = { RefreshNotificationGrid }
			try { $null = $grid.BeginInvoke([System.Action]$rebuildDeferred) } catch { }
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
					$details = "$($entry.Title): $($entry.Message -replace "`r`n", ' ')"
					$idx = $grid.Rows.Add($timeStr, $entry.Type, $details)
					if ($idx -ge 0)
					{
						$row = $grid.Rows[$idx]
						$row.Tag = $entry
						$color = switch ($entry.Type) { 'Warning' { Get-Theme 'NotificationStripWarning' } 'Error' { Get-Theme 'NotificationStripError' } default { Get-Theme 'NotificationStripInfo' } }
						$row.Cells[1].Style.ForeColor = $color
					}
				}
			}
		}
		catch { Write-Verbose "RefreshNotificationGrid Error: $_" }
		finally { $grid.ResumeLayout() }
	}
}

function Start-HotkeysGridSafetyTimer
{
	[CmdletBinding()]
	param()

	if ($script:HotkeysGridSafetyTimer) { return }

	$timer = New-Object System.Windows.Forms.Timer
	$timer.Interval = 400
	$timer.Add_Tick({
		param($s, $e)
		try { $s.Stop(); $s.Dispose() } catch { }
		$script:HotkeysGridSafetyTimer = $null
		if ($script:HotkeysGridEventDepth -eq 0 -and -not $script:HotkeysGridMouseActive) {
			if ($script:HotkeysListRefreshPending) { RefreshHotkeysList }
		}
	})
	$script:HotkeysGridSafetyTimer = $timer
	try { $timer.Start() } catch { }
}

function RefreshHotkeysList
{
	[CmdletBinding()]
	param()

	$grid = $global:DashboardConfig.UI.HotkeysGrid
	if (-not $grid) { return }

	if ($script:HotkeysGridMouseActive) {
		$script:HotkeysListRefreshPending = $true
		return
	}

	$script:HotkeysListRefreshPending = $true

	if ($script:HotkeysGridEventDepth -gt 0) { return }
	if ($script:HotkeysListRefreshPosted) { return }

	if (-not $grid.IsHandleCreated)
	{
		$script:HotkeysListRefreshPending = $false
		try { Update-HotkeysListContent } catch { }
		return
	}

	$script:HotkeysListRefreshPosted = $true
	$rebuild = {
		$script:HotkeysListRefreshPosted = $false
		if ($script:HotkeysGridEventDepth -gt 0)
		{
			RefreshHotkeysList
			return
		}
		$script:HotkeysListRefreshPending = $false
		try { Update-HotkeysListContent } catch { }
	}
	try { $null = $grid.BeginInvoke([System.Action]$rebuild) }
	catch {
		$script:HotkeysListRefreshPosted = $false
		$script:HotkeysListRefreshPending = $false
		try { Update-HotkeysListContent } catch { }
	}
}


function Update-HotkeysListContent
{
	[CmdletBinding()]
	param()

	try
	{
		$grid = $global:DashboardConfig.UI.HotkeysGrid
		if (-not $grid) { return }
		$grid.Rows.Clear()

		if ($grid.Columns.Count -lt 3)
		{
			Write-Verbose "UI: RefreshHotkeysList aborted because HotkeysGrid has $($grid.Columns.Count) columns."
			return
		}

		$AddHotkeysGridRow = {
			param($rowValues)
			if (-not $rowValues -or $rowValues.Count -gt $grid.Columns.Count) { return -1 }
			$rowIndex = $grid.Rows.Add()
			if ($rowIndex -lt 0) { return -1 }
			for ($i = 0; $i -lt $rowValues.Count; $i++)
			{
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

			$titleByPid = @{}
			$g = $global:DashboardConfig.UI.DataGridMain
			if ($g)
			{
				foreach ($row in $g.Rows)
				{
					if ($null -eq $row.Cells -or $row.Cells.Count -le 2) { continue }
					$cellValue = $row.Cells[2].Value
					if ($null -eq $cellValue) { continue }
					$key = $cellValue.ToString()
					if ([string]::IsNullOrEmpty($key)) { continue }
					if ($titleByPid.ContainsKey($key)) { continue }
					$title = if ($row.Cells.Count -gt 1) { $row.Cells[1].Value } else { '' }
					$titleByPid[$key] = $title
				}
			}

			$GetDecoratedTitle = {
				param($pidStr)
				$title = ''
				if ($null -ne $pidStr -and -not [string]::IsNullOrEmpty([string]$pidStr))
				{
					if ($titleByPid.ContainsKey([string]$pidStr)) { $title = $titleByPid[[string]$pidStr] }
				}
				return $title
			}

			$groups = [ordered]@{
				'Master'        = @()
				'Ftool'         = @()
				'Macro'         = @()
				'Windows'       = @()
				'Groups'        = @()
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
							$displayOwner = 'Master Hotkey'
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
									$cmd = Get-QuickCommandById -Data $data -CmdId $cmdId
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
					if ($rowIndex -ge 0)
					{
						$grid.Rows[$rowIndex].DefaultCellStyle.BackColor = Get-Theme 'TertiaryBg'
						$grid.Rows[$rowIndex].DefaultCellStyle.ForeColor = Get-Theme 'SectionHeaderText'
						$grid.Rows[$rowIndex].DefaultCellStyle.Font = (Get-BoldUIFont $grid.Font)
						$grid.Rows[$rowIndex].Tag = @{ Header = $true }
						$grid.Rows[$rowIndex].ReadOnly = $true
						$grid.Rows[$rowIndex].DefaultCellStyle.SelectionBackColor = $grid.Rows[$rowIndex].DefaultCellStyle.BackColor
					}
					$sortedItems = $items | Sort-Object Key
					foreach ($entry in $sortedItems)
					{
						$rowIndex = & $AddHotkeysGridRow @($entry.Key, $entry.Owner, $entry.Action)
						if ($rowIndex -ge 0)
						{
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
	if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options'].Contains('ShowBossImages'))
	{
		$showImages = ([int]$global:DashboardConfig.Config['Options']['ShowBossImages']) -eq 1
	}

	$bName = $btn.Text
	$type = 'Normal'
	if ($btn.Tag -is [System.Collections.IDictionary])
	{
		$bName = $btn.Tag['BossName']
		$type = $btn.Tag['Type']
	}
 elseif ($btn.Tag -is [string])
	{
		$type = $btn.Tag
	}

	if ($showImages)
	{
		$bData = $global:DashboardConfig.Resources.BossData[$bName]
		if ($bData -and $bData.url)
		{
			$fileName = $bData.url.Split('/')[-1]
			$localPath = Join-Path $global:DashboardConfig.Paths.Bosses $fileName

			if (Test-Path $localPath)
			{
				$fileStream = $null
				$img = $null
				try
				{
					$fileStream = New-Object System.IO.FileStream($localPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
					$originalImage = [System.Drawing.Image]::FromStream($fileStream)
					$img = $originalImage.Clone()
					$originalImage.Dispose()
				}
				catch
				{
					Write-Verbose "Error drawing boss image for $bName : $_"
					$btn.BackgroundImage = $null
					$btn.Text = $bName
				}
				finally
				{
					if ($fileStream) { $fileStream.Dispose() }
				}

				if ($img)
				{
					$btn.BackgroundImage = $img
					$btn.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
					$btn.BackColor = Get-Theme 'Transparent'
					$btn.Text = ''
				}
			}
			else
			{
				$btn.BackgroundImage = $null
				$btn.Text = $bName

				if (-not $global:DashboardConfig.Resources.ActiveDownloads) { $global:DashboardConfig.Resources.ActiveDownloads = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new()) }
				$tempPath = "$localPath.$([Guid]::NewGuid()).tmp"

				try
				{
					[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
					$wc = New-Object System.Net.WebClient
					$wc.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)')
					$wc.Add_DownloadFileCompleted({
							param($s, $e)
							try
							{
								if ($global:DashboardConfig.Resources.ActiveDownloads.Contains($s)) { $global:DashboardConfig.Resources.ActiveDownloads.Remove($s) }
								if (-not $e.Error -and -not $e.Cancelled)
								{
									try
									{
										if (Test-Path $localPath) { Remove-Item -LiteralPath $localPath -Force }
										Move-Item -LiteralPath $tempPath -Destination $localPath -Force
									}
									catch
         {
										Write-Verbose "Error moving downloaded file '$tempPath' to '$localPath': $_"
										return
									}

									if ($btn -and -not $btn.IsDisposed)
									{
										$invoker = if ($btn.IsHandleCreated) { $btn } elseif ($global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.IsHandleCreated) { $global:DashboardConfig.UI.MainForm } else { $null }
										if ($invoker)
										{
											try { $invoker.Invoke([Action] { Update-BossButtonImage $btn }) } catch {}
										}
									}
								}
							}
							finally
       {
								if (Test-Path $tempPath) { try { Remove-Item -LiteralPath $tempPath -Force } catch {} }
								$s.Dispose()
							}
						}.GetNewClosure())
					$global:DashboardConfig.Resources.ActiveDownloads.Add($wc) | Out-Null
					$wc.DownloadFileAsync([Uri]$bData.url, $tempPath)
				}
				catch {}
			}
		}
	}
	else
	{
		$btn.BackgroundImage = $null
		$btn.Text = $bName
		$btn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
		$btn.Font = (Get-UIFont 'Segoe UI Semibold' 9)
		$btn.ForeColor = Get-Theme 'WhiteText'
		$btn.BackColor = Get-Theme 'BossButtonBg'
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
function Get-QuickCommandById
{
	[CmdletBinding()]
	param(
		$Data,
		[string]$CmdId
	)

	if (-not $Data) { return $null }

	$cmd = $null
	if ($Data.SavedCommands)
	{
		foreach ($c in $Data.SavedCommands)
		{
			if ($c -and $c.Id -eq $CmdId) { $cmd = $c; break }
		}
	}
	if ((-not $cmd) -and $Data.GlobalCommandList)
	{
		foreach ($c in $Data.GlobalCommandList)
		{
			if ($c -and $c.Id -eq $CmdId) { $cmd = $c; break }
		}
	}
	return $cmd
}

function Set-HotkeyGroupBinding
{
	[CmdletBinding()]
	param(
		[string]$KeySection = 'Hotkeys',
		[string]$GroupSection = 'HotkeyGroups',
		[string]$DialogTitle = 'Hotkey Group',
		[string]$MessageTitle = 'Hotkeys Created'
	)

	$grid = $global:DashboardConfig.UI.DataGridMain
	if (-not $grid) { return }
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

	$groupName = ShowInputBox -Title $DialogTitle -Prompt 'Enter a name for this selection group:' -DefaultText $defaultName
	if ([string]::IsNullOrWhiteSpace($groupName)) { return }

	$currentKey = 'none'
	if ($global:DashboardConfig.Config.Contains($KeySection))
	{
		if ($global:DashboardConfig.Config[$KeySection].Contains($groupName))
		{
			$currentKey = $global:DashboardConfig.Config[$KeySection][$groupName]
		}
	}

	$newKey = Show-KeyCaptureDialog -currentKey $currentKey -Owner $global:DashboardConfig.UI.MainForm
	
	if ($null -eq $newKey -or ($newKey -eq $currentKey -and $currentKey -ne 'none')) { return }

	if (-not $global:DashboardConfig.Config.Contains($KeySection))
	{
		$global:DashboardConfig.Config[$KeySection] = @{}
	}
	if (-not $global:DashboardConfig.Config.Contains($GroupSection))
	{
		$global:DashboardConfig.Config[$GroupSection] = @{}
	}

	$global:DashboardConfig.Config[$KeySection][$groupName] = $newKey
	$global:DashboardConfig.Config[$GroupSection][$groupName] = ($identities -join ',')

	if (Test-DashboardCommand 'WriteConfig') { WriteConfig }

	SyncConfigToUI
	RegisterConfiguredHotkeys
	RefreshHotkeysList
	
	Show-DarkMessageBox "Hotkey '$newKey' assigned as '$groupName' to:
$($identities -join "
")" $MessageTitle 'Ok' 'Information' 'success'
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
				if ($Matches[1] -eq '1') { return 'DirectX 11' }
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
		$apiValue = if ($DirectXString -match '11') { '1' } else { '0' }
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

#region Login coordinate defaults (auto populated per resolution)

function GetLoginCoordinateLayout
{
	[CmdletBinding()]
	[OutputType([System.Collections.Specialized.OrderedDictionary])]
	param()

	return [ordered]@{
		'LauncherPlay'   = @{ Mode = 'Fixed';     X = 860;  Y = 520 }
		'CollectorStart' = @{ Mode = 'RightTop';  X = 70;   Y = 305 }
		'DisconnectOK'   = @{ Mode = 'CentreTop'; X = 13;   Y = 215 }
		'Server1'        = @{ Mode = 'Centre';    X = -140; Y = -74 }
		'Server2'        = @{ Mode = 'Centre';    X = -140; Y = -54 }
		'Channel1'       = @{ Mode = 'Centre';    X = 60;   Y = -74 }
		'Channel2'       = @{ Mode = 'Centre';    X = 60;   Y = -54 }
		'Char1'          = @{ Mode = 'Centre';    X = -162; Y = 36 }
		'Char2'          = @{ Mode = 'Centre';    X = 10;   Y = 36 }
		'Char3'          = @{ Mode = 'Centre';    X = 180;  Y = 36 }
		'Char4'          = @{ Mode = 'Centre';    X = -162; Y = 36 }
		'Char5'          = @{ Mode = 'Centre';    X = 10;   Y = 36 }
		'Char6'          = @{ Mode = 'Centre';    X = 180;  Y = 36 }
		'FirstNick'      = @{ Mode = 'Centre';    X = 38;   Y = 6 }
		'ScrollDown'     = @{ Mode = 'Centre';    X = 183;  Y = 81 }
		'LoginDetailsOK' = @{ Mode = 'Centre';    X = 3;    Y = 76 }
		'SwitchLeft'     = @{ Mode = 'Centre';    X = -262; Y = 26 }
		'SwitchRight'    = @{ Mode = 'Centre';    X = 275;  Y = 26 }
	}
}

function GetLoginCoordinateDefaults
{
	[CmdletBinding()]
	[OutputType([System.Collections.IDictionary])]
	param(
		[Parameter()]
		[string]$ResolutionString
	)

	$width = 0
	$height = 0
	if ($ResolutionString -match '(\d+)\s*x\s*(\d+)')
	{
		$width = [int]$Matches[1]
		$height = [int]$Matches[2]
	}

	if ($width -le 0 -or $height -le 0)
	{
		try
		{
			$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
			$width = $bounds.Width
			$height = $bounds.Height
		}
		catch
		{
			$width = 1024
			$height = 768
		}
	}

	$centreX = [int][math]::Round($width / 2)
	$centreY = [int][math]::Round($height / 2)

	$layout = GetLoginCoordinateLayout
	$defaults = [ordered]@{}

	foreach ($key in $layout.Keys)
	{
		$rule = $layout[$key]
		$x = 0
		$y = 0
		switch ($rule['Mode'])
		{
			'Fixed'     { $x = [int]$rule['X']; $y = [int]$rule['Y'] }
			'RightTop'  { $x = $width - [int]$rule['X']; $y = [int]$rule['Y'] }
			'CentreTop' { $x = $centreX + [int]$rule['X']; $y = [int]$rule['Y'] }
			default     { $x = $centreX + [int]$rule['X']; $y = $centreY + [int]$rule['Y'] }
		}
		$defaults[$key] = "$x,$y"
	}

	return $defaults
}

function GetLoginCoordinateResolutionString
{
	[CmdletBinding()]
	[OutputType([string])]
	param(
		[Parameter()]
		[string]$ProfileName
	)

	try
	{
		$UI = $global:DashboardConfig.UI
		if ($UI -and $UI.ResolutionSelector -and $UI.ResolutionSelector.SelectedItem)
		{
			$selectedProfile = 'Default'
			if ($UI.LoginProfileSelector -and $UI.LoginProfileSelector.SelectedItem) { $selectedProfile = $UI.LoginProfileSelector.SelectedItem.ToString() }
			if ($selectedProfile -eq $ProfileName) { return [string]$UI.ResolutionSelector.SelectedItem }
		}
	}
	catch {}

	$profPath = $null
	try
	{
		if ([string]::IsNullOrWhiteSpace($ProfileName) -or $ProfileName -eq 'Default')
		{
			if ($global:DashboardConfig.Config['LauncherPath'] -and $global:DashboardConfig.Config['LauncherPath']['LauncherPath'])
			{
				$profPath = Split-Path -Parent $global:DashboardConfig.Config['LauncherPath']['LauncherPath']
			}
		}
		elseif ($global:DashboardConfig.Config['Profiles'] -and $global:DashboardConfig.Config['Profiles'].Contains($ProfileName))
		{
			$profPath = $global:DashboardConfig.Config['Profiles'][$ProfileName]
		}
	}
	catch {}

	if ($profPath)
	{
		$resolution = GetNeuzResolution -ProfilePath $profPath
		if ($resolution) { return $resolution }
	}

	return ''
}

function RepairLoginConfigSection
{
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[System.Collections.IDictionary]$LoginConfig
	)

	$flatKeys = @()
	foreach ($key in @($LoginConfig.Keys))
	{
		if ($LoginConfig[$key] -is [string] -and $key -match '^([^_]+)_(.+)$') { $flatKeys += $key }
	}

	foreach ($flatKey in $flatKeys)
	{
		$flatKey -match '^([^_]+)_(.+)$' | Out-Null
		$profileName = $Matches[1]
		$settingName = $Matches[2]

		if (-not $LoginConfig.Contains($profileName) -or $LoginConfig[$profileName] -isnot [System.Collections.IDictionary])
		{
			$LoginConfig[$profileName] = [ordered]@{}
		}
		if (-not $LoginConfig[$profileName].Contains($settingName))
		{
			$LoginConfig[$profileName][$settingName] = $LoginConfig[$flatKey]
		}
		$LoginConfig.Remove($flatKey)
	}
}

function GetLoginCoordsOverrideState
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter()]
		[System.Collections.IDictionary]$LoginConfig
	)

	if ($null -eq $LoginConfig) { return $false }
	if ($LoginConfig.Contains('ManualCoordsOverride')) { return ($LoginConfig['ManualCoordsOverride'] -eq 'Yes') }

	foreach ($key in (GetLoginCoordinateLayout).Keys)
	{
		if ($key -eq 'LauncherPlay') { continue }
		$value = $LoginConfig["${key}Coords"]
		if ($value -and $value -ne '0,0' -and $value -ne 'Error') { return $true }
	}
	return $false
}

function SyncLoginCoordinateDefaults
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter()]
		[string]$ProfileName
	)

	if (-not $global:DashboardConfig -or -not $global:DashboardConfig.Config) { return $false }
	if (-not $global:DashboardConfig.Config.Contains('LoginConfig')) { $global:DashboardConfig.Config['LoginConfig'] = [ordered]@{} }

	$loginConfig = $global:DashboardConfig.Config['LoginConfig']
	RepairLoginConfigSection -LoginConfig $loginConfig

	$profiles = @('Default')
	if ($global:DashboardConfig.Config['Profiles'])
	{
		foreach ($name in $global:DashboardConfig.Config['Profiles'].Keys)
		{
			if ($name -notin $profiles) { $profiles += $name }
		}
	}
	if ($ProfileName -and $ProfileName -notin $profiles) { $profiles += $ProfileName }

	foreach ($name in $profiles)
	{
		if (-not $loginConfig.Contains($name) -or $loginConfig[$name] -isnot [System.Collections.IDictionary]) { $loginConfig[$name] = [ordered]@{} }
		$lc = $loginConfig[$name]

		$manual = GetLoginCoordsOverrideState -LoginConfig $lc
		$lc['ManualCoordsOverride'] = if ($manual) { 'Yes' } else { 'No' }
		if ($manual) { continue }

		$defaults = GetLoginCoordinateDefaults -ResolutionString (GetLoginCoordinateResolutionString -ProfileName $name)
		foreach ($key in $defaults.Keys) { $lc["${key}Coords"] = $defaults[$key] }
	}

	return $true
}

function UpdateLoginCoordinateFields
{
	[CmdletBinding()]
	param()

	$UI = $global:DashboardConfig.UI
	if (-not $UI -or -not $UI.LoginPickers) { return }

	$profileName = 'Default'
	try
	{
		if ($UI.LoginProfileSelector -and $UI.LoginProfileSelector.SelectedItem) { $profileName = $UI.LoginProfileSelector.SelectedItem.ToString() }
	}
	catch {}

	SyncLoginCoordinateDefaults -ProfileName $profileName | Out-Null

	$lc = $null
	try
	{
		if ($global:DashboardConfig.Config['LoginConfig'] -and $global:DashboardConfig.Config['LoginConfig'].Contains($profileName))
		{
			$lc = $global:DashboardConfig.Config['LoginConfig'][$profileName]
		}
	}
	catch {}

	$manual = if ($UI.ManualCoordsOverride) { [bool]$UI.ManualCoordsOverride.Checked } else { GetLoginCoordsOverrideState -LoginConfig $lc }
	$defaults = GetLoginCoordinateDefaults -ResolutionString (GetLoginCoordinateResolutionString -ProfileName $profileName)

	foreach ($key in $UI.LoginPickers.Keys)
	{
		$picker = $UI.LoginPickers[$key]
		$value = ''

		if ($manual)
		{
			if ($lc) { $value = $lc["Manual_${key}Coords"] }
			if (-not $value -and $lc) { $value = $lc["${key}Coords"] }
			if (-not $value) { $value = '0,0' }
		}
		else
		{
			if ($lc -and $lc["${key}Coords"]) { $value = $lc["${key}Coords"] }
			elseif ($defaults[$key]) { $value = $defaults[$key] }
			else { $value = '0,0' }
		}

		$picker.Text.Text = $value
		$picker.Text.ReadOnly = (-not $manual)
		if ($picker.Button) { $picker.Button.Enabled = $manual }
	}
}

function StashUserLoginCoordinates
{
	[CmdletBinding()]
	param()

	$UI = $global:DashboardConfig.UI
	if (-not $UI -or -not $UI.LoginPickers) { return }

	$profileName = 'Default'
	try
	{
		if ($UI.LoginProfileSelector -and $UI.LoginProfileSelector.SelectedItem) { $profileName = $UI.LoginProfileSelector.SelectedItem.ToString() }
	}
	catch {}

	$lc = $null
	try
	{
		$loginConfig = $global:DashboardConfig.Config['LoginConfig']
		if (-not $loginConfig.Contains($profileName)) { $loginConfig[$profileName] = [ordered]@{} }
		$lc = $loginConfig[$profileName]
	}
	catch { return }

	$defaults = GetLoginCoordinateDefaults -ResolutionString (GetLoginCoordinateResolutionString -ProfileName $profileName)

	foreach ($key in $UI.LoginPickers.Keys)
	{
		$value = $UI.LoginPickers[$key].Text.Text
		if (-not $value -or $value -notmatch '^\s*-?\d+\s*,\s*-?\d+\s*$') { continue }
		if ($value -eq '0,0' -or $value -eq $defaults[$key]) { continue }
		$lc["Manual_${key}Coords"] = $value
	}
}

function SaveLoginCoordinateFields
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$ProfileName
	)

	$UI = $global:DashboardConfig.UI
	if (-not ($UI -and $UI.LoginPickers -and $global:DashboardConfig.Config)) { return $false }
	if (-not $global:DashboardConfig.Config.Contains('LoginConfig')) { $global:DashboardConfig.Config['LoginConfig'] = [ordered]@{} }

	$loginConfig = $global:DashboardConfig.Config['LoginConfig']
	if (-not $loginConfig.Contains($ProfileName) -or $loginConfig[$ProfileName] -isnot [System.Collections.IDictionary]) { $loginConfig[$ProfileName] = [ordered]@{} }
	$profileLoginConfig = $loginConfig[$ProfileName]

	$manualOverride = ($UI.ManualCoordsOverride -and $UI.ManualCoordsOverride.Checked)
	$wasManualOverride = GetLoginCoordsOverrideState -LoginConfig $profileLoginConfig
	$profileLoginConfig['ManualCoordsOverride'] = if ($manualOverride) { 'Yes' } else { 'No' }

	$coordDefaults = GetLoginCoordinateDefaults -ResolutionString (GetLoginCoordinateResolutionString -ProfileName $ProfileName)

	foreach ($key in $UI.LoginPickers.Keys)
	{
		$coordKey = "${key}Coords"
		$coordText = $UI.LoginPickers[$key].Text.Text
		$profileLoginConfig[$coordKey] = $coordText

		$coordTextValid = ($coordText -and $coordText -match '^\s*-?\d+\s*,\s*-?\d+\s*$' -and $coordText -ne '0,0')
		if ($manualOverride)
		{
			if ($coordTextValid) { $profileLoginConfig["Manual_$coordKey"] = $coordText }
		}
		elseif ($wasManualOverride -and $coordTextValid -and $coordDefaults[$key] -and $coordText -ne $coordDefaults[$key])
		{
			$profileLoginConfig["Manual_$coordKey"] = $coordText
		}
	}

	if (-not $manualOverride)
	{
		foreach ($key in $coordDefaults.Keys) { $profileLoginConfig["${key}Coords"] = $coordDefaults[$key] }
	}

	return $true
}

#endregion

function Copy-BinFoldersToProfiles
{
	[CmdletBinding()]
	[OutputType([pscustomobject])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$SourceDir,
		[Parameter(Mandatory = $true)]
		[AllowEmptyCollection()]
		[object[]]$Targets
	)

	$result = [pscustomobject]@{
		Profiles       = 0
		Copied         = 0
		Unchanged      = 0
		Skipped        = New-Object System.Collections.Generic.List[string]
		MissingFolders = New-Object System.Collections.Generic.List[string]
	}

	$sourceBins = @()
	foreach ($binName in @('bin32', 'bin64'))
	{
		$src = Join-Path $SourceDir $binName
		if (Test-Path -LiteralPath $src) { $sourceBins += $src }
	}
	if ($sourceBins.Count -eq 0) { return $result }

	foreach ($target in $Targets)
	{
		$targetName = ''
		$targetPath = ''
		if ($target -is [System.Collections.IDictionary])
		{
			if ($target.Contains('Name')) { $targetName = [string]$target['Name'] }
			if ($target.Contains('Path')) { $targetPath = [string]$target['Path'] }
		}
		else
		{
			$nameProperty = $target.PSObject.Properties['Name']
			if ($nameProperty) { $targetName = [string]$nameProperty.Value }
			$pathProperty = $target.PSObject.Properties['Path']
			if ($pathProperty) { $targetPath = [string]$pathProperty.Value }
		}
		if ([string]::IsNullOrWhiteSpace($targetName)) { $targetName = $targetPath }
		if ([string]::IsNullOrWhiteSpace($targetPath))
		{
			$result.MissingFolders.Add("$targetName : no path")
			continue
		}
		if (-not (Test-Path -LiteralPath $targetPath))
		{
			$result.MissingFolders.Add("$targetName : $targetPath not found")
			continue
		}

		foreach ($srcBin in $sourceBins)
		{
			$binLeaf = [System.IO.Path]::GetFileName($srcBin)
			$dstRoot = Join-Path $targetPath $binLeaf

			try
			{
				if (-not (Test-Path -LiteralPath $dstRoot))
				{
					New-Item -ItemType Directory -Path $dstRoot -Force -ErrorAction Stop | Out-Null
				}

				foreach ($srcDir in @(Get-ChildItem -LiteralPath $srcBin -Recurse -Directory -ErrorAction SilentlyContinue))
				{
					$rel = $srcDir.FullName.Substring($srcBin.Length).TrimStart('\').TrimStart('/')
					$dstDir = Join-Path $dstRoot $rel
					if (-not (Test-Path -LiteralPath $dstDir))
					{
						try { New-Item -ItemType Directory -Path $dstDir -Force -ErrorAction Stop | Out-Null }
						catch { $result.Skipped.Add("$dstDir : $($_.Exception.Message)") }
					}
				}
			}
			catch
			{
				$result.Skipped.Add("$dstRoot : $($_.Exception.Message)")
				continue
			}

			foreach ($srcFile in @(Get-ChildItem -LiteralPath $srcBin -Recurse -File -ErrorAction SilentlyContinue))
			{
				$rel = $srcFile.FullName.Substring($srcBin.Length).TrimStart('\').TrimStart('/')
				$dstFile = Join-Path $dstRoot $rel

				try
				{
					$dstParent = Split-Path -Path $dstFile -Parent
					if (-not (Test-Path -LiteralPath $dstParent))
					{
						New-Item -ItemType Directory -Path $dstParent -Force -ErrorAction Stop | Out-Null
					}

					if (Test-Path -LiteralPath $dstFile)
					{
						$dstItem = Get-Item -LiteralPath $dstFile -Force -ErrorAction Stop
						if ($dstItem.Length -eq $srcFile.Length -and $dstItem.LastWriteTimeUtc -eq $srcFile.LastWriteTimeUtc)
						{
							$result.Unchanged++
							continue
						}
					}

					Copy-Item -LiteralPath $srcFile.FullName -Destination $dstFile -Force -ErrorAction Stop
					$result.Copied++
				}
				catch
				{
					$result.Skipped.Add("$binLeaf\$rel : $($_.Exception.Message)")
				}
			}
		}

		$result.Profiles++
	}

	return $result
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

		$global:DashboardConfig.Config['Profiles'].Clear()

		foreach ($row in $UI.ProfileGrid.Rows)
		{
			if ($null -ne $row.Cells -and $row.Cells.Count -gt 1 -and $null -ne $row.Cells[0].Value -and $null -ne $row.Cells[1].Value)
			{
				$key = $row.Cells[0].Value.ToString()
				$val = $row.Cells[1].Value.ToString()

				$global:DashboardConfig.Config['Profiles'][$key] = $val
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

function SyncClientGridToConfig
{
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	try
	{
		$UI = $global:DashboardConfig.UI
		if (-not ($UI -and $UI.ClientGrid -and $global:DashboardConfig.Config)) { return $false }

		$section = Get-RememberedClientSection
		if (-not $section) { return $false }

		$section.Clear()
		foreach ($row in $UI.ClientGrid.Rows)
		{
			if (-not $row.Cells -or $row.Cells.Count -lt 7) { continue }

			$profileName = if ($null -ne $row.Cells[0].Value) { $row.Cells[0].Value.ToString().Trim() } else { '' }
			$clientName = if ($null -ne $row.Cells[1].Value) { $row.Cells[1].Value.ToString().Trim() } else { '' }
			if ([string]::IsNullOrWhiteSpace($profileName) -or [string]::IsNullOrWhiteSpace($clientName)) { continue }

			$flagAt = {
				param($cell)
				if ($cell.Value -eq $true -or $cell.Value -eq 1) { '1' } else { '0' }
			}
			$key = "$profileName|$clientName"
			if ($section.Contains($key)) { continue }

			$section[$key] = "$(& $flagAt $row.Cells[4]),$(& $flagAt $row.Cells[5]),$(& $flagAt $row.Cells[6]),$(& $flagAt $row.Cells[2]),$(& $flagAt $row.Cells[3])"
		}

		Write-Verbose "  UI: ClientGrid synced to remembered clients."
		return $true
	}
	catch
	{
		Write-Verbose "  UI: Failed to sync ClientGrid to remembered clients: $_"
		return $false
	}
}

function SyncConfigToClientGrid
{
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	try
	{
		$UI = $global:DashboardConfig.UI
		if (-not ($UI -and $UI.ClientGrid)) { return $false }

		if ($global:DashboardConfig.State) { $global:DashboardConfig.State.ClientGridPopulating = $true }
		try
		{
			$UI.ClientGrid.Rows.Clear()

			$section = Get-RememberedClientSection
			if ($section)
			{
				foreach ($key in @($section.Keys))
				{
					$parts = [string]$key -split '\|', 2
					if ($parts.Count -ne 2) { continue }

					$tokens = @(([string]$section[$key]) -split ',')
					$isFtool = ($tokens.Count -gt 0 -and $tokens[0].Trim() -eq '1')
					$isMacro = ($tokens.Count -gt 1 -and $tokens[1].Trim() -eq '1')
					$isQuick = ($tokens.Count -gt 2 -and $tokens[2].Trim() -eq '1')

					$windowFlags = Get-ClientWindowFlags -ProfileName $parts[0] -ClientName $parts[1]

					$UI.ClientGrid.Rows.Add($parts[0], $parts[1], $windowFlags.Reconnect, $windowFlags.Hide, $isFtool, $isMacro, $isQuick) | Out-Null
				}
			}

			Restore-ClientGridSelection | Out-Null
		}
		finally
		{
			if ($global:DashboardConfig.State) { $global:DashboardConfig.State.ClientGridPopulating = $false }
		}

		return $true
	}
	catch
	{
		if ($global:DashboardConfig.State) { $global:DashboardConfig.State.ClientGridPopulating = $false }
		Write-Verbose "  UI: Error syncing ClientGrid: $_"
		return $false
	}
}

function RemoveLegacyProfileAutoFlags
{
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	try
	{
		$config = $global:DashboardConfig.Config
		if (-not $config) { return $false }

		$removed = $false
		foreach ($legacySection in @('AutoFtoolProfiles', 'AutoMacroProfiles', 'AutoQuickProfiles', 'ReconnectProfiles'))
		{
			if ($config.Contains($legacySection))
			{
				$config.Remove($legacySection)
				$removed = $true
			}
		}

		if ($config.Contains('Options') -and $config['Options'] -and $config['Options'].Contains('ReconnectProfiles'))
		{
			$config['Options'].Remove('ReconnectProfiles')
			$removed = $true
		}
		return $removed
	}
	catch { return $false }
}

function MigrateLegacyHideProfiles
{
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	try
	{
		$config = $global:DashboardConfig.Config
		if (-not $config -or -not $config.Contains('HideProfiles')) { return $false }

		$section = $config['HideProfiles']

		$hasOption = ($config.Contains('Options') -and $config['Options'] -and $config['Options'].Contains('HideProfiles'))

		if (-not $hasOption -and $section -and @($section.Keys).Count -gt 0)
		{
			$hide = $false
			foreach ($flagKey in @($section.Keys))
			{
				if (([string]$section[$flagKey]).Trim() -eq '1') { $hide = $true; break }
			}

			if (-not $config.Contains('Options') -or $null -eq $config['Options']) { $config['Options'] = [ordered]@{} }
			$config['Options']['HideProfiles'] = if ($hide) { '1' } else { '0' }
		}

		$config.Remove('HideProfiles')
		return $true
	}
	catch { return $false }
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
		$global:DashboardConfig.Config['ProcessName']['ProcessName'] = $UI.InputProcess.Text
		$global:DashboardConfig.Config['MaxClients']['MaxClients'] = $UI.InputMax.Text
		$global:DashboardConfig.Config['Paths']['JunctionTarget'] = $UI.InputJunction.Text
		
		$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin'] = if ($UI.NeverRestartingCollectorLogin.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['WorldBossListener'] = if ($UI.WorldBossListener.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['ShowBossImages'] = if ($UI.ShowBossImages.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['AutoStartWithWindows'] = if ($UI.AutoStartWithWindows.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['AutoStartSetup'] = if ($UI.AutoStartSetup.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['HideProfiles'] = if ($UI.HideProfiles.Checked) { '1' } else { '0' }
		$global:DashboardConfig.Config['Options']['NotificationPosition'] = $UI.NotificationPositionSelector.SelectedItem

		$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
		$appName = 'EntropiaDashboard'

		$taskName = $appName

		try
		{
			if ($UI.AutoStartWithWindows.Checked)
			{
			
				if ($global:DashboardConfig.State.IsRunningAsExe)
   {
					$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
					$action = New-ScheduledTaskAction -Execute $exePath
				}
				else
				{
					$startPath = Join-Path $global:DashboardConfig.Paths.Source 'start.ps1'
					$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$startPath`""
				}

				$trigger = New-ScheduledTaskTrigger -AtLogOn

				$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest

				try
				{
					Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
					Write-Verbose 'Auto-start task registered successfully.'
				}
				catch
				{
					Write-Verbose 'Failed to register scheduled task. Ensure you are running this setup script as Administrator.'
				}

			}
			else
			{
			
				if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue)
   {
					Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
				}
			
			}
		}
		catch
		{
			Write-Verbose "  UI: Could not update the Windows auto-start task: $_"
		}

		$null = SyncProfilesToConfig

		SyncClientGridToConfig | Out-Null

		RemoveLegacyProfileAutoFlags | Out-Null
		MigrateLegacyHideProfiles | Out-Null

		$global:DashboardConfig.Config['Options']['LauncherTimeout'] = $UI.InputLauncherTimeout.Text
		if (-not $global:DashboardConfig.Config.Contains('SelectedClients') -or $null -eq $global:DashboardConfig.Config['SelectedClients'])
		{
			$global:DashboardConfig.Config['SelectedClients'] = [ordered]@{}
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

		SaveLoginCoordinateFields -ProfileName $targetProfile | Out-Null

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

		if ($global:DashboardConfig.State) { $global:DashboardConfig.State.LoadingLoginSettings = $true }

		RemoveLegacyProfileAutoFlags | Out-Null
		MigrateLegacyHideProfiles | Out-Null

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
		$UI.HideProfiles.Checked = Test-HideProfilesDefault
		if ($global:DashboardConfig.Config.Contains('Options') -and $global:DashboardConfig.Config['Options'].Contains('NotificationPosition'))
		{ 
			$pos = $global:DashboardConfig.Config['Options']['NotificationPosition']
			if ($UI.NotificationPositionSelector.Items.Contains($pos)) { $UI.NotificationPositionSelector.SelectedItem = $pos }
		}
		else { $UI.NotificationPositionSelector.SelectedIndex = 0 }

		if ($global:DashboardConfig.Config.Contains('Options') -and $global:DashboardConfig.Config['Options'].Contains('Theme'))
		{
			$savedTheme = $global:DashboardConfig.Config['Options']['Theme']
			if (Test-DashboardCommand 'Set-Theme') { Set-Theme $savedTheme }
		}


		try
		{
			$UI.ProfileGrid.Rows.Clear()

			if ($global:DashboardConfig.Config['Profiles'])
			{
				$profiles = $global:DashboardConfig.Config['Profiles']
				foreach ($key in $profiles.Keys)
				{
					$path = $profiles[$key]
					$UI.ProfileGrid.Rows.Add($key, $path) | Out-Null
				}
			}
		}
		catch { Write-Verbose "  UI: Error syncing ProfileGrid: $_" }

		try { SyncConfigToClientGrid | Out-Null }
		catch { Write-Verbose "  UI: Error syncing ClientGrid: $_" }

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

		SyncLoginCoordinateDefaults | Out-Null

		if ($global:DashboardConfig.Config['LoginConfig'] -and $global:DashboardConfig.Config['LoginConfig'][$selectedLoginProfile])
		{
			$lc = $global:DashboardConfig.Config['LoginConfig'][$selectedLoginProfile]
		}
		else
		{
			$lc = [hashtable]@{}
		}

		if ($UI.ManualCoordsOverride) { $UI.ManualCoordsOverride.Checked = (GetLoginCoordsOverrideState -LoginConfig $lc) }
		UpdateLoginCoordinateFields

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
				$UI.MasterHotkeysStatus.ForeColor = Get-Theme 'WarningColor'
			}
			else
			{
				$UI.MasterHotkeysStatus.Text = ([char]0x25CF).ToString() + ' Active'
				$UI.MasterHotkeysStatus.ForeColor = Get-Theme 'SuccessColor'
			}
		}

		if (Test-DashboardCommand 'RefreshNoteGrid') { RefreshNoteGrid }

		Write-Verbose '  UI: Config synced to UI'
		return $true
	}
	catch
	{
		Write-Verbose "  UI: Failed to sync config to UI: $_"
		return $false
	}
	finally
	{
		if ($global:DashboardConfig.State) { $global:DashboardConfig.State.LoadingLoginSettings = $false }
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
		'Entropia King'    = @{
			name = 'Entropia King'
			url  = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Entropia-King.png'
			role = '1030608207460171837'
		}
		'Clockworks'       = @{
			name = 'Clockworks'
			url  = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Clockworks.png'
			role = '1030608298266869791'
		}
		'Royal Knight'     = @{
			name = 'Royal Knight'
			url  = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Royal-Knight.png'
			role = '1030608553741930536'
		}
		'Solarion'         = @{
			name = 'Solarion'
			url  = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Solarion.png'
			role = '1418869138016960512'
		}
		'C-A01'            = @{
			name = 'C-A01'
			url  = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/C-A01.png'
			role = '1210685566010658926'
		}
		'Great Venux Tree' = @{
			name = 'Great Venux Tree'
			url  = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Great-Venux-Tree.png'
			role = '1210685569164640296'
		}
		'Grinch'           = @{
			name = 'Grinch'
			url  = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/Bosses/Grinch.png'
			role = '1444911271978729616'
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

	$p = @{ type = 'Form'; visible = $false; width = 470; height = 440; id = 'MainForm'; text = 'Entropia Dashboard'; startPosition = 'Manual'; formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None }
	$mainForm = SetUIElement @p
	$p = @{ type = 'Form'; width = 600; height = 795; id = 'SettingsForm'; text = 'Settings'; startPosition = 'CenterScreen'; formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None; topMost = $false; opacity = 0.0 }
	$settingsForm = SetUIElement @p
	$settingsForm.Owner = $mainForm
	$p = @{ type = 'Form'; width = 400; height = 600; id = 'ExtraForm'; text = 'Extra'; startPosition = 'CenterScreen'; formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None; topMost = $false; opacity = 0.0 }
	$extraForm = SetUIElement @p
	$extraForm.Owner = $mainForm

	if ($wp = $global:DashboardConfig.Config['WindowPosition'])
	{
		try
		{
			if (($h = [int]$wp.MainFormHeight) -ge 440)
			{
				$mainForm.Height = $h; Write-Verbose "UI: Applied height: $h"
			}
			$x, $y = [int]$wp.MainFormX, [int]$wp.MainFormY
			$rect = [System.Drawing.Rectangle]::new($x, $y, $mainForm.Width, $mainForm.Height)
			if ([System.Windows.Forms.Screen]::AllScreens.WorkingArea.IntersectsWith($rect) -contains $true)
			{
				$mainForm.Location = [System.Drawing.Point]::new($x, $y)
			}
			if ($wp.Contains('ExtraFormX') -and $wp.Contains('ExtraFormY'))
			{
				$ex, $ey = [int]$wp.ExtraFormX, [int]$wp.ExtraFormY
				$eRect = [System.Drawing.Rectangle]::new($ex, $ey, $extraForm.Width, $extraForm.Height)
				if ([System.Windows.Forms.Screen]::AllScreens.WorkingArea.IntersectsWith($eRect) -contains $true)
				{
					$extraForm.StartPosition = 'Manual'; $extraForm.Location = [System.Drawing.Point]::new($ex, $ey)
				}
			}
			if ($wp.Contains('SettingsFormX') -and $wp.Contains('SettingsFormY'))
			{
				$sx, $sy = [int]$wp.SettingsFormX, [int]$wp.SettingsFormY
				$sRect = [System.Drawing.Rectangle]::new($sx, $sy, $settingsForm.Width, $settingsForm.Height)
				if ([System.Windows.Forms.Screen]::AllScreens.WorkingArea.IntersectsWith($sRect) -contains $true)
				{
					$settingsForm.StartPosition = 'Manual'; $settingsForm.Location = [System.Drawing.Point]::new($sx, $sy)
				}
			}
		}
		catch { Write-Verbose "UI: Failed to apply saved window position: $_" }
	}
	else { $mainForm.StartPosition = 'CenterScreen' }

	if (-not $global:DashboardConfig.UI) { $global:DashboardConfig | Add-Member -MemberType NoteProperty -Name UI -Value ([PSCustomObject]@{}) -Force }
	if (-not $global:DashboardConfig.Resources.ActiveDownloads) { $global:DashboardConfig.Resources.ActiveDownloads = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new()) }

	$global:DashboardConfig.UI.ToolTipMain = ($toolTipMain = New-ThemedToolTip)
	$global:DashboardConfig.UI.ToolTipSettings = ($toolTipSettings = New-ThemedToolTip)
	$global:DashboardConfig.UI.ToolTipExtra = ($toolTipExtra = New-ThemedToolTip)

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

	$p = @{ type = 'Panel'; width = 470; height = 30; id = 'TopBar'; tag = 'TopBar' }
	$topBar = SetUIElement @p
	Register-ThemedTopBarPaint $topBar

	$p = @{ type = 'Label'; width = 140; height = 12; top = 5; left = 10; id = 'TitleLabel'; text = 'Entropia Dashboard'; tag = 'Title'; font = (Get-UIFont 'Segoe UI' 8 'Bold') }
	$titleLabelForm = SetUIElement @p

	$p = @{ type = 'Label'; width = 160; height = 10; top = 16; left = 10; id = 'CopyrightLabel'; text = [char]0x00A9 + ' Immortal / Divine 2026 - v3.0.2'; tag = 'Subtitle'; font = (Get-UIFont 'Segoe UI' 6 'Italic') }
	$copyrightLabelForm = SetUIElement @p

	$p = @{ type = 'Button'; width = 30; height = 24; top = 3; left = 327; id = 'ThemeToggle'; text = [char]0x25D0; fs = 'Flat'; tag = 'HeaderButton'; font = (Get-UIFont 'Segoe UI' 20 'Bold'); tooltip = "Theme Manager`nOpen theme selector and custom theme editor."; noCustomPaint = $true }
	$btnThemeToggle = SetUIElement @p
	$btnThemeToggle.FlatAppearance.BorderSize = 0

	$p = @{ type = 'Button'; width = 40; height = 24; top = 3; left = 360; id = 'InfoForm'; text = 'Info'; fs = 'Flat'; tag = 'HeaderButton'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Help / Info`nOpens the project website for documentation and support."; noCustomPaint = $true }
	$btnInfoForm = SetUIElement @p
	$btnInfoForm.FlatAppearance.BorderSize = 0

	$p = @{ type = 'Button'; width = 30; height = 24; top = 3; left = 405; id = 'MinForm'; text = '_'; fs = 'Flat'; tag = 'HeaderButton'; font = (Get-UIFont 'Segoe UI' 10 'Bold'); tooltip = "Minimize`nMinimize the dashboard to the taskbar."; noCustomPaint = $true }
	$btnMinimizeForm = SetUIElement @p
	$btnMinimizeForm.FlatAppearance.BorderSize = 0

	$p = @{ type = 'Button'; width = 30; height = 24; top = 3; left = 437; id = 'CloseForm'; text = [char]0x166D; fs = 'Flat'; tag = 'Danger'; font = (Get-UIFont 'Segoe UI' 10 'Bold'); tooltip = "Exit`nClose the Entropia Dashboard application."; noCustomPaint = $true }
	$btnCloseForm = SetUIElement @p
	$btnCloseForm.FlatAppearance.BorderSize = 0

	$p = @{ type = 'Button'; width = 125; height = 30; top = 40; left = 15; id = 'Launch'; text = 'Launch ' + [char]0x25BE; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Launch`nLeft Click: Open launch menu.`nRight Click: Open menu for launch options."; noCustomPaint = $true }
	$btnLaunch = SetUIElement @p

	$LaunchContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
	$LaunchContextMenu.Name = 'LaunchContextMenu'
	$LaunchContextMenu.Items.Add('Loading...') | Out-Null
	$btnLaunch.ContextMenuStrip = $LaunchContextMenu

	$p = @{ type = 'Button'; width = 125; height = 30; top = 40; left = 150; id = 'Login'; text = 'Login'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Login`nLogs in selected clients using profile settings."; noCustomPaint = $true }
	$btnLogin = SetUIElement @p

	$p = @{ type = 'Button'; width = 80; height = 30; top = 40; left = 285; id = 'Settings'; text = 'Settings'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Settings`nOpen the configuration menu to adjust paths, profiles, and options."; noCustomPaint = $true }
	$btnSettings = SetUIElement @p

	$p = @{ type = 'Button'; width = 80; height = 30; top = 40; left = 375; id = 'Terminate'; text = 'Terminate ' + [char]0x25BE; fs = 'Flat'; tag = 'Danger'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Terminate`nLeft Click: Instantly close selected clients.`nRight Click: Disconnect selected clients (TCP close)."; noCustomPaint = $true }
	$btnStop = SetUIElement @p

	$p = @{ type = 'Button'; width = 125; height = 30; top = 75; left = 15; id = 'Ftool'; text = 'Ftool ' + [char]0x25BE; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Ftool`nLeft Click: Open Ftool for selected client.`nRight Click: Save/Load session template for selected client."; noCustomPaint = $true }
	$btnFtool = SetUIElement @p
	$btnFtool.Add_Click({
			if ($global:DashboardConfig.UI.DataGridMain.SelectedRows.Count -eq 0) { return }
			$row = $global:DashboardConfig.UI.DataGridMain.SelectedRows[0]
			FtoolSelectedRow $row
		})

	$p = @{ type = 'Button'; width = 125; height = 30; top = 75; left = 150; id = 'Macro'; text = 'Macro ' + [char]0x25BE; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Macro`nLeft Click: Open Macro for selected client.`nRight Click: Save/Load session template for selected client."; noCustomPaint = $true }
	$btnMacro = SetUIElement @p
	$btnMacro.Add_Click({
			if ($global:DashboardConfig.UI.DataGridMain.SelectedRows.Count -eq 0) { return }
			$row = $global:DashboardConfig.UI.DataGridMain.SelectedRows[0]
			MacroSelectedRow $row
		})

	$p = @{ type = 'Button'; width = 80; height = 30; top = 75; left = 285; id = 'Extra'; text = 'Extra ' + [char]0x25BE; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Extra`nLeft Click: Open the Extra panel for World Boss timers, Notes, Tools, and Notifications.`nRight Click: Open QuickCommands for selected client."; noCustomPaint = $true }
	$btnExtra = SetUIElement @p

	$p = @{ type = 'Button'; width = 80; height = 30; top = 75; left = 375; id = 'Wiki'; text = 'Map ' + [char]0x25BE; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Map`nLeft Click: Open the Node Maps (every field, dungeon & event map).`nRight Click: Open the Wiki panel for Entropia."; noCustomPaint = $true }
	$btnWiki = SetUIElement @p

	Register-ThemedButtonFadeEffect -Button $btnInfoForm
	Register-ThemedButtonFadeEffect -Button $btnThemeToggle
	Register-ThemedButtonFadeEffect -Button $btnMinimizeForm
	Register-ThemedButtonFadeEffect -Button $btnCloseForm -Variant 'Danger'
	Register-ThemedButtonFadeEffect -Button $btnLaunch
	Register-ThemedButtonFadeEffect -Button $btnLogin
	Register-ThemedButtonFadeEffect -Button $btnSettings
	Register-ThemedButtonFadeEffect -Button $btnStop -Variant 'Danger'
	Register-ThemedButtonFadeEffect -Button $btnFtool
	Register-ThemedButtonFadeEffect -Button $btnMacro
	Register-ThemedButtonFadeEffect -Button $btnExtra
	Register-ThemedButtonFadeEffect -Button $btnWiki

	$p = @{ type = 'DataGridView'; width = 450; height = 300; top = 115; left = 10; id = 'DataGridMain'; text = ''; font = (Get-UIFont 'Segoe UI' 9) }
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

	$p = @{ type = 'Panel'; height = 8; dock = 'Bottom'; cursor = 'SizeNS'; id = 'ResizeGrip'; tag = 'MainBg'; tooltip = "Resize`nClick and drag to resize the dashboard window." }
	$resizeGrip = SetUIElement @p
	Register-ThemedResizeGripPaint $resizeGrip

	#endregion

	$script:CurrentBuilderToolTip = $toolTipSettings

	#region SettingsForm

	$settingsTabs = New-Object Custom.DarkTabControl
	$settingsTabs.Dock = 'Top'
	$settingsTabs.Height = 740
	$settingsTabs.FillWidth = $true
	$settingsTabs.ItemSize = New-Object System.Drawing.Size(145, 30)
	$settingsTabs.AllowFormDrag = $true

	$tabGeneral = New-Object System.Windows.Forms.TabPage 'General'
	$tabLaunchSettings = New-Object System.Windows.Forms.TabPage 'Launch Settings'
	$tabLoginSettings = New-Object System.Windows.Forms.TabPage 'Login Settings'
	$tabLoginSettings.AutoScroll = $true
	$tabHotkeys = New-Object System.Windows.Forms.TabPage 'Hotkeys'

	$settingsTabs.TabPages.Add($tabGeneral)
	$settingsTabs.TabPages.Add($tabLaunchSettings)
	$settingsTabs.TabPages.Add($tabLoginSettings)
	$settingsTabs.TabPages.Add($tabHotkeys)

	$settingsForm.Controls.Add($settingsTabs)

	#region General

	$p = @{ type = 'Label'; width = 127; height = 20; top = 31; left = 20; id = 'LabelLauncher'; text = 'Select Main Launcher:'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Main Launcher`nSelect the 'Launcher.exe' of your main game client.`nThis path is used as the base for all operations." }
	$lblLauncher = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 230; height = 30; top = 56; left = 20; id = 'InputLauncher'; text = ''; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Main Launcher`nSelect the 'Launcher.exe' of your main game client.`nThis path is used as the base for all operations." }
	$txtLauncher = SetUIElement @p
	$p = @{ type = 'Button'; width = 55; height = 25; top = 26; left = 195; id = 'Browse'; text = 'Browse'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Main Launcher`nSelect the 'Launcher.exe' of your main game client.`nThis path is used as the base for all operations."; noCustomPaint = $true }
	$btnBrowseLauncher = SetUIElement @p

	$p = @{ type = 'Label'; width = 85; height = 20; top = 101; left = 20; id = 'LabelProcess'; text = 'Process Name:'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Process Name`nThe name of the game process (usually 'neuz').`nUsed to detect running clients." }
	$lblProcessName = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 230; height = 30; top = 126; left = 20; id = 'InputProcess'; text = ''; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Process Name`nThe name of the game process (usually 'neuz').`nUsed to detect running clients." }
	$txtProcessName = SetUIElement @p

	$p = @{ type = 'Label'; width = 230; height = 20; top = 171; left = 20; id = 'LabelMax'; text = 'Max Total Clients For Selection:'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Max Clients`nLimit the total number of clients the dashboard can launch for the default launcher or selected profile.`nPrevents accidental mass-launching." }
	$lblMaxClients = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 230; height = 30; top = 196; left = 20; id = 'InputMax'; text = ''; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Max Clients`nLimit the total number of clients the dashboard can launch for the default launcher or selected profile`nPrevents accidental mass-launching." }
	$txtMaxClients = SetUIElement @p

	$p = @{ type = 'CheckBox'; width = 200; height = 20; top = 20; left = 0; id = 'AutoStartWithWindows'; text = 'Auto Start with Windows'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Auto Start`nEnable this to have the Entropia Dashboard start automatically when you log into Windows." }
	$chkAutoStartWindows = SetUIElement @p
	
	$p = @{ type = 'Label'; width = 120; height = 20; top = 60; left = 20; id = 'LabelNotificationPosition'; text = 'Notification Position:'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Notification Position`nSets the corner of the screen where notifications will appear." }
	$lblNotificationPosition = SetUIElement @p
	
	$p = @{ type = 'Custom.DarkComboBox'; width = 105; height = 25; top = 57; left = 145; id = 'NotificationPositionSelector'; font = (Get-UIFont 'Segoe UI' 9); dropDownStyle = 'DropDownList' }
	$ddlNotificationPosition = SetUIElement @p
	$ddlNotificationPosition.Items.AddRange(@('Bottom Right', 'Bottom Left', 'Top Right', 'Top Left')) | Out-Null

	$p = @{ type = 'Label'; width = 320; height = 20; top = 31; left = 4; id = 'LabelProfiles'; text = 'Select Client Profile for Default Launch:'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Default Profile`nSelect the profile to be used when clicking the main 'Launch' button.`nTip: clicking a client name in the list below selects that client and its profile for the Launch button - every profile remembers its own selected client." }
	$lblProfiles = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 320; height = 210; top = 56; left = 4; id = 'ProfileGrid'; text = ''; font = (Get-UIFont 'Segoe UI' 9) }
	$ProfileGrid = SetUIElement @p
	$ProfileGrid.AllowUserToAddRows = $false
	$ProfileGrid.RowHeadersVisible = $false
	$ProfileGrid.EditMode = 'EditProgrammatically'
	$ProfileGrid.SelectionMode = 'FullRowSelect'
	$ProfileGrid.AutoSizeColumnsMode = 'Fill'
	$ProfileGrid.ColumnHeadersHeight = 30
	$ProfileGrid.RowTemplate.Height = 25
	$ProfileGrid.MultiSelect = $true

	$profileGridCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Name'; FillWeight = 26; ReadOnly = $true; ToolTipText = "Profile Name`nThe unique name of this client profile." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Path'; FillWeight = 74; ReadOnly = $true; ToolTipText = "Path`nLocation on disk. Double-click to open folder." })
	)
	$ProfileGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$profileGridCols)

	$p = @{ type = 'Button'; width = 60; height = 25; top = 276; left = 4; id = 'AddProfile'; text = 'Search'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Search Profile`nSearch for an existing game folder and register it as a profile."; noCustomPaint = $true }
	$btnAddProfile = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 276; left = 68; id = 'StartJunction'; text = 'Create'; fs = 'Flat'; tag = 'Success'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Create Profile`nCreates a lightweight copy (Junction) of your client into the Profiles Folder below.`nAllows separate settings for each profile."; noCustomPaint = $true }
	$btnStartJunction = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 276; left = 132; id = 'RenameProfile'; text = 'Rename'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Rename`nChange the display name of the selected profile."; noCustomPaint = $true }
	$btnRenameProfile = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 276; left = 196; id = 'RemoveProfile'; text = 'Remove'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Remove`nRemove the selected profile from the list (does not delete files)."; noCustomPaint = $true }
	$btnRemoveProfile = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 276; left = 260; id = 'DeleteProfile'; text = 'Delete'; fs = 'Flat'; tag = 'Danger'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Delete`nPermanently delete the selected profile and its files from the disk."; noCustomPaint = $true }
	$btnDeleteProfile = SetUIElement @p

	$p = @{ type = 'Label'; width = 140; height = 20; top = 316; left = 4; id = 'LabelJunction'; text = 'Select Profiles Folder:'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Profiles Folder`nThe directory where your client profiles (copies) will be stored." }
	$lblJunction = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 320; height = 30; top = 341; left = 4; id = 'InputJunction'; text = ''; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Profiles Folder`nThe directory where your client profiles (copies) will be stored." }
	$txtJunction = SetUIElement @p

	$p = @{ type = 'Button'; width = 55; height = 25; top = 311; left = 154; id = 'BrowseJunction'; text = 'Browse'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Browse Folder`nChoose the directory for storing client profiles."; noCustomPaint = $true }
	$btnBrowseJunction = SetUIElement @p

	$p = @{ type = 'Button'; width = 120; height = 25; top = 381; left = 4; id = 'CopyNeuz'; text = 'Copy Neuz.exe'; fs = 'Flat'; tag = 'Primary'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Copy neuz.exe / game folders`nCopies the complete 'bin32' and 'bin64' folders from the main client to the profiles selected above (all profiles if none is selected).`nUse this after patching the main client to update the profiles.`nFiles that cannot be overwritten (e.g. running or read-only) are skipped and reported - the copy continues."; noCustomPaint = $true }
	$btnCopyNeuz = SetUIElement @p

	$p = @{ type = 'Label'; width = 570; height = 20; top = 32; left = 15; id = 'LabelClients'; text = 'Remembered clients per profile - tick what runs automatically after login:'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Auto-Start Clients`nEvery client that logs into a profile is remembered here automatically and can be maintained by hand (Add / Rename / Remove).`n`nReconnect? / Hide? decide for this single client whether it is reconnected on disconnect / hidden from Taskbar+Alt-Tab. A client without its own Hide? tick (one that has not logged in yet) follows the 'Hide Profiles' switch next to the buttons below.`nFtool? / Macro? / QuickCmd? start that tool for this single client once it has logged in - never while launching, logging in or logged out. Macro windows are opened but left idle.`nClick a client name to make it the selected client of its profile (the Launch button follows it)." }
	$lblClients = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 570; height = 185; top = 56; left = 15; id = 'ClientGrid'; text = ''; font = (Get-UIFont 'Segoe UI' 9) }
	$ClientGrid = SetUIElement @p
	$ClientGrid.AllowUserToAddRows = $false
	$ClientGrid.RowHeadersVisible = $false
	$ClientGrid.EditMode = 'EditProgrammatically'
	$ClientGrid.SelectionMode = 'FullRowSelect'
	$ClientGrid.AutoSizeColumnsMode = 'Fill'
	$ClientGrid.ColumnHeadersHeight = 30
	$ClientGrid.RowTemplate.Height = 25
	$ClientGrid.MultiSelect = $false

	$clientGridCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Profile'; FillWeight = 14; ReadOnly = $true; ToolTipText = "Profile`nThe profile this client logs into.`nSet when the client is added (uses the profile selected in the list above)." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Client'; FillWeight = 26; ReadOnly = $true; ToolTipText = "Client`nThe window title of the remembered client.`nClients are recorded automatically the moment they log in.`nClick the name to make this the selected client of its profile (used by the Launch button)." }),
		(New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ HeaderText = 'Reconnect?'; FillWeight = 15; ReadOnly = $false; ToolTipText = "Auto-Reconnect`nIf checked, this single client is reconnected when it disconnects.`nOther clients of the same profile are not affected.`nA client without a row here yet (not logged in for the first time) is not reconnected." }),
		(New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ HeaderText = 'Hide?'; FillWeight = 11; ReadOnly = $false; ToolTipText = "Hide Window`nIf checked, this single client window is hidden from Taskbar/Alt-Tab when minimized.`nClients that have no row here yet (not logged in for the first time) follow the 'Hide Profiles' switch below." }),
		(New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ HeaderText = 'Ftool?'; FillWeight = 10; ReadOnly = $false; ToolTipText = "Auto F-Tool`nStarts the F-Tool for this single client once it has finished logging in.`nNever starts while launching, while logging in, or while logged out." }),
		(New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ HeaderText = 'Macro?'; FillWeight = 10; ReadOnly = $false; ToolTipText = "Auto Macro`nOpens the Macro Helper for this single client once it has finished logging in - it stays idle.`nNever starts while launching, while logging in, or while logged out." }),
		(New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ HeaderText = 'QuickCmd?'; FillWeight = 14; ReadOnly = $false; ToolTipText = "Auto Quick Command`nRuns a Quick Command for this single client once it has finished logging in.`nNever runs while launching, while logging in, or while logged out." })
	)
	$ClientGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$clientGridCols)

	$p = @{ type = 'Button'; width = 60; height = 25; top = 248; left = 15; id = 'AddClient'; text = 'Add'; fs = 'Flat'; tag = 'Success'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Add Client`nRemember a client under the profile selected above and tick what should start for it."; noCustomPaint = $true }
	$btnAddClient = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 248; left = 79; id = 'RenameClient'; text = 'Rename'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Rename`nChange the window title of the selected remembered client (the auto-start ticks are kept)."; noCustomPaint = $true }
	$btnRenameClient = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 248; left = 143; id = 'RemoveClient'; text = 'Remove'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Remove`nForget the selected client - it will be remembered again on its next login."; noCustomPaint = $true }
	$btnRemoveClient = SetUIElement @p

	$p = @{ type = 'CheckBox'; width = 160; height = 20; top = 251; left = 425; id = 'HideProfiles'; text = 'Hide Profiles'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Hide Profiles`nChecked: every client without its own Hide? tick in the list above is hidden from Taskbar/Alt-Tab while it is minimized - for example a client that has not logged in yet.`nUnchecked: those clients stay visible.`nA client's own Hide? tick always wins over this switch." }
	$chkHideProfiles = SetUIElement @p

	$grpRequired = New-Object System.Windows.Forms.GroupBox
	$grpRequired.Text = 'Required Settings'
	$grpRequired.Location = New-Object System.Drawing.Point(0, 4)
	$grpRequired.Size = New-Object System.Drawing.Size(256, 240)
	$grpRequired.Font = Get-UIFont 'Segoe UI' 9

	$grpPreferences = New-Object System.Windows.Forms.GroupBox
	$grpPreferences.Text = 'Preferences'
	$grpPreferences.Location = New-Object System.Drawing.Point(0, 250)
	$grpPreferences.Size = New-Object System.Drawing.Size(256, 94)
	$grpPreferences.Font = Get-UIFont 'Segoe UI' 9

	$grpProfiles = New-Object System.Windows.Forms.GroupBox
	$grpProfiles.Text = 'Client Profiles'
	$grpProfiles.Location = New-Object System.Drawing.Point(261, 4)
	$grpProfiles.Size = New-Object System.Drawing.Size(335, 416)
	$grpProfiles.Font = Get-UIFont 'Segoe UI' 9

	$grpClients = New-Object System.Windows.Forms.GroupBox
	$grpClients.Text = 'Auto-Start Clients'
	$grpClients.Location = New-Object System.Drawing.Point(0, 424)
	$grpClients.Size = New-Object System.Drawing.Size(596, 278)
	$grpClients.Font = Get-UIFont 'Segoe UI' 9

	$grpRequired.Controls.AddRange(@($lblLauncher, $txtLauncher, $btnBrowseLauncher, $lblProcessName, $txtProcessName, $lblMaxClients, $txtMaxClients))
	$grpPreferences.Controls.AddRange(@($chkAutoStartWindows, $lblNotificationPosition, $ddlNotificationPosition))
	$grpProfiles.Controls.AddRange(@($lblProfiles, $ProfileGrid, $btnAddProfile, $btnStartJunction, $btnRenameProfile, $btnRemoveProfile, $btnDeleteProfile, $lblJunction, $txtJunction, $btnBrowseJunction, $btnCopyNeuz))
	$grpClients.Controls.AddRange(@($lblClients, $ClientGrid, $btnAddClient, $btnRenameClient, $btnRemoveClient, $chkHideProfiles))

	$tabGeneral.Controls.AddRange(@($grpRequired, $grpPreferences, $grpProfiles, $grpClients))
	if (Test-DashboardCommand 'ApplyControlInternal')
	{
		$themeColors = Get-Theme
		foreach ($grp in @($grpRequired, $grpPreferences, $grpProfiles, $grpClients)) { ApplyControlInternal $grp $themeColors }
	}
	$uiPropertiesToAdd.NotificationPositionSelector = $ddlNotificationPosition
	#endregion

	#region Launch Settings

	$p = @{ type = 'Label'; width = 300; height = 20; top = 8; left = 15; text = 'SAVED SETUPS'; tag = 'Section'; font = (Get-UIFont 'Segoe UI' 9 'Bold') }
	$lblSetupsSection = SetUIElement @p

	$p = @{ type = 'Label'; width = 565; height = 20; top = 33; left = 15; id = 'LabelSetups'; text = 'Select clients in the main list, then click Add to save them as a launch setup:'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Saved Setups`nList of pre-configured client launch groups." }
	$lblSetups = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 565; height = 320; top = 60; left = 15; id = 'SetupGrid'; text = ''; font = (Get-UIFont 'Segoe UI' 9) }
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

	$p = @{ type = 'Button'; width = 60; height = 25; top = 400; left = 15; id = 'AddSetup'; text = 'Add'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Add Setup`nSave the currently selected clients in the main list as a new launch setup."; noCustomPaint = $true }
	$btnAddSetup = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 400; left = 80; id = 'EditSetup'; text = 'Edit'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Edit Setup`nModify the clients included in the selected setup."; noCustomPaint = $true }
	$btnEditSetup = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 400; left = 145; id = 'RenameSetup'; text = 'Rename'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Rename`nChange the name of the selected setup."; noCustomPaint = $true }
	$btnRenameSetup = SetUIElement @p
	$p = @{ type = 'Button'; width = 60; height = 25; top = 400; left = 210; id = 'DeleteSetup'; text = 'Delete'; fs = 'Flat'; tag = 'Danger'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Delete`nRemove the selected setup configuration."; noCustomPaint = $true }
	$btnDeleteSetup = SetUIElement @p

	$p = @{ type = 'CheckBox'; width = 250; height = 20; top = 445; left = 0; id = 'AutoStartSetup'; text = 'Auto Start Selected Setup'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Auto Setup`nAutomatically launch the currently selected setup when the Dashboard starts." }
	$chkAutoStartSetup = SetUIElement @p

	$p = @{ type = 'Label'; width = 230; height = 20; top = 440; left = 300; id = 'LabelLauncherTimeout'; text = 'Launcher Timeout (Seconds):'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Launcher Timeout`nSeconds to wait for the launcher to open the game client before retrying." }
	$lblLauncherTimeout = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 230; height = 30; top = 465; left = 300; id = 'InputLauncherTimeout'; text = '60'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Launcher Timeout`nSeconds to wait for the launcher to open the game client before retrying." }
	$txtLauncherTimeout = SetUIElement @p

	$tabLaunchSettings.Controls.AddRange(@($lblSetupsSection, $lblSetups, $SetupGrid, $btnAddSetup, $btnEditSetup, $btnRenameSetup, $btnDeleteSetup, $chkAutoStartSetup, $lblLauncherTimeout, $txtLauncherTimeout))
	#endregion

	#region Login Settings

	$p = @{ type = 'Label'; width = 80; height = 20; top = 10; left = 10; text = 'Edit Profile:'; font = (Get-UIFont 'Segoe UI' 9) }
	$lblProfSel = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 430; height = 25; top = 7; left = 100; id = 'LoginProfileSelector'; font = (Get-UIFont 'Segoe UI' 9); dropDownStyle = 'DropDownList' }
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

	$p = @{ type = 'Label'; width = 75; height = 20; top = 40; left = 10; text = 'Resolution:'; font = (Get-UIFont 'Segoe UI' 9) }
	$lblResolution = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 180; height = 25; top = 37; left = 85; id = 'ResolutionSelector'; font = (Get-UIFont 'Segoe UI' 9); dropDownStyle = 'DropDownList' }
	$ddlResolution = SetUIElement @p
    
	$resolutions = @(
		'1024 x 768', '1280 x 720', '1280 x 768', '1280 x 800', '1280 x 1024',
		'1360 x 768', '1440 x 900', '1600 x 815', '1600 x 900', '1600 x 1200',
		'1680 x 1050', '1920 x 1080', '1920 x 1440', '2560 x 1080', '2560 x 1440',
		'3440 x 1440', '3840 x 2160', '4096 x 3072', '5120 x 1440'
	)
	foreach ($res in $resolutions) { $ddlResolution.Items.Add($res) | Out-Null }

	$p = @{ type = 'Label'; width = 60; height = 20; top = 40; left = 285; text = 'DirectX:'; font = (Get-UIFont 'Segoe UI' 9) }
	$lblDirectX = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 185; height = 25; top = 37; left = 345; id = 'DirectXSelector'; font = (Get-UIFont 'Segoe UI' 9); dropDownStyle = 'DropDownList' }
	$ddlDirectX = SetUIElement @p
	$dxVersions = @('DirectX 9', 'DirectX 11')
	foreach ($dx in $dxVersions) { $ddlDirectX.Items.Add($dx) | Out-Null }

	$tabLoginSettings.Controls.AddRange(@($lblResolution, $ddlResolution, $lblDirectX, $ddlDirectX))

	$Pickers = @{}
	$rowY = 45

	$AddPickerRow = {
		param($LabelText, $KeyName, $Top, $Col, $Tooltip)
		$leftOffset = if ($Col -eq 2) { 300 } else { 10 }

		$p = @{ type = 'Label'; width = 100; height = 20; top = $Top; left = $leftOffset; text = $LabelText; font = (Get-UIFont 'Segoe UI' 8); tooltip = $Tooltip }
		$l = SetUIElement @p
		$p = @{ type = 'TextBox'; width = 80; height = 20; top = $Top; left = ($leftOffset + 100); id = "txt$KeyName"; text = '0,0'; font = (Get-UIFont 'Segoe UI' 8); tooltip = $Tooltip }
		$t = SetUIElement @p
		$p = @{ type = 'Button'; width = 40; height = 20; top = $Top; left = ($leftOffset + 185); id = "btnPick$KeyName"; text = 'Set'; fs = 'Flat'; tag = 'Primary'; font = (Get-UIFont 'Segoe UI' 7); tooltip = $Tooltip; noCustomPaint = $true }
		$b = SetUIElement @p
		Register-ThemedButtonFadeEffect -Button $b -Variant 'Primary'
		return @{ Label = $l; Text = $t; Button = $b }
	}

	$p = @{ type = 'Label'; width = 300; height = 20; top = ($rowY + 20); left = 15; text = 'BASIC SETTINGS'; tag = 'Section'; font = (Get-UIFont 'Segoe UI' 9 'Bold') }
	$tabLoginSettings.Controls.Add((SetUIElement @p))

	$p = @{ type = 'CheckBox'; width = 220; height = 20; top = ($rowY + 20); left = 330; id = 'ManualCoordsOverride'; text = 'Manual Coordinates'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Manual Coordinates`nOff (default): the coordinates below are calculated from the resolution this profile runs at - change the resolution and they follow on their own.`nOn: the values below are used exactly as you enter them; the boxes and their 'Set' buttons become editable.`nYour own values are never deleted when you untick the box, they are simply ignored until you tick it again." }
	$chkManualCoords = SetUIElement @p
	$tabLoginSettings.Controls.Add($chkManualCoords)

	$p = &$AddPickerRow 'Launcher Play:' 'LauncherPlay' ($rowY + 40) 1 "Set Coordinate`n1. Click your Launcher window.`n2. Hover over the Play/Start button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['LauncherPlay'] = $p
	$p = &$AddPickerRow 'First Nickname:' 'FirstNick' ($rowY + 40) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the 1st Nickname in the list.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['FirstNick'] = $p

	$p = @{ type = 'Label'; width = 300; height = 20; top = ($rowY + 75); left = 15; text = 'ADVANCED SETTINGS'; tag = 'Section'; font = (Get-UIFont 'Segoe UI' 9 'Bold') }
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

	$p = @{ type = 'Label'; width = 100; height = 20; top = ($rowY + 270); left = 300; text = 'Login Delay (s):'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Post-Login Delay`nSeconds to wait after login before performing actions (e.g., Collector)." }
	$lblPostLoginDelay = SetUIElement @p
	$p = @{ type = 'TextBox'; width = 80; height = 20; top = ($rowY + 270); left = 400; id = 'txtPostLoginDelayInput'; text = '5'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Post-Login Delay`nSeconds to wait after login before performing actions (e.g., Collector)." }
	$txtPostLoginDelay = SetUIElement @p
	$tabLoginSettings.Controls.AddRange(@($lblPostLoginDelay, $txtPostLoginDelay))

	$p = @{ type = 'Label'; width = 300; height = 20; top = ($rowY + 305); left = 15; text = 'RECONNECT SETTINGS'; tag = 'Section'; font = (Get-UIFont 'Segoe UI' 9 'Bold') }
	$tabLoginSettings.Controls.Add((SetUIElement @p))

	$p = @{ type = 'CheckBox'; width = 200; height = 20; top = ($rowY + 305); left = 330; id = 'NeverRestartingCollectorLogin'; text = 'Collector Double Click'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Collector Fix`nEnable if the Collector start button needs a double-click to work.`nApplies to the auto-collector after login and reconnection." }
	$chkNeverRestartingLogin = SetUIElement @p
	$tabLoginSettings.Controls.Add($chkNeverRestartingLogin)

	$p = &$AddPickerRow 'Login Wrong:' 'LoginDetailsOK' ($rowY + 325) 1 "Set Coordinate`n1. Click your game window.`n2. Hover over the Wrong Password 'OK' button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['LoginDetailsOK'] = $p
	$p = &$AddPickerRow 'Disconnect:' 'DisconnectOK' ($rowY + 325) 2 "Set Coordinate`n1. Click your game window.`n2. Hover over the Disconnect 'OK' button.`n3. Wait 3 seconds."; $tabLoginSettings.Controls.AddRange(@($p.Label, $p.Text, $p.Button)); $Pickers['DisconnectOK'] = $p

	$gridTop = $rowY + 370
	$gridHeight = 285

	$p = @{ type = 'DataGridView'; width = 560; height = $gridHeight; top = $gridTop; left = 10; id = 'LoginConfigGrid'; text = ''; font = (Get-UIFont 'Segoe UI' 9) }
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

	$p = @{ type = 'Panel'; width = 560; height = 50; top = 10; left = 10; id = 'MasterHotkeysPanel'; tag = 'SecondaryBg'; text = '' }
	$MasterHotkeysPanel = SetUIElement @p
	$tabHotkeys.Controls.Add($MasterHotkeysPanel)

	$p = @{ type = 'Label'; width = 115; height = 20; top = 15; left = 10; id = 'MasterPauseLabel'; text = 'Master Pause:'; tag = 'Subtitle'; font = (Get-UIFont 'Segoe UI' 9) }
	$lblMasterPause = SetUIElement @p
	$lblMasterPause.TextAlign = 'MiddleRight'
	$MasterHotkeysPanel.Controls.Add($lblMasterPause)

	$p = @{ type = 'Button'; width = 80; height = 26; top = 12; left = 130; id = 'MasterPauseHotkeyBtn'; text = 'none'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Master Pause Hotkey`nAssign a global hotkey to pause all ftool/macro hotkeys at once."; noCustomPaint = $true }
	$btnMasterPauseHotkey = SetUIElement @p
	$MasterHotkeysPanel.Controls.Add($btnMasterPauseHotkey)

	$p = @{ type = 'Label'; width = 115; height = 20; top = 15; left = 220; id = 'MasterResumeLabel'; text = 'Master Resume:'; tag = 'Subtitle'; font = (Get-UIFont 'Segoe UI' 9) }
	$lblMasterResume = SetUIElement @p
	$lblMasterResume.TextAlign = 'MiddleRight'
	$MasterHotkeysPanel.Controls.Add($lblMasterResume)

	$p = @{ type = 'Button'; width = 80; height = 26; top = 12; left = 340; id = 'MasterResumeHotkeyBtn'; text = 'none'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Master Resume Hotkey`nAssign a global hotkey to resume all ftool/macro hotkeys at once."; noCustomPaint = $true }
	$btnMasterResumeHotkey = SetUIElement @p
	$MasterHotkeysPanel.Controls.Add($btnMasterResumeHotkey)

	$p = @{ type = 'Label'; width = 110; height = 20; top = 15; left = 430; id = 'MasterHotkeysStatus'; text = '● Active'; tag = 'Status'; font = (Get-UIFont 'Segoe UI' 9 'Bold') }
	$lblMasterHotkeysStatus = SetUIElement @p
	$lblMasterHotkeysStatus.TextAlign = 'MiddleCenter'
	$MasterHotkeysPanel.Controls.Add($lblMasterHotkeysStatus)

	$p = @{ type = 'DataGridView'; width = 560; height = 520; top = 70; left = 10; id = 'HotkeysGrid'; text = ''; font = (Get-UIFont 'Segoe UI' 9) }
	$HotkeysGrid = SetUIElement @p
	$HotkeysGrid.AllowUserToAddRows = $false
	$HotkeysGrid.RowHeadersVisible = $false
	$HotkeysGrid.EditMode = 'EditProgrammatically'
	$HotkeysGrid.SelectionMode = 'FullRowSelect'
	$HotkeysGrid.AutoSizeColumnsMode = 'Fill'
	$HotkeysGrid.ColumnHeadersHeight = 35
	$HotkeysGrid.RowTemplate.Height = 25	
	$HotkeysGrid.MultiSelect = $true
	$HotkeysGrid.Add_MouseDown({
		param($s, $e) $script:HotkeysGridMouseActive = $true
	})
	$HotkeysGrid.Add_MouseUp({
		param($s, $e)
		$script:HotkeysGridMouseActive = $false
		if ($script:HotkeysListRefreshPending -and $script:HotkeysGridEventDepth -eq 0) {
			RefreshHotkeysList
		}
	})
	$hotkeysGridCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = "Hotkey`nEditable"; FillWeight = 25; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Assigned To'; FillWeight = 45; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Action'; FillWeight = 30; ReadOnly = $true })
	)
	$HotkeysGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$hotkeysGridCols)
	$tabHotkeys.Controls.Add($HotkeysGrid)

	$p = @{ type = 'Button'; width = 200; height = 40; top = 600; left = 195; id = 'UnregisterHotkey'; text = 'Unregister and delete selected Hotkeys permanently'; fs = 'Flat'; tag = 'Danger'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Unregister Hotkey`nDelete the selected hotkey assignment.`nDouble-click a row to edit the key instead."; noCustomPaint = $true }
	$btnUnregisterHotkey = SetUIElement @p
	$tabHotkeys.Controls.Add($btnUnregisterHotkey)

	#endregion

	$p = @{ type = 'Button'; width = 120; height = 40; top = 735; left = 20; id = 'Save'; text = 'Save'; fs = 'Flat'; tag = 'Success'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Save`nApply and save all changes to the configuration file."; noCustomPaint = $true }
	$btnSave = SetUIElement @p

	$p = @{ type = 'Button'; width = 120; height = 40; top = 735; left = 150; id = 'Cancel'; text = 'Cancel'; fs = 'Flat'; tag = 'Danger'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Cancel`nDiscard changes and close the settings window."; noCustomPaint = $true }
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
	$topBar.Controls.AddRange(@($titleLabelForm, $copyrightLabelForm, $btnInfoForm, $btnThemeToggle, $btnMinimizeForm, $btnCloseForm))

	#endregion

	$script:CurrentBuilderToolTip = $toolTipExtra

	#region ExtraForm


	$p = @{ type = 'Label'; text = 'EXTRA CONTROL PANEL'; font = (Get-UIFont 'Segoe UI' 10 'Bold'); tag = 'Danger'; top = 10; left = 15 }
	$lblTitle = SetUIElement @p
	$lblTitle.AutoSize = $true
	$lblTitle.Add_MouseDown({ param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.ExtraForm.Handle, 0xA1, 0x2, 0) })

	$p = @{ type = 'Label'; text = 'X'; font = (Get-UIFont 'Segoe UI' 12 'Bold'); tag = 'Subtitle'; top = 8; left = 370; cursor = 'Hand' }
	$btnClose = SetUIElement @p
	$btnClose.Add_Click({ HideExtraForm })

	$extraTabs = New-Object Custom.DarkTabControl
	$extraTabs.Location = New-Object System.Drawing.Point(0, 35)
	$extraTabs.Size = New-Object System.Drawing.Size(400, ($extraForm.Height - 35))
	$extraTabs.Anchor = 'Top, Bottom, Left, Right'
	$extraTabs.FillWidth = $true
	$extraTabs.ItemSize = New-Object System.Drawing.Size(99, 30)

	$tabMarket = New-Object System.Windows.Forms.TabPage 'Market'
	$tabTools = New-Object System.Windows.Forms.TabPage 'Tools'
	$tabWorldBoss = New-Object System.Windows.Forms.TabPage 'World Boss'
	$tabNotes = New-Object System.Windows.Forms.TabPage 'Notes'
	$tabNotifications = New-Object System.Windows.Forms.TabPage 'Notifications'

	$extraTabs.TabPages.Add($tabMarket)
	$extraTabs.TabPages.Add($tabTools)
	$extraTabs.TabPages.Add($tabNotes)

	#region World Boss

	$BossList = @(
		'Entropia King',
		'Clockworks',
		'Royal Knight',
		'Solarion',
		'C-A01',
		'Great Venux Tree',
		'Grinch'
	)

	$p = @{ type = 'Label'; text = 'ACCESS CODE'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 10; left = 15; tooltip = "Access Code`nEnter a valid code to send/receive World Boss alerts.`nTo get an acces code, you must be member of the Divine Discord.`nUse the Access Code TEST to test this feature out locally." }
	$lblCode = SetUIElement @p
	$lblCode.AutoSize = $true

	$p = @{ type = 'Label'; text = 'USERNAME'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 10; left = 200; tooltip = "Username`nYour name will be shown in Discord alerts you trigger." }
	$lblUser = SetUIElement @p
	$lblUser.AutoSize = $true

	$p = @{ type = 'TextBox'; width = 175; height = 25; top = 30; left = 15; font = (Get-UIFont 'Segoe UI' 10); tooltip = "Access Code`nEnter a valid code to send/receive World Boss alerts.`nTo get an acces code, you must be member of the Divine Discord.`nUse the Access Code TEST to test this feature out locally." }
	$codeBox = SetUIElement @p
	$codeBox.PasswordChar = '*'
	$codeBox.TextAlign = 'Left'
	
	if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['AccessCode'])
 {
		$codeBox.Text = $global:DashboardConfig.Config['Options']['AccessCode']
	}
	$codeBox.Add_TextChanged({
			if (-not $global:DashboardConfig.Config.Contains('Options')) { $global:DashboardConfig.Config['Options'] = [ordered]@{} }
			$global:DashboardConfig.Config['Options']['AccessCode'] = $this.Text
		})
	$codeBox.Add_Leave({ if (Test-DashboardCommand 'WriteConfig') { WriteConfig } })

	$p = @{ type = 'TextBox'; width = 185; height = 25; top = 30; left = 200; font = (Get-UIFont 'Segoe UI' 10); tooltip = "Username`nYour name will be shown in Discord alerts you trigger." }
	$userBox = SetUIElement @p
	$userBox.TextAlign = 'Left'

	if ($global:DashboardConfig.Config['User'] -and $global:DashboardConfig.Config['User']['Username'])
	{
		$userBox.Text = $global:DashboardConfig.Config['User']['Username']
	}

	$userBox.Add_TextChanged({
			if (-not $global:DashboardConfig.Config.Contains('User')) { $global:DashboardConfig.Config['User'] = [ordered]@{} }
			$global:DashboardConfig.Config['User']['Username'] = $this.Text
		})
	$userBox.Add_Leave({ if (Test-DashboardCommand 'WriteConfig') { WriteConfig } })

	$p = @{ type = 'Label'; text = 'Show Images'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 60; left = 15; tooltip = "Show Images`nToggle displaying boss images on buttons and in toast notifications." }
	$lblShowImages = SetUIElement @p
	$lblShowImages.AutoSize = $true

	$p = @{ type = 'Toggle'; width = 30; height = 17; top = 59; left = 90; cursor = 'Hand'; tooltip = "Show Images`nToggle displaying boss images on buttons and in toast notifications." }
	$chkShowImages = SetUIElement @p
	$chkShowImages.Name = 'ShowBossImages'

	$p = @{ type = 'Label'; text = 'World Boss Listener'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 60; left = 140; tooltip = "Listener`nEnable to receive network alerts when a World Boss spawns." }
	$lblWorldBossListener = SetUIElement @p
	$lblWorldBossListener.AutoSize = $true

	$p = @{ type = 'Toggle'; width = 30; height = 17; top = 59; left = 255; cursor = 'Hand'; tooltip = "Listener`nEnable to receive network alerts when a World Boss spawns." }
	$chkWorldBossListener = SetUIElement @p
	$chkWorldBossListener.Name = 'WorldBossListener'

	$p = @{ type = 'Button'; width = 80; height = 20; top = 57; left = 295; id = 'RefreshBosses'; text = 'Update ' + [char]0x21BB; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Refresh Schedule`nTrigger the daily schedule notification again with updated data."; noCustomPaint = $true }
	$btnRefreshBosses = SetUIElement @p

	$p = @{ type = 'Label'; text = 'KHELDOR'; font = (Get-UIFont 'Segoe UI' 9 'Bold'); tag = 'Title'; top = 85; left = 60; width = 140 }
	$lblNormalHeader = SetUIElement @p
	$lblNormalHeader.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

	$p = @{ type = 'Label'; text = 'Listen'; font = (Get-UIFont 'Segoe UI' 7); tag = 'Subtitle'; top = 105; left = 15; width = 44; tooltip = "Listen`nToggle notifications for this specific boss." }
	$lblListenNormal = SetUIElement @p
	$lblListenNormal.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

	$p = @{ type = 'Label'; text = 'CLICK TO REPORT'; font = (Get-UIFont 'Segoe UI' 7 'Bold'); tag = 'Danger'; top = 105; left = 59; width = 140; tooltip = "Report`nClick the button below to report this boss as spawned (Notifies everyone!)." }
	$lblReportNormal = SetUIElement @p
	$lblReportNormal.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

	$ButtonPanel = New-Object System.Windows.Forms.TableLayoutPanel
	$ButtonPanel.Location = New-Object System.Drawing.Point(15, 125) 
	$ButtonPanel.Size = New-Object System.Drawing.Size(370, 380) 
	$ButtonPanel.ColumnCount = 4
	$ButtonPanel.AutoScroll = $true 
	
	$ButtonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 12)))
	$ButtonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 38)))
	$ButtonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 12)))
	$ButtonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 38)))
	
	$percentage = 100 / $BossList.Count
	foreach ($i in 1..$BossList.Count)
	{
		$ButtonPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $percentage)))
	}

	if (-not $global:DashboardConfig.Resources.ActiveDownloads) { $global:DashboardConfig.Resources.ActiveDownloads = [System.Collections.ArrayList]::new() }

	foreach ($Boss in $BossList)
	{
		$bossNameNormal = $Boss
		$p = @{ type = 'Toggle'; width = 30; height = 17; cursor = 'Hand'; tooltip = "Toggle Listen`nEnable/Disable alerts for $bossNameNormal." }
		$tglNormal = SetUIElement @p
		$tglNormal.Anchor = [System.Windows.Forms.AnchorStyles]::None
		$tglNormal.Checked = $true
		if ($global:DashboardConfig.Config['BossFilter'] -and $global:DashboardConfig.Config['BossFilter'].Contains($bossNameNormal))
		{
			$tglNormal.Checked = [bool]::Parse($global:DashboardConfig.Config['BossFilter'][$bossNameNormal])
		}
		$tglNormal.Add_Click({
				param($s, $e)
				if (-not $global:DashboardConfig.Config.Contains('BossFilter')) { $global:DashboardConfig.Config['BossFilter'] = [ordered]@{} }
				$global:DashboardConfig.Config['BossFilter'][$bossNameNormal] = $s.Checked
				if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
			}.GetNewClosure())

		$p = @{ type = 'Button'; text = $Boss; height = 50; fs = 'Flat'; font = (Get-UIFont 'Segoe UI Semibold' 9); cursor = 'Hand'; dock = 'Fill'; noCustomPaint = $true }
		$btnNormal = SetUIElement @p
		$btnNormal.Tag = @{ Type = 'Boss'; BossName = $bossNameNormal }
		$btnNormal.Margin = New-Object System.Windows.Forms.Padding(2)
		$btnNormal.FlatAppearance.BorderSize = 0

		$btnNormal.Add_Click({
				if ([string]::IsNullOrWhiteSpace($codeBox.Text))
				{
					Show-DarkMessageBox 'Access Code Required' 'Wrong Access Code' 'Ok' 'Error'
					return
				}
				if ([string]::IsNullOrWhiteSpace($userBox.Text))
    {
					Show-DarkMessageBox 'Username Required' 'Missing Username' 'Ok' 'Error'
					return
				}
				$bName = $this.Tag['BossName']
				if ((Show-DarkMessageBox "Are you sure you want to report '$bName'?`nThis will notify everyone!" 'Confirm Report' 'YesNo' 'Question') -eq 'Yes')
				{
					Send-Message -code $codeBox.Text -bossName $bName
				}
			}.GetNewClosure())
		
		Update-BossButtonImage $btnNormal
		
		$ButtonPanel.Controls.Add($tglNormal)
		$ButtonPanel.Controls.Add($btnNormal)
	}

	if (-not $global:DashboardConfig.Resources.BossSchedule)
	{
		$global:DashboardConfig.Resources.BossSchedule = @{
			'Entropia King' = [TimeSpan]::Parse('18:30')
			'Solarion'      = [TimeSpan]::Parse('19:45')
			'Clockworks'    = [TimeSpan]::Parse('20:00')
			'Royal Knight'  = [TimeSpan]::Parse('20:15')
		}
	}
	$BossSchedule = $global:DashboardConfig.Resources.BossSchedule

	try { $tz = [TimeZoneInfo]::FindSystemTimeZoneById('W. Europe Standard Time') } catch { $tz = [TimeZoneInfo]::Local }
	$fontTimerImage = Get-UIFont 'Segoe UI' 9 'Bold'
	$fontTimerNoImage = Get-UIFont 'Segoe UI Semibold' 7

	$scheduledItems = [System.Collections.ArrayList]::new()
	if ($ButtonPanel)
	{
		foreach ($btn in $ButtonPanel.Controls)
		{
			if ($btn -is [System.Windows.Forms.Button])
			{
				$bName = if ($btn.Tag -is [System.Collections.IDictionary]) { $btn.Tag['BossName'] } else { $btn.Text }
				if ($BossSchedule.ContainsKey($bName))
				{
					$scheduledItems.Add(@{ Button = $btn; BossName = $bName }) | Out-Null
				}
			}
		}
	}

	if ($global:DashboardConfig.Resources.Timers['BossTimer'])
	{
		$global:DashboardConfig.Resources.Timers['BossTimer'].Stop()
		$global:DashboardConfig.Resources.Timers['BossTimer'].Dispose()
	}

	$bossTimer = New-Object System.Windows.Forms.Timer
	$bossTimerTick = {
		$nextInterval = 60000
		try
		{
			$this.Stop()
			if (-not $global:DashboardConfig.UI.MainForm -or $global:DashboardConfig.UI.MainForm.IsDisposed) { return }

			$serverTime = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
			$theme = Get-Theme

			foreach ($item in $scheduledItems)
			{
				$btn = $item.Button
				if ($btn.IsDisposed) { continue }
				
				$bName = $item.BossName
				if (-not $BossSchedule.ContainsKey($bName)) { continue }
				
				$spawnTime = $BossSchedule[$bName]
				$todaySpawn = $serverTime.Date.Add($spawnTime)
				
				if ($serverTime -gt $todaySpawn) { $nextSpawn = $todaySpawn.AddDays(1) } 
				else { $nextSpawn = $todaySpawn }
				
				$diff = $nextSpawn - $serverTime
				$timeStr = $diff.ToString('hh\:mm')
				$fixedtime = $spawnTime.ToString('hh\:mm')
				
				$hasImage = ($null -ne $btn.BackgroundImage)
				if ($hasImage)
				{
					$newText = "Fixed Spawn in $($timeStr)h ($($fixedtime))"
					if ($btn.Text -ne $newText)
					{
						$btn.Text = $newText
						$btn.TextAlign = [System.Drawing.ContentAlignment]::BottomCenter
						$btn.ForeColor = $theme['BossTimerText']
						if ($btn.Font -ne $fontTimerImage) { $btn.Font = $fontTimerImage }
					}
				}
				else
				{
					$newText = "$bName`nFixed Spawn in $($timeStr)h`n($($fixedtime))"
					if ($btn.Text -ne $newText)
					{
						$btn.Text = $newText
						$btn.TextAlign = [System.Drawing.ContentAlignment]::BottomCenter
						$btn.ForeColor = $theme['BossTimerText']
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
		catch
		{ 
			Write-Verbose "Boss timer tick failed: $_" 
			$nextInterval = 15000
		} 
		finally
		{
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

	$p = @{ type = 'DataGridView'; width = 370; height = 450; top = 10; left = 15; id = 'NoteGrid'; text = ''; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Notes`nYour saved notes and timers.`nDouble-click a row to edit." }
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

	$p = @{ type = 'Button'; width = 80; height = 25; top = 470; left = 15; id = 'AddNote'; text = 'Add'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Add Note`nCreate a new note, timer, or reminder."; noCustomPaint = $true }
	$btnAddNote = SetUIElement @p

	$p = @{ type = 'Button'; width = 80; height = 25; top = 470; left = 105; id = 'EditNote'; text = 'Edit'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Edit Note`nModify the selected note."; noCustomPaint = $true }
	$btnEditNote = SetUIElement @p

	$p = @{ type = 'Button'; width = 80; height = 25; top = 470; left = 195; id = 'RemoveNote'; text = 'Remove'; fs = 'Flat'; tag = 'Danger'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Remove Note`nDelete the selected note permanently."; noCustomPaint = $true }
	$btnRemoveNote = SetUIElement @p

	$tabNotes.Controls.AddRange(@(
			$NoteGrid,
			$btnAddNote,
			$btnEditNote,
			$btnRemoveNote
		))

	#endregion

	#region Notifications

	$p = @{ type = 'DataGridView'; width = 370; height = 450; top = 10; left = 15; id = 'NotificationGrid'; text = ''; font = (Get-UIFont 'Segoe UI' 9); tooltip = "History`nLog of recent notifications.`nDouble-click to view details." }
	$NotificationGrid = SetUIElement @p
	$NotificationGrid.AllowUserToAddRows = $false
	$NotificationGrid.RowHeadersVisible = $false
	$NotificationGrid.SelectionMode = 'FullRowSelect'
	$NotificationGrid.MultiSelect = $false
	$NotificationGrid.AutoSizeColumnsMode = 'Fill'
	$NotificationGrid.ColumnHeadersHeight = 30
	$NotificationGrid.RowTemplate.Height = 25
	$NotificationGrid.Add_MouseDown({
		param($s, $e) $script:NotificationGridMouseActive = $true
	})
	$NotificationGrid.Add_MouseUp({
		param($s, $e) $script:NotificationGridMouseActive = $false
	})

	$notificationCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Time'; FillWeight = 25; ReadOnly = $true; ToolTipText = "Time`nTimestamp of the event." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Type'; FillWeight = 20; ReadOnly = $true; ToolTipText = "Type`nSeverity (Info, Warning, Error)." }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Details'; FillWeight = 55; ReadOnly = $true; ToolTipText = "Details`nMessage content." })
	)
	$NotificationGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$notificationCols)

	$p = @{ type = 'Button'; width = 120; height = 25; top = 470; left = 15; id = 'ShowNotification'; text = 'Show Notification'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Show`nRe-display the selected notification toast."; noCustomPaint = $true }
	$btnShowNotification = SetUIElement @p

	$p = @{ type = 'Button'; width = 130; height = 25; top = 470; left = 145; id = 'HideAllNotifications'; text = 'Hide all notifications'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Hide All`nClose all currently visible notification toasts."; noCustomPaint = $true }
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
	$tabChestOpener = New-Object System.Windows.Forms.TabPage 'Chest Opener'

	$toolsTabs.TabPages.Add($tabWorldBoss)
	$toolsTabs.TabPages.Add($tabUpgradeCalc)
	$toolsTabs.TabPages.Add($tabChestOpener)
	$tabTools.Controls.Add($toolsTabs)

	$p = @{ type = 'Label'; width = 150; text = 'UPGRADE CALCULATOR'; font = (Get-UIFont 'Segoe UI' 9 'Bold'); tag = 'Subtitle'; top = 10; left = 15 }
	$lblCalcTitle = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 120; height = 25; top = 35; left = 15; id = 'UpgradeItemSelector'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Item Category`nSelect the type of item you want to upgrade." }
	$ddlUpgradeItem = SetUIElement @p
	foreach ($key in $global:DashboardConfig.Resources.UpgradeData.psobject.Properties.Name) { $ddlUpgradeItem.Items.Add($key) | Out-Null }

	$p = @{ type = 'Custom.DarkComboBox'; width = 60; height = 25; top = 35; left = 140; id = 'UpgradeTypeSelector'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Weapon Type`n1H or 2H (2H weapons cost double)." }
	$ddlUpgradeType = SetUIElement @p
	$ddlUpgradeType.Items.AddRange(@('1H', '2H')) | Out-Null

	$p = @{ type = 'Label'; text = 'Lvl:'; font = (Get-UIFont 'Segoe UI' 7); tag = 'Subtitle'; top = 38; left = 205; width = 20 }
	$lblLvlStart = SetUIElement @p
	$p = @{ type = 'Custom.DarkComboBox'; width = 45; height = 25; top = 35; left = 225; id = 'UpgradeLvlStart'; font = (Get-UIFont 'Segoe UI' 8) }
	$ddlUpgradeLvlStart = SetUIElement @p

	$p = @{ type = 'Label'; text = 'To:'; font = (Get-UIFont 'Segoe UI' 7); tag = 'Subtitle'; top = 38; left = 275; width = 20 }
	$lblLvlEnd = SetUIElement @p
	$p = @{ type = 'Custom.DarkComboBox'; width = 45; height = 25; top = 35; left = 295; id = 'UpgradeLvlEnd'; font = (Get-UIFont 'Segoe UI' 8) }
	$ddlUpgradeLvlEnd = SetUIElement @p

	$p = @{ type = 'Button'; width = 350; height = 25; top = 65; left = 15; id = 'AddUpgrade'; text = 'Add to List'; fs = 'Flat'; tag = 'Primary'; font = (Get-UIFont 'Segoe UI' 8 'Bold'); noCustomPaint = $true }
	$btnAddUpgrade = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 350; height = 120; top = 98; left = 15; id = 'UpgradeCartGrid'; font = (Get-UIFont 'Segoe UI' 8) }
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

	$p = @{ type = 'Button'; width = 170; height = 25; top = 225; left = 15; id = 'RemoveUpgrade'; text = 'Remove'; fs = 'Flat'; tag = 'Danger'; font = (Get-UIFont 'Segoe UI' 8); noCustomPaint = $true }
	$btnRemoveUpgrade = SetUIElement @p
	$p = @{ type = 'Button'; width = 170; height = 25; top = 225; left = 195; id = 'ResetUpgrade'; text = 'Reset All'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); noCustomPaint = $true }
	$btnResetUpgrade = SetUIElement @p

	$p = @{ type = 'Label'; width = 150; text = 'TOTAL MATERIALS REQUIRED'; font = (Get-UIFont 'Segoe UI' 8 'Bold'); tag = 'Subtitle'; top = 258; left = 15 }
	$lblMaterialsTitle = SetUIElement @p
	$p = @{ type = 'CheckBox'; width = 150; height = 20; top = 258; left = 225; id = 'UpgradeHideZero'; text = 'Hide Zeros'; tag = 'Subtitle'; font = (Get-UIFont 'Segoe UI' 7); checked = $true }
	$chkUpgradeHideZero = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 350; height = 150; top = 282; left = 15; id = 'UpgradeResultsGrid'; font = (Get-UIFont 'Consolas' 9) }
	$UpgradeResultsGrid = SetUIElement @p
	$UpgradeResultsGrid.ColumnHeadersVisible = $false
	$resCols = @(
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Material'; FillWeight = 70; ReadOnly = $true }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ HeaderText = 'Count'; FillWeight = 30; ReadOnly = $true })
	)
	$UpgradeResultsGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$resCols)
	$UpgradeResultsGrid.DefaultCellStyle.Alignment = 'MiddleLeft'

	$p = @{ type = 'Label'; width = 170; height = 20; top = 440; left = 15; id = 'UpgradeTotalPerin'; text = 'Total Perin: 0'; tag = 'Perin'; font = (Get-UIFont 'Segoe UI' 9 'Bold') }
	$lblTotalPerin = SetUIElement @p
	$p = @{ type = 'Label'; width = 170; height = 20; top = 440; left = 195; id = 'UpgradeTotalEcp'; text = 'Total ECP: 0'; tag = 'Ecp'; font = (Get-UIFont 'Segoe UI' 9 'Bold') }
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

	$p = @{ type = 'Label'; width = 180; text = 'CHEST OPENER'; font = (Get-UIFont 'Segoe UI' 9 'Bold'); tag = 'Subtitle'; top = 10; left = 15 }
	$lblChestTitle = SetUIElement @p

	$p = @{ type = 'Label'; text = 'TARGET WINDOW'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; height = 15; top = 35; left = 15; tooltip = "Target Window`nSelect the game client instance to execute chest clicks." }
	$lblChestWindow = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 350; height = 25; top = 52; left = 15; id = 'ChestWindowSelector'; font = (Get-UIFont 'Segoe UI' 8); dropDownStyle = 'DropDownList'; tooltip = "Target Window`nSelect the game client instance to execute chest clicks." }
	$ddlChestWindow = SetUIElement @p

	$p = @{ type = 'Label'; text = 'CHEST PRESET'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; height = 15; top = 85; left = 15; tooltip = "Chest Preset`nSelect the chest type from database presets." }
	$lblChestPreset = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 350; height = 25; top = 102; left = 15; id = 'ChestPresetSelector'; font = (Get-UIFont 'Segoe UI' 8); dropDownStyle = 'DropDownList'; tooltip = "Chest Preset`nSelect the chest type from database presets." }
	$ddlChestPreset = SetUIElement @p

	$p = @{ type = 'Panel'; width = 350; height = 90; top = 125; left = 15; tag = 'MainBg'; id = 'ChestRewardsPanel' }
	$pnlChestRewards = SetUIElement @p
	$pnlChestRewards.AutoScroll = $true

	$p = @{ type = 'Label'; text = 'CHEST KEY:'; font = (Get-UIFont 'Segoe UI' 8 'Bold'); tag = 'Subtitle'; width = 65; top = 225; left = 15; tooltip = "Chest Key`nKey pressed in-game to open the chest selection window before reward clicks." }
	$lblChestItemHotkey = SetUIElement @p

	$p = @{ type = 'Button'; width = 90; height = 24; top = 222; left = 85; id = 'ChestItemHotkeyBtn'; text = 'Hotkey'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8 'Bold'); tooltip = "Chest Key`nKey pressed in-game to open the chest selection window before reward clicks."; noCustomPaint = $true }
	$btnChestItemHotkey = SetUIElement @p

	$p = @{ type = 'Label'; text = 'START/STOP:'; font = (Get-UIFont 'Segoe UI' 8 'Bold'); tag = 'Subtitle'; width = 70; top = 225; left = 185; tooltip = "Toggle Hotkey`nKeyboard shortcut to start/stop opening chests." }
	$lblChestHotkey = SetUIElement @p

	$p = @{ type = 'Button'; width = 105; height = 24; top = 222; left = 260; id = 'ChestHotkeyBtn'; text = 'Hotkey'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8 'Bold'); tooltip = "Toggle Hotkey`nKeyboard shortcut to start/stop opening chests."; noCustomPaint = $true }
	$btnChestHotkey = SetUIElement @p

	$p = @{ type = 'Label'; text = 'AMOUNT:'; font = (Get-UIFont 'Segoe UI' 8 'Bold'); tag = 'Subtitle'; width = 55; top = 255; left = 15; tooltip = "Loop Count`nNumber of times to open the chest in sequence." }
	$lblChestCount = SetUIElement @p

	$p = @{ type = 'TextBox'; width = 50; height = 24; top = 252; left = 75; id = 'ChestCountInput'; text = '1'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Loop Count`nNumber of times to open the chest in sequence." }
	$txtChestCount = SetUIElement @p

	$p = @{ type = 'Label'; width = 230; height = 24; top = 252; left = 135; id = 'ChestStatusLabel'; text = 'Status: IDLE'; tag = 'Status'; font = (Get-UIFont 'Segoe UI' 8 'Bold') }
	$lblChestStatus = SetUIElement @p

	$p = @{ type = 'Button'; width = 165; height = 24; top = 282; left = 15; id = 'AddChestQueue'; text = '+ Add to Queue'; fs = 'Flat'; tag = 'Primary'; font = (Get-UIFont 'Segoe UI' 8 'Bold'); tooltip = "Add Chest to Queue`nAdd the current chest preset, reward choices, key, and amount to execution queue."; noCustomPaint = $true }
	$btnAddChestQueue = SetUIElement @p

	$p = @{ type = 'Button'; width = 85; height = 24; top = 282; left = 185; id = 'RemoveChestQueue'; text = 'Remove'; fs = 'Flat'; tag = 'Danger'; font = (Get-UIFont 'Segoe UI' 8); tooltip = 'Remove selected chest from queue.'; noCustomPaint = $true }
	$btnRemoveChestQueue = SetUIElement @p

	$p = @{ type = 'Button'; width = 90; height = 24; top = 282; left = 275; id = 'ResetChestQueue'; text = 'Clear All'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = 'Clear all chests from queue.'; noCustomPaint = $true }
	$btnResetChestQueue = SetUIElement @p

	$p = @{ type = 'DataGridView'; width = 350; height = 155; top = 312; left = 15; id = 'ChestQueueGrid'; font = (Get-UIFont 'Segoe UI' 8) }
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

	$uiPropertiesToAdd.ChestWindowSelector = $ddlChestWindow
	$uiPropertiesToAdd.ChestPresetSelector = $ddlChestPreset
	$uiPropertiesToAdd.ChestRewardsPanel = $pnlChestRewards
	$uiPropertiesToAdd.ChestCountInput = $txtChestCount
	$uiPropertiesToAdd.ChestItemHotkeyBtn = $btnChestItemHotkey
	$uiPropertiesToAdd.ChestHotkeyBtn = $btnChestHotkey
	$uiPropertiesToAdd.ChestHotkeySelector = $btnChestHotkey
	$uiPropertiesToAdd.ChestStatusLabel = $lblChestStatus
	$uiPropertiesToAdd.ChestQueueGrid = $ChestQueueGrid
	$uiPropertiesToAdd.AddChestQueue = $btnAddChestQueue
	$uiPropertiesToAdd.RemoveChestQueue = $btnRemoveChestQueue
	$uiPropertiesToAdd.ResetChestQueue = $btnResetChestQueue
	if (-not $global:DashboardConfig.Resources.ChestQueue) { $global:DashboardConfig.Resources.ChestQueue = [System.Collections.ArrayList]::new() }
	$uiPropertiesToAdd.ChestQueue = $global:DashboardConfig.Resources.ChestQueue

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
			if (Test-DashboardCommand 'Update-ChestRewardControls')
			{
				Update-ChestRewardControls
			}
		}.GetNewClosure())

	$ddlChestWindow.Add_DropDown({
			if (Test-DashboardCommand 'Update-ChestOpenerWindowList')
			{
				Update-ChestOpenerWindowList
			}
		}.GetNewClosure())

	$btnAddChestQueue.Add_Click({
			if (Test-DashboardCommand 'Add-ChestToQueue')
			{
				Add-ChestToQueue
			}
		}.GetNewClosure())

	$btnRemoveChestQueue.Add_Click({
			if (Test-DashboardCommand 'Remove-ChestFromQueue')
			{
				Remove-ChestFromQueue
			}
		}.GetNewClosure())

	$btnResetChestQueue.Add_Click({
			if (Test-DashboardCommand 'Reset-ChestQueue')
			{
				Reset-ChestQueue
			}
		}.GetNewClosure())

	$btnChestItemHotkey.Add_Click({
			$currentKey = $btnChestItemHotkey.Text
			if ($currentKey -eq 'Hotkey' -or $currentKey -eq 'none') { $currentKey = $null }

			$ownerForm = if ($global:DashboardConfig.UI.MainForm) { $global:DashboardConfig.UI.MainForm } else { $this.FindForm() }
			$newKey = Show-KeyCaptureDialog -currentKey $currentKey -OwnerForm $ownerForm

			if ($null -ne $newKey -and $newKey -ne $currentKey)
			{
				$btnChestItemHotkey.Text = $newKey
				if (-not $global:DashboardConfig.Config.Contains('ChestOpener'))
				{
					$global:DashboardConfig.Config['ChestOpener'] = [ordered]@{}
				}
				$global:DashboardConfig.Config['ChestOpener']['ItemHotkey'] = $newKey
				if (Test-DashboardCommand 'WriteConfig' -Verbose:$false) { WriteConfig }
			}
			elseif (-not $newKey)
   {
				$btnChestItemHotkey.Text = 'Hotkey'
				if (-not $global:DashboardConfig.Config.Contains('ChestOpener'))
				{
					$global:DashboardConfig.Config['ChestOpener'] = [ordered]@{}
				}
				$global:DashboardConfig.Config['ChestOpener']['ItemHotkey'] = ''
				if (Test-DashboardCommand 'WriteConfig' -Verbose:$false) { WriteConfig }
			}
		}.GetNewClosure())

	$btnChestHotkey.Add_Click({
			$currentKey = $btnChestHotkey.Text
			if ($currentKey -eq 'Hotkey' -or $currentKey -eq 'none') { $currentKey = $null }

			$ownerForm = if ($global:DashboardConfig.UI.MainForm) { $global:DashboardConfig.UI.MainForm } else { $this.FindForm() }
			$newKey = Show-KeyCaptureDialog -currentKey $currentKey -OwnerForm $ownerForm

			$oldHotkeyId = $global:DashboardConfig.State['ChestHotkeyId']
			if ($null -ne $newKey -and $newKey -ne $currentKey)
			{
				$btnChestHotkey.Text = $newKey

				$scriptBlock = {
					$fgHwnd = [Custom.Native]::GetForegroundWindow()

					$selectedWindow = $global:DashboardConfig.UI.ChestWindowSelector.SelectedItem

					if ($selectedWindow)
					{
						$targetHwnd = [IntPtr]::Zero

						if ($selectedWindow -is [IntPtr])
						{
							$targetHwnd = $selectedWindow
						}
						elseif ($selectedWindow.HWnd)
      {
							$targetHwnd = [IntPtr]$selectedWindow.HWnd
						}
						elseif ($selectedWindow.Handle)
      {
							$targetHwnd = [IntPtr]$selectedWindow.Handle
						}
						elseif ($selectedWindow.MainWindowHandle)
      {
							$targetHwnd = [IntPtr]$selectedWindow.MainWindowHandle
						}

						if ($targetHwnd -ne [IntPtr]::Zero -and $fgHwnd -ne $targetHwnd)
						{
							return
						}
					}

					if ($script:SyncState.IsRunning)
					{
						Stop-ChestOpener
					}
					else
     {
						Start-ChestOpener
					}
				}

				try
				{
					$global:DashboardConfig.State['ChestHotkeyId'] = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey 'ChestOpener' -OwnerLabel 'Chest Opener' -OldHotkeyId $oldHotkeyId
				}
				catch
    {
					Write-Warning "ChestOpener: Failed to register hotkey $newKey : $_"
				}

				if (-not $global:DashboardConfig.Config.Contains('ChestOpener'))
				{
					$global:DashboardConfig.Config['ChestOpener'] = [ordered]@{}
				}
				$global:DashboardConfig.Config['ChestOpener']['Hotkey'] = $newKey
				if (Test-DashboardCommand 'WriteConfig' -Verbose:$false) { WriteConfig }
			}
			elseif (-not $newKey -and $oldHotkeyId)
   {
				try { UnregisterHotkeyInstance -Id $oldHotkeyId -OwnerKey 'ChestOpener' } catch {}
				$global:DashboardConfig.State['ChestHotkeyId'] = $null
				$btnChestHotkey.Text = 'Hotkey'
				if (-not $global:DashboardConfig.Config.Contains('ChestOpener'))
				{
					$global:DashboardConfig.Config['ChestOpener'] = [ordered]@{}
				}
				$global:DashboardConfig.Config['ChestOpener']['Hotkey'] = ''
				if (Test-DashboardCommand 'WriteConfig' -Verbose:$false) { WriteConfig }
			}
		}.GetNewClosure())

	#region Market

	$p = @{ type = 'Label'; width = 62; height = 20; text = 'MARKET'; font = (Get-UIFont 'Segoe UI' 9 'Bold'); tag = 'Subtitle'; top = 12; left = 15 }
	$lblMarketTitle = SetUIElement @p

	$p = @{ type = 'Label'; width = 74; height = 16; text = 'ACCESS CODE'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 16; left = 82 }
	$lblMarketAccess = SetUIElement @p

	$p = @{ type = 'TextBox'; width = 100; height = 24; top = 11; left = 158; id = 'MarketCodeInput'; font = (Get-UIFont 'Segoe UI' 9); tooltip = "Access Code`nEnter your access code." }
	$marketCodeBox = SetUIElement @p
	$marketCodeBox.PasswordChar = '*'
	$marketCodeBox.TextAlign = 'Left'

	$p = @{ type = 'Button'; width = 58; height = 24; top = 11; left = 262; id = 'MarketUnlockBtn'; text = 'Unlock'; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 6 'Bold'); tooltip = "Unlock / Lock`nUnlock or lock the market view with the access code."; noCustomPaint = $true }
	$btnMarketUnlock = SetUIElement @p

	$p = @{ type = 'Label'; width = 61; height = 18; text = 'local'; font = (Get-UIFont 'Segoe UI' 9 'Bold'); tag = 'Subtitle'; top = 12; left = 324 }
	$lblMarketStatus = SetUIElement @p
	$lblMarketStatus.TextAlign = 'MiddleRight'

	$p = @{ type = 'Label'; width = 61; height = 14; text = 'journal only'; font = (Get-UIFont 'Segoe UI' 7); tag = 'Subtitle'; top = 36; left = 324 }
	$lblMarketScope = SetUIElement @p
	$lblMarketScope.TextAlign = 'MiddleRight'

	$p = @{ type = 'Label'; width = 200; height = 16; text = 'ITEM SEARCH'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 56; left = 15 }
	$lblMarketSearch = SetUIElement @p

	$p = @{ type = 'TextBox'; width = 300; height = 25; top = 74; left = 15; id = 'MarketSearchInput'; font = (Get-UIFont 'Segoe UI' 10); tooltip = "Item Search`nStart typing to fuzzy-match known item names.`nDown / Up to choose, Enter to accept, Esc to close the list." }
	$marketSearchBox = SetUIElement @p
	$marketSearchBox.TextAlign = 'Left'

	$p = @{ type = 'Button'; width = 44; height = 25; top = 74; left = 321; id = 'MarketFindBtn'; text = ''; fs = 'Flat'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "Find`nFocus the search box and show matching item names."; noCustomPaint = $true }
	$btnMarketFind = SetUIElement @p

	$marketSuggest = New-Object System.Windows.Forms.ListBox
	$marketSuggest.Name = 'MarketSuggestions'
	$marketSuggest.Location = New-Object System.Drawing.Point(15, 103)
	$marketSuggest.Size = New-Object System.Drawing.Size(300, 144)
	$marketSuggest.IntegralHeight = $false
	$marketSuggest.BorderStyle = 'FixedSingle'
	$marketSuggest.Font = Get-UIFont 'Segoe UI' 9
	$marketSuggest.DrawMode = 'OwnerDrawFixed'
	$marketSuggest.ItemHeight = 18
	$marketSuggest.Visible = $false
	if (Test-DashboardCommand 'ApplyControlInternal') { ApplyControlInternal $marketSuggest (Get-Theme) }

	$p = @{ type = 'Label'; width = 150; height = 16; text = 'PRICE'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 106; left = 15; tooltip = "Price`nEnter the observed price. Trade currencies: Penya, Perin, Entropia Chip, Entropia Chip Perin, Red Diamond.`nPress Enter to store it in your local journal." }
	$lblMarketPrice = SetUIElement @p

	$p = @{ type = 'TextBox'; width = 70; height = 25; top = 124; left = 15; id = 'MarketPriceInput'; font = (Get-UIFont 'Segoe UI' 10); tooltip = "Price`nEnter the observed price in one of: Penya, Perin, Entropia Chip, Entropia Chip Perin, Red Diamond.`nPress Enter to store it in your local journal." }
	$marketPriceBox = SetUIElement @p
	$marketPriceBox.TextAlign = 'Left'

	$p = @{ type = 'Button'; width = 80; height = 25; top = 124; left = 90; id = 'MarketAddBtn'; text = 'Add Price'; fs = 'Flat'; tag = 'Primary'; font = (Get-UIFont 'Segoe UI' 8 'Bold'); tooltip = "Add Price`nStore the typed price in your local journal for the selected item."; noCustomPaint = $true }
	$btnMarketAdd = SetUIElement @p

	$p = @{ type = 'Custom.DarkComboBox'; width = 104; height = 25; top = 124; left = 172; id = 'MarketCurCombo'; font = (Get-UIFont 'Segoe UI' 8); dropDownStyle = 'DropDownList'; tooltip = "Price currency`nThe trade currency the price you are adding is in: Penya, Perin, Entropia Chip, Entropia Chip Perin or Red Diamond." }
	$ddlMarketCur = SetUIElement @p

	$p = @{ type = 'Label'; width = 30; height = 16; text = 'LAST'; font = (Get-UIFont 'Segoe UI' 7); tag = 'Subtitle'; top = 128; left = 278 }
	$lblMarketRangeLbl = SetUIElement @p
	$lblMarketRangeLbl.TextAlign = 'MiddleRight'

	$p = @{ type = 'Custom.DarkComboBox'; width = 42; height = 25; top = 123; left = 310; id = 'MarketTrendRange'; font = (Get-UIFont 'Segoe UI' 8); dropDownStyle = 'DropDownList'; tooltip = "Trend Window`nHow many of the most recent sales are compared.`nStored in the config when changed." }
	$ddlMarketRange = SetUIElement @p

	$p = @{ type = 'Label'; width = 40; height = 16; text = 'sales'; font = (Get-UIFont 'Segoe UI' 7); tag = 'Subtitle'; top = 128; left = 354 }
	$lblMarketRangeUnit = SetUIElement @p

	$p = @{ type = 'Label'; width = 80; height = 16; text = 'PRICE TREND'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 156; left = 15 }
	$lblMarketTrend = SetUIElement @p

	$p = @{ type = 'Label'; width = 30; height = 26; text = ''; font = (Get-UIFont 'Segoe UI' 16 'Bold'); top = 151; left = 96 }
	$lblMarketTrendArrow = SetUIElement @p

	$p = @{ type = 'Label'; width = 257; height = 18; text = ''; font = (Get-UIFont 'Segoe UI' 9 'Bold'); top = 156; left = 128 }
	$lblMarketTrendText = SetUIElement @p

	$p = @{ type = 'Label'; width = 370; height = 16; text = ''; font = (Get-UIFont 'Segoe UI' 7); tag = 'Muted'; top = 178; left = 15; tooltip = "Trend Detail`nOlder vs. recent average of the selected window.`nPrices are in Penya, Perin, Entropia Chip, Entropia Chip Perin or Red Diamond." }
	$lblMarketTrendDetail = SetUIElement @p

	$p = @{ type = 'Label'; width = 55; height = 16; text = 'ITEM'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 200; left = 15 }
	$lblMarketOnline = SetUIElement @p

	$p = @{ type = 'Label'; width = 60; height = 16; text = 'HISTORY'; font = (Get-UIFont 'Segoe UI' 8); tag = 'Subtitle'; top = 248; left = 15 }
	$lblMarketHistory = SetUIElement @p

	$p = @{ type = 'Label'; width = 126; height = 16; text = ''; font = (Get-UIFont 'Segoe UI' 7); top = 249; left = 78 }
	$lblMarketFeedback = SetUIElement @p

	$p = @{ type = 'Label'; width = 101; height = 16; text = ''; font = (Get-UIFont 'Segoe UI' 7); tag = 'Muted'; top = 249; left = 284 }
	$lblMarketHistoryCount = SetUIElement @p
	$lblMarketHistoryCount.TextAlign = 'MiddleRight'

	$p = @{ type = 'DataGridView'; width = 370; height = 250; top = 268; left = 15; id = 'MarketHistoryGrid'; font = (Get-UIFont 'Segoe UI' 8); tooltip = "History`nAll recorded prices, newest first.`nPick an item to filter the list down to that item." }
	$marketGrid = SetUIElement @p
	$marketGrid.ColumnHeadersHeight = 25
	$marketGrid.MultiSelect = $false
	$marketGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
	$marketGrid.ReadOnly = $true
	$marketGrid.AutoGenerateColumns = $false

	$marketCols = @(
		(New-Object System.Windows.Forms.DataGridViewImageColumn -Property @{ Name = 'MarketColIcon'; HeaderText = ''; FillWeight = 7; ReadOnly = $true; DataPropertyName = 'Icon'; ImageLayout = [System.Windows.Forms.DataGridViewImageCellLayout]::Zoom; }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'MarketColWhen'; HeaderText = 'When'; FillWeight = 17; ReadOnly = $true; DataPropertyName = 'When'; }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'MarketColItem'; HeaderText = 'Item'; FillWeight = 25; ReadOnly = $true; DataPropertyName = 'Item'; }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'MarketColPrice'; HeaderText = 'Price'; FillWeight = 11; ReadOnly = $true; DataPropertyName = 'Price'; }),
		(New-Object System.Windows.Forms.DataGridViewImageColumn -Property @{ Name = 'MarketColCur'; HeaderText = ''; FillWeight = 4; ReadOnly = $true; DataPropertyName = 'Currency'; ImageLayout = [System.Windows.Forms.DataGridViewImageCellLayout]::Zoom; }),
		(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'MarketColSource'; HeaderText = 'Source'; FillWeight = 13; ReadOnly = $true; DataPropertyName = 'Source'; })
	)
	$marketGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]$marketCols)

	$marketInv = [System.Globalization.CultureInfo]::InvariantCulture
	$marketGrid.Columns['MarketColIcon'].DefaultCellStyle.NullValue = $null
	$marketGrid.Columns['MarketColCur'].DefaultCellStyle.NullValue = $null
	$marketGrid.Columns['MarketColWhen'].DefaultCellStyle.Format = 'MM-dd HH:mm'
	$marketGrid.Columns['MarketColWhen'].DefaultCellStyle.FormatProvider = $marketInv
	$marketGrid.Columns['MarketColWhen'].DefaultCellStyle.NullValue = '-'
	$marketGrid.Columns['MarketColPrice'].DefaultCellStyle.Format = '0.##'
	$marketGrid.Columns['MarketColPrice'].DefaultCellStyle.FormatProvider = $marketInv
	$smallFont = Get-UIFont 'Segoe UI' 7
	$marketGrid.Columns['MarketColItem'].DefaultCellStyle.Font = $smallFont
	try
	{
		$gridProp = $marketGrid.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance, NonPublic')
		if ($gridProp) { $gridProp.SetValue($marketGrid, $true, $null) }
	}
	catch { }

	$tabMarket.Controls.AddRange(@(
			$lblMarketTitle,
			$lblMarketStatus,
			$lblMarketAccess,
			$marketCodeBox,
			$btnMarketUnlock,
			$lblMarketScope,
			$lblMarketSearch,
			$marketSearchBox,
			$btnMarketFind,
			$lblMarketPrice,
			$marketPriceBox,
			$btnMarketAdd,
			$ddlMarketCur,
			$lblMarketTrend,
			$lblMarketRangeLbl,
			$ddlMarketRange,
			$lblMarketRangeUnit,
			$lblMarketTrendArrow,
			$lblMarketTrendText,
			$lblMarketTrendDetail,
			$lblMarketOnline,
			$lblMarketHistory,
			$lblMarketFeedback,
			$lblMarketHistoryCount,
			$marketGrid
		))
	$tabMarket.Controls.Add($marketSuggest)
	$marketSuggest.BringToFront()

	$marketSearchBox.Add_TextChanged({
			if (Test-DashboardCommand 'Update-MarketSuggestions') { Update-MarketSuggestions }
			try { if (Test-DashboardCommand 'Update-MarketOfferGrid') { Update-MarketOfferGrid } } catch { }
		})

	$marketSearchBox.Add_KeyDown({
			param($src, $e)
			if ($null -eq $e) { return }
			$key = [System.Windows.Forms.Keys]$e.KeyCode
			if ($key -eq [System.Windows.Forms.Keys]::Down)
			{
				if (Test-DashboardCommand 'Move-MarketSuggestion') { Move-MarketSuggestion -Delta 1 }
				$e.SuppressKeyPress = $true
			}
			elseif ($key -eq [System.Windows.Forms.Keys]::Up)
			{
				if (Test-DashboardCommand 'Move-MarketSuggestion') { Move-MarketSuggestion -Delta -1 }
				$e.SuppressKeyPress = $true
			}
			elseif ($key -eq [System.Windows.Forms.Keys]::Enter)
			{
				$showing = $false
				if (Test-DashboardCommand 'Get-MarketUIControl')
				{
					$lst = Get-MarketUIControl -Name 'MarketSuggestions'
					if ($lst) { try { $showing = ($lst.Visible -and $lst.Items.Count -gt 0) } catch { $showing = $false } }
				}
				if ($showing -and (Test-DashboardCommand 'Accept-MarketSelection')) { Accept-MarketSelection }
				$e.SuppressKeyPress = $true
			}
			elseif ($key -eq [System.Windows.Forms.Keys]::Escape)
			{
				if (Test-DashboardCommand 'Hide-MarketSuggestions') { Hide-MarketSuggestions }
				$e.SuppressKeyPress = $true
			}
		})

	$marketSearchBox.Add_Leave({
			if (Test-DashboardCommand 'Hide-MarketSuggestionsIfOutside') { Hide-MarketSuggestionsIfOutside }
		})

	$marketSuggest.Add_MouseDown({
			param($src, $e)
			if ($null -eq $e) { return }
			if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
			if (Test-DashboardCommand 'Select-MarketSuggestionAt')
			{
				Select-MarketSuggestionAt -X ([int]$e.X) -Y ([int]$e.Y)
			}
		})

	$marketSuggest.Add_DrawItem({
			param($src, $e)
			try { Draw-MarketSuggestionItem -Sender $src -EventArgs $e }
			catch
			{
				try
				{
					if ($null -eq $e -or $e.Index -lt 0) { return }
					$e.DrawBackground()
					$t = ''
					if ($e.Index -lt $src.Items.Count) { $t = [string]$src.Items[$e.Index] }
					$brush = New-Object System.Drawing.SolidBrush($e.ForeColor)
					try { $e.Graphics.DrawString($t, $src.Font, $brush, $e.Bounds.X + 4, $e.Bounds.Y + 2) }
					finally { $brush.Dispose() }
				}
				catch { }
			}
		})

	$marketGrid.Add_CellFormatting({
			param($src, $e)
			try { Format-MarketHistoryCell -Sender $src -EventArgs $e } catch { }
		})

	$marketPriceBox.Add_KeyDown({
			param($src, $e)
			if ($null -eq $e) { return }
			if ([System.Windows.Forms.Keys]$e.KeyCode -eq [System.Windows.Forms.Keys]::Enter)
			{
				if (Test-DashboardCommand 'Add-MarketPriceFromUI') { Add-MarketPriceFromUI }
				$e.SuppressKeyPress = $true
			}
		})

	$marketCodeBox.Add_TextChanged({
			if ($null -eq $global:DashboardConfig.Config) { return }
			if (-not $global:DashboardConfig.Config.Contains('Market')) { $global:DashboardConfig.Config['Market'] = [ordered]@{} }
			$global:DashboardConfig.Config['Market']['AccessCode'] = $this.Text
		})
	$marketCodeBox.Add_Leave({ if (Test-DashboardCommand 'WriteConfig' -Verbose:$false) { WriteConfig } })

	$btnMarketUnlock.Add_Click({
			if (Test-DashboardCommand 'Invoke-MarketUnlockAction') { Invoke-MarketUnlockAction }
		})

	$btnMarketFind.Add_Click({
			if (Test-DashboardCommand 'Update-MarketSuggestions')
			{
				if (Test-DashboardCommand 'Get-MarketUIControl')
				{
					$tb = Get-MarketUIControl -Name 'MarketSearchInput'
					if ($tb) { try { [void]$tb.Focus() } catch { } }
				}
				Update-MarketSuggestions
				try { if (Test-DashboardCommand 'Update-MarketOfferGrid') { Update-MarketOfferGrid } } catch { }
			}
		})

	$btnMarketAdd.Add_Click({
			if (Test-DashboardCommand 'Add-MarketPriceFromUI') { Add-MarketPriceFromUI }
		})

	$ddlMarketCur.Add_SelectedIndexChanged({
			if (Test-DashboardCommand 'Set-MarketAddCurrency') { Set-MarketAddCurrency }
		})

	$ddlMarketRange.Add_SelectedIndexChanged({
			if (Test-DashboardCommand 'Set-MarketTrendRange') { Set-MarketTrendRange }
		})

	$uiPropertiesToAdd.MarketTab = $tabMarket
	$tabMarket.Add_VisibleChanged({
			param($src, $e)
			try
			{
				if (-not $src.Visible) { return }
				if (Test-DashboardCommand 'Open-MarketTab') { Open-MarketTab }
			}
			catch { }
		})
	$uiPropertiesToAdd.MarketTitle = $lblMarketTitle
	$uiPropertiesToAdd.MarketStatus = $lblMarketStatus
	$uiPropertiesToAdd.MarketAccessLabel = $lblMarketAccess
	$uiPropertiesToAdd.MarketCodeInput = $marketCodeBox
	$uiPropertiesToAdd.MarketUnlockBtn = $btnMarketUnlock
	$uiPropertiesToAdd.MarketScopeHint = $lblMarketScope
	$uiPropertiesToAdd.MarketSearchLabel = $lblMarketSearch
	$uiPropertiesToAdd.MarketSearchInput = $marketSearchBox
	$uiPropertiesToAdd.MarketFindBtn = $btnMarketFind
	$uiPropertiesToAdd.MarketSuggestions = $marketSuggest
	$uiPropertiesToAdd.MarketPriceLabel = $lblMarketPrice
	$uiPropertiesToAdd.MarketPriceInput = $marketPriceBox
	$uiPropertiesToAdd.MarketAddBtn = $btnMarketAdd
	$uiPropertiesToAdd.MarketCurCombo = $ddlMarketCur
	$uiPropertiesToAdd.MarketTrendLabel = $lblMarketTrend
	$uiPropertiesToAdd.MarketRangeLabel = $lblMarketRangeLbl
	$uiPropertiesToAdd.MarketTrendRange = $ddlMarketRange
	$uiPropertiesToAdd.MarketRangeUnit = $lblMarketRangeUnit
	$uiPropertiesToAdd.MarketTrendArrow = $lblMarketTrendArrow
	$uiPropertiesToAdd.MarketTrendText = $lblMarketTrendText
	$uiPropertiesToAdd.MarketTrendDetail = $lblMarketTrendDetail
	$uiPropertiesToAdd.MarketOnlineLabel = $lblMarketOnline
	$uiPropertiesToAdd.MarketHistoryLabel = $lblMarketHistory
	$uiPropertiesToAdd.MarketFeedback = $lblMarketFeedback
	$uiPropertiesToAdd.MarketHistoryCount = $lblMarketHistoryCount
	$uiPropertiesToAdd.MarketHistoryGrid = $marketGrid

	#endregion

	@( $btnBrowseLauncher, $btnBrowseJunction, $btnAddProfile, $btnRenameProfile, $btnRemoveProfile, $btnAddClient, $btnRenameClient, $btnRemoveClient, $btnAddSetup, $btnEditSetup, $btnRenameSetup, $btnMasterPauseHotkey, $btnMasterResumeHotkey, $btnRefreshBosses, $btnAddNote, $btnEditNote, $btnShowNotification, $btnHideAllNotifications, $btnResetUpgrade, $btnChestHotkey, $btnChestItemHotkey, $btnResetChestQueue, $btnMarketFind, $btnMarketUnlock ) | ForEach-Object {
		Register-ThemedButtonFadeEffect -Button $_ -Variant 'Normal'
	}

	@( $btnDeleteProfile, $btnDeleteSetup, $btnUnregisterHotkey, $btnCancel, $btnRemoveNote, $btnRemoveUpgrade, $btnRemoveChestQueue ) | ForEach-Object {
		Register-ThemedButtonFadeEffect -Button $_ -Variant 'Danger'
	}

	@( $btnStartJunction, $btnSave ) | ForEach-Object {
		Register-ThemedButtonFadeEffect -Button $_ -Variant 'Success'
	}

	@( $btnCopyNeuz, $btnAddUpgrade, $btnAddChestQueue, $btnMarketAdd ) | ForEach-Object {
		Register-ThemedButtonFadeEffect -Button $_ -Variant 'Primary'
	}

	$ExtraForm.Controls.AddRange(@(
			$lblTitle,
			$btnClose,
			$extraTabs
		))

	#endregion
	#endregion

	$global:DashboardConfig.UI = [PSCustomObject]@{
		MainForm                      = $mainForm
		SettingsForm                  = $settingsForm
		ExtraForm                     = $ExtraForm
		SettingsTabs                  = $settingsTabs
		TopBar                        = $topBar
		CloseForm                     = $btnCloseForm
		MinForm                       = $btnMinimizeForm
		InfoForm                      = $btnInfoForm
		ThemeToggle                   = $btnThemeToggle
		ResizeGrip                    = $resizeGrip
		DataGridMain                  = $DataGridMain
		LoginButton                   = $btnLogin
		Ftool                         = $btnFtool
		Macro                         = $btnMacro
		Extra                         = $btnExtra
		Wiki                          = $btnWiki
		Settings                      = $btnSettings
		Exit                          = $btnStop
		Launch                        = $btnLaunch
		LaunchContextMenu             = $LaunchContextMenu
		ToolTipMain                   = $toolTipMain
		ToolTipSettings               = $toolTipSettings
		ToolTipExtra                  = $toolTipExtra
		ToolTipFtool                  = $null
		InputLauncher                 = $txtLauncher
		InputJunction                 = $txtJunction
		InputUser                     = $userBox 
		StartJunction                 = $btnStartJunction
		InputProcess                  = $txtProcessName
		InputMax                      = $txtMaxClients
		Browse                        = $btnBrowseLauncher
		InputLauncherTimeout          = $txtLauncherTimeout
		BrowseJunction                = $btnBrowseJunction
		NeverRestartingCollectorLogin = $chkNeverRestartingLogin
		AutoStartWithWindows          = $chkAutoStartWindows
		AutoStartSetup                = $chkAutoStartSetup
		WorldBossListener             = $chkWorldBossListener
		RefreshBosses                 = $btnRefreshBosses
		ShowBossImages                = $chkShowImages
		ProfileGrid                   = $ProfileGrid
		ClientGrid                    = $ClientGrid
		AddClient                     = $btnAddClient
		RenameClient                  = $btnRenameClient
		RemoveClient                  = $btnRemoveClient
		HideProfiles                  = $chkHideProfiles
		AddProfile                    = $btnAddProfile
		ButtonPanel                   = $ButtonPanel
		RenameProfile                 = $btnRenameProfile
		RemoveProfile                 = $btnRemoveProfile
		DeleteProfile                 = $btnDeleteProfile
		SetupGrid                     = $SetupGrid
		AddSetup                      = $btnAddSetup
		EditSetup                     = $btnEditSetup
		RenameSetup                   = $btnRenameSetup
		DeleteSetup                   = $btnDeleteSetup
		CopyNeuz                      = $btnCopyNeuz
		LoginConfigGrid               = $LoginConfigGrid
		LoginPickers                  = $Pickers
		InputPostLoginDelay           = $txtPostLoginDelay
		LoginProfileSelector          = $ddlLoginProfile
		ManualCoordsOverride          = $chkManualCoords
		ResolutionSelector            = $ddlResolution
		DirectXSelector               = $ddlDirectX
		HotkeysGrid                   = $HotkeysGrid
		MasterPauseHotkeyBtn          = $btnMasterPauseHotkey
		MasterResumeHotkeyBtn         = $btnMasterResumeHotkey
		MasterHotkeysStatus           = $lblMasterHotkeysStatus
		UnregisterHotkey              = $btnUnregisterHotkey
		Save                          = $btnSave
		Cancel                        = $btnCancel
		NoteGrid                      = $NoteGrid
		NotificationGrid              = $NotificationGrid
		ShowNotification              = $btnShowNotification
		HideAllNotifications          = $btnHideAllNotifications
		AddNote                       = $btnAddNote
		EditNote                      = $btnEditNote
		RemoveNote                    = $btnRemoveNote
		ContextMenuFront              = $itmFront
		ContextMenuBack               = $itmBack
		ContextMenuResizeAndCenter    = $itmResizeCenter
		ContextMenuSavePos            = $itmSavePos
		ContextMenuLoadPos            = $itmLoadPos
		SetHotkeyShow                 = $itmSetHotkeyShow
		SetHotkeyHide                 = $itmSetHotkeyHide
		SetHotkeyLoadPos              = $itmSetHotkeyLoadPos
		Relog                         = $itmRelog
	}
	if ($null -ne $uiPropertiesToAdd)
	{
		$uiPropertiesToAdd.GetEnumerator() | ForEach-Object {
			$global:DashboardConfig.UI | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value -Force
		}
	}

	RegisterUIEventHandlers
	$script:CurrentBuilderToolTip = $null
	if (Test-DashboardCommand 'Set-Theme') { Set-Theme }
	if (Test-DashboardCommand 'Initialize-MarketTab') { Initialize-MarketTab }
	return $true
}

function Set-HotkeyFromGridRow
{
	[CmdletBinding()]
	param(
		[System.Windows.Forms.DataGridView]$Grid,
		[int]$RowIndex
	)

		if ($RowIndex -lt 0) { return }
		
		$row = $Grid.Rows[$RowIndex]
		if (-not $row.Tag -or ($row.Tag -is [System.Collections.IDictionary] -and $row.Tag.Contains('Header') -and $row.Tag.Header)) { return }
		if (-not $row.Tag.Id -or -not $row.Tag.OwnerKey) { return }
		
		$null = $id; $id = $row.Tag.Id
		$ownerKey = $row.Tag.OwnerKey
		$currentKeyString = $row.Cells[0].Value.ToString()
		if ($currentKeyString -match '\(Paused\)$') { $currentKeyString = $currentKeyString -replace ' \(Paused\)$', '' }

		$newKey = Show-KeyCaptureDialog -currentKey $currentKeyString -Owner $global:DashboardConfig.UI.MainForm
		
		if ($null -eq $newKey -or ($newKey -eq $currentKeyString)) { return }
		
		if ($ownerKey -match '^GroupHotkey(?:Hide|LoadPos)?_(.+)')
		{
			$groupName = $Matches[1]
			$keySection = 'Hotkeys'
			if ($ownerKey -match '^GroupHotkeyHide_') { $keySection = 'HotkeysHide' }
			elseif ($ownerKey -match '^GroupHotkeyLoadPos_') { $keySection = 'HotkeysLoadPos' }

			if (-not $global:DashboardConfig.Config.Contains($keySection)) { $global:DashboardConfig.Config[$keySection] = [ordered]@{} }
			$global:DashboardConfig.Config[$keySection][$groupName] = $newKey

			try { UnregisterHotkeyInstance -Id $id -OwnerKey $ownerKey } catch {}
		}
		elseif ($ownerKey -eq 'master_pause_hotkey' -or $ownerKey -eq 'master_resume_hotkey')
		{
			if (-not $global:DashboardConfig.Config.Contains('GlobalHotkeys')) { $global:DashboardConfig.Config['GlobalHotkeys'] = [ordered]@{} }
			$configName = if ($ownerKey -eq 'master_pause_hotkey') { 'MasterPauseHotkey' } else { 'MasterResumeHotkey' }
			$global:DashboardConfig.Config['GlobalHotkeys'][$configName] = $newKey

			try { UnregisterHotkeyInstance -Id $id -OwnerKey $ownerKey } catch {}
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
					$cmd = Get-QuickCommandById -Data $data -CmdId $cmdId
					if ($cmd)
					{
						if (-not $global:DashboardConfig.Config.Contains('ChatCommander')) { $global:DashboardConfig.Config['ChatCommander'] = [ordered]@{} }
						$cmd.Hotkey = $newKey
						$cmd.HotkeyId = $null
						try { UnregisterHotkeyInstance -Id $id -OwnerKey "quickcmd_$($instanceId)_$($cmdId)" } catch {}
						try { UnregisterHotkeyInstance -Id $id -OwnerKey $instanceId } catch {}
						if ($newKey)
						{
							$scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($instanceId)' -CommandId '$($cmdId)'")
							$ownerKeyNew = "quickcmd_$($instanceId)_$($cmdId)"
							$ownerLabel = $null
							if ($cmd.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($cmd.Name)) { $ownerLabel = $cmd.Name }
							elseif ($cmd.PSObject.Properties['CommandText'] -and -not [string]::IsNullOrEmpty($cmd.CommandText)) { $ownerLabel = $cmd.CommandText }
							else { $ownerLabel = "Invoke Command: $cmdId" }
							try { $cmd.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKeyNew -OwnerLabel $ownerLabel -OldHotkeyId $id } catch {}
						}
						if (Test-DashboardCommand 'UpdateChatCommanderSettings') { UpdateChatCommanderSettings $data -forceWrite }
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
					if ($newKey)
					{
						$script = [scriptblock]::Create("ToggleMacroInstance -InstanceId '$($instanceId)'")
						try { $data.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $script -OwnerKey "macroinst_$($instanceId)" -OldHotkeyId $id } catch {}
					}
					if (Test-DashboardCommand 'UpdateMacroSettings') { UpdateMacroSettings $data -forceWrite }
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
					try
					{
						$data.GlobalHotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKey -OldHotkeyId $oldId
						if (Test-DashboardCommand 'UpdateSettings') { UpdateSettings $data -forceWrite }
					}
					catch
					{
						Show-DarkMessageBox "Failed to update hotkey: $_" 'Error' 'OK' 'Error'
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
				try
				{
					$extData.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKey -OldHotkeyId $oldId
					$form = $extData.Panel.FindForm()
					if ($form -and $form.Tag) { if (Test-DashboardCommand 'UpdateSettings') { UpdateSettings $form.Tag $extData -forceWrite } }
				}
				catch
				{
					Show-DarkMessageBox "Failed to update hotkey: $_" 'Error' 'OK' 'Error'
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
				try
				{
					$data.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKey -OldHotkeyId $oldId
					if (Test-DashboardCommand 'UpdateSettings') { UpdateSettings $data -forceWrite }
				}
				catch
				{
					Show-DarkMessageBox "Failed to update hotkey: $_" 'Error' 'OK' 'Error'
				}
			}
		}
		
		if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
		SyncConfigToUI
		RegisterConfiguredHotkeys
		RefreshHotkeysList
}


function RegisterUIEventHandlers
{
	[CmdletBinding()]
	param()

	if ($null -eq $global:DashboardConfig.UI) { return }

	$eventMappings = @{
		MainForm                     = @{
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
						if (-not $global:DashboardConfig.Config.Contains('WindowPosition'))
      {
							$global:DashboardConfig.Config['WindowPosition'] = [ordered]@{}
						}
						$global:DashboardConfig.Config['WindowPosition']['MainFormX'] = $src.Location.X
						$global:DashboardConfig.Config['WindowPosition']['MainFormY'] = $src.Location.Y
						$global:DashboardConfig.Config['WindowPosition']['MainFormHeight'] = $src.Height
						if ($global:DashboardConfig.UI.DataGridMain) { $global:DashboardConfig.Config['WindowPosition']['DataGridHeight'] = $global:DashboardConfig.UI.DataGridMain.Height }
						
						if ($global:DashboardConfig.UI.ExtraForm -and -not $global:DashboardConfig.UI.ExtraForm.IsDisposed)
      {
							$ef = $global:DashboardConfig.UI.ExtraForm
							if ($ef.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal)
							{
								$global:DashboardConfig.Config['WindowPosition']['ExtraFormX'] = $ef.Location.X; $global:DashboardConfig.Config['WindowPosition']['ExtraFormY'] = $ef.Location.Y
							}
						}
						
						if ($global:DashboardConfig.UI.SettingsForm -and -not $global:DashboardConfig.UI.SettingsForm.IsDisposed)
      {
							$sf = $global:DashboardConfig.UI.SettingsForm
							if ($sf.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal)
							{
								$global:DashboardConfig.Config['WindowPosition']['SettingsFormX'] = $sf.Location.X; $global:DashboardConfig.Config['WindowPosition']['SettingsFormY'] = $sf.Location.Y
							}
						}
						Write-Verbose "Saving Window Position: X=$($src.Location.X), Y=$($src.Location.Y), H=$($src.Height)"
						if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
					}
				}
				catch
				{
				}
				if (Test-DashboardCommand 'RemoveAllHotkeys') { RemoveAllHotkeys }
				if (Test-DashboardCommand 'StopDashboard') { StopDashboard }
			}
		}
		SettingsForm                 = @{
			MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.SettingsForm.Handle, 0xA1, 0x2, 0) }
			Load = { 
				SyncConfigToUI; RegisterConfiguredHotkeys; RefreshHotkeysList
				if (-not ('Custom.TabStripDragHelper' -as [Type])) {
					Write-Verbose 'TabStripDragHelper type not found. Skipping tab strip drag helper.'
				} else {
					$global:DashboardConfig.State.SettingsDragHelper = New-Object Custom.TabStripDragHelper(
						$global:DashboardConfig.UI.SettingsForm.Handle,
						$global:DashboardConfig.UI.SettingsTabs
					)
				}
			}
		}
		ExtraForm                    = @{
			MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.ExtraForm.Handle, 0xA1, 0x2, 0) }
		}
		#region MainForm
		TopBar                       = @{ MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.MainForm.Handle, 0xA1, 0x2, 0) } }
		InfoForm                     = @{ MouseDown = { 
				$cms = New-Object System.Windows.Forms.ContextMenuStrip
				if ('Custom.DarkRenderer' -as [Type]) { $cms.Renderer = New-Object Custom.DarkRenderer }
				$itmWeb = $cms.Items.Add('Open Website')
				$itmWeb.Image = if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon))
				{ 
					try
					{
						$icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon)
						$bmp = $icon.ToBitmap()
						$icon.Dispose()
						$bmp
					}
					catch { $null }
				}
				else { $null }
				$itmWeb.add_Click({ Start-Process 'https://immortal-divine.github.io/Entropia_Dashboard/' })
				
				$itmGuide = $cms.Items.Add('Start Interactive Guide')
				try
				{
					$guideBmp = New-Object System.Drawing.Bitmap(16, 16)
					$g = [System.Drawing.Graphics]::FromImage($guideBmp)
					$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
					$g.FillEllipse([System.Drawing.Brushes]::MediumSeaGreen, 0, 0, 15, 15)
					$sf = New-Object System.Drawing.StringFormat
					$sf.Alignment = [System.Drawing.StringAlignment]::Center
					$sf.LineAlignment = [System.Drawing.StringAlignment]::Center
					$g.DrawString('?', (Get-UIFont 'Segoe UI' 9 'Bold'), [System.Drawing.Brushes]::White, 8, 8, $sf)
					$g.Dispose()
					$itmGuide.Image = $guideBmp
				}
				catch {}
				$itmGuide.add_Click({ if (Test-DashboardCommand 'Show-Guide') { Show-Guide } })
				
				$itmKofi = $cms.Items.Add('Support me on Ko-fi')
				$kofiPath = Join-Path $global:DashboardConfig.Paths.App 'kofi.png'
				if (-not (Test-Path $kofiPath))
				{
					try
					{
						[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
						$wc = New-Object System.Net.WebClient
						$wc.DownloadFile('https://storage.ko-fi.com/cdn/cup-border.png', $kofiPath)
						$wc.Dispose()
					}
					catch {}
				}
				if (Test-Path $kofiPath)
				{
					try { $itmKofi.Image = [System.Drawing.Image]::FromFile($kofiPath) } catch {}
				}
				$itmKofi.add_Click({ Start-Process 'https://ko-fi.com/U7U61S0EGT' })
				
				$itmPayPal = $cms.Items.Add('Support me with PayPal')
				$paypalPath = Join-Path $global:DashboardConfig.Paths.App 'paypal.png'
				if (-not (Test-Path $paypalPath))
				{
					try
					{
						[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
						$wc = New-Object System.Net.WebClient
						$wc.DownloadFile('https://www.paypalobjects.com/webstatic/icon/pp258.png', $paypalPath)
						$wc.Dispose()
					}
					catch {}
				}
				if (Test-Path $paypalPath)
				{
					try { $itmPayPal.Image = [System.Drawing.Image]::FromFile($paypalPath) } catch {}
				}
				$itmPayPal.add_Click({ Start-Process 'https://ko-fi.com/U7U61S0EGT' })
				
				$cms.Show($Sender, 0, $Sender.Height)
			} 
		}
		MinForm                      = @{ Click = { $global:DashboardConfig.UI.MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized } }
		CloseForm                    = @{ Click = { $global:DashboardConfig.UI.MainForm.Close() } }
		ThemeToggle                  = @{ Click = { Show-ThemeSelectionForm } }

		Launch                       = @{
			Click = {
				if ($global:DashboardConfig.State.LaunchActive)
				{
					try { StopClientLaunch } catch { Show-DarkMessageBox $_.Exception.Message }
				}
				else
				{
					$Sender.ContextMenuStrip.Show($Sender, 0, $Sender.Height)
				}
			}
		}
		LaunchContextMenu            = @{
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

				$launchLabel = "Launch $($howManyUntilMax)x '$($selectedProfileName)'"
				$selectedClientName = Get-SelectedClientForProfile -ProfileName $selectedProfileName
				if (-not [string]::IsNullOrWhiteSpace($selectedClientName))
				{
					$launchLabel = "$launchLabel - $selectedClientName"
				}

				$header = $s.Items.Add($launchLabel)
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
				try { ApplyControlInternal $s (Get-Theme) } catch { }
			}
		}
		LoginButton                  = @{
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
		Settings                     = @{ Click = { ShowSettingsForm; RegisterConfiguredHotkeys; RefreshHotkeysList } }
		Exit                         = @{
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
		Ftool                        = @{ 
			MouseDown = {
				param($s, $e)
				if ($e.Button -eq 'Left')
				{
					Write-Verbose 'Ftool: Left Click'
					if (Test-DashboardCommand 'FtoolSelectedRow')
					{
						$global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { FtoolSelectedRow $_ }
					}
					else
					{
						Write-Error 'Ftool: FtoolSelectedRow command not found!'
					}
				}
				elseif ($e.Button -eq 'Right')
				{
					Write-Verbose 'Ftool: Right Click'
					try
					{
						$selectedRows = $global:DashboardConfig.UI.DataGridMain.SelectedRows
						if ($selectedRows.Count -eq 0) { return }
						$row = $selectedRows[0]
						$rowIdentity = Get-RowIdentity -Row $row
						$safeIdentity = $rowIdentity -replace '[^a-zA-Z0-9_\-]', '_'
						$sessionKey = "Session_Ftool_$safeIdentity"

						if ($null -eq $global:DashboardConfig.Config['Ftool'] -or $global:DashboardConfig.Config['Ftool'] -is [string]) { $global:DashboardConfig.Config['Ftool'] = [ordered]@{} }
						if ($null -eq $global:DashboardConfig.Config['Macro'] -or $global:DashboardConfig.Config['Macro'] -is [string]) { $global:DashboardConfig.Config['Macro'] = [ordered]@{} }

						$activeForms = [System.Collections.ArrayList]@()
						if ($global:DashboardConfig.Resources.FtoolForms)
						{
							foreach ($f in $global:DashboardConfig.Resources.FtoolForms.Values)
							{
								$isFtool = $false
								if ($f -and -not $f.IsDisposed -and $f.Tag -and $f.Tag -isnot [string])
								{
									if ($null -eq $f.Tag.Type -or $f.Tag.Type -ne 'Macro')
									{
										if ($f.Tag.SelectedWindow -eq $row.Tag.MainWindowHandle)
										{
											$isFtool = $true
										}
									}
								}
                                
								if ($isFtool)
								{
									$activeForms.Add($f) | Out-Null
									Write-Verbose "Ftool Debug: Captured form $($f.Name)"
								}
							}
						}

						if ($activeForms.Count -gt 0)
						{
							$global:DashboardConfig.Config['Ftool'][$sessionKey] = $activeForms.Count
							for ($i = 0; $i -lt $activeForms.Count; $i++)
							{
								$f = $activeForms[$i]
								$data = $f.Tag
								$configKey = "${sessionKey}_$i"
								if (-not $global:DashboardConfig.Config['Ftool'].Contains($configKey) -or $global:DashboardConfig.Config['Ftool'][$configKey] -isnot [System.Collections.IDictionary]) { $global:DashboardConfig.Config['Ftool'][$configKey] = [ordered]@{} }
								$config = $global:DashboardConfig.Config['Ftool'][$configKey]
								$config['key'] = if ($data.BtnKeySelect) { $data.BtnKeySelect.Text } else { '' }
								$config['interval'] = if ($data.Interval) { $data.Interval.Text } else { '' }
								$config['name'] = if ($data.Name) { $data.Name.Text } else { '' }
                                
								$instanceId = $data.InstanceId
								$instanceKey = "instance_$instanceId"
								if ($global:DashboardConfig.Resources.ExtensionTracking -and $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
								{
									$activeExtensions = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions
									$config['ExtCount'] = $activeExtensions.Count
									for ($j = 0; $j -lt $activeExtensions.Count; $j++)
									{
										$extKey = $activeExtensions[$j]
										if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
										{
											$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
											$extConfig = [ordered]@{}
											$extConfig['key'] = if ($extData.BtnKeySelect) { $extData.BtnKeySelect.Text } else { '' }
											$extConfig['interval'] = if ($extData.Interval) { $extData.Interval.Text } else { '' }
											$extConfig['name'] = if ($extData.Name) { $extData.Name.Text } else { '' }
											$extConfig['hotkey'] = if ($extData.Hotkey) { $extData.Hotkey } else { '' }
											$config["Ext_$j"] = $extConfig
										}
									}
								}
							}
							WriteConfig
                            
							foreach ($f in $activeForms)
							{
								$f.Close()
							}
							Show-DarkMessageBox "Saved $($activeForms.Count) active Ftool instances for PID $pid." 'Session Saved' 'OK' 'Information' 'success'
						}
						else
						{
							if ($global:DashboardConfig.Config['Ftool'].Contains($sessionKey))
       {
								$savedValue = $global:DashboardConfig.Config['Ftool'][$sessionKey]
								$count = [int]$savedValue
								for ($i = 0; $i -lt $count; $i++)
								{
									$configKey = "${sessionKey}_$i"
									$config = $null
									if ($global:DashboardConfig.Config['Ftool'].Contains($configKey))
									{
										$configRaw = $global:DashboardConfig.Config['Ftool'][$configKey]
										if ($configRaw -is [string] -and $configRaw.StartsWith('{'))
										{
											$config = $configRaw | ConvertFrom-Json
										}
										else { $config = $configRaw }
									}
									FtoolSelectedRow $row -ForceNew -Config $config -InstanceIndex ($i + 1)
								}
								Show-DarkMessageBox "Restored $count Ftool instances for $($row.Cells[1].Value)." 'Session Restored' 'OK' 'Information' 'success'
							}
							else
							{
								Show-DarkMessageBox 'No active Ftool instances to close and no saved session to restore.' 'No Action' 'OK' 'Information'
							}
						}
					}
					catch
					{
						Show-DarkMessageBox "Error handling Ftool session: $($_.Exception.Message)" 'Session Error' 'OK' 'Error'
					}
				}
			} 
		}
		Macro                        = @{ 
			MouseDown = {
				param($s, $e)
				if ($e.Button -eq 'Left')
				{
					if (Test-DashboardCommand 'MacroSelectedRow') { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { MacroSelectedRow $_ } }
				}
				elseif ($e.Button -eq 'Right')
				{
					try
					{
						$selectedRows = $global:DashboardConfig.UI.DataGridMain.SelectedRows
						if ($selectedRows.Count -eq 0) { return }
						$row = $selectedRows[0]

						$rowIdentity = Get-RowIdentity -Row $row
						$safeIdentity = $rowIdentity -replace '[^a-zA-Z0-9_\-]', '_'
						$sessionKey = "Session_Macro_Macro_$safeIdentity"
						Write-Verbose "Macro Right-Click: Identity=$rowIdentity, SessionKey=$sessionKey"

						if ([string]::IsNullOrWhiteSpace($sessionKey) -or $rowIdentity -eq '')
						{
							Show-DarkMessageBox 'Failed to generate a valid session key. Row data missing.' 'Session Error' 'OK' 'Error'
							return
						}

						if ($null -eq $global:DashboardConfig.Config['Ftool'] -or $global:DashboardConfig.Config['Ftool'] -is [string]) { $global:DashboardConfig.Config['Ftool'] = [ordered]@{} }
						if ($null -eq $global:DashboardConfig.Config['Macro'] -or $global:DashboardConfig.Config['Macro'] -is [string]) { $global:DashboardConfig.Config['Macro'] = [ordered]@{} }

						$activeForms = [System.Collections.ArrayList]@()
						if ($global:DashboardConfig.Resources.FtoolForms)
						{
							foreach ($k in $global:DashboardConfig.Resources.FtoolForms.Keys)
							{
								if ($k -eq "Macro_$safeIdentity" -or $k -match "^Macro_$([regex]::Escape($safeIdentity))_\d+$")
								{
									$f = $global:DashboardConfig.Resources.FtoolForms[$k]
									if ($f -and -not $f.IsDisposed -and $f.Tag -and $f.Tag -isnot [string])
									{
										$type = if ($f.Tag.PSObject.Properties['Type']) { $f.Tag.Type } else { $null }
										if ($type -eq 'Macro')
										{
											$activeForms.Add($f) | Out-Null
										}
									}
								}
							}
						}

						if ($activeForms.Count -gt 0)
						{
							$global:DashboardConfig.Config['Macro'][$sessionKey] = $activeForms.Count
							WriteConfig

							foreach ($f in $activeForms)
							{
								$f.Close()
							}
							Show-DarkMessageBox "Saved $($activeForms.Count) active Macro instances for $rowIdentity." 'Session Saved' 'OK' 'Information' 'success'
						}
						else
						{
							if ($global:DashboardConfig.Config['Macro'].Contains($sessionKey))
       {
								$count = [int]$global:DashboardConfig.Config['Macro'][$sessionKey]
								for ($i = 1; $i -le $count; $i++)
								{
									MacroSelectedRow $row -ForceNew -InstanceIndex $i
								}
								Show-DarkMessageBox "Restored $count Macro instances for $rowIdentity." 'Session Restored' 'OK' 'Information' 'success'
							}
							else
							{
								Show-DarkMessageBox 'No active Macro instances to close and no saved session to restore.' 'No Action' 'OK' 'Information'
							}
						}
					}
					catch
					{
						Show-DarkMessageBox "Error handling Macro session: $($_.Exception.Message)" 'Session Error' 'OK' 'Error'
					}
				}
			} 
		}
		Extra                        = @{ MouseDown = {
				param($s,$e)
				if ($e.Button -eq 'Left')
				{
					ShowExtraForm
					if (Test-DashboardCommand 'RefreshNoteGrid') { RefreshNoteGrid }
					if (Test-DashboardCommand 'RefreshNotificationGrid') { RefreshNotificationGrid }
					if (Test-DashboardCommand 'Initialize-ChestOpener') { Initialize-ChestOpener }
					if (Test-DashboardCommand 'Update-ChestOpenerWindowList') { Update-ChestOpenerWindowList }
					if (Test-DashboardCommand 'Initialize-MarketTab') { Initialize-MarketTab }
				}
				if ($e.Button -eq 'Right')
				{
					if (Test-DashboardCommand 'ChatCommanderSelectedRow') { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { ChatCommanderSelectedRow $_ } }
				}
			} 
		}
		Wiki                         = @{ MouseDown = { 
			param($s,$e)
			if ($e.Button -eq 'Left')
			{
				try
				{
					$atlasCommand = Get-Command Show-Atlas -ErrorAction Stop -Verbose:$False
					& $atlasCommand | Out-Null
				}
				catch
				{
					$errorMessage = "Node Maps action failed.`n`nCould not find or execute the 'Show-Atlas' function. The 'atlas.psm1' module may have failed to load correctly.`n`nTechnical Details: $($_.Exception.Message)"
					try { ShowErrorDialog -Message $errorMessage }
					catch { Show-DarkMessageBox $errorMessage 'Node Maps Error' 'Ok' 'Error' }
				}
			}
			if ($e.Button -eq 'Right')
			{
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
		}
		DataGridMain                 = @{
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
		ContextMenuFront             = @{ Click = { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { $h = $_.Tag.MainWindowHandle; SetWindowToolStyle -hWnd $h -Hide $false; [Custom.Native]::BringToFront($h); SetWindowToolStyle -hWnd $h -Hide $false } } }                
		ContextMenuBack              = @{ Click = { $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { [Custom.Native]::SendToBack($_.Tag.MainWindowHandle); [Custom.Native]::EmptyWorkingSet($_.Tag.Id) } } }
		ContextMenuResizeAndCenter   = @{ Click = { $scr = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea; $global:DashboardConfig.UI.DataGridMain.SelectedRows | ForEach-Object { [Custom.Native]::PositionWindow($_.Tag.MainWindowHandle, [Custom.Native]::TopWindowHandle, [int](($scr.Width - 1040) / 2), [int](($scr.Height - 807) / 2), 1040, 807, 0x0010) } } }
		ContextMenuSavePos           = @{ Click = { if (Test-DashboardCommand 'Save-WindowPositions') { Save-WindowPositions } } }
		ContextMenuLoadPos           = @{ Click = { if (Test-DashboardCommand 'Restore-WindowPositions') { Restore-WindowPositions } } }
		Relog                        = @{
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
		HotkeysGrid                  = @{
			CellDoubleClick = {
				param($s, $e)
				$script:HotkeysGridEventDepth++
				try
				{
					Set-HotkeyFromGridRow -Grid $s -RowIndex $e.RowIndex
				}
				finally
				{
					$script:HotkeysGridEventDepth--
					if ($script:HotkeysGridEventDepth -eq 0) {
						if ($script:HotkeysGridMouseActive) {
							Start-HotkeysGridSafetyTimer
						}
						else {
							RefreshHotkeysList
						}
					}
				}
			}
		}
		SetHotkeyShow                = @{
			Click = { Set-HotkeyGroupBinding -KeySection 'Hotkeys' -GroupSection 'HotkeyGroups' -DialogTitle 'Hotkey Group (Show)' -MessageTitle 'Hotkeys Created' }
		}
		SetHotkeyHide                = @{
			Click = { Set-HotkeyGroupBinding -KeySection 'HotkeysHide' -GroupSection 'HotkeyGroupsHide' -DialogTitle 'Hotkey Group (Hide)' -MessageTitle 'Hotkeys (Hide) Created' }
		}
		SetHotkeyLoadPos             = @{
			Click = { Set-HotkeyGroupBinding -KeySection 'HotkeysLoadPos' -GroupSection 'HotkeyGroupsLoadPos' -DialogTitle 'Hotkey Group (Load Pos)' -MessageTitle 'Hotkeys (Load Pos) Created' }
		}
		ResizeGrip                   = @{
			MouseDown = {
				param($s, $e)
				if ($e.Button -eq 'Left')
				{
					[Custom.Native]::ReleaseCapture()
					[Custom.Native]::SendMessage($global:DashboardConfig.UI.MainForm.Handle, 0x112, 0xF006, 0)
				}
				if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
			}
		}
		#endregion
		#region SettingsForm
		settingsTabs                 = @{ 
			MouseDown = { param($src, $e); [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($global:DashboardConfig.UI.SettingsForm.Handle, 0xA1, 0x2, 0) } 
		}
		Browse                       = @{ 
			Click = { $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Executable Files (*.exe)|*.exe|All Files (*.*)|*.*'; if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $global:DashboardConfig.UI.InputLauncher.Text = $d.FileName } } 
		}
		BrowseJunction               = @{
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
		InputJunction                = @{
			Leave = {
				if ([string]::IsNullOrWhiteSpace($global:DashboardConfig.UI.InputJunction.Text))
				{
					$global:DashboardConfig.UI.InputJunction.Text = $global:DashboardConfig.Paths.Profiles
				}
			}
		}
		StartJunction                = @{
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
		CopyNeuz                     = @{ 
			Click = {
				$confirm = Show-DarkMessageBox "Do you want to copy the complete 'bin32' and 'bin64' folders of your main client to the profiles selected in the list above (all profiles if none is selected)?`nProfiles that were not created through junctions or have a different game folder might break.`n`nFiles that cannot be overwritten are skipped and reported - the copy continues.`n`nThis action is only working if you patched your Main Launcher folder first!`nProcced?" 'Confirm Folder Copy' 'YesNo' 'Warning'

				if ($confirm -eq 'Yes')
				{
					$mainLauncherPath = $global:DashboardConfig.UI.InputLauncher.Text
					if ([string]::IsNullOrWhiteSpace($mainLauncherPath) -or -not (Test-Path $mainLauncherPath))
					{
						Show-DarkMessageBox 'Please select a valid Main Launcher Path first.' 'Error' 'OK' 'Error'
						return
					}

					$mainClientDir = Split-Path $mainLauncherPath -Parent

					$sourceBins = @()
					foreach ($binName in @('bin32', 'bin64'))
					{
						if (Test-Path -LiteralPath (Join-Path $mainClientDir $binName)) { $sourceBins += $binName }
					}
					if ($sourceBins.Count -eq 0)
					{
						Show-DarkMessageBox "No 'bin32' or 'bin64' folder was found in '$mainClientDir'. Patch the Main Launcher folder first." 'Nothing to copy' 'OK' 'Error'
						return
					}

					$profiles = $global:DashboardConfig.Config['Profiles']
					if (-not $profiles -or $profiles.Count -eq 0)
					{
						Show-DarkMessageBox 'No profiles found. Please create or add profiles first.' 'No Profiles found' 'OK' 'Information'
						return
					}

					$targets = @()
					foreach ($gridRow in $global:DashboardConfig.UI.ProfileGrid.SelectedRows)
					{
						if ($null -eq $gridRow.Cells[0].Value -or $null -eq $gridRow.Cells[1].Value) { continue }
						$targets += [pscustomobject]@{ Name = $gridRow.Cells[0].Value.ToString(); Path = $gridRow.Cells[1].Value.ToString() }
					}

					$targetNote = 'the selected profiles'
					if ($targets.Count -eq 0)
					{
						$targetNote = 'all profiles'
						foreach ($profileName in $profiles.Keys)
						{
							$targets += [pscustomobject]@{ Name = $profileName; Path = [string]$profiles[$profileName] }
						}
					}
					if ($targets.Count -eq 0)
					{
						Show-DarkMessageBox 'No profiles to copy to.' 'No Profiles found' 'OK' 'Information'
						return
					}

					$settingsForm = $global:DashboardConfig.UI.SettingsForm
					try { $settingsForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor } catch { }

					try
					{
						$copyResult = Copy-BinFoldersToProfiles -SourceDir $mainClientDir -Targets $targets
					}
					finally
					{
						try { $settingsForm.Cursor = [System.Windows.Forms.Cursors]::Default } catch { }
					}

					$summary = "Copied $($copyResult.Copied) file(s) to $($copyResult.Profiles) profile(s): $targetNote.`n$($copyResult.Unchanged) file(s) were already up to date."

					if ($copyResult.MissingFolders.Count -gt 0)
					{
						$missingPreview = ($copyResult.MissingFolders | Select-Object -First 5) -join "`n"
						$summary += "`n`nSkipped profiles:`n$missingPreview"
					}

					if ($copyResult.Skipped.Count -gt 0)
					{
						$skipPreview = ($copyResult.Skipped | Select-Object -First 6) -join "`n"
						if ($copyResult.Skipped.Count -gt 6) { $skipPreview += "`n... and $($copyResult.Skipped.Count - 6) more" }
						$summary += "`n`n$($copyResult.Skipped.Count) file(s) could not be overwritten and were skipped, the copy continued:`n$skipPreview"
						Show-DarkMessageBox $summary 'Copy finished with skipped files' 'OK' 'Warning'
					}
					else
					{
						Show-DarkMessageBox $summary 'Patch complete!' 'OK' 'Information' 'success'
					}
				}
			}
		}
		ProfileGrid                  = @{
			CellClick       = {
				param($s, $e)
				if ($e.RowIndex -lt 0) { return }


				$s.Tag = $e.RowIndex

				$row = $s.Rows[$e.RowIndex]
				$profileName = [string]$row.Cells[0].Value
				if (-not [string]::IsNullOrWhiteSpace($profileName) -and $global:DashboardConfig.Config.Contains('Options'))
				{
					$global:DashboardConfig.Config['Options']['SelectedProfile'] = $profileName
					WriteConfig
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
		ClientGrid                    = @{
			CellClick       = {
				param($s, $e)
				if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
				if ($global:DashboardConfig.State.ClientGridPopulating) { return }

				if ($e.ColumnIndex -ge 2)
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

					SyncClientGridToConfig
					WriteConfig

					if ($e.ColumnIndex -ge 4)
					{
						Reset-AutoToolStateForClient -ProfileName ([string]$row.Cells[0].Value) -ClientName ([string]$row.Cells[1].Value)
					}
				}
				else
				{
					$s.Tag = $e.RowIndex

					$selRow = $s.Rows[$e.RowIndex]
					$selProfile = [string]$selRow.Cells[0].Value
					$selClient = [string]$selRow.Cells[1].Value
					if (-not [string]::IsNullOrWhiteSpace($selProfile) -and -not [string]::IsNullOrWhiteSpace($selClient))
					{
						Set-SelectedClientForProfile -ProfileName $selProfile -ClientName $selClient | Out-Null
						if ($global:DashboardConfig.Config.Contains('Options'))
						{
							$global:DashboardConfig.Config['Options']['SelectedProfile'] = $selProfile
						}
						WriteConfig
					}
				}
			}
		}
		AddProfile                   = @{
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
		RenameProfile                = @{
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

						Rename-RememberedClientsProfile -OldProfile $oldName -NewProfile $newName

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
		RemoveProfile                = @{
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

					Remove-RememberedClientsProfile -ProfileName $profileName
					$global:DashboardConfig.UI.ProfileGrid.Rows.Remove($row)
				}
				SyncProfilesToConfig
				RefreshLoginProfileSelector
				WriteConfig
			}
		}
		AddClient                     = @{
			Click = {
				$UI = $global:DashboardConfig.UI
				$profiles = $global:DashboardConfig.Config['Profiles']
				if (-not $profiles -or @($profiles.Keys).Count -eq 0)
				{
					Show-DarkMessageBox 'Create a profile first.' 'No Profiles' 'OK' 'Warning'
					return
				}

				$profileName = ''
				if ($UI.ClientGrid.SelectedRows.Count -gt 0 -and $null -ne $UI.ClientGrid.SelectedRows[0].Cells[0].Value)
				{
					$profileName = $UI.ClientGrid.SelectedRows[0].Cells[0].Value.ToString()
				}
				if ([string]::IsNullOrWhiteSpace($profileName) -and $UI.ProfileGrid.SelectedRows.Count -gt 0 -and $null -ne $UI.ProfileGrid.SelectedRows[0].Cells[0].Value)
				{
					$profileName = $UI.ProfileGrid.SelectedRows[0].Cells[0].Value.ToString()
				}
				if ([string]::IsNullOrWhiteSpace($profileName)) { $profileName = @($profiles.Keys)[0] }
				if (-not $profiles.Contains($profileName)) { $profileName = @($profiles.Keys)[0] }

				$typed = ShowInputBox -Title 'Add Client' -Prompt "Window title of the client to remember under '$profileName':" -DefaultText ''
				if ([string]::IsNullOrWhiteSpace($typed)) { return }
				$clientName = $typed.Trim()

				if (Get-RememberedClientKey -ProfileName $profileName -ClientName $clientName)
				{
					Show-DarkMessageBox "'$clientName' is already remembered for '$profileName'." 'Already Added' 'OK' 'Warning'
					return
				}

				if (-not (Set-RememberedClientFlags -ProfileName $profileName -ClientName $clientName)) { return }
				SyncConfigToClientGrid | Out-Null
				WriteConfig
				Reset-AutoToolStateForClient -ProfileName $profileName -ClientName $clientName

				foreach ($row in $UI.ClientGrid.Rows)
				{
					if ([string]$row.Cells[0].Value -eq $profileName -and [string]$row.Cells[1].Value -eq $clientName)
					{
						try
						{
							$UI.ClientGrid.ClearSelection()
							$row.Selected = $true
							$UI.ClientGrid.CurrentCell = $row.Cells[1]
							$UI.ClientGrid.Tag = $row.Index
						}
						catch {}
						break
					}
				}
			}
		}
		RenameClient                  = @{
			Click = {
				$UI = $global:DashboardConfig.UI
				$rows = @($UI.ClientGrid.SelectedRows)
				if ($rows.Count -eq 0 -or $null -eq $rows[0].Cells[1].Value)
				{
					Show-DarkMessageBox 'Please select a client first.' 'Select Client' 'OK' 'Warning'
					return
				}

				$row = $rows[0]
				$profileName = [string]$row.Cells[0].Value
				$oldName = [string]$row.Cells[1].Value

				$typed = ShowInputBox -Title 'Rename Client' -Prompt "New window title for '$oldName':" -DefaultText $oldName
				if ([string]::IsNullOrWhiteSpace($typed)) { return }
				$newName = $typed.Trim()
				if ($newName -eq $oldName) { return }

				if (Get-RememberedClientKey -ProfileName $profileName -ClientName $newName)
				{
					Show-DarkMessageBox "'$newName' is already remembered for '$profileName'." 'Already Exists' 'OK' 'Warning'
					return
				}

				Rename-RememberedClient -ProfileName $profileName -OldName $oldName -NewName $newName | Out-Null
				SyncConfigToClientGrid | Out-Null
				WriteConfig
				Reset-AutoToolStateForClient -ProfileName $profileName -ClientName $newName
			}
		}
		RemoveClient                  = @{
			Click = {
				$UI = $global:DashboardConfig.UI
				$rows = @($UI.ClientGrid.SelectedRows)
				if ($rows.Count -eq 0 -or $null -eq $rows[0].Cells[1].Value)
				{
					Show-DarkMessageBox 'Please select a client first.' 'Select Client' 'OK' 'Warning'
					return
				}

				foreach ($row in $rows)
				{
					$profileName = [string]$row.Cells[0].Value
					$clientName = [string]$row.Cells[1].Value
					Remove-RememberedClient -ProfileName $profileName -ClientName $clientName | Out-Null
				}
				SyncConfigToClientGrid | Out-Null
				WriteConfig
			}
		}
		DeleteProfile                = @{
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
							Remove-RememberedClientsProfile -ProfileName $profileName
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
		SetupGrid                    = @{
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
		EditSetup                    = @{
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
					if (Test-DashboardCommand 'WriteConfig') { WriteConfig }

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
							$colRemove.DefaultCellStyle.BackColor = Get-Theme 'ButtonBg'
							$colRemove.DefaultCellStyle.ForeColor = Get-Theme 'WhiteText'
							$colRemove.DefaultCellStyle.SelectionBackColor = Get-Theme 'ButtonBg'
							$colRemove.DefaultCellStyle.SelectionForeColor = Get-Theme 'WhiteText'
							$colRemove.DefaultCellStyle.Font = (Get-UIFont 'Segoe UI' 8)

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
		AddSetup                     = @{
			Click = {
				
				$grid = $global:DashboardConfig.UI.DataGridMain
				$selectedRows = $grid.SelectedRows
				if ($selectedRows.Count -eq 0)
				{
					Show-DarkMessageBox 'Please select the clients you want to include in the setup.' 'No Selection' 'OK' 'Warning'
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
				if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
				SyncConfigToUI

				Show-DarkMessageBox "Setup '$setupName' saved successfully with $($detailedStateList.Count) clients." 'Setup Saved' 'OK' 'Information' 'success'
			}
		}
		RenameSetup                  = @{
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
							if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
							SyncConfigToUI
						}
					}
				}
			}
		}
		DeleteSetup                  = @{
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
								if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
								SyncConfigToUI
							}
						}
					}
				}
			}
		}
		LoginConfigGrid              = @{
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
		LoginProfileSelector         = @{
			SelectedIndexChanged = {
				param($s, $e)
				$oldProfile = $s.Tag
				if (-not $oldProfile) { $oldProfile = 'Default' }

				Write-Verbose "  UI: Saving settings for '$oldProfile' before switching..."
				SyncUIToConfig -ProfileToSave $oldProfile
				if (Test-DashboardCommand 'WriteConfig') { WriteConfig }

				Write-Verbose '  UI: Loading settings for new selection...'
				SyncConfigToUI
			}
		}
		MasterPauseHotkeyBtn         = @{
			Click = {
				$currentKey = $global:DashboardConfig.UI.MasterPauseHotkeyBtn.Text
				$newKey = Show-KeyCaptureDialog -currentKey $currentKey -OwnerForm $global:DashboardConfig.UI.SettingsForm
				if ($null -ne $newKey -and $newKey -ne $currentKey)
				{
					if (-not $global:DashboardConfig.Config.Contains('GlobalHotkeys')) { $global:DashboardConfig.Config['GlobalHotkeys'] = [ordered]@{} }
					$global:DashboardConfig.Config['GlobalHotkeys']['MasterPauseHotkey'] = $newKey
					$global:DashboardConfig.UI.MasterPauseHotkeyBtn.Text = $newKey

					if (Test-DashboardCommand 'WriteConfig') { WriteConfig }

					RegisterConfiguredHotkeys
					RefreshHotkeysList
				}
			}
		}
		MasterResumeHotkeyBtn        = @{
			Click = {
				$currentKey = $global:DashboardConfig.UI.MasterResumeHotkeyBtn.Text
				$newKey = Show-KeyCaptureDialog -currentKey $currentKey -OwnerForm $global:DashboardConfig.UI.SettingsForm
				if ($null -ne $newKey -and $newKey -ne $currentKey)
				{
					if (-not $global:DashboardConfig.Config.Contains('GlobalHotkeys')) { $global:DashboardConfig.Config['GlobalHotkeys'] = [ordered]@{} }
					$global:DashboardConfig.Config['GlobalHotkeys']['MasterResumeHotkey'] = $newKey
					$global:DashboardConfig.UI.MasterResumeHotkeyBtn.Text = $newKey

					if (Test-DashboardCommand 'WriteConfig') { WriteConfig }

					RegisterConfiguredHotkeys
					RefreshHotkeysList
				}
			}
		}
		Save                         = @{
			Click = {
				SyncUIToConfig
				WriteConfig
				HideSettingsForm
			}
		}
		Cancel                       = @{ 
			Click = { HideSettingsForm } 
		}
		UnregisterHotkey             = @{
			Click = {
				$grid = $global:DashboardConfig.UI.HotkeysGrid
				if (-not $grid) { return }
				
				$selectedRows = $grid.SelectedRows
				if ($selectedRows.Count -eq 0) { return }

				if ((Show-DarkMessageBox "Are you sure you want to unregister and delete $($selectedRows.Count) hotkey(s)?" 'Confirm Unregister' 'YesNo' 'Question') -ne 'Yes')
				{ 
					return 
				}

				foreach ($row in $selectedRows)
				{
					if (-not $row.Tag -or -not $row.Tag.Id) { continue }
					if ($row.Tag)
					{
						$id = $row.Tag.Id
						$ownerKey = $row.Tag.OwnerKey
						
						if (Test-DashboardCommand 'UnregisterHotkeyInstance')
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
										$cmd = Get-QuickCommandById -Data $data -CmdId $cmdId
										if ($cmd) { $cmd.Hotkey = $null; $cmd.HotkeyId = $null }
										if (Test-DashboardCommand 'UpdateChatCommanderSettings') { UpdateChatCommanderSettings $data -forceWrite }
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
										if (Test-DashboardCommand 'UpdateMacroSettings') { UpdateMacroSettings $data -forceWrite }
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
										if (Test-DashboardCommand 'UpdateSettings') { UpdateSettings $data -forceWrite }
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
									if ($form -and $form.Tag) { if (Test-DashboardCommand 'UpdateSettings') { UpdateSettings $form.Tag $extData -forceWrite } }
								}
							}
							elseif ($global:DashboardConfig.Resources.FtoolForms.Contains($ownerKey))
							{
								$form = $global:DashboardConfig.Resources.FtoolForms[$ownerKey]
								if ($form -and -not $form.IsDisposed) { $data = $form.Tag; $data.Hotkey = $null; $data.HotkeyId = $null; $data.BtnHotKey.Text = 'Hotkey'; if (Test-DashboardCommand 'UpdateSettings') { UpdateSettings $data -forceWrite } }
							}
						}
					}
				}
				if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
				SyncConfigToUI
				RegisterConfiguredHotkeys
				RefreshHotkeysList
			}
		}

		#endregion
		#region ExtraForm
		WorldBossListener            = @{
			CheckedChanged = {
				param($s, $e)

				if (-not $global:DashboardConfig.Config.Contains('Options')) { $global:DashboardConfig.Config['Options'] = [ordered]@{} }
				$newValue = if ($s.Checked) { '1' } else { '0' }
				if ($global:DashboardConfig.Config['Options']['WorldBossListener'] -ne $newValue)
				{
					$global:DashboardConfig.Config['Options']['WorldBossListener'] = $newValue
					if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
				}

				if ($s.Checked)
				{
					if (Test-DashboardCommand 'Start-WorldBossListener') { Start-WorldBossListener }
				}
				else
				{
					if (Test-DashboardCommand 'Stop-WorldBossListener') { Stop-WorldBossListener }
				}
			}
		}
		RefreshBosses                = @{
			Click = {
				if ($global:DashboardConfig.State.WorldBossListener -and $global:DashboardConfig.State.WorldBossListener.Timer)
				{
					$global:DashboardConfig.State.WorldBossListener.IsFirstRun = $true
					$global:DashboardConfig.State.WorldBossListener.NextRunTime = [DateTime]::MinValue
				}
				else
				{
					if (Test-DashboardCommand 'ShowToast') { ShowToast -Title 'World Boss' -Message 'Listener is not running.' -Type 'Warning' }
				}
			}
		}
		NotificationPositionSelector = @{
			SelectedIndexChanged = {
				param($s, $e)
				if (-not $global:DashboardConfig.Config.Contains('Options')) { $global:DashboardConfig.Config['Options'] = [ordered]@{} }
				$newValue = $s.SelectedItem
				if ($newValue -and $global:DashboardConfig.Config['Options']['NotificationPosition'] -ne $newValue)
				{
					$global:DashboardConfig.Config['Options']['NotificationPosition'] = $newValue
					if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
					if (Test-DashboardCommand 'UpdateNotificationPositions') { UpdateNotificationPositions }
				}
			}
		}
		ShowBossImages               = @{
			CheckedChanged = {
				param($s, $e)
				if (-not $global:DashboardConfig.Config.Contains('Options')) { $global:DashboardConfig.Config['Options'] = [ordered]@{} }
				$newValue = if ($s.Checked) { '1' } else { '0' }
				if ($global:DashboardConfig.Config['Options']['ShowBossImages'] -ne $newValue)
				{
					$global:DashboardConfig.Config['Options']['ShowBossImages'] = $newValue
					if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
				}

				if ($global:DashboardConfig.UI.ButtonPanel)
				{
					foreach ($ctrl in $global:DashboardConfig.UI.ButtonPanel.Controls)
					{
						if ($ctrl -is [System.Windows.Forms.Button])
						{
							Update-BossButtonImage $ctrl
						}
					}
				}

				if ($global:DashboardConfig.Resources.Timers['BossTimer'])
				{
					$t = $global:DashboardConfig.Resources.Timers['BossTimer']
					$t.Stop()
					$t.Interval = 1000
					$t.Start()
				}
			}
		}
		NoteGrid                     = @{
			DoubleClick = { if (Test-DashboardCommand 'EditNote') { EditNote } }
		}
		NotificationGrid             = @{
			CellDoubleClick = {
				param($s, $e)
				$script:NotificationGridEventDepth++
				try
				{
					if ($e.RowIndex -ge 0)
					{
						$row = $s.Rows[$e.RowIndex]
						if ($row.Tag)
						{
							$entry = $row.Tag
							if ($entry.ExtraData -and $entry.ExtraData.IsInteractive)
							{
								if (Test-DashboardCommand 'ShowInteractiveNotification') { ShowInteractiveNotification -Title $entry.Title -Message $entry.Message -Buttons $entry.ExtraData.Buttons -Type $entry.Type -Key $entry.Key -TimeoutSeconds $entry.TimeoutSeconds -Progress $entry.Progress -IgnoreCancellation:$entry.IgnoreCancellation }
							}
							else
							{
								if (Test-DashboardCommand 'ShowToast') { ShowToast -Title $entry.Title -Message $entry.Message -Type $entry.Type -Key $entry.Key -TimeoutSeconds $entry.TimeoutSeconds -Progress $entry.Progress -IgnoreCancellation:$entry.IgnoreCancellation }
							}
						}
					}
				}
				finally
				{
					$script:NotificationGridEventDepth--
				}
			}
		}
		ShowNotification             = @{
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
							if (Test-DashboardCommand 'ShowInteractiveNotification') { ShowInteractiveNotification -Title $entry.Title -Message $entry.Message -Buttons $entry.ExtraData.Buttons -Type $entry.Type -Key $entry.Key -TimeoutSeconds $entry.TimeoutSeconds -Progress $entry.Progress -IgnoreCancellation:$entry.IgnoreCancellation }
						}
						else
						{
							if (Test-DashboardCommand 'ShowToast') { ShowToast -Title $entry.Title -Message $entry.Message -Type $entry.Type -Key $entry.Key -TimeoutSeconds $entry.TimeoutSeconds -Progress $entry.Progress -IgnoreCancellation:$entry.IgnoreCancellation }
						}
					}
				}
			}
		}
		HideAllNotifications         = @{
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
		AddNote                      = @{
			Click = { if (Test-DashboardCommand 'AddNote') { AddNote } }
		}
		EditNote                     = @{
			Click = { if (Test-DashboardCommand 'EditNote') { EditNote } }
		}
		RemoveNote                   = @{
			Click = { if (Test-DashboardCommand 'RemoveNote') { RemoveNote } }
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
				Show-DarkMessageBox "1. Click on your client. It must be in focus!`n2. Hover with the mouse above the target button.`n3. Wait 3 seconds until the settings page opens again.`n`nClick OK to start the timer." 'Set Profile Coordinates' 'Ok' 'Information'
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

	$manualCoordsBox = $global:DashboardConfig.UI.ManualCoordsOverride
	if ($manualCoordsBox)
	{
		$manualCoordsBox.Add_CheckedChanged({
				param($s, $e)
				if ($global:DashboardConfig.State -and $global:DashboardConfig.State.LoadingLoginSettings) { return }
				try
				{
					$manualProfileName = 'Default'
					if ($global:DashboardConfig.UI.LoginProfileSelector -and $global:DashboardConfig.UI.LoginProfileSelector.SelectedItem) { $manualProfileName = $global:DashboardConfig.UI.LoginProfileSelector.SelectedItem.ToString() }
					$manualCfg = $global:DashboardConfig.Config
					if ($manualCfg -and $manualCfg.Contains('LoginConfig'))
					{
						$manualLoginCfg = $manualCfg['LoginConfig']
						if (-not $manualLoginCfg.Contains($manualProfileName) -or $manualLoginCfg[$manualProfileName] -isnot [System.Collections.IDictionary]) { $manualLoginCfg[$manualProfileName] = [ordered]@{} }
						$manualLoginCfg[$manualProfileName]['ManualCoordsOverride'] = if ($s.Checked) { 'Yes' } else { 'No' }
					}
				}
				catch { }
				if (-not $s.Checked) { StashUserLoginCoordinates }
				UpdateLoginCoordinateFields
			})
	}

	$resolutionBox = $global:DashboardConfig.UI.ResolutionSelector
	if ($resolutionBox)
	{
		$resolutionBox.Add_SelectedIndexChanged({
				param($s, $e)
				if ($global:DashboardConfig.State -and $global:DashboardConfig.State.LoadingLoginSettings) { return }
				$box = $global:DashboardConfig.UI.ManualCoordsOverride
				if ($box -and $box.Checked) { return }
				UpdateLoginCoordinateFields
			})
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

	try
	{
		if (Test-DashboardCommand 'StartProfileAutoToolsWatcher') { StartProfileAutoToolsWatcher }
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to start watcher from InitializeUI: $_" }

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
		[object]$tag,
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
		'Form'               { New-Object System.Windows.Forms.Form }
		'Panel'              { New-Object System.Windows.Forms.Panel }
		'Button'             { New-Object Custom.KeySafeButton }
		'Label'              { New-Object System.Windows.Forms.Label }
		'DataGridView'       { New-Object Custom.DarkDataGridView }
		'TextBox'            { New-Object System.Windows.Forms.TextBox }
		'ComboBox'           { New-Object System.Windows.Forms.ComboBox }
		'Custom.DarkComboBox'{ New-Object Custom.DarkComboBox }
		'CheckBox'           { New-Object Custom.KeySafeCheckBox }
		'Toggle'             { New-Object Custom.Toggle }
		'ProgressBar'        { New-Object System.Windows.Forms.ProgressBar }
		'TextProgressBar'    { New-Object Custom.TextProgressBar }
		default              { throw "Invalid element type specified: $type" }
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
		$el.ShowCellToolTips = $false
		$el.ColumnHeadersHeightSizeMode = 'DisableResizing'
		$el.RowHeadersWidthSizeMode = 'DisableResizing'
		$el.DefaultCellStyle.Alignment = 'MiddleCenter'
		$el.ColumnHeadersDefaultCellStyle.Alignment = 'MiddleCenter'
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
		if ($PSBoundParameters.ContainsKey('font')) { $el.Font = $font }
		if ($PSBoundParameters.ContainsKey('tag')) { $el.Tag = $tag }
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
				$el.FlatAppearance.BorderSize = 1
				if (-not $noCustomPaint -and (Test-DashboardCommand 'Register-ThemedFlatButtonPaint'))
				{
					Register-ThemedFlatButtonPaint $el
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
		}
		'ComboBox'
		{
			if ($null -ne $dropDownStyle) { try { $el.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::$dropDownStyle } catch {} }
			if ($null -ne $fs) { $el.FlatStyle = $fs }
		}
		'Custom.DarkComboBox'
		{
			if ($null -ne $dropDownStyle) { try { $el.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::$dropDownStyle } catch {} }
			if ($null -ne $fs) { $el.FlatStyle = $fs }
			$el.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
			$el.IntegralHeight = $false
			if ($PSBoundParameters.ContainsKey('font')) { $el.Font = $font }
		}
		'CheckBox'
		{
			if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
			if ($PSBoundParameters.ContainsKey('checked')) { $el.Checked = $checked }
			$el.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
			$el.FlatAppearance.BorderSize = 1
			$el.UseVisualStyleBackColor = $false
			$el.CheckAlign = [System.Drawing.ContentAlignment]::MiddleLeft
			$el.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
			$el.Padding = [System.Windows.Forms.Padding]::new(20, 0, 0, 0)
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

	if (Test-DashboardCommand 'ApplyControlInternal')
	{
		ApplyControlInternal $el (Get-Theme)
	}

	return $el
}

#endregion

#region Theme Selection & Editor Dialog

function Show-CustomColorPicker {
	param(
		[System.Drawing.Color]$InitialColor,
		[System.Windows.Forms.Form]$OwnerForm,
		[scriptblock]$OnColorChange
	)

	function Hsl-ToRgb([double]$h, [double]$s, [double]$l) {
		$h = $h - [Math]::Floor($h / 360.0) * 360.0
		$c = (1.0 - [Math]::Abs(2.0 * $l - 1.0)) * $s
		$hp = $h / 60.0
		$x = $c * (1.0 - [Math]::Abs(($hp % 2.0) - 1.0))
		$m = $l - $c / 2.0
		$r = 0.0; $g = 0.0; $b = 0.0
		if     ($h -lt 60)  { $r = $c; $g = $x }
		elseif ($h -lt 120) { $r = $x; $g = $c }
		elseif ($h -lt 180) { $g = $c; $b = $x }
		elseif ($h -lt 240) { $g = $x; $b = $c }
		elseif ($h -lt 300) { $r = $x; $b = $c }
		else                { $r = $c; $b = $x }
		return [System.Drawing.Color]::FromArgb(
			[Math]::Min(255, [int](($r + $m) * 255)),
			[Math]::Min(255, [int](($g + $m) * 255)),
			[Math]::Min(255, [int](($b + $m) * 255)))
	}

	function Rgb-ToHsl([System.Drawing.Color]$col) {
		[double]$r = $col.R / 255.0; [double]$g = $col.G / 255.0; [double]$b = $col.B / 255.0
		$max = [Math]::Max($r, [Math]::Max($g, $b))
		$min = [Math]::Min($r, [Math]::Min($g, $b))
		$state.L = ($max + $min) / 2.0
		if ($max -eq $min) { $state.H = 0.0; $state.S = 0.0; return }
		$d = $max - $min
		$state.S = if ($state.L -gt 0.5) { $d / (2.0 - $max - $min) } else { $d / ($max + $min) }
		$h = 0.0
		if ($max -eq $r)     { $h = (($g - $b) / $d) * 60 }
		elseif ($max -eq $g) { $h = (($b - $r) / $d) * 60 + 120 }
		else                 { $h = (($r - $g) / $d) * 60 + 240 }
		if ($h -lt 0) { $h += 360 }
		$state.H = $h
	}

	$state = @{ H = 0.0; S = 0.0; L = 0.5; Suppress = $false }
	Rgb-ToHsl $InitialColor

	$tc = if (Test-DashboardCommand 'Get-Theme') { Get-Theme } else { @{} }
	$dlgBg     = if ($tc.ContainsKey('SecondaryBg'))        { $tc['SecondaryBg'] }        else { [System.Drawing.Color]::FromArgb(30, 33, 40) }
	$dlgFg     = if ($tc.ContainsKey('PrimaryText'))        { $tc['PrimaryText'] }        else { [System.Drawing.Color]::White }
	$labelClr  = if ($tc.ContainsKey('HeaderSubtitleText')) { $tc['HeaderSubtitleText'] } else { [System.Drawing.Color]::FromArgb(140, 145, 155) }
	$accentClr = if ($tc.ContainsKey('AccentBlue'))          { $tc['AccentBlue'] }         else { [System.Drawing.Color]::FromArgb(31, 111, 235) }
	$buttonBg  = if ($tc.ContainsKey('ButtonBg'))            { $tc['ButtonBg'] }           else { [System.Drawing.Color]::FromArgb(55, 60, 70) }

	$dlg = New-Object System.Windows.Forms.Form
	$dlg.Text = 'Pick Color'
	$dlg.Size = New-Object System.Drawing.Size(275, 320)
	$dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.ShowInTaskbar = $false
	$dlg.BackColor = $dlgBg; $dlg.ForeColor = $dlgFg
	$dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
	if ($OwnerForm) { $dlg.Owner = $OwnerForm }

	$wheelSize = 200
	$wheelBmp  = New-Object System.Drawing.Bitmap($wheelSize, $wheelSize)
	$center    = $wheelSize / 2.0
	$maxRadius = $center - 1.0

	[Custom.ColorWheelHelper]::RenderWheel($wheelBmp, $state.L)

	$wheelPic = New-Object System.Windows.Forms.PictureBox
	$wheelPic.Size     = New-Object System.Drawing.Size($wheelSize, $wheelSize)
	$wheelPic.Location = New-Object System.Drawing.Point(10, 10)
	$wheelPic.Image    = $wheelBmp
	$wheelPic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Normal
	$wheelPic.Cursor   = [System.Windows.Forms.Cursors]::Cross
	$dlg.Controls.Add($wheelPic)

	$wheelPic.Add_Paint({
		param($s, $e)
		$rad = $maxRadius * $state.S
		$angleRad = $state.H * [Math]::PI / 180.0
		$cx = [int]($center + $rad * [Math]::Cos($angleRad))
		$cy = [int]($center + $rad * [Math]::Sin($angleRad))

		$penB = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 2)
		$penW = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 1)

		$e.Graphics.DrawEllipse($penB, $cx - 6, $cy - 6, 12, 12)
		$e.Graphics.DrawEllipse($penW, $cx - 5, $cy - 5, 10, 10)

		$penB.Dispose(); $penW.Dispose()
	})

	$brightnessPanel = New-Object System.Windows.Forms.Panel
	$brightnessPanel.Size     = New-Object System.Drawing.Size(25, $wheelSize)
	$brightnessPanel.Location = New-Object System.Drawing.Point(218, 10)
	$brightnessPanel.Cursor   = [System.Windows.Forms.Cursors]::Hand
	$dlg.Controls.Add($brightnessPanel)

	$brightnessPanel.Add_Paint({
		param($s, $e)
		$g = $e.Graphics
		$h = $brightnessPanel.Height; $w = $brightnessPanel.Width

		$pureColor = Hsl-ToRgb $state.H $state.S 0.5
		$rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
		$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, [System.Drawing.Color]::White, [System.Drawing.Color]::Black, 90.0)

		$cb = New-Object System.Drawing.Drawing2D.ColorBlend
		$cb.Positions = @(0.0, 0.5, 1.0)
		$cb.Colors    = @([System.Drawing.Color]::White, $pureColor, [System.Drawing.Color]::Black)
		$brush.InterpolationColors = $cb

		$g.FillRectangle($brush, $rect)
		$brush.Dispose()

		$iy = [int]((1.0 - $state.L) * ($h - 1))
		$penB = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 3)
		$penW = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 1)
		$g.DrawLine($penB, 0, $iy, $w, $iy)
		$g.DrawLine($penW, 0, $iy, $w, $iy)
		$penB.Dispose(); $penW.Dispose()
	})

	$hexLbl = New-Object System.Windows.Forms.Label
	$hexLbl.Text = '#'
	$hexLbl.Location = New-Object System.Drawing.Point(10, 218)
	$hexLbl.Size = New-Object System.Drawing.Size(12, 18)
	$hexLbl.ForeColor = $labelClr
	$dlg.Controls.Add($hexLbl)

	$hexBox = New-Object System.Windows.Forms.TextBox
	$hexBox.Location = New-Object System.Drawing.Point(22, 215)
	$hexBox.Size = New-Object System.Drawing.Size(70, 22)
	$hexBox.BackColor = $dlgBg
	$hexBox.ForeColor = $dlgFg
	$hexBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	$hexBox.Font = Get-UIFont 'Consolas' 9
	$hexBox.MaxLength = 6
	$dlg.Controls.Add($hexBox)

	$previewBox = New-Object System.Windows.Forms.Panel
	$previewBox.Location = New-Object System.Drawing.Point(102, 215)
	$previewBox.Size = New-Object System.Drawing.Size(141, 22)
	$previewBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	$previewBox.BackColor = $InitialColor
	$dlg.Controls.Add($previewBox)

	$syncColor = {
		param([System.Drawing.Color]$c, [bool]$updateWheel)
		$previewBox.BackColor = $c
		$state.Suppress = $true
		$hexBox.Text = ('{0:X2}{1:X2}{2:X2}' -f $c.R, $c.G, $c.B)
		$state.Suppress = $false

		if ($updateWheel) {
			[Custom.ColorWheelHelper]::RenderWheel($wheelBmp, $state.L)
			$wheelPic.Image = $wheelBmp
		}

		$wheelPic.Invalidate()
		$brightnessPanel.Invalidate()
		if ($OnColorChange) { & $OnColorChange $c }
	}

	$wheelDown = $false
	$handleWheel = {
		param($s, $e)
		$dx = $e.X - $center; $dy = $e.Y - $center
		$dist = [Math]::Sqrt($dx * $dx + $dy * $dy)
		if ($dist -gt $maxRadius) { $dist = $maxRadius }
		$state.S = $dist / $maxRadius
		$angle = [Math]::Atan2($dy, $dx) * 180.0 / [Math]::PI
		if ($angle -lt 0) { $angle += 360.0 }
		$state.H = $angle

		$c = Hsl-ToRgb $state.H $state.S $state.L
		& $syncColor $c $false
	}
	$wheelPic.Add_MouseDown({ $wheelDown = $true; & $handleWheel $args[0] $args[1] })
	$wheelPic.Add_MouseMove({ if ($wheelDown) { & $handleWheel $args[0] $args[1] } })
	$wheelPic.Add_MouseUp({   $wheelDown = $false })

	$brightnessDown = $false
	$handleBrightness = {
		param($s, $e)
		$y = [Math]::Max(0, [Math]::Min($s.Height - 1, $e.Y))
		$state.L = 1.0 - ($y / [double]($s.Height - 1))
		$c = Hsl-ToRgb $state.H $state.S $state.L
		& $syncColor $c $true
	}
	$brightnessPanel.Add_MouseDown({ $brightnessDown = $true; & $handleBrightness $args[0] $args[1] })
	$brightnessPanel.Add_MouseMove({ if ($brightnessDown) { & $handleBrightness $args[0] $args[1] } })
	$brightnessPanel.Add_MouseUp({   $brightnessDown = $false })

	$hexBox.Add_TextChanged({
		if ($state.Suppress) { return }
		$h = $hexBox.Text.Trim().TrimStart('#')
		if ($h.Length -eq 6) {
			try {
				$rv = [Convert]::ToInt32($h.Substring(0, 2), 16)
				$gv = [Convert]::ToInt32($h.Substring(2, 2), 16)
				$bv = [Convert]::ToInt32($h.Substring(4, 2), 16)
				$col = [System.Drawing.Color]::FromArgb($rv, $gv, $bv)
				Rgb-ToHsl $col
				& $syncColor $col $true
			} catch { }
		}
	})

	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = 'OK'; $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$btnOK.Location = New-Object System.Drawing.Point(48, 248); $btnOK.Size = New-Object System.Drawing.Size(80, 28)
	$btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$btnOK.BackColor = $accentClr; $btnOK.ForeColor = [System.Drawing.Color]::White
	$btnOK.FlatAppearance.BorderSize = 0
	$dlg.Controls.Add($btnOK)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = 'Cancel'; $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$btnCancel.Location = New-Object System.Drawing.Point(140, 248); $btnCancel.Size = New-Object System.Drawing.Size(80, 28)
	$btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$btnCancel.BackColor = $buttonBg; $btnCancel.ForeColor = [System.Drawing.Color]::White
	$btnCancel.FlatAppearance.BorderSize = 0
	$dlg.Controls.Add($btnCancel)

	$dlg.AcceptButton = $btnOK
	$dlg.CancelButton = $btnCancel

	$hexBox.Text = ('{0:X2}{1:X2}{2:X2}' -f $InitialColor.R, $InitialColor.G, $InitialColor.B)

	$result = $dlg.ShowDialog($OwnerForm)
	$selectedColor = $previewBox.BackColor
	$wheelBmp.Dispose()
	$dlg.Dispose()

	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		return $selectedColor
	}
	return $null
}

function Show-LiveColorPicker {
	param(
		[System.Windows.Forms.Button]$Button,
		[string]$ColorKey,
		[System.Collections.Generic.Dictionary[string, System.Drawing.Color]]$ColorMap,
		[System.Windows.Forms.Form]$OwnerForm,
		[scriptblock]$OnApply
	)
	$originalCol = $Button.BackColor
	$appliedColor = $originalCol

	$applyColor = {
		param($newCol)
		$Button.BackColor = $newCol
		if ($OnApply) {
			& $OnApply $ColorKey $newCol
		} else {
			$ColorMap[$ColorKey] = $newCol
			if ($global:DashboardConfig.State.CurrentTheme -eq 'Custom') {
				Save-CustomThemeColors $ColorMap -NoConfigWrite
				Set-Theme 'Custom' -NoSave
				#Reapply-GridStatusStyling
			}
		}
	}

	$onLiveChange = {
		param($newCol)
		& $applyColor $newCol
	}

	$newCol = Show-CustomColorPicker -InitialColor $originalCol -OwnerForm $OwnerForm -OnColorChange $onLiveChange

	if ($null -eq $newCol) {
		& $applyColor $originalCol
	}
}

function Show-ThemeSelectionForm {
	[CmdletBinding()]
	param()

	$c = Get-Theme
	$currentTheme = if ($global:DashboardConfig.State.CurrentTheme) { $global:DashboardConfig.State.CurrentTheme } else { 'OriginalGray' }
	$initialTheme = $currentTheme
	$editorState = @{ IsSavedAndApplied = $false }

	$customColorMap = [System.Collections.Generic.Dictionary[string, System.Drawing.Color]]::new([System.StringComparer]::OrdinalIgnoreCase)
	$customThemeDict = Get-ThemeColors 'Custom'
	if ($customThemeDict) {
		foreach ($k in $customThemeDict.Keys) { $customColorMap[$k] = $customThemeDict[$k] }
	}

	$colorButtons = @{}
	$themeCards = @{}
	$allThemes = Set-Theme -List

	$themeForm = New-Object System.Windows.Forms.Form
	$themeForm.Text = 'Theme Selector & Editor'
	$themeForm.Size = New-Object System.Drawing.Size(680, 580)
	$themeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
	$themeForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
	if ($global:DashboardConfig.UI.MainForm) { $themeForm.Owner = $global:DashboardConfig.UI.MainForm }
	$themeForm.BackColor = $c['MainBg']
	$themeForm.ForeColor = $c['PrimaryText']

	$topBar = New-Object System.Windows.Forms.Panel
	$topBar.Dock = [System.Windows.Forms.DockStyle]::Top
	$topBar.Height = 32
	$topBar.Add_MouseDown({ [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($args[0].FindForm().Handle, 0xA1, 0x2, 0) })
	Register-ThemedTopBarPaint $topBar

	$titleLbl = New-Object System.Windows.Forms.Label
	$titleLbl.Text = 'THEME SELECTOR & EDITOR'
	$titleLbl.Font = Get-UIFont 'Segoe UI' 9 'Bold'
	$titleLbl.ForeColor = $c['HeaderTitleText']
	$titleLbl.AutoSize = $true
	$titleLbl.Location = New-Object System.Drawing.Point(12, 7)
	$titleLbl.Add_MouseDown({ [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($args[0].FindForm().Handle, 0xA1, 0x2, 0) })
	$topBar.Controls.Add($titleLbl)

	$btnClose = New-Object System.Windows.Forms.Button
	$btnClose.Text = [char]0x166D
	$btnClose.Font = Get-UIFont 'Segoe UI' 10 'Bold'
	$btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$btnClose.FlatAppearance.BorderSize = 0
	$btnClose.Size = New-Object System.Drawing.Size(30, 24)
	$btnClose.Location = New-Object System.Drawing.Point(645, 3)
	$btnClose.ForeColor = $c['PrimaryText']
	$btnClose.Add_Click({ $args[0].FindForm().Close() })
	Register-ThemedButtonFadeEffect -Button $btnClose -Variant 'Danger'
	$topBar.Controls.Add($btnClose)
	$themeForm.Controls.Add($topBar)

	$themeTabControl = New-Object Custom.DarkTabControl
	$themeTabControl.Location = New-Object System.Drawing.Point(8, 38)
	$themeTabControl.Size = New-Object System.Drawing.Size(335, 500)
	$themeTabControl.Font = Get-UIFont 'Segoe UI' 8 'Bold'
	$themeTabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Normal
	$themeTabControl.ItemSize = New-Object System.Drawing.Size(0, 26)
	$themeTabControl.BackColor = $c['TabBg']
	$themeTabControl.ForeColor = $c['PrimaryText']
	$themeTabControl.TabBgColor = $c['TabBg']
	$themeTabControl.TabActiveBgColor = $c['TabActive']
	$themeTabControl.TabInactiveBgColor = $c['TabInactive']
	$themeTabControl.TabActiveTextColor = $c['TabActiveText']
	$themeTabControl.TabInactiveTextColor = $c['TabInactiveText']
	$themeTabControl.TabUnderlineColor = $c['TabUnderline']
	$themeTabControl.TabInactiveUnderlineColor = $c['TabInactiveUnderline']

	$lightNames = @('ArcticFrost','DeepOceanLight','FlyffClassic','Light','QuartzStone','SunnyMeadow','TerraCotta')
	$darkNames  = @('CrimsonTwilight','CyberViolet','Dark','DeepOcean','FlyffUniverse','ForestSage','MidnightSlate','OriginalGray')

	$drawMiniPreview = {
		param($gfx, $rect, $colorMap)
		$getColor = {
			param($key, $defaultColor)
			if ($colorMap -and $colorMap.ContainsKey($key) -and $colorMap[$key] -is [System.Drawing.Color]) {
				return $colorMap[$key]
			}
			return $defaultColor
		}
		$mainBg    = &$getColor 'MainBg' (Get-Theme 'MainBg')
		$topBg     = &$getColor 'TopBarBg' $mainBg
		$secBg     = &$getColor 'SecondaryBg' $mainBg
		$txtColor  = &$getColor 'PrimaryText' (Get-Theme 'WhiteText')
		$btnColor  = &$getColor 'ButtonBg' $secBg
		$accentCol = &$getColor 'ComboSelectedItem' (Get-Theme 'ComboSelectedItem')
		$gridBg    = &$getColor 'GridBg' $mainBg
		$gridAlt   = &$getColor 'GridAlternateRow' $secBg
		$borderCol = &$getColor 'BorderColor' ([System.Drawing.Color]::Gray)

		$rx = [int]$rect.X; $ry = [int]$rect.Y; $rw = [int]$rect.Width; $rh = [int]$rect.Height

		$bgB = New-Object System.Drawing.SolidBrush -ArgumentList @($mainBg)
		$gfx.FillRectangle($bgB, $rx, $ry, $rw, $rh); $bgB.Dispose()

		$topB = New-Object System.Drawing.SolidBrush -ArgumentList @($topBg)
		$gfx.FillRectangle($topB, $rx, $ry, $rw, 12); $topB.Dispose()

		$txtB = New-Object System.Drawing.SolidBrush -ArgumentList @($txtColor)
		$gfx.FillRectangle($txtB, $rx + 3, $ry + 3, 30, 4)
		$gfx.FillRectangle($txtB, $rx + $rw - 12, $ry + 3, 8, 4); $txtB.Dispose()

		$btnB = New-Object System.Drawing.SolidBrush -ArgumentList @($btnColor)
		$gfx.FillRectangle($btnB, $rx + 4, $ry + 16, 22, 10)
		$gfx.FillRectangle($btnB, $rx + 29, $ry + 16, 22, 10)
		$gfx.FillRectangle($btnB, $rx + 54, $ry + 16, 16, 10); $btnB.Dispose()

		$accB = New-Object System.Drawing.SolidBrush -ArgumentList @($accentCol)
		$gfx.FillRectangle($accB, $rx + 73, $ry + 16, 22, 10); $accB.Dispose()

		$gBgB = New-Object System.Drawing.SolidBrush -ArgumentList @($gridBg)
		$gfx.FillRectangle($gBgB, $rx + 4, $ry + 30, $rw - 8, $rh - 34); $gBgB.Dispose()

		$rAltB = New-Object System.Drawing.SolidBrush -ArgumentList @($gridAlt)
		$gfx.FillRectangle($rAltB, $rx + 4, $ry + 38, $rw - 8, 6)
		$gfx.FillRectangle($rAltB, $rx + 4, $ry + 50, $rw - 8, 6); $rAltB.Dispose()

		$gSelB = New-Object System.Drawing.SolidBrush -ArgumentList @($accentCol)
		$gfx.FillRectangle($gSelB, $rx + 4, $ry + 44, $rw - 8, 6); $gSelB.Dispose()

		$pen = New-Object System.Drawing.Pen -ArgumentList @($borderCol)
		$gfx.DrawRectangle($pen, $rx, $ry, $rw - 1, $rh - 1); $pen.Dispose()
	}

	$selectThemeHandler = {
		param($s, $e)
		$frm = $s.FindForm()
		$ft = $frm.Tag

		$targetCard = $s
		while ($targetCard -and $targetCard -isnot [System.Windows.Forms.Panel] -and $targetCard.Parent) {
			$targetCard = $targetCard.Parent
		}
		$selectedName = $targetCard.Tag
		if (-not $selectedName) { return }

		Set-Theme $selectedName
		#Reapply-GridStatusStyling
		$ft.EditorState.IsSavedAndApplied = $true

		$tc = $ft.ThemeCards
		$curC = Get-Theme
		foreach ($k in $tc.Keys) {
			$cCard = $tc[$k]
			$isSel = ($k -eq $selectedName)
			$cCard.BackColor = if ($isSel) { $curC['ComboSelectedItem'] } else { $curC['MainBg'] }
			foreach ($ctrl in $cCard.Controls) {
				if ($ctrl -is [System.Windows.Forms.Label]) {
					if ($ctrl.Font.Bold) {
						$ctrl.ForeColor = if ($isSel) { Get-Theme 'WhiteText' } else { $curC['PrimaryText'] }
					} else {
						$ctrl.ForeColor = $curC['HeaderSubtitleText']
					}
				}
			}
			$cCard.Invalidate()
		}

		$curC = Get-Theme
		$frm.BackColor = $curC['MainBg']
		$frm.ForeColor = $curC['PrimaryText']

		foreach ($ctrl in $frm.Controls) {
			if ($ctrl -is [System.Windows.Forms.TabControl]) {
				$ctrl.BackColor = $curC['PanelBg']
				if ($ctrl -is [Custom.DarkTabControl]) {
					$ctrl.TabBgColor                = $curC['TabBg']
					$ctrl.TabActiveBgColor          = $curC['TabActive']
					$ctrl.TabInactiveBgColor        = $curC['TabInactive']
					$ctrl.TabActiveTextColor        = $curC['TabActiveText']
					$ctrl.TabInactiveTextColor      = $curC['TabInactiveText']
					$ctrl.TabUnderlineColor         = $curC['TabUnderline']
					$ctrl.TabInactiveUnderlineColor = $curC['TabInactiveUnderline']
					$ctrl.BackColor                 = $curC['TabBg']
					$ctrl.ForeColor                 = $curC['TabActiveText']
				}
				foreach ($tp in $ctrl.TabPages) {
					$tp.BackColor = $curC['SecondaryBg']
					foreach ($inner in $tp.Controls) {
						if ($inner -is [System.Windows.Forms.FlowLayoutPanel]) {
							$inner.BackColor = $curC['SecondaryBg']
						}
					}
				}
			}
			elseif ($ctrl -is [System.Windows.Forms.Label]) {
				$ctrl.ForeColor = $curC['PrimaryText']
			}
			elseif ($ctrl -is [System.Windows.Forms.ComboBox]) {
				$ctrl.BackColor = $curC['InputBg']
				$ctrl.ForeColor = $curC['InputFg']
			}
			elseif ($ctrl -is [System.Windows.Forms.Button]) {
				$tagStr = if ($ctrl.Tag -is [System.Collections.IDictionary]) { [string]$ctrl.Tag['Type'] } else { [string]$ctrl.Tag }
				$ctrl.ForeColor = $curC['PrimaryText']
				$ctrl.FlatAppearance.BorderColor = $curC['BorderColor']

				$fadeProp = $ctrl.PSObject.Properties['FadeState']
				if ($fadeProp -and $fadeProp.Value) {
					$fs = $fadeProp.Value
					$prefix = switch ($tagStr) {
						'Danger'  { 'BtnDanger' }
						'Success' { 'BtnSuccess' }
						'Primary' { 'BtnPrimary' }
						default   { 'Btn' }
					}
					$top = $curC["${prefix}GradientNormalTop"]
					$bot = $curC["${prefix}GradientNormalBot"]
					$fs.NormalStartColor       = $top
					$fs.NormalEndColor         = $bot
					$fs.HoverStartColor        = $curC["${prefix}GradientHoverTop"]
					$fs.HoverEndColor          = $curC["${prefix}GradientHoverBot"]
					$fs.LeftPressedStartColor  = $curC["${prefix}GradientLeftPressTop"]
					$fs.LeftPressedEndColor    = $curC["${prefix}GradientLeftPressBot"]
					$fs.RightPressedStartColor = $curC["${prefix}GradientRightPressTop"]
					$fs.RightPressedEndColor   = $curC["${prefix}GradientRightPressBot"]
					$fs.TargetStartColor = $top
					$fs.TargetEndColor   = $bot
					$fs.CurrentStartR    = [double]$top.R
					$fs.CurrentStartG    = [double]$top.G
					$fs.CurrentStartB    = [double]$top.B
					$fs.CurrentEndR      = [double]$bot.R
					$fs.CurrentEndG      = [double]$bot.G
					$fs.CurrentEndB      = [double]$bot.B
				}
				$ctrl.Invalidate()
			}
		}

		$frm.Controls[0].Invalidate()
	}

	function New-ThemeCardInternal([string]$ThemeName, [string]$DisplayDesc, [string]$CurrentTheme) {
		$card = New-Object System.Windows.Forms.Panel
		$card.Size = New-Object System.Drawing.Size(310, 52)
		$card.Margin = New-Object System.Windows.Forms.Padding(2)
		$card.Cursor = [System.Windows.Forms.Cursors]::Hand
		$card.Tag = $ThemeName
		$card.BackColor = if ($ThemeName -eq $CurrentTheme) { (Get-Theme)['ComboSelectedItem'] } else { (Get-Theme)['MainBg'] }

		$pic = New-Object System.Windows.Forms.PictureBox
		$pic.Size = New-Object System.Drawing.Size(80, 46)
		$pic.Location = New-Object System.Drawing.Point(3, 3)
		$pic.Tag = $ThemeName
		$pic.Add_Paint({
			param($s, $e)
			$tn = $s.Tag
			$cdMap = if ($tn -eq 'Custom') { $customColorMap } else { Get-ThemeColors $tn }
			$pRect = New-Object System.Drawing.Rectangle(0, 0, $s.Width, $s.Height)
			& $drawMiniPreview $e.Graphics $pRect $cdMap
		})
		$card.Controls.Add($pic)

		$lbl = New-Object System.Windows.Forms.Label
		$lbl.Text = $ThemeName
		$lbl.Font = Get-UIFont 'Segoe UI' 9 'Bold'
		$lbl.Location = New-Object System.Drawing.Point(88, 5)
		$lbl.AutoSize = $true
		$lbl.ForeColor = if ($ThemeName -eq $CurrentTheme) { Get-Theme 'WhiteText' } else { $c['PrimaryText'] }
		$card.Controls.Add($lbl)

		$sub = New-Object System.Windows.Forms.Label
		$sub.Text = $DisplayDesc
		$sub.Font = Get-UIFont 'Segoe UI' 7 'Italic'
		$sub.Location = New-Object System.Drawing.Point(88, 25)
		$sub.AutoSize = $true
		$sub.ForeColor = $c['HeaderSubtitleText']
		$card.Controls.Add($sub)

		$card.Add_Click($selectThemeHandler)
		$pic.Add_Click($selectThemeHandler)
		$lbl.Add_Click($selectThemeHandler)
		$sub.Add_Click($selectThemeHandler)

		return $card
	}

	$lightTab = New-Object System.Windows.Forms.TabPage "Light"
	$lightTab.BackColor = $c['SecondaryBg']
	$lightFlow = New-Object System.Windows.Forms.FlowLayoutPanel
	$lightFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
	$lightFlow.WrapContents = $false
	$lightFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
	$lightFlow.BackColor = $c['SecondaryBg']
	$lightFlow.AutoScroll = $false

	foreach ($tn in $lightNames) {
		$card = New-ThemeCardInternal $tn 'Preset Theme' $currentTheme
		$themeCards[$tn] = $card
		$lightFlow.Controls.Add($card)
	}
	$lightTab.Controls.Add($lightFlow)
	$themeTabControl.TabPages.Add($lightTab)

	$darkTab = New-Object System.Windows.Forms.TabPage "Dark"
	$darkTab.BackColor = $c['SecondaryBg']
	$darkFlow = New-Object System.Windows.Forms.FlowLayoutPanel
	$darkFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
	$darkFlow.WrapContents = $false
	$darkFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
	$darkFlow.BackColor = $c['SecondaryBg']
	$darkFlow.AutoScroll = $false

	foreach ($tn in $darkNames) {
		$card = New-ThemeCardInternal $tn 'Preset Theme' $currentTheme
		$themeCards[$tn] = $card
		$darkFlow.Controls.Add($card)
	}
	$darkTab.Controls.Add($darkFlow)
	$themeTabControl.TabPages.Add($darkTab)

	$customTab = New-Object System.Windows.Forms.TabPage "Custom"
	$customTab.BackColor = $c['SecondaryBg']
	$customFlow = New-Object System.Windows.Forms.FlowLayoutPanel
	$customFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
	$customFlow.WrapContents = $false
	$customFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
	$customFlow.BackColor = $c['SecondaryBg']
	$customFlow.AutoScroll = $false

	$customCard = New-ThemeCardInternal 'Custom' 'User Defined Theme' $currentTheme
	$themeCards['Custom'] = $customCard
	$customFlow.Controls.Add($customCard)
	$customTab.Controls.Add($customFlow)
	$themeTabControl.TabPages.Add($customTab)

	$themeForm.Controls.Add($themeTabControl)

	$editorTabControl = New-Object Custom.DarkTabControl
	$editorTabControl.Location = New-Object System.Drawing.Point(350, 38)
	$editorTabControl.Size = New-Object System.Drawing.Size(318, 430)
	$editorTabControl.Font = Get-UIFont 'Segoe UI' 7.5 'Bold'
	$editorTabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Normal
	$editorTabControl.ItemSize = New-Object System.Drawing.Size(0, 24)
	$editorTabControl.BackColor = $c['TabBg']
	$editorTabControl.ForeColor = $c['PrimaryText']
	$editorTabControl.TabBgColor = $c['TabBg']
	$editorTabControl.TabActiveBgColor = $c['TabActive']
	$editorTabControl.TabInactiveBgColor = $c['TabInactive']
	$editorTabControl.TabActiveTextColor = $c['TabActiveText']
	$editorTabControl.TabInactiveTextColor = $c['TabInactiveText']
	$editorTabControl.TabUnderlineColor = $c['TabUnderline']
	$editorTabControl.TabInactiveUnderlineColor = $c['TabInactiveUnderline']

	$editorSections = [ordered]@{
		'Surfaces'    = @('MainBg','SecondaryBg','TopBarBg','BorderColor','PanelBg','TertiaryBg','TopBarAccent','GroupBoxBg','GroupBoxText')
		'Text'        = @('PrimaryText','SecondaryText','WhiteText','HeaderTitleText','HeaderSubtitleText','SectionHeaderText','AccentBlue','AccentDanger','StatusActive','CurrencyPerin','CurrencyEcp')
		'Semantic'    = @('SuccessColor','SuccessColorAlt','WarningColor','WarningColorAlt','ErrorColor','ErrorColorAlt')
		'Buttons'     = @('ButtonBg','HeaderButtonBg','BossButtonBg','BossTimerText')
		'Inputs'      = @('InputBg','InputBgAlt','InputFg','ComboDropdownBg','ComboItemText','ComboBorderNormal','ComboArrow','ComboSelectedItem','ComboSelectedText')
		'Grid'        = @('GridBg','GridBorder','GridSecondaryBg','GridAlternateRow','GridSelectedRow','GridSelectedText','GridHeaderBg','GridHeaderText','GridRowHeaderBg','GridRowHeaderText')
		'Tabs'        = @('TabBg','TabActive','TabInactive','TabActiveText','TabInactiveText','TabUnderline','TabInactiveUnderline','TabPageBg','TabPageText')
		'Controls'    = @('ToggleOnColor','ToggleOffColor','ProgressBarBg','ProgressBarFill','TreeViewBg','TreeViewHover','ListItemBg','ListItemText','ScrollThumbNormal','ScrollThumbHover','ScrollThumbActive')
		'Status'      = @('StatusDisconnected','StatusDisconnectedText','StatusFlashing','StatusFlashingText','StatusFlashingSelect')
		'Notif.'      = @('NotificationBg','NotificationMsgText','NotificationStripBg','NotificationStripWarning','NotificationStripInfo','NotificationStripError','NotificationProgressHigh','NotificationProgressMid','NotificationProgressLow','NotificationBtnBg','NotificationBtnHover','NotificationBtnPressed')
		'Tips'        = @('ToolTipBg','ToolTipBorder','ToolTipFg','TooltipBorderColor')
		'Gradients'   = @('BtnGradientNormalTop','BtnGradientNormalBot','BtnDangerGradientNormalTop','BtnDangerGradientNormalBot','BtnSuccessGradientNormalTop','BtnSuccessGradientNormalBot','BtnPrimaryGradientNormalTop','BtnPrimaryGradientNormalBot','BtnGradientHoverTop','BtnGradientHoverBot','BtnDangerGradientHoverTop','BtnDangerGradientHoverBot','BtnSuccessGradientHoverTop','BtnSuccessGradientHoverBot','BtnPrimaryGradientHoverTop','BtnPrimaryGradientHoverBot','BtnGradientLeftPressTop','BtnGradientLeftPressBot','BtnDangerGradientLeftPressTop','BtnDangerGradientLeftPressBot','BtnSuccessGradientLeftPressTop','BtnSuccessGradientLeftPressBot','BtnPrimaryGradientLeftPressTop','BtnPrimaryGradientLeftPressBot','BtnGradientRightPressTop','BtnGradientRightPressBot','BtnDangerGradientRightPressTop','BtnDangerGradientRightPressBot','BtnSuccessGradientRightPressTop','BtnSuccessGradientRightPressBot','BtnPrimaryGradientRightPressTop','BtnPrimaryGradientRightPressBot')
	}

	foreach ($sectionName in $editorSections.Keys) {
		$tab = New-Object System.Windows.Forms.TabPage $sectionName
		$tab.BackColor = $c['SecondaryBg']

		$flow = New-Object System.Windows.Forms.FlowLayoutPanel
		$flow.Dock = [System.Windows.Forms.DockStyle]::Fill
		$flow.WrapContents = $false
		$flow.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
		$flow.AutoScroll = ($editorSections[$sectionName].Count -gt 16)
		$flow.BackColor = $c['SecondaryBg']

		foreach ($key in $editorSections[$sectionName]) {
			$rowPnl = New-Object System.Windows.Forms.Panel
			$rowPnl.Size = New-Object System.Drawing.Size(290, 26)
			$rowPnl.Margin = New-Object System.Windows.Forms.Padding(2)

			$lblC = New-Object System.Windows.Forms.Label
			$lblC.Text = $key
			$lblC.Font = Get-UIFont 'Segoe UI' 7.5
			$lblC.Location = New-Object System.Drawing.Point(2, 5)
			$lblC.Size = New-Object System.Drawing.Size(190, 18)
			$lblC.ForeColor = $c['PrimaryText']
			$rowPnl.Controls.Add($lblC)

			$btnColor = New-Object System.Windows.Forms.Button
			$btnColor.Size = New-Object System.Drawing.Size(55, 20)
			$btnColor.Location = New-Object System.Drawing.Point(195, 3)
			$btnColor.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
			$btnColor.FlatAppearance.BorderSize = 1
			$btnColor.Tag = @{ Key = $key; Map = $customColorMap; OwnerForm = $themeForm }

			$curCol = if ($customColorMap.ContainsKey($key)) { $customColorMap[$key] } else { [System.Drawing.Color]::Gray }
			$btnColor.BackColor = $curCol

			$btnColor.Add_Click({
				$t = $args[0].Tag
				Show-LiveColorPicker -Button $args[0] -ColorKey $t.Key -ColorMap $t.Map -OwnerForm $t.OwnerForm -OnApply $null
			})

			$colorButtons[$key] = $btnColor
			$rowPnl.Controls.Add($btnColor)
			$flow.Controls.Add($rowPnl)
		}

		$tab.Controls.Add($flow)
		$editorTabControl.TabPages.Add($tab)
	}

	$themeForm.Controls.Add($editorTabControl)

	$btnY1 = 476
	$btnY2 = 510

	$lblPreset = New-Object System.Windows.Forms.Label
	$lblPreset.Text = 'Preset:'
	$lblPreset.Font = Get-UIFont 'Segoe UI' 8
	$lblPreset.ForeColor = $c['PrimaryText']
	$lblPreset.Location = New-Object System.Drawing.Point(352, ($btnY1 + 3))
	$lblPreset.AutoSize = $true
	$themeForm.Controls.Add($lblPreset)

	$cboPreset = New-Object System.Windows.Forms.ComboBox
	$cboPreset.Size = New-Object System.Drawing.Size(130, 24)
	$cboPreset.Location = New-Object System.Drawing.Point(402, $btnY1)
	$cboPreset.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$cboPreset.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	$cboPreset.BackColor = $c['InputBg']
	$cboPreset.ForeColor = $c['InputFg']
	$cboPreset.Font = Get-UIFont 'Segoe UI' 8
	foreach ($pn in ($allThemes | Where-Object { $_ -ne 'Custom' })) { [void]$cboPreset.Items.Add($pn) }
	$themeForm.Controls.Add($cboPreset)

	$btnApplyPreset = New-Object System.Windows.Forms.Button
	$btnApplyPreset.Text = 'Apply'
	$btnApplyPreset.Size = New-Object System.Drawing.Size(52, 24)
	$btnApplyPreset.Location = New-Object System.Drawing.Point(538, $btnY1)
	$btnApplyPreset.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$btnApplyPreset.Font = Get-UIFont 'Segoe UI' 8
	Register-ThemedButtonFadeEffect -Button $btnApplyPreset -Variant 'Normal'

	$btnApplyPreset.Add_Click({
		$frm = $args[0].FindForm()
		$ft = $frm.Tag
		$cbo = $ft.CboPreset
		if ($cbo.SelectedIndex -ge 0) {
			$selectedPreset = $cbo.SelectedItem
			$presetDict = Get-ThemeColors $selectedPreset
			if ($presetDict) {
				$cm = $ft.CustomColorMap
				$cm.Clear()
				foreach ($pk in $presetDict.Keys) { $cm[$pk] = $presetDict[$pk] }
				foreach ($entry in $ft.ColorButtons.GetEnumerator()) {
					$ck = $entry.Key; $cb = $entry.Value
					if ($cm.ContainsKey($ck)) { $cb.BackColor = $cm[$ck] }
				}
				if ($ft.ThemeCards.ContainsKey('Custom')) {
					$cc = $ft.ThemeCards['Custom']
					foreach ($ctrl in $cc.Controls) {
						if ($ctrl -is [System.Windows.Forms.PictureBox]) { $ctrl.Invalidate() }
					}
				}
			}
		}
	})
	$themeForm.Controls.Add($btnApplyPreset)

	$btnImport = New-Object System.Windows.Forms.Button
	$btnImport.Text = 'Import'
	$btnImport.Size = New-Object System.Drawing.Size(78, 26)
	$btnImport.Location = New-Object System.Drawing.Point(352, $btnY2)
	$btnImport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$btnImport.Font = Get-UIFont 'Segoe UI' 8
	Register-ThemedButtonFadeEffect -Button $btnImport -Variant 'Normal'

	$btnImport.Add_Click({
		$frm = $args[0].FindForm()
		$ft = $frm.Tag
		$ofd = New-Object System.Windows.Forms.OpenFileDialog
		$ofd.Title = 'Import Custom Theme'
		$ofd.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
		$ofd.InitialDirectory = if ($global:DashboardConfig.ConfigFilePath) { [System.IO.Path]::GetDirectoryName($global:DashboardConfig.ConfigFilePath) } else { [Environment]::GetFolderPath('MyDocuments') }
		if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
			try {
				$json = Get-Content $ofd.FileName -Raw -Encoding UTF8
				$jsonObj = $json | ConvertFrom-Json
				$cm = $ft.CustomColorMap
				$cm.Clear()
				foreach ($prop in $jsonObj.PSObject.Properties) {
					$iv = $prop.Value
					if ($iv -is [string]) {
						$parts = $iv.Split(',')
						if ($parts.Count -eq 3) {
							$cm[$prop.Name] = [System.Drawing.Color]::FromArgb([int]$parts[0], [int]$parts[1], [int]$parts[2])
						}
					}
				}
				foreach ($entry in $ft.ColorButtons.GetEnumerator()) {
					$ck = $entry.Key; $cb = $entry.Value
					if ($cm.ContainsKey($ck)) { $cb.BackColor = $cm[$ck] }
					else { $cb.BackColor = [System.Drawing.Color]::Gray }
				}
				if ($ft.ThemeCards.ContainsKey('Custom')) {
					$cc = $ft.ThemeCards['Custom']
					foreach ($ctrl in $cc.Controls) {
						if ($ctrl -is [System.Windows.Forms.PictureBox]) { $ctrl.Invalidate() }
					}
				}
			} catch {
				[System.Windows.Forms.MessageBox]::Show("Failed to import theme: $_", 'Import Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			}
		}
		$ofd.Dispose()
	})
	$themeForm.Controls.Add($btnImport)

	$btnExport = New-Object System.Windows.Forms.Button
	$btnExport.Text = 'Export'
	$btnExport.Size = New-Object System.Drawing.Size(78, 26)
	$btnExport.Location = New-Object System.Drawing.Point(436, $btnY2)
	$btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$btnExport.Font = Get-UIFont 'Segoe UI' 8
	Register-ThemedButtonFadeEffect -Button $btnExport -Variant 'Normal'

	$btnExport.Add_Click({
		$frm = $args[0].FindForm()
		$ft = $frm.Tag
		$sfd = New-Object System.Windows.Forms.SaveFileDialog
		$sfd.Title = 'Export Custom Theme'
		$sfd.Filter = 'JSON files (*.json)|*.json'
		$sfd.FileName = 'CustomTheme.json'
		$sfd.InitialDirectory = if ($global:DashboardConfig.ConfigFilePath) { [System.IO.Path]::GetDirectoryName($global:DashboardConfig.ConfigFilePath) } else { [Environment]::GetFolderPath('MyDocuments') }
		if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
			try {
				$cm = $ft.CustomColorMap
				$exportData = [ordered]@{}
				foreach ($ek in $cm.Keys) {
					$col = $cm[$ek]
					$exportData[$ek] = "$($col.R),$($col.G),$($col.B)"
				}
				($exportData | ConvertTo-Json -Depth 5) | Set-Content -Path $sfd.FileName -Encoding UTF8 -Force
			} catch {
				[System.Windows.Forms.MessageBox]::Show("Failed to export theme: $_", 'Export Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			}
		}
		$sfd.Dispose()
	})
	$themeForm.Controls.Add($btnExport)

	$btnSaveCustom = New-Object System.Windows.Forms.Button
	$btnSaveCustom.Text = 'Save & Apply Custom Theme'
	$btnSaveCustom.Font = Get-UIFont 'Segoe UI' 8 'Bold'
	$btnSaveCustom.Size = New-Object System.Drawing.Size(148, 26)
	$btnSaveCustom.Location = New-Object System.Drawing.Point(520, $btnY2)
	$btnSaveCustom.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	Register-ThemedButtonFadeEffect -Button $btnSaveCustom -Variant 'Primary'

	$btnSaveCustom.Add_Click({
		$frm = $args[0].FindForm()
		$ft = $frm.Tag
		Save-CustomThemeColors $ft.CustomColorMap
		Set-Theme 'Custom'
		#Reapply-GridStatusStyling
		$ft.EditorState.IsSavedAndApplied = $true

		if ($ft.ThemeCards.ContainsKey('Custom')) {
			$customCard = $ft.ThemeCards['Custom']
			$selectThemeHandler.Invoke($customCard, $null)
		}
	})
	$themeForm.Controls.Add($btnSaveCustom)

	$themeForm.Tag = @{
		EditorState    = $editorState
		InitialTheme   = $initialTheme
		ThemeCards     = $themeCards
		CustomColorMap = $customColorMap
		ColorButtons   = $colorButtons
		CboPreset      = $cboPreset
	}

	$themeForm.Add_FormClosing({
		$s = $args[0]
		$ft = $s.Tag
		if (-not $ft.EditorState.IsSavedAndApplied) {
			Register-CustomThemeFromConfig -Force
			Set-Theme $ft.InitialTheme
			#Reapply-GridStatusStyling
		}
	})

	$themeForm.ShowDialog() | Out-Null
	$themeForm.Dispose()
}

#endregion

#region Profile Auto-Tools Watcher (F-Tool / Macro / QuickCommand after login)

function Get-RememberedClientSection
{
	[CmdletBinding()]
	param()

	try
	{
		$config = $global:DashboardConfig.Config
		if (-not $config) { return $null }

		if (-not $config.Contains('RememberedClients') -or $null -eq $config['RememberedClients'] -or $config['RememberedClients'] -is [string])
		{
			$config['RememberedClients'] = [ordered]@{}
		}

		return $config['RememberedClients']
	}
	catch { return $null }
}

function Get-RememberedClientKey
{
	[CmdletBinding()]
	[OutputType([string])]
	param(
		[Parameter(Mandatory = $false)]
		[string]$ProfileName,
		[Parameter(Mandatory = $false)]
		[string]$ClientName
	)

	$section = Get-RememberedClientSection
	if (-not $section) { return $null }
	if ([string]::IsNullOrWhiteSpace($ProfileName) -or [string]::IsNullOrWhiteSpace($ClientName)) { return $null }

	foreach ($key in @($section.Keys))
	{
		$parts = [string]$key -split '\|', 2
		if ($parts.Count -eq 2 -and $parts[0] -ieq $ProfileName -and $parts[1] -ieq $ClientName) { return [string]$key }
	}

	return $null
}

function Get-AutoToolFlags
{
	[CmdletBinding()]
	[OutputType([hashtable])]
	param(
		[Parameter(Mandatory = $false)]
		[string]$ProfileName,
		[Parameter(Mandatory = $false)]
		[string]$ClientName
	)

	$flags = @{ Ftool = $false; Macro = $false; QuickCmd = $false }

	try
	{
		$key = Get-RememberedClientKey -ProfileName $ProfileName -ClientName $ClientName
		if (-not $key) { return $flags }

		$section = Get-RememberedClientSection
		if (-not $section) { return $flags }

		$tokens = @(([string]$section[$key]) -split ',')
		$flags.Ftool = ($tokens.Count -gt 0 -and $tokens[0].Trim() -eq '1')
		$flags.Macro = ($tokens.Count -gt 1 -and $tokens[1].Trim() -eq '1')
		$flags.QuickCmd = ($tokens.Count -gt 2 -and $tokens[2].Trim() -eq '1')
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to read client flags: $_" }

	return $flags
}

function Get-ClientWindowFlags
{
	[CmdletBinding()]
	[OutputType([hashtable])]
	param(
		[Parameter(Mandatory = $false)]
		[string]$ProfileName,
		[Parameter(Mandatory = $false)]
		[string]$ClientName
	)

	$flags = @{ Reconnect = $false; Hide = $false }

	try
	{
		$flags.Hide = Test-HideProfilesDefault

		$key = Get-RememberedClientKey -ProfileName $ProfileName -ClientName $ClientName
		if (-not $key) { return $flags }

		$section = Get-RememberedClientSection
		if (-not $section) { return $flags }

		$tokens = @(([string]$section[$key]) -split ',')
		if ($tokens.Count -gt 3) { $flags.Reconnect = ($tokens[3].Trim() -eq '1') }
		if ($tokens.Count -gt 4) { $flags.Hide = ($tokens[4].Trim() -eq '1') }
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to read window flags: $_" }

	return $flags
}

function Test-HideProfilesDefault
{
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	try
	{
		$config = $global:DashboardConfig.Config
		if (-not $config) { return $false }

		if ($config.Contains('Options'))
		{
			$options = $config['Options']
			if ($options -and $options.Contains('HideProfiles'))
			{
				return (([string]$options['HideProfiles']).Trim() -eq '1')
			}
		}

		if ($config.Contains('HideProfiles') -and $config['HideProfiles'])
		{
			$section = $config['HideProfiles']
			foreach ($flagKey in @($section.Keys))
			{
				if (([string]$section[$flagKey]).Trim() -eq '1') { return $true }
			}
		}
	}
	catch { }

	return $false
}

function Set-RememberedClientFlags
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$ProfileName,
		[Parameter(Mandatory = $true)]
		[string]$ClientName,
		[bool]$Ftool = $false,
		[bool]$Macro = $false,
		[bool]$QuickCmd = $false
	)

	try
	{
		$section = Get-RememberedClientSection
		if (-not $section) { return $false }

		$key = Get-RememberedClientKey -ProfileName $ProfileName -ClientName $ClientName
		if (-not $key) { $key = "$ProfileName|$ClientName" }

		$f = if ($Ftool) { '1' } else { '0' }
		$m = if ($Macro) { '1' } else { '0' }
		$q = if ($QuickCmd) { '1' } else { '0' }

		$windowFlags = Get-ClientWindowFlags -ProfileName $ProfileName -ClientName $ClientName
		$r = if ($windowFlags.Reconnect) { '1' } else { '0' }
		$h = if ($windowFlags.Hide) { '1' } else { '0' }

		$section[$key] = "$f,$m,$q,$r,$h"
		return $true
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to store client flags: $_"; return $false }
}

function Remove-RememberedClient
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$ProfileName,
		[Parameter(Mandatory = $true)]
		[string]$ClientName
	)

	try
	{
		$section = Get-RememberedClientSection
		if (-not $section) { return $false }

		$key = Get-RememberedClientKey -ProfileName $ProfileName -ClientName $ClientName
		if (-not $key) { return $false }

		$section.Remove($key)

		if ((Get-SelectedClientForProfile -ProfileName $ProfileName) -ieq $ClientName)
		{
			Set-SelectedClientForProfile -ProfileName $ProfileName -ClientName '' | Out-Null
		}
		return $true
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to remove client: $_"; return $false }
}

function Rename-RememberedClient
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$ProfileName,
		[Parameter(Mandatory = $true)]
		[string]$OldName,
		[Parameter(Mandatory = $true)]
		[string]$NewName
	)

	try
	{
		if ([string]::IsNullOrWhiteSpace($NewName)) { return $false }

		$section = Get-RememberedClientSection
		if (-not $section) { return $false }

		$oldKey = Get-RememberedClientKey -ProfileName $ProfileName -ClientName $OldName
		if (-not $oldKey) { return $false }

		$value = [string]$section[$oldKey]
		$section.Remove($oldKey)

		$newKey = Get-RememberedClientKey -ProfileName $ProfileName -ClientName $NewName
		if ($newKey) { return $true }

		$section["$ProfileName|$NewName"] = $value

		if ((Get-SelectedClientForProfile -ProfileName $ProfileName) -ieq $OldName)
		{
			Set-SelectedClientForProfile -ProfileName $ProfileName -ClientName $NewName | Out-Null
		}
		return $true
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to rename client: $_"; return $false }
}

function Remove-RememberedClientsProfile
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$ProfileName
	)

	try
	{
		$section = Get-RememberedClientSection
		if (-not $section) { return $false }

		$removed = $false
		foreach ($key in @($section.Keys))
		{
			$parts = [string]$key -split '\|', 2
			if ($parts.Count -eq 2 -and $parts[0] -ieq $ProfileName)
			{
				$section.Remove($key)
				$removed = $true
			}
		}

		Remove-SelectedProfileEntry -ProfileName $ProfileName | Out-Null

		if ($removed) { SyncConfigToClientGrid | Out-Null }
		return $removed
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to drop profile clients: $_"; return $false }
}

function Rename-RememberedClientsProfile
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$OldProfile,
		[Parameter(Mandatory = $true)]
		[string]$NewProfile
	)

	try
	{
		if ([string]::IsNullOrWhiteSpace($NewProfile)) { return $false }

		$section = Get-RememberedClientSection
		if (-not $section) { return $false }

		$moved = $false
		foreach ($key in @($section.Keys))
		{
			$parts = [string]$key -split '\|', 2
			if ($parts.Count -ne 2 -or -not ($parts[0] -ieq $OldProfile)) { continue }

			$value = $section[$key]
			$section.Remove($key)

			$newKey = "$NewProfile|$($parts[1])"
			if (-not $section.Contains($newKey)) { $section[$newKey] = $value }
			$moved = $true
		}

		Rename-SelectedProfileEntry -OldProfile $OldProfile -NewProfile $NewProfile | Out-Null

		if ($moved) { SyncConfigToClientGrid | Out-Null }
		return $moved
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to rename profile clients: $_"; return $false }
}

function Get-SelectedClientSection
{
	$config = $global:DashboardConfig.Config
	if (-not $config) { return $null }
	if (-not $config.Contains('SelectedClients') -or $null -eq $config['SelectedClients'] -or $config['SelectedClients'] -is [string])
	{
		$config['SelectedClients'] = [ordered]@{}
	}
	return $config['SelectedClients']
}

function Get-SelectedClientForProfile
{
	[CmdletBinding()]
	[OutputType([string])]
	param(
		[Parameter(Mandatory = $false)]
		[string]$ProfileName
	)

	try
	{
		$section = Get-SelectedClientSection
		if (-not $section -or [string]::IsNullOrWhiteSpace($ProfileName)) { return '' }

		foreach ($key in @($section.Keys))
		{
			if ([string]$key -ieq $ProfileName) { return ([string]$section[$key]).Trim() }
		}
	}
	catch { }

	return ''
}

function Set-SelectedClientForProfile
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $false)]
		[string]$ProfileName,
		[Parameter(Mandatory = $false)]
		[string]$ClientName
	)

	try
	{
		$section = Get-SelectedClientSection
		if (-not $section -or [string]::IsNullOrWhiteSpace($ProfileName)) { return $false }

		$existingKey = $null
		foreach ($key in @($section.Keys))
		{
			if ([string]$key -ieq $ProfileName) { $existingKey = $key; break }
		}

		if ([string]::IsNullOrWhiteSpace($ClientName))
		{
			if ($null -ne $existingKey) { $section.Remove($existingKey) }
			return $true
		}

		if ($null -ne $existingKey) { $section[$existingKey] = $ClientName.Trim() }
		else { $section[$ProfileName] = $ClientName.Trim() }
		return $true
	}
	catch { Write-Verbose "  UI: Failed to store selected client: $_"; return $false }
}

function Rename-SelectedProfileEntry
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$OldProfile,
		[Parameter(Mandatory = $true)]
		[string]$NewProfile
	)

	try
	{
		if ([string]::IsNullOrWhiteSpace($NewProfile)) { return $false }

		$section = Get-SelectedClientSection
		if (-not $section) { return $false }

		foreach ($key in @($section.Keys))
		{
			if ([string]$key -ieq $OldProfile)
			{
				$value = $section[$key]
				$section.Remove($key)
				$section[$NewProfile] = $value
				return $true
			}
		}
	}
	catch { }

	return $false
}

function Remove-SelectedProfileEntry
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$ProfileName
	)

	try
	{
		$section = Get-SelectedClientSection
		if (-not $section) { return $false }

		foreach ($key in @($section.Keys))
		{
			if ([string]$key -ieq $ProfileName) { $section.Remove($key); return $true }
		}
	}
	catch { }

	return $false
}

function Restore-ClientGridSelection
{
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	try
	{
		$UI = $global:DashboardConfig.UI
		if (-not ($UI -and $UI.ClientGrid -and $UI.ClientGrid.Rows.Count -gt 0)) { return $false }

		$selectedProfile = ''
		$config = $global:DashboardConfig.Config
		if ($config -and $config.Contains('Options') -and $config['Options'] -and $config['Options'].Contains('SelectedProfile'))
		{
			$selectedProfile = [string]$config['Options']['SelectedProfile']
		}
		$selectedClient = Get-SelectedClientForProfile -ProfileName $selectedProfile

		$targetRow = $null
		if (-not [string]::IsNullOrWhiteSpace($selectedProfile) -and -not [string]::IsNullOrWhiteSpace($selectedClient))
		{
			foreach ($row in $UI.ClientGrid.Rows)
			{
				if ([string]$row.Cells[0].Value -ieq $selectedProfile -and [string]$row.Cells[1].Value -ieq $selectedClient)
				{
					$targetRow = $row
					break
				}
			}
		}

		if (-not $targetRow) { return $false }

		$UI.ClientGrid.ClearSelection()
		$targetRow.Selected = $true
		try { $UI.ClientGrid.CurrentCell = $targetRow.Cells[1] } catch {}
		$UI.ClientGrid.Tag = $targetRow.Index
		return $true
	}
	catch { return $false }
}

function RememberProfileClient
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $false)]
		[string]$ProfileName,
		[Parameter(Mandatory = $false)]
		[string]$ClientName
	)

	try
	{
		if ([string]::IsNullOrWhiteSpace($ProfileName) -or [string]::IsNullOrWhiteSpace($ClientName)) { return $false }

		$section = Get-RememberedClientSection
		if (-not $section) { return $false }
		if (Get-RememberedClientKey -ProfileName $ProfileName -ClientName $ClientName) { return $false }

		$windowFlags = Get-ClientWindowFlags -ProfileName $ProfileName -ClientName $ClientName
		$r = if ($windowFlags.Reconnect) { '1' } else { '0' }
		$h = if ($windowFlags.Hide) { '1' } else { '0' }

		$section["$ProfileName|$ClientName"] = "0,0,0,$r,$h"
		Write-Verbose "AUTO-TOOLS: remembered client '$ClientName' of profile '$ProfileName'"

		if (Test-DashboardCommand 'WriteConfig') { WriteConfig }

		try { SyncConfigToClientGrid | Out-Null } catch {}

		return $true
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to remember client: $_"; return $false }
}

function Get-AutoToolRowInfo
{
	[CmdletBinding()]
	param($Row)

	try
	{
		if (-not $Row -or -not $Row.Cells -or $Row.Cells.Count -lt 4) { return $null }

		$titleObj = $Row.Cells[1].Value
		$pidObj = $Row.Cells[2].Value
		if ($null -eq $titleObj -or $null -eq $pidObj) { return $null }

		$fullTitle = $titleObj.ToString()
		$pidStr = $pidObj.ToString()
		if ([string]::IsNullOrWhiteSpace($fullTitle) -or [string]::IsNullOrWhiteSpace($pidStr)) { return $null }

		$profileName = 'Default'
		$cleanTitle = $fullTitle
		if ($fullTitle -match '^\[(?<pName>[^\]]+)\](?<rest>.*)')
		{
			$profileName = $Matches['pName']
			$cleanTitle = $Matches['rest']
		}
		elseif ($Row.Tag -and $Row.Tag.PSObject.Properties['Profile'] -and -not [string]::IsNullOrWhiteSpace([string]$Row.Tag.Profile))
		{
			$profileName = [string]$Row.Tag.Profile
		}

		$status = ''
		try { $status = [string]$Row.Cells[3].Value } catch {}

		$isLoggedIn = ($cleanTitle -match ' - ' -and $status -ne 'Exited')

		$clientName = $cleanTitle.Trim()

		return @{ ProfileName = $profileName; Client = $clientName; Pid = $pidStr; IsLoggedIn = $isLoggedIn; Status = $status }
	}
	catch { return $null }
}

function Reset-AutoToolStateForClient
{
	[CmdletBinding()]
	[OutputType([void])]
	param(
		[Parameter(Mandatory = $false)]
		[string]$ProfileName,
		[Parameter(Mandatory = $false)]
		[string]$ClientName
	)

	try
	{
		$state = $global:DashboardConfig.State
		$ui = $global:DashboardConfig.UI
		if (-not $state -or -not $ui -or -not $ui.DataGridMain) { return }
		if ([string]::IsNullOrWhiteSpace($ProfileName) -or [string]::IsNullOrWhiteSpace($ClientName)) { return }

		$matchedPid = $null
		foreach ($row in @($ui.DataGridMain.Rows))
		{
			$info = Get-AutoToolRowInfo -Row $row
			if ($info -and $info.ProfileName -ieq $ProfileName -and $info.Client -ieq $ClientName)
			{
				$matchedPid = $info.Pid
				break
			}
		}

		if (-not $matchedPid) { return }

		$changed = $false
		if ($state.AutoToolsStartedPids -and $state.AutoToolsStartedPids.Contains($matchedPid))
		{
			[void]$state.AutoToolsStartedPids.Remove($matchedPid)
			$changed = $true
		}
		if ($state.AutoToolsAttempts -and $state.AutoToolsAttempts.Contains($matchedPid))
		{
			$state.AutoToolsAttempts.Remove($matchedPid)
			$changed = $true
		}
		if ($changed)
		{
			$state.AutoToolsLastDiag = ''
			Write-Verbose "AUTO-TOOLS: re-armed '$ProfileName/$ClientName' (pid $matchedPid) after a settings change"
		}
	}
	catch { Write-Verbose "AUTO-TOOLS: Failed to re-arm client: $_" }
}

function Test-AutoToolMacroConfigured
{
	[CmdletBinding()]
	[OutputType([bool])]
	param($FormData)

	try
	{
		if (-not $FormData) { return $false }

		$keys = @()
		$mainKey = [string]$FormData.BtnKeySelect.Text
		if ($mainKey) { $keys += $mainKey }
		if ($FormData.SequencePanels)
		{
			foreach ($panel in @($FormData.SequencePanels))
			{
				if ($panel -and -not $panel.IsDisposed -and $panel.Tag)
				{
					$seqKey = [string]$panel.Tag.BtnKeySelect.Text
					if ($seqKey) { $keys += $seqKey }
				}
			}
		}

		$keys = @($keys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne 'none' -and $_.Trim() -ne 'Select' })
		if ($keys.Count -eq 0) { return $false }

		if ((Test-DashboardCommand 'GetVirtualKeyMappings') -and (Test-DashboardCommand 'ParseKeyString'))
		{
			$keyMappings = GetVirtualKeyMappings
			foreach ($k in $keys)
			{
				$parsedKey = ParseKeyString -KeyCombinationString $k
				if ($parsedKey -and $parsedKey.Primary -and $null -ne $keyMappings[$parsedKey.Primary]) { return $true }
			}
			return $false
		}

		return $true
	}
	catch { return $true }
}

function Start-AutoToolInstance
{
	[CmdletBinding()]
	[OutputType([string])]
	param(
		[Parameter(Mandatory = $true)][string]$Kind,
		[Parameter(Mandatory = $true)]$Form,
		[Parameter(Mandatory = $true)]$Notes
	)

	if (-not $Form -or $Form.IsDisposed) { return 'pending' }
	$data = $Form.Tag
	if (-not $data) { return 'pending' }

	try
	{
		switch ($Kind)
		{
			'Ftool' {
				if ($data.RunningSpammer) { return 'ok' }

				if (Test-DashboardCommand 'FindOrCreateProfile')
				{
					$templateKey = "Template_$(FindOrCreateProfile $data.WindowTitle -ProfileHint (Get-RowProfileHint -formData $data))"
					$ftoolCfg = $global:DashboardConfig.Config['Ftool']
					if ($ftoolCfg -and $templateKey -and $templateKey -ne 'Template_')
					{
						$templateCfg = $ftoolCfg[$templateKey]
						if ($templateCfg)
						{
							if ($templateCfg.Contains('key1')) { $data.BtnKeySelect.Text = $templateCfg['key1'] }
							if ($templateCfg.Contains('inpt1')) { $data.Interval.Text = $templateCfg['inpt1'] }
							if ($templateCfg.Contains('name1')) { $data.Name.Text = $templateCfg['name1'] }
							try { UpdateSettings $data -forceWrite } catch {}
							$Notes.Add('ftool: applied saved template')
						}
					}
				}

				$keyValue = [string]$data.BtnKeySelect.Text
				if ([string]::IsNullOrWhiteSpace($keyValue) -or $keyValue.Trim() -eq 'Select')
				{
					$Notes.Add('ftool: not started (no key configured)')
					return 'ok'
				}

				$intervalNum = 0
				if (-not [int]::TryParse([string]$data.Interval.Text, [ref]$intervalNum) -or $intervalNum -lt 10) { $intervalNum = 100 }

				if ((Test-DashboardCommand 'ToggleButtonState') -and (Test-DashboardCommand 'CreateSpammerTimer'))
				{
					ToggleButtonState $data.BtnStart $data.BtnStop $true
					$spamTimer = CreateSpammerTimer $data.SelectedWindow $keyValue $data.InstanceId $intervalNum
					$data.RunningSpammer = $spamTimer
					$global:DashboardConfig.Resources.Timers["ExtSpammer_$($data.InstanceId)_1"] = $spamTimer
					$Notes.Add("ftool: started ($keyValue)")
					return 'ok'
				}
				return 'pending'
			}

			'Macro' {
				if ($data.RunningTimer) { return 'ok' }
				if (-not (Test-AutoToolMacroConfigured -FormData $data))
				{
					$Notes.Add('macro: not started (no key configured)')
					return 'ok'
				}
				StartMacroSequence $data
				if ($data.RunningTimer) { $Notes.Add('macro: started'); return 'ok' }
				return 'pending'
			}

			'QuickCmd' {
				if ($data.RunningTimer) { return 'ok' }

				if ($data.PSObject.Properties['ChkRunOnStart'] -and $data.ChkRunOnStart -and -not $data.ChkRunOnStart.Checked)
				{
					$Notes.Add('quickcmd: not started (run on open disabled)')
					return 'ok'
				}

				$savedCount = 0
				if ($data.PSObject.Properties['panelSavedCommands'] -and $data.panelSavedCommands)
				{
					foreach ($control in $data.panelSavedCommands.Controls)
					{
						if ($control -is [System.Windows.Forms.Panel] -and $control.Tag -is [PSCustomObject]) { $savedCount++ }
					}
				}
				if ($savedCount -eq 0)
				{
					$Notes.Add('quickcmd: not started (no saved commands)')
					return 'ok'
				}

				StartChatCommanderSequence $data -RunSavedSequence -AutoStart
				if ($data.RunningTimer) { $Notes.Add("quickcmd: started ($savedCount command(s))"); return 'ok' }
				return 'pending'
			}
		}
	}
	catch { return 'pending' }

	return 'ok'
}

function Invoke-ProfileAutoToolsPass
{
	[CmdletBinding()]
	param()

	$state = $global:DashboardConfig.State
	$ui = $global:DashboardConfig.UI
	if (-not $state -or -not $ui -or -not $ui.DataGridMain -or $ui.DataGridMain.IsDisposed) { return }

	$startedPids = $state.AutoToolsStartedPids
	$attempts = $state.AutoToolsAttempts
	if (-not $startedPids -or -not $attempts) { return }

	$loggedInPids = New-Object 'System.Collections.Generic.HashSet[string]'
	if (-not $state.AutoToolsSeenPids) { $state.AutoToolsSeenPids = New-Object 'System.Collections.Generic.HashSet[string]' }
	$seenPids = $state.AutoToolsSeenPids
	foreach ($row in @($ui.DataGridMain.Rows))
	{
		$info = Get-AutoToolRowInfo -Row $row
		if ($info -and $info.IsLoggedIn)
		{
			[void]$loggedInPids.Add($info.Pid)
			[void]$seenPids.Add($info.Pid)

			RememberProfileClient -ProfileName $info.ProfileName -ClientName $info.Client | Out-Null
		}
	}

	foreach ($stalePid in @($startedPids))
	{
		if (-not $loggedInPids.Contains($stalePid))
		{
			[void]$startedPids.Remove($stalePid)
			if ($attempts.Contains($stalePid)) { $attempts.Remove($stalePid) }
		}
	}

	$notes = New-Object 'System.Collections.Generic.List[string]'

	$blockReason = ''
	if ($state.LoginActive) { $blockReason = 'LoginActive' }
	elseif ($state.LaunchActive) { $blockReason = 'LaunchActive' }
	elseif ($state.SequenceActive) { $blockReason = 'SequenceActive' }

	if ($blockReason)
	{
		$notes.Add("blocked:$blockReason")
	}
	else
	{
		foreach ($row in @($ui.DataGridMain.Rows))
		{
			$info = Get-AutoToolRowInfo -Row $row
			if (-not $info) { continue }

			$flags = Get-AutoToolFlags -ProfileName $info.ProfileName -ClientName $info.Client
			$anyFlag = $flags.Ftool -or $flags.Macro -or $flags.QuickCmd

			if (-not $info.IsLoggedIn)
			{
				if ($anyFlag) { $notes.Add("waiting:$($info.ProfileName)/$($info.Pid) logged-out") }
				continue
			}

			if (-not $anyFlag)
			{
				$notes.Add("no-flags:$($info.ProfileName)/$($info.Client) ($($info.Pid))")
				continue
			}

			if ($startedPids.Contains($info.Pid)) { continue }

			if ($state.ActiveLoginPids -and $state.ActiveLoginPids.Count -gt 0)
			{
				$pidInt = 0
				if ([int]::TryParse($info.Pid, [ref]$pidInt) -and $state.ActiveLoginPids.Contains($pidInt))
				{
					$notes.Add("waiting:$($info.ProfileName)/$($info.Pid) logging-in")
					continue
				}
			}

			$toolForms = $global:DashboardConfig.Resources.FtoolForms
			$expectedIds = @()
			$requested = @()
			$macroId = $null
			$qcId = $null
			$openError = ''

			try
			{
				if ($flags.Ftool -and (Test-DashboardCommand 'FtoolSelectedRow'))
				{
					$expectedIds += $info.Pid
					$requested += 'Ftool'
					if (-not ($toolForms -and $toolForms.Contains($info.Pid))) { FtoolSelectedRow $row }
				}

				if ($flags.Macro -and (Test-DashboardCommand 'MacroSelectedRow'))
				{
					$macroId = 'Macro_' + ((Get-RowIdentity -Row $row) -replace '[^a-zA-Z0-9_\-]', '_')
					$requested += 'Macro'

					$sessionCount = 0
					$macroCfg = $global:DashboardConfig.Config['Macro']
					if ($macroCfg -and $macroCfg -isnot [string] -and $macroCfg.Contains("Session_Macro_$macroId"))
					{
						try { $sessionCount = [int]$macroCfg["Session_Macro_$macroId"] } catch { $sessionCount = 0 }
						if ($sessionCount -gt 50) { $sessionCount = 0 }
					}

					if ($sessionCount -gt 0)
					{
						$notes.Add("macro: restoring $sessionCount saved instance(s) for $($info.ProfileName)/$($info.Pid)")
						for ($macroIdx = 1; $macroIdx -le $sessionCount; $macroIdx++)
						{
							$macroInstanceId = "${macroId}_${macroIdx}"
							$expectedIds += $macroInstanceId
							if (-not ($toolForms -and $toolForms.Contains($macroInstanceId))) { MacroSelectedRow $row -ForceNew -InstanceIndex $macroIdx }
						}
					}
					else
					{
						$expectedIds += $macroId
						if (-not ($toolForms -and $toolForms.Contains($macroId))) { MacroSelectedRow $row }
					}
				}

				if ($flags.QuickCmd -and (Test-DashboardCommand 'ChatCommanderSelectedRow'))
				{
					$qcId = 'ChatCommander_' + $info.Pid
					$expectedIds += $qcId
					$requested += 'QuickCmd'
					if (-not ($toolForms -and $toolForms.Contains($qcId))) { ChatCommanderSelectedRow $row }
				}
			}
			catch { $openError = $_.Exception.Message }

			$pending = $false
			foreach ($formId in $expectedIds)
			{
				if (-not ($toolForms -and $toolForms.Contains($formId))) { $pending = $true; break }
			}

			$startPending = $false
			if (-not $pending)
			{
				if ($flags.Ftool -and $toolForms -and $toolForms.Contains($info.Pid))
				{
					if ((Start-AutoToolInstance -Kind 'Ftool' -Form $toolForms[$info.Pid] -Notes $notes) -eq 'pending') { $startPending = $true }
				}
				if ($flags.QuickCmd -and $qcId -and $toolForms -and $toolForms.Contains($qcId))
				{
					if ((Start-AutoToolInstance -Kind 'QuickCmd' -Form $toolForms[$qcId] -Notes $notes) -eq 'pending') { $startPending = $true }
				}
			}

			if (-not $pending -and -not $startPending)
			{
				[void]$startedPids.Add($info.Pid)
				if ($attempts.Contains($info.Pid)) { $attempts.Remove($info.Pid) }
				$notes.Add("done:$($info.ProfileName)/$($info.Pid) [$($requested -join '+')]")
			}
			else
			{
				$attemptCount = 1
				if ($attempts.Contains($info.Pid)) { $attemptCount = [int]$attempts[$info.Pid] + 1 }
				$attempts[$info.Pid] = $attemptCount

				if ($attemptCount -ge 5)
				{
					$reason = if ($openError) { " ($openError)" } else { '' }
					$notes.Add("give-up:$($info.ProfileName)/$($info.Pid) after $attemptCount attempts$reason")
					[void]$startedPids.Add($info.Pid)
					$attempts.Remove($info.Pid)
				}
				else
				{
					$notes.Add("retry:$($info.ProfileName)/$($info.Pid) #$attemptCount (open=$(-not $pending), start=$(-not $startPending))")
				}
			}
		}
	}

	$diag = if ($notes.Count -gt 0) { $notes -join ' | ' } else { 'idle' }
	if ($state.AutoToolsLastDiag -ne $diag)
	{
		$state.AutoToolsLastDiag = $diag
		Write-Verbose "AUTO-TOOLS: $diag"
	}
}

function Invoke-ProfileAutoToolsCleanup
{
	[CmdletBinding()]
	param()

	$state = $global:DashboardConfig.State
	$ui = $global:DashboardConfig.UI
	if (-not $state -or -not $ui -or -not $ui.DataGridMain -or $ui.DataGridMain.IsDisposed) { return }

	if ($state.LoginActive -or $state.LaunchActive -or $state.SequenceActive) { return }

	$toolForms = $global:DashboardConfig.Resources.FtoolForms
	if (-not $toolForms -or $toolForms.Count -eq 0) { return }

	$seenPids = $state.AutoToolsSeenPids
	if (-not $seenPids)
	{
		$seenPids = New-Object 'System.Collections.Generic.HashSet[string]'
		$state.AutoToolsSeenPids = $seenPids
	}

	$byPid = @{}
	foreach ($row in @($ui.DataGridMain.Rows))
	{
		$info = Get-AutoToolRowInfo -Row $row
		if ($info) { $byPid[$info.Pid] = $info }
	}

	$closed = New-Object 'System.Collections.Generic.List[string]'

	foreach ($key in @($toolForms.Keys))
	{
		$form = $null
		try { $form = $toolForms[$key] } catch { continue }
		if (-not $form -or $form.IsDisposed) { continue }

		$data = $null
		try { $data = $form.Tag } catch {}
		if (-not $data) { continue }

		$kind = $null
		if ($key -match '^\d+(_\d+)?$') { $kind = 'Ftool' }
		elseif ($key -like 'ChatCommander_*') { $kind = 'QuickCmd' }
		elseif ($key -like 'Macro_*') { $kind = 'Macro' }
		if (-not $kind) { continue }

		$ownerPid = ''
		if ($data.PSObject.Properties['Row'] -and $data.Row -and $data.Row.Cells -and $data.Row.Cells.Count -gt 2)
		{
			try { $v = $data.Row.Cells[2].Value; if ($null -ne $v) { $ownerPid = [string]$v } } catch {}
		}
		if ([string]::IsNullOrWhiteSpace($ownerPid) -and $key -match '^(?:ChatCommander_)?(\d+)(?:_\d+)?$') { $ownerPid = $Matches[1] }
		if ([string]::IsNullOrWhiteSpace($ownerPid)) { continue }

		$entry = $null
		if ($byPid.ContainsKey($ownerPid)) { $entry = $byPid[$ownerPid] }

		$reason = $null
		if ($entry)
		{
			if ($entry.Status -eq 'Exited') { $reason = 'client closed' }
			elseif (-not $entry.IsLoggedIn -and $seenPids.Contains($ownerPid)) { $reason = 'client logged off' }
		}
		else
		{
			$alive = $false
			try { $null = [System.Diagnostics.Process]::GetProcessById([int]$ownerPid); $alive = $true } catch {}
			if (-not $alive) { $reason = 'client closed' }
		}

		if (-not $reason) { continue }

		try
		{
			if ($data.PSObject.Properties['BtnClose'] -and $data.BtnClose)
			{
				$data.BtnClose.PerformClick()
			}
			elseif ($kind -eq 'Macro' -and (Test-DashboardCommand 'StopMacroForm')) { StopMacroForm $form }
			elseif ($kind -eq 'QuickCmd' -and (Test-DashboardCommand 'StopChatCommanderForm')) { StopChatCommanderForm $form }
			else
			{
				if (Test-DashboardCommand 'CleanupInstanceResources') { try { CleanupInstanceResources $key } catch {} }
				try { if ($toolForms.Contains($key)) { $toolForms.Remove($key) } } catch {}
				try { $form.Close() } catch {}
				try { $form.Dispose() } catch {}
			}
			$closed.Add("$kind '$key' -> $reason")
		}
		catch { $closed.Add("$kind '$key' -> close failed: $($_.Exception.Message)") }
	}

	if ($closed.Count -gt 0)
	{
		$diag = $closed -join ' | '
		if ($state.AutoToolsLastCleanup -ne $diag)
		{
			$state.AutoToolsLastCleanup = $diag
			Write-Verbose "AUTO-TOOLS: closed $diag"
		}
	}
}

function StartProfileAutoToolsWatcher
{
	[CmdletBinding()]
	param()

	if (-not $global:DashboardConfig -or -not $global:DashboardConfig.Resources) { return }

	$timers = $global:DashboardConfig.Resources.Timers
	if ($timers -and $timers.Contains('ProfileAutoTools')) { return }

	$state = $global:DashboardConfig.State
	if (-not $state.AutoToolsStartedPids) { $state.AutoToolsStartedPids = New-Object 'System.Collections.Generic.HashSet[string]' }
	if (-not $state.AutoToolsAttempts) { $state.AutoToolsAttempts = @{} }

	$timer = New-Object System.Windows.Forms.Timer
	$timer.Interval = 2000
	$timer.Add_Tick({
		if ($global:ProfileAutoToolsBusy) { return }
		$global:ProfileAutoToolsBusy = $true
		try
		{
			try { Invoke-ProfileAutoToolsPass }
			catch { Write-Verbose "AUTO-TOOLS: $($_.Exception.Message)" }
			try { Invoke-ProfileAutoToolsCleanup }
			catch { Write-Verbose "AUTO-TOOLS: cleanup: $($_.Exception.Message)" }
		}
		finally { $global:ProfileAutoToolsBusy = $false }
	})

	$timers['ProfileAutoTools'] = $timer
	$timer.Start()
	Write-Verbose 'Profile Auto-Tools watcher started (opens F-Tool / Macro / QuickCommand after login).'
}

function StopProfileAutoToolsWatcher
{
	[CmdletBinding()]
	param()

	$timers = $global:DashboardConfig.Resources.Timers
	if ($timers -and $timers.Contains('ProfileAutoTools'))
	{
		$timer = $timers['ProfileAutoTools']
		try { $timer.Stop(); $timer.Dispose() } catch {}
		try { $timers.Remove('ProfileAutoTools') } catch {}
	}
	$global:ProfileAutoToolsBusy = $false
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion