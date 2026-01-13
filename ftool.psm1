<# ftool.psm1 #>

#region Hotkey Management

if (-not $global:RegisteredHotkeys) { $global:RegisteredHotkeys = @{} }
if (-not $global:RegisteredHotkeyByString) { $global:RegisteredHotkeyByString = @{} }

function ToggleSpecificFtoolInstance
{
	param(
		[Parameter(Mandatory = $true)]
		[string]$InstanceId,
		[Parameter(Mandatory = $false)]
		[string]$ExtKey
	)

	if (-not $global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId))
	{
		return
	}

	$form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
	if ($form -and -not $form.IsDisposed)
	{
		if ($form.InvokeRequired)
		{
			$form.BeginInvoke([System.Action]{ ToggleSpecificFtoolInstance -InstanceId $InstanceId -ExtKey $ExtKey }) | Out-Null
			return
		}

		$formData = $form.Tag

		if (-not [string]::IsNullOrEmpty($ExtKey))
		{
			if ($global:DashboardConfig.Resources.ExtensionData.Contains($ExtKey))
			{
				$extData = $global:DashboardConfig.Resources.ExtensionData[$ExtKey]
				if ($extData.RunningSpammer)
				{
					$extData.BtnStop.PerformClick()
				}
				else
				{
					$extData.BtnStart.PerformClick()
				}
			}
		}
		else
		{
			if ($formData.RunningSpammer)
			{
				$formData.BtnStop.PerformClick()
			}
			else
			{
				$formData.BtnStart.PerformClick()
			}
		}
	}
}

function ToggleInstanceHotkeys
{
	param(
		[Parameter(Mandatory = $true)]
		[string]$InstanceId,
		[Parameter(Mandatory = $true)]
		[bool]$ToggleState
	)

	if ($global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId))
	{
		$form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
		if ($form -and -not $form.IsDisposed -and $form.InvokeRequired)
		{
			$form.BeginInvoke([System.Action]{ ToggleInstanceHotkeys -InstanceId $InstanceId -ToggleState $ToggleState }) | Out-Null
			return
		}
	}
    
	if (-not $global:DashboardConfig.Resources.InstanceHotkeysPaused) { $global:DashboardConfig.Resources.InstanceHotkeysPaused = @{} }
    
	$global:DashboardConfig.Resources.InstanceHotkeysPaused[$InstanceId] = (-not $ToggleState)

	try
	{
		if ($ToggleState)
		{
			try { ResumeHotkeysForOwner -OwnerKey $InstanceId } catch {}
			if ($global:DashboardConfig.Resources.ExtensionData)
			{
				$extKeys = $global:DashboardConfig.Resources.ExtensionData.Keys | Where-Object { $_ -like "ext_${InstanceId}_*" }
				foreach ($extKey in $extKeys)
				{
					try { ResumeHotkeysForOwner -OwnerKey $extKey } catch {}
				}
			}
		}
		else
		{
			try { PauseHotkeysForOwner -OwnerKey $InstanceId } catch {}
			if ($global:DashboardConfig.Resources.ExtensionData)
			{
				$extKeys = $global:DashboardConfig.Resources.ExtensionData.Keys | Where-Object { $_ -like "ext_${InstanceId}_*" }
				foreach ($extKey in $extKeys)
				{
					try { PauseHotkeysForOwner -OwnerKey $extKey } catch {}
				}
			}
		}
	}
 catch {}

	if ($global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId))
	{
		$form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
		if ($form -and -not $form.IsDisposed)
		{
			UpdateSettings -formData $form.Tag -forceWrite
		}
	}
}


#endregion

#region Helper Functions


function Invoke-FtoolSpamAction
{
	param($timerData)
	try
	{
		if (-not $timerData -or $timerData['WindowHandle'] -eq [IntPtr]::Zero) { return }

		$keyCombinationString = $timerData['Key']
		$parsedKey = ParseKeyString -KeyCombinationString $keyCombinationString
		if (-not $parsedKey -or -not $parsedKey.Primary)
		{
			Write-Verbose "FTOOL: Invalid key to spam: '$keyCombinationString'"
			return
		}

		$primaryKeyName = $parsedKey.Primary
		$modifierNames = $parsedKey.Modifiers

		$keyMappings = GetVirtualKeyMappings
		$virtualKeyCode = $keyMappings[$primaryKeyName]

		if ($virtualKeyCode)
		{
			$WM_KEYDOWN = 0x0100; $WM_KEYUP = 0x0101
			$WM_LBUTTONDOWN = 0x0201; $WM_LBUTTONUP = 0x0202
			$WM_RBUTTONDOWN = 0x0204; $WM_RBUTTONUP = 0x0205
			$WM_MBUTTONDOWN = 0x0207; $WM_MBUTTONUP = 0x0208
			$WM_XBUTTONDOWN = 0x020B; $WM_XBUTTONUP = 0x020C

			$point = New-Object Custom.Win32Point
			[Custom.Win32MouseUtils]::GetCursorPos([ref]$point) | Out-Null
			[Custom.Win32MouseUtils]::ScreenToClient($timerData['WindowHandle'], [ref]$point) | Out-Null
			$lParam = [Custom.Win32MouseUtils]::MakeLParam($point.X, $point.Y)

			$dwellTime = Get-Random -Minimum 2 -Maximum 9

			if ($virtualKeyCode -ge 0x01 -and $virtualKeyCode -le 0x06)
			{
				$MK_LBUTTON = 0x0001; $MK_RBUTTON = 0x0002; $MK_SHIFT = 0x0004
				$MK_CONTROL = 0x0008; $MK_MBUTTON = 0x0010; $MK_XBUTTON1 = 0x0020
				$MK_XBUTTON2 = 0x0040

				$wDown = 0

				if ($modifierNames -contains 'Ctrl') { $wDown = $wDown -bor $MK_CONTROL }
				if ($modifierNames -contains 'Shift') { $wDown = $wDown -bor $MK_SHIFT }

				$targetButtonFlag = 0; $xButtonData = 0
				$msgDown = 0; $msgUp = 0

				switch ($virtualKeyCode)
				{
					0x01 { $targetButtonFlag = $MK_LBUTTON; $msgDown = $WM_LBUTTONDOWN; $msgUp = $WM_LBUTTONUP }
					0x02 { $targetButtonFlag = $MK_RBUTTON; $msgDown = $WM_RBUTTONDOWN; $msgUp = $WM_RBUTTONUP }
					0x04 { $targetButtonFlag = $MK_MBUTTON; $msgDown = $WM_MBUTTONDOWN; $msgUp = $WM_MBUTTONUP }
					0x05 { $targetButtonFlag = $MK_XBUTTON1; $xButtonData = 0x00010000; $msgDown = $WM_XBUTTONDOWN; $msgUp = $WM_XBUTTONUP }
					0x06 { $targetButtonFlag = $MK_XBUTTON2; $xButtonData = 0x00020000; $msgDown = $WM_XBUTTONDOWN; $msgUp = $WM_XBUTTONUP }
				}

				$wDown = $wDown -bor $targetButtonFlag
				$wUp = $wDown -bxor $targetButtonFlag

				if ($xButtonData -ne 0)
				{
					$wDown = $wDown -bor $xButtonData
					$wUp = $wUp -bor $xButtonData
				}

				if ($modifierNames -contains 'Alt')
				{
					[Custom.Ftool]::fnPostMessage($timerData['WindowHandle'], $WM_KEYDOWN, 0x12, 0)
				}

				[Custom.Ftool]::fnPostMessage($timerData['WindowHandle'], $msgDown, [IntPtr]$wDown, $lParam)
				Start-Sleep -Milliseconds $dwellTime
				[Custom.Ftool]::fnPostMessage($timerData['WindowHandle'], $msgUp, [IntPtr]$wUp, $lParam)

				if ($modifierNames -contains 'Alt')
				{
					[Custom.Ftool]::fnPostMessage($timerData['WindowHandle'], $WM_KEYUP, 0x12, 0xC0000000)
				}
			}
			else
			{
				$modifierVks = @()
				foreach ($modName in $modifierNames)
				{
					switch ($modName.ToUpper())
					{
						'CTRL' { $modifierVks += 0x11 }
						'ALT' { $modifierVks += 0x12 }
						'SHIFT' { $modifierVks += 0x10 }
					}
				}

				foreach ($modVk in $modifierVks)
				{
					[Custom.Ftool]::fnPostMessage($timerData['WindowHandle'], $WM_KEYDOWN, $modVk, 0)
				}

				[Custom.Ftool]::fnPostMessage($timerData['WindowHandle'], $WM_KEYDOWN, $virtualKeyCode, 0)
				Start-Sleep -Milliseconds $dwellTime
				[Custom.Ftool]::fnPostMessage($timerData['WindowHandle'], $WM_KEYUP, $virtualKeyCode, 0xC0000000)

				$reversedMods = $modifierVks | Sort-Object -Descending
				foreach ($modVk in $reversedMods)
				{
					[Custom.Ftool]::fnPostMessage($timerData['WindowHandle'], $WM_KEYUP, $modVk, 0xC0000000)
				}
			}
		}
	}
	catch { Write-Verbose ('FTOOL: Spammer Error: {0}' -f $_.Exception.Message) }
}

function LoadFtoolSettings
{
	param($formData)
    
	$profilePrefix = FindOrCreateProfile $formData.WindowTitle
	$formData.ProfilePrefix = $profilePrefix

	if ($profilePrefix)
	{
		$globalHotkeyName = "GlobalHotkey_$profilePrefix"
		if ($global:DashboardConfig.Config['Ftool'].Contains($globalHotkeyName))
		{
			$formData.GlobalHotkey = $global:DashboardConfig.Config['Ftool'][$globalHotkeyName]
		}
        
		$hotkeyName = "Hotkey_$profilePrefix"
		if ($global:DashboardConfig.Config['Ftool'].Contains($hotkeyName))
		{
			$formData.Hotkey = $global:DashboardConfig.Config['Ftool'][$hotkeyName]
			$formData.BtnHotKey.Text = $formData.Hotkey 
		}
		else
		{
			$formData.Hotkey = $null 
			$formData.BtnHotKey.Text = 'Hotkey' 
		}

		$hotkeysEnabledName = "hotkeys_enabled_$profilePrefix"
		if ($global:DashboardConfig.Config['Ftool'].Contains($hotkeysEnabledName))
		{
			$value = $global:DashboardConfig.Config['Ftool'][$hotkeysEnabledName]
			try { $formData.BtnHotkeyToggle.Checked = [bool]::Parse($value) } catch { $formData.BtnHotkeyToggle.Checked = $true }
		}
		else
		{
			$formData.BtnHotkeyToggle.Checked = $true
		}

		$keyName = "key1_$profilePrefix"
		$intervalName = "inpt1_$profilePrefix"
		$nameName = "name1_$profilePrefix"
		$positionXName = "pos1X_$profilePrefix"
		$positionYName = "pos1Y_$profilePrefix"
    
		if ($global:DashboardConfig.Config['Ftool'].Contains($keyName))
		{
			$formData.BtnKeySelect.Text = $global:DashboardConfig.Config['Ftool'][$keyName]
		}
		else
		{
			$formData.BtnKeySelect.Text = 'none'
		}
        
		if ($global:DashboardConfig.Config['Ftool'].Contains($intervalName))
		{
			$intervalValue = [int]$global:DashboardConfig.Config['Ftool'][$intervalName]
			if ($intervalValue -lt 10)
			{
				$intervalValue = 100
			}
			$formData.Interval.Text = $intervalValue.ToString()
		}
		else
		{
			$formData.Interval.Text = '100'
		}

		if ($global:DashboardConfig.Config['Ftool'].Contains($nameName))
		{
			$nameValue = [string]$global:DashboardConfig.Config['Ftool'][$nameName]
			$formData.Name.Text = $nameValue.ToString()
		}
		else
		{
			$formData.Name.Text = 'Main'
		}

		if ($global:DashboardConfig.Config['Ftool'].Contains($positionXName))
		{
			$positionValue = [int]$global:DashboardConfig.Config['Ftool'][$positionXName]
			$formData.PositionSliderX.Value = $positionValue
		}
		else
		{
			$formData.PositionSliderX.Value = 0
		}

		if ($global:DashboardConfig.Config['Ftool'].Contains($positionYName))
		{
			$positionValue = [int]$global:DashboardConfig.Config['Ftool'][$positionYName]
			$formData.PositionSliderY.Value = $positionValue
		}
		else
		{
			$formData.PositionSliderY.Value = 100
		}
	}
	else
	{
		$formData.Hotkey = $null
		$formData.BtnHotKey.Text = 'Hotkey' 
		$formData.BtnKeySelect.Text = 'none'
		$formData.Interval.Text = '1000'
		$formData.Name.Text = 'Main'
	}
}

function FindOrCreateProfile
{
	param($windowTitle)
    
	$profilePrefix = $null
	$profileFound = $false
    
	if ($windowTitle)
	{
		foreach ($key in $global:DashboardConfig.Config['Ftool'].Keys)
		{
			if ($key -like 'profile_*' -and $global:DashboardConfig.Config['Ftool'][$key] -eq $windowTitle)
			{
				$profilePrefix = $key -replace 'profile_', ''
				$profileFound = $true
				break
			}
		}
        
		if (-not $profileFound)
		{
			$maxProfileNum = 0
			foreach ($key in $global:DashboardConfig.Config['Ftool'].Keys)
			{
				if ($key -match 'profile_(\d+)')
				{
					$profileNum = [int]$matches[1]
					if ($profileNum -gt $maxProfileNum)
					{
						$maxProfileNum = $profileNum
					}
				}
			}
            
			$profilePrefix = ($maxProfileNum + 1).ToString()
			$profileKey = "profile_$profilePrefix"
			$global:DashboardConfig.Config['Ftool'][$profileKey] = $windowTitle
		}
	}
	return $profilePrefix
}

function InitializeExtensionTracking
{
	param($instanceId)
    
	$instanceKey = "instance_$instanceId"
    
	if (-not $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
	{
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey] = @{
			NextExtNum       = 2
			ActiveExtensions = @()
			RemovedExtNums   = @()
		}
	}
    
	$validActiveExtensions = @()
	foreach ($key in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions)
	{
		if ($global:DashboardConfig.Resources.ExtensionData.Contains($key))
		{
			$validActiveExtensions += $key
		}
	}
	$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions = $validActiveExtensions
    
	if ($global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
	{
		$validRemovedNums = @()
		foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
		{
			$intNum = 0
			if ([int]::TryParse($num.ToString(), [ref]$intNum))
			{
                
				if ($validRemovedNums -notcontains $intNum)
				{
					$validRemovedNums += $intNum
				}
			}
		}
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums = $validRemovedNums
	}
}

function GetNextExtensionNumber
{
	param($instanceId)
    
	$instanceKey = "instance_$instanceId"
    
	if ($global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums.Count -gt 0)
	{
		$sortedNums = @()
		foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
		{
			$intNum = 0
			if ([int]::TryParse($num.ToString(), [ref]$intNum))
			{
				$sortedNums += $intNum
			}
		}
        
		$sortedNums = $sortedNums | Sort-Object
        
		if ($sortedNums.Count -gt 0)
		{
			$extNum = $sortedNums[0]
            
			$newRemovedNums = @()
			foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
			{
				$intNum = 0
				if ([int]::TryParse($num.ToString(), [ref]$intNum) -and $intNum -ne $extNum)
				{
					$newRemovedNums += $intNum
				}
			}
			$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums = $newRemovedNums
            
			Write-Verbose "FTOOL: Reusing extension number $extNum for $instanceId"
		}
		else
		{
			$extNum = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum
			$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum++
			Write-Verbose "FTOOL: Using new extension number $extNum for $instanceId (no valid reusable numbers)"
		}
	}
	else
	{
		$extNum = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum++
		Write-Verbose "FTOOL: Using new extension number $extNum for $instanceId"
	}
    
	return $extNum
}

function FindExtensionKeyByControl
{
	param($control, $controlType)
    
	if (-not $control)
	{
		return $null 
	}
	if (-not $controlType -or -not ($controlType -is [string]))
	{
		return $null 
	}
    
	$controlId = $control.GetHashCode()
	$form = $control.FindForm()
    
	if ($form -and $form.Tag -and $form.Tag.ControlToExtensionMap.Contains($controlId))
	{
		$extKey = $form.Tag.ControlToExtensionMap[$controlId]
        
		if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
		{
			return $extKey
		}
		else
		{
			$form.Tag.ControlToExtensionMap.Remove($controlId)
		}
	}
    
	foreach ($key in $global:DashboardConfig.Resources.ExtensionData.Keys)
	{
		$extData = $global:DashboardConfig.Resources.ExtensionData[$key]
		if ($extData -and $extData.$controlType -eq $control)
		{
			if ($form -and $form.Tag)
			{
				$form.Tag.ControlToExtensionMap[$controlId] = $key
			}
			return $key
		}
	}
    
	return $null
}

function LoadExtensionSettings
{
	param($extData, $profilePrefix)
    
	$extNum = $extData.ExtNum
    
    
	$keyName = "key${extNum}_$profilePrefix"
	$intervalName = "inpt${extNum}_$profilePrefix"
	$nameName = "name${extNum}_$profilePrefix"
	$extHotkeyName = "ExtHotkey_${extNum}_$profilePrefix" 
    
	if ($global:DashboardConfig.Config['Ftool'].Contains($extHotkeyName))
	{
		$extData.Hotkey = $global:DashboardConfig.Config['Ftool'][$extHotkeyName]
		$extData.BtnHotKey.Text = $extData.Hotkey
	}
	else
	{
		$extData.Hotkey = $null
		$extData.BtnHotKey.Text = 'Hotkey' 
	}

	if ($global:DashboardConfig.Config['Ftool'].Contains($keyName))
	{
		$keyValue = $global:DashboardConfig.Config['Ftool'][$keyName]
		$extData.BtnKeySelect.Text = $keyValue
	}
	else
	{
		$extData.BtnKeySelect.Text = 'none'
	}
    
	if ($global:DashboardConfig.Config['Ftool'].Contains($intervalName))
	{
		$intervalValue = [int]$global:DashboardConfig.Config['Ftool'][$intervalName]
		if ($intervalValue -lt 10)
		{
			$intervalValue = 100
		}
		$extData.Interval.Text = $intervalValue.ToString()
	}
	else
	{
		$extData.Interval.Text = '1000'
	}

	if ($global:DashboardConfig.Config['Ftool'].Contains($nameName))
	{
		$nameValue = [string]$global:DashboardConfig.Config['Ftool'][$nameName]

		$extData.Name.Text = $nameValue.ToString()
	}
	else
	{
		$extData.Name.Text = 'Name'
	}
}

function UpdateSettings
{
	param($formData, $extData = $null, [switch]$forceWrite)
    
	$profilePrefix = FindOrCreateProfile $formData.WindowTitle
    
	if ($profilePrefix)
	{
		if ($extData)
		{
			$extNum = $extData.ExtNum
			$keyName = "key${extNum}_$profilePrefix"
			$intervalName = "inpt${extNum}_$profilePrefix"
			$nameName = "name${extNum}_$profilePrefix"
			$extHotkeyName = "ExtHotkey_${extNum}_$profilePrefix" 
            
			if ($extData.Hotkey)
			{ 
				$global:DashboardConfig.Config['Ftool'][$extHotkeyName] = $extData.Hotkey
			}
			else
			{
				if ($global:DashboardConfig.Config['Ftool'].Contains($extHotkeyName))
				{
					$global:DashboardConfig.Config['Ftool'].Remove($extHotkeyName)
				}
			}
            
			if ($extData.BtnKeySelect -and $extData.BtnKeySelect.Text)
			{
				$global:DashboardConfig.Config['Ftool'][$keyName] = $extData.BtnKeySelect.Text
			}
            
			if ($extData.Interval)
			{
				$global:DashboardConfig.Config['Ftool'][$intervalName] = $extData.Interval.Text
			}

			if ($extData.Name)
			{
				$global:DashboardConfig.Config['Ftool'][$nameName] = $extData.Name.Text
			}
		}
		else
		{
			$global:DashboardConfig.Config['Ftool']["GlobalHotkey_$profilePrefix"] = $formData.GlobalHotkey
			$global:DashboardConfig.Config['Ftool']["Hotkey_$profilePrefix"] = $formData.Hotkey
			$global:DashboardConfig.Config['Ftool']["hotkeys_enabled_$profilePrefix"] = $formData.BtnHotkeyToggle.Checked
			$global:DashboardConfig.Config['Ftool']["key1_$profilePrefix"] = $formData.BtnKeySelect.Text
			$global:DashboardConfig.Config['Ftool']["inpt1_$profilePrefix"] = $formData.Interval.Text
			$global:DashboardConfig.Config['Ftool']["name1_$profilePrefix"] = $formData.Name.Text
			$global:DashboardConfig.Config['Ftool']["pos1X_$profilePrefix"] = $formData.PositionSliderX.Value
			$global:DashboardConfig.Config['Ftool']["pos1Y_$profilePrefix"] = $formData.PositionSliderY.Value
		}
        
		if ($forceWrite)
		{
			WriteConfig
			if ($global:DashboardConfig.Resources.Timers.Contains('ConfigWriteTimer'))
			{
				$global:DashboardConfig.Resources.Timers['ConfigWriteTimer'].Stop()
			}
		}
		else
		{
			if (-not $global:DashboardConfig.Resources.Timers.Contains('ConfigWriteTimer'))
			{
				$ConfigWriteTimer = New-Object System.Windows.Forms.Timer
				$ConfigWriteTimer.Interval = 1000
				$ConfigWriteTimer.Add_Tick({
						param($timerSender, $timerArgs)
						try
						{
							if ($timerSender)
							{
								$timerSender.Stop()
							}
							WriteConfig
						}
						catch
						{
							Write-Verbose ('FTOOL: Error in config write timer: {0}' -f $_.Exception.Message)
						}
					})
				$global:DashboardConfig.Resources.Timers['ConfigWriteTimer'] = $ConfigWriteTimer
			}
			else
			{
				$ConfigWriteTimer = $global:DashboardConfig.Resources.Timers['ConfigWriteTimer']
			}
            
			$ConfigWriteTimer.Stop()
			$ConfigWriteTimer.Start()
		}
	}
}


function CreatePositionTimer
{
	param($formData)
    
	$positionTimer = New-Object System.Windows.Forms.Timer
	$positionTimer.Interval = 10
	$positionTimer.Tag = @{
		WindowHandle = $formData.SelectedWindow
		FtoolForm    = $formData.Form
		InstanceId   = $formData.InstanceId
		FormData     = $formData
		FtoolZState  = 'unknown' 
	}
        
	$positionTimer.Add_Tick({
			param($s, $e)
			try
			{
				if (-not $s -or -not $s.Tag) { return }
				$timerData = $s.Tag
				if (-not $timerData -or $timerData['WindowHandle'] -eq [IntPtr]::Zero) { return }
				if (-not $timerData['FtoolForm'] -or $timerData['FtoolForm'].IsDisposed) { return }
            
				$rect = New-Object Custom.Native+RECT
                    
				if ([Custom.Native]::IsWindow($timerData['WindowHandle']) -and [Custom.Native]::GetWindowRect($timerData['WindowHandle'], [ref]$rect))
				{
					try
					{
                                
							$sliderValueX = $timerData.FormData.PositionSliderX.Value
							$maxLeft = $rect.Right - $timerData['FtoolForm'].Width - 8
							$targetLeft = $rect.Left + 8 + (($maxLeft - ($rect.Left + 8)) * $sliderValueX / 100)
                                
							$sliderValueY = 100 - $timerData.FormData.PositionSliderY.Value
							$maxTop = $rect.Bottom - $timerData['FtoolForm'].Height - 8
							$targetTop = $rect.Top + 8 + (($maxTop - ($rect.Top + 8)) * $sliderValueY / 100)
    
							$currentLeft = $timerData['FtoolForm'].Left
							$newLeft = $currentLeft + ($targetLeft - $currentLeft) * 0.2 
							$timerData['FtoolForm'].Left = [int]$newLeft
    
							$currentTop = $timerData['FtoolForm'].Top
							$newTop = $currentTop + ($targetTop - $currentTop) * 0.2 
							$timerData['FtoolForm'].Top = [int]$newTop
                                                    
                                                        
							$ftoolHandle = $timerData['FtoolForm'].Handle
							$linkedHandle = $timerData['WindowHandle']
							$foregroundWindow = [Custom.Native]::GetForegroundWindow()
							$flags = [Custom.Native]::SWP_NOMOVE -bor [Custom.Native]::SWP_NOSIZE -bor [Custom.Native]::SWP_NOACTIVATE

							$shouldBeTopMost = $false
							if ($foregroundWindow -eq $linkedHandle) {
								$shouldBeTopMost = $true
							} else {
								$activeCtrl = [System.Windows.Forms.Control]::FromHandle($foregroundWindow)
								if ($activeCtrl -and -not $activeCtrl.IsDisposed -and $activeCtrl.Tag) {
									if ($activeCtrl.Tag.PSObject.Properties['SelectedWindow'] -and $activeCtrl.Tag.SelectedWindow -eq $linkedHandle) {
										$shouldBeTopMost = $true
									}
								}
							}

							if ($shouldBeTopMost)
							{
								$forceUpdate = $false
								if ($foregroundWindow -eq $linkedHandle) {
									$testPrev = [Custom.Native]::GetWindow($ftoolHandle, 3)
									$checkCount = 0
									while ($testPrev -ne [IntPtr]::Zero -and $checkCount -lt 50) {
										if ($testPrev -eq $linkedHandle) { $forceUpdate = $true; break }
										$testPrev = [Custom.Native]::GetWindow($testPrev, 3)
										$checkCount++
									}
								}

								if ($timerData.FtoolZState -ne 'topmost' -or $forceUpdate)
								{
									if ($forceUpdate) {
										[Custom.Native]::PositionWindow($ftoolHandle, [Custom.Native]::HWND_NOTOPMOST, 0, 0, 0, 0, $flags) | Out-Null
									}
									[Custom.Native]::PositionWindow($ftoolHandle, [Custom.Native]::HWND_TOPMOST, 0, 0, 0, 0, $flags) | Out-Null
									$timerData.FtoolZState = 'topmost'
								}
							}
							else
							{
								if ($timerData.FtoolZState -ne 'standard')
								{
									[Custom.Native]::PositionWindow($ftoolHandle, [Custom.Native]::HWND_NOTOPMOST, 0, 0, 0, 0, $flags) | Out-Null
									$timerData.FtoolZState = 'standard'
								}

								$next = [Custom.Native]::GetWindow($ftoolHandle, 2) 
								$amIAboveGame = $false
								$loopCount = 0
								
								while ($next -ne [IntPtr]::Zero -and $loopCount -lt 50)
								{
									if ($next -eq $linkedHandle) { $amIAboveGame = $true; break }
									$next = [Custom.Native]::GetWindow($next, 2)
									$loopCount++
								}

								if (-not $amIAboveGame)
								{
									$bottomOfCluster = [Custom.Native]::GetWindow($linkedHandle, 3)
									if ($bottomOfCluster -ne [IntPtr]::Zero) { [Custom.Native]::PositionWindow($ftoolHandle, $bottomOfCluster, 0, 0, 0, 0, $flags) | Out-Null }
									else { [Custom.Native]::PositionWindow($ftoolHandle, [Custom.Native]::TopWindowHandle, 0, 0, 0, 0, $flags) | Out-Null }
								}
							}
							$timerData.LastForegroundWindow = $foregroundWindow
					}
					catch
					{
						Write-Verbose ('FTOOL: Position timer error: {0} for {1}' -f $_.Exception.Message, $($timerData['InstanceId']))
					}
				}
				else
				{
                    
					Write-Verbose ('FTOOL: Parent window handle no longer valid. Closing ftool {0}.' -f $timerData['InstanceId'])
					$timerData['FtoolForm'].Close()
					$s.Stop()
					$s.Dispose()
					$global:DashboardConfig.Resources.Timers.Remove("ftoolPosition_$($timerData.InstanceId)")
				}
			}
			catch
			{
				Write-Verbose ('FTOOL: Position timer critical error: {0}' -f $_.Exception.Message)
			}
		})
    
	$positionTimer.Start()
	$global:DashboardConfig.Resources.Timers["ftoolPosition_$($formData.InstanceId)"] = $positionTimer
}

function RepositionExtensions
{
	param($form, $instanceId)
    
	if (-not $form -or -not $instanceId)
	{
		return 
	}
    
	$instanceKey = "instance_$instanceId"
    
	if (-not $global:DashboardConfig.Resources.ExtensionTracking -or 
		-not $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
	{
		return
	}
    
	$form.SuspendLayout()
    
	try
	{
		$baseHeight = 130
        
		$activeExtensions = @()
		if ($global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions)
		{
			$activeExtensions = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions
		}
        
		$newHeight = 130  
		$position = 0
        
		if ($activeExtensions.Count -gt 0)
		{
			$sortedExtensions = @()
			foreach ($extKey in $activeExtensions)
			{
				if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
				{
					$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
					$sortedExtensions += [PSCustomObject]@{
						Key    = $extKey
						ExtNum = [int]$extData.ExtNum 
					}
				}
			}
			$sortedExtensions = $sortedExtensions | Sort-Object ExtNum
            
			foreach ($extObj in $sortedExtensions)
			{
				$extKey = $extObj.Key
				if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
				{
					$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
					if ($extData -and $extData.Panel -and -not $extData.Panel.IsDisposed)
					{
						$newTop = $baseHeight + ($position * 70)
                        
						$extData.Panel.Top = $newTop
						$extData.Position = $position
                        
						$position++
					}
				}
			}
            
			if ($position -gt 0)
			{
				$newHeight = 130 + ($position * 70)
			}
		}
        
		if (-not $form.Tag.IsCollapsed)
		{
			$form.Height = $newHeight
			$form.Tag.OriginalHeight = $newHeight
			$form.Tag.PositionSliderY.Height = $newHeight - 20
		}
	}
	catch
	{
		Write-Verbose ('FTOOL: Error in RepositionExtensions: {0}' -f $_.Exception.Message)
	}
	finally
	{
		$form.ResumeLayout()
	}
}

function CreateSpammerTimer
{
	param($windowHandle, $keyValue, $instanceId, $interval, $extNum = $null, $extKey = $null)
    
	$spamTimer = New-Object System.Windows.Forms.Timer
	$spamTimer.Interval = $interval
    
	$timerTag = @{
		WindowHandle = $windowHandle
		Key          = $keyValue
		InstanceId   = $instanceId
	}
    
	if ($null -ne $extNum)
	{
		$timerTag['ExtNum'] = $extNum
		$timerTag['ExtKey'] = $extKey
	}
    
	$spamTimer.Tag = $timerTag
    
	$spamTimer.Add_Tick({
			param($s, $evt)
			try
			{
				if (-not $s -or -not $s.Tag) { return }
				Invoke-FtoolSpamAction -timerData $s.Tag
			}
			catch
			{
				Write-Verbose ('FTOOL: Spammer Error: {0}' -f $_.Exception.Message)
			}
		})
    
	Invoke-FtoolSpamAction -timerData $timerTag

	$spamTimer.Start()
	return $spamTimer
}

function ToggleButtonState
{
	param($startBtn, $stopBtn, $isStarting)
    
	if ($isStarting)
	{
		$startBtn.Enabled = $false
		$startBtn.Visible = $false
		$stopBtn.Enabled = $true
		$stopBtn.Visible = $true
	}
	else
	{
		$startBtn.Enabled = $true
		$startBtn.Visible = $true
		$stopBtn.Enabled = $false
		$stopBtn.Visible = $false
	}
}

function CheckRateLimit
{
	param($controlId, $minInterval = 10)
    
	$currentTime = [DateTime]::Now
	if ($global:DashboardConfig.Resources.LastEventTimes.Contains($controlId) -and 
		($currentTime - $global:DashboardConfig.Resources.LastEventTimes[$controlId]).TotalMilliseconds -lt $minInterval)
	{
		return $false
	}
    
	$global:DashboardConfig.Resources.LastEventTimes[$controlId] = $currentTime
	return $true
}

function AddFormCleanupHandler
{
	param($form)
    
	if (-not $form.Tag.ExtensionCleanupAdded)
	{
		$form.Add_FormClosing({
				param($src, $e)
            
				try
				{
					$instanceId = $src.Tag.InstanceId
					if ($instanceId)
					{
						CleanupInstanceResources $instanceId
					}
				}
				catch
				{
					Write-Verbose ('FTOOL: Error during form cleanup: {0}' -f $_.Exception.Message)
				}
			})
        
		$form.Tag | Add-Member -NotePropertyName ExtensionCleanupAdded -NotePropertyValue $true -Force
	}
}

function CleanupInstanceResources
{
	param($instanceId)
    
	$instanceKey = "instance_$instanceId"
    
	$keysToRemove = @()
    
	$activeExtensions = @()
	if ($global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey) -and 
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions)
	{
		$activeExtensions = @() + $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions
	}
    
	foreach ($key in $activeExtensions)
	{
		if ($global:DashboardConfig.Resources.ExtensionData.Contains($key))
		{
			$extData = $global:DashboardConfig.Resources.ExtensionData[$key]
			if ($extData.RunningSpammer)
			{
				$extData.RunningSpammer.Stop()
				$extData.RunningSpammer.Dispose()
			}
			if ($extData.HotkeyId)
			{
				try
				{
					UnregisterHotkeyInstance -Id $extData.HotkeyId -OwnerKey $key
				}
				catch
				{
					Write-Warning "FTOOL: Failed to unregister hotkey ID $($extData.HotkeyId) for extension $key during cleanup. Error: $_"
				}
			}
			$keysToRemove += $key
		}
	}
	foreach ($key in $keysToRemove)
	{
		$global:DashboardConfig.Resources.ExtensionData.Remove($key)
	}
    
	$timerKeysToRemove = @()
	foreach ($key in $global:DashboardConfig.Resources.Timers.Keys)
	{
		if ($key -like "ExtSpammer_${instanceId}_*" -or $key -eq "ftoolPosition_$instanceId")
		{
			$timer = $global:DashboardConfig.Resources.Timers[$key]
			if ($timer)
			{
				$timer.Stop()
				$timer.Dispose()
			}
			$timerKeysToRemove += $key
		}
	}
    
	foreach ($key in $timerKeysToRemove)
	{
		$global:DashboardConfig.Resources.Timers.Remove($key)
	}
    
	if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
	{
		$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
		if ($form -and $form.Tag)
		{
            
			if ($form.Tag.ToolTipFtool)
			{
				$form.Tag.ToolTipFtool.Dispose()
			}
			if ($form.Tag.HotkeyId)
			{
				try
				{
					UnregisterHotkeyInstance -Id $form.Tag.HotkeyId -OwnerKey $form.Tag.InstanceId
				}
				catch
				{
					Write-Warning "FTOOL: Failed to unregister hotkey with ID $($form.Tag.HotkeyId) for instance $instanceId. Error: $_"
				}
			}
			if ($form.Tag.GlobalHotkeyId)
			{
				try
				{
					$globalOwnerKey = "global_toggle_$($form.Tag.InstanceId)"
					UnregisterHotkeyInstance -Id $form.Tag.GlobalHotkeyId -OwnerKey $globalOwnerKey
				}
				catch
				{
					Write-Warning "FTOOL: Failed to unregister global hotkey with ID $($form.Tag.GlobalHotkeyId) for instance $instanceId. Error: $_"
				}
			}
		}
	}

	if ($global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
	{
		$global:DashboardConfig.Resources.ExtensionTracking.Remove($instanceKey)
	}
    
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
}

function StopFtoolForm
{
	param($Form)
    
	if (-not $Form -or $Form.IsDisposed)
	{
		return
	}
    
	try
	{
		$instanceId = $Form.Tag.InstanceId
		if ($instanceId)
		{
			CleanupInstanceResources $instanceId
            
			if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
			{
				$global:DashboardConfig.Resources.FtoolForms.Remove($instanceId)
			}
		}
        
		$Form.Close()
		$Form.Dispose()
	}
	catch
	{
		Write-Verbose ('FTOOL: Error stopping Ftool form: {0}' -f $_.Exception.Message)
	}
}

function RemoveExtension
{
	param($form, $extKey)
    
	if (-not $form -or -not $extKey)
	{
		return $false 
	}
    
	if (-not $global:DashboardConfig.Resources.ExtensionData -or 
		-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
	{
		return $false
	}
    
	$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
	if (-not $extData)
	{
		return $false 
	}
    
	$extNum = [int]$extData.ExtNum 
	$instanceId = $extData.InstanceId
	$instanceKey = "instance_$instanceId"
    
	try
	{
		if ($extData.RunningSpammer)
		{
			$extData.RunningSpammer.Stop()
			$extData.RunningSpammer.Dispose()
			$extData.RunningSpammer = $null
            
			$timerKey = "ExtSpammer_${instanceId}_$extNum"
			if ($global:DashboardConfig.Resources.Timers.Contains($timerKey))
			{
				$global:DashboardConfig.Resources.Timers.Remove($timerKey)
			}
		}
        
		if ($extData.Panel -and -not $extData.Panel.IsDisposed)
		{
			$form.Controls.Remove($extData.Panel)
			$extData.Panel.Dispose()
		}
        
		if (-not $global:DashboardConfig.Resources.ExtensionTracking -or 
			-not $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
		{
			InitializeExtensionTracking $instanceId
		}
        
		$intExtNum = [int]$extNum 
        
		$found = $false
		foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
		{
			$intNum = 0
			if ([int]::TryParse($num.ToString(), [ref]$intNum) -and $intNum -eq $intExtNum)
			{
				$found = $true
				break
			}
		}
        
		if (-not $found)
		{
			$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums += $intExtNum
		}
        
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions = 
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions | 
        Where-Object { $_ -ne $extKey }
        
		$global:DashboardConfig.Resources.ExtensionData.Remove($extKey)

		if ($extData.HotkeyId)
		{
			try
			{
				UnregisterHotkeyInstance -Id $extData.HotkeyId -OwnerKey $extKey
			}
			catch
			{
				Write-Warning "FTOOL: Failed to unregister hotkey ID $($extData.HotkeyId) for extension $($extKey). Error: $_"
			}
		}
        
		RepositionExtensions $form $instanceId
        
		return $true
	}
	catch
	{
		Write-Verbose ('FTOOL: Error in RemoveExtension: {0}' -f $_.Exception.Message)
		return $false
	}
}

#endregion

#region Core Functions

function FtoolSelectedRow
{
	param($row)
    
	if (-not $row -or -not $row.Cells -or $row.Cells.Count -lt 3)
	{
		Write-Verbose 'FTOOL: Invalid row data, skipping'
		return
	}
    
	$instanceId = $row.Cells[2].Value.ToString()
	if (-not $row.Tag -or -not $row.Tag.MainWindowHandle)
	{
		Write-Verbose "FTOOL: Missing window handle for instance $instanceId, skipping"
		return
	}
    
	$windowHandle = $row.Tag.MainWindowHandle
	if (-not $instanceId)
	{
		Write-Verbose 'FTOOL: Missing instance ID, skipping'
		return
	}
    
	if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
	{
		$existingForm = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
		if (-not $existingForm.IsDisposed)
		{
			$existingForm.BringToFront()
			return
		}
		else
		{
			$global:DashboardConfig.Resources.FtoolForms.Remove($instanceId)
		}
	}
    
	$targetWindowRect = New-Object Custom.Native+RECT
	[Custom.Native]::GetWindowRect($windowHandle, [ref]$targetWindowRect)
    
	$windowTitle = if ($row.Tag -and $row.Tag.MainWindowTitle)
	{
		$row.Tag.MainWindowTitle
	}
	elseif ($row.Cells[1].Value)
	{
		$row.Cells[1].Value.ToString()
	}
	else
	{
		"Window_$instanceId"
	}
    
	$ftoolForm = CreateFtoolForm $instanceId $targetWindowRect $windowTitle $row
    
	$global:DashboardConfig.Resources.FtoolForms[$instanceId] = $ftoolForm
    
	$ftoolForm.Show()
	$ftoolForm.BringToFront()
}

function CreateFtoolForm
{
	param($instanceId, $targetWindowRect, $windowTitle, $row)
    
	
	$ftoolForm = New-Object Custom.FtoolFormWindow
    
	
	$ftoolForm.Width = 235
	$ftoolForm.Height = 130
	$ftoolForm.Top = ($targetWindowRect.Top + 30)
	$ftoolForm.Left = ($targetWindowRect.Left + 10)
	$ftoolForm.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
	$ftoolForm.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
	$ftoolForm.Text = "FTool - $instanceId"
	$ftoolForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	$ftoolForm.Opacity = 1.0 

	if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon))
	{
		try
		{
			$icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon)
			$ftoolForm.Icon = $icon
		}
		catch
		{
			Write-Verbose "FTOOL: Failed to load icon from $($global:DashboardConfig.Paths.Icon): $_"
		}
	}
	if (-not $ftoolForm)
	{
		Write-Verbose "FTOOL: Failed to create form for $instanceId, skipping"
		return $null
	}



    
    
	$ftoolToolTip = New-Object System.Windows.Forms.ToolTip
	$ftoolToolTip.AutoPopDelay = 5000
	$ftoolToolTip.InitialDelay = 100
	$ftoolToolTip.ReshowDelay = 10
	$ftoolToolTip.ShowAlways = $true
	$ftoolToolTip.OwnerDraw = $true
	$ftoolToolTip | Add-Member -MemberType NoteProperty -Name 'TipFont' -Value (New-Object System.Drawing.Font('Segoe UI', 9))
	$ftoolToolTip.Add_Draw({
			$g, $b, $c = $_.Graphics, $_.Bounds, [System.Drawing.Color]
			$g.FillRectangle((New-Object System.Drawing.SolidBrush $c::FromArgb(30,30,30)), $b)
			$g.DrawRectangle((New-Object System.Drawing.Pen $c::FromArgb(100,100,100)), $b.X, $b.Y, $b.Width-1, $b.Height-1)
			$g.DrawString($_.ToolTipText, $this.TipFont, (New-Object System.Drawing.SolidBrush $c::FromArgb(240,240,240)), 3, 3, [System.Drawing.StringFormat]::GenericTypographic)
		})
	$ftoolToolTip.Add_Popup({
			$g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
			$s = $g.MeasureString($this.GetToolTip($_.AssociatedControl), $this.TipFont, [System.Drawing.PointF]::new(0,0), [System.Drawing.StringFormat]::GenericTypographic)
			$g.Dispose(); $_.ToolTipSize = [System.Drawing.Size]::new($s.Width+12, $s.Height+8)
		})
    
    
	$previousToolTip = $global:DashboardConfig.UI.ToolTipFtool
    
    
	$global:DashboardConfig.UI.ToolTipFtool = $ftoolToolTip
    

	try
	{
		$headerPanel = SetUIElement -type 'Panel' -visible $true -width 250 -height 20 -top 0 -left 0 -bg @(40, 40, 40)
		$ftoolForm.Controls.Add($headerPanel)

		$labelWinTitle = SetUIElement -type 'Label' -visible $true -width 120 -height 20 -top 5 -left 5 -bg @(40, 40, 40, 0) -fg @(255, 255, 255) -text $row.Cells[1].Value -font (New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Regular))
		$headerPanel.Controls.Add($labelWinTitle)

		$btnInstanceHotkeyToggle = SetUIElement -type 'Label' -visible $true -width 15 -height 15 -top 2 -left 135 -bg @(40, 40, 40) -fg @(255, 255, 255) -text ([char]0x2328) -font (New-Object System.Drawing.Font('Segoe UI', 10)) -tooltip "Set Master Hotkey`nAssign a global hotkey to toggle all hotkeys for this instance."
		$headerPanel.Controls.Add($btnInstanceHotkeyToggle)

		$btnHotkeyToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 15 -top 3 -left 150 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true -tooltip "Toggle Hotkeys`nEnable or disable all hotkeys for this specific Ftool instance."
		$headerPanel.Controls.Add($btnHotkeyToggle)
			
		$btnAdd = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 180 -bg @(40, 80, 80) -fg @(255, 255, 255) -text ([char]0x2795) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Add Extension`nAdd another Ftool extension slot for this client."
		$headerPanel.Controls.Add($btnAdd)
			
		$btnShowHide = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 195 -bg @(60, 60, 100) -fg @(255, 255, 255) -text ([char]0x25B2) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Minimize/Expand`nCollapse or expand the Ftool window."
		$headerPanel.Controls.Add($btnShowHide)
			
		$btnClose = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 215 -bg @(150, 20, 20) -fg @(255, 255, 255) -text ([char]0x166D) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Close`nStops all actions and closes this Ftool window."
		$headerPanel.Controls.Add($btnClose)
        
		$panelSettings = SetUIElement -type 'Panel' -visible $true -width 190 -height 60 -top 60 -left 40 -bg @(50, 50, 50)
		$ftoolForm.Controls.Add($panelSettings)
        
		$btnKeySelect = SetUIElement -type 'Button' -visible $true -width 55 -height 25 -top 4 -left 3 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Bind Key`nClick here, then press a key to assign it to this spammer."
		$panelSettings.Controls.Add($btnKeySelect)
        
		$interval = SetUIElement -type 'TextBox' -visible $true -width 47 -height 15 -top 5 -left 59 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '1000' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Interval (ms)`nTime in milliseconds between key presses."
		$panelSettings.Controls.Add($interval)
        
		$name = SetUIElement -type 'TextBox' -visible $true -width 37 -height 17 -top 5 -left 108 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'Main' -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular)) -tooltip "Name`nGive this spammer a name."
		$panelSettings.Controls.Add($name)
        
		$btnStart = SetUIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 10 -bg @(0, 120, 215) -fg @(255, 255, 255) -text 'Start' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Start`nBegin spamming the assigned key."
		$panelSettings.Controls.Add($btnStart)
        
		$btnStop = SetUIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 67 -bg @(200, 50, 50) -fg @(255, 255, 255) -text 'Stop' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Stop`nStop spamming the key."
		$btnStop.Enabled = $false
		$btnStop.Visible = $false
		$panelSettings.Controls.Add($btnStop)

		$btnHotKey = SetUIElement -type 'Button' -visible $true -width 40 -height 24 -top 4 -left 146 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'Hotkey' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 6)) -tooltip "Global Hotkey`nAssign a global hotkey to toggle this specific spammer on/off."
		$panelSettings.Controls.Add($btnHotKey)
        
		$positionSliderY = New-Object System.Windows.Forms.TrackBar
		$positionSliderY.Orientation = 'Vertical'
		$positionSliderY.Minimum = -18
		$positionSliderY.Maximum = 118
		$positionSliderY.TickFrequency = 300
		$positionSliderY.Value = 0
		$positionSliderY.Size = New-Object System.Drawing.Size(1, 110)
		$positionSliderY.Location = New-Object System.Drawing.Point(5, 20)
		$ftoolForm.Controls.Add($positionSliderY)
            
		$positionSliderX = New-Object System.Windows.Forms.TrackBar
		$positionSliderX.Minimum = -25
		$positionSliderX.Maximum = 125
		$positionSliderX.TickFrequency = 300
		$positionSliderX.Value = 0
		$positionSliderX.Size = New-Object System.Drawing.Size(190, 1)
		$positionSliderX.Location = New-Object System.Drawing.Point(45, 25)
		$ftoolForm.Controls.Add($positionSliderX)

	}
 	finally
	{
		$global:DashboardConfig.UI.ToolTipFtool = $previousToolTip
	}

	$formData = [PSCustomObject]@{
		InstanceId              = $instanceId
		SelectedWindow          = $row.Tag.MainWindowHandle
		BtnKeySelect            = $btnKeySelect
		Interval                = $interval
		Name                    = $name
		BtnStart                = $btnStart
		BtnStop                 = $btnStop
		BtnHotKey               = $btnHotKey
		BtnInstanceHotkeyToggle = $btnInstanceHotkeyToggle
		BtnHotkeyToggle         = $btnHotkeyToggle
		BtnAdd                  = $btnAdd
		BtnClose                = $btnClose
		BtnShowHide             = $btnShowHide
		PositionSliderX         = $positionSliderX
		PositionSliderY         = $positionSliderY
		Form                    = $ftoolForm
		ToolTipFtool            = $ftoolToolTip 
		RunningSpammer          = $null
		WindowTitle             = $windowTitle
		Process                 = $row.Tag
		OriginalLeft            = $targetWindowRect.Left
		IsCollapsed             = $false
		LastExtensionAdded      = 0
		ExtensionCount          = 0
		ControlToExtensionMap   = @{}
		OriginalHeight          = 130
		HotkeyId                = $null 
		GlobalHotkeyId          = $null 
		ProfilePrefix           = $null 
		Hotkey                  = $null 
		GlobalHotkey            = $null 
	}
    
	$ftoolForm.Tag = $formData
    
	LoadFtoolSettings $formData
	ToggleInstanceHotkeys -InstanceId $formData.InstanceId -ToggleState $formData.BtnHotkeyToggle.Checked
    
	if (-not [string]::IsNullOrEmpty($formData.Hotkey))
	{
		try
		{
			$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($formData.InstanceId)' -ExtKey `$null")
			$newHotkeyId = SetHotkey -KeyCombinationString $formData.Hotkey -Action $scriptBlock -OwnerKey $formData.InstanceId
			$formData.HotkeyId = $newHotkeyId
			$hotkeyIdDisplay = if ($formData.HotkeyId) { $formData.HotkeyId } else { 'None' }
			Write-Verbose "FTOOL: Registered hotkey $($formData.Hotkey) (ID: $hotkeyIdDisplay) for main instance $($formData.InstanceId) on load."
		}
		catch
		{
			Write-Warning "FTOOL: Failed to register hotkey $($formData.Hotkey) for main instance $($formData.InstanceId) on load. Error: $_"
			$formData.HotkeyId = $null 
		}
	}

	if (-not [string]::IsNullOrEmpty($formData.GlobalHotkey))
	{
		try
		{
			$ownerKey = "global_toggle_$($formData.InstanceId)"
			$script = @"
Write-Verbose "FTOOL: Global-toggle hotkey triggered for instance '$($formData.InstanceId)'"
if (`$global:DashboardConfig.Resources.FtoolForms.Contains('$($formData.InstanceId)')) {
    `$f = `$global:DashboardConfig.Resources.FtoolForms['$($formData.InstanceId)']
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
                ToggleInstanceHotkeys -InstanceId '$($formData.InstanceId)' -ToggleState `$toggle.Checked
                Write-Verbose "FTOOL: Global toggle action completed for instance '$($formData.InstanceId)'."
            } catch {
                Write-Warning "FTOOL: Error during global toggle action for instance '$($formData.InstanceId)'. Error: `$_"
            }
        }
    }
}
"@
			$scriptBlock = [scriptblock]::Create($script)
			$formData.GlobalHotkeyId = SetHotkey -KeyCombinationString $formData.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey
			Write-Verbose "FTOOL: Registered global-toggle hotkey $($formData.GlobalHotkey) (ID: $($formData.GlobalHotkeyId)) for instance $($formData.InstanceId) on load."
		}
		catch
		{
			Write-Warning "FTOOL: Failed to register global-toggle hotkey $($formData.GlobalHotkey) for instance $($formData.InstanceId) on load. Error: $_"
			$formData.GlobalHotkeyId = $null
		}
	}

	CreatePositionTimer $formData
	AddFtoolEventHandlers $formData
    
	return $ftoolForm
}

function AddFtoolEventHandlers
{
	param($formData)
    
	$formData.Interval.Add_TextChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			$intervalValue = 0
			if (-not [int]::TryParse($this.Text, [ref]$intervalValue) -or $intervalValue -lt 10)
			{
				$this.Text = '100'
			}
        
			UpdateSettings $data -forceWrite
		})

	$formData.Name.Add_TextChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
    
			UpdateSettings $data -forceWrite
		})

	$formData.PositionSliderX.Add_ValueChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return
			}
			$data = $form.Tag
			UpdateSettings $data -forceWrite
		})

	$formData.PositionSliderY.Add_ValueChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return
			}
			$data = $form.Tag
			UpdateSettings $data -forceWrite
		})
    
	$formData.BtnKeySelect.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			$currentKey = $data.BtnKeySelect.Text
			$newKey = Show-KeyCaptureDialog $currentKey -OwnerForm $form
        
			if ($newKey -and $newKey -ne $currentKey)
			{
				if ($data.RunningSpammer)
				{
					$data.RunningSpammer.Stop()
					$data.RunningSpammer.Dispose()
					$data.RunningSpammer = $null
				}
            
				$data.BtnKeySelect.Text = $newKey
				UpdateSettings $data -forceWrite
            
				ToggleButtonState $data.BtnStart $data.BtnStop $false
			}
		})
    
	$formData.BtnStart.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			$keyValue = $data.BtnKeySelect.Text
			if (-not $keyValue -or $keyValue.Trim() -eq '')
			{
				Show-DarkMessageBox 'Please select a key' 'Missing Ftool Key' 'Ok' 'Information'
				return
			}
        
			$intervalNum = 0
			if (-not [int]::TryParse($data.Interval.Text, [ref]$intervalNum) -or $intervalNum -lt 10)
			{
				$intervalNum = 100
				$data.Interval.Text = '100'
			}
        
			ToggleButtonState $data.BtnStart $data.BtnStop $true
        
			$spamTimer = CreateSpammerTimer $data.SelectedWindow $keyValue $data.InstanceId $intervalNum
        
			$data.RunningSpammer = $spamTimer
			$global:DashboardConfig.Resources.Timers["ExtSpammer_$($data.InstanceId)_1"] = $spamTimer
		})
    
	$formData.BtnStop.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			if ($data.RunningSpammer)
			{
				$data.RunningSpammer.Stop()
				$data.RunningSpammer.Dispose()
				$data.RunningSpammer = $null
			}
        
			ToggleButtonState $data.BtnStart $data.BtnStop $false
        
			if ($global:DashboardConfig.Resources.Timers.Contains("ExtSpammer_$($data.InstanceId)_1"))
			{
				$global:DashboardConfig.Resources.Timers.Remove("ExtSpammer_$($data.InstanceId)_1")
			}
		})

	$formData.BtnHotKey.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag

			$currentHotkeyText = $data.BtnHotKey.Text
			if ($currentHotkeyText -eq 'Hotkey') { $currentHotkeyText = $null } 

			$oldHotkeyIdToUnregister = $data.HotkeyId

			$newHotkey = Show-KeyCaptureDialog $currentHotkeyText -OwnerForm $form
        
			if ($newHotkey -and $newHotkey -ne $currentHotkeyText)
			{ 
        
				$data.Hotkey = $newHotkey       
				try
				{
					$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($data.InstanceId)' -ExtKey `$null")
					$data.HotkeyId = SetHotkey -KeyCombinationString $data.Hotkey -Action $scriptBlock -OwnerKey $data.InstanceId -OldHotkeyId $oldHotkeyIdToUnregister
					$data.BtnHotKey.Text = $newHotkey 
					Write-Verbose "FTOOL: Registered hotkey $($data.Hotkey) (ID: $($data.HotkeyId)) for main instance $($data.InstanceId)."
				}
				catch
				{
					Write-Warning "FTOOL: Failed to register hotkey $($data.Hotkey) for main instance $($data.InstanceId). Error: $_"
					$data.HotkeyId = $null 
					$data.Hotkey = $currentHotkeyText
					$data.BtnHotKey.Text = $currentHotkeyText -or 'Hotkey'
				}
            
				UpdateSettings $data -forceWrite 
			}
			elseif (-not $newHotkey -and $oldHotkeyIdToUnregister)
			{ 
				try
				{
					UnregisterHotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $data.InstanceId
					Write-Verbose "FTOOL: Unregistered hotkey (ID: $($oldHotkeyIdToUnregister)) for main instance $($data.InstanceId) due to user clear."
				}
				catch
				{
					Write-Warning "FTOOL: Failed to unregister hotkey (ID: $($oldHotkeyIdToUnregister)) for main instance $($data.InstanceId) on clear. Error: $_"
				}
				$data.HotkeyId = $null
				$data.Hotkey = $null
				$data.BtnHotKey.Text = 'Hotkey'
				UpdateSettings $data -forceWrite
			}
		})

	$formData.BtnHotkeyToggle.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag) { return }
			$data = $form.Tag
			$toggleOn = $this.Checked
			ToggleInstanceHotkeys -InstanceId $data.InstanceId -ToggleState $toggleOn
		})

	$formData.BtnInstanceHotkeyToggle.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag) { return }
			$data = $form.Tag
        
			$currentHotkeyText = $data.GlobalHotkey
        
			$oldHotkeyIdToUnregister = $data.GlobalHotkeyId
			$ownerKey = "global_toggle_$($data.InstanceId)"
        
			$newHotkey = Show-KeyCaptureDialog $currentHotkeyText -OwnerForm $form
        
			if ($newHotkey -and $newHotkey -ne $currentHotkeyText)
			{
				if (TestHotkeyConflict -KeyCombinationString $newHotkey -NewHotkeyType ([Custom.HotkeyManager+HotkeyActionType]::GlobalToggle) -OwnerKeyToExclude $ownerKey)
				{
					Show-DarkMessageBox "This hotkey combination ('$newHotkey') is already assigned to another function or instance. Please choose a different key." 'Hotkey Conflict' 'Ok' 'Information'
					$data.GlobalHotkey = $currentHotkeyText
					return 
				}
        
				$data.GlobalHotkey = $newHotkey
				try
				{
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
					$data.GlobalHotkeyId = SetHotkey -KeyCombinationString $data.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey -OldHotkeyId $oldHotkeyIdToUnregister
					Write-Verbose "FTOOL: Registered global-toggle hotkey $($data.GlobalHotkey) (ID: $($data.GlobalHotkeyId)) for instance $($data.InstanceId)."
				}
				catch
				{
					Write-Warning "FTOOL: Failed to register global-toggle hotkey $($data.GlobalHotkey) for instance $($data.InstanceId). Error: $_"
					$data.GlobalHotkeyId = $null
					$data.GlobalHotkey = $currentHotkeyText
				}
        
				UpdateSettings $data -forceWrite
			}
			elseif (-not $newHotkey -and $oldHotkeyIdToUnregister)
			{
				try
				{
					UnregisterHotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $ownerKey
					Write-Verbose "FTOOL: Unregistered global-toggle hotkey (ID: $($oldHotkeyIdToUnregister)) for instance $($data.InstanceId) due to user clear."
				}
				catch
				{
					Write-Warning "FTOOL: Failed to unregister global-toggle hotkey (ID: $($oldHotkeyIdToUnregister)) for instance $($data.InstanceId) on clear. Error: $_"
				}
				$data.GlobalHotkeyId = $null
				$data.GlobalHotkey = $null
				UpdateSettings $data -forceWrite
			}
		})

    
	$formData.BtnShowHide.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			if (-not (CheckRateLimit $this.GetHashCode() 10))
			{
				return 
			}
        
			if ($data.IsCollapsed)
			{
				$form.Height = $data.OriginalHeight
				$data.IsCollapsed = $false
				$this.Text = [char]0x25B2  
			}
			else
			{
				$data.OriginalHeight = $form.Height
				$form.Height = 25
				$data.IsCollapsed = $true
				$this.Text = [char]0x25BC  
			}
		})
    
	$formData.BtnAdd.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag -or $form.Tag.IsCollapsed)
			{
				return 
			}
			$data = $form.Tag
        
			if (-not (CheckRateLimit $this.GetHashCode() 10))
			{
				return 
			}
        
			if ($data.ExtensionCount -ge 8)
			{
				Show-DarkMessageBox "Maximum number of extensions reached.`nOnly 10 per client allowed!" 'Spam Protection' 'Ok' 'Information'
				return
			}
        
			$data.ExtensionCount++
        
			InitializeExtensionTracking $data.InstanceId
        
			$extNum = GetNextExtensionNumber $data.InstanceId
        
			$extData = CreateExtensionPanel $form $currentHeight $extNum $data.InstanceId $data.SelectedWindow
			$extKeyForScriptBlock = "ext_$($data.InstanceId)_$extNum"
        
			$profilePrefix = FindOrCreateProfile $data.WindowTitle
        
			LoadExtensionSettings $extData $profilePrefix
            
			if (-not [string]::IsNullOrEmpty($extData.Hotkey))
			{
				try
				{
					$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($extData.InstanceId)' -ExtKey '$($extKeyForScriptBlock)'")
					$newHotkeyId = SetHotkey -KeyCombinationString $extData.Hotkey -Action $scriptBlock -OwnerKey $extKeyForScriptBlock
					$extData.HotkeyId = $newHotkeyId
					$extHotKeyIdDisplay = if ($extData.HotkeyId) { $extData.HotkeyId } else { 'None' }
					Write-Verbose "FTOOL: Registered hotkey $($extData.Hotkey) (ID: $extHotKeyIdDisplay) for extension $($extKeyForScriptBlock) on load."
				}
				catch
				{
					Write-Warning "FTOOL: Failed to register hotkey $($extData.Hotkey) for extension $($extKeyForScriptBlock) on load. Error: $_"
					$extData.HotkeyId = $null 
				}
			}
        
			AddExtensionEventHandlers $extData $data
        
			AddFormCleanupHandler $form
        
			RepositionExtensions $form $extData.InstanceId
                
			$form.Refresh()
		})
    
	$formData.BtnClose.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			CleanupInstanceResources $data.InstanceId
        
			if ($global:DashboardConfig.Resources.FtoolForms.Contains($data.InstanceId))
			{
				$global:DashboardConfig.Resources.FtoolForms.Remove($data.InstanceId)
			}
        
			$form.Close()
			$form.Dispose()
		})
    
	$formData.Form.Add_FormClosed({
			param($src, $e)
        
			$instanceId = $src.Tag.InstanceId
			if ($instanceId)
			{
				CleanupInstanceResources $instanceId
			}
		})
}

function CreateExtensionPanel
{
	param($form, $currentHeight, $extNum, $instanceId, $windowHandle)
    
    
    
	$previousToolTip = $global:DashboardConfig.UI.ToolTipFtool
    
    
	if ($form.Tag -and $form.Tag.ToolTipFtool)
	{
		$global:DashboardConfig.UI.ToolTipFtool = $form.Tag.ToolTipFtool
	}
    

	try
	{

		$panelExt = SetUIElement -type 'Panel' -visible $true -width 190 -height 60 -top 0 -left 40 -bg @(50, 50, 50)
		$form.Controls.Add($panelExt)
		$panelExt.BringToFront()
    
		$btnKeySelectExt = SetUIElement -type 'Button' -visible $true -width 55 -height 25 -top 4 -left 3 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Bind Key`nClick here, then press a key to assign it to this spammer."
		$panelExt.Controls.Add($btnKeySelectExt)
    
		$intervalExt = SetUIElement -type 'TextBox' -visible $true -width 47 -height 15 -top 5 -left 59 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '1000' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Interval (ms)`nTime in milliseconds between key presses."
		$panelExt.Controls.Add($intervalExt)
    
		$btnStartExt = SetUIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 10 -bg @(0, 120, 215) -fg @(255, 255, 255) -text 'Start' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Start`nBegin spamming the assigned key."
		$panelExt.Controls.Add($btnStartExt)
    
		$btnStopExt = SetUIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 67 -bg @(200, 50, 50) -fg @(255, 255, 255) -text 'Stop' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Stop`nStop spamming the key."
		$btnStopExt.Enabled = $false
		$btnStopExt.Visible = $false
		$panelExt.Controls.Add($btnStopExt)

		$btnHotKeyExt = SetUIElement -type 'Button' -visible $true -width 40 -height 24 -top 4 -left 146 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'Hotkey' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 6)) -tooltip "Global Hotkey`nAssign a global hotkey to toggle this specific spammer on/off."
		$panelExt.Controls.Add($btnHotKeyExt)
    
		$nameExt = SetUIElement -type 'TextBox' -visible $true -width 37 -height 17 -top 5 -left 108 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'Name' -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular)) -tooltip "Name`nGive this spammer a name."
		$panelExt.Controls.Add($nameExt)
    
		$btnRemoveExt = SetUIElement -type 'Button' -visible $true -width 40 -height 20 -top 35 -left 120 -bg @(150, 50, 50) -fg @(255, 255, 255) -text 'Close' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 7)) -tooltip "Remove Extension`nDelete this Ftool extension slot."
		$panelExt.Controls.Add($btnRemoveExt)

	}
 finally
	{
        
		$global:DashboardConfig.UI.ToolTipFtool = $previousToolTip
	}

	$extKey = "ext_${instanceId}_$extNum"
	$instanceKey = "instance_$instanceId"
    
	$extData = [PSCustomObject]@{
		Panel          = $panelExt
		BtnKeySelect   = $btnKeySelectExt
		Interval       = $intervalExt
		BtnStart       = $btnStartExt
		BtnStop        = $btnStopExt
		BtnHotKey      = $btnHotKeyExt
		BtnRemove      = $btnRemoveExt
		Name           = $nameExt
		ExtNum         = $extNum
		InstanceId     = $instanceId
		WindowHandle   = $windowHandle
		RunningSpammer = $null
		Position       = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions.Count
		Hotkey         = $null 
		HotkeyId       = $null 
	}
    
	$global:DashboardConfig.Resources.ExtensionData[$extKey] = $extData
    
	if (-not $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions.Contains($extKey))
	{
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions += $extKey
	}
    
	return $extData
}

function AddExtensionEventHandlers
{
	param($extData, $formData)
    
	$extData.Interval.Add_TextChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			if (-not (CheckRateLimit $this.GetHashCode() 100))
			{
				return 
			}
        
			$extKey = FindExtensionKeyByControl $this 'Interval'
			if (-not $extKey)
			{
				return 
			}
        
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
        
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
        
			$intervalValue = 0
			if (-not [int]::TryParse($this.Text, [ref]$intervalValue) -or $intervalValue -lt 10)
			{
				$this.Text = '100'
				$intervalValue = 100
			}
        
			UpdateSettings $data $extData -forceWrite
		})

	$extData.Name.Add_TextChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
    
			if (-not (CheckRateLimit $this.GetHashCode() 100))
			{
				return 
			}
    
			$extKey = FindExtensionKeyByControl $this 'Name'
			if (-not $extKey)
			{
				return 
			}
    
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
    
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
    
			UpdateSettings $data $extData -forceWrite
		})
    
	$extData.BtnKeySelect.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			$extKey = FindExtensionKeyByControl $this 'BtnKeySelect'
			if (-not $extKey)
			{
				return 
			}
        
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
        
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
        
			$currentKey = $extData.BtnKeySelect.Text
			$newKey = Show-KeyCaptureDialog $currentKey -OwnerForm $form
        
			if ($newKey -and $newKey -ne $currentKey)
			{
				if ($extData.RunningSpammer)
				{
					$extData.RunningSpammer.Stop()
					$extData.RunningSpammer.Dispose()
					$extData.RunningSpammer = $null
				}
            
				$extData.BtnKeySelect.Text = $newKey
				UpdateSettings $data $extData -forceWrite
            
				ToggleButtonState $extData.BtnStart $extData.BtnStop $false
			}
		})
    
	$extData.BtnStart.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			if (-not (CheckRateLimit $this.GetHashCode() 1))
			{
				return 
			}
        
			$extKey = FindExtensionKeyByControl $this 'BtnStart'
			if (-not $extKey)
			{
				return 
			}
        
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
        
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
        
			$extNum = $extData.ExtNum
        
			$keyValue = $extData.BtnKeySelect.Text
			if (-not $keyValue -or $keyValue.Trim() -eq '')
			{
				Show-DarkMessageBox 'Please select a key' 'Missing Ftool Key' 'Ok' 'Information'
				return
			}
        
			$intervalNum = 0
        
			if (-not $extData.Interval -or -not $extData.Interval.Text)
			{
				$intervalNum = 100
			}
			else
			{
				if (-not [int]::TryParse($extData.Interval.Text, [ref]$intervalNum) -or $intervalNum -lt 10)
				{
					$intervalNum = 100
					if ($extData.Interval)
					{
						$extData.Interval.Text = '100'
					}
				}
			}
        
			ToggleButtonState $extData.BtnStart $extData.BtnStop $true
        
			$timerKey = "ExtSpammer_$($data.InstanceId)_$extNum"
			if ($global:DashboardConfig.Resources.Timers.Contains($timerKey))
			{
				$existingTimer = $global:DashboardConfig.Resources.Timers[$timerKey]
				if ($existingTimer)
				{
					$existingTimer.Stop()
					$existingTimer.Dispose()
					$global:DashboardConfig.Resources.Timers.Remove($timerKey)
				}
			}
        
			$spamTimer = CreateSpammerTimer $data.SelectedWindow $keyValue $data.InstanceId $intervalNum $extNum $extKey
        
			$extData.RunningSpammer = $spamTimer
			$global:DashboardConfig.Resources.Timers[$timerKey] = $spamTimer
		})
    
	$extData.BtnStop.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
        
			if (-not (CheckRateLimit $this.GetHashCode() 1))
			{
				return 
			}
        
			$extKey = FindExtensionKeyByControl $this 'BtnStop'
			if (-not $extKey)
			{
				return 
			}
        
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
        
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
        
			$extNum = $extData.ExtNum
        
			if ($extData.RunningSpammer)
			{
				$extData.RunningSpammer.Stop()
				$extData.RunningSpammer.Dispose()
				$extData.RunningSpammer = $null
			}
        
			ToggleButtonState $extData.BtnStart $extData.BtnStop $false
        
			$timerKey = "ExtSpammer_$($data.InstanceId)_$extNum"
			if ($global:DashboardConfig.Resources.Timers.Contains($timerKey))
			{
				$global:DashboardConfig.Resources.Timers.Remove($timerKey)
			}
		})

	$extData.BtnHotKey.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$formData = $form.Tag 

			$extKey = FindExtensionKeyByControl $this 'BtnHotKey'
			if (-not $extKey)
			{
				return 
			}
        
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
        
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
        
			$currentHotkeyText = $extData.BtnHotKey.Text
			if ($currentHotkeyText -eq 'Hotkey') { $currentHotkeyText = $null } 

			$oldHotkeyIdToUnregister = $extData.HotkeyId

			$newHotkey = Show-KeyCaptureDialog $currentHotkeyText -OwnerForm $form
        
			if ($newHotkey -and $newHotkey -ne $currentHotkeyText)
			{ 

				$extData.Hotkey = $newHotkey
            
				try
				{
					$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($extData.InstanceId)' -ExtKey '$($extKey)'")
					$extData.HotkeyId = SetHotkey -KeyCombinationString $extData.Hotkey -Action $scriptBlock -OwnerKey $extKey -OldHotkeyId $oldHotkeyIdToUnregister
					$extData.BtnHotKey.Text = $newHotkey 
					$extHotKeyIdDisplay = if ($extData.HotkeyId) { $extData.HotkeyId } else { 'None' }
					Write-Verbose "FTOOL: Registered hotkey $($extData.Hotkey) (ID: $extHotKeyIdDisplay) for extension $($extKey)."
				}
				catch
				{
					Write-Warning "FTOOL: Failed to register hotkey $($extData.Hotkey) for extension $($extKey). Error: $_"
					$extData.HotkeyId = $null 
					$extData.Hotkey = $currentHotkeyText
					$extData.BtnHotKey.Text = $currentHotkeyText -or 'Hotkey'
				}
            
				UpdateSettings $formData $extData -forceWrite
			}
			elseif (-not $newHotkey -and $oldHotkeyIdToUnregister)
			{ 
				try
				{
					UnregisterHotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $extKey
					Write-Verbose "FTOOL: Unregistered hotkey (ID: $($oldHotkeyIdToUnregister)) for extension $($extKey) due to user clear."
				}
				catch
				{
					Write-Warning "FTOOL: Failed to unregister hotkey (ID: $($oldHotkeyIdToUnregister)) for extension $($extKey) on clear. Error: $_"
				}
				$extData.HotkeyId = $null
				$extData.Hotkey = $null
				$extData.BtnHotKey.Text = 'Hotkey'
				UpdateSettings $formData $extData -forceWrite
			}
		})
    
	$extData.BtnRemove.Add_Click({
			try
			{
				$form = $this.FindForm()
				if (-not $form -or -not $form.Tag)
				{
					return 
				}
            
				if (-not (CheckRateLimit $this.GetHashCode() 100))
				{
					return 
				}
            
				$extKey = FindExtensionKeyByControl $this 'BtnRemove'
				if (-not $extKey)
				{
					return 
				}
            
				$this.Enabled = $false
            
				RemoveExtension $form $extKey
            
				$form.Tag.ExtensionCount--
			}
			catch
			{
				Write-Verbose ('FTOOL: Error in Remove button handler: {0}' -f $_.Exception.Message)
			}
		})
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion
#endregion