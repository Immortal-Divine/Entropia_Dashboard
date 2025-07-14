<# ftool.psm1 
	.SYNOPSIS
		Interface for Ftool.

	.DESCRIPTION
		This module creates and manages the complete ftool interface for Entropia Dashboard:
		- Builds the main application window and all dialog forms
		- Creates interactive controls (buttons, panels, grids, text boxes)
		- Handles window positioning
		- Implements settings management
		- Activates Ftool

	.NOTES
		Author: Immortal / Divine
		Version: 1.1
		Requires: PowerShell 5.1, .NET Framework 4.5+, classes.psm1, ini.psm1, datagrid.psm1, ftool.dll
#>

#region Helper Functions

<#
.SYNOPSIS
Gets virtual key code mappings for all supported keys
#>
function Get-VirtualKeyMappings
{
	return @{
		# Function Keys
		'F1' = 0x70; 'F2' = 0x71; 'F3' = 0x72; 'F4' = 0x73; 'F5' = 0x74; 'F6' = 0x75
		'F7' = 0x76; 'F8' = 0x77; 'F9' = 0x78; 'F10' = 0x79; 'F11' = 0x7A; 'F12' = 0x7B
		'F13' = 0x7C; 'F14' = 0x7D; 'F15' = 0x7E; 'F16' = 0x7F; 'F17' = 0x80; 'F18' = 0x81
		'F19' = 0x82; 'F20' = 0x83; 'F21' = 0x84; 'F22' = 0x85; 'F23' = 0x86; 'F24' = 0x87

		# Number Keys (Top Row)
		'0' = 0x30; '1' = 0x31; '2' = 0x32; '3' = 0x33; '4' = 0x34
		'5' = 0x35; '6' = 0x36; '7' = 0x37; '8' = 0x38; '9' = 0x39

		# Letter Keys
		'A' = 0x41; 'B' = 0x42; 'C' = 0x43; 'D' = 0x44; 'E' = 0x45; 'F' = 0x46; 'G' = 0x47; 'H' = 0x48
		'I' = 0x49; 'J' = 0x4A; 'K' = 0x4B; 'L' = 0x4C; 'M' = 0x4D; 'N' = 0x4E; 'O' = 0x4F; 'P' = 0x50
		'Q' = 0x51; 'R' = 0x52; 'S' = 0x53; 'T' = 0x54; 'U' = 0x55; 'V' = 0x56; 'W' = 0x57; 'X' = 0x58
		'Y' = 0x59; 'Z' = 0x5A

		# Main Special Keys
		'SPACE' = 0x20; 'ENTER' = 0x0D; 'TAB' = 0x09; 'ESCAPE' = 0x1B; 'SHIFT' = 0x10; 'CONTROL' = 0x11; 'ALT' = 0x12
		'UP_ARROW' = 0x26; 'DOWN_ARROW' = 0x28; 'LEFT_ARROW' = 0x25; 'RIGHT_ARROW' = 0x27; 'HOME' = 0x24; 'END' = 0x23
		'PAGE_UP' = 0x21; 'PAGE_DOWN' = 0x22; 'INSERT' = 0x2D; 'DELETE' = 0x2E; 'BACKSPACE' = 0x08

		# Lock & Modifier Keys
		'CAPS_LOCK' = 0x14; 'NUM_LOCK' = 0x90; 'SCROLL_LOCK' = 0x91; 'PRINT_SCREEN' = 0x2C; 'PAUSE_BREAK' = 0x13
		'LEFT_WINDOWS' = 0x5B; 'RIGHT_WINDOWS' = 0x5C; 'APPLICATION' = 0x5D; 'LEFT_SHIFT' = 0xA0; 'RIGHT_SHIFT' = 0xA1
		'LEFT_CONTROL' = 0xA2; 'RIGHT_CONTROL' = 0xA3; 'LEFT_ALT' = 0xA4; 'RIGHT_ALT' = 0xA5; 'SLEEP' = 0x5F

		# Numpad Keys
		'NUMPAD_0' = 0x60; 'NUMPAD_1' = 0x61; 'NUMPAD_2' = 0x62; 'NUMPAD_3' = 0x63; 'NUMPAD_4' = 0x64
		'NUMPAD_5' = 0x65; 'NUMPAD_6' = 0x66; 'NUMPAD_7' = 0x67; 'NUMPAD_8' = 0x68; 'NUMPAD_9' = 0x69
		'NUMPAD_MULTIPLY' = 0x6A; 'NUMPAD_ADD' = 0x6B; 'NUMPAD_SEPARATOR' = 0x6C; 'NUMPAD_SUBTRACT' = 0x6D; 'NUMPAD_DECIMAL' = 0x6E; 'NUMPAD_DIVIDE' = 0x6F

		# Punctuation Keys (US Layout)
		'SEMICOLON' = 0xBA; 'EQUALS' = 0xBB; 'COMMA' = 0xBC; 'MINUS' = 0xBD; 'PERIOD' = 0xBE
		'FORWARD_SLASH' = 0xBF; 'BACKTICK' = 0xC0; 'LEFT_BRACKET' = 0xDB; 'BACKSLASH' = 0xDC; 'RIGHT_BRACKET' = 0xDD
		'APOSTROPHE' = 0xDE

		# Media and Browser Keys
		'BROWSER_BACK' = 0xA6; 'BROWSER_FORWARD' = 0xA7; 'BROWSER_REFRESH' = 0xA8; 'BROWSER_STOP' = 0xA9
		'BROWSER_SEARCH' = 0xAA; 'BROWSER_FAVORITES' = 0xAB; 'BROWSER_HOME' = 0xAC; 'VOLUME_MUTE' = 0xAD
		'VOLUME_DOWN' = 0xAE; 'VOLUME_UP' = 0xAF; 'MEDIA_NEXT_TRACK' = 0xB0; 'MEDIA_PREVIOUS_TRACK' = 0xB1
		'MEDIA_STOP' = 0xB2; 'MEDIA_PLAY_PAUSE' = 0xB3; 'LAUNCH_MAIL' = 0xB4; 'LAUNCH_MEDIA_PLAYER' = 0xB5
		'LAUNCH_MY_COMPUTER' = 0xB6; 'LAUNCH_CALCULATOR' = 0xB7

		# International & Other Keys
		'IME_KANA_HANGUL' = 0x15; 'IME_JUNJA' = 0x17; 'IME_FINAL' = 0x18; 'IME_HANJA_KANJI' = 0x19
		'IME_CONVERT' = 0x1C; 'IME_NONCONVERT' = 0x1D; 'IME_ACCEPT' = 0x1E; 'IME_MODE_CHANGE' = 0x1F; 'IME_PROCESS' = 0xE5
		'SELECT' = 0x29; 'PRINT' = 0x2A; 'EXECUTE' = 0x2B; 'HELP' = 0x2F; 'CLEAR' = 0x0C
		'ATTN' = 0xF6; 'CRSEL' = 0xF7; 'EXSEL' = 0xF8; 'ERASE_EOF' = 0xF9; 'PLAY' = 0xFA; 'ZOOM' = 0xFB
		'PA1' = 0xFD; 'OEM_CLEAR' = 0xFE
	}
}

<#
.SYNOPSIS
Creates key capture dialog
#>
function Show-KeyCaptureDialog
{
	param($currentKey = 'F1')
	
	# Create key capture form
	$captureForm = New-Object System.Windows.Forms.Form
	$captureForm.Text = "Press a Key"
	$captureForm.Size = New-Object System.Drawing.Size(300, 150)
	$captureForm.StartPosition = 'CenterParent'
	$captureForm.FormBorderStyle = 'FixedDialog'
	$captureForm.MaximizeBox = $false
	$captureForm.MinimizeBox = $false
	$captureForm.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
	$captureForm.ForeColor = [System.Drawing.Color]::White
	
	# Create instruction label
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Press any key to capture it.`nCurrent: $currentKey`n`nPress ESC to cancel."
	$label.Size = New-Object System.Drawing.Size(280, 60)
	$label.Location = New-Object System.Drawing.Point(10, 10)
	$label.TextAlign = 'MiddleCenter'
	$label.Font = New-Object System.Drawing.Font('Segoe UI', 10)
	$captureForm.Controls.Add($label)
	
	# Create result label
	$resultLabel = New-Object System.Windows.Forms.Label
	$resultLabel.Text = "Waiting for key press..."
	$resultLabel.Size = New-Object System.Drawing.Size(280, 25)
	$resultLabel.Location = New-Object System.Drawing.Point(10, 75)
	$resultLabel.TextAlign = 'MiddleCenter'
	$resultLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
	$resultLabel.ForeColor = [System.Drawing.Color]::Yellow
	$captureForm.Controls.Add($resultLabel)
	
	# Variable to store captured key
	$script:capturedKey = $null
	
	# Add KeyDown event handler
	$captureForm.Add_KeyDown({
		param($form, $e)
		
		# Handle ESC to cancel
		if ($e.KeyCode -eq 'Escape')
		{
			$script:capturedKey = $null
			$captureForm.DialogResult = 'Cancel'
			$captureForm.Close()
			return
		}
		
		# Get key mappings
		$keyMappings = Get-VirtualKeyMappings
		$keyName = $null
		
		# Find the key name from the virtual key code
		foreach ($kvp in $keyMappings.GetEnumerator())
		{
			if ($kvp.Value -eq $e.KeyValue)
			{
				$keyName = $kvp.Key
				break
			}
		}
		
		# Handle special cases for keys not in our mapping
		if (-not $keyName)
		{
			$keyName = $e.KeyCode.ToString().ToUpper()
			# Verify the key name exists in our mappings, if not use a fallback
			$keyMappingsRecheck = Get-VirtualKeyMappings
			if (-not $keyMappingsRecheck.Contains($keyName))
			{
				Write-Verbose "FTOOL: Key '$keyName' not found in virtual key mappings, using fallback" -ForegroundColor Yellow
				$keyName = 'F1' # Fallback to F1 if key is not supported
			}
		}
		
		$script:capturedKey = $keyName
		$resultLabel.Text = "Captured: $keyName"
		$resultLabel.ForeColor = [System.Drawing.Color]::Green
		
		# Auto-close after short delay
		$timer = New-Object System.Windows.Forms.Timer
		$timer.Interval = 100
		$timer.Add_Tick({
			param($timerSender, $timerArgs)
			try
			{
				if ($timerSender)
				{
					$timerSender.Stop()
				}
				if ($captureForm -and -not $captureForm.IsDisposed)
				{
					$captureForm.DialogResult = 'OK'
					$captureForm.Close()
				}
			}
			catch
			{
				# Silent error handling for timer cleanup
			}
		})
		$timer.Start()
	})
	
	# Set form to capture all key events
	$captureForm.KeyPreview = $true
	$captureForm.TopMost = $true
	
	# Show dialog and return result
	$result = $captureForm.ShowDialog()
	
	if ($result -eq 'OK' -and $script:capturedKey)
	{
		return $script:capturedKey
	}
	else
	{
		return $currentKey  # Return current key if cancelled
	}
}

<#
.SYNOPSIS
Loads settings for the FTool form
#>
function LoadFtoolSettings
{
	param($formData)
	
	# Initialize settings if needed
	if (-not $global:DashboardConfig.Config)
	{ 
		$global:DashboardConfig.Config = [ordered]@{} 
	}
	if (-not $global:DashboardConfig.Config.Contains('Ftool'))
	{ 
		$global:DashboardConfig.Config['Ftool'] = [ordered]@{} 
	}
	
	# Look for profile-specific settings
	$profilePrefix = FindOrCreateProfile $formData.WindowTitle
	
	# Set Data from profile or default
	if ($profilePrefix)
	{
		$keyName = "key1_$profilePrefix"
		$intervalName = "inpt1_$profilePrefix"
		$nameName = "name1_$profilePrefix"

		# Set Key from profile or default		
		if ($global:DashboardConfig.Config['Ftool'].Contains($keyName))
		{
			$formData.BtnKeySelect.Text = $global:DashboardConfig.Config['Ftool'][$keyName]
		}
		else
		{
			$formData.BtnKeySelect.Text = 'F1'
		}
		
		# Set interval from profile or default
		if ($global:DashboardConfig.Config['Ftool'].Contains($intervalName))
		{
			$intervalValue = [int]$global:DashboardConfig.Config['Ftool'][$intervalName]
			if ($intervalValue -lt 100)
			{
				$intervalValue = 1000
			}
			$formData.Interval.Text = $intervalValue.ToString()
		}
		else
		{
			$formData.Interval.Text = '1000'
		}

		# Set interval from profile or default
		if ($global:DashboardConfig.Config['Ftool'].Contains($nameName))
		{
			$nameValue = [string]$global:DashboardConfig.Config['Ftool'][$nameName]
			$formData.Name.Text = $nameValue.ToString()
		}
		else
		{
			$formData.Name.Text = 'Main'
		}
	}
	else
	{
		$formData.BtnKeySelect.Text = 'F1'
		$formData.Interval.Text = '1000'
		$formData.Name.Text = 'Main'
	}
}

<#
.SYNOPSIS
Finds existing profile or creates a new one
#>
function FindOrCreateProfile
{
	param($windowTitle)
	
	$profilePrefix = $null
	$profileFound = $false
	
	if ($windowTitle)
	{
		# Check if we have a profile for this window title
		foreach ($key in $global:DashboardConfig.Config['Ftool'].Keys)
		{
			if ($key -like 'profile_*' -and $global:DashboardConfig.Config['Ftool'][$key] -eq $windowTitle)
			{
				$profilePrefix = $key -replace 'profile_', ''
				$profileFound = $true
				break
			}
		}
		
		# Create new profile if not found
		if (-not $profileFound)
		{
			# Find next available profile number
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

<#
.SYNOPSIS
Initializes tracking for extensions
#>
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
	
	# Clean up invalid extensions
	$validActiveExtensions = @()
	foreach ($key in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions)
	{
		if ($global:DashboardConfig.Resources.ExtensionData.Contains($key))
		{
			$validActiveExtensions += $key
		}
	}
	$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions = $validActiveExtensions
	
	# Ensure RemovedExtNums contains only valid integers
	if ($global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
	{
		$validRemovedNums = @()
		foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
		{
			$intNum = 0
			if ([int]::TryParse($num.ToString(), [ref]$intNum))
			{
				# Only add if not already in the list
				if ($validRemovedNums -notcontains $intNum)
				{
					$validRemovedNums += $intNum
				}
			}
		}
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums = $validRemovedNums
	}
}

<#
.SYNOPSIS
Gets next available extension number
#>
function GetNextExtensionNumber
{
	param($instanceId)
	
	$instanceKey = "instance_$instanceId"
	
	# First check if we have any removed extension numbers we can reuse
	if ($global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums.Count -gt 0)
	{
		# Get the smallest removed extension number
		# Convert to integers before sorting to ensure proper numeric sorting
		$sortedNums = @()
		foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
		{
			$intNum = 0
			if ([int]::TryParse($num.ToString(), [ref]$intNum))
			{
				$sortedNums += $intNum
			}
		}
		
		# Sort the numbers numerically
		$sortedNums = $sortedNums | Sort-Object
		
		if ($sortedNums.Count -gt 0)
		{
			$extNum = $sortedNums[0]
			
			# Remove it from the list of removed numbers
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
			
			Write-Verbose "FTOOL: Reusing extension number $extNum for $instanceId" -ForegroundColor DarkGray
		}
		else
		{
			# Fallback if no valid numbers were found
			$extNum = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum
			$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum++
			Write-Verbose "FTOOL: Using new extension number $extNum for $instanceId (no valid reusable numbers)" -ForegroundColor Yellow
		}
	}
	else
	{
		# Use the next available extension number
		$extNum = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum++
		Write-Verbose "FTOOL: Using new extension number $extNum for $instanceId" -ForegroundColor DarkGray
	}
	
	return $extNum
}

<#
.SYNOPSIS
Finds extension key by control reference
#>
function FindExtensionKeyByControl
{
	param($control, $controlType)
	
	# Validate input parameters
	if (-not $control)
	{
		return $null 
	}
	if (-not $controlType -or -not ($controlType -is [string]))
	{
		return $null 
	}
	
	# Use efficient lookup method with control hash code
	$controlId = $control.GetHashCode()
	$form = $control.FindForm()
	
	# Check if we have a cached mapping
	if ($form -and $form.Tag -and $form.Tag.ControlToExtensionMap.Contains($controlId))
	{
		$extKey = $form.Tag.ControlToExtensionMap[$controlId]
		
		# Verify that the extension still exists
		if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
		{
			return $extKey
		}
		else
		{
			# Remove stale mapping
			$form.Tag.ControlToExtensionMap.Remove($controlId)
		}
	}
	
	# If we don"t have a cached mapping, search for the extension
	foreach ($key in $global:DashboardConfig.Resources.ExtensionData.Keys)
	{
		$extData = $global:DashboardConfig.Resources.ExtensionData[$key]
		if ($extData -and $extData.$controlType -eq $control)
		{
			# Store the mapping for future lookups
			if ($form -and $form.Tag)
			{
				$form.Tag.ControlToExtensionMap[$controlId] = $key
			}
			return $key
		}
	}
	
	return $null
}

<#
.SYNOPSIS
Loads settings for extension
#>
function LoadExtensionSettings
{
	param($extData, $profilePrefix)
	
	$extNum = $extData.ExtNum
	
	# Load settings from config if they exist
	$keyName = "key${extNum}_$profilePrefix"
	$intervalName = "inpt${extNum}_$profilePrefix"
	$nameName = "name${extNum}_$profilePrefix"
	
	if ($global:DashboardConfig.Config['Ftool'].Contains($keyName))
	{
		$keyValue = $global:DashboardConfig.Config['Ftool'][$keyName]
		$extData.BtnKeySelect.Text = $keyValue
	}
	else
	{
		$extData.BtnKeySelect.Text = 'F1'
	}
	
	if ($global:DashboardConfig.Config['Ftool'].Contains($intervalName))
	{
		$intervalValue = [int]$global:DashboardConfig.Config['Ftool'][$intervalName]
		if ($intervalValue -lt 100)
		{
			$intervalValue = 1000
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

<#
.SYNOPSIS
Creates or updates settings in config
#>
function UpdateSettings
{
	param($formData, $extData = $null)
	
	$profilePrefix = FindOrCreateProfile $formData.WindowTitle
	
	if ($profilePrefix)
	{
		if ($extData)
		{
			# Update extension settings
			$extNum = $extData.ExtNum
			$keyName = "key${extNum}_$profilePrefix"
			$intervalName = "inpt${extNum}_$profilePrefix"
			$nameName = "name${extNum}_$profilePrefix"
			
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
			# Update main form settings
			$global:DashboardConfig.Config['Ftool']["key1_$profilePrefix"] = $formData.BtnKeySelect.Text
			$global:DashboardConfig.Config['Ftool']["inpt1_$profilePrefix"] = $formData.Interval.Text
			$global:DashboardConfig.Config['Ftool']["name1_$profilePrefix"] = $formData.Name.Text
		}
		
		# Batch write operations to reduce disk I/O
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
						Write-Config
					}
					catch
					{
						Write-Verbose "FTOOL: Error in config write timer: $($_.Exception.Message)" -ForegroundColor Red
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

<#
.SYNOPSIS
Creates timer to update form position
#>
function CreatePositionTimer
{
	param($formData)
	
	$positionTimer = New-Object System.Windows.Forms.Timer
	# Use 10ms interval for better responsiveness
	$positionTimer.Interval = 10
	$positionTimer.Tag = @{
		WindowHandle = $formData.SelectedWindow
		FtoolForm    = $formData.Form
		InstanceId   = $formData.InstanceId
		FormData     = $formData
	}
	
	$positionTimer.Add_Tick({
			param($s, $e)
			try
			{
				if (-not $s -or -not $s.Tag)
				{
					return
				}
				$timerData = $s.Tag
				if (-not $timerData -or $timerData['WindowHandle'] -eq [IntPtr]::Zero)
				{
					return
				}
			
				# Verify that form exists and is not disposed
				if (-not $timerData['FtoolForm'] -or $timerData['FtoolForm'].IsDisposed)
				{
					return
				}
		
				# Get current window position
				$rect = New-Object Native+RECT
				if ([Native]::GetWindowRect($timerData['WindowHandle'], [ref]$rect))
				{
					try
					{
						# Update main form position
						$timerData['FtoolForm'].Top = $rect.Top + 30
						$timerData['FtoolForm'].Left = $rect.Left + 8
					
						# Get foreground window to check if target window is active
						$foregroundWindow = [Native]::GetForegroundWindow()
					
						# Manage TopMost state based on which window has focus
						if ($foregroundWindow -eq $timerData['WindowHandle'])
						{
							# If target window has focus, make ftool topmost
							if (-not $timerData['FtoolForm'].TopMost)
							{
								$timerData['FtoolForm'].TopMost = $true
								$timerData['FtoolForm'].BringToFront()
							}
						}
						elseif ($foregroundWindow -ne $timerData['FtoolForm'].Handle)
						{
							# If neither target window nor ftool has focus, remove topmost
							if ($timerData['FtoolForm'].TopMost)
							{
								$timerData['FtoolForm'].TopMost = $false
							}
						}
					}
					catch
					{
						Write-Verbose "FTOOL: Position timer error: $($_.Exception.Message) for $($timerData['InstanceId'])" -ForegroundColor Red
					}
				}
			}
			catch
			{
				Write-Verbose "FTOOL: Position timer critical error: $($_.Exception.Message)" -ForegroundColor Red
			}
		})
	
	$positionTimer.Start()
	$global:DashboardConfig.Resources.Timers["ftoolPosition_$($formData.InstanceId)"] = $positionTimer
}

<#
.SYNOPSIS
Repositions extensions to fix layout issues
#>
function RepositionExtensions
{
	param($form, $instanceId)
	
	# Validate input parameters
	if (-not $form -or -not $instanceId)
	{
		return 
	}
	
	$instanceKey = "instance_$instanceId"
	
	# Check if extension tracking exists
	if (-not $global:DashboardConfig.Resources.ExtensionTracking -or 
		-not $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
	{
		return
	}
	
	# Suspend layout to prevent flickering during repositioning
	$form.SuspendLayout()
	
	try
	{
		# Get base height of form (without extensions)
		$baseHeight = 90
		
		# Get active extensions
		$activeExtensions = @()
		if ($global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions)
		{
			$activeExtensions = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions
		}
		
		# Calculate new form height based on number of extensions
		$newHeight = 90  # Base height with no extensions
		$position = 0
		
		if ($activeExtensions.Count -gt 0)
		{
			# Sort extensions by their extension number to maintain proper order
			$sortedExtensions = @()
			foreach ($extKey in $activeExtensions)
			{
				if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
				{
					$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
					$sortedExtensions += [PSCustomObject]@{
						Key    = $extKey
						ExtNum = [int]$extData.ExtNum  # Ensure ExtNum is treated as integer
					}
				}
			}
			$sortedExtensions = $sortedExtensions | Sort-Object ExtNum
			
			# Reposition each extension panel
			foreach ($extObj in $sortedExtensions)
			{
				$extKey = $extObj.Key
				if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
				{
					$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
					if ($extData -and $extData.Panel -and -not $extData.Panel.IsDisposed)
					{
						# Calculate new top position with proper spacing
						$newTop = $baseHeight + ($position * 70)
						
						# Update panel position
						$extData.Panel.Top = $newTop
						$extData.Position = $position
						
						$position++
					}
				}
			}
			
			# Calculate height based on actual number of extensions
			if ($position -gt 0)
			{
				$newHeight = 90 + ($position * 70)
			}
		}
		
		# Resize form to fit extensions if not collapsed
		if (-not $form.Tag.IsCollapsed)
		{
			$form.Height = $newHeight
			$form.Tag.OriginalHeight = $newHeight
		}
	}
	catch
	{
		Write-Verbose "FTOOL: Error in RepositionExtensions: $($_.Exception.Message)" -ForegroundColor Red
	}
	finally
	{
		# Resume layout to apply all changes at once
		$form.ResumeLayout()
	}
}

<#
.SYNOPSIS
Creates spammer timer for F-key
#>
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
				if (-not $s -or -not $s.Tag)
				{
					return
				}
				$timerData = $s.Tag
				if (-not $timerData -or $timerData['WindowHandle'] -eq [IntPtr]::Zero)
				{
					return
				}
				
				# Get virtual key code for the selected key
				$keyMappings = Get-VirtualKeyMappings
				$virtualKeyCode = $keyMappings[$timerData['Key']]
				
				if ($virtualKeyCode)
				{
					# Send key message to window (WM_KEYDOWN = 256)
					[Ftool]::fnPostMessage($timerData['WindowHandle'], 256, $virtualKeyCode, 0)
				}
				else
				{
					Write-Verbose "FTOOL: Unknown key '$($timerData['Key'])' for $($timerData['InstanceId'])" -ForegroundColor Yellow
				}
			}
			catch
			{
				Write-Verbose "FTOOL: Spammer timer error: $($_.Exception.Message)" -ForegroundColor Red
			}
		})
	
	$spamTimer.Start()
	return $spamTimer
}

<#
.SYNOPSIS
Toggles button state for start/stop
#>
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

<#
.SYNOPSIS
Rate limiting for UI events
#>
function CheckRateLimit
{
	param($controlId, $minInterval = 100)
	
	$currentTime = Get-Date
	if ($global:DashboardConfig.Resources.LastEventTimes.Contains($controlId) -and 
	($currentTime - $global:DashboardConfig.Resources.LastEventTimes[$controlId]).TotalMilliseconds -lt $minInterval)
	{
		return $false
	}
	
	$global:DashboardConfig.Resources.LastEventTimes[$controlId] = $currentTime
	return $true
}

<#
.SYNOPSIS
Adds form cleanup handler
#>
function AddFormCleanupHandler
{
	param($form)
	
	# Check if cleanup handler is already added
	if (-not $form.Tag.ExtensionCleanupAdded)
	{
		$form.Add_FormClosing({
				param($src, $e)
			
				# Use try/catch to ensure cleanup happens even if errors occur
				try
				{
					# Clean up all extension timers for this form
					$instanceId = $src.Tag.InstanceId
					if ($instanceId)
					{
						CleanupInstanceResources $instanceId
					}
				}
				catch
				{
					# Log the error but continue with form closing
					Write-Verbose "FTOOL: Error during form cleanup: $($_.Exception.Message)" -ForegroundColor Red
				}
			})
		
		# Mark that we"ve added the cleanup handler
		$form.Tag | Add-Member -NotePropertyName ExtensionCleanupAdded -NotePropertyValue $true -Force
	}
}

<#
.SYNOPSIS
Cleans up resources for an instance
#>
function CleanupInstanceResources
{
	param($instanceId)
	
	$instanceKey = "instance_$instanceId"
	
	# Find and remove all extension data for this instance
	$keysToRemove = @()
	
	# Make a copy of the active extensions array
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
			$keysToRemove += $key
		}
	}
	
	# Remove extension data
	foreach ($key in $keysToRemove)
	{
		$global:DashboardConfig.Resources.ExtensionData.Remove($key)
	}
	
	# Remove timers
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
	
	# Remove timer references
	foreach ($key in $timerKeysToRemove)
	{
		$global:DashboardConfig.Resources.Timers.Remove($key)
	}
	
	# Remove extension tracking for this instance
	if ($global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
	{
		$global:DashboardConfig.Resources.ExtensionTracking.Remove($instanceKey)
	}
	
	# Force garbage collection to release memory
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
}

<#
.SYNOPSIS
Stops and cleans up an Ftool form
#>
function Stop-FtoolForm
{
	param($Form)
	
	if (-not $Form -or $Form.IsDisposed)
	{
		return
	}
	
	try
	{
		# Get instance ID from form tag
		$instanceId = $Form.Tag.InstanceId
		if ($instanceId)
		{
			# Clean up instance resources
			CleanupInstanceResources $instanceId
			
			# Remove from global forms collection
			if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
			{
				$global:DashboardConfig.Resources.FtoolForms.Remove($instanceId)
			}
		}
		
		# Close and dispose form
		$Form.Close()
		$Form.Dispose()
	}
	catch
	{
		Write-Verbose "FTOOL: Error stopping Ftool form: $($_.Exception.Message)" -ForegroundColor Red
	}
}

<#
.SYNOPSIS
Removes an extension from a form
#>
function RemoveExtension
{
	param($form, $extKey)
	
	# Validate input parameters
	if (-not $form -or -not $extKey)
	{
		return $false 
	}
	
	# Check if extension data exists
	if (-not $global:DashboardConfig.Resources.ExtensionData -or 
		-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
	{
		return $false
	}
	
	# Get extension data
	$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
	if (-not $extData)
	{
		return $false 
	}
	
	# Extract extension information
	$extNum = [int]$extData.ExtNum  # Ensure extNum is an integer
	$instanceId = $extData.InstanceId
	$instanceKey = "instance_$instanceId"
	
	try
	{
		# Stop timer if running
		if ($extData.RunningSpammer)
		{
			$extData.RunningSpammer.Stop()
			$extData.RunningSpammer.Dispose()
			$extData.RunningSpammer = $null
			
			# Remove from global timers collection
			$timerKey = "ExtSpammer_${instanceId}_$extNum"
			if ($global:DashboardConfig.Resources.Timers.Contains($timerKey))
			{
				$global:DashboardConfig.Resources.Timers.Remove($timerKey)
			}
		}
		
		# Remove panel from form
		if ($extData.Panel -and -not $extData.Panel.IsDisposed)
		{
			$form.Controls.Remove($extData.Panel)
			$extData.Panel.Dispose()
		}
		
		# Initialize tracking if it doesn"t exist
		if (-not $global:DashboardConfig.Resources.ExtensionTracking -or 
			-not $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
		{
			InitializeExtensionTracking $instanceId
		}
		
		# Add extension number to reuse list - ensure it"s stored as an integer
		$intExtNum = [int]$extNum  # Convert to integer explicitly
		
		# Check if the number is already in the list
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
		
		# Remove from active extensions list
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions = 
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions | 
		Where-Object { $_ -ne $extKey }
		
		# Remove from global collection
		$global:DashboardConfig.Resources.ExtensionData.Remove($extKey)
		
		# Reposition extensions
		RepositionExtensions $form $instanceId
		
		return $true
	}
	catch
	{
		Write-Verbose "FTOOL: Error in RemoveExtension: $($_.Exception.Message)" -ForegroundColor Red
		return $false
	}
}

#endregion

#region Core Functions

<#
.SYNOPSIS
Processes selected row data and creates FTool form
#>
function FtoolSelectedRow
{
	param($row)
	
	# Validate row data
	if (-not $row -or -not $row.Cells -or $row.Cells.Count -lt 3)
	{
		Write-Verbose 'FTOOL: Invalid row data, skipping' -ForegroundColor Yellow
		return
	}
	
	# Get instance ID and window handle
	$instanceId = $row.Cells[2].Value.ToString()
	if (-not $row.Tag -or -not $row.Tag.MainWindowHandle)
	{
		Write-Verbose "FTOOL: Missing window handle for instance $instanceId, skipping" -ForegroundColor Yellow
		return
	}
	
	$windowHandle = $row.Tag.MainWindowHandle
	if (-not $instanceId)
	{
		Write-Verbose 'FTOOL: Missing instance ID, skipping' -ForegroundColor Yellow
		return
	}
	
	# Check if form already exists for this instance
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
	
	# Get target window position
	$targetWindowRect = New-Object Native+RECT
	[Native]::GetWindowRect($windowHandle, [ref]$targetWindowRect)
	
	# Get window title
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
	
	# Create the Ftool form
	$ftoolForm = CreateFtoolForm $instanceId $targetWindowRect $windowTitle $row
	
	# Add form to global collection
	$global:DashboardConfig.Resources.FtoolForms[$instanceId] = $ftoolForm
	
	# Show and activate form$ftoolForm.Show()
	$ftoolForm.BringToFront()
}

<#
.SYNOPSIS
Creates main FTool form with all controls
#>
function CreateFtoolForm
{
	param($instanceId, $targetWindowRect, $windowTitle, $row)
	
	# Create main Ftool form
	$ftoolForm = Set-UIElement -type 'Form' -visible $true -width 200 -height 130 -top ($targetWindowRect.Top + 30) -left ($targetWindowRect.Left + 10) -bg @(30, 30, 30) -fg @(255, 255, 255) -text "FTool - $instanceId" -startPosition 'Manual' -formBorderStyle 0 -opacity 1
	# Load custom icon if it exists
	if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon))
	{
		try
		{
			$icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon)
			$ftoolForm.Icon = $icon
		}
		catch
		{
			Write-Verbose "FTOOL: Failed to load icon from $($global:DashboardConfig.Paths.Icon): $_" -ForegroundColor Red
		}
	}
	if (-not $ftoolForm)
	{
		Write-Verbose "FTOOL: Failed to create form for $instanceId, skipping" -ForegroundColor Red
		return $null
	}
	
	# Create UI elements
	$headerPanel = Set-UIElement -type 'Panel' -visible $true -width 200 -height 20 -top 0 -left 0 -bg @(40, 40, 40)
	$ftoolForm.Controls.Add($headerPanel)
	
	# Create window title label
	$labelWinTitle = Set-UIElement -type 'Label' -visible $true -width 100 -height 20 -top 5 -left 5 -bg @(40, 40, 40, 0) -fg @(255, 255, 255) -text $row.Cells[1].Value -font (New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Regular))
	$headerPanel.Controls.Add($labelWinTitle)
	
	# Create Add button
	$btnAdd = Set-UIElement -type 'Button' -visible $true -width 14 -height 14 -top 3 -left 120 -bg @(40, 80, 80) -fg @(255, 255, 255) -text ([char]0x2795) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11))
	$headerPanel.Controls.Add($btnAdd)
	
	# Create Toggle Height button
	$btnShowHide = Set-UIElement -type 'Button' -visible $true -width 14 -height 14 -top 3 -left 140 -bg @(60, 60, 100) -fg @(255, 255, 255) -text ([char]0x25B2) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11))
	$headerPanel.Controls.Add($btnShowHide)
	
	# Create Close button
	$btnClose = Set-UIElement -type 'Button' -visible $true -width 14 -height 14 -top 3 -left 160 -bg @(200, 45, 45) -fg @(255, 255, 255) -text ([char]0x166D) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11))
	$headerPanel.Controls.Add($btnClose)
	
	# Create settings panel
	$panelSettings = Set-UIElement -type 'Panel' -visible $true -width 165 -height 60 -top 25 -left 10 -bg @(50, 50, 50)
	$ftoolForm.Controls.Add($panelSettings)
	
	# Create key selection button
	$btnKeySelect = Set-UIElement -type 'Button' -visible $true -width 55 -height 25 -top 4 -left 5 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'F1' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8))
	$panelSettings.Controls.Add($btnKeySelect)
	
	# Create interval text box
	$interval = Set-UIElement -type 'TextBox' -visible $true -width 50 -height 15 -top 5 -left 65 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '1000' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$panelSettings.Controls.Add($interval)
	
	# Create label for main control
	$name = Set-UIElement -type 'TextBox' -visible $true -width 40 -height 17 -top 5 -left 120 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'Main' -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular))
	$panelSettings.Controls.Add($name)
	
	# Create Start button
	$btnStart = Set-UIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 10 -bg @(0, 120, 215) -fg @(255, 255, 255) -text 'Start' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$panelSettings.Controls.Add($btnStart)
	
	# Create Stop button
	$btnStop = Set-UIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 67 -bg @(200, 50, 50) -fg @(255, 255, 255) -text 'Stop' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$btnStop.Enabled = $false
	$btnStop.Visible = $false
	$panelSettings.Controls.Add($btnStop)
	
	# Create form data object
	$formData = [PSCustomObject]@{
		InstanceId            = $instanceId
		SelectedWindow        = $row.Tag.MainWindowHandle
		BtnKeySelect          = $btnKeySelect
		Interval              = $interval
		Name                  = $name
		BtnStart              = $btnStart
		BtnStop               = $btnStop
		BtnAdd                = $btnAdd
		BtnClose              = $btnClose
		BtnShowHide           = $btnShowHide
		Form                  = $ftoolForm
		RunningSpammer        = $null
		WindowTitle           = $windowTitle
		Process               = $row.Tag
		OriginalLeft          = $targetWindowRect.Left
		IsCollapsed           = $false
		LastExtensionAdded    = 0
		ExtensionCount        = 0
		ControlToExtensionMap = @{}
		OriginalHeight        = 120
	}
	
	# Store form data in Tag property
	$ftoolForm.Tag = $formData
	
	# Load settings, create position timer, and add event handlers
	LoadFtoolSettings $formData
	CreatePositionTimer $formData
	AddFtoolEventHandlers $formData
	
	return $ftoolForm
}

<#
.SYNOPSIS
Adds event handlers to FTool form controls
#>
function AddFtoolEventHandlers
{
	param($formData)
	
	# Interval text box change event
	$formData.Interval.Add_TextChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			# Enforce minimum interval
			$intervalValue = 0
			if (-not [int]::TryParse($this.Text, [ref]$intervalValue) -or $intervalValue -lt 100)
			{
				$this.Text = '1000'
			}
		
			# Update settings
			UpdateSettings $data
		})

	# Name text box change event
	$formData.Name.Add_TextChanged({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$data = $form.Tag
	
		# Update settings
		UpdateSettings $data
	})
	
	# Key selection button click event
	$formData.BtnKeySelect.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$data = $form.Tag
		
		# Show key capture dialog
		$currentKey = $data.BtnKeySelect.Text
		$newKey = Show-KeyCaptureDialog $currentKey
		
		if ($newKey -and $newKey -ne $currentKey)
		{
			# Stop existing spammer if running before changing key
			if ($data.RunningSpammer)
			{
				$data.RunningSpammer.Stop()
				$data.RunningSpammer.Dispose()
				$data.RunningSpammer = $null
			}
			
			$data.BtnKeySelect.Text = $newKey
			UpdateSettings $data
			
			# Reset button states since spammer was stopped
			ToggleButtonState $data.BtnStart $data.BtnStop $false
		}
	})
	
	# Start button click event
	$formData.BtnStart.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			# Get key and interval values
			$keyValue = $data.BtnKeySelect.Text
			if (-not $keyValue -or $keyValue.Trim() -eq '')
			{
			[System.Windows.Forms.MessageBox]::Show('Please select a key', 'Error')
			return
			}
		
			# Enforce minimum interval
			$intervalNum = 0
			if (-not [int]::TryParse($data.Interval.Text, [ref]$intervalNum) -or $intervalNum -lt 100)
			{
				$intervalNum = 1000
				$data.Interval.Text = '1000'
			}
		
			# Update UI state
			ToggleButtonState $data.BtnStart $data.BtnStop $true
		
			# Create spammer timer
			$spamTimer = CreateSpammerTimer $data.SelectedWindow $keyValue $data.InstanceId $intervalNum
		
			# Store timer references
			$data.RunningSpammer = $spamTimer
			$global:DashboardConfig.Resources.Timers["ExtSpammer_$($data.InstanceId)_1"] = $spamTimer
		})
	
	# Stop button click event
	$formData.BtnStop.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			# Stop and clean up timer
			if ($data.RunningSpammer)
			{
				$data.RunningSpammer.Stop()
				$data.RunningSpammer.Dispose()
				$data.RunningSpammer = $null
			}
		
			# Update UI state
			ToggleButtonState $data.BtnStart $data.BtnStop $false
		
			# Remove from global timers collection
			if ($global:DashboardConfig.Resources.Timers.Contains("ExtSpammer_$($data.InstanceId)_1"))
			{
				$global:DashboardConfig.Resources.Timers.Remove("ExtSpammer_$($data.InstanceId)_1")
			}
		})
	
	# ShowHide button click event
	$formData.BtnShowHide.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			# Check rate limit
			if (-not (CheckRateLimit $this.GetHashCode() 200))
			{
				return 
			}
		
			if ($data.IsCollapsed)
			{
				# Restore original height
				$form.Height = $data.OriginalHeight
				$data.IsCollapsed = $false
				$this.Text = [char]0x25B2  # Unicode UP ARROW
			}
			else
			{
				# Store current height and collapse
				$data.OriginalHeight = $form.Height
				$form.Height = 35
				$data.IsCollapsed = $true
				$this.Text = [char]0x25BC  # Unicode DOWN ARROW
			}
		})
	
	# Add button click event with rate limiting
	$formData.BtnAdd.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag -or $form.Tag.IsCollapsed)
			{
				return 
			}
			$data = $form.Tag
		
			# Check rate limit
			if (-not (CheckRateLimit $this.GetHashCode() 200))
			{
				return 
			}
		
			# Limit the number of extensions
			if ($data.ExtensionCount -ge 8)
			{
				[System.Windows.Forms.MessageBox]::Show('Maximum number of extensions reached.', 'Warning')
				return
			}
		
			# Increment extension count
			$data.ExtensionCount++
		
			# Initialize instance-specific extension tracking
			InitializeExtensionTracking $data.InstanceId
		
			# Get the next extension number for this instance
			$extNum = GetNextExtensionNumber $data.InstanceId
		
			# Create extension panel and controls
			$extData = CreateExtensionPanel $form $currentHeight $extNum $data.InstanceId $data.SelectedWindow
		
			# Find or create profile
			$profilePrefix = FindOrCreateProfile $data.WindowTitle
		
			# Load settings for extension
			LoadExtensionSettings $extData $profilePrefix
		
			# Add event handlers for extension
			AddExtensionEventHandlers $extData $data
		
			# Add form cleanup handler if not already added
			AddFormCleanupHandler $form
		
			RepositionExtensions $form $extData.InstanceId
		
			$form.Refresh()
		})
	
	# Close button click event
	$formData.BtnClose.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			# Clean up resources
			CleanupInstanceResources $data.InstanceId
		
			# Remove from global collections
			if ($global:DashboardConfig.Resources.FtoolForms.Contains($data.InstanceId))
			{
				$global:DashboardConfig.Resources.FtoolForms.Remove($data.InstanceId)
			}
		
			# Close form
			$form.Close()
			$form.Dispose()
		})
	
	# Form closed event for main form
	$formData.Form.Add_FormClosed({
			param($src, $e)
		
			# Clean up all extension timers for this form
			$instanceId = $src.Tag.InstanceId
			if ($instanceId)
			{
				CleanupInstanceResources $instanceId
			}
		})
}

<#
.SYNOPSIS
Creates extension panel with controls
#>
function CreateExtensionPanel
{
	param($form, $currentHeight, $extNum, $instanceId, $windowHandle)
	
	# Create extension panel
	$panelExt = Set-UIElement -type 'Panel' -visible $true -width 165 -height 60 -top 0 -left 10 -bg @(50, 50, 50)
	$form.Controls.Add($panelExt)
	
	# Create key selection button for extension
	$btnKeySelectExt = Set-UIElement -type 'Button' -visible $true -width 55 -height 25 -top 4 -left 5 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'F1' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8))
	$panelExt.Controls.Add($btnKeySelectExt)
	
	# Create interval text box for extension
	$intervalExt = Set-UIElement -type 'TextBox' -visible $true -width 50 -height 15 -top 5 -left 65 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '1000' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$panelExt.Controls.Add($intervalExt)
	
	# Create Start button for extension
	$btnStartExt = Set-UIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 10 -bg @(0, 120, 215) -fg @(255, 255, 255) -text 'Start' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$panelExt.Controls.Add($btnStartExt)
	
	# Create Stop button for extension
	$btnStopExt = Set-UIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 67 -bg @(200, 50, 50) -fg @(255, 255, 255) -text 'Stop' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$btnStopExt.Enabled = $false
	$btnStopExt.Visible = $false
	$panelExt.Controls.Add($btnStopExt)
	
	# Create extension name
	$nameExt = Set-UIElement -type 'TextBox' -visible $true -width 40 -height 17 -top 5 -left 120 -bg @(40, 40, 40) -fg @(255, 255, 255) -text "Name" -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular))
	$panelExt.Controls.Add($nameExt)
	
	# Create Remove button for the extension
	$btnRemoveExt = Set-UIElement -type 'Button' -visible $true -width 40 -height 20 -top 35 -left 120 -bg @(150, 50, 50) -fg @(255, 255, 255) -text 'Close' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 7))
	$panelExt.Controls.Add($btnRemoveExt)
	
	# Store extension data in global collection
	$extKey = "ext_${instanceId}_$extNum"
	$instanceKey = "instance_$instanceId"
	
	# Use efficient data structure
	$extData = [PSCustomObject]@{
		Panel          = $panelExt
		BtnKeySelect   = $btnKeySelectExt
		Interval       = $intervalExt
		BtnStart       = $btnStartExt
		BtnStop        = $btnStopExt
		BtnRemove      = $btnRemoveExt
		Name           = $nameExt
		ExtNum         = $extNum
		InstanceId     = $instanceId
		WindowHandle   = $windowHandle
		RunningSpammer = $null
		Position       = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions.Count
	}
	
	$global:DashboardConfig.Resources.ExtensionData[$extKey] = $extData
	
	# Add to active extensions list
	if (-not $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions.Contains($extKey))
	{
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions += $extKey
	}
	
	return $extData
}

<#
.SYNOPSIS
Adds event handlers to extension controls
#>
function AddExtensionEventHandlers
{
	param($extData, $formData)
	
	# Interval text box change event for extension
	$extData.Interval.Add_TextChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			# Check rate limit
			if (-not (CheckRateLimit $this.GetHashCode()))
			{
				return 
			}
		
			# Find which extension this belongs to
			$extKey = FindExtensionKeyByControl $this 'Interval'
			if (-not $extKey)
			{
				return 
			}
		
			# Check if extension data still exists
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
		
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
		
			# Enforce minimum interval
			$intervalValue = 0
			if (-not [int]::TryParse($this.Text, [ref]$intervalValue) -or $intervalValue -lt 100)
			{
				$this.Text = '1000'
				$intervalValue = 1000
			}
		
			# Update settings
			UpdateSettings $data $extData
		})

	# Name text box change event for extension
	$extData.Name.Add_TextChanged({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$data = $form.Tag
	
		# Check rate limit
		if (-not (CheckRateLimit $this.GetHashCode()))
		{
			return 
		}
	
		# Find which extension this belongs to
		$extKey = FindExtensionKeyByControl $this 'Name'
		if (-not $extKey)
		{
			return 
		}
	
		# Check if extension data still exists
		if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
		{
			return 
		}
	
		$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
		if (-not $extData)
		{
			return 
		}
	
		# Update settings
		UpdateSettings $data $extData
	})
	
	# Key selection button click event for extension
	$extData.BtnKeySelect.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$data = $form.Tag
		
		# Find which extension this belongs to
		$extKey = FindExtensionKeyByControl $this 'BtnKeySelect'
		if (-not $extKey)
		{
			return 
		}
		
		# Check if extension data still exists
		if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
		{
			return 
		}
		
		$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
		if (-not $extData)
		{
			return 
		}
		
		# Show key capture dialog
		$currentKey = $extData.BtnKeySelect.Text
		$newKey = Show-KeyCaptureDialog $currentKey
		
		if ($newKey -and $newKey -ne $currentKey)
		{
			# Stop existing extension spammer if running before changing key
			if ($extData.RunningSpammer)
			{
				$extData.RunningSpammer.Stop()
				$extData.RunningSpammer.Dispose()
				$extData.RunningSpammer = $null
			}
			
			$extData.BtnKeySelect.Text = $newKey
			UpdateSettings $data $extData
			
			# Reset extension button states since spammer was stopped
			ToggleButtonState $extData.BtnStart $extData.BtnStop $false
		}
	})
	
	# Start button click event for extension
	$extData.BtnStart.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			# Check rate limit
			if (-not (CheckRateLimit $this.GetHashCode() 500))
			{
				return 
			}
		
			# Find which extension this belongs to
			$extKey = FindExtensionKeyByControl $this 'BtnStart'
			if (-not $extKey)
			{
				return 
			}
		
			# Check if extension data still exists
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
		
			# Get key and interval values
			$keyValue = $extData.BtnKeySelect.Text
			 if (-not $keyValue -or $keyValue.Trim() -eq '')
			{
			 [System.Windows.Forms.MessageBox]::Show('Please select a key', 'Error')
			 return
			}
		
			# Enforce minimum interval
			$intervalNum = 0
		
			# Check if Interval exists before accessing it
			if (-not $extData.Interval -or -not $extData.Interval.Text)
			{
				$intervalNum = 1000
			}
			else
			{
				if (-not [int]::TryParse($extData.Interval.Text, [ref]$intervalNum) -or $intervalNum -lt 100)
				{
					$intervalNum = 1000
					if ($extData.Interval)
					{
						$extData.Interval.Text = '1000'
					}
				}
			}
		
			# Update UI state
			ToggleButtonState $extData.BtnStart $extData.BtnStop $true
		
			# Check if a timer already exists and clean it up
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
		
			# Create spammer timer
			$spamTimer = CreateSpammerTimer $data.SelectedWindow $keyValue $data.InstanceId $intervalNum $extNum $extKey
		
			# Store timer references
			$extData.RunningSpammer = $spamTimer
			$global:DashboardConfig.Resources.Timers[$timerKey] = $spamTimer
		})
	
	# Stop button click event for extension
	$extData.BtnStop.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			# Check rate limit
			if (-not (CheckRateLimit $this.GetHashCode() 500))
			{
				return 
			}
		
			# Find which extension this belongs to
			$extKey = FindExtensionKeyByControl $this 'BtnStop'
			if (-not $extKey)
			{
				return 
			}
		
			# Check if extension data still exists
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
		
			# Stop and clean up timer
			if ($extData.RunningSpammer)
			{
				$extData.RunningSpammer.Stop()
				$extData.RunningSpammer.Dispose()
				$extData.RunningSpammer = $null
			}
		
			# Update UI state
			ToggleButtonState $extData.BtnStart $extData.BtnStop $false
		
			# Remove from global timers collection
			$timerKey = "ExtSpammer_$($data.InstanceId)_$extNum"
			if ($global:DashboardConfig.Resources.Timers.Contains($timerKey))
			{
				$global:DashboardConfig.Resources.Timers.Remove($timerKey)
			}
		})
	
	# Remove button click event for extension
	$extData.BtnRemove.Add_Click({
			try
			{
				$form = $this.FindForm()
				if (-not $form -or -not $form.Tag)
				{
					return 
				}
			
				# Check rate limit
				if (-not (CheckRateLimit $this.GetHashCode() 1000))
				{
					return 
				}
			
				# Find which extension this belongs to
				$extKey = FindExtensionKeyByControl $this 'BtnRemove'
				if (-not $extKey)
				{
					return 
				}
			
				# Disable the button while removing to prevent multiple clicks
				$this.Enabled = $false
			
				# Remove the extension
				RemoveExtension $form $extKey
			
				# Decrement extension count
				$form.Tag.ExtensionCount--
			}
			catch
			{
				Write-Verbose "FTOOL: Error in Remove button handler: $($_.Exception.Message)" -ForegroundColor Red
			}
		})
}

#endregion

#region Module Exports

# Export public functions
Export-ModuleMember -Function FtoolSelectedRow, Stop-FtoolForm, Show-KeyCaptureDialog, Get-VirtualKeyMappings

#endregion