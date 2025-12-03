<# login.psm1 
	.SYNOPSIS
		Login Automation Module for Entropia Dashboard.

	.DESCRIPTION
		This module provides a comprehensive login automation system for the Entropia Dashboard:
		- Processes multiple game clients sequentially based on selection
		- Monitors client processes and window states
		- Provides thread-safe logging of login operations
		- Handles cleanup of resources when operations complete

	.NOTES
		Author: Immortal / Divine
		Version: 1.2
		Requires: PowerShell 5.1, .NET Framework 4.5+, classes.psm1, datagrid.psm1
#>

#region Helper Functions

<#
.SYNOPSIS
Restores a minimized window
#>
function Restore-Window
{
	param(
		[System.Diagnostics.Process]$Process
	)
	
	if ($Process -and $Process.MainWindowHandle -ne [IntPtr]::Zero)
	{
		if ([Custom.Native]::IsWindowMinimized($Process.MainWindowHandle))
		{
			Write-Verbose "LOGIN: Restoring minimized window for PID $($Process.Id)" -ForegroundColor DarkGray
			[Custom.Native]::BringToFront($Process.MainWindowHandle)
			Start-Sleep -Milliseconds 100
		}
	}
}

<#
.SYNOPSIS
Brings window to foreground with validation
#>
function Set-WindowForeground
{
	param(
		[System.Diagnostics.Process]$Process
	)
	
	if (-not $Process -or $Process.MainWindowHandle -eq [IntPtr]::Zero)
	{
		return $false
	}
	
	$script:ScriptInitiatedMove = $true
	$result = [Custom.Native]::BringToFront($Process.MainWindowHandle)
	Write-Verbose "LOGIN: Brought window to front: $result" -ForegroundColor Green
	Start-Sleep -Milliseconds 100
	
	# Reset the script-initiated move flag
	$script:ScriptInitiatedMove = $false
	
	# Validate if the window is now the foreground window
	$foregroundHandle = [Custom.Native]::GetForegroundWindow()
	if ($foregroundHandle -ne $Process.MainWindowHandle)
	{
		Write-Verbose "LOGIN: Failed to bring window to foreground for PID $($Process.Id)" -ForegroundColor Red
		return $false
	}
	
	return $true
}

<#
.SYNOPSIS
Checks if user moved the mouse or changed focus
#>
function Test-UserMouseIntervention
{
	param()
	
	# If we're currently performing a script-initiated move, don't detect as intervention
	if ($script:ScriptInitiatedMove)
	{
		return $false
	}
	
	# Get current mouse position
	$currentPosition = [System.Windows.Forms.Cursor]::Position
	$currentTime = Get-Date
	
	# If the mouse position has changed significantly since our last script action
	if ([Math]::Abs($currentPosition.X - $script:LastScriptMouseTarget.X) -gt 5 -or 
		[Math]::Abs($currentPosition.Y - $script:LastScriptMouseTarget.Y) -gt 5)
	{
		
		# Check if enough time has passed since our last script-initiated move
		$timeSinceLastMove = New-TimeSpan -Start $script:LastScriptMouseMoveTime -End $currentTime
		
		# Only consider it user intervention if it's been more than 500ms since our last script move
		# This prevents false positives when the script itself is moving the mouse
		if ($timeSinceLastMove.TotalMilliseconds -gt 500)
		{
			Write-Verbose 'LOGIN: User mouse intervention detected' -ForegroundColor Yellow
			return $true
		}
	}
	
	return $false
}

<#
.SYNOPSIS
Waits for an application to be responsive
#>
function Wait-ForResponsive
{
	param(
		[System.Diagnostics.Process]$Monitor
	)
	
	$waitInterval = 100
	$maxAttempts = 40  # 8 seconds max
	$isResponsive = $false
	
	for ($i = 0; $i -lt $maxAttempts; $i++)
	{
		# Check if process is still valid
		if (-not $Monitor -or $Monitor.HasExited -or $Monitor.MainWindowHandle -eq [IntPtr]::Zero)
		{
			return $false
		}
		
		$responsiveTask = [Custom.Native]::ResponsiveAsync($Monitor.MainWindowHandle, 100)
		if ($responsiveTask.Result)
		{
			$isResponsive = $true
			Start-Sleep -Milliseconds $waitInterval
			break
		}
		
		Start-Sleep -Milliseconds $waitInterval
	}
	
	if (-not $isResponsive)
	{
		Write-Verbose "LOGIN: Window unresponsive for PID $($Monitor.Id)" -ForegroundColor Yellow
	}
	
	return $isResponsive
}

<#
.SYNOPSIS
Waits for file to be accessible
#>
function Wait-ForFileAccess
{
	param(
		[string]$FilePath
	)
	
	$maxAttempts = 40  # 4 second max
	$waitInterval = 100  # 50ms
	
	for ($i = 0; $i -lt $maxAttempts; $i++)
	{
		try
		{
			# Check if file exists and can be opened
			if (Test-Path -Path $FilePath)
			{
				$fileStream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
				if ($fileStream)
				{
					$fileStream.Close()
					$fileStream.Dispose()
					return $true
				}
			}
		}
		catch
		{
			# File is locked or doesn't exist
		}
		
		Start-Sleep -Milliseconds $waitInterval
	}
	
	return $false
}

<#
.SYNOPSIS
Thread-safe log writing
#>
function Write-LogWithRetry
{
	param(
		[string]$FilePath,
		[string]$Value
	)
	
	$maxAttempts = 10
	$waitInterval = 100
	
	for ($i = 0; $i -lt $maxAttempts; $i++)
	{
		try
		{
			if (-not (Test-Path -Path (Split-Path -Path $FilePath -Parent)))
			{
				New-Item -Path (Split-Path -Path $FilePath -Parent) -ItemType Directory -Force | Out-Null
			}
			
			Set-Content -Path $FilePath -Value $Value -Force
			return $true
		}
		catch
		{
			Start-Sleep -Milliseconds $waitInterval
		}
	}
	
	Write-Verbose "LOGIN: Failed to write to log file $FilePath" -ForegroundColor Red
	return $false
}

<#
.SYNOPSIS
Simulates a mouse click at specific coordinates without using ftool.dll
#>
function Invoke-MouseClick
{
	param(
		[int]$X,
		[int]$Y
	)
	
	# Store the current cursor position to restore later
	$originalPosition = [System.Windows.Forms.Cursor]::Position
	
	# Track that this movement is script-initiated
	$script:ScriptInitiatedMove = $true
	$script:LastScriptMouseTarget = New-Object System.Drawing.Point($X, $Y)
	$script:LastScriptMouseMoveTime = Get-Date
	
	# Get the handle of the foreground window
	$hWnd = [Custom.Native]::GetForegroundWindow()
	
	try
	{
		Write-Verbose "LOGIN: Moving cursor from ($($originalPosition.X),$($originalPosition.Y)) to ($X,$Y)" -ForegroundColor DarkGray
		
		# Force cursor position
		Start-Sleep -Milliseconds 10
		[void][Custom.Native]::SetCursorPos($X, $Y)
		Start-Sleep -Milliseconds 10
			[void][Custom.Native]::SetCursorPos($X, $Y)
		Start-Sleep -Milliseconds 10
		
		$newPos = [System.Windows.Forms.Cursor]::Position
		Start-Sleep -Milliseconds 10
		Write-Verbose "LOGIN: Cursor position after move: ($($newPos.X),$($newPos.Y))" -ForegroundColor DarkGray
		
		# If position is still off by more than 5 pixels, try one more time
		if ([Math]::Abs($newPos.X - $X) -gt 5 -or [Math]::Abs($newPos.Y - $Y) -gt 5)
		{
			Write-Verbose "LOGIN: Cursor position verification failed. Expected: ($X,$Y), Actual: ($($newPos.X),$($newPos.Y))" -ForegroundColor Red
			
			# Try one more time with a longer delay
			Start-Sleep -Milliseconds 50
			[void][Custom.Native]::SetCursorPos($X, $Y)
			Start-Sleep -Milliseconds 50
			[void][Custom.Native]::SetCursorPos($X, $Y)
			Start-Sleep -Milliseconds 50
			
			# Check again
			$newPos = [System.Windows.Forms.Cursor]::Position
			Start-Sleep -Milliseconds 50
			Write-Verbose "LOGIN: Cursor position after second attempt: ($($newPos.X),$($newPos.Y))" -ForegroundColor Gray
			if ([Math]::Abs($newPos.X - $X) -gt 5 -or [Math]::Abs($newPos.Y - $Y) -gt 5)
			{
				Write-Verbose "LOGIN: Cursor position verification failed. Expected: ($X,$Y), Actual: ($($newPos.X),$($newPos.Y)). Stopping login." -ForegroundColor Red
				return $false
			}
		}
		
		# Use mouse_event for more reliable clicking at the CURRENT position
		# This is important - we click where the cursor actually is
		$MOUSEEVENTF_LEFTDOWN = 0x0002
		$MOUSEEVENTF_LEFTUP = 0x0004
		
		# Perform click
		[Custom.Native]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
		Start-Sleep -Milliseconds 10
		[Custom.Native]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
		Start-Sleep -Milliseconds 10
		
		$neverRestarting = $false # Default to false if setting doesn't exist
		if ($global:DashboardConfig.Config.Contains('Login') -and 
			$global:DashboardConfig.Config['Login'].Contains('NeverRestartingCollectorLogin'))
		{
			$neverRestarting = [bool]([int]$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin'])
		}

		# Fallback to SendMessage if needed - using the ACTUAL coordinates (collector DC workaround)
		if ($hWnd -ne [IntPtr]::Zero -and $neverRestarting -eq $true)
		{
			$currentPos = [System.Windows.Forms.Cursor]::Position
			$lparam = ($currentPos.Y -shl 16) -bor $currentPos.X
			
			$windowsMouseDown = 0x0201  # WM_LBUTTONDOWN
			$windowsMouseUp = 0x0202    # WM_LBUTTONUP
			
			[Custom.Native]::SendMessage($hWnd, $windowsMouseDown, 0, $lparam)
			Start-Sleep -Milliseconds 20
			[Custom.Native]::SendMessage($hWnd, $windowsMouseUp, 0, $lparam)
			Start-Sleep -Milliseconds 20
			Write-Verbose "Mouse click was performed." -ForegroundColor DarkGray
		}
		
		# Small delay to ensure any next clicks register
		Start-Sleep -Milliseconds 50
	} 
	catch
	{
		Write-Verbose "LOGIN: Mouse click simulation failed: $_" -ForegroundColor Red
	}
	finally
	{
		
		# Reset script-initiated flag after a short delay
		Start-Sleep -Milliseconds 50
		$script:ScriptInitiatedMove = $false
		Start-Sleep -Milliseconds 50
	}
}

<#
.SYNOPSIS
Simulates a key press
#>
function Invoke-KeyPress
{
	param(
		[int]$VirtualKeyCode
	)
	
	# Get the handle of the foreground window
	$hWnd = [Custom.Native]::GetForegroundWindow()
	[Custom.Native]::BringToFront($Process.MainWindowHandle) | Out-Null

	# Simulate key press (this is still using ftool.dll)
	[Custom.Ftool]::fnPostMessage($hWnd, 0x0100, $VirtualKeyCode, 0) # WM_KEYDOWN
	Start-Sleep -Milliseconds 20
	[Custom.Ftool]::fnPostMessage($hWnd, 0x0101, $VirtualKeyCode, 0) # WM_KEYUP
	
	Start-Sleep -Milliseconds 100
}

#endregion

#region Core Functions

# Login state tracking
$script:LoginState = @{
	Active      = $false
	LastAttempt = $null
	RetryCount  = 0
	MaxRetries  = 1
	Timeout     = 120 # seconds
}

<#
.SYNOPSIS
Logs into selected game client
#>
function LoginSelectedRow
{
	param(
		[System.Windows.Forms.DataGridViewRow]$Row,
		[string]$LogFolder = ($global:DashboardConfig.Config['LauncherPath']['LauncherPath'] -replace '\\Launcher\.exe$', ''),
		[string]$LogFilePath = "$($LogFolder)\Log\network_$(Get-Date -Format 'yyyyMMdd').log"
	)
	
	$global:DashboardConfig.State.LoginActive = $true
	
	# Initialize global variables if they don't exist
	if (-not $script:LastScriptMouseTarget)
	{
		$script:LastScriptMouseTarget = New-Object System.Drawing.Point(0, 0)
		$script:LastScriptMouseMoveTime = Get-Date
		$script:ScriptInitiatedMove = $false
	}
	
	# Get UI reference
	$UI = $global:DashboardConfig.UI
	if (-not $UI)
	{
		Write-Verbose 'LOGIN: UI reference not found' -ForegroundColor Red
		return
	}
	
	# Ensure log file exists
	if (-not (Test-Path $LogFilePath))
	{
		Write-Verbose 'LOGIN: Creating new log file' -ForegroundColor Yellow
		try
		{
			New-Item -Path $LogFilePath -ItemType File -Force | Out-Null
		}
		catch
		{
			Write-Verbose "LOGIN: Error creating log file: $_" -ForegroundColor Red
			return
		}
	}
	else
	{
		# Clear log file
		Write-LogWithRetry -FilePath $LogFilePath -Value ''
	}
	
	# Simple property change handler that updates button appearance
	if ($global:DashboardConfig.State.LoginActive -eq $true) {
		$global:DashboardConfig.UI.LoginButton.FlatStyle = 'Popup'
	} else {
		$global:DashboardConfig.UI.LoginButton.FlatStyle = 'Flat'
	}
	
	Write-Verbose 'LOGIN: Login process started' -ForegroundColor Cyan
	
	# Check if rows are selected
	if ($UI.DataGridFiller.SelectedRows.Count -eq 0)
	{
		Write-Verbose 'LOGIN: No clients selected' -ForegroundColor Yellow
		return
	}
	
	# Sort rows by index to process in order
	$sortedRows = $UI.DataGridFiller.SelectedRows | Sort-Object { $_.Cells[0].Value -as [int] }
	
	# Process each selected row
	foreach ($row in $sortedRows)
	{
		try
		{
			$process = $row.Tag
			if (-not $process)
			{
				Write-Verbose 'LOGIN: No process associated with row' -ForegroundColor Yellow
				continue 
			}
			
			# Reset state flags
			$null = $foundCERT
			$null = $foundLogin
			$null = $foundCacheJoin
			$foundCERT = $false
			$foundLogin = $false
			$foundCacheJoin = $false
			
			# Start log monitoring job with timeout mechanism
			$logMonitorJob = Start-Job -ArgumentList $LogFilePath -ScriptBlock {
				param($LogFilePath)
				
				$null = $foundCERT
				$null = $foundLogin
				$null = $foundCacheJoin
				$foundCERT = $false
				$foundLogin = $false
				$foundCacheJoin = $false
				
				if (-not (Test-Path $LogFilePath))
				{
					Write-Verbose "Log file does not exist: $LogFilePath" -ForegroundColor Yellow
					return
				}
				
				# Add a timeout mechanism to prevent infinite loops
				$startTime = Get-Date
				$timeout = New-TimeSpan -Minutes 2  # 2 minute timeout
				
				while ((New-TimeSpan -Start $startTime -End (Get-Date)) -lt $timeout)
				{
					try
					{
						$line = Get-Content -Path $LogFilePath -Tail 1 -ErrorAction SilentlyContinue
						$lines = Get-Content -Path $LogFilePath -Tail 6 -ErrorAction SilentlyContinue
						
						if ($line -eq '2 - CERT_SRVR_LIST')
						{
							Start-Sleep -Milliseconds 50
							$foundCERT = $true
							Write-Output 'CERT_FOUND'
							Start-Sleep -Milliseconds 50
						}
						elseif ($line -eq '6 - LOGIN_PLAYER_LIST')
						{
							Start-Sleep -Milliseconds 50
							$foundLogin = $true
							Write-Output 'LOGIN_FOUND'
							Start-Sleep -Milliseconds 50
						}
						
						# Check for cache join pattern
						if ($lines -and $lines.Count -ge 6)
						{
							for ($i = 0; $i -le $lines.Count - 3; $i++)
							{
								if ($lines[$i] -match '13 - CACHE_ACK_JOIN' -and 
									$lines[$i + 2] -match '13 - CACHE_ACK_JOIN' -and 
									$lines[$i + 4] -match '13 - CACHE_ACK_JOIN')
								{
									Start-Sleep -Milliseconds 50
									$foundCacheJoin = $true
									Write-Output 'CACHE_FOUND'
									Start-Sleep -Milliseconds 50
									break
								}
							}
						}
					}
					catch
					{
						# Ignore errors during log reading
						Write-Output "ERROR: $_"
						Start-Sleep -Milliseconds 10
					}
					
					Start-Sleep -Milliseconds 100
				}
				
				# If we reach here, we've timed out
				Write-Output 'TIMEOUT'
			}
			
			# Get client position from the index column (column 0)
			# This is the actual displayed index, not the row position
			$entryNumber = [int]$row.Cells[0].Value
			Write-Verbose "LOGIN: Processing entry $entryNumber (PID $($process.Id))" -ForegroundColor Cyan
			
			try
			{
				# Restore window if minimized
				Restore-Window -Process $process 
				
				
				# Bring window to foreground
				if (-not (Set-WindowForeground -Process $process ))
				{
					
					continue
				}
				
				
				# Calculate target click position (center of client window)
				$rect = New-Object Custom.Native+RECT
				if (-not [Custom.Native]::GetWindowRect($process.MainWindowHandle, [ref]$rect))
				{
					
					continue
				}
				
				$centerX = [int](($rect.Left + $rect.Right) / 2) + 25
				$centerY = [int](($rect.Top + $rect.Bottom) / 2)
				
				
				
				
				# Adjust Y position based on row index value from column 0
				$adjustedY = $centerY
				if ($entryNumber -ge 1 -and $entryNumber -le 5)
				{
					$yOffset = ($entryNumber - 3) * 18
					$adjustedY = $centerY + $yOffset
					
				}
				
				if ($entryNumber -ge 6 -and $entryNumber -le 10)
				{
					$yOffset = ($entryNumber - 8) * 18
					$adjustedY = $centerY + $yOffset
					
				}
				
				
				if ($entryNumber -ge 6 -and $entryNumber -le 10)
				{
					$scrollCenterX = $centerX + 145
					$scrollCenterY = $centerY + 28
					$adjustedY = $centerY + $yOffset
					
					Invoke-MouseClick -X $scrollCenterX -Y $scrollCenterY
					
				}
				
				# Perform first click with explicit coordinates
				
				Invoke-MouseClick -X $centerX -Y $adjustedY
				Invoke-MouseClick -X $centerX -Y $adjustedY		
				
				# Wait for process to be responsive
				
				Wait-ForResponsive -Monitor $process
				
				# Wait for CERT screen with timeout
				$certTimeout = New-TimeSpan -Seconds 20
				$certStartTime = Get-Date
				while (-not $foundCERT -and ((New-TimeSpan -Start $certStartTime -End (Get-Date)) -lt $certTimeout))
				{
					if (-not (Wait-ForFileAccess -FilePath $LogFilePath))
					{
						Write-Verbose 'LOGIN: File lock timeout, retrying...' -ForegroundColor Yellow
						Start-Sleep -Milliseconds 200
						continue
					}
					
					$logStatus = Receive-Job -Job $logMonitorJob -Keep
					if ($logStatus -contains 'CERT_FOUND')
					{
						$foundCERT = $true
						Start-Sleep -Milliseconds 100
						Write-Verbose "LOGIN: CERT found for PID $($process.Id)" -ForegroundColor Green
					}
					
					if ($logStatus -contains 'TIMEOUT')
					{
						Write-Verbose 'LOGIN: Log monitoring timed out' -ForegroundColor Red
						break
					}
					
					if (Test-UserMouseIntervention)
					{
						Write-Verbose 'LOGIN: User intervention detected - stopping' -ForegroundColor Yellow
						if ($logMonitorJob)
						{
							Stop-Job -Job $logMonitorJob -ErrorAction SilentlyContinue
							Remove-Job -Job $logMonitorJob -Force -ErrorAction SilentlyContinue
						}
						$global:DashboardConfig.State.LoginActive = $false
						return
					}
					

				}
				
				if (-not $foundCERT)
				{
					Write-Verbose 'LOGIN: CERT not found within timeout period' -ForegroundColor Yellow
					continue
				}
				
				# Press Enter to continue
				Write-LogWithRetry -FilePath $LogFilePath -Value ''
				Write-Verbose 'LOGIN: Pressing Enter at CERT' -ForegroundColor Green
				Invoke-KeyPress -VirtualKeyCode 0x0D  # Enter key
				Wait-ForResponsive -Monitor $process
				
				# Wait for LOGIN screen with timeout
				$loginTimeout = New-TimeSpan -Seconds 60
				$loginStartTime = Get-Date
				while (-not $foundLogin -and ((New-TimeSpan -Start $loginStartTime -End (Get-Date)) -lt $loginTimeout))
				{
					if (-not (Wait-ForFileAccess -FilePath $LogFilePath))
					{
						Write-Verbose 'LOGIN: File lock timeout, retrying...' -ForegroundColor Yellow
						Start-Sleep -Milliseconds 200
						continue
					}
					
					$logStatus = Receive-Job -Job $logMonitorJob -Keep
					if ($logStatus -contains 'LOGIN_FOUND')
					{
						$foundLogin = $true
						Start-Sleep -Milliseconds 100
						Write-Verbose "LOGIN: LOGIN found for PID $($process.Id)" -ForegroundColor Green
					}
					
					if ($logStatus -contains 'TIMEOUT')
					{
						Write-Verbose 'LOGIN: Log monitoring timed out' -ForegroundColor Red
						break
					}
					
					if (Test-UserMouseIntervention)
					{
						Write-Verbose 'LOGIN: User intervention detected - stopping' -ForegroundColor Yellow
						if ($logMonitorJob)
						{
							Stop-Job -Job $logMonitorJob -ErrorAction SilentlyContinue
							Remove-Job -Job $logMonitorJob -Force -ErrorAction SilentlyContinue
						}
						$global:DashboardConfig.State.LoginActive = $false
						return
					}
					
				}
				
				if (-not $foundLogin)
				{
					Write-Verbose 'LOGIN: LOGIN not found within timeout period' -ForegroundColor Yellow
					continue
				}
				
				# Select login position if configured
				$loginPosSetting = $null
				if ($global:DashboardConfig.Config -and 
					$global:DashboardConfig.Config['Login'] -and 
					$global:DashboardConfig.Config['Login']['Login'])
				{
					
					$loginPositions = $global:DashboardConfig.Config['Login']['Login'] -split ','
					if ($entryNumber -le $loginPositions.Count)
					{
						$loginPosSetting = $loginPositions[$entryNumber - 1]
					}
				}
				
				if ($loginPosSetting)
				{
					$num = [int]$loginPosSetting
					$rightArrowCount = $num - 1
					
					for ($r = 1; $r -le $rightArrowCount; $r++)
					{
						Write-Verbose 'LOGIN: Pressing right arrow for login position selection' -ForegroundColor Cyan
						Invoke-KeyPress -VirtualKeyCode 0x27  # Right arrow key
					}
					
					Write-Verbose "LOGIN: Selected login position $num for PID $($process.Id)" -ForegroundColor Green
				}
				
				# Press Enter to login
				Write-LogWithRetry -FilePath $LogFilePath -Value ''
				Write-Verbose 'LOGIN: Pressing Enter to login' -ForegroundColor Green
				Invoke-KeyPress -VirtualKeyCode 0x0D  # Enter key
				Start-Sleep -Milliseconds 500
				Wait-ForResponsive -Monitor $process
				
				# Wait for CACHE_JOIN with timeout
				$cacheTimeout = New-TimeSpan -Seconds 60  # Longer timeout for login
				$cacheStartTime = Get-Date
				while (-not $foundCacheJoin -and ((New-TimeSpan -Start $cacheStartTime -End (Get-Date)) -lt $cacheTimeout))
				{
					if (-not (Wait-ForFileAccess -FilePath $LogFilePath))
					{
						Write-Verbose 'LOGIN: File lock timeout, retrying...' -ForegroundColor Yellow
						Start-Sleep -Milliseconds 200
						continue
					}
					
					$logStatus = Receive-Job -Job $logMonitorJob -Keep
					if ($logStatus -contains 'CACHE_FOUND')
					{
						$foundCacheJoin = $true
						Start-Sleep -Milliseconds 100
						Write-Verbose "LOGIN: Cache confirmed for PID $($process.Id)" -ForegroundColor Green
					}
					
					if ($logStatus -contains 'TIMEOUT')
					{
						Write-Verbose 'LOGIN: Log monitoring timed out' -ForegroundColor Red
						break
					}
					
					if (Test-UserMouseIntervention)
					{
						Write-Verbose 'LOGIN: User intervention detected - stopping' -ForegroundColor Yellow
						if ($logMonitorJob)
						{
							Stop-Job -Job $logMonitorJob -ErrorAction SilentlyContinue
							Remove-Job -Job $logMonitorJob -Force -ErrorAction SilentlyContinue
						}
						$global:DashboardConfig.State.LoginActive = $false
						return
					}
					
				}
				
				if (-not $foundCacheJoin)
				{
					Write-Verbose 'LOGIN: Cache not detected within timeout period' -ForegroundColor Yellow
					continue
				}
				else
				{
					# Check if finalize collector login is enabled in settings
					$finalizeLogin = $false # Default to false if setting doesn't exist
					if ($global:DashboardConfig.Config.Contains('Login') -and 
						$global:DashboardConfig.Config['Login'].Contains('FinalizeCollectorLogin'))
					{
						$finalizeLogin = [bool]([int]$global:DashboardConfig.Config['Login']['FinalizeCollectorLogin'])
					}
					
					if ($finalizeLogin)
					{
						# Click again to finalize login
						$adjustedX = $centerX + 400
						$adjustedY = $centerY - 100
						
						Write-Verbose "LOGIN: Clicking to finalize collector login at X:$adjustedX Y:$adjustedY" -ForegroundColor Cyan
						Start-Sleep -Milliseconds 500
						Invoke-MouseClick -X $adjustedX -Y $adjustedY
					}
					else
					{
						Write-Verbose "LOGIN: Skipping finalize collector login (disabled in settings)" -ForegroundColor Yellow
					}

					Start-Sleep -Milliseconds 500
					
					Write-Verbose 'LOGIN: Minimizing...' -ForegroundColor Cyan
					[Custom.Native]::SendToBack($process.MainWindowHandle)
					Write-Verbose 'LOGIN: Optimizing...' -ForegroundColor Cyan
					[Custom.Native]::EmptyWorkingSet($process.Handle)

					Write-Verbose "LOGIN: Login complete for PID $($process.Id)" -ForegroundColor Green
				}
				
			}
			catch
			{
				Write-Verbose "LOGIN: Window setup error for PID $($process.Id): $_" -ForegroundColor Red
				continue
			}
		}
		catch
		{
			Write-Verbose "LOGIN: Login process error for row $entryNumber`: $_" -ForegroundColor Red
			$global:DashboardConfig.State.LoginActive = $false
		}
		finally
		{
			# Clean up resources
			if ($logMonitorJob)
			{
				Stop-Job -Job $logMonitorJob -ErrorAction SilentlyContinue
				Remove-Job -Job $logMonitorJob -Force -ErrorAction SilentlyContinue
			}
			$global:DashboardConfig.State.LoginActive = $false
		}
	}

	# Simple property change handler that updates button appearance
	if ($global:DashboardConfig.State.LoginActive -eq $true) {
		$global:DashboardConfig.UI.LoginButton.FlatStyle = 'Popup'
	} else {
		$global:DashboardConfig.UI.LoginButton.FlatStyle = 'Flat'
	}
	
	Write-Verbose 'LOGIN: All selected clients processed' -ForegroundColor Green
}

#endregion

#region Module Exports

# Export module functions
Export-ModuleMember -Function LoginSelectedRow

#endregion