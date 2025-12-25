<# reconnect.psm1 #>

#region Global Configuration & State Wrapper

if (-not $global:DashboardConfig)
{
	$global:DashboardConfig = @{
		Settings = @{ Paths = @{ GameLogFile = $null } }
		State    = @{}
		Config   = @{ Profiles = @{}; Options = @{ AutoReconnect = '0' } }
	}
}

if (-not $global:LoginCancellation)
{
	$global:LoginCancellation = [hashtable]::Synchronized(@{
			IsCancelled = $false
		})
}

$script:ReconnectScriptInitiatedMove = $false

if (-not $global:DashboardConfig.State) { $global:DashboardConfig.State = @{} }
if (-not $global:DashboardConfig.State.ReconnectQueue) { $global:DashboardConfig.State.ReconnectQueue = [System.Collections.Queue]::new() }
if (-not $global:DashboardConfig.State.PidCooldowns) { $global:DashboardConfig.State.PidCooldowns = @{} }
if (-not $global:DashboardConfig.State.ScheduledReconnects) { $global:DashboardConfig.State.ScheduledReconnects = @{} }
if (-not $global:DashboardConfig.State.ContainsKey('DisconnectActive')) { $global:DashboardConfig.State.DisconnectActive = $false }

if (-not $global:DashboardConfig.State.NotificationActionQueue)
{
	$global:DashboardConfig.State.NotificationActionQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
}

if (-not $global:DashboardConfig.State.ManualReconnectOverrides) { $global:DashboardConfig.State.ManualReconnectOverrides = [System.Collections.Generic.HashSet[int]]::new() }

if (-not $global:DashboardConfig.State.ContainsKey('QueuePaused')) { $global:DashboardConfig.State.QueuePaused = $false }

if (-not $global:DashboardConfig.State.WasInGamePids) { $global:DashboardConfig.State.WasInGamePids = [System.Collections.Generic.HashSet[int]]::new() }
if (-not $global:DashboardConfig.State.FlashingPids) { $global:DashboardConfig.State.FlashingPids = @{} }

if (-not $global:DashboardConfig.State.Timers) { $global:DashboardConfig.State.Timers = @{} }
if (-not $global:DashboardConfig.State.ActiveLoginPids) { $global:DashboardConfig.State.ActiveLoginPids = [System.Collections.Generic.HashSet[int]]::new() }

$global:WatcherBusy = $false
$global:WorkerBusy = $false

if (-not $global:DashboardConfig.State.NotificationMap) { $global:DashboardConfig.State.NotificationMap = @{} }
if (-not $global:NotificationStack) { $global:NotificationStack = [System.Collections.ArrayList]::new() }

#endregion

#region Internal Helper Functions

function InvokeGuardedAction
{
	param([ScriptBlock]$Action)
	$script:ReconnectScriptInitiatedMove = $true
	try { & $Action } finally
	{
		$script:ReconnectScriptInitiatedMove = $false
	}
}

function EnsureWindowReady
{
	param([IntPtr]$hWnd)
	if ($hWnd -eq [IntPtr]::Zero) { return }
	SetWindowToolStyle -hWnd $hWnd -Hide $false
	$stateChanged = $false
    
	if ([Custom.Native]::IsWindowMinimized($hWnd))
	{
		InvokeGuardedAction { [Custom.Native]::ShowWindow($hWnd, 9) }
		$stateChanged = $true
	}
    
	if ([Custom.Native]::GetForegroundWindow() -ne $hWnd)
	{
		InvokeGuardedAction { [Custom.Native]::SetForegroundWindow($hWnd) }
		InvokeGuardedAction { 
			if (-not [Custom.Native]::SetForegroundWindow($hWnd))
			{
				[Custom.Native]::keybd_event(0x12, 0, 0, [IntPtr]::Zero)
				Start-Sleep -Milliseconds 10
				[Custom.Native]::SetForegroundWindow($hWnd)
				[Custom.Native]::keybd_event(0x12, 0, 2, [IntPtr]::Zero)
			}
		}
		$stateChanged = $true
	}
    
	if ($stateChanged) { SleepWithCancel -Milliseconds 250 }

	if ([Custom.Native]::GetForegroundWindow() -ne $hWnd)
	{
		InvokeGuardedAction { [Custom.Native]::ShowWindow($hWnd, 9); [Custom.Native]::SetForegroundWindow($hWnd) }
	}

	SetWindowToolStyle -hWnd $hWnd -Hide $false
}

function SleepWithCancel
{
	param([int]$Milliseconds)
	$sw = [System.Diagnostics.Stopwatch]::StartNew()
	while ($sw.Elapsed.TotalMilliseconds -lt $Milliseconds)
	{
		if ($CancellationContext.IsCancelled) { throw 'LoginCancelled' }
		Start-Sleep -Milliseconds 100
	}
}

function EnsureWindowResponsive
{
	if ($Global:CurrentActiveProcessId -eq 0) { return }
	$proc = Get-Process -Id $Global:CurrentActiveProcessId -ErrorAction SilentlyContinue
	if (-not $proc) { throw "Process with ID $Global:CurrentActiveProcessId has terminated." }

	$sw = [System.Diagnostics.Stopwatch]::StartNew()
	while (-not $proc.Responding)
	{
		if ($sw.Elapsed.TotalSeconds -gt 15) { throw 'Timeout: Window is Not Responding (Hung).' }
		CheckCancel
		$proc.Refresh()
		Start-Sleep -Milliseconds 10
	}
	try { $proc.WaitForInputIdle(20) | Out-Null } catch {}
}

function Ensure-WindowState
{
	$hWnd = $Global:CurrentActiveWindowHandle
	if ($hWnd -eq [IntPtr]::Zero) { return }
                
	SetWindowToolStyle -hWnd $hWnd -Hide $false
	$changed = $false
	$CancellationContext.ScriptInitiatedMove = $true
	try
	{
		if ([Custom.Native]::IsWindowMinimized($hWnd))
		{
			[Custom.Native]::ShowWindow($hWnd, 9)
			$changed = $true
		}
                    
		$fg = [Custom.Native]::GetForegroundWindow()
		if ($fg -ne $hWnd)
		{
			[Custom.Native]::SetForegroundWindow($hWnd)
			if (-not [Custom.Native]::SetForegroundWindow($hWnd))
			{
				[Custom.Native]::keybd_event(0x12, 0, 0, [IntPtr]::Zero)
				Start-Sleep -Milliseconds 10
				[Custom.Native]::SetForegroundWindow($hWnd)
				[Custom.Native]::keybd_event(0x12, 0, 2, [IntPtr]::Zero)
			}
			$changed = $true
		}
	}
 finally
	{
		$CancellationContext.ScriptInitiatedMove = $false
	}
                
	if ($changed)
	{
		SleepWithCancel -Milliseconds 500
	}

	$fg = [Custom.Native]::GetForegroundWindow()
	if ($fg -ne $hWnd)
	{
		[Custom.Native]::ShowWindow($hWnd, 9)
		[Custom.Native]::SetForegroundWindow($hWnd)
	}
}

function Wait-ForDashboardClientReady
{
	param([int]$ProcessId)
	CheckCancel
	$timeout = New-TimeSpan -Seconds 120
	$sw = [System.Diagnostics.Stopwatch]::StartNew()
                
	while ($sw.Elapsed -lt $timeout)
	{
		CheckCancel
                    
		try
		{
			if ($DashboardConfig -and $DashboardConfig.UI -and $DashboardConfig.UI.DataGridFiller)
			{
				$grid = $DashboardConfig.UI.DataGridFiller
				$currentState = $null
                            
				foreach ($row in $grid.Rows)
				{
					if ($row.Tag -and $row.Tag.Id -eq $ProcessId)
					{
						$currentState = $row.Cells[3].Value
						break
					}
				}
                            
				if ($currentState -and $currentState -eq 'Loading...')
				{
					Start-Sleep -Milliseconds 250
					continue
				}
				elseif ($currentState -match 'Normal|Min|Reconnecting|Disconnected')
				{
					return $true
				}
			}
		}
		catch {}
                    
		Start-Sleep -Milliseconds 250
	}
                
	throw "Timeout waiting for client $ProcessId to leave Loading state"
}

function Invoke-MouseClick
{
	param([int]$X, [int]$Y)
	CheckCancel; EnsureWindowResponsive; CheckCancel
                
	Ensure-WindowState

	$CancellationContext.ScriptInitiatedMove = $true
	try
	{
		[Custom.Native]::SetCursorPos($X, $Y); Start-Sleep -Milliseconds 20
		[Custom.Native]::SetCursorPos($X, $Y); Start-Sleep -Milliseconds 30
		$MOUSEEVENTF_LEFTDOWN = 0x0002; $MOUSEEVENTF_LEFTUP = 0x0004
		[Custom.Native]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0); Start-Sleep -Milliseconds 50
		[Custom.Native]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0); Start-Sleep -Milliseconds 50
	}
 catch { Write-Verbose "Mouse click failed: $_" } finally
	{ 
		Start-Sleep -Milliseconds 50; $CancellationContext.ScriptInitiatedMove = $false; CheckCancel 
		Ensure-WindowState
	}
}

function Invoke-KeyPress
{
	param([int]$VirtualKeyCode)
	CheckCancel; EnsureWindowResponsive; CheckCancel
                
	Ensure-WindowState

	try
	{
		$hWnd = $Global:CurrentActiveWindowHandle
		if ($hWnd -eq [IntPtr]::Zero) { $hWnd = [Custom.Native]::GetForegroundWindow() }
		[Custom.Ftool]::fnPostMessage($hWnd, 0x0100, $VirtualKeyCode, 0); SleepWithCancel -Milliseconds 25
		[Custom.Ftool]::fnPostMessage($hWnd, 0x0101, $VirtualKeyCode, 0); SleepWithCancel -Milliseconds 25
	}
 catch { if ($_.Exception.Message -eq 'LoginCancelled') { throw }; throw "KeyPress Failed: $($_.Exception.Message)" }
                
	Ensure-WindowState
}

function ParseCoordinates
{
	param([string]$ConfigString)
	if ([string]::IsNullOrWhiteSpace($ConfigString) -or $ConfigString -notmatch ',') { return $null }
	$parts = $ConfigString.Split(',')
	if ($parts.Count -eq 2) { return @{ X = [int]$parts[0].Trim(); Y = [int]$parts[1].Trim() } }
	return $null
}

function Wait-ForLogEntry
{
	param($LogPath, $SearchStrings, $TimeoutSeconds)
	if ($TimeoutSeconds -le 0) { $TimeoutSeconds = 60 }
	$waitSw = [System.Diagnostics.Stopwatch]::StartNew()
	while ($waitSw.Elapsed.TotalSeconds -lt $TimeoutSeconds)
	{
		CheckCancel
		if (-not [string]::IsNullOrWhiteSpace($LogPath) -and (Test-Path $LogPath))
		{
			try
			{
				$lastLines = Get-Content $LogPath -Tail 20 -ErrorAction SilentlyContinue
				if ($lastLines) { foreach ($str in $SearchStrings) { if ($lastLines -match [regex]::Escape($str)) { return $true } } }
			}
			catch {}
		}
		SleepWithCancel -Milliseconds 250
	}
	return $false
}

function Wait-UntilWorldLoaded
{
	param($LogPath, $Config, $TotalSteps, $CurrentStep, $ClientIdx, $ClientCount, $EntryNum)
	$threshold = 3; $searchStr = '13 - CACHE_ACK_JOIN'
	if ($Config['WorldLoadLogThreshold']) { $threshold = [int]$Config['WorldLoadLogThreshold'] }
	if ($Config['WorldLoadLogEntry']) { $searchStr = $Config['WorldLoadLogEntry'] }
	$foundCount = 0; $timeout = New-TimeSpan -Minutes 2; $sw = [System.Diagnostics.Stopwatch]::StartNew()
	$pct = [int](($CurrentStep / $TotalSteps) * 100)
	$currentStep++; Report-LoginProgress -Action 'Waiting for World Load...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
	while ($foundCount -lt $threshold)
	{
		if ($sw.Elapsed -gt $timeout) { throw 'World load timeout' }
		CheckCancel
		if (-not [string]::IsNullOrWhiteSpace($LogPath) -and (Test-Path $LogPath))
		{
			try
			{
				$lines = Get-Content $LogPath -Tail 50 -ErrorAction SilentlyContinue
				if ($lines) { $foundCount = ($lines | Select-String -SimpleMatch $searchStr).Count }
			}
			catch {}
		}
		SleepWithCancel -Milliseconds 250
	}
}

function Write-LogWithRetry
{
	param([string]$FilePath, [string]$Value)
	if ([string]::IsNullOrWhiteSpace($FilePath)) { return }
	for ($i = 0; $i -lt 5; $i++) { try { Set-Content -Path $FilePath -Value $Value -Force -ErrorAction Stop; return } catch { Start-Sleep -Milliseconds 50 } }
}

function Report-LoginProgress
{
	param($Action, $Step, $TotalSteps, $ClientIdx, $ClientCount, $EntryNum, $ProfileName)
	$pct = [int](($Step / $TotalSteps) * 100)
	Write-Progress -Activity 'Login' -Status "Total: $pct% | Client $ClientIdx/$ClientCount`nProfile: $ProfileName | Account: ($EntryNum)`n$Action" -PercentComplete $pct
}

#endregion

#region Core Logic: Watcher & Worker

function StartDisconnectWatcher
{
	[CmdletBinding()]
	param()

	StopDisconnectWatcher

	$global:WatcherBusy = $false
	$global:WorkerBusy = $false
	$global:DashboardConfig.State.DisconnectActive = $true

	if ($global:DashboardConfig.State.WasInGamePids) { $global:DashboardConfig.State.WasInGamePids.Clear() }
	if ($global:DashboardConfig.State.FlashingPids) { $global:DashboardConfig.State.FlashingPids.Clear() }
	if ($global:DashboardConfig.State.ScheduledReconnects) { $global:DashboardConfig.State.ScheduledReconnects.Clear() }
	if ($global:DashboardConfig.State.ManualReconnectOverrides) { $global:DashboardConfig.State.ManualReconnectOverrides.Clear() }
	if ($global:DashboardConfig.State.NotificationActionQueue) { $global:DashboardConfig.State.NotificationActionQueue.Clear() }

	$global:DashboardConfig.State.QueuePaused = $false

	Write-Verbose 'Starting Disconnect Supervisor (Separated Logic Mode)...'

	$TimerWatcher = New-Object System.Windows.Forms.Timer
	$TimerWatcher.Interval = 2000

	$null = $JobWatcher; $JobWatcher = Register-ObjectEvent -InputObject $TimerWatcher -EventName Tick -SourceIdentifier 'DisconnectWatcherTick' -Action {
		if ($global:WatcherBusy) { return }
		$global:WatcherBusy = $true

		try
		{
			$now = Get-Date
			if ($global:DashboardConfig.State.ScheduledReconnects.Count -gt 0)
			{
				$scheduledPids = [int[]]@($global:DashboardConfig.State.ScheduledReconnects.Keys)

				$isPaused = ($global:DashboardConfig.State.LoginActive -or $global:DashboardConfig.State.ReconnectQueue.Count -gt 0 -or $global:DashboardConfig.State.NotificationHoverActive)

				foreach ($sPid in $scheduledPids)
				{
					$triggerTime = $global:DashboardConfig.State.ScheduledReconnects[$sPid]

					if ($isPaused)
					{
						$global:DashboardConfig.State.ScheduledReconnects[$sPid] = $triggerTime.AddMilliseconds(2000)
						continue
					}

					if ($now -ge $triggerTime)
					{
						$details = ''
						if ($global:DashboardConfig.UI.DataGridFiller)
						{
							$row = $global:DashboardConfig.UI.DataGridFiller.Rows | Where-Object { $_.Tag -is [System.Diagnostics.Process] -and $_.Tag.Id -eq $sPid } | Select-Object -First 1
							if ($row)
							{
								$proc = $row.Tag
								$profileName = [string]$row.Cells[1].Value
								$pName = 'Default'; $wTitle = $profileName
								if ($profileName -match '^\[(.*?)\](.*)') { $pName = $matches[1]; $wTitle = $matches[2] }
								$launcherPath = if ($global:DashboardConfig.Config['LauncherPath'] -and $global:DashboardConfig.Config['LauncherPath']['LauncherPath']) { $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] } else { '' }
									
								$isProfileAutoReconnect = $false
								$rp = $global:DashboardConfig.Config['ReconnectProfiles']
								if ($rp)
								{
									$pClean = $profileName
									if ($pClean -match '\[(.*?)\]') { $pClean = $matches[1] }
									if ($rp -is [System.Collections.Hashtable]) { if ($rp.ContainsKey($profileName) -or $rp.ContainsKey($pClean)) { $isProfileAutoReconnect = $true } } 
									elseif ($rp.Contains($profileName) -or $rp.Contains($pClean)) { $isProfileAutoReconnect = $true }
								}
								$autoRecFlag = if ($isProfileAutoReconnect) { 'Yes' } else { 'No' }
								$details = "Profile: $pName | Title: $wTitle`nPID: $sPid | Proc: $($proc.ProcessName)`nAuto-Reconnect: $autoRecFlag`nTime: $(Get-Date -Format 'HH:mm:ss')"
							}
						}
						if (-not $global:DashboardConfig.State.ReconnectQueue.Contains($sPid))
						{
							Write-Verbose "Auto-Reconnect Triggered for PID $sPid"
							$global:DashboardConfig.State.ReconnectQueue.Enqueue($sPid)
						}
						$global:DashboardConfig.State.ScheduledReconnects.Remove($sPid)
						CloseToast -Key $sPid
						ShowReconnectInteractiveNotification -Title 'Reconnecting...' -Message "Initiating reconnection sequence...`n$details" -Type 'Info' -RelatedPid $sPid -TimeoutSeconds 0

					}
				}
			}

			if (-not $global:DashboardConfig.UI.DataGridFiller) { return }

			$PidConnectionCounts = @{}
			$CheckFailed = $false
			try
			{
				
				
				$AllConns = @(Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue)
				if ($AllConns)
				{
					foreach ($c in $AllConns)
					{
						$p = [int]$c.OwningProcess
						if (-not $PidConnectionCounts.ContainsKey($p)) { $PidConnectionCounts[$p] = 0 }
						$PidConnectionCounts[$p]++
					}
				}
			}
			catch
			{
				
				Write-Verbose "RECONNECT: Unexpected error polling TCP connections: $($_.Exception.Message)"
				$CheckFailed = $true
			}

			if (-not $CheckFailed)
			{
				
				$WasInGame = $global:DashboardConfig.State.WasInGamePids
				$Flashing = $global:DashboardConfig.State.FlashingPids
				$Cooldowns = $global:DashboardConfig.State.PidCooldowns

				$globalAutoSetting = $false
				if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['AutoReconnect'])
				{
					if ("$($global:DashboardConfig.Config['Options']['AutoReconnect'])" -match '1|true') { $globalAutoSetting = $true }
				}

				foreach ($row in $global:DashboardConfig.UI.DataGridFiller.Rows)
				{
					if (-not ($row.Tag -is [System.Diagnostics.Process])) { continue }

					$proc = $row.Tag
					$pidInt = $proc.Id
					$profileName = [string]$row.Cells[1].Value

					if ($proc.HasExited)
					{
						if ($WasInGame.Contains($pidInt)) { $WasInGame.Remove($pidInt) }
						if ($Flashing.ContainsKey($pidInt)) { $Flashing.Remove($pidInt) }
						if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($pidInt))
						{
							$global:DashboardConfig.State.ScheduledReconnects.Remove($pidInt)
						}
						continue
					}

					$CurrentCount = if ($PidConnectionCounts.ContainsKey($pidInt)) { $PidConnectionCounts[$pidInt] } else { 0 }

					
					if (-not (Get-Variable -Name 'ReconnectConnectionThreshold' -Scope Script -ErrorAction SilentlyContinue)) { Set-Variable -Scope Script -Name ReconnectConnectionThreshold -Value 2 -Option ReadOnly }

					if ($CurrentCount -ge $script:ReconnectConnectionThreshold)
					{
						if (-not $WasInGame.Contains($pidInt)) { [void]$WasInGame.Add($pidInt); Write-Verbose "RECONNECT: PID $pidInt marked in-game (connections=$CurrentCount)" }
						if ($Flashing.ContainsKey($pidInt)) { $Flashing.Remove($pidInt) }
						if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($pidInt))
						{
							$global:DashboardConfig.State.ScheduledReconnects.Remove($pidInt)
						}
						CloseToast -Key $pidInt
					}

					if ($CurrentCount -eq 0 -and $WasInGame.Contains($pidInt))
					{

						if ($global:DashboardConfig.State.ActiveLoginPids.Contains($pidInt)) { continue }
						if ($global:DashboardConfig.State.ReconnectQueue.Contains($pidInt)) { continue }
						if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($pidInt)) { continue }

						$isFocused = $false
						try
						{
							$fg = [Custom.Native]::GetForegroundWindow()
							if ($proc.MainWindowHandle -ne [IntPtr]::Zero -and $proc.MainWindowHandle -eq $fg) { $isFocused = $true }
						}
						catch {}

						if ($isFocused)
						{
							$WasInGame.Remove($pidInt)
							CloseToast -Key $pidInt
							continue
						}

						$isCooldown = $false
						if ($Cooldowns.ContainsKey($pidInt))
						{
							if ((Get-Date) -lt $Cooldowns[$pidInt].AddSeconds(120)) { $isCooldown = $true }
						}

						$WasInGame.Remove($pidInt)
						$global:DashboardConfig.State.FlashingPids[$pidInt] = $true
						$row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black

						$pName = 'Default'; $wTitle = $profileName
						if ($profileName -match '^\[(.*?)\](.*)') { $pName = $matches[1]; $wTitle = $matches[2] }
						$launcherPath = if ($global:DashboardConfig.Config['LauncherPath'] -and $global:DashboardConfig.Config['LauncherPath']['LauncherPath']) { $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] } else { '' }

						$isProfileAutoReconnect = $false

						$rp = $global:DashboardConfig.Config['ReconnectProfiles']
						if ($rp)
						{
							$pClean = $profileName
							if ($pClean -match '\[(.*?)\]') { $pClean = $matches[1] }

							if ($rp -is [System.Collections.Hashtable])
							{
								if ($rp.ContainsKey($profileName) -or $rp.ContainsKey($pClean)) { $isProfileAutoReconnect = $true }
							}
							elseif ($rp.Contains($profileName) -or $rp.Contains($pClean))
							{
								$isProfileAutoReconnect = $true
							}
						}

						if (-not $rp -and $globalAutoSetting)
						{
							$isProfileAutoReconnect = $true
						}

						$autoRecFlag = if ($isProfileAutoReconnect) { 'Yes' } else { 'No' }
						$details = "Profile: $pName | Title: $wTitle`nPID: $pidInt | Proc: $($proc.ProcessName)`nAuto-Reconnect: $autoRecFlag`nTime: $(Get-Date -Format 'HH:mm:ss')"

						$btns = [ordered]@{
							'Reconnect Now'                = 'Reconnect'
							'Dismiss All'                  = 'DismissAll'
							'Dismiss'                      = 'Dismiss'
							'Delay Reconnect by 2 minutes' = 'Delay'
						}

						Write-Verbose "RECONNECT: PID $pidInt disconnected; isCooldown=$isCooldown; isProfileAutoReconnect=$isProfileAutoReconnect"
						if ($isProfileAutoReconnect)
						{
							if ($isCooldown)
							{
								$row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Orange
								ShowReconnectInteractiveNotification -Title 'Cooldown Active' -Message "Disconnected again. Auto-reconnect paused.`n$details" -Type 'Warning' -RelatedPid $pidInt -TimeoutSeconds 15 -Buttons $btns
							}
							else
							{
								$row.DefaultCellStyle.BackColor = [System.Drawing.Color]::DarkRed

								$global:DashboardConfig.State.ScheduledReconnects[$pidInt] = (Get-Date).AddSeconds(15)

								ShowReconnectInteractiveNotification -Title 'Connection Lost' -Message "Auto-reconnect in 15s...`n$details" -Type 'Warning' -RelatedPid $pidInt -TimeoutSeconds 15 -Buttons $btns
							}
						}
						else
						{
							$row.DefaultCellStyle.BackColor = [System.Drawing.Color]::DarkRed
							if ($isCooldown)
							{
								ShowReconnectInteractiveNotification -Title 'Cooldown Active' -Message "Disconnected again.`n$details" -Type 'Warning' -RelatedPid $pidInt -TimeoutSeconds 15 -Buttons $btns
							}
							else
							{
								ShowReconnectInteractiveNotification -Title 'Connection Lost' -Message "Disconnected.`n$details" -Type 'Warning' -RelatedPid $pidInt -TimeoutSeconds 15 -Buttons $btns
							}
						}
					}
				}
			}
		}
		catch {} finally { $global:WatcherBusy = $false }
	}

	$TimerWorker = New-Object System.Windows.Forms.Timer
	$TimerWorker.Interval = 500

	$null = $JobWorker; $JobWorker = Register-ObjectEvent -InputObject $TimerWorker -EventName Tick -SourceIdentifier 'DisconnectWorkerTick' -Action {
		if ($global:WorkerBusy) { return }
		$global:WorkerBusy = $true

		try
		{
			$cmdQueue = $global:DashboardConfig.State.NotificationActionQueue
			while ($cmdQueue -and $cmdQueue.Count -gt 0)
			{
				$cmd = $cmdQueue.Dequeue()
				$cPid = $cmd.Pid
				$action = $cmd.Action

				switch ($action)
				{
					'Reconnect'
					{
						if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($cPid))
						{
							$global:DashboardConfig.State.ScheduledReconnects.Remove($cPid)
						}
						if ($global:DashboardConfig.State.ManualReconnectOverrides) { [void]$global:DashboardConfig.State.ManualReconnectOverrides.Add($cPid) }

						if (-not $global:DashboardConfig.State.ReconnectQueue.Contains($cPid))
						{
							$global:DashboardConfig.State.ReconnectQueue.Enqueue($cPid)
						}

						$details = ''
						if ($global:DashboardConfig.UI.DataGridFiller)
						{
							$row = $global:DashboardConfig.UI.DataGridFiller.Rows | Where-Object { $_.Tag -is [System.Diagnostics.Process] -and $_.Tag.Id -eq $cPid } | Select-Object -First 1
							if ($row)
							{
								$proc = $row.Tag
								$profileName = [string]$row.Cells[1].Value
								$pName = 'Default'; $wTitle = $profileName
								if ($profileName -match '^\[(.*?)\](.*)') { $pName = $matches[1]; $wTitle = $matches[2] }
								$launcherPath = if ($global:DashboardConfig.Config['LauncherPath'] -and $global:DashboardConfig.Config['LauncherPath']['LauncherPath']) { $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] } else { '' }
                                
								$isProfileAutoReconnect = $false
								$rp = $global:DashboardConfig.Config['ReconnectProfiles']
								if ($rp)
								{
									$pClean = $profileName
									if ($pClean -match '\[(.*?)\]') { $pClean = $matches[1] }
									if ($rp -is [System.Collections.Hashtable]) { if ($rp.ContainsKey($profileName) -or $rp.ContainsKey($pClean)) { $isProfileAutoReconnect = $true } } 
									elseif ($rp.Contains($profileName) -or $rp.Contains($pClean)) { $isProfileAutoReconnect = $true }
								}
								$autoRecFlag = if ($isProfileAutoReconnect) { 'Yes' } else { 'No' }
								$details = "Profile: $pName | Title: $wTitle`nPID: $cPid | Proc: $($proc.ProcessName)`nAuto-Reconnect: $autoRecFlag`nTime: $(Get-Date -Format 'HH:mm:ss')"
							}
						}
						ShowReconnectInteractiveNotification -Title 'Reconnecting...' -Message "Initiating reconnection sequence...`n$details" -Type 'Info' -RelatedPid $cPid -TimeoutSeconds 0
					}

					'Dismiss'
					{
						if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($cPid))
						{
							$global:DashboardConfig.State.ScheduledReconnects.Remove($cPid)
						}
						CloseToast -Key $cPid
					}

					'Delay'
					{
						$global:DashboardConfig.State.ScheduledReconnects[$cPid] = (Get-Date).AddMinutes(2)
						if ($global:DashboardConfig.State.ManualReconnectOverrides) { [void]$global:DashboardConfig.State.ManualReconnectOverrides.Add($cPid) }
						Write-Verbose "PID $cPid Reconnect delayed 2m"
                        
						$details = ''
						if ($global:DashboardConfig.UI.DataGridFiller)
						{
							$row = $global:DashboardConfig.UI.DataGridFiller.Rows | Where-Object { $_.Tag -is [System.Diagnostics.Process] -and $_.Tag.Id -eq $cPid } | Select-Object -First 1
							if ($row)
							{
								$proc = $row.Tag
								$profileName = [string]$row.Cells[1].Value
								$pName = 'Default'; $wTitle = $profileName
								if ($profileName -match '^\[(.*?)\](.*)') { $pName = $matches[1]; $wTitle = $matches[2] }
								$launcherPath = if ($global:DashboardConfig.Config['LauncherPath'] -and $global:DashboardConfig.Config['LauncherPath']['LauncherPath']) { $global:DashboardConfig.Config['LauncherPath']['LauncherPath'] } else { '' }
                                
								$isProfileAutoReconnect = $false
								$rp = $global:DashboardConfig.Config['ReconnectProfiles']
								if ($rp)
								{
									$pClean = $profileName
									if ($pClean -match '\[(.*?)\]') { $pClean = $matches[1] }
									if ($rp -is [System.Collections.Hashtable]) { if ($rp.ContainsKey($profileName) -or $rp.ContainsKey($pClean)) { $isProfileAutoReconnect = $true } } 
									elseif ($rp.Contains($profileName) -or $rp.Contains($pClean)) { $isProfileAutoReconnect = $true }
								}
								$autoRecFlag = if ($isProfileAutoReconnect) { 'Yes' } else { 'No' }
								$details = "Profile: $pName | Title: $wTitle`nPID: $cPid | Proc: $($proc.ProcessName)`nAuto-Reconnect: $autoRecFlag`nTime: $(Get-Date -Format 'HH:mm:ss')"
							}
						}

						$btns = [ordered]@{
							'Reconnect Now' = 'Reconnect'
							'Dismiss All'   = 'DismissAll'
						}
						ShowReconnectInteractiveNotification -Title 'Reconnect Delayed' -Message "Auto-reconnect paused for 2 minutes.`n$details" -Type 'Info' -RelatedPid $cPid -TimeoutSeconds 120 -Buttons $btns
					}

					'DismissAll'
					{
						$targets = [int[]]@($global:DashboardConfig.State.FlashingPids.Keys)
						foreach ($tPid in $targets)
						{
							if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($tPid))
							{
								$global:DashboardConfig.State.ScheduledReconnects.Remove($tPid)
							}
							CloseToast -Key $tPid
						}
					}
				}
			}

			if ($global:DashboardConfig.State.LoginActive) { return }
			if ($global:DashboardConfig.State.QueuePaused) { return }
			if ($global:DashboardConfig.State.ReconnectQueue.Count -gt 0)
			{
				$PidToReconnect = $global:DashboardConfig.State.ReconnectQueue.Dequeue()
				InvokeReconnectionSequence -PidToReconnect $PidToReconnect
			}
		}
		catch { } finally { $global:WorkerBusy = $false }
	}

	$global:DashboardConfig.State.Timers['Watcher'] = $TimerWatcher
	$global:DashboardConfig.State.Timers['Worker'] = $TimerWorker
	$TimerWatcher.Start()
	$TimerWorker.Start()
}

function StopDisconnectWatcher
{
	[CmdletBinding()]
	param()
	Write-Verbose 'Stopping Disconnect Supervisor...'

	if ($global:DashboardConfig.State.Timers)
	{
		if ($global:DashboardConfig.State.Timers['Watcher'])
		{
			$t = $global:DashboardConfig.State.Timers['Watcher']
			try { $t.Stop(); $t.Dispose() } catch {}
			$global:DashboardConfig.State.Timers.Remove('Watcher')
		}
		if ($global:DashboardConfig.State.Timers['Worker'])
		{
			$t = $global:DashboardConfig.State.Timers['Worker']
			try { $t.Stop(); $t.Dispose() } catch {}
			$global:DashboardConfig.State.Timers.Remove('Worker')
		}
	}

	Get-EventSubscriber -SourceIdentifier 'DisconnectWatcherTick' -ErrorAction SilentlyContinue | Unregister-Event -Force
	Get-EventSubscriber -SourceIdentifier 'DisconnectWorkerTick' -ErrorAction SilentlyContinue | Unregister-Event -Force

	if ($global:DashboardConfig.State.NotificationMap)
	{
		$keysToClose = [int[]]@($global:DashboardConfig.State.NotificationMap.Keys)
		foreach ($key in $keysToClose)
		{
			CloseToast -Key $key
		}
	}
	if ($global:NotificationStack)
	{
		foreach ($f in $global:NotificationStack) { try { $f.Close() } catch {} }
		$global:NotificationStack.Clear()
	}

	if ($global:DashboardConfig.State)
	{
		$global:DashboardConfig.State.DisconnectActive = $false
		if ($global:DashboardConfig.State.ReconnectQueue) { $global:DashboardConfig.State.ReconnectQueue.Clear() }
		if ($global:DashboardConfig.State.WasInGamePids) { $global:DashboardConfig.State.WasInGamePids.Clear() }
		if ($global:DashboardConfig.State.FlashingPids) { $global:DashboardConfig.State.FlashingPids.Clear() }
		if ($global:DashboardConfig.State.ManualReconnectOverrides) { $global:DashboardConfig.State.ManualReconnectOverrides.Clear() }
		if ($global:DashboardConfig.State.NotificationActionQueue) { $global:DashboardConfig.State.NotificationActionQueue.Clear() }
	}

	$global:WatcherBusy = $false
	$global:WorkerBusy = $false
}

#endregion

#region Core Logic: Reconnection Sequence

function InvokeReconnectionSequence
{
	param([int]$PidToReconnect)

	$Row = $null
	if ($global:DashboardConfig.UI.DataGridFiller)
	{
		$Row = $global:DashboardConfig.UI.DataGridFiller.Rows | Where-Object { $_.Tag -is [System.Diagnostics.Process] -and $_.Tag.Id -eq $PidToReconnect } | Select-Object -First 1
	}

	if (-not $Row) { return }

	$CachedEntryNumber = [int]$Row.Cells[0].Value
	$CachedProfileName = [string]$Row.Cells[1].Value
	if ($CachedProfileName -match '\[(.*?)\]') { $CachedProfileName = $matches[1] }

	$global:LoginCancellation.IsCancelled = $false
	$Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::DarkRed
	$Row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black

	$hookCallback = [Custom.MouseHookManager+HookProc] {
		param($nCode, $wParam, $lParam)
		if ($nCode -ge 0 -and $wParam -eq 0x0200)
		{
			if (-not $script:ReconnectScriptInitiatedMove)
			{
				if (-not $global:LoginCancellation.IsCancelled)
				{
					$global:LoginCancellation.IsCancelled = $true
				}
			}
		}
		return [Custom.MouseHookManager]::CallNextHookEx([Custom.MouseHookManager]::HookId, $nCode, $wParam, $lParam)
	}

	$hWnd = [IntPtr]::Zero

	try
	{
		[Custom.MouseHookManager]::Start($hookCallback)

		$Process = $Row.Tag
		if ($Process.HasExited)
		{
			$Row.Cells[3].Value = 'Exited'
			return
		}

		$hWnd = $Process.MainWindowHandle

		Write-Verbose "Auto-Reconnect: Processing PID $PidToReconnect (Profile: $CachedProfileName)..."

		$isManuallyForced = $false
		if ($global:DashboardConfig.State.ManualReconnectOverrides -and $global:DashboardConfig.State.ManualReconnectOverrides.Contains($PidToReconnect))
		{
			$isManuallyForced = $true
			[void]$global:DashboardConfig.State.ManualReconnectOverrides.Remove($PidToReconnect)
		}

		if (-not $isManuallyForced)
		{
			if (-not ($global:DashboardConfig.Config['ReconnectProfiles'] -and $global:DashboardConfig.Config['ReconnectProfiles'].Contains($CachedProfileName)))
			{
				$Row.Cells[3].Value = 'Disconnected'
				ShowReconnectInteractiveNotification -Title 'Reconnect Blocked' -Message "Profile '$CachedProfileName' not enabled for auto-reconnect." -Type 'Info' -RelatedPid $PidToReconnect -TimeoutSeconds 10
				return
			}
		}

		$global:DashboardConfig.State.PidCooldowns[$PidToReconnect] = Get-Date

		$loginConfig = $global:DashboardConfig.Config['LoginConfig']
		$profileConfig = $null
		if ($loginConfig.Contains($CachedProfileName)) { $profileConfig = $loginConfig[$CachedProfileName] }
		elseif ($loginConfig.Contains('Default')) { $profileConfig = $loginConfig['Default'] }
		if (-not $profileConfig) { $profileConfig = @{} }

		$DisconnectCoordsString = if ($profileConfig['DisconnectOKCoords']) { $profileConfig['DisconnectOKCoords'] } else { '0,0' }
		$LoginDetailsOKString = if ($profileConfig['LoginDetailsOKCoords']) { $profileConfig['LoginDetailsOKCoords'] } else { '0,0' }
		$FirstNickString = if ($profileConfig['FirstNickCoords']) { $profileConfig['FirstNickCoords'] } else { '0,0' }
		$ScrollDownString = if ($profileConfig['ScrollDownCoords']) { $profileConfig['ScrollDownCoords'] } else { '0,0' }

		if ($DisconnectCoordsString -eq '0,0') { return }

		$ParseXY = { param($s) if ($s -match ',') { $p = $s.Split(','); return @{X = [int]$p[0]; Y = [int]$p[1]} } return $null }

		$disCoords = &$ParseXY $DisconnectCoordsString
		$logDetCoords = &$ParseXY $LoginDetailsOKString
		$firstNickCoords = &$ParseXY $FirstNickString
		$scrollDownCoords = &$ParseXY $ScrollDownString

		EnsureWindowReady $hWnd

		SleepWithCancel -Milliseconds 50
		$maxWait = 25000; $counter = 0
		while (-not $Process.Responding -and $counter -lt $maxWait)
		{
			SleepWithCancel -Milliseconds 500; $Process.Refresh(); $counter += 500
		}
		SleepWithCancel -Milliseconds 50

		EnsureWindowReady $hWnd

		$rect = New-Object Custom.Native+RECT
		EnsureWindowReady $hWnd
		if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect))
		{
			$absX = $rect.Left + $disCoords.X
			$absY = $rect.Top + $disCoords.Y
			InvokeGuardedAction {
				[Custom.Native]::SetCursorPos($absX, $absY)
				SleepWithCancel -Milliseconds 50
				[Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); SleepWithCancel -Milliseconds 50
				[Custom.Native]::mouse_event(0x04, 0, 0, 0, 0)
			}
		}
		EnsureWindowReady $hWnd

		SleepWithCancel -Milliseconds 50
		$maxWait = 25000; $counter = 0
		while (-not $Process.Responding -and $counter -lt $maxWait)
		{
			SleepWithCancel -Milliseconds 500; $Process.Refresh(); $counter += 500
		}
		SleepWithCancel -Milliseconds 50

		EnsureWindowReady $hWnd
		if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect))
		{
			$entryToClick = $CachedEntryNumber
			InvokeGuardedAction {
				[Custom.Native]::SetForegroundWindow($hWnd)
				Start-Sleep -Milliseconds 50
				if ($entryToClick -ge 6 -and $entryToClick -le 10)
				{
					$absX = $rect.Left + ($scrollDownCoords.X + 5)
					$absY = $rect.Top + ($scrollDownCoords.Y - 5)
					[Custom.Native]::SetCursorPos($absX, $absY)
					SleepWithCancel -Milliseconds 50
					[Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); SleepWithCancel -Milliseconds 50
					[Custom.Native]::mouse_event(0x04, 0, 0, 0, 0); SleepWithCancel -Milliseconds 50
					$absX = $rect.Left + $firstNickCoords.X
					$absY = $rect.Top + $firstNickCoords.Y
					$targetY = $absY + (($entryToClick - 6) * 18)
				}
				elseif ($entryToClick -ge 1 -and $entryToClick -le 5)
				{
					$absX = $rect.Left + $firstNickCoords.X
					$absY = $rect.Top + $firstNickCoords.Y
					$targetY = $absY + (($entryToClick - 1) * 18)
				}
				[Custom.Native]::SetCursorPos($absX, $targetY)
			}
			SleepWithCancel -Milliseconds 50
			InvokeGuardedAction { [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); SleepWithCancel -Milliseconds 50; [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0) }
			SleepWithCancel -Milliseconds 50
			InvokeGuardedAction { [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); SleepWithCancel -Milliseconds 50; [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0) }
		}
		EnsureWindowReady $hWnd

		SleepWithCancel -Milliseconds 1000
		$maxWait = 25000; $counter = 0
		while (-not $Process.Responding -and $counter -lt $maxWait)
		{
			SleepWithCancel -Milliseconds 2000; $Process.Refresh(); $counter += 2000
		}
		SleepWithCancel -Milliseconds 1000

		EnsureWindowReady $hWnd
		if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect))
		{
			$absX = $rect.Left + $logDetCoords.X; $absY = $rect.Top + $logDetCoords.Y
			InvokeGuardedAction {
				[Custom.Native]::SetCursorPos($absX, $absY)
				SleepWithCancel -Milliseconds 50
				[Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); SleepWithCancel -Milliseconds 50
				[Custom.Native]::mouse_event(0x04, 0, 0, 0, 0)
			}
			SleepWithCancel -Milliseconds 50
		}

		EnsureWindowReady $hWnd

		SleepWithCancel -Milliseconds 50
		$maxWait = 25000; $counter = 0
		while (-not $Process.Responding -and $counter -lt $maxWait)
		{
			SleepWithCancel -Milliseconds 500; $Process.Refresh(); $counter += 500
		}
		SleepWithCancel -Milliseconds 50

		EnsureWindowReady $hWnd
		if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect))
		{
			InvokeGuardedAction {
				SleepWithCancel -Milliseconds 50
				if ($entryToClick -ge 6 -and $entryToClick -le 10)
				{
					$absX = $rect.Left + $firstNickCoords.X
					$absY = $rect.Top + $firstNickCoords.Y
					$targetY = $absY + (($entryToClick - 6) * 18)
				}
				elseif ($entryToClick -ge 1 -and $entryToClick -le 5)
				{
					$absX = $rect.Left + $firstNickCoords.X
					$absY = $rect.Top + $firstNickCoords.Y
					$targetY = $absY + (($entryToClick - 1) * 18)
				}
				[Custom.Native]::SetCursorPos($absX, $targetY)
			}
			SleepWithCancel -Milliseconds 50
			InvokeGuardedAction { [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); SleepWithCancel -Milliseconds 50; [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0) }
			SleepWithCancel -Milliseconds 50
			InvokeGuardedAction { [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); SleepWithCancel -Milliseconds 50; [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0) }
		}
		EnsureWindowReady $hWnd

		SleepWithCancel -Milliseconds 50
		[Custom.MouseHookManager]::Stop()

		CloseToast -Key $PidToReconnect

		
		
		
		try
		{
			try { $logPath = GetClientLogPath -Row $Row } catch { $logPath = $null }

			$mainForm = if ($global:DashboardConfig.UI) { $global:DashboardConfig.UI.MainForm } else { $null }
			$rowRef = $Row; $hWndRef = $hWnd; $logRef = $logPath

			if ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated)
			{
				Write-Verbose "RECONNECT: Scheduling UI-threaded LoginSelectedRow for PID $PidToReconnect (Entry=$CachedEntryNumber)"
				$mainForm.Invoke([Action]({
							try
							{
								
								if (-not $rowRef)
								{
									Write-Verbose 'RECONNECT: UI Thread Closure failed to capture row. Aborting UI selection.'
									return 
								}

								if ($global:DashboardConfig.UI -and $global:DashboardConfig.UI.DataGridFiller)
								{
									try { $global:DashboardConfig.UI.DataGridFiller.ClearSelection() } catch {}
									try { $rowRef.Selected = $true } catch {}
									try { $global:DashboardConfig.UI.DataGridFiller.CurrentCell = $rowRef.Cells[0] } catch {}
								}
                        
								
								$safeHwnd = [IntPtr]::Zero
								if ($hWndRef) { $safeHwnd = $hWndRef }

								Write-Verbose "RECONNECT (UI): Invoking LoginSelectedRow for PID $($rowRef.Tag.Id) (Entry=$CachedEntryNumber)"
								if (Get-Command 'LoginSelectedRow' -ErrorAction SilentlyContinue)
								{ 
									LoginSelectedRow -RowInput $rowRef -WindowHandle $safeHwnd -LogFilePath $logRef 
								}
							}
							catch { Write-Verbose "RECONNECT: UI invocation error: $_" }
						}.GetNewClosure())) | Out-Null
			}
			else
			{
				Write-Verbose "RECONNECT: UI not available, invoking LoginSelectedRow directly for PID $PidToReconnect (Entry=$CachedEntryNumber)"
				if (Get-Command 'LoginSelectedRow' -ErrorAction SilentlyContinue) { LoginSelectedRow -RowInput $Row -WindowHandle $hWnd -LogFilePath $logPath }
			}
		}
		catch { Write-Verbose "RECONNECT: Error while scheduling LoginSelectedRow: $_" }

	}
 catch
	{
		if ($hWnd -ne [IntPtr]::Zero)
		{
			try { [Custom.Native]::ShowWindow($hWnd, 6) } catch {}
		}
		if ($_.Exception.Message -eq 'LoginCancelled' -or $_.ToString() -eq 'LoginCancelled')
		{
			$Row.Cells[3].Value = 'Cancelled'
		}
		else
		{
			$Row.Cells[3].Value = 'Error'
			Write-Verbose "Reconnect Error: $_"
		}
	}
 finally
	{
		try { [Custom.MouseHookManager]::Stop() } catch {}
		$global:LoginCancellation.IsCancelled = $false
		$script:ReconnectScriptInitiatedMove = $false
	}
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion
