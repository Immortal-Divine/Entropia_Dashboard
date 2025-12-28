<# login.psm1 #>

#region Configuration

$global:LoginCancellation = [hashtable]::Synchronized(@{
		IsCancelled         = $false
		ScriptInitiatedMove = $false
	})
$TOTAL_STEPS_PER_CLIENT = 15
$script:LastToastUpdateTime = $null

if (-not $global:LoginResources)
{
	$global:LoginResources = @{
		PowerShellInstance  = $null
		Runspace            = $null
		EventSubscriptionId = $null
		EventSubscriber     = $null
		InfoSubscription    = $null
		AsyncResult         = $null
		IsStopping          = $false
		IsMouseHookActive   = $false
	}
}

if (-not $global:DashboardConfig) { $global:DashboardConfig = @{} }
if (-not $global:DashboardConfig.State) { $global:DashboardConfig.State = @{} }
if ($global:DashboardConfig.State -is [System.Collections.IDictionary])
{
	if (-not $global:DashboardConfig.State.Contains('LoginNotificationMap')) { $global:DashboardConfig.State['LoginNotificationMap'] = @{} }
}
else
{
	if (-not $global:DashboardConfig.State.LoginNotificationMap) { $global:DashboardConfig.State.LoginNotificationMap = @{} }
}
if (-not $global:LoginNotificationStack) { $global:LoginNotificationStack = [System.Collections.ArrayList]::new() }


#endregion

#region Helper Functions

function GetClientLogPath
{
	param(
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.DataGridViewRow]$Row
	)
	$determinedLogBaseFolder = ''
	$process = $Row.Tag
	$profileName = ''
	$processExeBaseFolder = ''
	$processTitle = $Row.Cells[1].Value
	$null = $entryNum; $entryNum = $Row.Cells[0].Value

	if ($processTitle -match '^\[([^\]]+)\]') { $profileName = $Matches[1] }

	if ([string]::IsNullOrEmpty($profileName) -and $process -and $process.Id)
	{
		$profileName = GetProcessProfile -Process $process
	}

	if ($process -and $process.Id)
	{
		$processExePath = [Custom.Native]::GetProcessPathById($process.Id)
		if (-not [string]::IsNullOrEmpty($processExePath))
		{
			$processExeBaseFolder = Split-Path -Parent -Path $processExePath
		}
	}

	if ([string]::IsNullOrEmpty($determinedLogBaseFolder) -and
		$global:DashboardConfig -and $global:DashboardConfig.Config -and
		$global:DashboardConfig.Config.Contains('Profiles'))
	{
		$foundProfileKey = $null
		foreach ($key in $global:DashboardConfig.Config.Profiles.Keys)
		{
			if ($key.ToLowerInvariant() -eq $profileName.ToLowerInvariant())
			{
				$foundProfileKey = $key
				break
			}
		}
		if ($foundProfileKey)
		{
			$profilePath = $global:DashboardConfig.Config.Profiles[$foundProfileKey]
			if (-not [string]::IsNullOrEmpty($profilePath))
			{
				$determinedLogBaseFolder = $profilePath
			}
		}
	}

	if ([string]::IsNullOrEmpty($determinedLogBaseFolder) -and -not [string]::IsNullOrEmpty($processExeBaseFolder))
	{
		$determinedLogBaseFolder = $processExeBaseFolder
	}

	if ([string]::IsNullOrEmpty($determinedLogBaseFolder))
	{
		$launcherPathConfig = $global:DashboardConfig.Config['LauncherPath']
		if ($launcherPathConfig -and $launcherPathConfig.Contains('LauncherPath'))
		{
			$launcherPath = $launcherPathConfig['LauncherPath']
			if (-not [string]::IsNullOrEmpty($launcherPath))
			{
				$determinedLogBaseFolder = Split-Path -Path $launcherPath -Parent
			}
		}
	}

	if ([string]::IsNullOrEmpty($determinedLogBaseFolder))
	{
		$determinedLogBaseFolder = Join-Path -Path $env:APPDATA -ChildPath 'Entropia_Dashboard'
	}

	$actualLogPath = Join-Path -Path $determinedLogBaseFolder -ChildPath "Log\network_$(Get-Date -Format 'yyyyMMdd').log"
	return $actualLogPath
}

#endregion

#region Core Function

function LoginSelectedRow
{
	param(
		[Parameter(Mandatory = $false)]
		$RowInput,
		[Parameter(Mandatory = $false)]
		[Object]$WindowHandle = [IntPtr]::Zero,
		[string]$LogFilePath
	)
    
	
	if ($WindowHandle -isnot [IntPtr])
	{
		if ($WindowHandle)
		{
			try { $WindowHandle = [IntPtr]$WindowHandle } catch { $WindowHandle = [IntPtr]::Zero }
		}
		else
		{
			$WindowHandle = [IntPtr]::Zero
		}
	}

	if ($global:DashboardConfig.State['LoginActive']) { return }
	$global:DashboardConfig.State['LoginActive'] = $true

	try 
	{
		$null = $rowsToProcess; $rowsToProcess = @()
		$rawSelection = @()

		if ($RowInput)
		{
			$rawSelection += $RowInput
		}
		elseif ($global:DashboardConfig.UI.DataGridFiller.SelectedRows.Count -gt 0)
		{
			$rawSelection = $global:DashboardConfig.UI.DataGridFiller.SelectedRows | Sort-Object Index
		}
		else
		{
			$global:DashboardConfig.State['LoginActive'] = $false
			return
		}

		$enrichedData = $rawSelection | ForEach-Object {
			$item = $_
			$actualRow = $null
			$entryNum = 0

			
			if ($item.PSTypeNames -contains 'LoginOverrideWrapper')
			{
				$actualRow = $item.Row
				$entryNum = $item.OverrideAccountID -as [int]
			} 
			else
			{
				$actualRow = $item
				$entryNum = $actualRow.Cells[0].Value -as [int]
			}

			if ($null -ne $actualRow)
			{
				$title = $actualRow.Cells[1].Value
				$profileName = 'Default'
				if ($title -match '^\[([^\]]+)\]') { $profileName = $Matches[1] }
                
				[PSCustomObject]@{ 
					OriginalRow = $actualRow
					EntryNum    = $entryNum 
					Profile     = $profileName 
				}
			}
		}
		
		$profilePriorityMap = [ordered]@{}
		$priorityCounter = 0
		foreach ($item in $enrichedData)
		{
			$pKey = $item.Profile.ToString()
			if (-not $profilePriorityMap.Contains($pKey)) { $profilePriorityMap[$pKey] = $priorityCounter++ }
		}
		
		$sortedData = $enrichedData | Sort-Object @{Expression = {$profilePriorityMap[$_.Profile.ToString()]}; Ascending = $true}, @{Expression = {$_.EntryNum}; Ascending = $true}

		$jobs = @()
		foreach ($dataItem in $sortedData)
		{
			$row = $dataItem.OriginalRow
			$entryNum = [int]$dataItem.EntryNum 
            
			$process = $row.Tag
			$processTitle = $row.Cells[1].Value

			$profileName = 'Default'
			if ($global:DashboardConfig.Config['LoginConfig'])
			{
				$allProfileNames = $global:DashboardConfig.Config['LoginConfig'].Keys | Where-Object { $_ -ne 'Default' }
				foreach ($knownProfile in $allProfileNames)
				{
					if ($processTitle -match "\[$([regex]::Escape($knownProfile))\]")
					{
						$profileName = $knownProfile
						break
					}
				}
			}

			$actualLogPath = $LogFilePath
			if ([string]::IsNullOrEmpty($actualLogPath))
			{
				$actualLogPath = GetClientLogPath -Row $row
			}
			if ($null -eq $actualLogPath) { $actualLogPath = '' }

			$clientLoginConfig = @{}
			if ($global:DashboardConfig.Config['LoginConfig'] -and $global:DashboardConfig.Config['LoginConfig'][$profileName])
			{
				$clientLoginConfig = $global:DashboardConfig.Config['LoginConfig'][$profileName]
			}

			$thisJobHandle = [IntPtr]::Zero
			if ($WindowHandle -ne [IntPtr]::Zero)
			{
				if ($RowInput -eq $row -or ($RowInput.Row -eq $row))
				{
					$thisJobHandle = $WindowHandle
				}
			}

			$jobs += [PSCustomObject]@{
				EntryNumber    = $entryNum
				ProcessId      = if ($process) { $process.Id } else { 0 }
				ExplicitHandle = $thisJobHandle
				LogPath        = $actualLogPath
				Config         = $clientLoginConfig
				ProfileName    = $profileName
			}
		}
		
		ShowToast -Title 'Login Process' -Message 'Starting Login Process...' -Type 'Info' -Key 9999 -TimeoutSeconds 0

		$global:DashboardConfig.UI.LoginButton.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
		$global:DashboardConfig.UI.LoginButton.Text = 'Running...'

		$global:LoginCancellation.IsCancelled = $false
		$global:LoginCancellation.ScriptInitiatedMove = $false
		$global:LoginResources['IsStopping'] = $false

		$hookCallback = [Custom.MouseHookManager+HookProc] {
			param($nCode, $wParam, $lParam)
			if ($nCode -ge 0 -and $wParam -eq 0x0200)
			{
				if (-not $global:LoginCancellation.ScriptInitiatedMove)
				{
					
					if ($global:LoginResources -and $global:LoginResources.PowerShellInstance -and -not $global:LoginResources['IsStopping'])
					{
						$psInstance = $global:LoginResources.PowerShellInstance
						if ($psInstance.InvocationStateInfo.State -eq 'Running')
						{
							Write-Verbose 'LOGIN: Mouse movement detected, stopping login pipeline.'
							$global:LoginCancellation.IsCancelled = $true
							
							$psInstance.Stop()
						}
					}
				}
			}
			return [Custom.MouseHookManager]::CallNextHookEx([Custom.MouseHookManager]::HookId, $nCode, $wParam, $lParam)
		}
		if (-not $global:LoginResources.IsMouseHookActive)
		{
			try
			{
				[Custom.MouseHookManager]::Start($hookCallback)
				$global:LoginResources.IsMouseHookActive = $true
				Write-Verbose 'LOGIN: Mouse hook started.'
			}
			catch
			{
				Write-Warning "LOGIN: Failed to start mouse hook: $_"
			}
		}

		
		$varsToInject = @{
			'DashboardConfig'        = $global:DashboardConfig
			'TOTAL_STEPS_PER_CLIENT' = $TOTAL_STEPS_PER_CLIENT
			'CancellationContext'    = $global:LoginCancellation
			'LoginResources'         = $global:LoginResources
		}

		$assembliesToAdd = [System.Collections.Generic.HashSet[string]]::new()
		$customTypesToLoad = @('Custom.Native', 'Custom.Ftool', 'Custom.MouseHookManager', 'Custom.ColorWriter')
        
		foreach ($typeName in $customTypesToLoad)
		{
			$type = $typeName -as [Type]
			if ($type -and -not [string]::IsNullOrEmpty($type.Assembly.Location))
			{
				$assembliesToAdd.Add($type.Assembly.Location) | Out-Null
			}
		}
        
		
		if ($assembliesToAdd.Count -eq 0)
		{
			$loadedAssemblies = [AppDomain]::CurrentDomain.GetAssemblies()
			foreach ($asm in $loadedAssemblies)
			{
				try
				{
					if ($asm.IsDynamic) { continue }
					if ([string]::IsNullOrWhiteSpace($asm.Location)) { continue }

					foreach ($t in $asm.GetExportedTypes())
					{
						if ($customTypesToLoad -contains $t.FullName)
						{
							$assembliesToAdd.Add($asm.Location) | Out-Null
							break
						}
					}
				}
				catch {}
			}
		}

		
		$localRunspace = NewManagedRunspace `
			-Name 'LoginRunspace' `
			-MinRunspaces 1 `
			-MaxRunspaces 1 `
			-SessionVariables $varsToInject `
			-Assemblies ($assembliesToAdd | Select-Object -Unique)

		
		$global:LoginResources = @{
			PowerShellInstance   = $null
			Runspace             = $localRunspace
			EventSubscriptionId  = $null
			EventSubscriber      = $null
			ProgressSubscription = $null
			AsyncResult          = $null
			IsStopping           = $false
			IsMouseHookActive    = $global:LoginResources.IsMouseHookActive
		}

		
		$loginScript = {
			param($Jobs, $TotalStepsPerClient, $GlobalOptions, $LoginConfig, $CancellationContext, $State)
            
			if (-not $TotalStepsPerClient -or $TotalStepsPerClient -eq 0) { $TotalStepsPerClient = 15 }

			if (-not ('Custom.Native' -as [Type]))
			{
				throw 'CRITICAL ERROR: [Custom.Native] type is missing in Background Runspace. Automation cannot proceed.'
			}
            
			$Global:CurrentActiveProcessId = 0
			$Global:CurrentActiveWindowHandle = [IntPtr]::Zero
			$InformationPreference = 'SilentlyContinue'

			function CheckCancel
			{
				if ($CancellationContext.IsCancelled) { throw 'LoginCancelled' }
			}

			function SleepWithCancel
			{
				param([int]$Milliseconds)
				$sw = [System.Diagnostics.Stopwatch]::StartNew()
				while ($sw.Elapsed.TotalMilliseconds -lt $Milliseconds)
				{
					if ($CancellationContext.IsCancelled) { throw 'LoginCancelled' }
					Start-Sleep -Milliseconds 50
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

			function SetWindowToolStyle
			{
				param([IntPtr]$hWnd, [bool]$Hide = $true)
				$GWL_EXSTYLE = -20; $WS_EX_TOOLWINDOW = 0x00000080; $WS_EX_APPWINDOW = 0x00040000
				try
				{
					if ('Custom.TaskbarTool' -as [Type]) { $null = [Custom.TaskbarTool]::SetTaskbarState($hWnd, -not $Hide) }
					[IntPtr]$currentExStylePtr = [Custom.Native]::GetWindowLongPtr($hWnd, $GWL_EXSTYLE)
					[long]$currentExStyle = $currentExStylePtr.ToInt64(); $newExStyle = $currentExStyle
					if ($Hide) { $newExStyle = ($currentExStyle -bor $WS_EX_TOOLWINDOW) -band (-bnot $WS_EX_APPWINDOW) }
					else { $newExStyle = ($currentExStyle -band (-bnot $WS_EX_TOOLWINDOW)) -bor $WS_EX_APPWINDOW }
					if ($newExStyle -ne $currentExStyle)
					{
						$isMinimized = [Custom.Native]::IsMinimized($hWnd)
						[Custom.Native]::ShowWindow($hWnd, [Custom.Native]::SW_HIDE)
						[IntPtr]$newExStylePtr = [IntPtr]$newExStyle
						[Custom.Native]::SetWindowLongPtr($hWnd, $GWL_EXSTYLE, $newExStylePtr) | Out-Null
						if ($isMinimized) { [Custom.Native]::ShowWindow($hWnd, [Custom.Native]::SW_SHOWMINNOACTIVE) }
						else { [Custom.Native]::ShowWindow($hWnd, [Custom.Native]::SW_SHOWNA) }
					}
				}
				catch {}
			}

			function EnsureWindowState
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
						[Custom.Native]::ShowWindow($hWnd, 6)
						[Custom.Native]::ShowWindow($hWnd, 9)
						$changed = $true
					}
                    
					$fg = [Custom.Native]::GetForegroundWindow()
					if ($fg -ne $hWnd)
					{
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

			function Invoke-MouseClick
			{
				param([int]$X, [int]$Y)
				CheckCancel; EnsureWindowResponsive; CheckCancel
                
				EnsureWindowState

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
					EnsureWindowState
				}
			}

			function Invoke-KeyPress
			{
				param([int]$VirtualKeyCode)
				CheckCancel; EnsureWindowResponsive; CheckCancel
                
				EnsureWindowState

				try
				{
					$hWnd = $Global:CurrentActiveWindowHandle
					if ($hWnd -eq [IntPtr]::Zero) { $hWnd = [Custom.Native]::GetForegroundWindow() }
					[Custom.Ftool]::fnPostMessage($hWnd, 0x0100, $VirtualKeyCode, 0); SleepWithCancel -Milliseconds 25
					[Custom.Ftool]::fnPostMessage($hWnd, 0x0101, $VirtualKeyCode, 0); SleepWithCancel -Milliseconds 25
				}
				catch { if ($_.Exception.Message -eq 'LoginCancelled') { throw }; throw "KeyPress Failed: $($_.Exception.Message)" }
                
				EnsureWindowState
			}

			function Set-WindowForeground
			{
				param([System.Diagnostics.Process]$Process, [IntPtr]$ExplicitHwnd)
				CheckCancel
				$targetHwnd = if ($ExplicitHwnd -ne [IntPtr]::Zero) { $ExplicitHwnd } else { $Process.MainWindowHandle }
				if ($targetHwnd -eq [IntPtr]::Zero) { return $false }
				$CancellationContext.ScriptInitiatedMove = $true
				try { [Custom.Native]::BringToFront($targetHwnd); Start-Sleep -Milliseconds 100 } finally { Start-Sleep -Milliseconds 50; $CancellationContext.ScriptInitiatedMove = $false; CheckCancel }
				return $true
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
				param($LogPath, $Config, $TotalSteps, $CurrentStep, $ClientIdx, $ClientCount, $EntryNum, $ProfileName)
				$threshold = 3; $searchStr = '13 - CACHE_ACK_JOIN'
				if ($Config['WorldLoadLogThreshold']) { $threshold = [int]$Config['WorldLoadLogThreshold'] }
				if ($Config['WorldLoadLogEntry']) { $searchStr = $Config['WorldLoadLogEntry'] }
				$foundCount = 0; $timeout = New-TimeSpan -Minutes 2; $sw = [System.Diagnostics.Stopwatch]::StartNew()
				$null = $pct; $pct = 0; if ($TotalSteps -gt 0) { $pct = [int](($CurrentStep / $TotalSteps) * 100) }
				$currentStep++; ReportLoginProgress -Action 'Waiting for World Load...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
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

			function Write-Verbose
			{
				[CmdletBinding()]
				param([string]$Message, [string]$ForegroundColor = 'DarkGray')
                
				$dateStr = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
				$callerName = 'LoginSelectedRow'
				$bracketedCaller = "[$callerName]"
				$paddedCaller = $bracketedCaller.PadRight(35)
				$prefix = " | $dateStr - $paddedCaller - "
				$finalMsg = "$prefix$Message"

				if ('Custom.ColorWriter' -as [Type])
				{
					[Custom.ColorWriter]::WriteColored($finalMsg, $ForegroundColor)
				}
				else
				{
					Write-Host $finalMsg -ForegroundColor $ForegroundColor
				}
			}

			function ReportLoginProgress
			{
				param($Action, $Step, $TotalSteps, $ClientIdx, $ClientCount, $EntryNum, $ProfileName)
				$pct = 0; if ($TotalSteps -gt 0) { $pct = [int](($Step / $TotalSteps) * 100) }
				$msg = "Total: $pct% | Client $ClientIdx/$ClientCount`nProfile: $ProfileName | Account: ($EntryNum)`n$Action"
				Write-Verbose -Message $msg
				Write-Information -MessageData @{ Text = $msg; Percent = $pct } -Tags 'LoginStatus'
			}

			try
			{
				$totalClients = $Jobs.Count
				$totalGlobalSteps = $totalClients * $TotalStepsPerClient
				$currentClientIndex = 0

				if ($State -and -not $State['LoginGracePids']) { $State['LoginGracePids'] = [System.Collections.Hashtable]::Synchronized(@{}) }

				foreach ($job in $Jobs)
				{
					CheckCancel; EnsureWindowResponsive; CheckCancel; SleepWithCancel -Milliseconds 50
					$currentClientIndex++
					$entryNumber = $job.EntryNumber
					$processId = $job.ProcessId
					$explicitHandle = if ($job.ExplicitHandle -and $job.ExplicitHandle -ne 0) { $job.ExplicitHandle } else { [IntPtr]::Zero }
					$Global:CurrentActiveProcessId = $processId
					$profileName = $job.ProfileName

					if ($processId -gt 0) { if ($State -and $State['ActiveLoginPids']) { [void]$State['ActiveLoginPids'].Add($processId) } }
                
					$logPath = $job.LogPath
					$config = $job.Config
					$stepBase = ($currentClientIndex - 1) * $TotalStepsPerClient
					$currentStep = $stepBase

					$currentStep++; ReportLoginProgress -Action 'Starting Login Process...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
                
					if ($processId -eq 0) { throw "Process ID not found for Client $entryNumber" }
					$process = Get-Process -Id $processId -ErrorAction SilentlyContinue
					if (-not $process) { throw "Process $processId is gone." }

					$serverID = '1'; $channelID = '1'; $charSlot = '1'; $startCollector = 'No'
					$settingKey = "Client${entryNumber}_Settings"
					if ($config[$settingKey])
					{
						$parts = $config[$settingKey] -split ','
						if ($parts.Count -eq 4) { $serverID = $parts[0]; $channelID = $parts[1]; $charSlot = $parts[2]; $startCollector = $parts[3] }
					}

					Write-LogWithRetry -FilePath $logPath -Value ''
					$workingHwnd = if ($explicitHandle -ne [IntPtr]::Zero) { $explicitHandle } else { $process.MainWindowHandle }
					$Global:CurrentActiveWindowHandle = $workingHwnd
					$currentStep++; ReportLoginProgress -Action 'Getting Client ready...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName

					EnsureWindowState; CheckCancel; EnsureWindowResponsive; CheckCancel; SleepWithCancel -Milliseconds 50

					$rect = New-Object Custom.Native+RECT
					[Custom.Native]::GetWindowRect($workingHwnd, [ref]$rect)
					$x = $rect.Left + 100; $y = $rect.Top + 100
                
					$currentStep++; ReportLoginProgress -Action "Log into Account $($entryNumber)..." -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName

					$firstNickCoords = ParseCoordinates $config['FirstNickCoords']
					$scrollDownCoords = ParseCoordinates $config['ScrollDownCoords']
					$scrollbaseX = if ($scrollDownCoords) { $rect.Left + $scrollDownCoords.X } else { [int](($rect.Left + $rect.Right) / 2) + 160 }
					$scrollbaseY = if ($scrollDownCoords) { $rect.Top + $scrollDownCoords.Y } else { [int](($rect.Top + $rect.Bottom) / 2) + 46 }

					if ($firstNickCoords)
					{
						$baseX = $rect.Left + $firstNickCoords.X
						$baseY = $rect.Top + $firstNickCoords.Y
						if ($entryNumber -ge 6 -and $entryNumber -le 10)
						{
							CheckCancel; EnsureWindowResponsive; CheckCancel
							Invoke-MouseClick -X ($scrollbaseX + 5) -Y ($scrollbaseY - 5); SleepWithCancel -Milliseconds 50
							$targetY = $baseY + (($entryNumber - 6) * 18)
							Invoke-MouseClick -X $baseX -Y $targetY; Invoke-MouseClick -X $baseX -Y $targetY
							Invoke-MouseClick -X $baseX -Y $targetY; Invoke-MouseClick -X $baseX -Y $targetY
						}
						elseif ($entryNumber -ge 1 -and $entryNumber -le 5)
						{
							$targetY = $baseY + (($entryNumber - 1) * 18)
							CheckCancel; EnsureWindowResponsive; CheckCancel
							Invoke-MouseClick -X $baseX -Y $targetY; Invoke-MouseClick -X $baseX -Y $targetY
							Invoke-MouseClick -X $baseX -Y $targetY; Invoke-MouseClick -X $baseX -Y $targetY
						}
					}
					else
					{
						$centerX = [int](($rect.Left + $rect.Right) / 2) + 25
						$centerY = [int](($rect.Top + $rect.Bottom) / 2) + 18
						$adjustedY = $centerY
						if ($entryNumber -ge 6 -and $entryNumber -le 10)
						{
							$yOffset = ($entryNumber - 8) * 18
							$adjustedY = $centerY + $yOffset
							Invoke-MouseClick -X ($centerX + 145) -Y ($centerY + 28); SleepWithCancel -Milliseconds 50
						}
						elseif ($entryNumber -ge 1 -and $entryNumber -le 5)
						{
							$yOffset = ($entryNumber - 3) * 18
							$adjustedY = $centerY + $yOffset
						}
						elseif ($entryNumber -ge 11) { return }
						CheckCancel; EnsureWindowResponsive; CheckCancel
						Invoke-MouseClick -X $centerX -Y $adjustedY; Invoke-MouseClick -X $centerX -Y $adjustedY
						Invoke-MouseClick -X $centerX -Y $adjustedY; Invoke-MouseClick -X $centerX -Y $adjustedY
						SleepWithCancel -Milliseconds 50
					}
					SleepWithCancel -Milliseconds 50
                
					
					CheckCancel; EnsureWindowResponsive; CheckCancel
					if ($config["Server${serverID}Coords"])
					{
						$coords = ParseCoordinates $config["Server${serverID}Coords"]
						if ($coords)
						{
							$currentStep++; ReportLoginProgress -Action "Selecting Server $serverID..." -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
							Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
							Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
							SleepWithCancel -Milliseconds 50
						}
					}
					SleepWithCancel -Milliseconds 50; CheckCancel; EnsureWindowResponsive; CheckCancel
                
					
					if ($config["Channel${channelID}Coords"])
					{
						$coords = ParseCoordinates $config["Channel${channelID}Coords"]
						if ($coords)
						{
							$currentStep++; ReportLoginProgress -Action "Selecting Channel $channelID..." -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
							Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
							Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
							SleepWithCancel -Milliseconds 50
						}
					}
					SleepWithCancel -Milliseconds 50; CheckCancel; EnsureWindowResponsive; CheckCancel
					$currentStep++; ReportLoginProgress -Action 'Entering Character Selection...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
					Invoke-KeyPress -VirtualKeyCode 0x0D

					$currentStep++
					$loginReady = Wait-ForLogEntry -LogPath $logPath -SearchStrings @('6 - LOGIN_PLAYER_LIST') -TimeoutSeconds 25
					if (-not $loginReady) { throw "CERT Login Timeout for Client $entryNumber" }
					SleepWithCancel -Milliseconds 50; CheckCancel; EnsureWindowResponsive; CheckCancel
                
					
					if ($config["Char${charSlot}Coords"] -and $config["Char${charSlot}Coords"] -ne '0,0')
					{
						$coords = ParseCoordinates $config["Char${charSlot}Coords"]
						if ($coords)
						{
							$currentStep++; ReportLoginProgress -Action "Selecting Character in Slot $charSlot" -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
							Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
							SleepWithCancel -Milliseconds 50
						}
					}
					CheckCancel; EnsureWindowResponsive; CheckCancel
					Invoke-KeyPress -VirtualKeyCode 0x0D; SleepWithCancel -Milliseconds 50

					$currentStep++
					$cacheJoin = Wait-ForLogEntry -LogPath $logPath -SearchStrings @('13 - CACHE_ACK_JOIN') -TimeoutSeconds 60
					if (-not $cacheJoin) { throw "Main Login Timeout for Client $entryNumber" }

					$currentStep++; Wait-UntilWorldLoaded -LogPath $logPath -Config $config -TotalSteps $totalGlobalSteps -CurrentStep $currentStep -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
					CheckCancel; EnsureWindowResponsive; CheckCancel
                
					$delay = if ($config['PostLoginDelay']) { [int]$config['PostLoginDelay'] } else { 1 }
					$currentStep++; ReportLoginProgress -Action "Post Login Delay ($delay s)" -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
					SleepWithCancel -Milliseconds ($delay * 1000)
					CheckCancel; EnsureWindowResponsive; CheckCancel

					
					if ($startCollector -eq 'Yes' -and $config['CollectorStartCoords'] -and $config['CollectorStartCoords'] -ne '0,0')
					{
						$coords = ParseCoordinates $config['CollectorStartCoords']
						if ($coords)
						{
							$currentStep++; ReportLoginProgress -Action 'Starting Collector...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
							Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
							SleepWithCancel -Milliseconds 1000
						}
					}

					$currentStep++; ReportLoginProgress -Action 'Minimizing...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
					[Custom.Native]::SendMessage($workingHwnd, 0x0112, 0xF020, 0)
					$currentStep++; ReportLoginProgress -Action 'Optimizing...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName
					[Custom.Native]::EmptyWorkingSet($process.Handle)
					$currentStep++; ReportLoginProgress -Action 'Finished...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $currentClientIndex -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName

				}
			}
			catch
			{
				if ($Global:CurrentActiveWindowHandle -ne [IntPtr]::Zero)
				{
					try { [Custom.Native]::ShowWindow($Global:CurrentActiveWindowHandle, 6) } catch {}
				}
				if ($_.Exception.Message -eq 'LoginCancelled' -or $_.ToString() -eq 'LoginCancelled')
				{
					Write-Verbose 'LOGIN: Background processing cancelled.'
				}
				else
				{
					throw $_
				}
			}
		}

		$infoEvent = {
			param($s, $e)
			if ($global:LoginResources['IsStopping']) { return }
			$now = Get-Date
			if ($null -eq $script:LastToastUpdateTime -or ($now - $script:LastToastUpdateTime).TotalMilliseconds -ge 150)
			{
				$script:LastToastUpdateTime = $now
				
				$record = $s[$e.Index]
				if ($record -and $record.Tags -contains 'LoginStatus')
				{
					$data = $record.MessageData
					$text = $data.Text
					$pct = $data.Percent

					$mainForm = $global:DashboardConfig.UI.MainForm
					if ($mainForm -and -not $mainForm.IsDisposed)
					{
						$mainForm.BeginInvoke([Action] {
								try
								{
									
									ShowToast -Title 'Login Progress' -Message $text -Type 'Info' -Key 9999 -TimeoutSeconds 0 -Progress $pct
								}
								catch
        						{
									Write-Warning "Toast Update Failed: $($_.Exception.Message)"
								}
							})
					}
				}
			}
		}
		$completionEvent = {
			param($s, $e)
			$state = $e.InvocationStateInfo.State
			$mainForm = $global:DashboardConfig.UI.GlobalProgressBar.FindForm()
			if ($state -eq 'Failed') { Write-Verbose "LOGIN EXCEPTION: $($e.InvocationStateInfo.Reason)" }
			if ($state -match 'Completed|Failed|Stopped')
			{
				if ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated)
				{
					$mainForm.BeginInvoke([Action] { CleanUpLoginResources -globalLoginResourcesRef $global:DashboardConfig})
				}
				else { CleanUpLoginResources -globalLoginResourcesRef $global:DashboardConfig }
			}
		}
		$eventName = 'LoginOp_' + [Guid]::NewGuid().ToString('N')

		
		$stepsVal = if ($TOTAL_STEPS_PER_CLIENT) { $TOTAL_STEPS_PER_CLIENT } else { 15 }
		$result = InvokeInManagedRunspace -RunspacePool $localRunspace -ScriptBlock $loginScript -AsJob -ArgumentList $jobs, $stepsVal, $global:DashboardConfig.Config['Options'], $global:DashboardConfig.Config['LoginConfig'], $global:LoginCancellation, $global:DashboardConfig.State

		if ($result -and $result.PowerShell)
		{
			$loginPS = $result.PowerShell

			
			$infoSub = Register-ObjectEvent -InputObject $loginPS.Streams.Information -EventName DataAdded -Action $infoEvent
			$global:LoginResources['InfoSubscription'] = $infoSub

			$eventSub = Register-ObjectEvent -InputObject $loginPS -EventName InvocationStateChanged -SourceIdentifier $eventName -Action $completionEvent
			$global:LoginResources['EventSubscriptionId'] = $eventName
			$global:LoginResources['EventSubscriber'] = $eventSub

			$global:LoginResources['PowerShellInstance'] = $loginPS
			$global:LoginResources['AsyncResult'] = $result.AsyncResult

			Write-Verbose 'LOGIN: Background process started.'
		}
		else
		{
			throw 'Failed to start login background job'
		}

	}
 catch
	{
		Write-Verbose "LOGIN: Error starting login sequence: $_"
		CleanUpLoginResources -globalLoginResourcesRef $global:DashboardConfig
	}
}

function CleanUpLoginResources
{
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		$globalLoginResourcesRef
	)

	
	if ($global:LoginResources['IsStopping'])
	{ 
		Write-Verbose 'LOGIN: CleanUpLoginResources called but already stopping.'
		return 
	}
	$global:LoginResources['IsStopping'] = $true

	Write-Verbose 'LOGIN: Cleaning up resources...' 

	if (-not $globalLoginResourcesRef) { $globalLoginResourcesRef = $global:DashboardConfig }

	
	if ($global:DashboardConfig.State.ActiveLoginPids)
	{
		$pidsToGrace = @($global:DashboardConfig.State.ActiveLoginPids)
		foreach ($pidToGrace in $pidsToGrace)
		{
			$global:DashboardConfig.State.LoginGracePids[$pidToGrace] = (Get-Date).AddSeconds(120)
		}
		$global:DashboardConfig.State.ActiveLoginPids.Clear()
	}

	
	if ($global:LoginResources.IsMouseHookActive)
	{
		try
		{
			[Custom.MouseHookManager]::Stop()
			$global:LoginResources.IsMouseHookActive = $false
			Write-Verbose 'LOGIN: Mouse hook stopped.'
		}
		catch
		{
			Write-Error "LOGIN: Failed to stop mouse hook: $_" 
		}
	}

	
	if (-not ($global:DashboardConfig.State.ReconnectQueue -and $global:DashboardConfig.State.ReconnectQueue.Count -gt 0))
	{
		CloseToast -Key 9999
	}

	
	$uiCleanupAction = {
		$ui = $global:DashboardConfig.UI
		if ($ui.LoginButton)
		{
			$ui.LoginButton.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
			$ui.LoginButton.Text = 'Login'
		}
        
		if ($global:DashboardConfig.State)
		{
			$global:DashboardConfig.State['LoginActive'] = $false
		}
	}
	
	$mainForm = $global:DashboardConfig.UI.GlobalProgressBar.FindForm()
    
	
	
	if ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated -and $mainForm.InvokeRequired)
	{
		try
		{ 
			
			$mainForm.BeginInvoke($uiCleanupAction)
		}
		catch
		{
			Write-Warning "LOGIN: Failed to BeginInvoke UI cleanup action: $_" 
			
			& $uiCleanupAction $globalLoginResourcesRef
		}
	}
 	elseif ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated)
	{
		
		& $uiCleanupAction $globalLoginResourcesRef
	}
 	else
	{
		
		Write-Warning 'LOGIN: UI form not available for cleanup, skipping UI updates.'
		$global:DashboardConfig.State['LoginActive'] = $false 
	}

	
	$localRes = $global:LoginResources.Clone()

	
	if ($localRes.EventSubscriptionId)
	{
		Unregister-Event -SourceIdentifier $localRes.EventSubscriptionId -ErrorAction SilentlyContinue
		Write-Verbose "LOGIN: Unregistered InvocationStateChanged event: $($localRes.EventSubscriptionId)" 
	}
	if ($localRes.InfoSubscription)
	{
		Unregister-Event -SubscriptionId $localRes.InfoSubscription.Id -ErrorAction SilentlyContinue
		Write-Verbose "LOGIN: Unregistered Information.DataAdded event (ID: $($localRes.InfoSubscription.Id))" 
	}

	[System.Threading.Tasks.Task]::Run([Action]{
		try
		{
			$ps = $localRes.PowerShellInstance; $rs = $localRes.Runspace; $ar = $localRes.AsyncResult
			if ($ps)
			{
				try { if ($ps.InvocationStateInfo.State -eq 'Running') { $ps.Stop() } } catch {}
				if ($ar) { try { $ps.EndInvoke($ar) } catch {} }
				try { $ps.Dispose() } catch {}
			}
			if ($rs) { try { $rs.Dispose() } catch {} }
		} catch {}
	}) | Out-Null
	
	$global:LoginResources = @{
		PowerShellInstance  = $null
		Runspace            = $null
		EventSubscriptionId = $null
		EventSubscriber     = $null
		AsyncResult         = $null
		IsStopping          = $false
		InfoSubscription    = $null
		IsMouseHookActive   = $false
	}
}
#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion