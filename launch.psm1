<# launch.psm1 #>

#region Configuration and Constants

$global:LaunchCancellation = [hashtable]::Synchronized(@{
		IsCancelled = $false
	})

$script:LauncherTimeout = 30
$script:LaunchDelay = 3
$script:MaxRetryAttempts = 3
$script:LastLaunchToastUpdateTime = $null

$script:ProcessConfig = @{
	MaxRetries = 3
	RetryDelay = 500
	Timeout    = 60000
}

#endregion

#region Launch Management Functions

function StartClientLaunch
{
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		[string]$ProfileNameOverride,

		[Parameter(Mandatory = $false)]
		[switch]$OneClientOnly,

		[Parameter(Mandatory = $false)]
		[int]$ClientAddCount = 0,

		[Parameter(Mandatory = $false)]
		[switch]$SavedLaunchLoginConfig,

		[Parameter(Mandatory = $false)]
		[switch]$FromSequence
	)

	if ($SavedLaunchLoginConfig)
	{
		$global:LaunchCancellation.IsCancelled = $false
		InvokeSavedLaunchSequence
		return
	}

	if ($global:DashboardConfig.State.LaunchActive -and -not $FromSequence)
	{
		[Custom.DarkMessageBox]::Show('Launch operation already in progress', 'Launch', 'Ok', 'Information')
		return
	}

	$global:LaunchCancellation.IsCancelled = $false
	$global:DashboardConfig.State.LaunchActive = $true

	$settingsDict = [ordered]@{}
	foreach ($section in $global:DashboardConfig.Config.Keys)
	{
		$settingsDict[$section] = [ordered]@{}
		foreach ($key in $global:DashboardConfig.Config[$section].Keys)
		{
			$settingsDict[$section][$key] = $global:DashboardConfig.Config[$section][$key]
		}
	}

	if (Get-Command ReadConfig -ErrorAction SilentlyContinue) { ReadConfig }

	$neuzName = $settingsDict['ProcessName']['ProcessName']
	$launcherPath = $settingsDict['LauncherPath']['LauncherPath']

	$profileToUse = $null
	if (-not [string]::IsNullOrEmpty($ProfileNameOverride))
	{
		$profileToUse = $ProfileNameOverride
		Write-Verbose "LAUNCH: Overriding profile with '$profileToUse'"
	}
	elseif ($settingsDict['Options'] -and $settingsDict['Options']['SelectedProfile'])
	{
		$profileToUse = $settingsDict['Options']['SelectedProfile']
	}
	$profileToUse = 'Default'
	if (-not [string]::IsNullOrEmpty($ProfileNameOverride))
	{
		$profileToUse = $ProfileNameOverride
		Write-Verbose "LAUNCH: Overriding profile with '$profileToUse'"
	}
	elseif ($settingsDict['Options'] -and $settingsDict['Options']['SelectedProfile'])
	{
		if (-not [string]::IsNullOrWhiteSpace($settingsDict['Options']['SelectedProfile']))
		{
			$profileToUse = $settingsDict['Options']['SelectedProfile']
		}
	}

	if ($profileToUse)
	{
		if ($settingsDict['Profiles'] -and $settingsDict['Profiles'][$profileToUse])
		{
			$profilePath = $settingsDict['Profiles'][$profileToUse]
			if (Test-Path $profilePath)
			{
				$exeName = [System.IO.Path]::GetFileName($launcherPath)
				$profileLauncherPath = Join-Path $profilePath $exeName
				if (Test-Path $profileLauncherPath)
				{
					$launcherPath = $profileLauncherPath
					Write-Verbose "LAUNCH: Using Profile '$profileToUse'"
				}
				else
				{
					Write-Verbose "LAUNCH: Profile executable not found at '$profileLauncherPath'. Falling back to default."
				}
			}
		}
	}

	Write-Verbose "LAUNCH: Using ProcessName: $neuzName"
	Write-Verbose "LAUNCH: Using LauncherPath: $launcherPath"

	if ([string]::IsNullOrEmpty($neuzName) -or [string]::IsNullOrEmpty($launcherPath) -or -not (Test-Path $launcherPath))
	{
		Write-Verbose 'LAUNCH: Invalid launch settings'
		$global:DashboardConfig.State.LaunchActive = $false
		return
	}

	$maxClients = 1
	if ($settingsDict['MaxClients'].Contains('MaxClients') -and -not [string]::IsNullOrEmpty($settingsDict['MaxClients']['MaxClients']))
	{
		if (-not ([int]::TryParse($settingsDict['MaxClients']['MaxClients'], [ref]$maxClients)))
		{
			$maxClients = 1
		}
	}

	if ($ClientAddCount -gt 0)
	{
		$currentCount = (Get-Process -Name $neuzName -ErrorAction SilentlyContinue).Count
		$maxClients = $currentCount + $ClientAddCount
		Write-Verbose "LAUNCH: Adding $ClientAddCount client(s). Target total: $maxClients"
	}
	elseif ($OneClientOnly)
	{
		$currentCount = (Get-Process -Name $neuzName -ErrorAction SilentlyContinue).Count
		$maxClients = $currentCount + 1
		Write-Verbose "LAUNCH: OneClientOnly override active. Target total: $maxClients"
	}
	$clientsToLaunch = 0
	$profileClientCount = 0

	if ($ClientAddCount -gt 0)
	{
		$clientsToLaunch = $ClientAddCount
		Write-Verbose "LAUNCH: Adding $ClientAddCount client(s) to profile '$profileToUse'."
	}
	elseif ($OneClientOnly)
	{
		$clientsToLaunch = 1
		Write-Verbose "LAUNCH: OneClientOnly override active for profile '$profileToUse'. Launching 1."
	}
	else
	{
		if (Get-Command GetProcessProfile -ErrorAction SilentlyContinue)
		{
			$allClients = Get-Process -Name $neuzName -ErrorAction SilentlyContinue
			foreach ($client in $allClients)
			{
				$clientProfile = GetProcessProfile -Process $client
				if (([string]::IsNullOrEmpty($clientProfile) -and $profileToUse -eq 'Default') -or $clientProfile -eq $profileToUse)
				{
					$profileClientCount++
				}
			}
		}
		$clientsToLaunch = [Math]::Max(0, $maxClients - $profileClientCount)
	}

	if ($clientsToLaunch -le 0 -and -not $SavedLaunchLoginConfig)
	{
		Write-Verbose "LAUNCH: No clients to launch for profile '$profileToUse'. (Max: $maxClients, Current: $profileClientCount)"
		$global:DashboardConfig.State.LaunchActive = $false
		return
	}

	
	$varsToInject = @{
		'LaunchCancellation' = $global:LaunchCancellation
		'DashboardConfig'    = $global:DashboardConfig
	}
    
	
	try
	{
		$localRunspace = NewManagedRunspace `
			-Name 'LaunchRunspace' `
			-MinRunspaces 1 `
			-MaxRunspaces 1 `
			-SessionVariables $varsToInject
	}
 catch
	{
		Write-Verbose "LAUNCH: Error creating runspace: $_"
		$global:DashboardConfig.State.LaunchActive = $false
		return
	}

	
	$global:LaunchResources = @{
		PowerShellInstance  = $null
		Runspace            = $localRunspace
		EventSubscriptionId = $null
		EventSubscriber     = $null
		InfoSubscription    = $null
		AsyncResult         = $null
		StartTime           = [DateTime]::Now
	}

	ShowToast -Title 'Launch Process' -Message 'Starting Launch Process...' -Type 'Info' -Key 9998 -TimeoutSeconds 0

	$launchScript = {
		param(
			$Settings,
			$LauncherPath,
			$NeuzName,
			$MaxClients,
			$ClientsToLaunch,
			$IniPath,
			$CancellationContext
		)

		function CheckCancel
		{
			if ($CancellationContext.IsCancelled) { throw 'LaunchCancelled' }
		}

		function ReportLaunchProgress
		{
			param($Action, $Step, $TotalSteps)
			$pct = [int](($Step / $TotalSteps) * 100)
			$msg = "Total: $pct% | $Action"
			Write-Progress -Activity 'Launch' -Status $msg -PercentComplete $pct
			Write-Verbose -Message $msg
			Write-Information -MessageData @{ Text = $msg; Percent = $pct } -Tags 'LaunchStatus'
		}

		function SleepWithCancel
		{
			param([int]$Milliseconds)
			$sw = [System.Diagnostics.Stopwatch]::StartNew()
			while ($sw.Elapsed.TotalMilliseconds -lt $Milliseconds)
			{
				if ($CancellationContext.IsCancelled) { throw 'LaunchCancelled' }
				Start-Sleep -Milliseconds 100
			}
		}

		$currentWindowHandle = [IntPtr]::Zero

		try
		{
			CheckCancel
			Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

			if (-not ('Custom.ColorWriter' -as [Type]))
			{
				InitializeClassesModule
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

			$launcherDir = [System.IO.Path]::GetDirectoryName($LauncherPath)
			$launcherName = [System.IO.Path]::GetFileNameWithoutExtension($LauncherPath)

			Write-Verbose 'LAUNCH: Checking clients...'
			$currentClients = @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue)
			$pCount = $currentClients.Count

			if ($pCount -gt 0) { Write-Verbose "LAUNCH: Found $pCount client(s)" }
			else { Write-Verbose 'LAUNCH: No clients found' }

			Write-Verbose "LAUNCH: Launching $ClientsToLaunch more"

			$existingPIDs = $currentClients | Select-Object -ExpandProperty Id
			$currentClients = $null

			$stepsPerClient = 8
			$totalSteps = $ClientsToLaunch * $stepsPerClient
			$currentStep = 0

			for ($attempt = 1; $attempt -le $ClientsToLaunch; $attempt++)
			{
				CheckCancel
				$currentStep++
				ReportLaunchProgress -Action "Client $attempt/$ClientsToLaunch Starting Launch Process..." -Step $currentStep -TotalSteps $totalSteps

				$launcherRunning = $null -ne (Get-Process -Name $launcherName -ErrorAction SilentlyContinue)
				$currentStep++
				if ($launcherRunning)
				{
					ReportLaunchProgress -Action "Client $attempt/$ClientsToLaunch Waiting for Launcher..." -Step $currentStep -TotalSteps $totalSteps
					$launcherTimeout = New-TimeSpan -Seconds 60
					$launcherStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
					$progressReported = @(5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60)

					while ($null -ne (Get-Process -Name $launcherName -ErrorAction SilentlyContinue))
					{
						CheckCancel
						$elapsedSeconds = [int]$launcherStopwatch.Elapsed.TotalSeconds

						if ($elapsedSeconds -in $progressReported)
						{
							$progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
							Write-Verbose "LAUNCH: Waiting. ($elapsedSeconds s / 60)"
						}

						if ($launcherStopwatch.Elapsed -gt $launcherTimeout)
						{
							Write-Verbose 'LAUNCH: Timeout - killing'
							try { Stop-Process -Name $launcherName -Force -ErrorAction SilentlyContinue } catch {}
							Start-Sleep -Seconds 1
							break
						}
						SleepWithCancel -Milliseconds 500
					}
				}
				else
				{
					ReportLaunchProgress -Action "Client $attempt/$ClientsToLaunch Launcher Ready..." -Step $currentStep -TotalSteps $totalSteps
				}

				CheckCancel
				$currentStep++
				ReportLaunchProgress -Action "Client $attempt/$ClientsToLaunch Executing Launcher..." -Step $currentStep -TotalSteps $totalSteps 
				$launcherProcess = Start-Process -FilePath $LauncherPath -WorkingDirectory $launcherDir -PassThru
				$launcherPID = $launcherProcess.Id
				Write-Verbose "LAUNCH: PID: $launcherPID"

				$currentStep++
				ReportLaunchProgress -Action "Client $attempt/$ClientsToLaunch Initializing Client..." -Step $currentStep -TotalSteps $totalSteps 
				SleepWithCancel -Milliseconds 1000

				$timeout = New-TimeSpan -Minutes 2
				$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
				$launcherClosed = $false
				$null = $launcherClosedNormally; $launcherClosedNormally = $false
				$progressReported = @(1, 5, 15, 30, 60, 90, 120)

				$currentStep++
				ReportLaunchProgress -Action "Client $attempt/$ClientsToLaunch Monitoring/Patching..." -Step $currentStep -TotalSteps $totalSteps 
				while (-not $launcherClosed -and $stopwatch.Elapsed -lt $timeout)
				{
					CheckCancel
					$elapsedSeconds = [int]$stopwatch.Elapsed.TotalSeconds
					if ($elapsedSeconds -in $progressReported)
					{
						$progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
						ReportLaunchProgress -Action "Patching... ($elapsedSeconds s / 120)" -Step $currentStep -TotalSteps $totalSteps 

					}

					$launcherExists = $false
					$launcherResponding = $true

					try
					{
						$tempProcess = Get-Process -Id $launcherPID -ErrorAction SilentlyContinue
						if ($tempProcess)
						{
							$launcherExists = $true
							$launcherResponding = $tempProcess.Responding
						}
					}
					catch { $launcherExists = $false }

					if (-not $launcherExists)
					{
						$launcherClosed = $true
						$launcherClosedNormally = $true
					}
					else
					{
						if (-not $launcherResponding)
						{
							Write-Verbose 'LAUNCH: Not responding - killing'
							try { Stop-Process -Id $launcherPID -Force -ErrorAction SilentlyContinue } catch {}
							$launcherClosed = $true
							$launcherClosedNormally = $false
						}
					}

					if (-not $launcherClosed) { SleepWithCancel -Milliseconds 500 }
				}

				if (-not $launcherClosed)
				{
					Write-Verbose 'LAUNCH: Timeout - killing'
					try { Stop-Process -Id $launcherPID -Force -ErrorAction SilentlyContinue } catch {}
					$launcherClosedNormally = $false
				}

				$clientStarted = $false
				$newClientPID = 0
				$stopwatch.Restart()
				$clientDetectionTimeout = New-TimeSpan -Seconds 30
				$progressReported = @(5, 10, 15, 20, 25)

				$currentStep++
				ReportLaunchProgress -Action "Client $attempt/$ClientsToLaunch Waiting for Client Process..." -Step $currentStep -TotalSteps $totalSteps 
				while (-not $clientStarted -and $stopwatch.Elapsed -lt $clientDetectionTimeout)
				{
					CheckCancel
					$elapsedSeconds = [int]$stopwatch.Elapsed.TotalSeconds
					if ($elapsedSeconds -in $progressReported)
					{
						$progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
						Write-Verbose "LAUNCH: Waiting. ($elapsedSeconds s / 60)"
					}

					$currentPIDs = @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
					$newPIDs = $currentPIDs | Where-Object { $_ -notin $existingPIDs }

					if ($newPIDs.Count -gt 0)
					{
						try
						{
							$tempNewClients = @(Get-Process -Id $newPIDs -ErrorAction SilentlyContinue)
							$tempNewClient = $tempNewClients | Sort-Object StartTime -Descending | Select-Object -First 1
							if ($tempNewClient)
							{
								$newClientPID = $tempNewClient.Id
								$clientStarted = $true
								$existingPIDs += $newClientPID
								Write-Verbose "LAUNCH: Client started PID: $newClientPID"
							}
						}
						catch
						{
							if ($newPIDs.Count -gt 0)
							{
								$newClientPID = $newPIDs[0]
								$clientStarted = $true
								$existingPIDs += $newClientPID
								Write-Verbose "LAUNCH: Using PID: $newClientPID (fallback)"
							}
						}

						if ($clientStarted)
						{
							$windowReady = $false
							$innerTimeout = New-TimeSpan -Seconds 30
							$innerStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

							$currentStep++
							ReportLaunchProgress -Action "Client $attempt/$ClientsToLaunch Waiting for Window..." -Step $currentStep -TotalSteps $totalSteps 
							while (-not $windowReady -and $innerStopwatch.Elapsed -lt $innerTimeout)
							{
								CheckCancel
								$clientExists = $false
								$clientWindowHandle = [IntPtr]::Zero
								$clientHandle = [IntPtr]::Zero
								$clientResponding = $false

								try
								{
									$tempClient = Get-Process -Id $newClientPID -ErrorAction SilentlyContinue
									if ($tempClient)
									{
										$clientExists = $true
										$clientResponding = $tempClient.Responding
										$clientWindowHandle = $tempClient.MainWindowHandle
										$clientHandle = $tempClient.Handle
									}
								}
								catch { $clientExists = $false }

								if (-not $clientExists)
								{
									Write-Verbose 'LAUNCH: Client terminated'
									$clientStarted = $false
									break
								}

								if ($clientResponding -and $clientWindowHandle -ne [IntPtr]::Zero)
								{
									$currentWindowHandle = $clientWindowHandle
									$windowReady = $true
									$currentStep++; ReportLaunchProgress -Action "Client $attempt/$ClientsToLaunch Finalizing..." -Step $currentStep -TotalSteps $totalSteps
									SleepWithCancel -Milliseconds 500
									[Custom.Native]::ShowWindow($clientWindowHandle, [Custom.Native]::SW_MINIMIZE)
									try { [Custom.Native]::EmptyWorkingSet($clientHandle) } catch {}
									Write-Verbose "LAUNCH: Client ready: $newClientPID"
								}
								SleepWithCancel -Milliseconds 500
							}
						}
					}
					SleepWithCancel -Milliseconds 500
				}

				if (-not $clientStarted)
				{
					Write-Verbose 'LAUNCH: No client detected'
				}

				SleepWithCancel -Milliseconds 2000
			}

		}
		catch
		{
			if ($currentWindowHandle -ne [IntPtr]::Zero)
			{
				try { [Custom.Native]::ShowWindow($currentWindowHandle, 6) } catch {}
			}
			if ($_.Exception.Message -eq 'LaunchCancelled')
			{
				Write-Verbose 'LAUNCH: Operation Cancelled by User.'
			}
			else
			{
				Write-Verbose "LAUNCH: $($_.Exception.Message)"
			}
		}
		finally
		{
			[System.GC]::Collect()
		}
	}


	try
	{
		$completionScriptBlock = {
			Write-Verbose 'LAUNCH: Launch operation completed'
			StopClientLaunch
		}

		$script:LaunchCompletionAction = $completionScriptBlock
		$eventName = 'LaunchOperation_' + [Guid]::NewGuid().ToString('N')

		$simpleEventAction = {
			param($src, $e)
			$state = $e.InvocationStateInfo.State
			if ($state -eq 'Completed')
			{
				if (Get-Command -Name StopClientLaunch -ErrorAction SilentlyContinue)
				{
					StopClientLaunch -StepCompleted
				}
			}
			elseif ($state -eq 'Failed' -or $state -eq 'Stopped')
			{
				if (Get-Command -Name StopClientLaunch -ErrorAction SilentlyContinue)
				{
					StopClientLaunch
				}
			}
		}

		
		$res = InvokeInManagedRunspace `
			-RunspacePool $localRunspace `
			-ScriptBlock $launchScript `
			-AsJob `
			-ArgumentList $settingsDict, $launcherPath, $neuzName, $maxClients, $clientsToLaunch, $global:DashboardConfig.Paths.Ini, $global:LaunchCancellation

		if (-not $res -or -not $res.PowerShell) { throw 'Failed to start launch background operation' }

		$launchPS = $res.PowerShell

		$launchInfoEvent = {
			param($s, $e)
			$now = Get-Date
			if ($null -eq $script:LastLaunchToastUpdateTime -or ($now - $script:LastLaunchToastUpdateTime).TotalMilliseconds -ge 150)
			{
				$script:LastLaunchToastUpdateTime = $now
				$record = $s[$e.Index]
				if ($record -and $record.Tags -contains 'LaunchStatus')
				{
					$data = $record.MessageData
					$text = $data.Text
					$pct = $data.Percent
					$mainForm = $global:DashboardConfig.UI.MainForm
					if ($mainForm -and -not $mainForm.IsDisposed)
					{
						$mainForm.BeginInvoke([Action] {
								try { ShowToast -Title 'Launch Progress' -Message $text -Type 'Info' -Key 9998 -TimeoutSeconds 0 -Progress $pct } catch {}
							})
					}
				}
			}
		}
		$infoSub = Register-ObjectEvent -InputObject $launchPS.Streams.Information -EventName DataAdded -Action $launchInfoEvent
		$global:LaunchResources.InfoSubscription = $infoSub

		$eventSub = Register-ObjectEvent -InputObject $launchPS -EventName InvocationStateChanged -SourceIdentifier $eventName -Action $simpleEventAction

		if ($null -eq $eventSub) { throw 'Failed to register event subscriber' }

		$global:LaunchResources.EventSubscriptionId = $eventName
		$global:LaunchResources.EventSubscriber = $eventSub
		$global:LaunchResources.PowerShellInstance = $launchPS

		$safetyTimer = New-Object System.Timers.Timer
		$safetyTimer.Interval = 600000
		$safetyTimer.AutoReset = $false

		$safetyTimer.Add_Elapsed({
				Write-Verbose 'LAUNCH: Safety timer elapsed'
				if (Get-Command -Name StopClientLaunch -ErrorAction SilentlyContinue) { StopClientLaunch }
			})

		$safetyTimer.Start()
		$global:DashboardConfig.Resources.Timers['launchSafetyTimer'] = $safetyTimer

		$global:LaunchResources.AsyncResult = $res.AsyncResult

		if ($global:DashboardConfig.State.LaunchActive)
		{
			$global:DashboardConfig.UI.Launch.FlatStyle = 'Popup'
			$global:DashboardConfig.UI.Launch.Text = 'Cancel Launch'
		}

		Write-Verbose 'LAUNCH: Launch operation started'
	}
	catch
	{
		Write-Verbose "LAUNCH: Error in launch setup: $_"
		StopClientLaunch
	}
}

function InvokeSavedLaunchSequence
{
	Write-Verbose 'LAUNCH: Initiating Smart Saved Launch Sequence...'

	if (-not $global:DashboardConfig.Config['SavedLaunchConfig'] -or $global:DashboardConfig.Config['SavedLaunchConfig'].Count -eq 0)
	{
		[Custom.DarkMessageBox]::Show("No saved configuration found.`nPlease setup your clients first and create a One-Click Setup!", 'One-Click Setup', 'OK', 'Warning')
		return
	}
	if ($global:DashboardConfig.State.LaunchActive)
	{
		[Custom.DarkMessageBox]::Show('Launch operation already in progress', 'One-Click Setup', 'Ok', 'Information')
		return
	}

	$global:LaunchCancellation.IsCancelled = $false
	$global:DashboardConfig.State.LaunchActive = $true
	$global:DashboardConfig.State.SequenceActive = $true

	
	$requirements = [ordered]@{}
	$configSection = $global:DashboardConfig.Config['SavedLaunchConfig']
    
	foreach ($key in $configSection.Keys)
	{
		$parts = $configSection[$key] -split ',', 5
		if ($parts.Length -ge 4)
		{
			$p = $parts[1]
			$loginNeeded = $false
			if ($parts.Length -eq 5)
			{
				$loginNeeded = ($parts[4] -eq '1')
			}
			else
			{
				if ($parts[3] -match ' - ') { $loginNeeded = $true }
			}
			if (-not $requirements.Contains($p))
			{
				$requirements[$p] = @{ Login = 0; NoLogin = 0 }
			}
			if ($loginNeeded) { $requirements[$p].Login++ }
			else { $requirements[$p].NoLogin++ }
		}
	}

	
	$currentState = @{}
	$grid = $global:DashboardConfig.UI.DataGridFiller

	foreach ($row in $grid.Rows)
	{
		$title = $row.Cells[1].Value.ToString()
		$p = 'Default'
		$cleanTitle = $title
		if ($title -match '^\[([^\]]+)\](.*)')
		{
			$p = $matches[1]
			$cleanTitle = $matches[2]
		}
		$isLoggedIn = ($cleanTitle -match ' - ')
		if (-not $currentState.Contains($p))
		{
			$currentState[$p] = @{ Login = 0; NoLogin = 0 }
		}
		if ($isLoggedIn) { $currentState[$p].Login++ }
		else { $currentState[$p].NoLogin++ }
	}

	
	$launchQueue = @()
	$loginTargets = @{}

	foreach ($p in $requirements.Keys)
	{
		$req = $requirements[$p]
		$cur = if ($currentState.Contains($p)) { $currentState[$p] } else { @{ Login = 0; NoLogin = 0 } }
		$reqTotal = $req.Login + $req.NoLogin
		$curTotal = $cur.Login + $cur.NoLogin
		$totalToLaunch = [Math]::Max(0, ($reqTotal - $curTotal))

		if ($totalToLaunch -gt 0)
		{
			$launchQueue += @{ Profile = $p; Count = $totalToLaunch }
			$missingLoginRaw = [Math]::Max(0, ($req.Login - $cur.Login))
			$countToLogin = [Math]::Min($totalToLaunch, $missingLoginRaw)
			$loginTargets[$p] = $countToLogin
			Write-Verbose "LAUNCH PLAN [$p]: Launching $totalToLaunch. (Will Auto-Login: $countToLogin)"
		}
	}

	$existingPIDs = @($grid.Rows | ForEach-Object { if ($_.Tag -and $_.Tag.Id) { $_.Tag.Id } })
	$global:DashboardConfig.Resources['RestoreSnapshotPIDs'] = $existingPIDs
	$global:DashboardConfig.Resources['LoginTargets'] = $loginTargets

	if ($launchQueue.Count -eq 0)
	{
		Write-Verbose 'LAUNCH SEQUENCE: State matches saved config. No action.'
		[Custom.DarkMessageBox]::Show('All clients already match the saved configuration.', 'One-Click Setup', 'OK', 'Information')
		$global:DashboardConfig.State.LaunchActive = $false
		$global:DashboardConfig.State.SequenceActive = $false
		return
	}

	$global:DashboardConfig.Resources['LaunchQueue'] = $launchQueue
	$global:DashboardConfig.Resources['LaunchQueueIndex'] = 0

	if ($global:DashboardConfig.Resources.Timers.Contains('SequenceTimer'))
	{
		$global:DashboardConfig.Resources.Timers['SequenceTimer'].Stop()
		$global:DashboardConfig.Resources.Timers['SequenceTimer'].Dispose()
		$global:DashboardConfig.Resources.Timers.Remove('SequenceTimer')
	}

	$global:DashboardConfig.UI.Launch.FlatStyle = 'Popup'
	$global:DashboardConfig.UI.Launch.Text = 'Cancel Launch'

	$seqTimer = New-Object System.Windows.Forms.Timer
	$seqTimer.Interval = 1000

	$seqTimer.Add_Tick({
			if ($global:LaunchCancellation.IsCancelled)
			{
				$this.Stop(); $this.Dispose()
				return
			}
			if ($global:LaunchResources) { return }

			$queue = $global:DashboardConfig.Resources['LaunchQueue']
			$idx = $global:DashboardConfig.Resources['LaunchQueueIndex']

			if ($idx -lt $queue.Count)
			{
				$item = $queue[$idx]
				$global:DashboardConfig.Resources['LaunchQueueIndex'] = $idx + 1

				if ($item.Profile -eq 'Default')
				{
					StartClientLaunch -ClientAddCount $item.Count -FromSequence
				}
				else
				{
					StartClientLaunch -ProfileNameOverride $item.Profile -ClientAddCount $item.Count -FromSequence
				}
			}
			else
			{
				$this.Stop(); $this.Dispose()
				$global:DashboardConfig.Resources.Timers.Remove('SequenceTimer')

				Write-Verbose 'LAUNCH SEQUENCE: Launches complete. Waiting 5s for clients to appear...'

				$loginWaitTimer = New-Object System.Windows.Forms.Timer
				$loginWaitTimer.Interval = 5000
				$loginWaitTimer.Tag = 'OneShot'
				$loginWaitTimer.Add_Tick({
						$this.Stop(); $this.Dispose()
						$global:DashboardConfig.Resources.Timers.Remove('LoginWaitTimer')
						if ($global:LaunchCancellation.IsCancelled) { return }

						$oldPIDs = $global:DashboardConfig.Resources['RestoreSnapshotPIDs']
						$targets = $global:DashboardConfig.Resources['LoginTargets']
						$grid = $global:DashboardConfig.UI.DataGridFiller

						
						$savedSlots = [System.Collections.Generic.List[PSObject]]::new()
						if ($global:DashboardConfig.Config['SavedLaunchConfig'])
						{
							$configSection = $global:DashboardConfig.Config['SavedLaunchConfig']
                    
							foreach ($key in $configSection.Keys)
							{
								$val = $configSection[$key]
								$parts = $val -split ',', 5
                        
								if ($parts.Length -ge 3)
								{
									$gPos = [int]$parts[0].Trim()
									$prof = $parts[1].Trim()
									$acc = [int]$parts[2].Trim()
									$titl = if ($parts.Length -ge 4) { $parts[3].Trim() } else { '' }
                            
									$savedSlots.Add([PSCustomObject]@{
											GridPos     = $gPos
											Profile     = $prof
											AccountID   = $acc
											Title       = $titl
											AssignedRow = $null
										})
								}
							}
						}
						$savedSlots.Sort({ $args[0].GridPos - $args[1].GridPos })

						
						$existingRows = [System.Collections.Generic.List[Object]]::new()
						$newRows = [System.Collections.Generic.List[Object]]::new()
						$sortedGridRows = $grid.Rows | Sort-Object Index

						foreach ($row in $sortedGridRows)
						{
							if ($row.Tag -and $row.Tag.Id)
							{
								if ($oldPIDs -contains $row.Tag.Id) { $existingRows.Add($row) } 
								else { $newRows.Add($row) }
							}
						}

						
						foreach ($row in $existingRows)
						{
							$rowTitle = $row.Cells[1].Value.ToString()
                    
							
							foreach ($slot in $savedSlots)
							{
								if ($null -eq $slot.AssignedRow -and $slot.Title -eq $rowTitle)
								{
									$slot.AssignedRow = $row
									
									break
								}
							}
                    
							
							if ($null -eq ($savedSlots | Where-Object { $_.AssignedRow -eq $row }))
							{
								$pName = 'Default'
								if ($rowTitle -match '^\[([^\]]+)\]') { $pName = $matches[1] }
								foreach ($slot in $savedSlots)
								{
									if ($null -eq $slot.AssignedRow -and $slot.Profile -eq $pName)
									{
										$slot.AssignedRow = $row
										break
									}
								}
							}
						}

						
						$finalLoginList = @()
						$loginCounts = @{}

						foreach ($row in $newRows)
						{
							$rowTitle = $row.Cells[1].Value.ToString()
							$pName = 'Default'
							if ($rowTitle -match '^\[([^\]]+)\]') { $pName = $matches[1] }
                    
							$matchedSlot = $null
                    
							
							foreach ($slot in $savedSlots)
							{
								if ($null -eq $slot.AssignedRow -and $slot.Title -eq $rowTitle)
								{
									$matchedSlot = $slot
									break
								}
							}

							
							if ($null -eq $matchedSlot)
							{
								$visualGridPos = $row.Cells[0].Value -as [int]
								foreach ($slot in $savedSlots)
								{
									if ($null -eq $slot.AssignedRow -and $slot.Profile -eq $pName -and $slot.GridPos -eq $visualGridPos)
									{
										$matchedSlot = $slot
										break
									}
								}
							}

							
							if ($null -eq $matchedSlot)
							{
								foreach ($slot in $savedSlots)
								{
									if ($null -eq $slot.AssignedRow -and $slot.Profile -eq $pName)
									{
										$matchedSlot = $slot
										break
									}
								}
							}

							if ($matchedSlot)
							{
								$matchedSlot.AssignedRow = $row 

								if ($targets.Contains($pName))
								{
									if (-not $loginCounts.Contains($pName)) { $loginCounts[$pName] = 0 }
									if ($loginCounts[$pName] -lt $targets[$pName])
									{
                                
										
										$wrapper = [PSCustomObject]@{
											Row               = $row
											OverrideAccountID = $matchedSlot.AccountID 
										}
										$wrapper.PSObject.TypeNames.Insert(0, 'LoginOverrideWrapper')
                                
										$finalLoginList += $wrapper
										$loginCounts[$pName]++
									}
								}
							}
						}

						$global:DashboardConfig.UI.Launch.FlatStyle = 'Flat'
						$global:DashboardConfig.UI.Launch.Text = 'Launch ' + [char]0x25BE

						if ($finalLoginList.Count -gt 0)
						{
							$details = $finalLoginList | ForEach-Object {
								$pName = 'Unknown'
								if ($_.Row.Cells[1].Value -match '^\[([^\]]+)\]') { $pName = $matches[1] }
								"[$pName (Row:$($_.Row.Cells[0].Value)) -> Enforce Acc:$($_.OverrideAccountID)]"
							}
							$detailString = $details -join ', '
                    
							Write-Verbose "LAUNCH SEQUENCE: Auto-logging into $($finalLoginList.Count) specific clients."
							Write-Verbose "LAUNCH PLAN MAP: $detailString"

							if (Get-Command LoginSelectedRow -ErrorAction SilentlyContinue)
							{
								LoginSelectedRow -RowInput $finalLoginList
							}
							$global:DashboardConfig.State.LaunchActive = $false
							$global:DashboardConfig.State.SequenceActive = $false
						}
						else
						{
							Write-Verbose 'LAUNCH SEQUENCE: No clients require login scripts.'
							$global:DashboardConfig.State.LaunchActive = $false
							$global:DashboardConfig.State.SequenceActive = $false
						}
					})
				$global:DashboardConfig.Resources.Timers['LoginWaitTimer'] = $loginWaitTimer
				$loginWaitTimer.Start()
			}
		})
	$seqTimer.Start()
	$global:DashboardConfig.Resources.Timers['SequenceTimer'] = $seqTimer
}

function StopClientLaunch
{
	[CmdletBinding()]
	param([switch]$StepCompleted)

	if ($StepCompleted)
	{
		Write-Verbose 'LAUNCH: Cleaning up completed launch step...'
	}
 else
	{
		Write-Verbose 'LAUNCH: Cancelling all launch operations...'
		$global:LaunchCancellation.IsCancelled = $true
	}

	if (-not ($StepCompleted -and $global:DashboardConfig.State.SequenceActive))
	{
		$global:DashboardConfig.State.LaunchActive = $false
		$global:DashboardConfig.State.SequenceActive = $false
	}

	
	if (-not $StepCompleted)
	{
		try
		{
			if ($global:DashboardConfig.Resources.Timers.Contains('SequenceTimer'))
			{
				$global:DashboardConfig.Resources.Timers['SequenceTimer'].Stop()
				$global:DashboardConfig.Resources.Timers['SequenceTimer'].Dispose()
				$global:DashboardConfig.Resources.Timers.Remove('SequenceTimer')
				Write-Verbose 'LAUNCH: Queue timer stopped.'
			}
			if ($global:DashboardConfig.Resources.Timers.Contains('LoginWaitTimer'))
			{
				$global:DashboardConfig.Resources.Timers['LoginWaitTimer'].Stop()
				$global:DashboardConfig.Resources.Timers['LoginWaitTimer'].Dispose()
				$global:DashboardConfig.Resources.Timers.Remove('LoginWaitTimer')
				Write-Verbose 'LAUNCH: Login wait timer stopped.'
			}
		}
		catch {}
		if ($global:DashboardConfig.Resources.Contains('LaunchQueue'))
		{
			$global:DashboardConfig.Resources['LaunchQueue'] = @()
			$global:DashboardConfig.Resources['LaunchQueueIndex'] = 0
		}
	}

	if ($null -eq $global:LaunchResources)
	{
		
		if ($global:DashboardConfig.UI.Launch)
		{
			$global:DashboardConfig.UI.Launch.FlatStyle = 'Flat'
			$global:DashboardConfig.UI.Launch.Text = 'Launch ' + [char]0x25BE
		}
		return
	}

	try
	{
		
		if ($null -ne $global:LaunchResources.EventSubscriptionId)
		{
			try { Unregister-Event -SourceIdentifier $global:LaunchResources.EventSubscriptionId -Force -ErrorAction SilentlyContinue } catch {}
		}
		if ($null -ne $global:LaunchResources.EventSubscriber)
		{
			try { Unregister-Event -SubscriptionId $global:LaunchResources.EventSubscriber.Id -ErrorAction SilentlyContinue } catch {}
		}
		if ($null -ne $global:LaunchResources.InfoSubscription)
		{
			try { Unregister-Event -SubscriptionId $global:LaunchResources.InfoSubscription.Id -ErrorAction SilentlyContinue } catch {}
		}

		
		$psInstance = $global:LaunchResources.PowerShellInstance
		$null = $runspace; $runspace = $global:LaunchResources.Runspace
		$asyncResult = $global:LaunchResources.AsyncResult

		if ($asyncResult -and -not $asyncResult.IsCompleted)
		{
			
			$completedGracefully = $asyncResult.AsyncWaitHandle.WaitOne(5000)
			if (-not $completedGracefully -and $psInstance)
			{
				Write-Verbose 'LAUNCH: Graceful cancellation timed out. Forcing stop.'
				if ($psInstance.InvocationStateInfo.State -eq 'Running')
				{
					try { $psInstance.Stop() } catch {}
				}
			}
		}

		
		DisposeManagedRunspace -JobResource $global:LaunchResources
		if ($global:DashboardConfig.Resources.Timers.Contains('launchSafetyTimer'))
		{
			try
			{
				$global:DashboardConfig.Resources.Timers['launchSafetyTimer'].Stop()
				$global:DashboardConfig.Resources.Timers['launchSafetyTimer'].Dispose()
				$global:DashboardConfig.Resources.Timers.Remove('launchSafetyTimer')
			}
			catch {}
		}

		$global:LaunchResources = $null

		CloseToast -Key 9998

		[System.GC]::Collect()

		Write-Verbose 'LAUNCH: Launch resource cleanup completed'
	}
	catch
	{
		Write-Verbose "LAUNCH: Error during launch cleanup: $_"
	}
	finally
	{
		
		$global:DashboardConfig.State.LaunchActive = $false
		if ($global:DashboardConfig.UI.Launch)
		{
			$global:DashboardConfig.UI.Launch.FlatStyle = 'Flat'
			$global:DashboardConfig.UI.Launch.Text = 'Launch ' + [char]0x25BE
		}
	}
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion
