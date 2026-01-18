<# launch.psm1 #>

#region Configuration and Constants

$global:LaunchCancellation = [hashtable]::Synchronized(@{
		IsCancelled = $false
	})

$script:LauncherTimeout = 30
$script:LaunchDelay = 3
$script:MaxRetryAttempts = 3
$script:LastLaunchToastUpdateTime = $null
$script:LastLaunchToastTick = 0

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
		[string]$SetupName,

		[Parameter(Mandatory = $false)]
		[int]$SequenceRemainingCount = 0,

		[Parameter(Mandatory = $false)]
		[switch]$FromSequence
	)

	if ($SavedLaunchLoginConfig)
	{
		$global:LaunchCancellation.IsCancelled = $false
		InvokeSavedLaunchSequence -SetupName $SetupName
		return
	}

	if ($global:DashboardConfig.State.LaunchActive -and -not $FromSequence)
	{
		Show-DarkMessageBox 'Launch operation already in progress' 'Launch' 'Ok' 'Information'
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

	if (Get-Command ReadConfig -ErrorAction SilentlyContinue -Verbose:$False) { ReadConfig }

	$neuzName = $settingsDict['ProcessName']['ProcessName']
	$launcherPath = $settingsDict['LauncherPath']['LauncherPath']

	$profileToUse = 'Default'

	if ($settingsDict['Options'] -and $settingsDict['Options']['SelectedProfile'])
	{
		if (-not [string]::IsNullOrWhiteSpace($settingsDict['Options']['SelectedProfile']))
		{
			$profileToUse = $settingsDict['Options']['SelectedProfile']
		}
	}

	if (-not [string]::IsNullOrEmpty($ProfileNameOverride))
	{
		$profileToUse = $ProfileNameOverride
		Write-Verbose "LAUNCH: Overriding profile with '$profileToUse'"
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
		if (Get-Command GetProcessProfile -ErrorAction SilentlyContinue -Verbose:$False)
		{
			$allClients = Get-Process -Name $neuzName -ErrorAction SilentlyContinue -Verbose:$False
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

	$launcherTimeoutSeconds = 60
	if ($settingsDict['Options'] -and $settingsDict['Options']['LauncherTimeout'])
	{
		if (-not [int]::TryParse($settingsDict['Options']['LauncherTimeout'], [ref]$launcherTimeoutSeconds)) { $launcherTimeoutSeconds = 60 }
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
    
	$assembliesToAdd = [System.Collections.Generic.HashSet[string]]::new()
	$customTypesToLoad = @('Custom.Native', 'Custom.ColorWriter')
    
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
	
	try
	{
		$localRunspace = NewManagedRunspace `
			-Name 'LaunchRunspace' `
			-MinRunspaces 1 `
			-MaxRunspaces 1 `
			-SessionVariables $varsToInject `
			-Assemblies ($assembliesToAdd | Select-Object -Unique)

		if (-not $localRunspace) { throw "Runspace creation returned null." }
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

	$global:DashboardConfig.State.LaunchStatus = [hashtable]::Synchronized(@{
		Text = "Starting Launch Process..."
		Percent = 0
		LastUpdateTick = [DateTime]::Now.Ticks
	})

	$launchScript = {
		param(
			$Settings,
			$LauncherPath,
			$NeuzName,
			$MaxClients,
			$ClientsToLaunch,
			$IniPath,
			$CancellationContext,
			$SetupName,
			$SequenceRemainingCount,
			$ProfileName,
			$LauncherTimeoutSeconds
		)

		function CheckCancel
		{
			if ($CancellationContext.IsCancelled) { throw 'LaunchCancelled' }
		}

		function ReportLaunchProgress
		{
			param($Action, $Step, $TotalSteps)
			try {
				$pct = 0; if ($TotalSteps -gt 0) { $pct = [int](($Step / $TotalSteps) * 100) }
				$msg = "Profile: $ProfileName | Total: $pct%`n$Action"
				Write-Verbose -Message $msg
				if ($DashboardConfig -and $DashboardConfig.State -and $DashboardConfig.State.LaunchStatus) {
					$status = $DashboardConfig.State.LaunchStatus
					$status['Text'] = $msg
					$status['Percent'] = $pct
					$status['LastUpdateTick'] = [DateTime]::Now.Ticks
				}
			} catch {}
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
			$InformationPreference = 'SilentlyContinue'

			function Write-Verbose
			{
				[CmdletBinding()]
				param([string]$Message, [string]$ForegroundColor = 'DarkGray')
                
				$dateStr = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
				$callerName = 'StartClientLaunch'
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
			
			$isFreshLaunch = ($pCount -eq 0)

			if ($isFreshLaunch) {
				Write-Verbose "LAUNCH: No clients detected. Starting Main Launcher for patching..."
				ReportLaunchProgress -Action "Patching Main Client..." -Step 0 -TotalSteps $totalSteps
				
				$mainLogDir = Join-Path $launcherDir "Log"
				$initialSuccessCount = 0
				
				if (Test-Path $mainLogDir) {
					$logFile = Get-ChildItem -Path $mainLogDir -Filter "launcher*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
					if ($logFile) {
						try {
							$stream = [System.IO.File]::Open($logFile.FullName, 'Open', 'Read', 'ReadWrite')
							$reader = New-Object System.IO.StreamReader($stream)
							$content = $reader.ReadToEnd()
							$reader.Close(); $stream.Close()
							$initialSuccessCount = ([regex]::Matches($content, "Status: Update completed successfully!")).Count
						} catch {}
					}
				}

				try {
					$mainLauncherProc = Start-Process -FilePath $LauncherPath -WorkingDirectory $launcherDir -PassThru -WindowStyle Normal
					if ($mainLauncherProc) {
						$patchSuccess = $false
						$patchTimeout = [DateTime]::Now.AddSeconds($LauncherTimeoutSeconds)
						
						while ([DateTime]::Now -lt $patchTimeout) {
							CheckCancel
							if ($mainLauncherProc.HasExited) { break }
							
							if (Test-Path $mainLogDir) {
								$logFile = Get-ChildItem -Path $mainLogDir -Filter "launcher*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
								if ($logFile) {
									try {
										$stream = [System.IO.File]::Open($logFile.FullName, 'Open', 'Read', 'ReadWrite')
										$reader = New-Object System.IO.StreamReader($stream)
										$content = $reader.ReadToEnd()
										$reader.Close(); $stream.Close()
										
										$currentCount = ([regex]::Matches($content, "Status: Update completed successfully!")).Count
										if ($currentCount -gt $initialSuccessCount) {
											$patchSuccess = $true
											break
										}
									} catch {}
								}
							}
							Start-Sleep -Milliseconds 1000
						}
						
						if ($patchSuccess) {
							Write-Verbose "LAUNCH: Main Patch Successful. Closing Main Launcher."
							Stop-Process -Id $mainLauncherProc.Id -Force -ErrorAction SilentlyContinue
							$mainLauncherProc.WaitForExit(5000)
						} else {
							Write-Verbose "LAUNCH: Main Patch timeout or launcher closed. Proceeding..."
							if (-not $mainLauncherProc.HasExited) {
								Stop-Process -Id $mainLauncherProc.Id -Force -ErrorAction SilentlyContinue
							}
						}
					}
				} catch {
					Write-Verbose "LAUNCH: Error during main patching: $_"
				}
			}

			if ($pCount -gt 0) { Write-Verbose "LAUNCH: Found $pCount client(s)" }
			else { Write-Verbose 'LAUNCH: No clients found' }

			Write-Verbose "LAUNCH: Launching $ClientsToLaunch more"

			$existingPIDs = New-Object System.Collections.Generic.List[int]
			if ($currentClients) {
				$currentClients | ForEach-Object { $existingPIDs.Add($_.Id) }
			}
			$currentClients = $null

			$launchTasks = New-Object System.Collections.Generic.LinkedList[PSObject]
			1..$ClientsToLaunch | ForEach-Object { $launchTasks.AddLast([PSCustomObject]@{ RetryCount = 0; Id = $_ }) }

			$stepsPerClient = 8
			$totalSteps = $ClientsToLaunch * $stepsPerClient
			$currentStep = 0
			$clientsLaunchedCount = 0

			while ($launchTasks.Count -gt 0)
			{
				$task = $launchTasks.First.Value
				$launchTasks.RemoveFirst()

				$actualRunning = @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue)
				foreach ($proc in $actualRunning) {
					if (-not $existingPIDs.Contains($proc.Id)) {
						$existingPIDs.Add($proc.Id)
					}
				}

				CheckCancel
				$currentStep++
				
				$clientsLeftInBatch = $launchTasks.Count + 1
				$totalLeft = $SequenceRemainingCount + $clientsLeftInBatch
				$prefix = ""
				if (-not [string]::IsNullOrEmpty($SetupName)) { $prefix = "Setup: $SetupName [$totalLeft Left] | " }
				$attemptDisplay = $clientsLaunchedCount + 1

				ReportLaunchProgress -Action "${prefix}Client $attemptDisplay/$ClientsToLaunch Starting..." -Step $currentStep -TotalSteps $totalSteps

				$GetRelevantLaunchers = {
					$candidates = @(Get-Process -Name $launcherName -ErrorAction SilentlyContinue)
					if ($NeuzName -and $NeuzName -ne $launcherName) {
						$candidates += @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue)
					}
					if ($candidates.Count -eq 0) { return @() }
					return $candidates | Where-Object { 
						if ($existingPIDs.Contains($_.Id)) { return $false }
						if ($_.ProcessName -eq $NeuzName) { return $true }
						if ($_.ProcessName -eq $launcherName) {
							try {
								return ($_.MainModule.FileName -eq $LauncherPath)
							} catch { return $false }
						}
						return $false
					}
				}

				$relevantLaunchers = &$GetRelevantLaunchers
				$currentStep++
				if ($relevantLaunchers.Count -gt 0)
				{
					ReportLaunchProgress -Action "${prefix}Client $attemptDisplay/$ClientsToLaunch Waiting for Launcher (Timeout: ${LauncherTimeoutSeconds}s)..." -Step $currentStep -TotalSteps $totalSteps
					$launcherTimeout = New-TimeSpan -Seconds $LauncherTimeoutSeconds
					$launcherStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

					while (($relevantLaunchers = &$GetRelevantLaunchers).Count -gt 0)
					{
						CheckCancel
						$elapsedSeconds = [int]$launcherStopwatch.Elapsed.TotalSeconds

						if ($launcherStopwatch.Elapsed -gt $launcherTimeout)
						{
							Write-Verbose 'LAUNCH: Timeout - killing stuck launcher(s)'
							try { 
								foreach ($p in $relevantLaunchers) { try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch {} }
							} catch {}
							Start-Sleep -Seconds 1
							break
						}
						SleepWithCancel -Milliseconds 500
					}
				}
				else
				{
					ReportLaunchProgress -Action "${prefix}Client $attemptDisplay/$ClientsToLaunch Launcher Ready..." -Step $currentStep -TotalSteps $totalSteps
				}

				CheckCancel
				$currentStep++
				ReportLaunchProgress -Action "${prefix}Client $attemptDisplay/$ClientsToLaunch Executing Launcher..." -Step $currentStep -TotalSteps $totalSteps 
				
				$launcherProcess = Start-Process -FilePath $LauncherPath -WorkingDirectory $launcherDir -PassThru -WindowStyle Normal
				
				if ($launcherProcess)
				{
					$launcherPID = $launcherProcess.Id
					Write-Verbose "LAUNCH: PID: $launcherPID"
					if ($launcherPID) {
						$playCoords = $null
						if ($Settings['LoginConfig'] -and $Settings['LoginConfig'][$ProfileName] -and $Settings['LoginConfig'][$ProfileName]['LauncherPlayCoords']) {
							$playCoords = $Settings['LoginConfig'][$ProfileName]['LauncherPlayCoords']
						}
						
						if ($playCoords -and $playCoords -ne '0,0') {
							Write-Verbose "LAUNCH: Waiting for Profile Patch..."
							$profLogDir = Join-Path $launcherDir "Log"
							$profInitialCount = 0
							
							if (Test-Path $profLogDir) {
								$logFile = Get-ChildItem -Path $profLogDir -Filter "launcher*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
								if ($logFile) {
									try {
										$stream = [System.IO.File]::Open($logFile.FullName, 'Open', 'Read', 'ReadWrite')
										$reader = New-Object System.IO.StreamReader($stream)
										$content = $reader.ReadToEnd()
										$reader.Close(); $stream.Close()
										$profInitialCount = ([regex]::Matches($content, "Status: Update completed successfully!")).Count
									} catch {}
								}
							}

							$profPatchSuccess = $false
							$profPatchTimeout = [DateTime]::Now.AddSeconds($LauncherTimeoutSeconds)
							
							while ([DateTime]::Now -lt $profPatchTimeout) {
								CheckCancel
								if ($launcherProcess.HasExited) { break }
								
								if (Test-Path $profLogDir) {
									$logFile = Get-ChildItem -Path $profLogDir -Filter "launcher*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
									if ($logFile) {
										try {
											$stream = [System.IO.File]::Open($logFile.FullName, 'Open', 'Read', 'ReadWrite')
											$reader = New-Object System.IO.StreamReader($stream)
											$content = $reader.ReadToEnd()
											$reader.Close(); $stream.Close()
											
											$currentCount = ([regex]::Matches($content, "Status: Update completed successfully!")).Count
											if ($currentCount -gt $profInitialCount) {
												$profPatchSuccess = $true
												break
											}
										} catch {}
									}
								}
								Start-Sleep -Milliseconds 500
							}
							
							if ($profPatchSuccess) {
								Write-Verbose "LAUNCH: Profile Patch Complete. Clicking Play..."
								try {
									$hWnd = $launcherProcess.MainWindowHandle
									if ($hWnd -ne [IntPtr]::Zero) {
										if ('Custom.Native' -as [Type]) {
											[Custom.Native]::ShowWindow($hWnd, 1)
											[Custom.Native]::SetForegroundWindow($hWnd)
											Start-Sleep -Milliseconds 500
											
											$rect = New-Object Custom.Native+RECT
											[Custom.Native]::GetWindowRect($hWnd, [ref]$rect)
											
											$coords = $playCoords -split ','
											if ($coords.Count -eq 2) {
												$x = [int]$coords[0] + $rect.Left
												$y = [int]$coords[1] + $rect.Top
												
												[Custom.Native]::SetCursorPos($x, $y)
												Start-Sleep -Milliseconds 100
												[Custom.Native]::mouse_event(0x02, 0, 0, 0, 0)
												[Custom.Native]::mouse_event(0x04, 0, 0, 0, 0)
												Write-Verbose "LAUNCH: Clicked Play at $x, $y"
											}
										}
									}
								} catch {
									Write-Verbose "LAUNCH: Failed to click Play: $_"
								}
							}
						}
					}
					$playClicked = $false
					$playCoords = $null
					$initialSuccessCount = 0
					$logDir = Join-Path $launcherDir "Log"

					if ($Settings['LoginConfig'] -and $Settings['LoginConfig'][$ProfileName] -and $Settings['LoginConfig'][$ProfileName]['LauncherPlayCoords']) {
						$playCoords = $Settings['LoginConfig'][$ProfileName]['LauncherPlayCoords']
					}

					if ($playCoords -and (Test-Path $logDir)) {
						$logFile = Get-ChildItem -Path $logDir -Filter "launcher_*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
						if ($logFile) {
							try {
								$stream = [System.IO.File]::Open($logFile.FullName, 'Open', 'Read', 'ReadWrite')
								$reader = New-Object System.IO.StreamReader($stream)
								$content = $reader.ReadToEnd()
								$reader.Close(); $stream.Close()
								$initialSuccessCount = ([regex]::Matches($content, "Status: Update completed successfully!")).Count
							} catch {}
						}
					}

					Start-Sleep -Milliseconds 500
					try {
						$lProc = Get-Process -Id $launcherPID -ErrorAction SilentlyContinue
						if ($lProc -and $lProc.MainWindowHandle -ne [IntPtr]::Zero) {
							if ('Custom.Native' -as [Type]) {
								[Custom.Native]::ShowWindow($lProc.MainWindowHandle, 1)
								[Custom.Native]::SetForegroundWindow($lProc.MainWindowHandle)
							}
						}
					} catch {}
				}
				else 
				{
					Write-Verbose "LAUNCH: Failed to start Launcher process."
					continue
				}

				$currentStep++
				ReportLaunchProgress -Action "${prefix}Client $attemptDisplay/$ClientsToLaunch Initializing Client..." -Step $currentStep -TotalSteps $totalSteps 
				SleepWithCancel -Milliseconds 1000

				$timeout = New-TimeSpan -Minutes 2
				$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
				$launcherClosed = $false
				$null = $launcherClosedNormally; $launcherClosedNormally = $false
				$progressReported = @(1, 5, 15, 30, 60, 90, 120)

				$currentStep++
				ReportLaunchProgress -Action "${prefix}Client $attemptDisplay/$ClientsToLaunch Monitoring/Patching (Timeout: ${LauncherTimeoutSeconds}s)..." -Step $currentStep -TotalSteps $totalSteps 
				while (-not $launcherClosed -and $stopwatch.Elapsed -lt (New-TimeSpan -Seconds $LauncherTimeoutSeconds))
				{
					CheckCancel

					$currentPIDs = @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
					$newPIDs = $currentPIDs | Where-Object { -not $existingPIDs.Contains($_) }
					if ($newPIDs.Count -gt 0)
					{
						$launcherClosed = $true
						break
					}

					$elapsedSeconds = [int]$stopwatch.Elapsed.TotalSeconds
					if ($elapsedSeconds -in $progressReported)
					{
						$progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
						ReportLaunchProgress -Action "${prefix}Patching... ($elapsedSeconds s / $LauncherTimeoutSeconds)" -Step $currentStep -TotalSteps $totalSteps 

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

					if (-not $launcherClosed -and -not $playClicked -and $playCoords) {
						if (Test-Path $logDir) {
							$logFile = Get-ChildItem -Path $logDir -Filter "launcher_*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
							if ($logFile) {
								try {
									$stream = [System.IO.File]::Open($logFile.FullName, 'Open', 'Read', 'ReadWrite')
									$reader = New-Object System.IO.StreamReader($stream)
									$content = $reader.ReadToEnd()
									$reader.Close(); $stream.Close()
									
									$currentSuccessCount = ([regex]::Matches($content, "Status: Update completed successfully!")).Count
									
									if ($currentSuccessCount -gt $initialSuccessCount) {
										Write-Verbose "LAUNCH: Patch complete detected. Clicking Play..."
										
										$lProc = Get-Process -Id $launcherPID -ErrorAction SilentlyContinue
										if ($lProc -and $lProc.MainWindowHandle -ne [IntPtr]::Zero) {
											$hWnd = $lProc.MainWindowHandle
											$coords = $playCoords -split ','
											if ($coords.Count -eq 2) {
												$x = [int]$coords[0]; $y = [int]$coords[1]
												
												[Custom.Native]::ShowWindow($hWnd, 1)
												[Custom.Native]::SetForegroundWindow($hWnd)
												Start-Sleep -Milliseconds 200
												
												$rect = New-Object Custom.Native+RECT
												[Custom.Native]::GetWindowRect($hWnd, [ref]$rect)
												
												$absX = $rect.Left + $x
												$absY = $rect.Top + $y
												
												[Custom.Native]::SetCursorPos($absX, $absY)
												Start-Sleep -Milliseconds 50
												[Custom.Native]::mouse_event(0x02, 0, 0, 0, 0)
												[Custom.Native]::mouse_event(0x04, 0, 0, 0, 0)
												
												$playClicked = $true
											}
										}
									}
								} catch {}
							}
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
				$clientDetectionTimeout = New-TimeSpan -Seconds $LauncherTimeoutSeconds

				$currentStep++
				ReportLaunchProgress -Action "${prefix}Client $attemptDisplay/$ClientsToLaunch Waiting for Client Process (Timeout: ${LauncherTimeoutSeconds}s)..." -Step $currentStep -TotalSteps $totalSteps 
				while (-not $clientStarted -and $stopwatch.Elapsed -lt $clientDetectionTimeout)
				{
					CheckCancel
					$elapsedSeconds = [int]$stopwatch.Elapsed.TotalSeconds

					$currentPIDs = @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
					$newPIDs = $currentPIDs | Where-Object { -not $existingPIDs.Contains($_) }

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
								$existingPIDs.Add($newClientPID)
								Write-Verbose "LAUNCH: Client started PID: $newClientPID"
							}
						}
						catch
						{
							if ($newPIDs.Count -gt 0)
							{
								$newClientPID = $newPIDs[0]
								$clientStarted = $true
								$existingPIDs.Add($newClientPID)
								Write-Verbose "LAUNCH: Using PID: $newClientPID (fallback)"
							}
						}

						if ($clientStarted)
						{
							$windowReady = $false
							$innerTimeout = New-TimeSpan -Seconds $LauncherTimeoutSeconds
							$innerStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

							$currentStep++
							ReportLaunchProgress -Action "${prefix}Client $attemptDisplay/$ClientsToLaunch Waiting for Window (Timeout: ${LauncherTimeoutSeconds}s)..." -Step $currentStep -TotalSteps $totalSteps 
							while (-not $windowReady -and $innerStopwatch.Elapsed -lt $innerTimeout)
							{
								CheckCancel
								$clientExists = $false
								$clientWindowHandle = [IntPtr]::Zero
								$clientHandle = [IntPtr]::Zero
								$clientResponding = $false
								$tempClient = $null

								try
								{
									$tempClient = Get-Process -Id $newClientPID -ErrorAction SilentlyContinue
									if ($tempClient)
									{
										$clientExists = $true
										try { $clientResponding = $tempClient.Responding } catch { $clientResponding = $false }
										try { $clientWindowHandle = $tempClient.MainWindowHandle } catch { $clientWindowHandle = [IntPtr]::Zero }
										try { $clientHandle = $tempClient.Handle } catch { $clientHandle = [IntPtr]::Zero }
									}
								}
								catch { $clientExists = $false }

								if (-not $clientExists)
								{
									Write-Verbose 'LAUNCH: Client terminated'
									$clientStarted = $false
									if ($tempClient) { $tempClient.Dispose() }
									break
								}

								if ($clientResponding -and $clientWindowHandle -ne [IntPtr]::Zero)
								{
									$currentWindowHandle = $clientWindowHandle
									$windowReady = $true
									$currentStep++; ReportLaunchProgress -Action "${prefix}Client $attemptDisplay/$ClientsToLaunch Finalizing..." -Step $currentStep -TotalSteps $totalSteps
									SleepWithCancel -Milliseconds 500
									[Custom.Native]::SendMessage($clientWindowHandle, 0x0112, 0xF020, 0)
									try { [Custom.Native]::EmptyWorkingSet($clientHandle) } catch {}
									Write-Verbose "LAUNCH: Client ready: $newClientPID"
								}
								
								if ($tempClient) { $tempClient.Dispose() }
								
								SleepWithCancel -Milliseconds 500
							}
						}
					}
					SleepWithCancel -Milliseconds 500
				}

				if (-not $clientStarted -or -not $windowReady)
				{
					Write-Verbose "LAUNCH: Client failed to start or become responsive (Started: $clientStarted, Ready: $windowReady)"
					if ($newClientPID -gt 0)
					{
						Write-Verbose "LAUNCH: Killing unresponsive client PID $newClientPID"
						try { Stop-Process -Id $newClientPID -Force -ErrorAction SilentlyContinue } catch {}
					}
					
					$task.RetryCount++
					if ($task.RetryCount -eq 1)
					{
						Write-Verbose "LAUNCH: Retrying client immediately (Attempt 2)"
						$launchTasks.AddFirst($task)
					}
					elseif ($task.RetryCount -eq 2)
					{
						Write-Verbose "LAUNCH: Retrying client at end of batch (Attempt 3)"
						$launchTasks.AddLast($task)
					}
					else
					{
						Write-Verbose "LAUNCH: Client failed 3 times. Skipping."
					}
				}
				else
				{
					$clientsLaunchedCount++
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
				if (Get-Command -Name StopClientLaunch -ErrorAction SilentlyContinue -Verbose:$False)
				{
					StopClientLaunch -StepCompleted
				}
			}
			elseif ($state -eq 'Failed' -or $state -eq 'Stopped')
			{
				if (Get-Command -Name StopClientLaunch -ErrorAction SilentlyContinue -Verbose:$False)
				{
					StopClientLaunch
				}
			}
		}

		
		$res = InvokeInManagedRunspace `
			-RunspacePool $localRunspace `
			-ScriptBlock $launchScript `
			-AsJob `
			-ArgumentList $settingsDict, $launcherPath, $neuzName, $maxClients, $clientsToLaunch, $global:DashboardConfig.Paths.Ini, $global:LaunchCancellation, $SetupName, $SequenceRemainingCount, $profileToUse, $launcherTimeoutSeconds

		if (-not $res -or -not $res.PowerShell) { throw 'Failed to start launch background operation' }

		$launchPS = $res.PowerShell

		$global:LaunchResources.InfoSubscription = $null

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
				if (Get-Command -Name StopClientLaunch -ErrorAction SilentlyContinue -Verbose:$False) { StopClientLaunch }
			})

		$safetyTimer.Start()
		$global:DashboardConfig.Resources.Timers['launchSafetyTimer'] = $safetyTimer

		$uiUpdateTimer = New-Object System.Windows.Forms.Timer
		$uiUpdateTimer.Interval = 100
		$uiUpdateTimer.Add_Tick({
			try {
				$status = $global:DashboardConfig.State.LaunchStatus
				if ($status -and $status.LastUpdateTick -gt $script:LastLaunchToastTick)
				{
					$script:LastLaunchToastTick = $status.LastUpdateTick
					ShowToast -Title 'Launch Progress' -Message $status.Text -Type 'Info' -Key 9998 -TimeoutSeconds 0 -Progress $status.Percent
				}
			} catch {}
		})
		$uiUpdateTimer.Start()
		$global:DashboardConfig.Resources.Timers['launchUITimer'] = $uiUpdateTimer

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
	param([string]$SetupName)

	Write-Verbose 'LAUNCH: Initiating Smart Saved Launch Sequence...'

	$configSection = $null
	if (-not [string]::IsNullOrEmpty($SetupName))
	{
		$sectionKey = "Setup_$SetupName"
		if ($global:DashboardConfig.Config.Contains($sectionKey)) { $configSection = $global:DashboardConfig.Config[$sectionKey] }
	}
	elseif ($global:DashboardConfig.Config['SavedLaunchConfig'])
	{
		$configSection = $global:DashboardConfig.Config['SavedLaunchConfig']
	}

	if (-not $configSection -or $configSection.Count -eq 0)
	{
		Show-DarkMessageBox "No saved configuration found for setup '$SetupName'.`nPlease create a setup first!" 'Launch Setup' 'OK' 'Warning' 
		return
	}
	if ($global:DashboardConfig.State.LaunchActive)
	{
		Show-DarkMessageBox 'Launch operation already in progress' 'One-Click Setup' 'Ok' 'Information'
		return
	}

	$global:LaunchCancellation.IsCancelled = $false
	$global:DashboardConfig.State.LaunchActive = $true
	$global:DashboardConfig.State.SequenceActive = $true

	
	$requirements = [ordered]@{}
    
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
	$grid = $global:DashboardConfig.UI.DataGridMain

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

		$missingLogin = [Math]::Max(0, ($req.Login - $cur.Login))
		if ($missingLogin -gt 0) { $loginTargets[$p] = $missingLogin }

		if ($totalToLaunch -gt 0)
		{
			$launchQueue += @{ Profile = $p; Count = $totalToLaunch }
			Write-Verbose "LAUNCH PLAN [$p]: Launching $totalToLaunch. (Will Auto-Login: $missingLogin)"
		}
		elseif ($missingLogin -gt 0)
		{
			Write-Verbose "LAUNCH PLAN [$p]: No launch needed. (Will Auto-Login: $missingLogin)"
		}
	}

	$existingPIDs = @($grid.Rows | ForEach-Object { if ($_.Tag -and $_.Tag.Id) { $_.Tag.Id } })
	$global:DashboardConfig.Resources['RestoreSnapshotPIDs'] = $existingPIDs
	$global:DashboardConfig.Resources['LoginTargets'] = $loginTargets

	if ($launchQueue.Count -eq 0)
	{
		if ($loginTargets.Count -eq 0)
		{
			Write-Verbose 'LAUNCH SEQUENCE: State matches saved config. No action.'
			Show-DarkMessageBox 'All clients already match the saved configuration.' 'One-Click Setup' 'OK' 'Information'
			$global:DashboardConfig.State.LaunchActive = $false
			$global:DashboardConfig.State.SequenceActive = $false
			return
		}
	}

	$global:DashboardConfig.Resources['LaunchQueue'] = $launchQueue
	$global:DashboardConfig.Resources['LaunchQueueIndex'] = 0
	$global:DashboardConfig.Resources['CurrentLaunchSetupName'] = $SetupName

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

				$remainingInSeq = 0
				for ($i = $idx + 1; $i -lt $queue.Count; $i++) { $remainingInSeq += $queue[$i].Count }
				$sName = $global:DashboardConfig.Resources['CurrentLaunchSetupName']

				if ($item.Profile -eq 'Default')
				{
					StartClientLaunch -ClientAddCount $item.Count -FromSequence -SetupName $sName -SequenceRemainingCount $remainingInSeq
				}
				else
				{
					StartClientLaunch -ProfileNameOverride $item.Profile -ClientAddCount $item.Count -FromSequence -SetupName $sName -SequenceRemainingCount $remainingInSeq
				}
			}
			else
			{
				$this.Stop(); $this.Dispose()
				$global:DashboardConfig.Resources.Timers.Remove('SequenceTimer')

				Write-Verbose 'LAUNCH SEQUENCE: Launches complete. Waiting for clients to appear...'

				$loginWaitTimer = New-Object System.Windows.Forms.Timer
				$loginWaitTimer.Interval = 1000
				$loginWaitTimer.Tag = 'OneShot'
				$loginWaitTimer.Add_Tick({
						$this.Stop(); $this.Dispose()
						$global:DashboardConfig.Resources.Timers.Remove('LoginWaitTimer')
						if ($global:LaunchCancellation.IsCancelled) { return }

						$oldPIDs = $global:DashboardConfig.Resources['RestoreSnapshotPIDs']
						$targets = $global:DashboardConfig.Resources['LoginTargets']
						$grid = $global:DashboardConfig.UI.DataGridMain

						$sName = $global:DashboardConfig.Resources['CurrentLaunchSetupName']
						$cSection = $null
						if (-not [string]::IsNullOrEmpty($sName))
						{
							$k = "Setup_$sName"
							if ($global:DashboardConfig.Config.Contains($k)) { $cSection = $global:DashboardConfig.Config[$k] }
						}
						elseif ($global:DashboardConfig.Config['SavedLaunchConfig'])
						{
							$cSection = $global:DashboardConfig.Config['SavedLaunchConfig']
						}
						
						$savedSlots = [System.Collections.Generic.List[PSObject]]::new()
						if ($cSection)
						{
                    
							foreach ($key in $cSection.Keys)
							{
								$val = $cSection[$key]
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
						$sortedSlots = @($savedSlots | Sort-Object GridPos)
						$savedSlots = [System.Collections.Generic.List[PSObject]]::new()
						$savedSlots.AddRange([PSObject[]]$sortedSlots)

						
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

						$allRows = [System.Collections.Generic.List[Object]]::new()
						$allRows.AddRange($existingRows)
						$allRows.AddRange($newRows)
						
						$assignedRows = [System.Collections.Generic.HashSet[Object]]::new()

						$rowProfileCounters = @{}
						$rowsWithInfo = [System.Collections.Generic.List[Object]]::new()

						foreach ($row in $allRows) {
							$rt = $row.Cells[1].Value.ToString()
							$pn = 'Default'
							$ct = $rt
							if ($rt -match '^\[([^\]]+)\]\s*(.*)') { 
								$pn = $matches[1]
								$ct = $matches[2].Trim()
							}
							
							if (-not $rowProfileCounters.ContainsKey($pn)) { $rowProfileCounters[$pn] = 0 }
							$rowProfileCounters[$pn]++
							
							$rowsWithInfo.Add([PSCustomObject]@{
								Row = $row
								Title = $rt
								Profile = $pn
								CleanTitle = $ct
								RelativePos = $rowProfileCounters[$pn]
							})
						}

						foreach ($info in $rowsWithInfo) {
							$row = $info.Row
							if ($assignedRows.Contains($row)) { continue }
							
							foreach ($slot in $savedSlots) {
								if ($null -eq $slot.AssignedRow) {
									if ($slot.Title -eq $info.Title -or $slot.Title -eq $info.CleanTitle) {
										$slot.AssignedRow = $row
										$assignedRows.Add($row) | Out-Null
										break
									}
								}
							}
						}

						foreach ($info in $rowsWithInfo) {
							$row = $info.Row
							if ($assignedRows.Contains($row)) { continue }

							foreach ($slot in $savedSlots) {
								if ($null -eq $slot.AssignedRow -and $slot.Profile -eq $info.Profile -and $slot.GridPos -eq $info.RelativePos) {
									if ($slot.Title -ne $info.CleanTitle -and $slot.Title -match ' - ' -and $info.CleanTitle -match ' - ') { continue }
									
									$slot.AssignedRow = $row
									$assignedRows.Add($row) | Out-Null
									break
								}
							}
						}

						foreach ($info in $rowsWithInfo) {
							$row = $info.Row
							if ($assignedRows.Contains($row)) { continue }

							foreach ($slot in $savedSlots) {
								if ($null -eq $slot.AssignedRow -and $slot.Profile -eq $info.Profile) {
									if ($slot.Title -ne $info.CleanTitle -and $slot.Title -match ' - ' -and $info.CleanTitle -match ' - ') { continue }
									
									$slot.AssignedRow = $row
									$assignedRows.Add($row) | Out-Null
									break
								}
							}
						}

						$finalLoginList = @()
						$loginCounts = @{}

						foreach ($slot in $savedSlots) {
							if ($slot.AssignedRow) {
								$row = $slot.AssignedRow
								$pName = $slot.Profile

								$rowTitle = $row.Cells[1].Value.ToString()
								$cleanTitle = $rowTitle
								if ($rowTitle -match '^\[([^\]]+)\](.*)') { $cleanTitle = $matches[2] }
								$isLoggedIn = ($cleanTitle -match ' - ')

								if (-not $isLoggedIn) {
									if ($targets.Contains($pName)) {
										if (-not $loginCounts.Contains($pName)) { $loginCounts[$pName] = 0 }
										if ($loginCounts[$pName] -lt $targets[$pName]) {
											$wrapper = [PSCustomObject]@{
												Row               = $row
												OverrideAccountID = $slot.AccountID
											}
											$wrapper.PSObject.TypeNames.Insert(0, 'LoginOverrideWrapper')
											$finalLoginList += $wrapper
											$loginCounts[$pName]++
										}
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

							if (Get-Command LoginSelectedRow -ErrorAction SilentlyContinue -Verbose:$False)
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
	param([switch]$StepCompleted, [switch]$Sync)

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
		$localRes = $global:LaunchResources.Clone()

        $cleanupAction = {
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
        }

        if ($Sync) {
            & $cleanupAction
        } else {
		    [System.Threading.Tasks.Task]::Run([Action]$cleanupAction) | Out-Null
        }

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
		if ($global:DashboardConfig.Resources.Timers.Contains('launchUITimer'))
		{
			try
			{
				$global:DashboardConfig.Resources.Timers['launchUITimer'].Stop()
				$global:DashboardConfig.Resources.Timers['launchUITimer'].Dispose()
				$global:DashboardConfig.Resources.Timers.Remove('launchUITimer')
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
