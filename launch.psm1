<# launch.psm1 #>

#region Configuration and Constants

$global:LaunchCancellation = [hashtable]::Synchronized(@{
    IsCancelled = $false
})

$script:LauncherTimeout = 30
$script:LaunchDelay = 3
$script:MaxRetryAttempts = 3

$script:ProcessConfig = @{
    MaxRetries = 3
    RetryDelay = 500
    Timeout    = 60000
}

#endregion

#region Launch Management Functions

function Start-ClientLaunch
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ProfileNameOverride,

        [Parameter(Mandatory=$false)]
        [switch]$OneClientOnly,

        [Parameter(Mandatory=$false)]
        [int]$ClientAddCount = 0,

		[Parameter(Mandatory=$false)]
        [switch]$SavedLaunchLoginConfig,

        [Parameter(Mandatory=$false)]
        [switch]$FromSequence
    )

    if ($SavedLaunchLoginConfig) {
        $global:LaunchCancellation.IsCancelled = $false
        Invoke-SavedLaunchSequence
        return
    }

    if ($global:DashboardConfig.State.LaunchActive -and -not $FromSequence)
    {
        [System.Windows.Forms.MessageBox]::Show('Launch operation already in progress', 'Information',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
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

    if (Get-Command Read-Config -ErrorAction SilentlyContinue) { Read-Config }

    $neuzName = $settingsDict['ProcessName']['ProcessName']
    $launcherPath = $settingsDict['LauncherPath']['LauncherPath']

    $profileToUse = $null
    if (-not [string]::IsNullOrEmpty($ProfileNameOverride)) {
        $profileToUse = $ProfileNameOverride
        Write-Verbose "LAUNCH: Overriding profile with '$profileToUse'" -ForegroundColor Cyan
    }
    elseif ($settingsDict['Options'] -and $settingsDict['Options']['SelectedProfile']) {
        $profileToUse = $settingsDict['Options']['SelectedProfile']
    }
	$profileToUse = "Default"
	if (-not [string]::IsNullOrEmpty($ProfileNameOverride)) {
		$profileToUse = $ProfileNameOverride
		Write-Verbose "LAUNCH: Overriding profile with '$profileToUse'" -ForegroundColor Cyan
	}
	elseif ($settingsDict['Options'] -and $settingsDict['Options']['SelectedProfile']) {
		if (-not [string]::IsNullOrWhiteSpace($settingsDict['Options']['SelectedProfile'])) {
			$profileToUse = $settingsDict['Options']['SelectedProfile']
		}
	}

    if ($profileToUse) {
        if ($settingsDict['Profiles'] -and $settingsDict['Profiles'][$profileToUse]) {
            $profilePath = $settingsDict['Profiles'][$profileToUse]
            if (Test-Path $profilePath) {
                $exeName = [System.IO.Path]::GetFileName($launcherPath)
                $profileLauncherPath = Join-Path $profilePath $exeName
                if (Test-Path $profileLauncherPath) {
                    $launcherPath = $profileLauncherPath
                    Write-Verbose "LAUNCH: Using Profile '$profileToUse'" -ForegroundColor Cyan
                } else {
                    Write-Verbose "LAUNCH: Profile executable not found at '$profileLauncherPath'. Falling back to default." -ForegroundColor Yellow
                }
            }
        }
    }

    Write-Verbose "LAUNCH: Using ProcessName: $neuzName" -ForegroundColor DarkGray
    Write-Verbose "LAUNCH: Using LauncherPath: $launcherPath" -ForegroundColor DarkGray

    if ([string]::IsNullOrEmpty($neuzName) -or [string]::IsNullOrEmpty($launcherPath) -or -not (Test-Path $launcherPath))
    {
        Write-Verbose 'LAUNCH: Invalid launch settings' -ForegroundColor Yellow
        $global:DashboardConfig.State.LaunchActive = $false
        return
    }

    $maxClients = 1
    if ($settingsDict['MaxClients'].Contains('MaxClients') -and -not [string]::IsNullOrEmpty($settingsDict['MaxClients']['MaxClients']))
    {
        if (-not ([int]::TryParse($settingsDict['MaxClients']['MaxClients'], [ref]$maxClients))) {
            $maxClients = 1
        }
    }

    if ($ClientAddCount -gt 0) {
        $currentCount = (Get-Process -Name $neuzName -ErrorAction SilentlyContinue).Count
        $maxClients = $currentCount + $ClientAddCount
        Write-Verbose "LAUNCH: Adding $ClientAddCount client(s). Target total: $maxClients" -ForegroundColor Cyan
    }
    elseif ($OneClientOnly) {
        $currentCount = (Get-Process -Name $neuzName -ErrorAction SilentlyContinue).Count
        $maxClients = $currentCount + 1
        Write-Verbose "LAUNCH: OneClientOnly override active. Target total: $maxClients" -ForegroundColor Cyan
    }
	$clientsToLaunch = 0
	$profileClientCount = 0

	if ($ClientAddCount -gt 0) {
		$clientsToLaunch = $ClientAddCount
		Write-Verbose "LAUNCH: Adding $ClientAddCount client(s) to profile '$profileToUse'." -ForegroundColor Cyan
	}
	elseif ($OneClientOnly) {
		$clientsToLaunch = 1
		Write-Verbose "LAUNCH: OneClientOnly override active for profile '$profileToUse'. Launching 1." -ForegroundColor Cyan
	}
	else {
		if (Get-Command Get-ProcessProfile -ErrorAction SilentlyContinue) {
			$allClients = Get-Process -Name $neuzName -ErrorAction SilentlyContinue
			foreach ($client in $allClients) {
				$clientProfile = Get-ProcessProfile -Process $client
				if (([string]::IsNullOrEmpty($clientProfile) -and $profileToUse -eq 'Default') -or $clientProfile -eq $profileToUse) {
					$profileClientCount++
				}
			}
		}
		$clientsToLaunch = [Math]::Max(0, $maxClients - $profileClientCount)
	}

	if ($clientsToLaunch -le 0 -and -not $SavedLaunchLoginConfig) {
		Write-Verbose "LAUNCH: No clients to launch for profile '$profileToUse'. (Max: $maxClients, Current: $profileClientCount)" -ForegroundColor Green
		$global:DashboardConfig.State.LaunchActive = $false
		return
	}
    $localRunspace = $null
    try
    {
        $localRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 1)
        $localRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
        $localRunspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $localRunspace.Open()
    }
    catch
    {
        Write-Verbose "LAUNCH: Error creating runspace: $_" -ForegroundColor Red
        $global:DashboardConfig.State.LaunchActive = $false
        return
    }

    $launchPS = [PowerShell]::Create()
    $launchPS.RunspacePool = $localRunspace

    $global:LaunchResources = @{
        PowerShellInstance  = $launchPS
        Runspace            = $localRunspace
        EventSubscriptionId = $null
        EventSubscriber     = $null
        AsyncResult         = $null
        StartTime           = [DateTime]::Now
    }

    $launchPS.AddScript({
            param(
                $Settings,
                $LauncherPath,
                $NeuzName,
                $MaxClients,
                $ClientsToLaunch,
                $IniPath,
                $CancellationContext
            )

            function CheckCancel {
                if ($CancellationContext.IsCancelled) { throw "LaunchCancelled" }
            }

            function SleepWithCancel {
                param([int]$Milliseconds)
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while ($sw.Elapsed.TotalMilliseconds -lt $Milliseconds) {
                    if ($CancellationContext.IsCancelled) { throw "LaunchCancelled" }
                    Start-Sleep -Milliseconds 100
                }
            }

            try
            {
                CheckCancel
                Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

                $classes = @'
                using System;
                public class ColorWriter
                {
                    public static void WriteColored(string message, string color)
                    {
                        ConsoleColor originalColor = Console.ForegroundColor;
                        switch(color.ToLower()) {
                            case "darkgray": Console.ForegroundColor = ConsoleColor.DarkGray; break;
                            case "yellow": Console.ForegroundColor = ConsoleColor.Yellow; break;
                            case "red": Console.ForegroundColor = ConsoleColor.Red; break;
                            case "cyan": Console.ForegroundColor = ConsoleColor.Cyan; break;
                            case "green": Console.ForegroundColor = ConsoleColor.Green; break;
                            default: Console.ForegroundColor = ConsoleColor.DarkGray; break;
                        }
                        Console.WriteLine(message);
                        Console.ForegroundColor = originalColor;
                    }
                }
'@
                Add-Type -TypeDefinition $classes -Language 'CSharp'

                function Write-Verbose {
                    [CmdletBinding()]
                    param([string]$Object, [string]$ForegroundColor = 'darkgray')
                    [ColorWriter]::WriteColored($Object, $ForegroundColor)
                }

                $launcherDir = [System.IO.Path]::GetDirectoryName($LauncherPath)
                $launcherName = [System.IO.Path]::GetFileNameWithoutExtension($LauncherPath)

                Write-Verbose 'LAUNCH: Checking clients...' -ForegroundColor DarkGray
                $currentClients = @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue)
                $pCount = $currentClients.Count

                if ($pCount -gt 0) { Write-Verbose "LAUNCH: Found $pCount client(s)" -ForegroundColor DarkGray }
                else { Write-Verbose 'LAUNCH: No clients found' -ForegroundColor DarkGray }

                Write-Verbose "LAUNCH: Launching $ClientsToLaunch more" -ForegroundColor Cyan

                $existingPIDs = $currentClients | Select-Object -ExpandProperty Id
                $currentClients = $null

                for ($attempt = 1; $attempt -le $ClientsToLaunch; $attempt++)
                {
                    CheckCancel
                    Write-Verbose "LAUNCH: Client $attempt/$ClientsToLaunch" -ForegroundColor Cyan

                    $launcherRunning = $null -ne (Get-Process -Name $launcherName -ErrorAction SilentlyContinue)
                    if ($launcherRunning)
                    {
                        Write-Verbose 'LAUNCH: Launcher running. Waiting (60s)...' -ForegroundColor DarkGray
                        $launcherTimeout = New-TimeSpan -Seconds 60
                        $launcherStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                        $progressReported = @(5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60)

                        while ($null -ne (Get-Process -Name $launcherName -ErrorAction SilentlyContinue))
                        {
                            CheckCancel
                            $elapsedSeconds = [int]$launcherStopwatch.Elapsed.TotalSeconds

                            if ($elapsedSeconds -in $progressReported) {
                                $progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
                                Write-Verbose "LAUNCH: Waiting. ($elapsedSeconds s / 60)" -ForegroundColor DarkGray
                            }

                            if ($launcherStopwatch.Elapsed -gt $launcherTimeout) {
                                Write-Verbose 'LAUNCH: Timeout - killing' -ForegroundColor Yellow
                                try { Stop-Process -Name $launcherName -Force -ErrorAction SilentlyContinue } catch {}
                                Start-Sleep -Seconds 1
                                break
                            }
                            SleepWithCancel -Milliseconds 500
                        }
                    }

                    CheckCancel
                    Write-Verbose 'LAUNCH: Starting launcher' -ForegroundColor DarkGray
                    $launcherProcess = Start-Process -FilePath $LauncherPath -WorkingDirectory $launcherDir -PassThru
                    $launcherPID = $launcherProcess.Id
                    Write-Verbose "LAUNCH: PID: $launcherPID" -ForegroundColor DarkGray

                    Write-Verbose 'LAUNCH: Initializing...' -ForegroundColor DarkGray
                    SleepWithCancel -Milliseconds 1000

                    $timeout = New-TimeSpan -Minutes 2
                    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    $launcherClosed = $false
                    $null = $launcherClosedNormally; $launcherClosedNormally = $false
                    $progressReported = @(1, 5, 15, 30, 60, 90, 120)

                    Write-Verbose 'LAUNCH: Monitoring (2min)' -ForegroundColor DarkGray
                    while (-not $launcherClosed -and $stopwatch.Elapsed -lt $timeout)
                    {
                        CheckCancel
                        $elapsedSeconds = [int]$stopwatch.Elapsed.TotalSeconds
                        if ($elapsedSeconds -in $progressReported) {
                            $progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
                            Write-Verbose "LAUNCH: Patching... ($elapsedSeconds s / 120)" -ForegroundColor Yellow
                        }

                        $launcherExists = $false
                        $launcherResponding = $true

                        try {
                            $tempProcess = Get-Process -Id $launcherPID -ErrorAction SilentlyContinue
                            if ($tempProcess) {
                                $launcherExists = $true
                                $launcherResponding = $tempProcess.Responding
                            }
                        } catch { $launcherExists = $false }

                        if (-not $launcherExists) {
                            $launcherClosed = $true
                            $launcherClosedNormally = $true
                        }
                        else {
                            if (-not $launcherResponding) {
                                Write-Verbose "LAUNCH: Not responding - killing" -ForegroundColor Yellow
                                try { Stop-Process -Id $launcherPID -Force -ErrorAction SilentlyContinue } catch {}
                                $launcherClosed = $true
                                $launcherClosedNormally = $false
                            }
                        }

                        if (-not $launcherClosed) { SleepWithCancel -Milliseconds 500 }
                    }

                    if (-not $launcherClosed) {
                        Write-Verbose 'LAUNCH: Timeout - killing' -ForegroundColor Red
                        try { Stop-Process -Id $launcherPID -Force -ErrorAction SilentlyContinue } catch {}
                        $launcherClosedNormally = $false
                    }

                    $clientStarted = $false
                    $newClientPID = 0
                    $stopwatch.Restart()
                    $clientDetectionTimeout = New-TimeSpan -Seconds 30
                    $progressReported = @(5, 10, 15, 20, 25)

                    Write-Verbose 'LAUNCH: Waiting for client' -ForegroundColor DarkGray
                    while (-not $clientStarted -and $stopwatch.Elapsed -lt $clientDetectionTimeout)
                    {
                        CheckCancel
                        $elapsedSeconds = [int]$stopwatch.Elapsed.TotalSeconds
                        if ($elapsedSeconds -in $progressReported) {
                            $progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
                            Write-Verbose "LAUNCH: Waiting. ($elapsedSeconds s / 60)" -ForegroundColor DarkGray
                        }

                        $currentPIDs = @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
                        $newPIDs = $currentPIDs | Where-Object { $_ -notin $existingPIDs }

                        if ($newPIDs.Count -gt 0)
                        {
                            try {
                                $tempNewClients = @(Get-Process -Id $newPIDs -ErrorAction SilentlyContinue)
                                $tempNewClient = $tempNewClients | Sort-Object StartTime -Descending | Select-Object -First 1
                                if ($tempNewClient) {
                                    $newClientPID = $tempNewClient.Id
                                    $clientStarted = $true
                                    $existingPIDs += $newClientPID
                                    Write-Verbose "LAUNCH: Client started PID: $newClientPID" -ForegroundColor DarkGray
                                }
                            } catch {
                                if ($newPIDs.Count -gt 0) {
                                    $newClientPID = $newPIDs[0]
                                    $clientStarted = $true
                                    $existingPIDs += $newClientPID
                                    Write-Verbose "LAUNCH: Using PID: $newClientPID (fallback)" -ForegroundColor Yellow
                                }
                            }

                            if ($clientStarted)
                            {
                                $windowReady = $false
                                $innerTimeout = New-TimeSpan -Seconds 30
                                $innerStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

                                Write-Verbose 'LAUNCH: Waiting for window' -ForegroundColor DarkGray
                                while (-not $windowReady -and $innerStopwatch.Elapsed -lt $innerTimeout)
                                {
                                    CheckCancel
                                    $clientExists = $false
                                    $clientWindowHandle = [IntPtr]::Zero
                                    $clientHandle = [IntPtr]::Zero
                                    $clientResponding = $false

                                    try {
                                        $tempClient = Get-Process -Id $newClientPID -ErrorAction SilentlyContinue
                                        if ($tempClient) {
                                            $clientExists = $true
                                            $clientResponding = $tempClient.Responding
                                            $clientWindowHandle = $tempClient.MainWindowHandle
                                            $clientHandle = $tempClient.Handle
                                        }
                                    } catch { $clientExists = $false }

                                    if (-not $clientExists) {
                                        Write-Verbose 'LAUNCH: Client terminated' -ForegroundColor Red
                                        $clientStarted = $false
                                        break
                                    }

                                    if ($clientResponding -and $clientWindowHandle -ne [IntPtr]::Zero) {
                                        $windowReady = $true
                                        Write-Verbose 'LAUNCH: Window ready' -ForegroundColor DarkGray
                                        SleepWithCancel -Milliseconds 500
                                        [Custom.Native]::ShowWindow($clientWindowHandle, [Custom.Native]::SW_MINIMIZE)
                                        try { [Custom.Native]::EmptyWorkingSet($clientHandle) } catch {}
                                        Write-Verbose "LAUNCH: Client ready: $newClientPID" -ForegroundColor Green
                                    }
                                    SleepWithCancel -Milliseconds 500
                                }
                            }
                        }
                        SleepWithCancel -Milliseconds 500
                    }

                    if (-not $clientStarted) {
                        Write-Verbose 'LAUNCH: No client detected' -ForegroundColor Yellow
                    }

                    SleepWithCancel -Milliseconds 2000
                }

            }
            catch
            {
                if ($_.Exception.Message -eq "LaunchCancelled") {
                    Write-Verbose "LAUNCH: Operation Cancelled by User." -ForegroundColor Yellow
                } else {
                    Write-Verbose "LAUNCH: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            finally
            {
                [System.GC]::Collect()
            }
        }).AddArgument($settingsDict).AddArgument($launcherPath).AddArgument($neuzName).AddArgument($maxClients).AddArgument($clientsToLaunch).AddArgument($global:DashboardConfig.Paths.Ini).AddArgument($global:LaunchCancellation)


    try
    {
        $completionScriptBlock = {
            Write-Verbose 'LAUNCH: Launch operation completed' -ForegroundColor Green
            Stop-ClientLaunch
        }

        $script:LaunchCompletionAction = $completionScriptBlock
        $eventName = 'LaunchOperation_' + [Guid]::NewGuid().ToString('N')

        $simpleEventAction = {
            param($src, $e)
            $state = $e.InvocationStateInfo.State
            if ($state -eq 'Completed')
            {
                if (Get-Command -Name Stop-ClientLaunch -ErrorAction SilentlyContinue) {
                    Stop-ClientLaunch -StepCompleted
                }
            }
            elseif ($state -eq 'Failed' -or $state -eq 'Stopped')
            {
                if (Get-Command -Name Stop-ClientLaunch -ErrorAction SilentlyContinue) {
                    Stop-ClientLaunch
                }
            }
        }

        $eventSub = Register-ObjectEvent -InputObject $launchPS -EventName InvocationStateChanged -SourceIdentifier $eventName -Action $simpleEventAction

        if ($null -eq $eventSub) { throw 'Failed to register event subscriber' }

        $global:LaunchResources.EventSubscriptionId = $eventName
        $global:LaunchResources.EventSubscriber = $eventSub

        $safetyTimer = New-Object System.Timers.Timer
        $safetyTimer.Interval = 600000
        $safetyTimer.AutoReset = $false

        $safetyTimer.Add_Elapsed({
                Write-Verbose 'LAUNCH: Safety timer elapsed' -ForegroundColor DarkGray
                if (Get-Command -Name Stop-ClientLaunch -ErrorAction SilentlyContinue) { Stop-ClientLaunch }
            })

        $safetyTimer.Start()
        $global:DashboardConfig.Resources.Timers['launchSafetyTimer'] = $safetyTimer

        $asyncResult = $launchPS.BeginInvoke()
        if ($null -eq $asyncResult) { throw 'Failed to start async operation' }

        $global:LaunchResources.AsyncResult = $asyncResult

        if ($global:DashboardConfig.State.LaunchActive) {
            $global:DashboardConfig.UI.Launch.FlatStyle = 'Popup'
            $global:DashboardConfig.UI.Launch.Text = "Cancel Launch"
        }

        Write-Verbose 'LAUNCH: Launch operation started' -ForegroundColor Green
    }
    catch
    {
        Write-Verbose "LAUNCH: Error in launch setup: $_" -ForegroundColor Red
        Stop-ClientLaunch
    }
}

function Invoke-SavedLaunchSequence
{
    Write-Verbose "LAUNCH: Initiating Smart Saved Launch Sequence..." -ForegroundColor Cyan

    if (-not $global:DashboardConfig.Config['SavedLaunchConfig'] -or $global:DashboardConfig.Config['SavedLaunchConfig'].Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No saved configuration found.", "Error", "OK", "Warning")
        return
    }
    if ($global:DashboardConfig.State.LaunchActive) {
        [System.Windows.Forms.MessageBox]::Show("Launch in progress.", "Busy", "OK", "Warning")
        return
    }

    $global:LaunchCancellation.IsCancelled = $false
    $global:DashboardConfig.State.LaunchActive = $true
    $global:DashboardConfig.State.SequenceActive = $true

    # --- 1. Calculate Requirements ---
    $requirements = [ordered]@{}
    $configSection = $global:DashboardConfig.Config['SavedLaunchConfig']
    
    foreach ($key in $configSection.Keys) {
        $parts = $configSection[$key] -split ',', 5
        if ($parts.Length -ge 4) {
            $p = $parts[1]
            $loginNeeded = $false
            if ($parts.Length -eq 5) {
                $loginNeeded = ($parts[4] -eq '1')
            } else {
                if ($parts[3] -match ' - ') { $loginNeeded = $true }
            }
            if (-not $requirements.Contains($p)) {
                $requirements[$p] = @{ Login = 0; NoLogin = 0 }
            }
            if ($loginNeeded) { $requirements[$p].Login++ }
            else { $requirements[$p].NoLogin++ }
        }
    }

    # --- 2. Check Current State ---
    $currentState = @{}
    $grid = $global:DashboardConfig.UI.DataGridFiller

    foreach ($row in $grid.Rows) {
        $title = $row.Cells[1].Value.ToString()
        $p = "Default"
        $cleanTitle = $title
        if ($title -match '^\[([^\]]+)\](.*)') {
            $p = $matches[1]
            $cleanTitle = $matches[2]
        }
        $isLoggedIn = ($cleanTitle -match ' - ')
        if (-not $currentState.Contains($p)) {
            $currentState[$p] = @{ Login = 0; NoLogin = 0 }
        }
        if ($isLoggedIn) { $currentState[$p].Login++ }
        else { $currentState[$p].NoLogin++ }
    }

    # --- 3. Build Queue ---
    $launchQueue = @()
    $loginTargets = @{}

    foreach ($p in $requirements.Keys) {
        $req = $requirements[$p]
        $cur = if ($currentState.Contains($p)) { $currentState[$p] } else { @{ Login=0; NoLogin=0 } }
        $reqTotal = $req.Login + $req.NoLogin
        $curTotal = $cur.Login + $cur.NoLogin
        $totalToLaunch = [Math]::Max(0, ($reqTotal - $curTotal))

        if ($totalToLaunch -gt 0) {
            $launchQueue += @{ Profile = $p; Count = $totalToLaunch }
            $missingLoginRaw = [Math]::Max(0, ($req.Login - $cur.Login))
            $countToLogin = [Math]::Min($totalToLaunch, $missingLoginRaw)
            $loginTargets[$p] = $countToLogin
            Write-Verbose "LAUNCH PLAN [$p]: Launching $totalToLaunch. (Will Auto-Login: $countToLogin)" -ForegroundColor DarkGray
        }
    }

    $existingPIDs = @($grid.Rows | ForEach-Object { if ($_.Tag -and $_.Tag.Id) { $_.Tag.Id } })
    $global:DashboardConfig.Resources['RestoreSnapshotPIDs'] = $existingPIDs
    $global:DashboardConfig.Resources['LoginTargets'] = $loginTargets

    if ($launchQueue.Count -eq 0) {
        Write-Verbose "LAUNCH SEQUENCE: State matches saved config. No action." -ForegroundColor Green
        [System.Windows.Forms.MessageBox]::Show("All clients match the saved configuration.", "Done", "OK", "Information")
        $global:DashboardConfig.State.LaunchActive = $false
        $global:DashboardConfig.State.SequenceActive = $false
        return
    }

    $global:DashboardConfig.Resources['LaunchQueue'] = $launchQueue
    $global:DashboardConfig.Resources['LaunchQueueIndex'] = 0

    if ($global:DashboardConfig.Resources.Timers.Contains('SequenceTimer')) {
        $global:DashboardConfig.Resources.Timers['SequenceTimer'].Stop()
        $global:DashboardConfig.Resources.Timers['SequenceTimer'].Dispose()
        $global:DashboardConfig.Resources.Timers.Remove('SequenceTimer')
    }

    $global:DashboardConfig.UI.Launch.FlatStyle = 'Popup'
    $global:DashboardConfig.UI.Launch.Text = "Cancel Launch"

    $seqTimer = New-Object System.Windows.Forms.Timer
    $seqTimer.Interval = 1000

    $seqTimer.Add_Tick({
        if ($global:LaunchCancellation.IsCancelled) {
             $this.Stop(); $this.Dispose()
             return
        }
        if ($global:LaunchResources) { return }

        $queue = $global:DashboardConfig.Resources['LaunchQueue']
        $idx = $global:DashboardConfig.Resources['LaunchQueueIndex']

        if ($idx -lt $queue.Count) {
            $item = $queue[$idx]
            $global:DashboardConfig.Resources['LaunchQueueIndex'] = $idx + 1

            if ($item.Profile -eq "Default") {
                Start-ClientLaunch -ClientAddCount $item.Count -FromSequence
            } else {
                Start-ClientLaunch -ProfileNameOverride $item.Profile -ClientAddCount $item.Count -FromSequence
            }
        }
        else {
            $this.Stop(); $this.Dispose()
            $global:DashboardConfig.Resources.Timers.Remove('SequenceTimer')

            Write-Verbose "LAUNCH SEQUENCE: Launches complete. Waiting 5s for clients to appear..." -ForegroundColor Cyan

            $loginWaitTimer = New-Object System.Windows.Forms.Timer
            $loginWaitTimer.Interval = 5000
            $loginWaitTimer.Tag = "OneShot"
            $loginWaitTimer.Add_Tick({
                $this.Stop(); $this.Dispose()
                $global:DashboardConfig.Resources.Timers.Remove('LoginWaitTimer')
                if ($global:LaunchCancellation.IsCancelled) { return }

                $oldPIDs = $global:DashboardConfig.Resources['RestoreSnapshotPIDs']
                $targets = $global:DashboardConfig.Resources['LoginTargets']
                $grid = $global:DashboardConfig.UI.DataGridFiller

                # A. Parse Config (Including Title)
                $savedSlots = [System.Collections.Generic.List[PSObject]]::new()
                if ($global:DashboardConfig.Config['SavedLaunchConfig']) {
                    $configSection = $global:DashboardConfig.Config['SavedLaunchConfig']
                    
                    foreach ($key in $configSection.Keys) {
                        $val = $configSection[$key]
                        $parts = $val -split ',', 5
                        
                        if ($parts.Length -ge 3) {
                            $gPos = [int]$parts[0].Trim()
                            $prof = $parts[1].Trim()
                            $acc  = [int]$parts[2].Trim()
                            $titl = if ($parts.Length -ge 4) { $parts[3].Trim() } else { "" }
                            
                            $savedSlots.Add([PSCustomObject]@{
                                GridPos    = $gPos
                                Profile    = $prof
                                AccountID  = $acc
                                Title      = $titl
                                AssignedRow = $null
                            })
                        }
                    }
                }
                $savedSlots.Sort({ $args[0].GridPos - $args[1].GridPos })

                # B. Separate Rows
                $existingRows = [System.Collections.Generic.List[Object]]::new()
                $newRows = [System.Collections.Generic.List[Object]]::new()
                $sortedGridRows = $grid.Rows | Sort-Object Index

                foreach ($row in $sortedGridRows) {
                    if ($row.Tag -and $row.Tag.Id) {
                        if ($oldPIDs -contains $row.Tag.Id) { $existingRows.Add($row) } 
                        else { $newRows.Add($row) }
                    }
                }

                # C. Consume Slots for EXISTING ROWS (Strict Title Match First)
                foreach ($row in $existingRows) {
                    $rowTitle = $row.Cells[1].Value.ToString()
                    
                    # 1. Try Exact Title Match
                    foreach ($slot in $savedSlots) {
                        if ($null -eq $slot.AssignedRow -and $slot.Title -eq $rowTitle) {
                            $slot.AssignedRow = $row
                            # Write-Verbose "MATCH (Title): Row '$rowTitle' matched to Slot Acc:$($slot.AccountID)"
                            break
                        }
                    }
                    
                    # 2. If no title match (unlikely for running clients), fall back to loose profile match (only if title is unassigned)
                    if ($null -eq ($savedSlots | Where-Object { $_.AssignedRow -eq $row })) {
                        $pName = "Default"
                        if ($rowTitle -match '^\[([^\]]+)\]') { $pName = $matches[1] }
                        foreach ($slot in $savedSlots) {
                            if ($null -eq $slot.AssignedRow -and $slot.Profile -eq $pName) {
                                $slot.AssignedRow = $row
                                break
                            }
                        }
                    }
                }

                # D. Map Remaining Slots to NEW ROWS
                $finalLoginList = @()
                $loginCounts = @{}

                foreach ($row in $newRows) {
                    $rowTitle = $row.Cells[1].Value.ToString()
                    $pName = "Default"
                    if ($rowTitle -match '^\[([^\]]+)\]') { $pName = $matches[1] }
                    
                    $matchedSlot = $null
                    
                    # 1. Try Title Match (If window title is already set correctly)
                    foreach ($slot in $savedSlots) {
                        if ($null -eq $slot.AssignedRow -and $slot.Title -eq $rowTitle) {
                            $matchedSlot = $slot
                            break
                        }
                    }

                    # 2. Try GridPos Match (Strict)
                    if ($null -eq $matchedSlot) {
                         $visualGridPos = $row.Cells[0].Value -as [int]
                         foreach ($slot in $savedSlots) {
                            if ($null -eq $slot.AssignedRow -and $slot.Profile -eq $pName -and $slot.GridPos -eq $visualGridPos) {
                                $matchedSlot = $slot
                                break
                            }
                        }
                    }

                    # 3. Fallback: Any available slot for this profile
                    if ($null -eq $matchedSlot) {
                        foreach ($slot in $savedSlots) {
                            if ($null -eq $slot.AssignedRow -and $slot.Profile -eq $pName) {
                                $matchedSlot = $slot
                                break
                            }
                        }
                    }

                    if ($matchedSlot) {
                        $matchedSlot.AssignedRow = $row 

                        if ($targets.Contains($pName)) {
                            if (-not $loginCounts.Contains($pName)) { $loginCounts[$pName] = 0 }
                            if ($loginCounts[$pName] -lt $targets[$pName]) {
                                
                                # --- CREATE WRAPPER ---
                                $wrapper = [PSCustomObject]@{
                                    Row = $row
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
                $global:DashboardConfig.UI.Launch.Text = "Launch " + [char]0x25BE

                if ($finalLoginList.Count -gt 0) {
                    $details = $finalLoginList | ForEach-Object {
                        $pName = "Unknown"
                        if ($_.Row.Cells[1].Value -match '^\[([^\]]+)\]') { $pName = $matches[1] }
                        "[$pName (Row:$($_.Row.Cells[0].Value)) -> Enforce Acc:$($_.OverrideAccountID)]"
                    }
                    $detailString = $details -join ', '
                    
                    Write-Verbose "LAUNCH SEQUENCE: Auto-logging into $($finalLoginList.Count) specific clients." -ForegroundColor Green
                    Write-Verbose "LAUNCH PLAN MAP: $detailString" -ForegroundColor Cyan

                    if (Get-Command LoginSelectedRow -ErrorAction SilentlyContinue) {
                        LoginSelectedRow -RowInput $finalLoginList
                    }
                    $global:DashboardConfig.State.LaunchActive = $false
                    $global:DashboardConfig.State.SequenceActive = $false
                } else {
                    Write-Verbose "LAUNCH SEQUENCE: No clients require login scripts." -ForegroundColor Yellow
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

function Stop-ClientLaunch
{
    [CmdletBinding()]
    param([switch]$StepCompleted)

    if ($StepCompleted) {
        Write-Verbose 'LAUNCH: Cleaning up completed launch step...' -ForegroundColor DarkGray
    } else {
        Write-Verbose 'LAUNCH: Cancelling all launch operations...' -ForegroundColor Yellow
        $global:LaunchCancellation.IsCancelled = $true
    }

    if (-not ($StepCompleted -and $global:DashboardConfig.State.SequenceActive)) {
        $global:DashboardConfig.State.LaunchActive = $false
        $global:DashboardConfig.State.SequenceActive = $false
    }

    try
    {
        if (-not $StepCompleted) {
            try {
                if ($global:DashboardConfig.Resources.Timers.Contains('SequenceTimer')) {
                    $global:DashboardConfig.Resources.Timers['SequenceTimer'].Stop()
                    $global:DashboardConfig.Resources.Timers['SequenceTimer'].Dispose()
                    $global:DashboardConfig.Resources.Timers.Remove('SequenceTimer')
                    Write-Verbose 'LAUNCH: Queue timer stopped.' -ForegroundColor DarkGray
                }
                if ($global:DashboardConfig.Resources.Timers.Contains('LoginWaitTimer')) {
                    $global:DashboardConfig.Resources.Timers['LoginWaitTimer'].Stop()
                    $global:DashboardConfig.Resources.Timers['LoginWaitTimer'].Dispose()
                    $global:DashboardConfig.Resources.Timers.Remove('LoginWaitTimer')
                    Write-Verbose 'LAUNCH: Login wait timer stopped.' -ForegroundColor DarkGray
                }
            } catch {}
            if ($global:DashboardConfig.Resources.Contains('LaunchQueue')) {
                $global:DashboardConfig.Resources['LaunchQueue'] = @()
                $global:DashboardConfig.Resources['LaunchQueueIndex'] = 0
            }
        }

        if ($null -eq $global:LaunchResources)
        {
            $global:DashboardConfig.UI.Launch.FlatStyle = 'Flat'
            $global:DashboardConfig.UI.Launch.Text = "Launch " + [char]0x25BE
            return
        }

        if ($null -ne $global:LaunchResources.EventSubscriptionId)
        {
            try { Unregister-Event -SourceIdentifier $global:LaunchResources.EventSubscriptionId -Force -ErrorAction SilentlyContinue } catch {}
        }

        if ($null -ne $global:LaunchResources.EventSubscriber)
        {
            try {
                if ($null -ne $global:LaunchResources.EventSubscriptionId) {
                    Unregister-Event -SourceIdentifier $global:LaunchResources.EventSubscriptionId -ErrorAction SilentlyContinue
                }
            } catch {}
        }

        $psInstance = $global:LaunchResources.PowerShellInstance
        $runspace = $global:LaunchResources.Runspace
        $asyncResult = $global:LaunchResources.AsyncResult

        if ($null -ne $psInstance)
        {
            try {
                if ($psInstance.InvocationStateInfo.State -eq 'Running') {
                    $psInstance.Stop()
                }
                # Wait for the pipeline to finish stopping by calling EndInvoke.
                # This will throw an exception if the pipeline was stopped, which is expected.
                if ($asyncResult -and -not $asyncResult.IsCompleted) {
                    $psInstance.EndInvoke($asyncResult)
                }
            } catch {
                Write-Verbose "LAUNCH: Caught expected exception during pipeline stop: $($_.Exception.Message)" -ForegroundColor DarkGray
            } finally {
                try { $psInstance.Dispose() } catch {}
            }
        }

        if ($null -ne $runspace)
        {
            try {
                $runspace.Close()
                $runspace.Dispose()
            } catch {}
        }

        if ($global:DashboardConfig.Resources.Timers.Contains('launchSafetyTimer'))
        {
            try {
                $global:DashboardConfig.Resources.Timers['launchSafetyTimer'].Stop()
                $global:DashboardConfig.Resources.Timers['launchSafetyTimer'].Dispose()
                $global:DashboardConfig.Resources.Timers.Remove('launchSafetyTimer')
            } catch {}
        }

        $global:LaunchResources = $null

        [System.GC]::Collect()

        Write-Verbose 'LAUNCH: Launch cleanup completed' -ForegroundColor Green

        $global:DashboardConfig.UI.Launch.FlatStyle = 'Flat'
        $global:DashboardConfig.UI.Launch.Text = "Launch " + [char]0x25BE
    }
    catch
    {
        Write-Verbose "LAUNCH: Error during launch cleanup: $_" -ForegroundColor Red
        $global:DashboardConfig.State.LaunchActive = $false
        $global:DashboardConfig.UI.Launch.FlatStyle = 'Flat'
        $global:DashboardConfig.UI.Launch.Text = "Launch " + [char]0x25BE
    }
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion
