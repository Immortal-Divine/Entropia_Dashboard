<# login.psm1 #>

#region Configuration

$global:LoginCancellation = [hashtable]::Synchronized(@{
    IsCancelled = $false
})
$TOTAL_STEPS_PER_CLIENT = 13
$script:ScriptInitiatedMove = $false

if (-not $global:LoginResources) {
    $global:LoginResources = @{
        PowerShellInstance  = $null
        Runspace            = $null
        EventSubscriptionId = $null
        EventSubscriber     = $null
        ProgressSubscription= $null
        AsyncResult         = $null
        IsStopping          = $false
        IsMouseHookActive   = $false
    }
}
#endregion

#region Helper Functions

function Update-Progress {
    param(
        [Parameter(Mandatory=$true)]
        $ProgressBarObject,
        [string]$Text,
        [int]$Percent
    )

    if (-not $ProgressBarObject -or $ProgressBarObject.IsDisposed) { return }

    try {
        if ($ProgressBarObject.InvokeRequired) {
            $ProgressBarObject.BeginInvoke([Action]{
                Update-Progress -ProgressBarObject $ProgressBarObject -Text $Text -Percent $Percent
            })
        } else {
            $ProgressBarObject.CustomText = $Text
            if ($Percent -ge 0 -and $Percent -le 100) {
                $ProgressBarObject.Value = $Percent
            }
        }
    } catch {}
}

function Global:Get-ClientLogPath {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.DataGridViewRow]$Row
    )
    $determinedLogBaseFolder = ''
    $process = $Row.Tag
    $profileName = ''
    $processExeBaseFolder = ''
    $processTitle = $Row.Cells[1].Value
    $null = $entryNum; $entryNum = $Row.Cells[0].Value

    if ($processTitle -match '^\[([^\]]+)\]') { $profileName = $Matches[1] }

    if ([string]::IsNullOrEmpty($profileName) -and $process -and $process.Id) {
        $profileName = Get-ProcessProfile -Process $process
    }

    if ($process -and $process.Id) {
        $processExePath = [Custom.Native]::GetProcessPathById($process.Id)
        if (-not [string]::IsNullOrEmpty($processExePath)) {
            $processExeBaseFolder = Split-Path -Parent -Path $processExePath
        }
    }

    if ([string]::IsNullOrEmpty($determinedLogBaseFolder) -and
        $global:DashboardConfig -and $global:DashboardConfig.Config -and
        $global:DashboardConfig.Config.Contains('Profiles'))
    {
        $foundProfileKey = $null
        foreach ($key in $global:DashboardConfig.Config.Profiles.Keys) {
            if ($key.ToLowerInvariant() -eq $profileName.ToLowerInvariant()) {
                $foundProfileKey = $key
                break
            }
        }
        if ($foundProfileKey) {
            $profilePath = $global:DashboardConfig.Config.Profiles[$foundProfileKey]
            if (-not [string]::IsNullOrEmpty($profilePath)) {
                $determinedLogBaseFolder = $profilePath
            }
        }
    }

    if ([string]::IsNullOrEmpty($determinedLogBaseFolder) -and -not [string]::IsNullOrEmpty($processExeBaseFolder)) {
        $determinedLogBaseFolder = $processExeBaseFolder
    }

    if ([string]::IsNullOrEmpty($determinedLogBaseFolder)) {
        $launcherPathConfig = $global:DashboardConfig.Config['LauncherPath']
        if ($launcherPathConfig -and $launcherPathConfig.ContainsKey('LauncherPath')) {
            $launcherPath = $launcherPathConfig['LauncherPath']
            if (-not [string]::IsNullOrEmpty($launcherPath)) {
                $determinedLogBaseFolder = Split-Path -Path $launcherPath -Parent
            }
        }
    }

    if ([string]::IsNullOrEmpty($determinedLogBaseFolder)) {
        $determinedLogBaseFolder = Join-Path -Path $env:APPDATA -ChildPath "Entropia_Dashboard"
    }

    $actualLogPath = Join-Path -Path $determinedLogBaseFolder -ChildPath "Log\network_$(Get-Date -Format 'yyyyMMdd').log"
    return $actualLogPath
}

#endregion

#region Core Function

function LoginSelectedRow {
    param(
        [Parameter(Mandatory=$false)]
        $RowInput,
        [Parameter(Mandatory=$false)]
        [IntPtr]$WindowHandle = [IntPtr]::Zero,
        [string]$LogFilePath
    )

    if ($global:DashboardConfig.State['LoginActive']) { return }
    $global:DashboardConfig.State['LoginActive'] = $true

    $rowsToProcess = @()
    $rawSelection = @()

    if ($RowInput) {
        $rawSelection += $RowInput
    } elseif ($global:DashboardConfig.UI.DataGridFiller.SelectedRows.Count -gt 0) {
        $rawSelection = $global:DashboardConfig.UI.DataGridFiller.SelectedRows | Sort-Object Index
    } else {
        $global:DashboardConfig.State['LoginActive'] = $false
        return
    }

    # --- INPUT PROCESSING ---
    $enrichedData = $rawSelection | ForEach-Object {
        $item = $_
        $actualRow = $null
        $entryNum = 0

        # Check for our Special Wrapper
        if ($item.PSTypeNames -contains 'LoginOverrideWrapper') {
            # OVERRIDE DETECTED: Use the Forced Account ID (e.g., 4)
            $actualRow = $item.Row
            $entryNum = $item.OverrideAccountID -as [int]
        } 
        else {
            # Standard Manual Selection: Use the Grid Cell Value
            $actualRow = $item
            $entryNum = $actualRow.Cells[0].Value -as [int]
        }

        if ($null -ne $actualRow) {
            $title = $actualRow.Cells[1].Value
            $profileName = "Default"
            if ($title -match '^\[([^\]]+)\]') { $profileName = $Matches[1] }
            
            [PSCustomObject]@{ 
                OriginalRow = $actualRow; 
                EntryNum = $entryNum; # This holds the Correct ID (4)
                Profile = $profileName 
            }
        }
    }
    # ------------------------

    $profilePriorityMap = [ordered]@{}
    $priorityCounter = 0
    foreach ($item in $enrichedData) {
        $pKey = $item.Profile.ToString()
        if (-not $profilePriorityMap.Contains($pKey)) { $profilePriorityMap[$pKey] = $priorityCounter++ }
    }

    $sortedData = $enrichedData | Sort-Object `
        @{Expression={$profilePriorityMap[$_.Profile.ToString()]}; Ascending=$true}, `
        @{Expression={$_.EntryNum}; Ascending=$true}

    $jobs = @()
    foreach ($dataItem in $sortedData) {
        $row = $dataItem.OriginalRow
        # The script will use this $entryNum (4) for the mouse click math
        $entryNum = [int]$dataItem.EntryNum 
        
        $process = $row.Tag
        $processTitle = $row.Cells[1].Value

        $profileName = "Default"
        if ($global:DashboardConfig.Config['LoginConfig']) {
            $allProfileNames = $global:DashboardConfig.Config['LoginConfig'].Keys | Where-Object { $_ -ne "Default" }
            foreach ($knownProfile in $allProfileNames) {
                if ($processTitle -match "\[$([regex]::Escape($knownProfile))\]") {
                    $profileName = $knownProfile
                    break
                }
            }
        }

        $actualLogPath = $LogFilePath
        if ([string]::IsNullOrEmpty($actualLogPath)) {
            $actualLogPath = Get-ClientLogPath -Row $row
        }
        if ($null -eq $actualLogPath) { $actualLogPath = "" }

        $clientLoginConfig = @{}
        if ($global:DashboardConfig.Config['LoginConfig'] -and $global:DashboardConfig.Config['LoginConfig'][$profileName]) {
             $clientLoginConfig = $global:DashboardConfig.Config['LoginConfig'][$profileName]
        }

        $thisJobHandle = [IntPtr]::Zero
        if ($WindowHandle -ne [IntPtr]::Zero) {
            if ($RowInput -eq $row -or ($RowInput.Row -eq $row)) {
                $thisJobHandle = $WindowHandle
            }
        }

        $jobs += [PSCustomObject]@{
            EntryNumber    = $entryNum
            ProcessId      = if ($process) { $process.Id } else { 0 }
            ExplicitHandle = $thisJobHandle
            LogPath        = $actualLogPath
            Config         = $clientLoginConfig
        }
    }

    $pb = $global:DashboardConfig.UI.GlobalProgressBar
    if ($pb) {
        $pb.Visible = $true
        $pb.Value = 0
        $pb.CustomText = "Starting Login Process..."
    }
    $global:DashboardConfig.UI.LoginButton.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $global:DashboardConfig.UI.LoginButton.Text = "Running..."

    $global:LoginCancellation.IsCancelled = $false
    $global:LoginResources['IsStopping'] = $false
    $script:ScriptInitiatedMove = $false

    $hookCallback = [Custom.MouseHookManager+HookProc] {
        param($nCode, $wParam, $lParam)
        if ($nCode -ge 0 -and $wParam -eq 0x0200) {
            if (-not $script:ScriptInitiatedMove) {
                if (-not $global:LoginCancellation.IsCancelled) {
                    $global:LoginCancellation.IsCancelled = $true
                    $pbLocal = $global:DashboardConfig.UI.GlobalProgressBar
                    if ($pbLocal -and $pbLocal.InvokeRequired) {
                        $pbLocal.BeginInvoke([Action]{ $pbLocal.CustomText = "Cancelling..." })
                    }
                }
            }
        }
        return [Custom.MouseHookManager]::CallNextHookEx([Custom.MouseHookManager]::HookId, $nCode, $wParam, $lParam)
    }
    if (-not $global:LoginResources.IsMouseHookActive) {
        try {
            [Custom.MouseHookManager]::Start($hookCallback)
            $global:LoginResources.IsMouseHookActive = $true
            Write-Verbose "LOGIN: Mouse hook started." -ForegroundColor Green
        } catch {
            Write-Warning "LOGIN: Failed to start mouse hook: $_" -ForegroundColor Yellow
        }
    }

    try {
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        
        $varsToInject = @{
            'DashboardConfig'         = $global:DashboardConfig
            'TOTAL_STEPS_PER_CLIENT'  = $TOTAL_STEPS_PER_CLIENT
            'CancellationContext'     = $global:LoginCancellation
            'LoginResources'          = $global:LoginResources
        }

        foreach ($entry in $varsToInject.GetEnumerator()) {
            $sessionVar = New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry(
                $entry.Key,
                $entry.Value,
                "Injected from main thread",
                [System.Management.Automation.ScopedItemOptions]::None
            )
            $iss.Variables.Add($sessionVar)
        }

        $customTypesToLoad = @('Custom.Native', 'Custom.Ftool', 'Custom.MouseHookManager')
        $assembliesToAdd = [System.Collections.Generic.HashSet[string]]::new()

        foreach ($typeName in $customTypesToLoad) {
            $type = $typeName -as [Type]
            if ($type -and -not [string]::IsNullOrEmpty($type.Assembly.Location)) {
                $assembliesToAdd.Add($type.Assembly.Location) | Out-Null
            }
        }

        if ($assembliesToAdd.Count -eq 0) {
            $loadedAssemblies = [AppDomain]::CurrentDomain.GetAssemblies()
            foreach ($asm in $loadedAssemblies) {
                try {
                    if ($asm.IsDynamic) { continue }
                    if ([string]::IsNullOrWhiteSpace($asm.Location)) { continue }

                    foreach ($t in $asm.GetExportedTypes()) {
                        if ($customTypesToLoad -contains $t.FullName) {
                            $assembliesToAdd.Add($asm.Location) | Out-Null
                            break
                        }
                    }
                } catch {}
            }
        }

        foreach ($loc in $assembliesToAdd) {
            if (-not [string]::IsNullOrWhiteSpace($loc)) {
                if (Test-Path -Path $loc -ErrorAction SilentlyContinue) {
                    Write-Verbose "LOGIN: Injecting assembly from $loc"
                    $iss.Assemblies.Add([System.Management.Automation.Runspaces.SessionStateAssemblyEntry]::new($loc))
                }
            }
        }

        $localRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 1, $iss, $Host)
        $localRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
        $localRunspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $localRunspace.Open()

        $loginPS = [PowerShell]::Create()
        $loginPS.RunspacePool = $localRunspace

        $global:LoginResources = @{
            PowerShellInstance  = $loginPS
            Runspace            = $localRunspace
            EventSubscriptionId = $null
            EventSubscriber     = $null
            ProgressSubscription= $null
            AsyncResult         = $null
            IsStopping          = $false
            IsMouseHookActive   = $global:LoginResources.IsMouseHookActive
        }

        # Runspace Script Body (Logic unchanged, reliant on $entryNumber which is now 4)
        $loginPS.AddScript({
            param($Jobs, $TotalStepsPerClient, $GlobalOptions, $LoginConfig, $CancellationContext, $State)

            if (-not ('Custom.Native' -as [Type])) {
                throw "CRITICAL ERROR: [Custom.Native] type is missing in Background Runspace. Automation cannot proceed."
            }
            
            $Global:CurrentActiveProcessId = 0
            $Global:CurrentActiveWindowHandle = [IntPtr]::Zero

            function Global:CheckCancel {
                if ($CancellationContext.IsCancelled) { throw "LoginCancelled" }
            }

            function Global:SleepWithCancel {
                param([int]$Milliseconds)
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while ($sw.Elapsed.TotalMilliseconds -lt $Milliseconds) {
                    if ($CancellationContext.IsCancelled) { throw "LoginCancelled" }
                    Start-Sleep -Milliseconds 100
                }
            }

            function Global:EnsureWindowResponsive {
                if ($Global:CurrentActiveProcessId -eq 0) { return }
                $proc = Get-Process -Id $Global:CurrentActiveProcessId -ErrorAction SilentlyContinue
                if (-not $proc) { throw "Process with ID $Global:CurrentActiveProcessId has terminated." }

                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while (-not $proc.Responding) {
                    if ($sw.Elapsed.TotalSeconds -gt 15) { throw "Timeout: Window is Not Responding (Hung)." }
                    CheckCancel
                    $proc.Refresh()
                    Start-Sleep -Milliseconds 10
                }
                try { $proc.WaitForInputIdle(20) | Out-Null } catch {}
            }

            function Global:Invoke-MouseClick {
                param([int]$X, [int]$Y)
                CheckCancel; EnsureWindowResponsive; CheckCancel
                $script:ScriptInitiatedMove = $true
                try {
                    [Custom.Native]::SetCursorPos($X, $Y); Start-Sleep -Milliseconds 20
                    [Custom.Native]::SetCursorPos($X, $Y); Start-Sleep -Milliseconds 30
                    $MOUSEEVENTF_LEFTDOWN = 0x0002; $MOUSEEVENTF_LEFTUP = 0x0004
                    [Custom.Native]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0); Start-Sleep -Milliseconds 50
                    [Custom.Native]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0); Start-Sleep -Milliseconds 50
                } catch { Write-Verbose "Mouse click failed: $_" } finally { Start-Sleep -Milliseconds 50; $script:ScriptInitiatedMove = $false; CheckCancel }
            }

            function Global:Invoke-KeyPress {
                param([int]$VirtualKeyCode)
                CheckCancel; EnsureWindowResponsive; CheckCancel
                try {
                    $hWnd = $Global:CurrentActiveWindowHandle
                    if ($hWnd -eq [IntPtr]::Zero) { $hWnd = [Custom.Native]::GetForegroundWindow() }
                    [Custom.Ftool]::fnPostMessage($hWnd, 0x0100, $VirtualKeyCode, 0); SleepWithCancel -Milliseconds 25
                    [Custom.Ftool]::fnPostMessage($hWnd, 0x0101, $VirtualKeyCode, 0); SleepWithCancel -Milliseconds 25
                } catch { if ($_.Exception.Message -eq "LoginCancelled") { throw }; throw "KeyPress Failed: $($_.Exception.Message)" }
            }

            function Global:Set-WindowForeground {
                param([System.Diagnostics.Process]$Process, [IntPtr]$ExplicitHwnd)
                CheckCancel
                $targetHwnd = if ($ExplicitHwnd -ne [IntPtr]::Zero) { $ExplicitHwnd } else { $Process.MainWindowHandle }
                if ($targetHwnd -eq [IntPtr]::Zero) { return $false }
                $script:ScriptInitiatedMove = $true
                try { [Custom.Native]::BringToFront($targetHwnd); Start-Sleep -Milliseconds 100 } finally { $script:ScriptInitiatedMove = $false; CheckCancel }
                return $true
            }

            function Global:ParseCoordinates {
                param([string]$ConfigString)
                if ([string]::IsNullOrWhiteSpace($ConfigString) -or $ConfigString -notmatch ',') { return $null }
                $parts = $ConfigString.Split(',')
                if ($parts.Count -eq 2) { return @{ X = [int]$parts[0].Trim(); Y = [int]$parts[1].Trim() } }
                return $null
            }

            function Global:Wait-ForLogEntry {
                param($LogPath, $SearchStrings, $TimeoutSeconds)
                if ($TimeoutSeconds -le 0) { $TimeoutSeconds = 60 }
                $waitSw = [System.Diagnostics.Stopwatch]::StartNew()
                while ($waitSw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
                    CheckCancel
                    if (-not [string]::IsNullOrWhiteSpace($LogPath) -and (Test-Path $LogPath)) {
                        try {
                            $lastLines = Get-Content $LogPath -Tail 20 -ErrorAction SilentlyContinue
                            if ($lastLines) { foreach ($str in $SearchStrings) { if ($lastLines -match [regex]::Escape($str)) { return $true } } }
                        } catch {}
                    }
                    SleepWithCancel -Milliseconds 100
                }
                return $false
            }

            function Global:Wait-UntilWorldLoaded {
                param($LogPath, $Config, $TotalSteps, $CurrentStep)
                $threshold = 3; $searchStr = "13 - CACHE_ACK_JOIN"
                if ($Config['WorldLoadLogThreshold']) { $threshold = [int]$Config['WorldLoadLogThreshold'] }
                if ($Config['WorldLoadLogEntry']) { $searchStr = $Config['WorldLoadLogEntry'] }
                $foundCount = 0; $timeout = New-TimeSpan -Minutes 2; $sw = [System.Diagnostics.Stopwatch]::StartNew()
                Write-Progress -Activity "Login" -Status "Waiting for World Load..." -PercentComplete ([int](($CurrentStep / $TotalSteps) * 100))
                while ($foundCount -lt $threshold) {
                    if ($sw.Elapsed -gt $timeout) { throw "World load timeout" }
                    CheckCancel
                    if (-not [string]::IsNullOrWhiteSpace($LogPath) -and (Test-Path $LogPath)) {
                        try {
                            $lines = Get-Content $LogPath -Tail 50 -ErrorAction SilentlyContinue
                            if ($lines) { $foundCount = ($lines | Select-String -SimpleMatch $searchStr).Count }
                        } catch {}
                    }
                    SleepWithCancel -Milliseconds 100
                }
            }

            function Global:Write-LogWithRetry {
                param([string]$FilePath, [string]$Value)
                if ([string]::IsNullOrWhiteSpace($FilePath)) { return }
                for ($i=0; $i -lt 5; $i++) { try { Set-Content -Path $FilePath -Value $Value -Force -ErrorAction Stop; return } catch { Start-Sleep -Milliseconds 50 } }
            }

            $totalClients = $Jobs.Count
            $totalGlobalSteps = $totalClients * $TotalStepsPerClient
            $currentClientIndex = 0

            if ($State -and -not $State['LoginGracePids']) { $State['LoginGracePids'] = [System.Collections.Hashtable]::Synchronized(@{}) }

            foreach ($job in $Jobs) {
                CheckCancel; EnsureWindowResponsive; CheckCancel
                $currentClientIndex++
                $entryNumber = $job.EntryNumber
                $processId = $job.ProcessId
                $explicitHandle = if ($job.ExplicitHandle -and $job.ExplicitHandle -ne 0) { $job.ExplicitHandle } else { [IntPtr]::Zero }
                $Global:CurrentActiveWindowHandle = $explicitHandle
                $Global:CurrentActiveProcessId = $processId

                if ($processId -gt 0) { if ($State -and $State['ActiveLoginPids']) { [void]$State['ActiveLoginPids'].Add($processId) } }
                
                $logPath = $job.LogPath
                $config = $job.Config
                $stepBase = ($currentClientIndex - 1) * $TotalStepsPerClient
                $currentStep = $stepBase

                $currentStep++; Write-Progress -Activity "Login" -Status "Client $entryNumber Starting" -PercentComplete ([int](($currentStep / $totalGlobalSteps) * 100))
                
                if ($processId -eq 0) { throw "Process ID not found for Client $entryNumber" }
                $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
                if (-not $process) { throw "Process $processId is gone." }

                $serverID="1"; $channelID="1"; $charSlot="1"; $startCollector="No"
                $settingKey = "Client${entryNumber}_Settings"
                if ($config[$settingKey]) {
                    $parts = $config[$settingKey] -split ','
                    if ($parts.Count -eq 4) { $serverID=$parts[0]; $channelID=$parts[1]; $charSlot=$parts[2]; $startCollector=$parts[3] }
                }

                Write-LogWithRetry -FilePath $logPath -Value ""
                $workingHwnd = if ($explicitHandle -ne [IntPtr]::Zero) { $explicitHandle } else { $process.MainWindowHandle }

                if ($workingHwnd -ne [IntPtr]::Zero) {
                    if ([Custom.Native]::IsWindowMinimized($workingHwnd)) {
                        [Custom.Native]::ShowWindow($workingHwnd, 6); SleepWithCancel -Milliseconds 100
                        [Custom.Native]::ShowWindow($workingHwnd, 9); SleepWithCancel -Milliseconds 100
                    }
                    Set-WindowForeground -Process $process -ExplicitHwnd $workingHwnd | Out-Null
                }
                CheckCancel; EnsureWindowResponsive; CheckCancel; SleepWithCancel -Milliseconds 100

                $rect = New-Object Custom.Native+RECT
                [Custom.Native]::GetWindowRect($workingHwnd, [ref]$rect)
                $x = $rect.Left + 100; $y = $rect.Top + 100
                [Custom.Native]::SetCursorPos($x, $y); SleepWithCancel -Milliseconds 10
                $MOUSEEVENTF_LEFTDOWN = 0x0002; $MOUSEEVENTF_LEFTUP = 0x0004
                [Custom.Native]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0); SleepWithCancel -Milliseconds 25
                [Custom.Native]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0); SleepWithCancel -Milliseconds 1000

                $currentStep++; Write-Progress -Activity "Login" -Status "Client $entryNumber Selecting Index" -PercentComplete ([int](($currentStep / $totalGlobalSteps) * 100))

                $firstNickCoords = ParseCoordinates $config['FirstNickCoords']
                $scrollDownCoords = ParseCoordinates $config['ScrollDownCoords']
                $scrollbaseX = if ($scrollDownCoords) { $rect.Left + $scrollDownCoords.X } else { [int](($rect.Left + $rect.Right) / 2) + 160 }
                $scrollbaseY = if ($scrollDownCoords) { $rect.Top + $scrollDownCoords.Y } else { [int](($rect.Top + $rect.Bottom) / 2) + 46 }

                if ($firstNickCoords) {
                    $baseX = $rect.Left + $firstNickCoords.X
                    $baseY = $rect.Top + $firstNickCoords.Y
                    if ($entryNumber -ge 6 -and $entryNumber -le 10) {
                        CheckCancel; EnsureWindowResponsive; CheckCancel
                        Invoke-MouseClick -X ($scrollbaseX + 5) -Y ($scrollbaseY -5); SleepWithCancel -Milliseconds 100
                        $targetY = $baseY + (($entryNumber - 6) * 18)
                        Invoke-MouseClick -X $baseX -Y $targetY; Invoke-MouseClick -X $baseX -Y $targetY
                        Invoke-MouseClick -X $baseX -Y $targetY; Invoke-MouseClick -X $baseX -Y $targetY
                    } elseif ($entryNumber -ge 1 -and $entryNumber -le 5) {
                        $targetY = $baseY + (($entryNumber - 1) * 18)
                        CheckCancel; EnsureWindowResponsive; CheckCancel
                        Invoke-MouseClick -X $baseX -Y $targetY; Invoke-MouseClick -X $baseX -Y $targetY
                        Invoke-MouseClick -X $baseX -Y $targetY; Invoke-MouseClick -X $baseX -Y $targetY
                    }
                } else {
                    $centerX = [int](($rect.Left + $rect.Right) / 2) + 25
                    $centerY = [int](($rect.Top + $rect.Bottom) / 2) + 18
                    $adjustedY = $centerY
                    if ($entryNumber -ge 6 -and $entryNumber -le 10) {
                        $yOffset = ($entryNumber - 8) * 18
                        $adjustedY = $centerY + $yOffset
                        Invoke-MouseClick -X ($centerX + 145) -Y ($centerY + 28); SleepWithCancel -Milliseconds 100
                    } elseif ($entryNumber -ge 1 -and $entryNumber -le 5) {
                        $yOffset = ($entryNumber - 3) * 18
                        $adjustedY = $centerY + $yOffset
                    } elseif ($entryNumber -ge 11) { return }
                    CheckCancel; EnsureWindowResponsive; CheckCancel
                    Invoke-MouseClick -X $centerX -Y $adjustedY; Invoke-MouseClick -X $centerX -Y $adjustedY
                    Invoke-MouseClick -X $centerX -Y $adjustedY; Invoke-MouseClick -X $centerX -Y $adjustedY
                    SleepWithCancel -Milliseconds 50
                }
                SleepWithCancel -Milliseconds 50
                
                # Server Selection
                CheckCancel; EnsureWindowResponsive; CheckCancel
                if ($config["Server${serverID}Coords"]) {
                    $coords = ParseCoordinates $config["Server${serverID}Coords"]
                    if ($coords) {
                        $currentStep++; Write-Progress -Activity "Login" -Status "Client $entryNumber Server $serverID" -PercentComplete ([int](($currentStep / $totalGlobalSteps) * 100))
                        Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
                        Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
                        SleepWithCancel -Milliseconds 50
                    }
                }
                SleepWithCancel -Milliseconds 50; CheckCancel; EnsureWindowResponsive; CheckCancel
                
                # Channel Selection
                if ($config["Channel${channelID}Coords"]) {
                    $coords = ParseCoordinates $config["Channel${channelID}Coords"]
                    if ($coords) {
                        $currentStep++; Write-Progress -Activity "Login" -Status "Client $entryNumber Channel $channelID" -PercentComplete ([int](($currentStep / $totalGlobalSteps) * 100))
                        Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
                        Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
                        SleepWithCancel -Milliseconds 50
                    }
                }
                SleepWithCancel -Milliseconds 50; CheckCancel; EnsureWindowResponsive; CheckCancel
                Invoke-KeyPress -VirtualKeyCode 0x0D

                $currentStep++
                $loginReady = Wait-ForLogEntry -LogPath $logPath -SearchStrings @("6 - LOGIN_PLAYER_LIST") -TimeoutSeconds 25
                if (-not $loginReady) { throw "CERT Login Timeout for Client $entryNumber" }
                SleepWithCancel -Milliseconds 50; CheckCancel; EnsureWindowResponsive; CheckCancel
                
                # Char Slot
                if ($config["Char${charSlot}Coords"] -and $config["Char${charSlot}Coords"] -ne '0,0') {
                    $coords = ParseCoordinates $config["Char${charSlot}Coords"]
                    if ($coords) {
                        $currentStep++; Write-Progress -Activity "Login" -Status "Client $entryNumber Char $charSlot" -PercentComplete ([int](($currentStep / $totalGlobalSteps) * 100))
                        Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
                        SleepWithCancel -Milliseconds 500
                    }
                }
                CheckCancel; EnsureWindowResponsive; CheckCancel
                Invoke-KeyPress -VirtualKeyCode 0x0D; SleepWithCancel -Milliseconds 50

                $currentStep++
                $cacheJoin = Wait-ForLogEntry -LogPath $logPath -SearchStrings @("13 - CACHE_ACK_JOIN") -TimeoutSeconds 60
                if (-not $cacheJoin) { throw "Main Login Timeout for Client $entryNumber" }

                $currentStep++; Wait-UntilWorldLoaded -LogPath $logPath -Config $config -TotalSteps $totalGlobalSteps -CurrentStep $currentStep
                CheckCancel; EnsureWindowResponsive; CheckCancel
                
                $delay = if ($config['PostLoginDelay']) { [int]$config['PostLoginDelay'] } else { 1 }
                $currentStep++; Write-Progress -Activity "Login" -Status "Client $entryNumber Optimization Delay ($delay s)" -PercentComplete ([int](($currentStep / $totalGlobalSteps) * 100))
                SleepWithCancel -Milliseconds ($delay * 1000)
                CheckCancel; EnsureWindowResponsive; CheckCancel

                # Collector
                if ($startCollector -eq "Yes" -and $config['CollectorStartCoords'] -and $config['CollectorStartCoords'] -ne '0,0') {
                    $coords = ParseCoordinates $config['CollectorStartCoords']
                    if ($coords) {
                        $currentStep++; Write-Progress -Activity "Login" -Status "Client $entryNumber Start Collector" -PercentComplete ([int](($currentStep / $totalGlobalSteps) * 100))
                        Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
                        SleepWithCancel -Milliseconds 1000
                    }
                }

                $currentStep++; Write-Progress -Activity "Login" -Status "Client $entryNumber Minimizing" -PercentComplete ([int](($currentStep / $totalGlobalSteps) * 100))
                [Custom.Native]::SendMessage($workingHwnd, 0x0112, 0xF020, 0)
                [Custom.Native]::EmptyWorkingSet($process.Handle)
            }
        }).AddArgument($jobs).AddArgument($TOTAL_STEPS_PER_CLIENT).AddArgument($global:DashboardConfig.Config['Options']).AddArgument($global:DashboardConfig.Config['LoginConfig']).AddArgument($global:LoginCancellation).AddArgument($global:DashboardConfig.State)

        $progressEvent = {
            param($s, $e)
            $uiPB = $global:DashboardConfig.UI.GlobalProgressBar
            if ($s -and $e -and $e.Index -ge 0) {
                $record = $s[$e.Index]
                if ($record -and $record.Activity -eq 'Login') {
                    $text = $record.StatusDescription; $val = $record.PercentComplete
                    if ($uiPB -and -not $uiPB.IsDisposed) {
                        if ($uiPB.InvokeRequired) { $uiPB.BeginInvoke([Action]{ Update-Progress -ProgressBarObject $uiPB -Text $text -Percent $val }) }
                        else { Update-Progress -ProgressBarObject $uiPB -Text $text -Percent $val }
                    }
                }
            }
        }
        $progressSub = Register-ObjectEvent -InputObject $loginPS.Streams.Progress -EventName DataAdded -Action $progressEvent
        $global:LoginResources['ProgressSubscription'] = $progressSub

        $completionEvent = {
            param($s, $e)
            $state = $e.InvocationStateInfo.State
            $mainForm = $global:DashboardConfig.UI.GlobalProgressBar.FindForm()
            if ($state -eq 'Failed') { Write-Verbose "LOGIN EXCEPTION: $($e.InvocationStateInfo.Reason)" -ForegroundColor Red }
            if ($state -match 'Completed|Failed|Stopped') {
                if ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated) {
                     $mainForm.BeginInvoke([Action]{ CleanUpLoginResources -globalLoginResourcesRef $global:DashboardConfig})
                } else { CleanUpLoginResources -globalLoginResourcesRef $global:DashboardConfig }
            }
        }
        $eventName = 'LoginOp_' + [Guid]::NewGuid().ToString('N')
        $eventSub = Register-ObjectEvent -InputObject $loginPS -EventName InvocationStateChanged -SourceIdentifier $eventName -Action $completionEvent
        
        $global:LoginResources['EventSubscriptionId'] = $eventName
        $global:LoginResources['EventSubscriber'] = $eventSub

        $asyncResult = $loginPS.BeginInvoke()
        $global:LoginResources['AsyncResult'] = $asyncResult

        Write-Verbose "LOGIN: Background process started." -ForegroundColor Green

    } catch {
        Write-Verbose "LOGIN: Failed to start background thread: $_" -ForegroundColor Red
        CleanUpLoginResources -globalLoginResourcesRef $global:DashboardConfig
    }
}

function CleanUpLoginResources {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        $globalLoginResourcesRef
    )

    # Prevent re-entry if already stopping
    if ($global:LoginResources['IsStopping'] -and $global:LoginResources.IsMouseHookActive) { 
        Write-Verbose "LOGIN: CleanUpLoginResources called but already stopping." -ForegroundColor DarkYellow
        return 
    }
    $global:LoginResources['IsStopping'] = $true

    Write-Verbose "LOGIN: Cleaning up resources..." -ForegroundColor Cyan

    if (-not $globalLoginResourcesRef) { $globalLoginResourcesRef = $global:DashboardConfig }

    # Move ActiveLoginPids to GracePids for monitoring, clear ActiveLoginPids
    if ($global:DashboardConfig.State.ActiveLoginPids) {
        $pidsToGrace = @($global:DashboardConfig.State.ActiveLoginPids)
        foreach ($pidToGrace in $pidsToGrace) {
            $global:DashboardConfig.State.LoginGracePids[$pidToGrace] = (Get-Date).AddSeconds(120)
        }
        $global:DashboardConfig.State.ActiveLoginPids.Clear()
    }

    # Stop mouse hook
	if ($global:LoginResources.IsMouseHookActive) {
        try {
            [Custom.MouseHookManager]::Stop()
            $global:LoginResources.IsMouseHookActive = $false
            Write-Verbose "LOGIN: Mouse hook stopped." -ForegroundColor Green
        } catch {
            Write-Error "LOGIN: Failed to stop mouse hook: $_" -ForegroundColor Red
        }
	}

    # Prepare UI cleanup action. This MUST be invoked on the UI thread.
    $uiCleanupAction = {
        $ui = $global:DashboardConfig.UI
        if ($ui.GlobalProgressBar) {
            $ui.GlobalProgressBar.Visible = $false
            $ui.GlobalProgressBar.Value = 0
            $ui.GlobalProgressBar.CustomText = ""
        }
        if ($ui.LoginButton) {
            $ui.LoginButton.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
            $ui.LoginButton.Text = "Login"
        }
        
		if ($global:DashboardConfig.State) {
            $global:DashboardConfig.State['LoginActive'] = $false
        }
    }
    # Get the main form.
    $mainForm = $global:DashboardConfig.UI.GlobalProgressBar.FindForm()
    
    # If on UI thread or can invoke, execute UI cleanup.
    # Check if we are on the UI thread using InvokeRequired
    if ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated -and $mainForm.InvokeRequired) {
        try { 
            # Use BeginInvoke to prevent deadlocks if the main thread is busy waiting
            $mainForm.BeginInvoke($uiCleanupAction)
        } catch {
            Write-Warning "LOGIN: Failed to BeginInvoke UI cleanup action: $_" -ForegroundColor Yellow
            # Fallback to direct call, but this might fail if not on UI thread
            & $uiCleanupAction $globalLoginResourcesRef
        }
    } elseif ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated) {
        # Already on UI thread
        & $uiCleanupAction $globalLoginResourcesRef
    } else {
        # UI is not available, log and proceed with non-UI cleanup
        Write-Warning "LOGIN: UI form not available for cleanup, skipping UI updates." -ForegroundColor Yellow
        $global:DashboardConfig.State['LoginActive'] = $false # Still update state
    }

    # Capture a copy of the resources *before* nulling global ones
    $localRes = $global:LoginResources.Clone()

    # Unregister PowerShell event subscriptions
    if ($localRes.EventSubscriptionId) {
        Unregister-Event -SourceIdentifier $localRes.EventSubscriptionId -ErrorAction SilentlyContinue
        Write-Verbose "LOGIN: Unregistered InvocationStateChanged event: $($localRes.EventSubscriptionId)" -ForegroundColor Green
    }
	if ($localRes.ProgressSubscription) {
        Unregister-Event -SubscriptionId $localRes.ProgressSubscription.Id -ErrorAction SilentlyContinue
        Write-Verbose "LOGIN: Unregistered Progress.DataAdded event (ID: $($localRes.ProgressSubscription.Id))" -ForegroundColor Green
    }

    # Dispose PowerShell instance and runspace synchronously
    try {
        if ($localRes.PowerShellInstance) {
            if ($localRes.PowerShellInstance.InvocationStateInfo.State -eq 'Running') {
                $localRes.PowerShellInstance.Stop()
            }
            if ($localRes.AsyncResult -and -not $localRes.AsyncResult.IsCompleted) {
                try {
                    $localRes.PowerShellInstance.EndInvoke($localRes.AsyncResult) | Out-Null
                    Write-Verbose "LOGIN: PowerShell EndInvoke completed." -ForegroundColor DarkGray
                } catch {
                    Write-Warning "LOGIN: Error during PowerShell EndInvoke: $_" -ForegroundColor Yellow
                }
            }
            $localRes.PowerShellInstance.Dispose()
            Write-Verbose "LOGIN: PowerShellInstance disposed." -ForegroundColor DarkGray
        }
        if ($localRes.Runspace) {
            $localRes.Runspace.Dispose()
            Write-Verbose "LOGIN: Runspace disposed." -ForegroundColor DarkGray
        }
    } catch {
        Write-Error "LOGIN: Error during resource disposal: $_" -ForegroundColor Red
    }

    # Reset global LoginResources to null AFTER the local copy is made and cleanup initiated
    $global:LoginResources = @{
        PowerShellInstance  = $null
        Runspace            = $null
        EventSubscriptionId = $null
        EventSubscriber     = $null
        AsyncResult         = $null
        IsStopping          = $false
		ProgressSubscription= $null
        IsMouseHookActive   = $false
    }

    [System.GC]::Collect() # Request garbage collection
    [System.GC]::WaitForPendingFinalizers()
}
#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion
