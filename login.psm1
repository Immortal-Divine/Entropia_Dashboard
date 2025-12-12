<# login.psm1 
    .SYNOPSIS
        Provides a multi-client automation engine for logging into the game. 
		It uses coordinate-based mouse simulation, direct key presses, and active log file monitoring to navigate the login sequence, with a built-in safety-cancellation feature triggered by user mouse movement.

    .DESCRIPTION
        This module implements the core "Login" functionality of the dashboard. It is a powerful automation script designed to log in multiple game clients sequentially based on user-defined settings. The process is fully automated, from restoring the client window to entering the game world, and provides detailed progress feedback to the user. A critical safety feature allows the user to immediately cancel the entire operation at any time by simply moving their mouse.

        The module's architecture and key features include:

        1.  **Automated Login Sequence (`ProcessSingleClient`):**
            *   For each selected client, the module executes a precise sequence of actions to navigate the login screen and character selection.
            *   **Coordinate-Based Clicks:** It uses pre-configured screen coordinates (set by the user in the Settings UI) to simulate mouse clicks on server, channel, and character selection buttons.
            *   **Simulated Key Presses:** It programmatically sends "Enter" key presses to confirm selections and enter the game.
            *   **State-Driven Logic:** Instead of relying on fixed delays, the automation actively waits for confirmation that the game has reached the next stage. It does this by monitoring the game's own network log file for specific entries (e.g., `CACHE_ACK_JOIN`), ensuring the process is robust and reliable.

        2.  **Safety and User Control (`LoginSelectedRow`):**
            *   Before starting, the script installs a system-wide mouse hook. If any mouse movement is detected that was not initiated by the script itself, a cancellation flag is triggered.
            *   The entire automation sequence is interspersed with checks for this flag. If detected, the process throws an exception and halts immediately, running a `finally` block to guarantee all resources (like the mouse hook) are cleaned up. This gives the user instant and reliable control to abort the process.

        3.  **Progress and UI Feedback:**
            *   Throughout the multi-step login process for each client, the module continuously updates a progress bar and status text on the main dashboard UI. This gives the user clear, real-time feedback on what stage the automation is in (e.g., "Waiting for World Load...").

        4.  **Window Management and Optimization:**
            *   The module automatically restores and brings the target client window to the foreground before interacting with it.
            *   After a successful login, it minimizes the client window and calls the `EmptyWorkingSet` API to reduce its memory footprint, which is essential for running multiple clients.
#>

#region Configuration
class CancellationState {
    [bool]$IsCancelled = $false
    [void]Cancel() {
        if (-not $this.IsCancelled) {
            $this.IsCancelled = $true
            Write-Verbose "LOGIN: Cancellation requested by user mouse movement." -ForegroundColor Yellow
        }
    }
    [void]Reset() { $this.IsCancelled = $false }
}

$global:LoginCancel = [CancellationState]::new()

$TOTAL_STEPS_PER_CLIENT = 13 

# Flag to indicate if mouse movement is script-initiated.
$script:ScriptInitiatedMove = $false

#endregion Configuration

#region Helper Functions

function Lock-MousePosition {
    param([int]$X, [int]$Y)
    $rect = New-Object Custom.Native+RECT
    $rect.Left = $X; $rect.Top = $Y; $rect.Right = $X + 1; $rect.Bottom = $Y + 1
    if ([Custom.Native].GetMethod("ClipCursor", [type[]]@([Custom.Native+RECT].MakeByRefType()))) {
        [Custom.Native]::ClipCursor([ref]$rect)
    }
}

function Unlock-MousePosition {
    if ([Custom.Native].GetMethod("ClipCursor", [type[]]@([IntPtr]))) {
        [Custom.Native]::ClipCursor([IntPtr]::Zero)
    }
}

function Update-Progress {
    param(
        [Parameter(Mandatory=$true)]
        [Custom.TextProgressBar]$ProgressBarObject,
        [string]$Text,
        [int]$Value = -1, 
        [int]$CurrentStep = -1, 
        [int]$TotalSteps = -1 
    )
    $pb = $ProgressBarObject
    if ($pb) {
        $pb.CustomText = $Text
        if ($CurrentStep -ne -1 -and $TotalSteps -gt 0) {
            $percent = [int](($CurrentStep / $TotalSteps) * 100)
            if ($percent -ge 0 -and $percent -le 100) {
                $pb.Value = $percent
            }
        } elseif ($Value -ne -1) {
            $pb.Value = $Value
        } else {
            $pb.Value = 0
        }
    }
}

function Restore-Window {
    param([System.Diagnostics.Process]$Process)
    CheckCancel
    if ($Process -and $Process.MainWindowHandle -ne [IntPtr]::Zero) {
        if ([Custom.Native]::IsWindowMinimized($Process.MainWindowHandle)) {
            [Custom.Native]::BringToFront($Process.MainWindowHandle)
            Start-Sleep -Milliseconds 250
        }
    }
    CheckCancel
}

function CheckCancel {
    if ($global:LoginCancel.IsCancelled) {
        throw "Login cancelled by user mouse movement."
    }
}


function Set-WindowForeground {
    param([System.Diagnostics.Process]$Process)
    
    # Check before action
    CheckCancel

    if (-not $Process -or $Process.MainWindowHandle -eq [IntPtr]::Zero) { return $false }
    
    $script:ScriptInitiatedMove = $true
    try {
        [Custom.Native]::BringToFront($Process.MainWindowHandle)
        Start-Sleep -Milliseconds 100
    } finally {
        $script:ScriptInitiatedMove = $false
        # Check after action (Buffered intervention)
        CheckCancel
    }
    return $true
}

function ParseCoordinates {
    param([string]$ConfigString)
    if ([string]::IsNullOrWhiteSpace($ConfigString) -or $ConfigString -notmatch ',') { return $null }
    $parts = $ConfigString.Split(',')
    if ($parts.Count -eq 2) { return @{ X = [int]$parts[0].Trim(); Y = [int]$parts[1].Trim() } }
    return $null
}

function Invoke-MouseClick {
    param([int]$X, [int]$Y)
    
    # Check before starting
    CheckCancel

    $script:ScriptInitiatedMove = $true
    
    try {
        [Custom.Native]::SetCursorPos($X, $Y)
        Start-Sleep -Milliseconds 20
        [Custom.Native]::SetCursorPos($X, $Y)
        Start-Sleep -Milliseconds 30
        
        $MOUSEEVENTF_LEFTDOWN = 0x0002
        $MOUSEEVENTF_LEFTUP = 0x0004
        
        [Custom.Native]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Start-Sleep -Milliseconds 50
        [Custom.Native]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        Start-Sleep -Milliseconds 50
    } catch {
        Write-Verbose "LOGIN: Mouse click failed: $_" -ForegroundColor Red
    } finally {
        Start-Sleep -Milliseconds 50
        $script:ScriptInitiatedMove = $false
        
        # Check after finishing (Buffered intervention)
        # If user moved mouse *during* the sleeps above, this throws now.
        CheckCancel
    }
}

function Invoke-KeyPress {
    param([int]$VirtualKeyCode)
    
    CheckCancel
    
    # KeyPress doesn't move mouse, but we mark InitiatedMove true 
    # to prevent race conditions in parallel checks if they existed (safeguard)
    $script:ScriptInitiatedMove = $true 
    try {
        $hWnd = [Custom.Native]::GetForegroundWindow()
        [Custom.Ftool]::fnPostMessage($hWnd, 0x0100, $VirtualKeyCode, 0) # WM_KEYDOWN
        Start-Sleep -Milliseconds 50
        [Custom.Ftool]::fnPostMessage($hWnd, 0x0101, $VirtualKeyCode, 0) # WM_KEYUP
        Start-Sleep -Milliseconds 100
    } finally {
        $script:ScriptInitiatedMove = $false
        CheckCancel
    }
}

function Write-LogWithRetry {
    param([string]$FilePath, [string]$Value)
    for ($i=0; $i -lt 5; $i++) {
        try {
            Set-Content -Path $FilePath -Value $Value -Force -ErrorAction Stop
            return
        } catch { Start-Sleep -Milliseconds 50 }
    }
}

function Wait-ForFileAccess {
    param([string]$FilePath)
    return (Test-Path $FilePath)
}

function Wait-UntilClientNormalState {
    param(
        $Row,
        [int]$currentGlobalStep, 
        [int]$totalGlobalSteps,
        [Parameter(Mandatory=$true)]
        [Custom.TextProgressBar]$ProgressBarObject
    )
    if (-not $Row) { return }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    
    while ($true) {
        # Continuous Check
        CheckCancel

        $state = $Row.Cells[3].Value 
        if ($state -eq 'Normal') {
            return
        }
        
        if ($sw.Elapsed.Milliseconds % 2000 -lt 100) {
            Update-Progress -ProgressBarObject $ProgressBarObject -Text "Waiting for Client Normal State..." -CurrentStep $currentGlobalStep -TotalSteps $totalGlobalSteps
        }
        
        Start-Sleep -Milliseconds 250
    }
}

function Wait-ForLogEntry {
    param(
        [string]$LogPath,
        [string[]]$SearchStrings,
        [int]$TimeoutSeconds = 60,
        [int]$currentGlobalStep,
        [int]$totalGlobalSteps,
        [Parameter(Mandatory=$true)]
        [Custom.TextProgressBar]$ProgressBarObject
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Update-Progress -ProgressBarObject $ProgressBarObject -Text "Watching log for: $($SearchStrings -join ', ')..." -CurrentStep $currentGlobalStep -TotalSteps $totalGlobalSteps

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        # Continuous Check
        CheckCancel
        
        if (Test-Path $LogPath) {
            try {
                $lines = Get-Content $LogPath -Tail 20 -ErrorAction SilentlyContinue
                foreach ($str in $SearchStrings) {
                    if ($lines -match [regex]::Escape($str)) {
                        Write-Verbose "Log Match Found: $str" -ForegroundColor Green
                        return $true
                    }
                }
            } catch {}
        }
        Start-Sleep -Milliseconds 250
    }
    return $false
}

function Wait-UntilWorldLoaded {
    param(
        $LogPath,
        [int]$currentGlobalStep, 
        [int]$totalGlobalSteps,
        [Parameter(Mandatory=$true)]
        [Custom.TextProgressBar]$ProgressBarObject
    )
    
    $threshold = 3
    $searchStr = "13 - CACHE_ACK_JOIN"
    if ($global:DashboardConfig.Config['LoginConfig']) {
        $cfg = $global:DashboardConfig.Config['LoginConfig']
        if ($cfg['WorldLoadLogThreshold']) { $threshold = [int]$cfg['WorldLoadLogThreshold'] }
        if ($cfg['WorldLoadLogEntry']) { $searchStr = $cfg['WorldLoadLogEntry'] }
    }
    
    $foundCount = 0
    $timeout = New-TimeSpan -Minutes 2
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    
    Update-Progress -ProgressBarObject $ProgressBarObject -Text "Waiting for World Load..." -CurrentStep $currentGlobalStep -TotalSteps $totalGlobalSteps

    while ($foundCount -lt $threshold) {
        if ($sw.Elapsed -gt $timeout) { throw "World load timeout" }
        
        # Continuous Check
        CheckCancel

        if (Test-Path $LogPath) {
            try {
                $lines = Get-Content $LogPath -Tail 50 -ErrorAction SilentlyContinue
                if ($lines) {
                    $foundCount = ($lines | Select-String -SimpleMatch $searchStr).Count
                }
            } catch {}
        }
        Start-Sleep -Seconds 1
    }
}

function ProcessSingleClient {
    param(
        $Row,
        $LogFilePath,
        $LoginConfig,
        [int]$clientIndex,
        [int]$totalClients,
        [Parameter(Mandatory=$true)]
        [Custom.TextProgressBar]$ProgressBarObject
    )
    # Initial Check
    CheckCancel
    $clientStepCounter = 0
    $process = $Row.Tag
    if (-not $process) { throw "No process attached to row" }
    $entryNumber = $Row.Cells[0].Value
    $clientStepCounter++; Update-Progress -ProgressBarObject $ProgressBarObject -Text "Starting login for Client $entryNumber" -CurrentStep (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter) -TotalSteps ($totalClients * $TOTAL_STEPS_PER_CLIENT)
    # --- READ CONFIGURATION ---
    $serverID = "1"; $channelID = "1"; $charSlot = "1"; $startCollector = "No"
    $settingKey = "Client${entryNumber}_Settings"
    if ($LoginConfig[$settingKey]) {
        $parts = $LoginConfig[$settingKey] -split ','
        if ($parts.Count -eq 4) {
            $serverID = $parts[0]
            $channelID = $parts[1]
            $charSlot = $parts[2]
            $startCollector = $parts[3]
        }
    }
    # 1. Clear Logs & Restore
    Write-LogWithRetry -FilePath $LogFilePath -Value ""
    Restore-Window -Process $process
    CheckCancel
    # 2. Wait for Responsive Window
    $clientStepCounter++; 
    $cgswait = (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter)
    $tgswait = ($totalClients * $TOTAL_STEPS_PER_CLIENT)
    Update-Progress -ProgressBarObject $ProgressBarObject -Text "Waiting for Responsive Window..." -CurrentStep $cgswait -TotalSteps $tgswait
    #Wait-UntilClientNormalState -ProgressBarObject $ProgressBarObject -Row $Row -currentGlobalStep $cgswait -totalGlobalSteps $tgswait
    Set-WindowForeground -Process $process | Out-Null
    # Note: Set-WindowForeground does internal intervention checks
    $rect = New-Object Custom.Native+RECT
    [Custom.Native]::GetWindowRect($process.MainWindowHandle, [ref]$rect)
    # --- STEP 1: INITIAL CLICK ---
    $clientStepCounter++; Update-Progress -ProgressBarObject $ProgressBarObject -Text "Selecting Client Index..." -CurrentStep (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter) -TotalSteps ($totalClients * $TOTAL_STEPS_PER_CLIENT)
    $centerX = [int](($rect.Left + $rect.Right) / 2) + 25
    $centerY = [int](($rect.Top + $rect.Bottom) / 2) + 18
    $adjustedY = $centerY
    if ($entryNumber -ge 6 -and $entryNumber -le 10) {
        $yOffset = ($entryNumber - 8) * 18
        $adjustedY = $centerY + $yOffset
        $scrollCenterX = $centerX + 145
        $scrollCenterY = $centerY + 28
        Invoke-MouseClick -X $scrollCenterX -Y $scrollCenterY
        Start-Sleep -Milliseconds 200
    } elseif ($entryNumber -ge 1 -and $entryNumber -le 5) {
        $yOffset = ($entryNumber - 3) * 18
        $adjustedY = $centerY + $yOffset
    }
    Invoke-MouseClick -X $centerX -Y $adjustedY
    Invoke-MouseClick -X $centerX -Y $adjustedY
    Start-Sleep -Milliseconds 200
    CheckCancel
    # --- STEP 2: SERVER SELECTION ---
    $coordKey = "Server${serverID}Coords"
    if ($LoginConfig[$coordKey]) {
        $coords = ParseCoordinates $LoginConfig[$coordKey]
        if ($coords) {
            $clientStepCounter++; Update-Progress -ProgressBarObject $ProgressBarObject -Text "Clicking Server $serverID..." -CurrentStep (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter) -TotalSteps ($totalClients * $TOTAL_STEPS_PER_CLIENT)
            Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
            Start-Sleep -Milliseconds 50
        }
    }
    CheckCancel
    # --- STEP 3: CHANNEL SELECTION ---
    $coordKey = "Channel${channelID}Coords"
    if ($LoginConfig[$coordKey]) {
        $coords = ParseCoordinates $LoginConfig[$coordKey]
        if ($coords) {
            $clientStepCounter++; Update-Progress -ProgressBarObject $ProgressBarObject -Text "Clicking Channel $channelID..." -CurrentStep (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter) -TotalSteps ($totalClients * $TOTAL_STEPS_PER_CLIENT)
            Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
            Start-Sleep -Milliseconds 50
        }
    }
    CheckCancel
    # --- STEP 4: CERT LOGIN ---
    Invoke-KeyPress -VirtualKeyCode 0x0D 
    $clientStepCounter++; 
    $cgscert = (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter)
    $tgscert = ($totalClients * $TOTAL_STEPS_PER_CLIENT)
    Update-Progress -ProgressBarObject $ProgressBarObject -Text "Waiting for Cert Login..." -CurrentStep $cgscert -TotalSteps $tgscert
    $loginReady = Wait-ForLogEntry -ProgressBarObject $ProgressBarObject -LogPath $LogFilePath -SearchStrings @("6 - LOGIN_PLAYER_LIST") -TimeoutSeconds 25 -currentGlobalStep $cgscert -totalGlobalSteps $tgscert
    if (-not $loginReady) { throw "CERT Login Timeout." }
    CheckCancel
    # --- STEP 5: CHARACTER SELECTION ---
    # This must happen BEFORE the next Enter press
    $coordKey = "Char${charSlot}Coords"
    if ($LoginConfig[$coordKey] -and $LoginConfig[$coordKey] -ne '0,0') {
        $coords = ParseCoordinates $LoginConfig[$coordKey]
        if ($coords) {
            $clientStepCounter++; Update-Progress -ProgressBarObject $ProgressBarObject -Text "Selecting Character Slot $charSlot..." -CurrentStep (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter) -TotalSteps ($totalClients * $TOTAL_STEPS_PER_CLIENT)
            Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
            Start-Sleep -Milliseconds 500 # Wait for selection to register
        }
    } else {
        if ($charSlot -ne "1") {
             Write-Verbose "WARNING: No coordinates for Char $charSlot. Defaulting to Char 1." -ForegroundColor Yellow
        }
    }
    CheckCancel
    # --- STEP 6: MAIN LOGIN (Enter World) ---
    $clientStepCounter++; 
    $cgsmain = (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter)
    $tgsmain = ($totalClients * $TOTAL_STEPS_PER_CLIENT)
    Update-Progress -ProgressBarObject $ProgressBarObject -Text "Entering World..." -CurrentStep $cgsmain -TotalSteps $tgsmain
    Invoke-KeyPress -VirtualKeyCode 0x0D
    Start-Sleep -Milliseconds 500
    CheckCancel
    $cacheJoin = Wait-ForLogEntry -ProgressBarObject $ProgressBarObject -LogPath $LogFilePath -SearchStrings @("13 - CACHE_ACK_JOIN") -TimeoutSeconds 60 -currentGlobalStep $cgsmain -totalGlobalSteps $tgsmain
    if (-not $cacheJoin) { throw "Main Login Timeout." }
    # --- STEP 7: WORLD LOAD ---
    $clientStepCounter++; 
    $cgsworld = (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter)
    $tgsworld = ($totalClients * $TOTAL_STEPS_PER_CLIENT)
    Update-Progress -ProgressBarObject $ProgressBarObject -Text "Waiting for World Load..." -CurrentStep $cgsworld -TotalSteps $tgsworld
    Wait-UntilWorldLoaded -ProgressBarObject $ProgressBarObject -LogPath $LogFilePath -currentGlobalStep $cgsworld -totalGlobalSteps $tgsworld
    # Finalization
    $delay = 1 
    if ($LoginConfig['PostLoginDelay']) { $delay = [int]$LoginConfig['PostLoginDelay'] }
    $clientStepCounter++; Update-Progress -ProgressBarObject $ProgressBarObject -Text "Optimization Delay ($delay s)..." -CurrentStep (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter) -TotalSteps ($totalClients * $TOTAL_STEPS_PER_CLIENT)
    # Sleep with check
    $swDelay = [System.Diagnostics.Stopwatch]::StartNew()
    while ($swDelay.Elapsed.TotalSeconds -lt $delay) {
        CheckCancel
        Start-Sleep -Milliseconds 250
    }
    # Collector Start
    if ($startCollector -eq "Yes" -and $LoginConfig['CollectorStartCoords'] -and $LoginConfig['CollectorStartCoords'] -ne '0,0') {
        $coords = ParseCoordinates $LoginConfig['CollectorStartCoords']
        if ($coords) {
            $clientStepCounter++; Update-Progress -ProgressBarObject $ProgressBarObject -Text "Starting Collector..." -CurrentStep (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter) -TotalSteps ($totalClients * $TOTAL_STEPS_PER_CLIENT)
            Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
            Start-Sleep -Seconds 1
        }
    }
    # Minimize
    $clientStepCounter++; Update-Progress -ProgressBarObject $ProgressBarObject -Text "Minimizing..." -CurrentStep (($clientIndex - 1) * $TOTAL_STEPS_PER_CLIENT + $clientStepCounter) -TotalSteps ($totalClients * $TOTAL_STEPS_PER_CLIENT)
    CheckCancel
    if ($global:DashboardConfig.Config['Options']['HideMinimizedWindows'] -eq '1') {
        # Hide window completely (SW_HIDE = 0)
        #[Custom.Native]::ShowWindow($process.MainWindowHandle, 0)
		[Custom.Native]::SendMessage($process.MainWindowHandle, 0x0112, 0xF020, 0)
    } else {
        # Minimize window gracefully using WM_SYSCOMMAND (SC_MINIMIZE = 0xF020)
        [Custom.Native]::SendMessage($process.MainWindowHandle, 0x0112, 0xF020, 0)
    }
    [Custom.Native]::EmptyWorkingSet($process.Handle)
}
#endregion

#region Core Function

function LoginSelectedRow {
    param(
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.DataGridViewRow]$RowInput,
        [string]$LogFilePath
    )
    $global:DashboardConfig.State.LoginActive = $true
    # Reset cancellation state and hook for this run
    $global:LoginCancel.Reset()
    $script:ScriptInitiatedMove = $false
    # Define and start the mouse hook
    $hookCallback = [Custom.MouseHookManager+HookProc] {
        param($nCode, $wParam, $lParam)
        # WM_MOUSEMOVE is 0x0200
        if ($nCode -ge 0 -and $wParam -eq 0x0200) {
            # Ignore mouse moves that this script initiated
            if (-not $script:ScriptInitiatedMove) {
                $global:LoginCancel.Cancel()
            }
        }
        # Always pass the event to the next hook in the chain
        return [Custom.MouseHookManager]::CallNextHookEx([Custom.MouseHookManager]::HookId, $nCode, $wParam, $lParam)
    }
    [Custom.MouseHookManager]::Start($hookCallback)
    # UI SETUP
    $pb = $global:DashboardConfig.UI.GlobalProgressBar
    if ($pb) {
        $pb.Visible = $true
        $pb.Value = 0
        $pb.CustomText = "Starting Login Process..."
    }
    $global:DashboardConfig.UI.LoginButton.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $global:DashboardConfig.UI.LoginButton.Text = "Running..."
    $rowsToProcess = @()
    if ($RowInput) {
        $rowsToProcess += $RowInput
    } elseif ($global:DashboardConfig.UI.DataGridFiller.SelectedRows.Count -gt 0) {
        $rowsToProcess = $global:DashboardConfig.UI.DataGridFiller.SelectedRows | Sort-Object { $_.Cells[0].Value -as [int] }
    } else {
        Write-Verbose "LOGIN: No rows selected." -ForegroundColor Yellow
        $global:DashboardConfig.State.LoginActive = $false
        # Reset Button to Ready
        $global:DashboardConfig.UI.LoginButton.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
        $global:DashboardConfig.UI.LoginButton.Text = "Login"
        return
    }
    $total = $rowsToProcess.Count
    $totalGlobalSteps = $total * $TOTAL_STEPS_PER_CLIENT
    Update-Progress -ProgressBarObject $pb -Text "Starting Login Process..." -CurrentStep 0 -TotalSteps $totalGlobalSteps
    $current = 0
    $loginConfig = $global:DashboardConfig.Config['LoginConfig']
    if (-not $loginConfig) { $loginConfig = @{} }
    try {
        foreach ($row in $rowsToProcess) {
            CheckCancel
            $current++
            $entryNum = $row.Cells[0].Value
            # Logic to find log path if not provided
            $actualLogPath = $LogFilePath
            if ([string]::IsNullOrEmpty($actualLogPath)) {
            	$LogFolder = ($global:DashboardConfig.Config['LauncherPath']['LauncherPath'] -replace '\\Launcher\.exe$', '')
        		$actualLogPath = Join-Path -Path $LogFolder -ChildPath "Log\network_$(Get-Date -Format 'yyyyMMdd').log"
        	}
            ProcessSingleClient -Row $row -LogFilePath $actualLogPath -LoginConfig $loginConfig -clientIndex $current -totalClients $total -ProgressBarObject $pb
		}
        Update-Progress -ProgressBarObject $pb -Text "Done" -CurrentStep $totalGlobalSteps -TotalSteps $totalGlobalSteps
        Write-Verbose "All selected clients processed." -ForegroundColor Green
    } catch {
        # Catch the Abort/Intervention errors here
        Update-Progress -ProgressBarObject $pb -Text "Aborted" -Value 0
        Write-Verbose "Login Process Stopped: $_" -ForegroundColor Red
        # Cleanup is handled in the 'finally' block.
    } finally {
        # Always unhook the mouse listener
        [Custom.MouseHookManager]::Stop()
        Unlock-MousePosition
        $global:DashboardConfig.State.LoginActive = $false
        # Reset UI
        if ($pb) {
            $pb.Visible = $false
            $pb.Value = 0
            $pb.CustomText = ""
        }
        $global:DashboardConfig.UI.LoginButton.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
        $global:DashboardConfig.UI.LoginButton.Text = "Login"
        # Force the message queue to process UI updates and prevent freezing.
    }
}

#endregion

#region Module Exports

# Export module functions
Export-ModuleMember -Function LoginSelectedRow, Restore-Window, Set-WindowForeground, Wait-ForFileAccess, Invoke-MouseClick, Invoke-KeyPress, Write-LogWithRetry, Update-Progress
#endregion Module Exports