<# reconnect.psm1 #>

#region Global Configuration & State Wrapper

if (-not $global:DashboardConfig) {
    $global:DashboardConfig = @{
        Settings = @{ Paths = @{ GameLogFile = $null } }
        State    = @{}
        Config   = @{ Profiles = @{}; Options = @{ AutoReconnect = '0' } }
    }
}

if (-not $global:LoginCancellation) {
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

if (-not $global:DashboardConfig.State.NotificationActionQueue) {
    $global:DashboardConfig.State.NotificationActionQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
}

if (-not $global:DashboardConfig.State.ManualReconnectOverrides) { $global:DashboardConfig.State.ManualReconnectOverrides = [System.Collections.Generic.HashSet[int]]::new() }

if (-not $global:DashboardConfig.State.ContainsKey('QueuePaused')) { $global:DashboardConfig.State.QueuePaused = $false }

if (-not $global:DashboardConfig.State.WasInGamePids) { $global:DashboardConfig.State.WasInGamePids = [System.Collections.Generic.HashSet[int]]::new() }
if (-not $global:DashboardConfig.State.FlashingPids) { $global:DashboardConfig.State.FlashingPids = @{} }

if (-not $global:DashboardConfig.State.Timers) { $global:DashboardConfig.State.Timers = @{} }
if (-not $global:DashboardConfig.State.ActiveLoginPids) { $global:DashboardConfig.State.ActiveLoginPids = [System.Collections.Generic.HashSet[int]]::new() }

$global:WatcherBusy = $false
$global:WorkerBusy  = $false

if (-not $global:DashboardConfig.State.NotificationMap) { $global:DashboardConfig.State.NotificationMap = @{} }
if (-not $global:NotificationStack) { $global:NotificationStack = [System.Collections.ArrayList]::new() }

#endregion

#region Internal Helper Functions

function Global:CheckCancel {
    if ($global:LoginCancellation.IsCancelled) { throw "LoginCancelled" }
}

function Global:SleepWithCancel {
    param([int]$Milliseconds)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $Milliseconds) {
        if ($global:LoginCancellation.IsCancelled) { throw "LoginCancelled" }
        Start-Sleep -Milliseconds 50
    }
    $sw.Stop()
}

function Global:Invoke-GuardedAction {
    param([ScriptBlock]$Action)
    CheckCancel
    $script:ReconnectScriptInitiatedMove = $true
    try { & $Action } finally {
        $script:ReconnectScriptInitiatedMove = $false
        CheckCancel
    }
}

function Global:Update-NotificationPositions {
    $global:NotificationStack = [System.Collections.ArrayList]@($global:NotificationStack | Where-Object { $_ -and -not $_.IsDisposed })

    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $baseX = $screen.Right - 330
    $baseY = $screen.Bottom - 145 - 10

    for ($i = 0; $i -lt $global:NotificationStack.Count; $i++) {
        $form = $global:NotificationStack[$i]
        $targetY = $baseY - ($i * 140)
        if ($form -and -not $form.IsDisposed) {
            $form.Location = New-Object System.Drawing.Point($baseX, $targetY)
        }
    }
}

function Global:Close-Notification {
    param([int]$PidKey, [bool]$TriggerRedraw = $true)
    if ($global:DashboardConfig.State.NotificationMap.ContainsKey($PidKey)) {
        try {
            $f = $global:DashboardConfig.State.NotificationMap[$PidKey]
            if ($f -and -not $f.IsDisposed) {
                $f.Close()
                if ($global:NotificationStack.Contains($f)) {
                    $global:NotificationStack.Remove($f)
                }
            }
        } catch {}
        $global:DashboardConfig.State.NotificationMap.Remove($PidKey)

        if ($TriggerRedraw -and $global:DashboardConfig.UI.MainForm) {
            $global:DashboardConfig.UI.MainForm.BeginInvoke([Action]{ Update-NotificationPositions })
        }
    }
}

function Global:Show-InteractiveNotification {
    param(
        [string]$Title,
        [string]$Message,
        [hashtable]$Buttons,
        [string]$Type = "Normal",
        [int]$RelatedPid = 0,
        [int]$TimeoutSeconds = 15
    )

    if ($RelatedPid -gt 0) {
        Close-Notification -PidKey $RelatedPid -TriggerRedraw:$false
    }

    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.Size = New-Object System.Drawing.Size(320, 135)
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)

    if ($RelatedPid -gt 0) { $global:DashboardConfig.State.NotificationMap[$RelatedPid] = $form }

    $form.Add_FormClosed({
        param($s, $e)
        if ($global:NotificationStack.Contains($s)) {
            $global:NotificationStack.Remove($s)
            $map = $global:DashboardConfig.State.NotificationMap.GetEnumerator() | Where-Object { $_.Value -eq $s } | Select-Object -First 1
            if ($map) { $global:DashboardConfig.State.NotificationMap.Remove($map.Key) }
        }
        if ($global:DashboardConfig.UI.MainForm) {
            $global:DashboardConfig.UI.MainForm.BeginInvoke([Action]{ Update-NotificationPositions })
        }
    })

    $pnl = New-Object System.Windows.Forms.Panel; $pnl.Dock = "Fill"; $pnl.BorderStyle = "FixedSingle"; $form.Controls.Add($pnl)

    $stripContainer = New-Object System.Windows.Forms.Panel
    $stripContainer.Size = New-Object System.Drawing.Size(5, 135); $stripContainer.Dock = "Left"; $stripContainer.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $pnl.Controls.Add($stripContainer)

    $strip = New-Object System.Windows.Forms.Panel
    $strip.Size = New-Object System.Drawing.Size(5, 135); $strip.Dock = "Bottom"
    $strip.BackColor = if ($Type -eq "Warning") { [System.Drawing.Color]::Orange } elseif ($Type -eq "Info") { [System.Drawing.Color]::CornflowerBlue } else { [System.Drawing.Color]::IndianRed }
    $stripContainer.Controls.Add($strip)

    $lblClose = New-Object System.Windows.Forms.Label
    $lblClose.Text = "X"; $lblClose.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $lblClose.ForeColor = [System.Drawing.Color]::Gray
    $lblClose.Location = New-Object System.Drawing.Point(298, 4); $lblClose.Size = New-Object System.Drawing.Size(20, 25); $lblClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $lblClose.Tag = $RelatedPid
    $lblClose.Add_Click({
                        $pidToClose = $this.Tag
                        if ($pidToClose -and $global:DashboardConfig.State.ScheduledReconnects.ContainsKey($pidToClose)) {
                             $global:DashboardConfig.State.ScheduledReconnects.Remove($pidToClose)
                        }
                        Close-Notification -PidKey $pidToClose
    })
    $pnl.Controls.Add($lblClose)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = $Title; $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold); $lblTitle.ForeColor = [System.Drawing.Color]::White
    $lblTitle.Location = New-Object System.Drawing.Point(15, 10); $lblTitle.AutoSize = $true
    $pnl.Controls.Add($lblTitle)

    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Text = $Message; $lblMsg.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblMsg.ForeColor = [System.Drawing.Color]::LightGray
    $lblMsg.Location = New-Object System.Drawing.Point(15, 32); $lblMsg.Size = New-Object System.Drawing.Size(290, 65)
    $pnl.Controls.Add($lblMsg)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 100

    $startTime = [DateTime]::Now
    $endTime   = $startTime.AddSeconds($TimeoutSeconds)

    $timer.Tag = @{
        Pid = $RelatedPid;
        Form = $form;
        StartTime = $startTime;
        EndTime = $endTime;
        TotalMs = ($TimeoutSeconds * 1000);
        Strip = $strip;
        InitialHeight = 135;
        ActionQueue = $global:DashboardConfig.State.NotificationActionQueue
    }

    if ($Buttons) {
        $count = 0
        foreach ($key in $Buttons.Keys) {
            $btn = New-Object System.Windows.Forms.Button
            $btn.Text = $key
            $btn.Size = New-Object System.Drawing.Size(90, 25)
            $btn.FlatStyle = "Flat"
            $btn.ForeColor = [System.Drawing.Color]::White
            $btn.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
            $btn.FlatAppearance.BorderSize = 0
            $btn.Location = New-Object System.Drawing.Point((215 - ($count * 100)), 100)

            $btn.Tag = @{ Command = $Buttons[$key]; Pid = $RelatedPid; Form = $form; Timer = $timer; ActionQueue = $global:DashboardConfig.State.NotificationActionQueue }

            $btn.Add_Click({
                $data = $this.Tag
                $cmd = $data.Command

                if ($cmd -eq "Delay") {
                    $data.Timer.Tag.EndTime = [DateTime]::Now.AddMinutes(2)
                    $data.Timer.Tag.TotalMs = 120000
                }

                if ($data.ActionQueue) {
                    $data.ActionQueue.Enqueue(@{ Action = $cmd; Pid = $data.Pid })
                }

                if ($cmd -ne "Delay" -and $data.Form -and -not $data.Form.IsDisposed) {
                    $data.Form.Close()
                }
            })
            $pnl.Controls.Add($btn)
            $count++
        }
    }

    $global:NotificationStack.Add($form)
    if ($global:DashboardConfig.UI.MainForm) {
        $global:DashboardConfig.UI.MainForm.BeginInvoke([Action]{ Update-NotificationPositions })
    }

    $timer.Add_Tick({
        param($s, $e)
        $tag = $s.Tag
        $now = [DateTime]::Now

        if ($tag.Form.IsDisposed) { $s.Stop(); $s.Dispose(); return }

        if ($tag.Pid -gt 0 -and $global:DashboardConfig.State.ScheduledReconnects.ContainsKey($tag.Pid)) {
             $schedTime = $global:DashboardConfig.State.ScheduledReconnects[$tag.Pid]
             if ($schedTime -gt $tag.EndTime) {
                 $tag.EndTime = $schedTime
             }
        }

        if ($now -ge $tag.EndTime) {
            $s.Stop()
            if ($tag.Form -and -not $tag.Form.IsDisposed) {
                 $tag.Form.Close()
            }
            $s.Dispose()
            return
        }

        $mousePos = [System.Windows.Forms.Cursor]::Position
        $isHovering = $tag.Form.DesktopBounds.Contains($mousePos)

        if (-not $isHovering) {
            $remaining = ($tag.EndTime - $now).TotalMilliseconds
            $pct = $remaining / $tag.TotalMs
            if ($pct -lt 0) { $pct = 0 }
            if ($pct -gt 1) { $pct = 1 }

            $newHeight = [int]($tag.InitialHeight * $pct)
            $tag.Strip.Height = $newHeight
        } else {
             $tag.EndTime = $tag.EndTime.AddMilliseconds(100)

             if ($tag.Pid -gt 0 -and $global:DashboardConfig.State.ScheduledReconnects.ContainsKey($tag.Pid)) {
                 $global:DashboardConfig.State.ScheduledReconnects[$tag.Pid] = $global:DashboardConfig.State.ScheduledReconnects[$tag.Pid].AddMilliseconds(100)
             }
        }
    })

    $timer.Start()
    $form.Show()
}

function Global:ClearGameLogs {
    $LogDirs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::InvariantCultureIgnoreCase)
    $LauncherPath = $global:DashboardConfig.Config['LauncherPath']['LauncherPath']
    if (-not [string]::IsNullOrWhiteSpace($LauncherPath)) {
        $Parent = Split-Path -Path $LauncherPath -Parent
        if (Test-Path "$Parent\Log") { [void]$LogDirs.Add("$Parent\Log") }
    }
    if ($global:DashboardConfig.Config['Profiles']) {
        foreach ($p in $global:DashboardConfig.Config['Profiles'].Values) {
            if ($p -and (Test-Path "$p\Log")) { [void]$LogDirs.Add("$p\Log") }
        }
    }
    foreach ($Dir in $LogDirs) {
        if (Test-Path $Dir) {
            try {
                $Files = Get-ChildItem -Path $Dir -Filter "network_*.log" -ErrorAction SilentlyContinue
                foreach ($File in $Files) {
                    try { Clear-Content -Path $File.FullName -Force -ErrorAction Stop } catch {}
                }
            } catch {}
        }
    }
}

#endregion

#region Core Logic: Watcher & Worker

function Global:Start-DisconnectWatcher {
    [CmdletBinding()]
    param()

    Stop-DisconnectWatcher
    ClearGameLogs

    $global:WatcherBusy = $false
    $global:WorkerBusy = $false
    $global:DashboardConfig.State.DisconnectActive = $true

    if ($global:DashboardConfig.State.WasInGamePids) { $global:DashboardConfig.State.WasInGamePids.Clear() }
    if ($global:DashboardConfig.State.FlashingPids) { $global:DashboardConfig.State.FlashingPids.Clear() }
    if ($global:DashboardConfig.State.ScheduledReconnects) { $global:DashboardConfig.State.ScheduledReconnects.Clear() }
    if ($global:DashboardConfig.State.ManualReconnectOverrides) { $global:DashboardConfig.State.ManualReconnectOverrides.Clear() }
    if ($global:DashboardConfig.State.NotificationActionQueue) { $global:DashboardConfig.State.NotificationActionQueue.Clear() }

    $global:DashboardConfig.State.QueuePaused = $false

    Write-Verbose "Starting Disconnect Supervisor (Separated Logic Mode)..."

    $TimerWatcher = New-Object System.Windows.Forms.Timer
    $TimerWatcher.Interval = 2000

    $null = $JobWatcher; $JobWatcher = Register-ObjectEvent -InputObject $TimerWatcher -EventName Tick -SourceIdentifier "DisconnectWatcherTick" -Action {
        if ($global:WatcherBusy) { return }
        $global:WatcherBusy = $true

        try {
            $now = Get-Date
            if ($global:DashboardConfig.State.ScheduledReconnects.Count -gt 0) {
                $scheduledPids = [int[]]@($global:DashboardConfig.State.ScheduledReconnects.Keys)

                foreach ($sPid in $scheduledPids) {
                    $triggerTime = $global:DashboardConfig.State.ScheduledReconnects[$sPid]

                    if ($now -ge $triggerTime) {
                        if ($global:DashboardConfig.State.LoginActive -or $global:DashboardConfig.State.ReconnectQueue.Count -gt 0) {
                             $global:DashboardConfig.State.ScheduledReconnects[$sPid] = $now.AddSeconds(1)
                        } else {
                            if (-not $global:DashboardConfig.State.ReconnectQueue.Contains($sPid)) {
                                Write-Verbose "Auto-Reconnect Triggered for PID $sPid"
                                $global:DashboardConfig.State.ReconnectQueue.Enqueue($sPid)
                            }
                            $global:DashboardConfig.State.ScheduledReconnects.Remove($sPid)
                            Close-Notification -PidKey $sPid
                        }
                    }
                }
            }

            if (-not $global:DashboardConfig.UI.DataGridFiller) { return }

            $PidConnectionCounts = @{}
            $CheckFailed = $false
            try {
                $AllConns = Get-NetTCPConnection -State Established -ErrorAction Stop
                if ($AllConns) {
                    foreach ($c in $AllConns) {
                        $p = [int]$c.OwningProcess
                        if (-not $PidConnectionCounts.ContainsKey($p)) { $PidConnectionCounts[$p] = 0 }
                        $PidConnectionCounts[$p]++
                    }
                }
            } catch { $CheckFailed = $true }

            if (-not $CheckFailed) {
                $WasInGame = $global:DashboardConfig.State.WasInGamePids
                $Flashing = $global:DashboardConfig.State.FlashingPids
                $Cooldowns = $global:DashboardConfig.State.PidCooldowns

                $globalAutoSetting = $false
                if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options']['AutoReconnect']) {
                    if ("$($global:DashboardConfig.Config['Options']['AutoReconnect'])" -match '1|true') { $globalAutoSetting = $true }
                }

                foreach ($row in $global:DashboardConfig.UI.DataGridFiller.Rows) {
                    if (-not ($row.Tag -is [System.Diagnostics.Process])) { continue }

                    $proc = $row.Tag
                    $pidInt = $proc.Id
                    $profileName = [string]$row.Cells[1].Value

                    if ($proc.HasExited) {
                        if ($WasInGame.Contains($pidInt)) { $WasInGame.Remove($pidInt) }
                        if ($Flashing.ContainsKey($pidInt)) { $Flashing.Remove($pidInt) }
                        if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($pidInt)) {
                             $global:DashboardConfig.State.ScheduledReconnects.Remove($pidInt)
                        }
                        continue
                    }

                    $CurrentCount = if ($PidConnectionCounts.ContainsKey($pidInt)) { $PidConnectionCounts[$pidInt] } else { 0 }

                    if ($CurrentCount -ge 2) {
                        if (-not $WasInGame.Contains($pidInt)) { [void]$WasInGame.Add($pidInt) }
                        if ($Flashing.ContainsKey($pidInt)) { $Flashing.Remove($pidInt) }
                        if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($pidInt)) {
                             $global:DashboardConfig.State.ScheduledReconnects.Remove($pidInt)
                        }
                        Close-Notification -PidKey $pidInt
                    }

                    if ($CurrentCount -eq 0 -and $WasInGame.Contains($pidInt)) {

                        if ($global:DashboardConfig.State.ActiveLoginPids.Contains($pidInt)) { continue }
                        if ($global:DashboardConfig.State.ReconnectQueue.Contains($pidInt)) { continue }
                        if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($pidInt)) { continue }

                        $isFocused = $false
                        try {
                             $fg = [Custom.Native]::GetForegroundWindow()
                             if ($proc.MainWindowHandle -ne [IntPtr]::Zero -and $proc.MainWindowHandle -eq $fg) { $isFocused = $true }
                        } catch {}

                        if ($isFocused) {
                            $WasInGame.Remove($pidInt)
                            Close-Notification -PidKey $pidInt
                            continue
                        }

                        $isCooldown = $false
                        if ($Cooldowns.ContainsKey($pidInt)) {
                            if ((Get-Date) -lt $Cooldowns[$pidInt].AddSeconds(120)) { $isCooldown = $isCooldown }
                        }

                        $WasInGame.Remove($pidInt)
                        $global:DashboardConfig.State.FlashingPids[$pidInt] = $true
                        $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black

                        $pName = "Default"; $wTitle = $profileName
                        if ($profileName -match '^\[(.*?)\](.*)') { $pName = $matches[1]; $wTitle = $matches[2] }
                        $details = "`nProfile: $pName`nTitel: $wTitle`nTime: $(Get-Date -Format 'HH:mm:ss')"

                        $isProfileAutoReconnect = $false

                        $rp = $global:DashboardConfig.Config['ReconnectProfiles']
                        if ($rp) {
                            $pClean = $profileName
                            if ($pClean -match '\[(.*?)\]') { $pClean = $matches[1] }

                            if ($rp -is [System.Collections.Hashtable]) {
                                if ($rp.ContainsKey($profileName) -or $rp.ContainsKey($pClean)) { $isProfileAutoReconnect = $true }
                            } elseif ($rp.Contains($profileName) -or $rp.Contains($pClean)) {
                                $isProfileAutoReconnect = $true
                            }
                        }

                        if (-not $rp -and $globalAutoSetting) {
                            $isProfileAutoReconnect = $true
                        }

                        $btns = [ordered]@{
                            "Dismiss All" = "DismissAll"
                            "Reconnect in 2m"      = "Delay"
                            "Reconnect Now"     = "Reconnect"
                        }

                        if ($isProfileAutoReconnect) {
                            if ($isCooldown) {
                                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Orange
                                Show-InteractiveNotification -Title "Cooldown Active" -Message "Disconnected again. Auto-reconnect paused.$details" -Type "Warning" -RelatedPid $pidInt -TimeoutSeconds 15 -Buttons $btns
                            } else {
                                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::DarkRed

                                $global:DashboardConfig.State.ScheduledReconnects[$pidInt] = (Get-Date).AddSeconds(15)

                                Show-InteractiveNotification -Title "Connection Lost" -Message "Auto-reconnect in 15s...$details" -Type "Warning" -RelatedPid $pidInt -TimeoutSeconds 15 -Buttons $btns
                            }
                        } else {
                             $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::DarkRed
                             if ($isCooldown) {
                                Show-InteractiveNotification -Title "Cooldown Active" -Message "Disconnected again.$details" -Type "Warning" -RelatedPid $pidInt -TimeoutSeconds 15 -Buttons $btns
                             } else {
                                Show-InteractiveNotification -Title "Connection Lost" -Message "Disconnected.$details" -Type "Warning" -RelatedPid $pidInt -TimeoutSeconds 15 -Buttons $btns
                             }
                        }
                    }
                }
            }
        } catch {} finally { $global:WatcherBusy = $false }
    }

    $TimerWorker = New-Object System.Windows.Forms.Timer
    $TimerWorker.Interval = 500

    $null = $JobWorker; $JobWorker = Register-ObjectEvent -InputObject $TimerWorker -EventName Tick -SourceIdentifier "DisconnectWorkerTick" -Action {
        if ($global:WorkerBusy) { return }
        $global:WorkerBusy = $true

        try {
            $cmdQueue = $global:DashboardConfig.State.NotificationActionQueue
            while ($cmdQueue -and $cmdQueue.Count -gt 0) {
                $cmd = $cmdQueue.Dequeue()
                $cPid = $cmd.Pid
                $action = $cmd.Action

                switch ($action) {
                    "Reconnect" {
                        if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($cPid)) {
                             $global:DashboardConfig.State.ScheduledReconnects.Remove($cPid)
                        }
                        if ($global:DashboardConfig.State.ManualReconnectOverrides) { [void]$global:DashboardConfig.State.ManualReconnectOverrides.Add($cPid) }

                        if (-not $global:DashboardConfig.State.ReconnectQueue.Contains($cPid)) {
                            $global:DashboardConfig.State.ReconnectQueue.Enqueue($cPid)
                        }
                        Close-Notification -PidKey $cPid
                    }

                    "Delay" {
                    	$global:DashboardConfig.State.ScheduledReconnects[$cPid] = (Get-Date).AddMinutes(2)
                        if ($global:DashboardConfig.State.ManualReconnectOverrides) { [void]$global:DashboardConfig.State.ManualReconnectOverrides.Add($cPid) }
                        Write-Verbose "PID $cPid Reconnect delayed 2m"
                    }

                    "DismissAll" {
                        $targets = [int[]]@($global:DashboardConfig.State.FlashingPids.Keys)
                        foreach ($tPid in $targets) {
                            if ($global:DashboardConfig.State.ScheduledReconnects.ContainsKey($tPid)) {
                                 $global:DashboardConfig.State.ScheduledReconnects.Remove($tPid)
                            }
                            Close-Notification -PidKey $tPid
                        }
                    }
                }
            }

            if ($global:DashboardConfig.State.LoginActive) { return }
            if ($global:DashboardConfig.State.QueuePaused) { return }
            if ($global:DashboardConfig.State.ReconnectQueue.Count -gt 0) {
                $PidToReconnect = $global:DashboardConfig.State.ReconnectQueue.Dequeue()
                Invoke-ReconnectionSequence -PidToReconnect $PidToReconnect
            }
        } catch { } finally { $global:WorkerBusy = $false }
    }

    $global:DashboardConfig.State.Timers['Watcher'] = $TimerWatcher
    $global:DashboardConfig.State.Timers['Worker'] = $TimerWorker
    $TimerWatcher.Start()
    $TimerWorker.Start()
}

function Global:Stop-DisconnectWatcher {
    [CmdletBinding()]
    param()
    Write-Verbose "Stopping Disconnect Supervisor..."

    if ($global:DashboardConfig.State.Timers) {
        if ($global:DashboardConfig.State.Timers['Watcher']) {
            $t = $global:DashboardConfig.State.Timers['Watcher']
            try { $t.Stop(); $t.Dispose() } catch {}
            $global:DashboardConfig.State.Timers.Remove('Watcher')
        }
        if ($global:DashboardConfig.State.Timers['Worker']) {
            $t = $global:DashboardConfig.State.Timers['Worker']
            try { $t.Stop(); $t.Dispose() } catch {}
            $global:DashboardConfig.State.Timers.Remove('Worker')
        }
    }

    Get-EventSubscriber -SourceIdentifier "DisconnectWatcherTick" -ErrorAction SilentlyContinue | Unregister-Event -Force
    Get-EventSubscriber -SourceIdentifier "DisconnectWorkerTick" -ErrorAction SilentlyContinue | Unregister-Event -Force

    if ($global:DashboardConfig.State.NotificationMap) {
        $keysToClose = [int[]]@($global:DashboardConfig.State.NotificationMap.Keys)
        foreach ($key in $keysToClose) {
            Close-Notification -PidKey $key
        }
    }
    if ($global:NotificationStack) {
        foreach ($f in $global:NotificationStack) { try { $f.Close() } catch {} }
        $global:NotificationStack.Clear()
    }

    if ($global:DashboardConfig.State) {
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

function Global:Invoke-ReconnectionSequence {
    param([int]$PidToReconnect)

    $Row = $null
    if ($global:DashboardConfig.UI.DataGridFiller) {
        $Row = $global:DashboardConfig.UI.DataGridFiller.Rows | Where-Object { $_.Tag -is [System.Diagnostics.Process] -and $_.Tag.Id -eq $PidToReconnect } | Select-Object -First 1
    }

    if (-not $Row) { return }

    $CachedEntryNumber = [int]$Row.Cells[0].Value
    $CachedProfileName = [string]$Row.Cells[1].Value
    if ($CachedProfileName -match '\[(.*?)\]') { $CachedProfileName = $matches[1] }

    $global:LoginCancellation.IsCancelled = $false
    $Row.Cells[3].Value = "Reconnecting..."
    $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::DarkRed
    $Row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black

    $hookCallback = [Custom.MouseHookManager+HookProc] {
        param($nCode, $wParam, $lParam)
        if ($nCode -ge 0 -and $wParam -eq 0x0200) {
            if (-not $script:ReconnectScriptInitiatedMove) {
                if (-not $global:LoginCancellation.IsCancelled) {
                    $global:LoginCancellation.IsCancelled = $true
                }
            }
        }
        return [Custom.MouseHookManager]::CallNextHookEx([Custom.MouseHookManager]::HookId, $nCode, $wParam, $lParam)
    }

    try {
        [Custom.MouseHookManager]::Start($hookCallback)
        CheckCancel

        $Process = $Row.Tag
        if ($Process.HasExited) {
            $Row.Cells[3].Value = "Exited"
            return
        }

        $hWnd = $Process.MainWindowHandle

        Write-Verbose "Auto-Reconnect: Processing PID $PidToReconnect (Profile: $CachedProfileName)..."

        $isManuallyForced = $false
        if ($global:DashboardConfig.State.ManualReconnectOverrides -and $global:DashboardConfig.State.ManualReconnectOverrides.Contains($PidToReconnect)) {
            $isManuallyForced = $true
            [void]$global:DashboardConfig.State.ManualReconnectOverrides.Remove($PidToReconnect)
        }

        if (-not $isManuallyForced) {
            if (-not ($global:DashboardConfig.Config['ReconnectProfiles'] -and $global:DashboardConfig.Config['ReconnectProfiles'].Contains($CachedProfileName))) {
                $Row.Cells[3].Value = "Disconnected"
                Show-InteractiveNotification -Title "Reconnect Blocked" -Message "Profile '$CachedProfileName' not enabled for auto-reconnect." -Type "Info" -RelatedPid $PidToReconnect -TimeoutSeconds 10
                return
            }
        }

        $global:DashboardConfig.State.PidCooldowns[$PidToReconnect] = Get-Date

        $loginConfig = $global:DashboardConfig.Config['LoginConfig']
        $profileConfig = $null
        if ($loginConfig.Contains($CachedProfileName)) { $profileConfig = $loginConfig[$CachedProfileName] }
        elseif ($loginConfig.Contains('Default')) { $profileConfig = $loginConfig['Default'] }
        if (-not $profileConfig) { $profileConfig = @{} }

        $DisconnectCoordsString = if ($profileConfig['DisconnectOKCoords']) { $profileConfig['DisconnectOKCoords'] } else { "0,0" }
        $LoginDetailsOKString = if ($profileConfig['LoginDetailsOKCoords']) { $profileConfig['LoginDetailsOKCoords'] } else { "0,0" }
        $FirstNickString = if ($profileConfig['FirstNickCoords']) { $profileConfig['FirstNickCoords'] } else { "0,0" }
        $ScrollDownString = if ($profileConfig['ScrollDownCoords']) { $profileConfig['ScrollDownCoords'] } else { "0,0" }

        if ($DisconnectCoordsString -eq "0,0") { return }

        $ParseXY = { param($s) if ($s -match ',') { $p=$s.Split(','); return @{X=[int]$p[0]; Y=[int]$p[1]} } return $null }

        $disCoords = &$ParseXY $DisconnectCoordsString
        $logDetCoords = &$ParseXY $LoginDetailsOKString
        $firstNickCoords = &$ParseXY $FirstNickString
        $scrollDownCoords = &$ParseXY $ScrollDownString

        CheckCancel

		if ($hWnd -ne [IntPtr]::Zero) {
			if ([Custom.Native]::IsWindowMinimized($hWnd)) {
			Invoke-GuardedAction {  [Custom.Native]::ShowWindow($hWnd, 6); SleepWithCancel -Milliseconds 250; [Custom.Native]::ShowWindow($hWnd, 9) }; SleepWithCancel -Milliseconds 250
        	}
        	Invoke-GuardedAction { [Custom.Native]::SetForegroundWindow($hWnd) }; SleepWithCancel -Milliseconds 500
		}
		$rect = New-Object Custom.Native+RECT
        if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect)) {
            $absX = $rect.Left + $disCoords.X
            $absY = $rect.Top + $disCoords.Y
            Invoke-GuardedAction {
                [Custom.Native]::SetCursorPos($absX, $absY)
                Start-Sleep -Milliseconds 300
                [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); Start-Sleep -Milliseconds 100
                [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0)
            }
        }

        SleepWithCancel -Milliseconds 1000
        $maxWait = 12000; $counter = 0
        while (-not $Process.Responding -and $counter -lt $maxWait) {
            CheckCancel; Start-Sleep -Milliseconds 1000; $Process.Refresh(); $counter += 1000
        }
        SleepWithCancel -Milliseconds 500

        Invoke-GuardedAction {
            [Custom.Native]::ShowWindow($hWnd, 9)
            [Custom.Native]::SetForegroundWindow($hWnd)
        }
        SleepWithCancel -Milliseconds 500
		while (-not $Process.Responding -and $counter -lt $maxWait) {
            CheckCancel; Start-Sleep -Milliseconds 1000; $Process.Refresh(); $counter += 1000
        }

        if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect)) {
            $entryToClick = $CachedEntryNumber
            Invoke-GuardedAction {
                [Custom.Native]::SetForegroundWindow($hWnd)
                Start-Sleep -Milliseconds 150
                if ($entryToClick -ge 6 -and $entryToClick -le 10) {
                    $absX = $rect.Left + ($scrollDownCoords.X + 5)
                    $absY = $rect.Top + ($scrollDownCoords.Y - 5)
                    [Custom.Native]::SetCursorPos($absX, $absY)
                    Start-Sleep -Milliseconds 50
                    [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); Start-Sleep -Milliseconds 50
                    [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0); Start-Sleep -Milliseconds 200
                    $absX = $rect.Left + $firstNickCoords.X
                    $absY = $rect.Top + $firstNickCoords.Y
                    $targetY = $absY + (($entryToClick - 6) * 18)
                } elseif ($entryToClick -ge 1 -and $entryToClick -le 5) {
                    $absX = $rect.Left + $firstNickCoords.X
                    $absY = $rect.Top + $firstNickCoords.Y
                    $targetY = $absY + (($entryToClick - 1) * 18)
                }
                [Custom.Native]::SetCursorPos($absX, $targetY)
            }
            SleepWithCancel -Milliseconds 100
            Invoke-GuardedAction { [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); Start-Sleep -Milliseconds 50; [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0) }
            SleepWithCancel -Milliseconds 100
            Invoke-GuardedAction { [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); Start-Sleep -Milliseconds 50; [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0) }
        }

        SleepWithCancel -Milliseconds 1000
        $maxWait = 25000; $counter = 0
        while (-not $Process.Responding -and $counter -lt $maxWait) {
            CheckCancel; Start-Sleep -Milliseconds 100; $Process.Refresh(); $counter += 100
        }

        Invoke-GuardedAction { [Custom.Native]::SetForegroundWindow($hWnd) }; SleepWithCancel -Milliseconds 1000

        if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect)) {
            $absX = $rect.Left + $logDetCoords.X; $absY = $rect.Top + $logDetCoords.Y
            Invoke-GuardedAction {
                [Custom.Native]::SetCursorPos($absX, $absY)
                Start-Sleep -Milliseconds 1000
                [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); Start-Sleep -Milliseconds 100
                [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0)
            }
            SleepWithCancel -Milliseconds 500
        }

        if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect)) {
            Invoke-GuardedAction {
                [Custom.Native]::SetForegroundWindow($hWnd)
                Start-Sleep -Milliseconds 150
                if ($entryToClick -ge 6 -and $entryToClick -le 10) {
                    $absX = $rect.Left + $firstNickCoords.X
                    $absY = $rect.Top + $firstNickCoords.Y
                    $targetY = $absY + (($entryToClick - 6) * 18)
                } elseif ($entryToClick -ge 1 -and $entryToClick -le 5) {
                    $absX = $rect.Left + $firstNickCoords.X
                    $absY = $rect.Top + $firstNickCoords.Y
                    $targetY = $absY + (($entryToClick - 1) * 18)
                }
                [Custom.Native]::SetCursorPos($absX, $targetY)
            }
            SleepWithCancel -Milliseconds 100
            Invoke-GuardedAction { [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); Start-Sleep -Milliseconds 50; [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0) }
            SleepWithCancel -Milliseconds 100
            Invoke-GuardedAction { [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); Start-Sleep -Milliseconds 50; [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0) }
        }

        SleepWithCancel -Milliseconds 100
        [Custom.MouseHookManager]::Stop()

        Close-Notification -PidKey $PidToReconnect

        if (Get-Command "LoginSelectedRow" -ErrorAction SilentlyContinue) {
            LoginSelectedRow -RowInput $Row -WindowHandle $hWnd
        }

    } catch {
        if ($_.Exception.Message -eq "LoginCancelled" -or $_.ToString() -eq "LoginCancelled") {
            $Row.Cells[3].Value = "Cancelled"
        } else {
            $Row.Cells[3].Value = "Error"
            Write-Verbose "Reconnect Error: $_"
        }
    } finally {
        try { [Custom.MouseHookManager]::Stop() } catch {}
        $global:LoginCancellation.IsCancelled = $false
        $script:ReconnectScriptInitiatedMove = $false
    }
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion
