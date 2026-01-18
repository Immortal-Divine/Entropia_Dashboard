<# quickcommands.psm1 #>

#region Helper Functions

function InitChatCommanderRunspace
{
    param($InstanceId)
    
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
    # Only load assemblies used by this module
    $assemblies = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { 
        $_.FullName -match 'Custom.Native' -or 
        $_.FullName -match 'Custom.Ftool'
    }
    
    foreach ($asm in $assemblies)
    {
        if (-not [string]::IsNullOrEmpty($asm.Location))
        {
            $iss.Assemblies.Add((New-Object System.Management.Automation.Runspaces.SessionStateAssemblyEntry($asm.Location)))
        }
    }

    $pool = [runspacefactory]::CreateRunspacePool(1, 5, $iss, $Host)
    $pool.ThreadOptions = 'ReuseThread'
    $pool.Open()
    
    return $pool
}

function Get-ChatCommanderScriptBlock
{
    return {
        param($Handle, $CommandText) # Removed $HoldTime, $Mods, $IsAlt parameters

        $WM_KEYDOWN = 0x0100; $WM_KEYUP = 0x0101
        $VK_RETURN = 0x0D # Enter key
        $VK_SHIFT = 0x10 # Shift
        $VK_BACK = 0x08 # Backspace
        $WM_CHAR = 0x0102

        # Open chat (e.g., press 'T' or Enter)
        # Using VK_RETURN to open chat as it's more common in games for chat input
        [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, $VK_RETURN, 1)
        Start-Sleep -Milliseconds 50 # Short delay for chat window to open
        [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYUP, $VK_RETURN, 0xC0000000)
        Start-Sleep -Milliseconds 50 # Another short delay

        # Return x 10 (Backspace to clear chat)
        for ($i = 0; $i -lt 10; $i++) {
            [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, $VK_BACK, 1)
            Start-Sleep -Milliseconds 5
            [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYUP, $VK_BACK, 0xC0000000)
            Start-Sleep -Milliseconds 5
        }

        # Type the command text
        foreach ($char in $CommandText.ToCharArray())
        {
            [Custom.Ftool]::fnPostMessage($Handle, $WM_CHAR, [int]$char, 1)
            Start-Sleep -Milliseconds 10
        }
        
        # Removed: HoldTime logic

        # Press Enter to send the command
        [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, $VK_RETURN, 1)
        Start-Sleep -Milliseconds 50 # Short delay for enter press
        [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYUP, $VK_RETURN, 0xC0000000)
    }
}

function Invoke-ChatCommanderAction
{
    param($formData, $commandString) # Removed $holdTime
    
    # Variable replacement
    $varText = ""
    if ($formData.PSObject.Properties['TxtVariable'] -and $formData.TxtVariable) {
        $varText = $formData.TxtVariable.Text
    }
    $finalCommand = $commandString.Replace('[var]', $varText)

    if ($formData.SelectedWindow -eq [IntPtr]::Zero) { return $null }

    if (-not $formData.RunspacePool -or $formData.RunspacePool.RunspacePoolStateInfo.State -ne 'Opened')
    {
        Write-Verbose "ChatCommander Runspace closed/broken. Attempting to restart..."
        try {
            if ($formData.RunspacePool) { $formData.RunspacePool.Dispose() }
            $formData.RunspacePool = InitChatCommanderRunspace -InstanceId $formData.InstanceId
        }
        catch {
            Write-Verbose "Failed to restart Runspace: $_"
            $formData.RunspacePool = $null # Explicitly set to null on failure
            return $null
        }
    }

    # Add an explicit check here to ensure RunspacePool is valid before proceeding
    if (-not $formData.RunspacePool -or $formData.RunspacePool.RunspacePoolStateInfo.State -ne 'Opened') {
        Write-Verbose "RunspacePool is not available or not opened after restart attempt."
        return $null
    }

    $ps = [System.Management.Automation.PowerShell]::Create()
    try {
        $ps.RunspacePool = $formData.RunspacePool
        $ps.AddScript((Get-ChatCommanderScriptBlock).ToString()) | Out-Null
        $ps.AddArgument($formData.SelectedWindow) | Out-Null
        $ps.AddArgument($finalCommand) | Out-Null
        # Removed: $ps.AddArgument($holdTime) | Out-Null

        return @{ 
            PowerShell = $ps
            AsyncResult = $ps.BeginInvoke()
        }
    } catch {
        Write-Verbose "ChatCommander Invoke Error: $_"
        if ($ps) { $ps.Dispose() }
        return $null
    }
}

function Invoke-SpecificChatCommand
{
    param($InstanceId, $CommandId)
    
    if ($global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId)) {
        $form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
        if ($form -and $form.Tag) {
            $data = $form.Tag
            $cmd = $data.SavedCommands | Where-Object { $_.Id -eq $CommandId } | Select-Object -First 1
            if ($cmd) {
                $job = Invoke-ChatCommanderAction -formData $data -commandString $cmd.CommandText
                if ($job) {
                    $data.ActiveJobs.Add($job) | Out-Null
                    
                    # If the main timer isn't running to clean up jobs, we handle it here asynchronously
                    if (-not $data.RunningTimer) {
                        [System.Threading.Tasks.Task]::Run([System.Action]{
                            $job.AsyncResult.AsyncWaitHandle.WaitOne() | Out-Null
                            try { $job.PowerShell.EndInvoke($job.AsyncResult); $job.PowerShell.Dispose() } catch {}
                            
                            if ($form -and -not $form.IsDisposed) {
                                $form.BeginInvoke([System.Action]{
                                    if ($data.ActiveJobs) { $data.ActiveJobs.Remove($job) }
                                })
                            }
                        }) | Out-Null
                    }
                }
            }
        }
    }
}

function CreateChatCommanderPositionTimer
{
    param($formData)
    
    $positionTimer = New-Object System.Windows.Forms.Timer
    $positionTimer.Interval = 10
    $positionTimer.Tag = @{
        WindowHandle = $formData.SelectedWindow
        Form         = $formData.Form
        InstanceId   = $formData.InstanceId
        FormData     = $formData
        ZState       = 'unknown' 
		RealLeft     = $formData.Form.Left
		RealTop      = $formData.Form.Top
    }
        
    $positionTimer.Add_Tick({
            param($s, $e)
            try
            {
                if (-not $s -or -not $s.Tag) { return }
                $timerData = $s.Tag
                if (-not $timerData -or $timerData['WindowHandle'] -eq [IntPtr]::Zero) { return }
                if (-not $timerData['Form'] -or $timerData['Form'].IsDisposed) { return }
            
                $rect = New-Object Custom.Native+RECT
                    
                if ([Custom.Native]::IsWindow($timerData['WindowHandle']) -and [Custom.Native]::GetWindowRect($timerData['WindowHandle'], [ref]$rect))
                {
                    try
                    {
                        $sliderValueX = $timerData.FormData.PositionSliderX.Value
                        $maxLeft = $rect.Right - $timerData['Form'].Width - 8
                        $targetLeft = $rect.Left + 8 + (($maxLeft - ($rect.Left + 8)) * $sliderValueX / 100)
                            
                        $sliderValueY = 100 - $timerData.FormData.PositionSliderY.Value
                        $maxTop = $rect.Bottom - $timerData['Form'].Height - 8
                        $targetTop = $rect.Top + 8 + (($maxTop - ($rect.Top + 8)) * $sliderValueY / 100)

                        # Use and update a high-precision 'Real' position to prevent snapping
                        $newRealLeft = $timerData.RealLeft + ($targetLeft - $timerData.RealLeft) * 0.2
                        $newRealTop = $timerData.RealTop + ($targetTop - $timerData.RealTop) * 0.2
                        
                        $timerData.RealLeft = $newRealLeft
                        $timerData.RealTop = $newRealTop

                        # Set the form's integer position for display
                        $timerData['Form'].Left = [int]$newRealLeft
                        $timerData['Form'].Top = [int]$newRealTop
                                                
                        $formHandle = $timerData['Form'].Handle
                        $linkedHandle = $timerData['WindowHandle']
                        $foregroundWindow = [Custom.Native]::GetForegroundWindow()
                        $flags = [Custom.Native]::SWP_NOMOVE -bor [Custom.Native]::SWP_NOSIZE -bor [Custom.Native]::SWP_NOACTIVATE
                        
                        $shouldBeTopMost = $false
                        if ($foregroundWindow -eq $linkedHandle) {
                            $shouldBeTopMost = $true
                        } else {
                            $activeCtrl = [System.Windows.Forms.Control]::FromHandle($foregroundWindow)
                            if ($activeCtrl -and -not $activeCtrl.IsDisposed -and $activeCtrl.Tag) {
                                if ($activeCtrl.Tag.PSObject.Properties['SelectedWindow'] -and $activeCtrl.Tag.SelectedWindow -eq $linkedHandle) {
                                    $shouldBeTopMost = $true
                                }
                            }
                        }

                        if ($shouldBeTopMost)
                        {
                            $forceUpdate = $false
                            if ($foregroundWindow -eq $linkedHandle) {
                                $testPrev = [Custom.Native]::GetWindow($formHandle, 3)
                                $checkCount = 0
                                while ($testPrev -ne [IntPtr]::Zero -and $checkCount -lt 50) {
                                    if ($testPrev -eq $linkedHandle) { $forceUpdate = $true; break }
                                    $testPrev = [Custom.Native]::GetWindow($testPrev, 3)
                                    $checkCount++
                                }
                            }

                            if ($timerData.ZState -ne 'topmost' -or $forceUpdate)
                            {
                                if ($forceUpdate) {
                                    [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::HWND_NOTOPMOST, 0, 0, 0, 0, $flags) | Out-Null
                                }
                                [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::HWND_TOPMOST, 0, 0, 0, 0, $flags) | Out-Null
                                $timerData.ZState = 'topmost'
                            }
                        }
                        else
                        {
                            if ($timerData.ZState -ne 'standard')
                            {
                                [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::HWND_NOTOPMOST, 0, 0, 0, 0, $flags) | Out-Null
                                $timerData.ZState = 'standard'
                            }

                            $next = [Custom.Native]::GetWindow($formHandle, 2) 
                            $amIAboveGame = $false
                            $loopCount = 0
                            
                            while ($next -ne [IntPtr]::Zero -and $loopCount -lt 50)
                            {
                                if ($next -eq $linkedHandle) { $amIAboveGame = $true; break }
                                $next = [Custom.Native]::GetWindow($next, 2)
                                $loopCount++
                            }

                            if (-not $amIAboveGame)
                            {
                                $bottomOfCluster = [Custom.Native]::GetWindow($linkedHandle, 3)
                                if ($bottomOfCluster -ne [IntPtr]::Zero) { [Custom.Native]::PositionWindow($formHandle, $bottomOfCluster, 0, 0, 0, 0, $flags) | Out-Null }
                                else { [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::TopWindowHandle, 0, 0, 0, 0, $flags) | Out-Null }
                            }
                        }
                        $timerData.LastForegroundWindow = $foregroundWindow
                    }
                    catch {}
                }
                else
                {
                    $timerData['Form'].Close()
                    $s.Stop()
                    $s.Dispose()
                    $global:DashboardConfig.Resources.Timers.Remove("ChatCommanderPosition_$($timerData.InstanceId)")
                }
            }
            catch {}
        })
    
    $positionTimer.Start()
    $global:DashboardConfig.Resources.Timers["ChatCommanderPosition_$($formData.InstanceId)"] = $positionTimer
}

function AdjustChatCommanderFormHeight # Renamed from RepositionChatCommandPanels
{
    param($form)
    
    if (-not $form -or $form.IsDisposed) { return }
    
    $form.SuspendLayout()
    try
    {
        $formData = $form.Tag
        
        if ($formData.ShowAllMode)
        {
            # In Show All mode, we use a larger fixed size
            $screen = [System.Windows.Forms.Screen]::FromPoint($form.Location)
            $workingArea = $screen.WorkingArea
            $distToBottom = $workingArea.Bottom - $form.Top
            $targetHeight = 600
            if ($targetHeight -gt $distToBottom - 20) { $targetHeight = $distToBottom - 20 }
            if ($targetHeight -lt 400) { $targetHeight = 400 }
            
            $form.Height = $targetHeight
            $formData.OriginalHeight = $targetHeight
            $formData.PositionSliderY.Height = $targetHeight - 20
            return
        }

        # Standard Mode (Pagination)
        $baseHeight = 285 # Adjusted base height for the top section (header + new command input + search)
        $panel = $formData.panelSavedCommands
        
        # Temporarily enable AutoSize to get preferred height
        $panel.AutoSize = $true
        $panel.PerformLayout()
        $preferredPanelHeight = $panel.Height
        
        # Add space for pagination controls if visible
        $paginationHeight = if ($formData.PanelPagination.Visible) { 35 } else { 0 }

        $newHeight = $baseHeight + $preferredPanelHeight + $paginationHeight + 10 # 10 for padding
        
        $screen = [System.Windows.Forms.Screen]::FromPoint($form.Location)
        $workingArea = $screen.WorkingArea
        $distToBottom = $workingArea.Bottom - $form.Top
        $maxHeight = $distToBottom - 20
        if ($maxHeight -lt 285) { $maxHeight = 285 } # Minimum usable height

        if ($newHeight -gt $maxHeight) {
            $newHeight = $maxHeight
            # Constrain panel
            $panel.AutoSize = $false
            $panel.Height = $maxHeight - $baseHeight - $paginationHeight - 10
            $panel.AutoScroll = $true
        } else {
            # Let panel grow
            $panel.AutoSize = $true
            $panel.AutoScroll = $false
        }

        if (-not $formData.IsCollapsed)
        {
            $finalHeight = $newHeight
            $form.Height = $finalHeight
            $formData.OriginalHeight = $finalHeight
            $formData.PositionSliderY.Height = $finalHeight - 20
        }
    }
    catch { Write-Verbose "AdjustHeight Error: $_" }
    finally
    {
        $form.ResumeLayout()
    }
}

function StartChatCommanderSequence
{
    param($formData, [switch]$RunSavedSequence)

    $steps = [System.Collections.ArrayList]::new()

    if ($RunSavedSequence)
    {
        # Run all currently visible saved commands as a sequence
        foreach ($control in $formData.panelSavedCommands.Controls)
        {
            if ($control -is [System.Windows.Forms.Panel] -and $control.Tag -is [PSCustomObject])
            {
                $savedCommand = $control.Tag
                $steps.Add([PSCustomObject]@{ 
                    Command = $savedCommand.CommandText; 
                })
            }
        }
        if ($steps.Count -eq 0)
        {
            Show-DarkMessageBox 'No saved commands are visible to start a sequence.' 'ChatCommander Error' 'Ok' 'Warning'
            return
        }
    }
    else
    {
        # Run the main command as a single step
        $mainCommand = $formData.TxtCommandText.Text
        
        if (-not [string]::IsNullOrWhiteSpace($mainCommand))
        {
            $steps.Add([PSCustomObject]@{ 
                Command = $mainCommand; 
            })
        }
        else
        {
            Show-DarkMessageBox 'No valid chat command defined in the main input to start a sequence.' 'ChatCommander Error' 'Ok' 'Warning'
            return
        }
    }

    $formData.ActiveSteps = $steps
    $formData.CurrentStepIndex = 0
    
    if (-not $formData.PSObject.Properties['ActiveJobs'])
    {
        $formData | Add-Member -MemberType NoteProperty -Name 'ActiveJobs' -Value ([System.Collections.ArrayList]::new())
    }
    else
    {
        foreach ($job in $formData.ActiveJobs) {
            try {
                if ($job.PowerShell.InvocationStateInfo.State -eq 'Running') {
                    $job.PowerShell.Stop()
                }
                $job.PowerShell.Dispose()
            } catch {}
        }
        $formData.ActiveJobs.Clear()
    }

    if (-not $formData.RunspacePool -or $formData.RunspacePool.RunspacePoolStateInfo.State -ne 'Opened')
    {
        if ($formData.RunspacePool) { try { $formData.RunspacePool.Dispose() } catch {} }
        $formData.RunspacePool = InitChatCommanderRunspace -InstanceId $formData.InstanceId
    }
    
    if ($formData.PSObject.Properties['BtnRunSequence']) { $formData.BtnRunSequence.Enabled = $false; $formData.BtnRunSequence.Visible = $false }

    if (-not $formData.PSObject.Properties['IsChatCommanderRunning'])
    {
        $formData | Add-Member -MemberType NoteProperty -Name 'IsChatCommanderRunning' -Value $true
    }
    $formData.IsChatCommanderRunning = $true

    if (-not $formData.PSObject.Properties['Stopping'])
    {
        $formData | Add-Member -MemberType NoteProperty -Name 'Stopping' -Value $false
    }
    $formData.Stopping = $false

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Tag = $formData
    
    $tickAction = {
        param($s, $e)
        $fData = $s.Tag
        
        $s.Stop()

        if (-not $fData.IsChatCommanderRunning) { return }

        if ($fData.Stopping)
        {
            StopChatCommanderSequence $fData -NaturalEnd
            return
        }

        if (-not $fData.Form -or $fData.Form.IsDisposed)
        {
            $s.Dispose()
            return
        }

        if ($fData.ActiveJobs) {
            $completed = @()
            foreach ($job in $fData.ActiveJobs) {
                if ($job.AsyncResult.IsCompleted) {
                    $completed += $job
                }
            }
            foreach ($job in $completed) {
                try { $job.PowerShell.EndInvoke($job.AsyncResult); $job.PowerShell.Dispose() } catch {}
                $fData.ActiveJobs.Remove($job)
            }
        }

        $idx = $fData.CurrentStepIndex
        if ($idx -lt $fData.ActiveSteps.Count)
        {
            $step = $fData.ActiveSteps[$idx]
            
            $job = Invoke-ChatCommanderAction -formData $fData -commandString $step.Command
            
            if ($job)
            {
                $fData.ActiveJobs.Add($job) | Out-Null
            }

            $nextIdx = ($idx + 1)
            $fData.CurrentStepIndex = $nextIdx
            
            # Sequence runs once, no looping
            $shouldStop = ($nextIdx -ge $fData.ActiveSteps.Count)

            # Always wait for job completion for sequential execution with a fixed delay
            $pollTimer = New-Object System.Windows.Forms.Timer
            $pollTimer.Interval = 2
            $pollTimer.Tag = @{ 
                Job = $job; 
                MainTimer = $s; 
                NextInterval = 50; # Fixed minimal delay between commands
                ShouldStop = $shouldStop;
                FormData = $fData
            }
            
            $pollTimer.Add_Tick({
                $pt = $this
                $d = $pt.Tag
                
                if (-not $d.FormData.IsChatCommanderRunning)
                {
                    $pt.Stop()
                    $pt.Dispose()
                    return
                }

                if ($d.Job.AsyncResult.IsCompleted)
                {
                    $pt.Stop()
                    
                    try {
                        $d.Job.PowerShell.EndInvoke($d.Job.AsyncResult)
                        $d.Job.PowerShell.Dispose()
                    } catch {}
                    
                    $d.FormData.ActiveJobs.Remove($d.Job)

                    if ($d.FormData.IsChatCommanderRunning)
                    {
                        if ($d.ShouldStop)
                        {
                            StopChatCommanderSequence $d.FormData -NaturalEnd
                        }
                        else
                        {
                            $d.MainTimer.Interval = $d.NextInterval
                            $d.MainTimer.Start()
                        }
                    }
                    $pt.Dispose()
                }
            })
            $pollTimer.Start()
        }
    }
    
    $timer.Add_Tick($tickAction)

    $formData.RunningTimer = $timer
    $global:DashboardConfig.Resources.Timers["ChatCommanderTimer_$($formData.InstanceId)"] = $timer

    & $tickAction $timer $null
}

function StopChatCommanderSequence
{
    param($formData, [switch]$NaturalEnd)

    if ($formData.PSObject.Properties['IsChatCommanderRunning'])
    {
        $formData.IsChatCommanderRunning = $false
    }

    if ($formData.RunningTimer)
    {
        $formData.RunningTimer.Stop()
        $formData.RunningTimer.Dispose()
        $formData.RunningTimer = $null
    }
    
    if ($global:DashboardConfig.Resources.Timers.Contains("ChatCommanderTimer_$($formData.InstanceId)"))
    {
        $t = $global:DashboardConfig.Resources.Timers["ChatCommanderTimer_$($formData.InstanceId)"]
        if ($t) { $t.Stop(); $t.Dispose() }
        $global:DashboardConfig.Resources.Timers.Remove("ChatCommanderTimer_$($formData.InstanceId)")
    }

    if (-not $NaturalEnd)
    {
        if ($formData.ActiveJobs)
        {
            foreach ($job in $formData.ActiveJobs)
            {
                try {
                    if ($job.PowerShell) { $job.PowerShell.Stop(); $job.PowerShell.Dispose() }
                } catch {}
            }
            $formData.ActiveJobs.Clear()
        }
    }
    if ($formData.PSObject.Properties['BtnRunSequence']) { $formData.BtnRunSequence.Enabled = $true; $formData.BtnRunSequence.Visible = $true }
}

function ToggleChatCommanderInstance
{
    param($InstanceId)

    if (-not $global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId)) { return }
    $form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
    
    if ($form -and -not $form.IsDisposed -and $form.Tag.Type -eq 'ChatCommander')
    {
        if ($form.InvokeRequired)
        {
            $form.BeginInvoke([System.Action]{ ToggleChatCommanderInstance -InstanceId $InstanceId }) | Out-Null
            return
        }

        $data = $form.Tag
        if ($data.RunningTimer)
        {
            StopChatCommanderSequence $data
        }
        else
        {
            StartChatCommanderSequence $data
        }
    }
}

function ToggleChatCommanderHotkeys
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstanceId,
        [Parameter(Mandatory = $true)]
        [bool]$ToggleState
    )
    
    if ($global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId))
    {
        $form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
        if ($form -and -not $form.IsDisposed -and $form.InvokeRequired)
        {
            $form.BeginInvoke([System.Action]{ ToggleChatCommanderHotkeys -InstanceId $InstanceId -ToggleState $ToggleState }) | Out-Null
            return
        }
    }

    if (-not $global:DashboardConfig.Resources.InstanceHotkeysPaused) { $global:DashboardConfig.Resources.InstanceHotkeysPaused = @{} }
    
    $global:DashboardConfig.Resources.InstanceHotkeysPaused[$InstanceId] = (-not $ToggleState)

    try
    {
        if ($ToggleState)
        {
            try { ResumeHotkeysForOwner -OwnerKey $InstanceId } catch {}
        }
        else
        {
            try { PauseHotkeysForOwner -OwnerKey $InstanceId } catch {}
        }
    }
    catch {}

    if ($global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId))
    {
        $form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
        if ($form -and -not $form.IsDisposed)
        {
            UpdateChatCommanderSettings $form.Tag -forceWrite
        }
    }
}

function CleanupChatCommanderResources
{
    param($instanceId)
    
    $timerKeysToRemove = @()
    foreach ($key in $global:DashboardConfig.Resources.Timers.Keys)
    {
        if ($key -eq "ChatCommanderTimer_$instanceId" -or $key -eq "ChatCommanderPosition_$instanceId")
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
    foreach ($key in $timerKeysToRemove)
    {
        $global:DashboardConfig.Resources.Timers.Remove($key)
    }

    if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
    {
        $form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
        if ($form -and $form.Tag)
        {
            if ($form.Tag.ToolTipFtool)
            {
                $form.Tag.ToolTipFtool.Dispose()
            }
            if ($form.Tag.ActiveJobs)
            {
                foreach ($job in $form.Tag.ActiveJobs) {
                    try {
                        if ($job.PowerShell) {
                            if ($job.PowerShell.InvocationStateInfo.State -eq 'Running') {
                                $job.PowerShell.Stop()
                            }
                            $job.PowerShell.Dispose()
                        }
                    } catch {
                        Write-Verbose "Error disposing ChatCommander Job: $_"
                    }
                }
                $form.Tag.ActiveJobs.Clear()
            }
            if ($form.Tag.RunspacePool)
            {
                try {
                    if ($form.Tag.RunspacePool.RunspacePoolStateInfo.State -ne 'Closed') {
                        $form.Tag.RunspacePool.Close()
                    }
                    $form.Tag.RunspacePool.Dispose()
                } catch {
                    Write-Verbose "Error disposing ChatCommander Runspace: $_"
                }
                $form.Tag.RunspacePool = $null
            }

            if ($form.Tag.HotkeyId) { try { UnregisterHotkeyInstance -Id $form.Tag.HotkeyId -OwnerKey $form.Tag.InstanceId } catch {} }
            if ($form.Tag.GlobalHotkeyId) { try { $globalOwnerKey = "global_toggle_$($form.Tag.InstanceId)"; UnregisterHotkeyInstance -Id $form.Tag.GlobalHotkeyId -OwnerKey $globalOwnerKey } catch {} }
            
            # Robust cleanup using pattern (covers saved commands quickcmd_*)
            try {
                if (Get-Command UnregisterHotkeysByOwnerPattern -ErrorAction SilentlyContinue) {
                    UnregisterHotkeysByOwnerPattern -OwnerPattern "quickcmd_${instanceId}_*"
                }
            } catch {}
        }
    }
    # Ensure the Hotkeys UI reflects any removals performed during cleanup
    if (Get-Command RefreshHotkeysList -ErrorAction SilentlyContinue) { try { RefreshHotkeysList } catch {} }
}

function LoadChatCommanderSettings
{
    param($formData)
    $formData.IsLoading = $true
    
    if (-not $global:DashboardConfig.Config.Contains('ChatCommander')) { $formData.IsLoading = $false; return }
    
    $profilePrefix = $null
    if (Get-Command FindOrCreateProfile -ErrorAction SilentlyContinue)
    {
        $profilePrefix = FindOrCreateProfile $formData.WindowTitle
    }
    
    if ($profilePrefix)
    {
        $configSuffix = ""
        $parts = $formData.InstanceId -split '_'
        if ($parts.Count -gt 2)
        {
            $configSuffix = "_Sub" + $parts[-1]
        }

        $cfg = $global:DashboardConfig.Config['ChatCommander']
        
        $p = "$profilePrefix$configSuffix"
        $visibleKey = "VisibleCommands_$p"

        if ($cfg.Contains("CommandText_$p")) { $formData.TxtCommandText.Text = $cfg["CommandText_$p"] }
        if ($cfg.Contains("Name_$p")) { $formData.Name.Text = $cfg["Name_$p"] }
        
        $globalHotkeyName = "GlobalHotkey_$p"
        if ($cfg.Contains($globalHotkeyName)) { $formData.GlobalHotkey = $cfg[$globalHotkeyName] }

        $hotkeysEnabledName = "hotkeys_enabled_$p"
        if ($cfg.Contains($hotkeysEnabledName))
        {
            $value = $cfg[$hotkeysEnabledName]
            try { $formData.BtnHotkeyToggle.Checked = [bool]::Parse($value) } catch { $formData.BtnHotkeyToggle.Checked = $true }
        }
        else { $formData.BtnHotkeyToggle.Checked = $true }

        if ($cfg.Contains("PosX_$p")) { $formData.PositionSliderX.Value = [int]$cfg["PosX_$p"] }
        if ($cfg.Contains("PosY_$p")) { $formData.PositionSliderY.Value = [int]$cfg["PosY_$p"] }

        # --- FIX: Robust Global Commands Loading ---
        $globalCommands = [System.Collections.ArrayList]::new()
        if ($cfg.Contains("GlobalCommands")) {
            try {
                $raw = $cfg["GlobalCommands"]
                if ($raw -is [System.Collections.IEnumerable] -and $raw -isnot [string]) { $raw = $raw -join '' }
                
                # 1. Try cleaning escaped quotes first if detected
                if ($raw -is [string]) {
                    # Replace literal \" with "
                    $cleanRaw = $raw -replace '\\"', '"'
                    $imported = ConvertFrom-Json $cleanRaw -ErrorAction SilentlyContinue
                    
                    if (-not $imported) {
                        # Fallback to original raw if clean failed
                         $imported = ConvertFrom-Json $raw -ErrorAction SilentlyContinue
                    }
                } else {
                     $imported = $raw
                }

                if ($imported) {
                    if ($imported -is [System.Collections.IEnumerable] -and $imported -isnot [string] -and $imported -isnot [System.Collections.IDictionary]) { 
                        $globalCommands = [System.Collections.ArrayList]$imported 
                    }
                    elseif ($imported -isnot [string]) { 
                        $globalCommands = [System.Collections.ArrayList]@($imported) 
                    }
                }
            } catch {}
        }
        
        # Migration: If global is empty but instance has commands, migrate them
        if ($globalCommands.Count -eq 0 -and $cfg.Contains("SavedCommands_$p")) {
             try {
                $json = $cfg["SavedCommands_$p"]
                $imported = ConvertFrom-Json $json -ErrorAction SilentlyContinue
                if ($imported) {
                    if ($imported -is [System.Collections.IEnumerable] -and $imported -isnot [string]) { $globalCommands = [System.Collections.ArrayList]$imported }
                    else { $globalCommands = [System.Collections.ArrayList]@($imported) }
                }
            } catch {}
        }
        
        $formData.GlobalCommandList = $globalCommands

        # --- FIX: Robust Visible Commands Loading ---
        $visibleIds = [System.Collections.Generic.HashSet[string]]::new()
        if ($cfg.Contains($visibleKey)) {
            try {
                $vRaw = $cfg[$visibleKey]
                if ($vRaw -is [System.Collections.IEnumerable] -and $vRaw -isnot [string]) { $vRaw = $vRaw -join '' }

                $vList = $null
                # Clean escaped quotes explicitly: [\"id\"] -> ["id"]
                if ($vRaw -is [string]) {
                    $cleanVRaw = $vRaw -replace '\\"', '"'
                    $vList = ConvertFrom-Json $cleanVRaw -ErrorAction SilentlyContinue
                    
                    if (-not $vList) {
                        $vList = ConvertFrom-Json $vRaw -ErrorAction SilentlyContinue
                    }
                }

                if ($vList) {
                    # Handle single string result (common failure point)
                    if ($vList -is [string]) {
                        $visibleIds.Add($vList) | Out-Null
                    }
                    elseif ($vList -is [System.Collections.IEnumerable]) {
                        foreach ($id in $vList) { $visibleIds.Add($id) | Out-Null }
                    }
                }
            } catch {}
        } elseif ($cfg.Contains("SavedCommands_$p")) {
             # Migration fallback
             foreach ($cmd in $globalCommands) { if ($cmd.Id) { $visibleIds.Add($cmd.Id) | Out-Null } }
        }

        # Populate Instance SavedCommands based on Visibility
        $formData.SavedCommands = [System.Collections.ArrayList]::new()
        foreach ($cmd in $formData.GlobalCommandList) {
            # Case-insensitive ID check for safety
            if ($cmd.Id -and ($visibleIds.Contains($cmd.Id) -or $visibleIds.Contains($cmd.Id.ToString()))) {
                $formData.SavedCommands.Add($cmd) | Out-Null
                
                # Register Hotkey if present
                if ($null -eq $cmd.PSObject.Properties['Hotkey']) { $cmd | Add-Member -MemberType NoteProperty -Name 'Hotkey' -Value $null -Force }
                if ($null -eq $cmd.PSObject.Properties['HotkeyId']) { $cmd | Add-Member -MemberType NoteProperty -Name 'HotkeyId' -Value $null -Force }
                if ($cmd.Hotkey) {
                    try {
                            $scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($formData.InstanceId)' -CommandId '$($cmd.Id)'")
                            # Use an owner key that includes the instance and the command id so the Hotkey Manager can show a clear action/text
                            $ownerKey = "quickcmd_$($formData.InstanceId)_$($cmd.Id)"
                            # Compute a user-friendly label for display (Name > CommandText > fallback to id)
                            $ownerLabel = $null
                            if ($cmd.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($cmd.Name)) { $ownerLabel = $cmd.Name }
                            elseif ($cmd.PSObject.Properties['CommandText'] -and -not [string]::IsNullOrEmpty($cmd.CommandText)) { $ownerLabel = $cmd.CommandText }
                            else { $ownerLabel = "Invoke Command: $($cmd.Id)" }
                            $cmd.HotkeyId = SetHotkey -KeyCombinationString $cmd.Hotkey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel
                        } catch {
                            $cmd.HotkeyId = $null
                        }
                }
            }
        }
        
        RefreshSavedCommandsUI $formData.Form

    }
    else
    {
        $formData.BtnHotkeyToggle.Checked = $true
    }
    $formData.IsLoading = $false
}

function UpdateChatCommanderSettings
{
    param($formData, [switch]$forceWrite)
    
    if ($formData.IsLoading) { return }
    
    if (-not $global:DashboardConfig.Config.Contains('ChatCommander'))
    { 
        $global:DashboardConfig.Config['ChatCommander'] = [ordered]@{} 
    }
    
    $profilePrefix = $null
    if (Get-Command FindOrCreateProfile -ErrorAction SilentlyContinue)
    {
        $profilePrefix = FindOrCreateProfile $formData.WindowTitle
    }

    if ($profilePrefix)
    {
        $configSuffix = ""
        $parts = $formData.InstanceId -split '_'
        if ($parts.Count -gt 2)
        {
            $configSuffix = "_Sub" + $parts[-1]
        }

        $cfg = $global:DashboardConfig.Config['ChatCommander']
        
        $p = "$profilePrefix$configSuffix"
        $visibleKey = "VisibleCommands_$p"
        
        $cfg["CommandText_$p"] = $formData.TxtCommandText.Text
        # Removed: Interval saving
        $cfg["Name_$p"] = $formData.Name.Text
        $cfg["PosX_$p"] = $formData.PositionSliderX.Value
        $cfg["PosY_$p"] = $formData.PositionSliderY.Value
        # Removed: LoopEnabled, HoldEnabled, HoldInterval, WaitEnabled saving
        $cfg["GlobalHotkey_$p"] = $formData.GlobalHotkey
        $cfg["hotkeys_enabled_$p"] = $formData.BtnHotkeyToggle.Checked

        # Save Global Commands
        $cfg["GlobalCommands"] = (ConvertTo-Json -InputObject $formData.GlobalCommandList -Compress -Depth 5)
        
        # Save Visible IDs
        $visibleIds = [System.Collections.ArrayList]::new()
        foreach ($cmd in $formData.SavedCommands) { if ($cmd.Id) { $visibleIds.Add($cmd.Id) | Out-Null } }
        $cfg[$visibleKey] = (ConvertTo-Json -InputObject $visibleIds -Compress -Depth 2)
        
        # Cleanup old key if exists
        if ($cfg.Contains("SavedCommands_$p")) { $cfg.Remove("SavedCommands_$p") }
        
        # Remove old sequence settings if they exist (now completely obsolete)
        $k = 1
        while ($cfg.Contains("Seq${k}_CommandText_$p"))
        {
            $cfg.Remove("Seq${k}_CommandText_$p")
            $cfg.Remove("Seq${k}_Interval_$p")
            $cfg.Remove("Seq${k}_Name_$p")
            $cfg.Remove("Seq${k}_HoldEnabled_$p")
            $cfg.Remove("Seq${k}_HoldInterval_$p")
            $cfg.Remove("Seq${k}_WaitEnabled_$p")
            $k++
        }
        if ($cfg.Contains("SeqCount_$p")) { $cfg.Remove("SeqCount_$p") }


        if ($forceWrite)
        {
            if (Get-Command WriteConfig -ErrorAction SilentlyContinue) { WriteConfig }
        }
    }
}

#endregion

#region Core Functions

function ChatCommanderSelectedRow
{
    param($row)
    
    if (-not $row -or -not $row.Cells -or $row.Cells.Count -lt 3) { return }
    
    $instanceId = "ChatCommander_" + $row.Cells[2].Value.ToString()

    if (-not $row.Tag -or -not $row.Tag.MainWindowHandle) { return }
    
    $windowHandle = $row.Tag.MainWindowHandle
    
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
    
    $targetWindowRect = New-Object Custom.Native+RECT
    [Custom.Native]::GetWindowRect($windowHandle, [ref]$targetWindowRect)
    
    $windowTitle = if ($row.Tag -and $row.Tag.MainWindowTitle) { $row.Tag.MainWindowTitle } else { $row.Cells[1].Value.ToString() }
    
    $ChatCommanderForm = CreateChatCommanderForm $instanceId $targetWindowRect $windowTitle $row
    
    $global:DashboardConfig.Resources.FtoolForms[$instanceId] = $ChatCommanderForm
    
    $ChatCommanderForm.Show()
    $ChatCommanderForm.BringToFront()
}

function StopChatCommanderForm
{
    param($Form)
    
    if (-not $Form -or $Form.IsDisposed) { return }
    
    try
    {
        $instanceId = $Form.Tag.InstanceId
        if ($instanceId)
        {
            CleanupChatCommanderResources $instanceId
            if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId)) { $global:DashboardConfig.Resources.FtoolForms.Remove($instanceId) }
        }
        $Form.Close()
        $Form.Dispose()
    }
    catch {}
}

function CreateChatCommanderForm
{
    param($instanceId, $targetWindowRect, $windowTitle, $row)
    
    $ChatCommanderForm = New-Object Custom.FtoolFormWindow
    $ChatCommanderForm.Width = 250
    $ChatCommanderForm.Height = 400 # Adjusted initial height for saved commands list
    $ChatCommanderForm.Top = if ($targetWindowRect) { ($targetWindowRect.Top + 30) } else { 200 }
    $ChatCommanderForm.Left = if ($targetWindowRect) { ($targetWindowRect.Right - $ChatCommanderForm.Width - 10) } else { 200 }
    $ChatCommanderForm.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $ChatCommanderForm.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $ChatCommanderForm.Text = "ChatCommander - $instanceId"
    $ChatCommanderForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $ChatCommanderForm.Opacity = 1.0 

    if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon))
    {
        try { $ChatCommanderForm.Icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon) } catch {}
    }

    $ChatCommanderToolTip = New-Object System.Windows.Forms.ToolTip
    $ChatCommanderToolTip.AutoPopDelay = 5000
    $ChatCommanderToolTip.InitialDelay = 100
    $ChatCommanderToolTip.ReshowDelay = 10
    $ChatCommanderToolTip.ShowAlways = $true
    $ChatCommanderToolTip.OwnerDraw = $true
    $ChatCommanderToolTip | Add-Member -MemberType NoteProperty -Name 'TipFont' -Value (New-Object System.Drawing.Font('Segoe UI', 9))
    $ChatCommanderToolTip.Add_Draw({
            $g, $b, $c = $_.Graphics, $_.Bounds, [System.Drawing.Color]
            $g.FillRectangle((New-Object System.Drawing.SolidBrush $c::FromArgb(30,30,30)), $b)
            $g.DrawRectangle((New-Object System.Drawing.Pen $c::FromArgb(100,100,100)), $b.X, $b.Y, $b.Width-1, $b.Height-1)
            $g.DrawString($_.ToolTipText, $this.TipFont, (New-Object System.Drawing.SolidBrush $c::FromArgb(240,240,240)), 3, 3, [System.Drawing.StringFormat]::GenericTypographic)
        })
    $ChatCommanderToolTip.Add_Popup({
            $g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
            $s = $g.MeasureString($this.GetToolTip($_.AssociatedControl), $this.TipFont, [System.Drawing.PointF]::new(0,0), [System.Drawing.StringFormat]::GenericTypographic)
            $g.Dispose(); $_.ToolTipSize = [System.Drawing.Size]::new($s.Width+12, $s.Height+8)
        })

    $previousToolTip = $global:DashboardConfig.UI.ToolTipFtool
    $global:DashboardConfig.UI.ToolTipFtool = $ChatCommanderToolTip

    try
    {

    $headerPanel = SetUIElement -type 'Panel' -visible $true -width 265 -height 20 -top 0 -left 0 -bg @(40, 40, 40)
    $ChatCommanderForm.Controls.Add($headerPanel)

    $labelWinTitle = SetUIElement -type 'Label' -visible $true -width 105 -height 20 -top 5 -left 5 -bg @(40, 40, 40, 0) -fg @(255, 255, 255) -text $row.Cells[1].Value -font (New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Regular))
    $headerPanel.Controls.Add($labelWinTitle)

    $btnInstanceHotkeyToggle = SetUIElement -type 'Label' -visible $true -width 15 -height 15 -top 2 -left 105 -bg @(40, 40, 40) -fg @(255, 255, 255) -text ([char]0x2328) -font (New-Object System.Drawing.Font('Segoe UI', 10)) -tooltip "Set Master Hotkey`nAssign a global hotkey to toggle all hotkeys for this instance."
    $headerPanel.Controls.Add($btnInstanceHotkeyToggle)

    $btnHotkeyToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 15 -top 3 -left 120 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true -tooltip "Toggle Hotkeys`nEnable or disable all hotkeys for this specific ChatCommander instance."
    $headerPanel.Controls.Add($btnHotkeyToggle)

    $btnImport = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 155 -bg @(40, 60, 80) -fg @(255, 255, 255) -text ([char]0x2193) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Import Settings`nLoad settings from a file."
    $headerPanel.Controls.Add($btnImport)

    $btnExport = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 170 -bg @(40, 60, 80) -fg @(255, 255, 255) -text ([char]0x2191) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Export Settings`nSave current settings to a file."
    $headerPanel.Controls.Add($btnExport)

    $btnAddInstance = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 185 -bg @(40, 40, 40) -fg @(255, 255, 255) -text ([char]0x2398) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Add Instance`nCreates another ChatCommander window for this specific client."
    $headerPanel.Controls.Add($btnAddInstance)
        
    $btnShowHide = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 200 -bg @(60, 60, 100) -fg @(255, 255, 255) -text ([char]0x25B2) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Minimize/Expand`nCollapse or expand the ChatCommander window to save screen space."
    $headerPanel.Controls.Add($btnShowHide)
        
    $btnReset = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 215 -bg @(100, 100, 100) -fg @(255, 255, 255) -text ([char]0x21BB) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Reset`nReset all settings for this instance to default."
    $headerPanel.Controls.Add($btnReset)

    $btnClose = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 230 -bg @(150, 20, 20) -fg @(255, 255, 255) -text ([char]0x166D) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Close`nStops the ChatCommander and closes this window."
    $headerPanel.Controls.Add($btnClose)

    # New Command Input Panel
    $panelNewCommand = SetUIElement -type 'Panel' -visible $true -width 200 -height 115 -top 60 -left 40 -bg @(50, 50, 50) # Adjusted height
    $ChatCommanderForm.Controls.Add($panelNewCommand)
    
    $lblCmd = SetUIElement -type 'Label' -visible $true -width 100 -height 15 -top 0 -left 3 -bg @(40, 40, 40, 0) -fg @(200, 200, 200) -text "Command:" -font (New-Object System.Drawing.Font('Segoe UI', 7))
    $panelNewCommand.Controls.Add($lblCmd)

    $txtCommandText = SetUIElement -type 'TextBox' -visible $true -width 194 -height 24 -top 15 -left 3 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Chat Command`nEnter the chat command to send."
    $panelNewCommand.Controls.Add($txtCommandText)
    
    $lblName = SetUIElement -type 'Label' -visible $true -width 100 -height 15 -top 40 -left 3 -bg @(40, 40, 40, 0) -fg @(200, 200, 200) -text "Name:" -font (New-Object System.Drawing.Font('Segoe UI', 7))
    $panelNewCommand.Controls.Add($lblName)

    $name = SetUIElement -type 'TextBox' -visible $true -width 130 -height 17 -top 55 -left 3 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'New' -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular)) -tooltip "Name`nGive this command a descriptive name for easier identification."
    $panelNewCommand.Controls.Add($name)

    $lblVar = SetUIElement -type 'Label' -visible $true -width 100 -height 15 -top 75 -left 3 -bg @(50, 50, 50, 0) -fg @(200, 200, 200) -text "Variable:" -font (New-Object System.Drawing.Font('Segoe UI', 7))
    $panelNewCommand.Controls.Add($lblVar)

    $txtVariable = SetUIElement -type 'TextBox' -visible $true -width 130 -height 17 -top 90 -left 3 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '' -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular)) -tooltip "Variable`nReplaces [var] in the command text with this value."
    $panelNewCommand.Controls.Add($txtVariable)

    # Removed: Interval textbox
    
    $btnAdd = SetUIElement -type 'Button' -visible $true -width 55 -height 20 -top 88 -left 138 -bg @(40, 80, 80) -fg @(255, 255, 255) -text ([char]0x2795) + ' Add' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Add Command`nAdds the current chat command to the saved list."
    $panelNewCommand.Controls.Add($btnAdd)

    # Removed: lblLoop, lblWait, lblHold

    $btnRunSequence = SetUIElement -type 'Button' -visible $true -width 200 -height 25 -top ($panelNewCommand.Bottom + 5) -left 40 -bg @(0, 120, 215) -fg @(255, 255, 255) -text "Run All Visible" -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Run Sequence`nExecutes all currently visible commands in the list from top to bottom."
    $ChatCommanderForm.Controls.Add($btnRunSequence)

    # Search bar for saved commands
    $lblSearch = SetUIElement -type 'Label' -visible $true -width 100 -height 15 -top ($btnRunSequence.Bottom + 5) -left 40 -bg @(40, 40, 40, 0) -fg @(200, 200, 200) -text "Search:" -font (New-Object System.Drawing.Font('Segoe UI', 7))
    $ChatCommanderForm.Controls.Add($lblSearch)

    $txtSearchCommands = SetUIElement -type 'TextBox' -visible $true -width 155 -height 24 -top ($lblSearch.Bottom + 0) -left 40 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Search Commands`nType to filter saved chat commands across all profiles."
    $ChatCommanderForm.Controls.Add($txtSearchCommands)

    $btnShowAll = SetUIElement -type 'Button' -visible $true -width 40 -height 24 -top ($lblSearch.Bottom + 0) -left 200 -bg @(60, 60, 100) -fg @(255, 255, 255) -text "ALL" -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 7)) -tooltip "Show All`nToggle table view of all commands."
    $ChatCommanderForm.Controls.Add($btnShowAll)

    # FlowLayoutPanel for saved commands
    $panelSavedCommands = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelSavedCommands.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $panelSavedCommands.WrapContents = $false
    $panelSavedCommands.AutoSize = $true
    $panelSavedCommands.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $panelSavedCommands.Width = 200
    $panelSavedCommands.Top = ($txtSearchCommands.Bottom + 5)
    $panelSavedCommands.Left = 40
    $panelSavedCommands.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $ChatCommanderForm.Controls.Add($panelSavedCommands)

    # Pagination Panel
    $panelPagination = SetUIElement -type 'Panel' -visible $true -width 200 -height 30 -top 0 -left 40 -bg @(50, 50, 50)
    $ChatCommanderForm.Controls.Add($panelPagination)

    $btnPrevPage = SetUIElement -type 'Button' -visible $true -width 30 -height 20 -top 5 -left 5 -bg @(40, 40, 40) -fg @(255, 255, 255) -text "<" -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8))
    $panelPagination.Controls.Add($btnPrevPage)

    $lblPageInfo = SetUIElement -type 'Label' -visible $true -width 120 -height 20 -top 8 -left 40 -bg @(50, 50, 50, 0) -fg @(200, 200, 200) -text "1 / 1" -font (New-Object System.Drawing.Font('Segoe UI', 8))
    $lblPageInfo.TextAlign = 'MiddleCenter'
    $panelPagination.Controls.Add($lblPageInfo)

    $btnNextPage = SetUIElement -type 'Button' -visible $true -width 30 -height 20 -top 5 -left 165 -bg @(40, 40, 40) -fg @(255, 255, 255) -text ">" -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8))
    $panelPagination.Controls.Add($btnNextPage)

    # DataGridView for Show All mode
    $gridAllCommands = SetUIElement -type 'DataGridView' -visible $false -width 400 -height 300 -top ($txtSearchCommands.Bottom + 5) -left 40 -bg @(40, 40, 40) -fg @(255, 255, 255)
    $gridAllCommands.AllowUserToAddRows = $false
    $gridAllCommands.RowHeadersVisible = $false
    $gridAllCommands.SelectionMode = 'FullRowSelect'
    $gridAllCommands.MultiSelect = $false
    $gridAllCommands.AutoSizeColumnsMode = 'Fill'
    
    $colVis = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colVis.HeaderText = "Show"
    $colVis.FillWeight = 20
    $colVis.ReadOnly = $false

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = "Name"
    $colName.ReadOnly = $true
    $colName.FillWeight = 100
    
    $colCmd = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colCmd.HeaderText = "Command"
    $colCmd.ReadOnly = $true
    $colCmd.FillWeight = 100

    $colExec = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $colExec.HeaderText = "Run"
    $colExec.Text = "Run"
    $colExec.UseColumnTextForButtonValue = $true
    $colExec.FillWeight = 20
    $colExec.FlatStyle = 'Flat'

    $colEdit = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $colEdit.HeaderText = "Edit"
    $colEdit.Text = "Edit"
    $colEdit.UseColumnTextForButtonValue = $true
    $colEdit.FillWeight = 20
    $colEdit.FlatStyle = 'Flat'

    $colDel = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $colDel.HeaderText = "Del"
    $colDel.Text = "X"
    $colDel.UseColumnTextForButtonValue = $true
    $colDel.FillWeight = 15
    $colDel.FlatStyle = 'Flat'
    $colDel.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Red

    $gridAllCommands.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colVis, $colName, $colCmd, $colExec, $colEdit, $colDel))
    $ChatCommanderForm.Controls.Add($gridAllCommands)

    # Handler for Visibility Checkbox
    $gridAllCommands.Add_CellContentClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -eq 0) { # Visibility Column
            $grid = $s
            $form = $grid.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag
            
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            $isChecked = $grid.Rows[$e.RowIndex].Cells[0].Value
            $cmd = $grid.Rows[$e.RowIndex].Tag
            
            if ($isChecked) {
                # Add to SavedCommands if not present
                $exists = $false
                foreach ($c in $data.SavedCommands) { if ($c.Id -eq $cmd.Id) { $exists = $true; break } }
                
                if (-not $exists) {
                    $data.SavedCommands.Add($cmd) | Out-Null
                    # Register Hotkey if present
                    if ($cmd.Hotkey) {
                        try {
                            $scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($data.InstanceId)' -CommandId '$($cmd.Id)'")
                            $ownerKey = "quickcmd_$($data.InstanceId)_$($cmd.Id)"
                            $ownerLabel = if ($cmd.Name) { $cmd.Name } else { $cmd.CommandText }
                            $cmd.HotkeyId = SetHotkey -KeyCombinationString $cmd.Hotkey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel
                            Write-Verbose "QuickCommand: Registered hotkey $($cmd.Hotkey) for command $($cmd.Name)"
                        } catch { 
                            $cmd.HotkeyId = $null 
                            Write-Warning "QuickCommand: Failed to register hotkey $($cmd.Hotkey) for command $($cmd.Name): $_"
                        }
                    }
                }
            } else {
                # Remove from SavedCommands
                $toRemove = $null
                foreach ($c in $data.SavedCommands) { if ($c.Id -eq $cmd.Id) { $toRemove = $c; break } }
                if ($toRemove) {
                    $data.SavedCommands.Remove($toRemove)
                    if ($toRemove.HotkeyId) {
                        try { UnregisterHotkeyInstance -Id $toRemove.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($toRemove.Id)" } catch {}
                        try { UnregisterHotkeyInstance -Id $toRemove.HotkeyId -OwnerKey $data.InstanceId } catch {}
                        $toRemove.HotkeyId = $null
                        Write-Verbose "QuickCommand: Unregistered hotkey for command $($toRemove.Name)"
                    }
                }
            }
            UpdateChatCommanderSettings $data -forceWrite
            if (Get-Command RefreshHotkeysList -ErrorAction SilentlyContinue) { try { RefreshHotkeysList } catch {} }
        }
    })

    $positionSliderY = New-Object System.Windows.Forms.TrackBar
    $positionSliderY.Orientation = 'Vertical'
    $positionSliderY.Minimum = -18
    $positionSliderY.Maximum = 118
    $positionSliderY.TickFrequency = 300
    $positionSliderY.Value = 100
    $positionSliderY.Size = New-Object System.Drawing.Size(15, 110)
    $positionSliderY.Location = New-Object System.Drawing.Point(5, 20)
    $ChatCommanderForm.Controls.Add($positionSliderY)
        
    $positionSliderX = New-Object System.Windows.Forms.TrackBar
    $positionSliderX.Minimum = -25
    $positionSliderX.Maximum = 125
    $positionSliderX.TickFrequency = 300
    $positionSliderX.Value = 100
    $positionSliderX.Size = New-Object System.Drawing.Size(190, 15)
    $positionSliderX.Location = New-Object System.Drawing.Point(45, 25)
    $ChatCommanderForm.Controls.Add($positionSliderX)

    }
    finally
    {
        $global:DashboardConfig.UI.ToolTipFtool = $previousToolTip
    }

    $formData = [PSCustomObject]@{
        Type            = 'ChatCommander'
        InstanceId      = $instanceId
        SelectedWindow  = $row.Tag.MainWindowHandle
        HeaderPanel     = $headerPanel
        PanelNewCommand = $panelNewCommand
        TxtVariable     = $txtVariable
        TxtCommandText  = $txtCommandText
        # Removed: Interval
        Name            = $name # For new command input
        # Removed: BtnLoopToggle, BtnWaitToggle, BtnHoldKeyToggle, TxtHoldKeyInterval
        BtnHotkeyToggle = $btnHotkeyToggle
        BtnInstanceHotkeyToggle = $btnInstanceHotkeyToggle
        BtnAdd          = $btnAdd # To add new command to saved list / Save Changes
        BtnExport       = $btnExport
        BtnImport       = $btnImport
        BtnAddInstance  = $btnAddInstance
        BtnClose        = $btnClose
        BtnShowHide     = $btnShowHide
        BtnReset        = $btnReset
        BtnRunSequence  = $btnRunSequence
        BtnShowAll      = $btnShowAll
        BtnPrevPage     = $btnPrevPage
        BtnNextPage     = $btnNextPage
        LblPageInfo     = $lblPageInfo
        PanelPagination = $panelPagination
        GridAllCommands = $gridAllCommands
        PositionSliderX = $positionSliderX
        PositionSliderY = $positionSliderY
        Form            = $ChatCommanderForm
        RunningTimer    = $null
        WindowTitle     = $windowTitle
        Row             = $row
        ToolTipFtool    = $ChatCommanderToolTip
        IsCollapsed     = $false
        OriginalHeight  = $ChatCommanderForm.Height # Initial height
        HotkeyId        = $null
        GlobalHotkeyId  = $null
        GlobalHotkey    = $null
        SavedCommands   = [System.Collections.ArrayList]::new() # New: List of saved command objects
        GlobalCommandList = [System.Collections.ArrayList]::new() # New: Global list
        txtSearchCommands = $txtSearchCommands # New: Search bar
        panelSavedCommands = $panelSavedCommands # New: FlowLayoutPanel for saved commands
        ActiveSteps     = @()
        CurrentStepIndex= 0
        ActiveJobs      = [System.Collections.ArrayList]::new()
        RunspacePool    = (InitChatCommanderRunspace -InstanceId $instanceId)
        IsLoading       = $false
        EditingCommandId = $null # New: To track which command is being edited
        CurrentPage     = 1
        ItemsPerPage    = 30
        ShowAllMode     = $false
    }
    
    $ChatCommanderForm.Tag = $formData
    
    CreateChatCommanderPositionTimer $formData | Out-Null
    LoadChatCommanderSettings $formData | Out-Null
    ToggleChatCommanderHotkeys -InstanceId $formData.InstanceId -ToggleState $formData.BtnHotkeyToggle.Checked | Out-Null

    if (-not [string]::IsNullOrEmpty($formData.GlobalHotkey))
    {
        try
        {
            $ownerKey = "global_toggle_$($formData.InstanceId)"
            $script = @"
if (`$global:DashboardConfig.Resources.FtoolForms.Contains('$($formData.InstanceId)')) {
    `$f = `$global:DashboardConfig.Resources.FtoolForms['$($formData.InstanceId)']
    if (`$f -and -not `$f.IsDisposed -and `$f.Tag) {
        `$toggle = `$f.Tag.BtnHotkeyToggle
        if (`$toggle) {
            if (`$toggle.InvokeRequired) { `$toggle.Invoke([System.Action]{ `$toggle.Checked = -not `$toggle.Checked }) } else { `$toggle.Checked = -not `$toggle.Checked }
            ToggleChatCommanderHotkeys -InstanceId '$($formData.InstanceId)' -ToggleState `$toggle.Checked
        }
    }
}
"@
            $scriptBlock = [scriptblock]::Create($script)
            # Friendly label for UI
            $ownerLabel = $null
            if ($formData.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($formData.Name)) { $ownerLabel = $formData.Name }
            elseif ($formData.PSObject.Properties['WindowTitle'] -and -not [string]::IsNullOrEmpty($formData.WindowTitle)) { $ownerLabel = $formData.WindowTitle }
            else { $ownerLabel = "Chat Commander: $($formData.InstanceId)" }
            $formData.GlobalHotkeyId = SetHotkey -KeyCombinationString $formData.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel
        }
        catch
        {
            $formData.GlobalHotkeyId = $null
        }
    }

    AddChatCommanderEventHandlers $formData | Out-Null
    
    return $ChatCommanderForm
}

function CreateSavedCommandEntryUI
{
    param($formData, $cmd)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(190, 26)
    $panel.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $panel.Tag = $cmd 
    $panel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = if ($cmd.Name) { $cmd.Name } else { $cmd.CommandText }
    $lbl.ForeColor = [System.Drawing.Color]::White
    $lbl.Location = New-Object System.Drawing.Point(3, 4)
    $lbl.Size = New-Object System.Drawing.Size(85, 18)
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lbl.AutoEllipsis = $true
    $lbl.Cursor = 'Hand'
    $panel.Controls.Add($lbl)
    
    $toolTip = $formData.ToolTipFtool
    if ($toolTip) { $toolTip.SetToolTip($lbl, "$($cmd.Name)`n$($cmd.CommandText)`nClick to Edit") }

    # Run Button
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text =  [char]0x25B6
    $btnRun.FlatStyle = 'Flat'
    $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(40, 80, 40)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Location = New-Object System.Drawing.Point(90, 1)
    $btnRun.Size = New-Object System.Drawing.Size(24, 24)
    $btnRun.Cursor = 'Hand'
    if ($toolTip) { $toolTip.SetToolTip($btnRun, "Execute Command") }
    $panel.Controls.Add($btnRun)

    # Hotkey Button
    $btnKey = New-Object System.Windows.Forms.Button
    $btnKey.Text = if ($cmd.Hotkey) { $cmd.Hotkey } else { "Key" }
    $btnKey.FlatStyle = 'Flat'
    $btnKey.FlatAppearance.BorderSize = 0
    $btnKey.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 80)
    $btnKey.ForeColor = [System.Drawing.Color]::White
    $btnKey.Location = New-Object System.Drawing.Point(116, 1)
    $btnKey.Size = New-Object System.Drawing.Size(45, 24)
    $btnKey.Font = New-Object System.Drawing.Font('Segoe UI', 6)
    $btnKey.Cursor = 'Hand'
    if ($toolTip) { $toolTip.SetToolTip($btnKey, "Assign Hotkey") }
    $panel.Controls.Add($btnKey)

    # Delete Button
    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = "X"
    $btnDel.FlatStyle = 'Flat'
    $btnDel.FlatAppearance.BorderSize = 0
    $btnDel.BackColor = [System.Drawing.Color]::FromArgb(80, 40, 40)
    $btnDel.ForeColor = [System.Drawing.Color]::White
    $btnDel.Location = New-Object System.Drawing.Point(163, 1)
    $btnDel.Size = New-Object System.Drawing.Size(24, 24)
    $btnDel.Cursor = 'Hand'
    if ($toolTip) { $toolTip.SetToolTip($btnDel, "Delete Command") }
    $panel.Controls.Add($btnDel)

    # Event Handlers
    $btnRun.Add_Click({
        $form = $this.FindForm()
        if ($form -and $form.Tag) {
            $data = $form.Tag
            if ($data.RunningTimer) { StopChatCommanderSequence $data }
            $job = Invoke-ChatCommanderAction -formData $data -commandString $cmd.CommandText
            if ($job) {
                $data.ActiveJobs.Add($job) | Out-Null
                [System.Threading.Tasks.Task]::Run([System.Action]{
                    $job.AsyncResult.AsyncWaitHandle.WaitOne() | Out-Null
                    try { $job.PowerShell.EndInvoke($job.AsyncResult); $job.PowerShell.Dispose() } catch {}
                    if ($form -and -not $form.IsDisposed) {
                        $form.BeginInvoke([System.Action]{
                            if ($data.ActiveJobs) { $data.ActiveJobs.Remove($job) }
                        })
                    }
                }) | Out-Null
            }
        }
    }.GetNewClosure())

    $btnKey.Add_Click({
        $form = $this.FindForm()
        if ($form -and $form.Tag) {
            $data = $form.Tag
            $currentKey = if ($cmd.Hotkey) { $cmd.Hotkey } else { $null }
            $newKey = Show-KeyCaptureDialog $currentKey -OwnerForm $form
            
            if ($newKey -ne $currentKey) {
                if ($cmd.HotkeyId) {
                    # Try to unregister both new-style ownerKey and legacy ownerKey (instance-only)
                    try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($cmd.Id)" } catch {}
                }

                $cmd.Hotkey = $newKey
                $cmd.HotkeyId = $null

                if ($newKey) {
                    try {
                        $scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($data.InstanceId)' -CommandId '$($cmd.Id)'")
                        $ownerKey = "quickcmd_$($data.InstanceId)_$($cmd.Id)"
                        # Compute friendly label for UI (Name > CommandText > fallback id)
                        $ownerLabel = $null
                        if ($cmd.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($cmd.Name)) { $ownerLabel = $cmd.Name }
                        elseif ($cmd.PSObject.Properties['CommandText'] -and -not [string]::IsNullOrEmpty($cmd.CommandText)) { $ownerLabel = $cmd.CommandText }
                        else { $ownerLabel = "Invoke Command: $($cmd.Id)" }
                        $cmd.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel
                    } catch {}
                }
                
                $this.Text = if ($newKey) { $newKey } else { "Key" }
                UpdateChatCommanderSettings $data -forceWrite
                if (Get-Command RefreshHotkeysList -ErrorAction SilentlyContinue) { try { RefreshHotkeysList } catch {} }
            }
        }
    }.GetNewClosure())

    $btnDel.Add_Click({
        $form = $this.FindForm()
        if ($form -and $form.Tag) {
            $data = $form.Tag
            if ($cmd.HotkeyId) { try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($cmd.Id)" } catch {} }
            $data.SavedCommands.Remove($cmd)
            RefreshSavedCommandsUI $form
            UpdateChatCommanderSettings $data -forceWrite
            if (Get-Command RefreshHotkeysList -ErrorAction SilentlyContinue) { try { RefreshHotkeysList } catch {} }
        }
    }.GetNewClosure())
    
    $lbl.Add_Click({
        $form = $this.FindForm()
        if ($form -and $form.Tag) {
            $data = $form.Tag
            $data.TxtCommandText.Text = $cmd.CommandText
            $data.Name.Text = $cmd.Name
            $data.EditingCommandId = $cmd.Id
            $data.BtnAdd.Text = ([char]0x2705) + ' Save'
            $data.BtnAdd.BackColor = [System.Drawing.Color]::FromArgb(40, 120, 40)
        }
    }.GetNewClosure())

    return $panel
}

function GetAllSavedCommands
{
    param($currentInstanceId)
    # Now using GlobalCommandList from formData, so this helper might be redundant or just return the global list
    # But since we are inside the form context usually, we can access formData directly.
    # If called from outside, we'd need to look at config.
    
    if ($global:DashboardConfig.Config.Contains('ChatCommander') -and $global:DashboardConfig.Config['ChatCommander'].Contains('GlobalCommands')) {
        try {
            return ConvertFrom-Json $global:DashboardConfig.Config['ChatCommander']['GlobalCommands']
        } catch { return @() }
    }
    return @()
}

function RefreshSavedCommandsUI
{
    param($form)
    $formData = $form.Tag

    $searchTerm = $formData.txtSearchCommands.Text.ToLower()
    $isSearching = -not [string]::IsNullOrWhiteSpace($searchTerm)

    # Get source list - In ShowAllMode we show Global, otherwise Visible (SavedCommands)
    $sourceList = if ($formData.ShowAllMode) { $formData.GlobalCommandList } else { $formData.SavedCommands }
    
    # Filter
    $filteredList = [System.Collections.ArrayList]::new()
    foreach ($cmd in $sourceList) {
        if ($null -eq $cmd) { continue }
        $cText = if ($cmd.PSObject.Properties['CommandText']) { $cmd.CommandText } else { '' }
        $cName = if ($cmd.PSObject.Properties['Name']) { $cmd.Name } else { '' }
        
        if ([string]::IsNullOrWhiteSpace($cText)) { continue }

        if (-not $isSearching -or 
            ($cName.ToLower().Contains($searchTerm)) -or 
            ($cText.ToLower().Contains($searchTerm))) {
            $filteredList.Add($cmd)
        }
    }

    if ($formData.ShowAllMode) {
        # Table Mode
        $formData.panelSavedCommands.Visible = $false
        $formData.PanelPagination.Visible = $false
        $formData.GridAllCommands.Visible = $true
        
        $formData.GridAllCommands.Rows.Clear()
        foreach ($cmd in $filteredList) {
            $isVisible = $false
            foreach ($c in $formData.SavedCommands) {
                if ($c.Id -eq $cmd.Id) { $isVisible = $true; break }
            }
            
            $idx = $formData.GridAllCommands.Rows.Add($isVisible, $cmd.Name, $cmd.CommandText)
            $formData.GridAllCommands.Rows[$idx].Tag = $cmd
        }
        
        # Resize form for table view
        AdjustChatCommanderFormHeight $form
    } else {
        # List Mode with Pagination
        $formData.GridAllCommands.Visible = $false
        $formData.panelSavedCommands.Visible = $true
        $formData.PanelPagination.Visible = $true
        
        $formData.panelSavedCommands.SuspendLayout()
        $formData.panelSavedCommands.Controls.Clear()
        
        $totalItems = $filteredList.Count
        $totalPages = [Math]::Ceiling($totalItems / $formData.ItemsPerPage)
        if ($totalPages -lt 1) { $totalPages = 1 }
        
        if ($formData.CurrentPage -gt $totalPages) { $formData.CurrentPage = $totalPages }
        if ($formData.CurrentPage -lt 1) { $formData.CurrentPage = 1 }
        
        $startIndex = ($formData.CurrentPage - 1) * $formData.ItemsPerPage
        
        $count = 0
        for ($i = $startIndex; $i -lt $totalItems; $i++) {
            if ($count -ge $formData.ItemsPerPage) { break }
            $cmd = $filteredList[$i]
            $entryUI = CreateSavedCommandEntryUI $formData $cmd
            $formData.panelSavedCommands.Controls.Add($entryUI)
            $count++
        }
        
        $formData.panelSavedCommands.ResumeLayout()
        
        # Update Pagination Controls
        $formData.LblPageInfo.Text = "$($formData.CurrentPage) / $totalPages"
        $formData.BtnPrevPage.Enabled = ($formData.CurrentPage -gt 1)
        $formData.BtnNextPage.Enabled = ($formData.CurrentPage -lt $totalPages)
        
        # Resize form for list view
        AdjustChatCommanderFormHeight $form
        
        # Position Pagination Panel
        $formData.PanelPagination.Top = $formData.panelSavedCommands.Bottom + 5
    }
}

function AdjustChatCommanderFormHeight # Renamed from RepositionChatCommandPanels
{
    param($form)
    
    if (-not $form -or $form.IsDisposed) { return }
    
    $form.SuspendLayout()
    try
    {
        $data = $form.Tag
        
        if ($data.ShowAllMode)
        {
            $form.Width = 500
            $data.HeaderPanel.Width = 500
            
            # Move Header Buttons to right
            $data.BtnClose.Left = 480
            $data.BtnReset.Left = 465
            $data.BtnShowHide.Left = 450
            $data.BtnAddInstance.Left = 435
            $data.BtnExport.Left = 420
            $data.BtnImport.Left = 405
            # Keep toggles on left or move them? Let's keep them relative or move them too?
            # Original: Toggle at 120. Let's keep toggles on left side to avoid gap issues with title.
            
            # Resize Main Controls
            $data.PanelNewCommand.Width = 440
            $data.TxtCommandText.Width = 434
            $data.BtnAdd.Left = 382
            $data.Name.Width = 374
            
            $data.BtnRunSequence.Width = 440
            $data.BtnShowAll.Left = 440
            $data.txtSearchCommands.Width = 395
            
            $data.GridAllCommands.Width = 440
            $data.GridAllCommands.Height = 300 # Default height, will be adjusted below if needed
            
            # In Show All mode, we use a larger fixed size or max available
            $screen = [System.Windows.Forms.Screen]::FromPoint($form.Location)
            $workingArea = $screen.WorkingArea
            $distToBottom = $workingArea.Bottom - $form.Top
            $targetHeight = 600
            if ($targetHeight -gt $distToBottom - 20) { $targetHeight = $distToBottom - 20 }
            if ($targetHeight -lt 400) { $targetHeight = 400 }
            
            $form.Height = $targetHeight
            $data.OriginalHeight = $targetHeight
            $data.PositionSliderY.Height = $targetHeight - 20
            
            $data.GridAllCommands.Height = $targetHeight - $data.GridAllCommands.Top - 20
            return
        }

        # Restore Standard Width
        $form.Width = 250
        $data.HeaderPanel.Width = 265
        
        # Restore Header Buttons
        $data.BtnClose.Left = 230
        $data.BtnReset.Left = 215
        $data.BtnShowHide.Left = 200
        $data.BtnAddInstance.Left = 185
        $data.BtnExport.Left = 170
        $data.BtnImport.Left = 155
        
        # Restore Main Controls
        $data.PanelNewCommand.Width = 200
        $data.TxtCommandText.Width = 194
        $data.BtnAdd.Left = 138
        $data.Name.Width = 130
        
        $data.BtnRunSequence.Width = 200
        $data.BtnShowAll.Left = 200
        $data.txtSearchCommands.Width = 155
        
        $data.GridAllCommands.Width = 400 # Hidden anyway

        # Standard Mode (Pagination)
        $baseHeight = 285 # Adjusted base height for the top section (header + new command input + search)
        $panel = $data.panelSavedCommands
        
        # Temporarily enable AutoSize to get preferred height
        $panel.AutoSize = $true
        $panel.PerformLayout()
        $preferredPanelHeight = $panel.Height
        
        # Add space for pagination controls if visible
        $paginationHeight = if ($data.PanelPagination.Visible) { 0 } else { 0 }

        $newHeight = $baseHeight + $preferredPanelHeight + $paginationHeight + 10 # 10 for padding
        
        $screen = [System.Windows.Forms.Screen]::FromPoint($form.Location)
        $workingArea = $screen.WorkingArea
        $distToBottom = $workingArea.Bottom - $form.Top
        $maxHeight = $distToBottom - 20
        if ($maxHeight -lt 285) { $maxHeight = 285 } # Minimum usable height

        if ($newHeight -gt $maxHeight) {
            $newHeight = $maxHeight
            # Constrain panel
            $panel.AutoSize = $false
            $panel.Height = $maxHeight - $baseHeight - $paginationHeight - 10
            $panel.AutoScroll = $true
        } else {
            # Let panel grow
            $panel.AutoSize = $true
            $panel.AutoScroll = $false
        }

        if (-not $data.IsCollapsed)
        {
            $finalHeight = $newHeight
            $form.Height = $finalHeight
            $data.OriginalHeight = $finalHeight
            $data.PositionSliderY.Height = $finalHeight - 20
            
            if ($data.PanelPagination.Visible) {
                $data.PanelPagination.Top = $panel.Bottom + 5
            }
        }
    }
    catch { Write-Verbose "AdjustHeight Error: $_" }
    finally
    {
        $form.ResumeLayout()
    }
}

function AddChatCommanderEventHandlers
{
    param($formData)

    $formData.TxtCommandText.Add_TextChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        # If editing, update the 'Save' button state
        if ($data.EditingCommandId) {
            $data.BtnAdd.Text = ([char]0x2705) + ' Save'
            $data.BtnAdd.BackColor = [System.Drawing.Color]::FromArgb(40, 120, 40)
        }
    })

    $formData.BtnAdd.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag

        $newCommandText = $data.TxtCommandText.Text.Trim()
        $newCommandName = $data.Name.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($newCommandText))
        {
            Show-DarkMessageBox 'Command text cannot be empty.' 'Add/Save Command Error' 'Ok' 'Error'
            return
        }
        if ([string]::IsNullOrWhiteSpace($newCommandName))
        {
            $newCommandName = "Command $($data.SavedCommands.Count + 1)"
        }

        # Check for duplicates in Global List (by Name)
        $existing = $data.GlobalCommandList | Where-Object { $_.Name -eq $newCommandName } | Select-Object -First 1
        
        if ($existing -and ($null -eq $data.EditingCommandId -or $existing.Id -ne $data.EditingCommandId)) {
             $newName = ShowInputBox -Title "Duplicate Name" -Prompt "A command with this name already exists.`nPlease enter a unique name:" -DefaultText "$newCommandName (1)"
             if ([string]::IsNullOrWhiteSpace($newName)) { return }
             $newCommandName = $newName
             # Re-check
             $existing2 = $data.GlobalCommandList | Where-Object { $_.Name -eq $newCommandName } | Select-Object -First 1
             if ($existing2) {
                Show-DarkMessageBox "Name still duplicate. Action cancelled." "Error" "OK" "Error"
                return
             }
        }

        if ($data.EditingCommandId)
        {
            # Editing existing command
            $commandToUpdate = $data.GlobalCommandList | Where-Object { $_.Id -eq $data.EditingCommandId } | Select-Object -First 1
            if ($commandToUpdate) {
                $commandToUpdate.CommandText = $newCommandText
                $commandToUpdate.Name = $newCommandName
            }
            # Also update in SavedCommands if present
            $savedCmd = $data.SavedCommands | Where-Object { $_.Id -eq $data.EditingCommandId } | Select-Object -First 1
            if ($savedCmd) { $savedCmd.CommandText = $newCommandText; $savedCmd.Name = $newCommandName }

            $data.EditingCommandId = $null # Reset editing state
            $data.BtnAdd.Text = ([char]0x2795) + ' Add' # Change button text back to 'Add'
            $data.BtnAdd.BackColor = [System.Drawing.Color]::FromArgb(40, 80, 80) # Reset color
        }
        else
        {
            # Adding new command
            $newCommand = [PSCustomObject]@{
                Id = [Guid]::NewGuid().ToString()
                CommandText = $newCommandText
                Name = $newCommandName
                Hotkey = $null
                HotkeyId = $null
            }
            $data.GlobalCommandList.Add($newCommand) | Out-Null
            $data.SavedCommands.Add($newCommand) | Out-Null
        }
        
        $data.TxtCommandText.Text = "" # Clear input after adding/saving
        $data.Name.Text = "New" # Reset name

        RefreshSavedCommandsUI $form
        UpdateChatCommanderSettings $data -forceWrite
    })

    $formData.BtnShowHide.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag

        if ($data.IsCollapsed)
        {
            $form.Height = $data.OriginalHeight
            $data.IsCollapsed = $false
            $this.Text = ([char]0x25B2) # Up arrow
        }
        else
        {
            $data.OriginalHeight = $form.Height
            $form.Height = 26
            $data.IsCollapsed = $true
            $this.Text = ([char]0x25BC) # Down arrow
        }
    })

    $formData.BtnRunSequence.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        StartChatCommanderSequence $form.Tag -RunSavedSequence
    })

    $formData.BtnClose.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $null = $data; $data = $form.Tag

        StopChatCommanderForm $form
    })

    $formData.BtnReset.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        
        if ((Show-DarkMessageBox "Are you sure you want to reset this ChatCommander instance to default settings? This will remove all commands and clear settings." "Reset Defaults" "YesNo" "Warning") -eq 'Yes')
        {
            # Stop ChatCommander if running
            if ($data.RunningTimer) {
                StopChatCommanderSequence $data
            }

            # Clear Saved Commands
            foreach ($cmd in $data.SavedCommands) {
                if ($cmd.HotkeyId) {
                    try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($cmd.Id)" } catch {}
                    try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey $data.InstanceId } catch {}
                }
            }
            $data.SavedCommands.Clear()
            # Note: We do NOT clear GlobalCommandList on reset, only the instance visibility (SavedCommands)

            # Unregister Global Hotkeys if any
            if ($data.GlobalHotkeyId) {
                 $ownerKey = "global_toggle_$($data.InstanceId)"
                 try { UnregisterHotkeyInstance -Id $data.GlobalHotkeyId -OwnerKey $ownerKey } catch {}
                 $data.GlobalHotkeyId = $null
                 $data.GlobalHotkey = $null
            }
            RefreshSavedCommandsUI $form

            # Reset Controls
            $data.TxtCommandText.Text = ""
            # Removed: Interval reset
            $data.Name.Text = "New"
            $data.BtnHotkeyToggle.Checked = $true
            $data.PositionSliderX.Value = 100
            $data.PositionSliderY.Value = 100
            $data.EditingCommandId = $null
            $data.BtnAdd.Text = ([char]0x2795) + ' Add'
            $data.BtnAdd.BackColor = [System.Drawing.Color]::FromArgb(40, 80, 80)

            # Clear Config specific to this profile
            $profilePrefix = $null
            if (Get-Command FindOrCreateProfile -ErrorAction SilentlyContinue)
            {
                $profilePrefix = FindOrCreateProfile $data.WindowTitle
            }

            if ($profilePrefix) {
                $configSuffix = ""
                $parts = $data.InstanceId -split '_'
                if ($parts.Count -gt 2)
                {
                    $configSuffix = "_Sub" + $parts[-1]
                }
                $p = "$profilePrefix$configSuffix"
                
                $keysToClear = [System.Collections.ArrayList]@()
                foreach ($k in $global:DashboardConfig.Config['ChatCommander'].Keys) {
                    if ($k -match "_$p$") {
                        $keysToClear.Add($k)
                    }
                }
                foreach ($k in $keysToClear) {
                    $global:DashboardConfig.Config['ChatCommander'].Remove($k)
                }
                if ($global:DashboardConfig.Config['ChatCommander'].Contains("GlobalHotkey_$p")) { $global:DashboardConfig.Config['ChatCommander'].Remove("GlobalHotkey_$p") }
            }

            # Save clean state
            UpdateChatCommanderSettings $data -forceWrite
            AdjustChatCommanderFormHeight $form
        }
    })

    $formData.BtnExport.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "JSON Files (*.json)|*.json"
        $sfd.FileName = "ChatCommands_$($data.InstanceId).json"
        
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $json = $data.SavedCommands | ConvertTo-Json -Depth 5
                $json | Set-Content -Path $sfd.FileName -Encoding UTF8
                Show-DarkMessageBox "Commands exported successfully." "Export" "Ok" "Information"
            } catch {
                Show-DarkMessageBox "Failed to export commands: $_" "Export Error" "Ok" "Error"
            }
        }
    })

    $formData.BtnImport.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag

        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "JSON Files (*.json)|*.json"
        
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $json = Get-Content -Path $ofd.FileName -Raw -Encoding UTF8
                $imported = [System.Collections.ArrayList]@(ConvertFrom-Json $json)
                
                if ($imported) {
                    # Clear existing hotkeys before overwriting
                    foreach ($cmd in $data.SavedCommands) {
                        if ($cmd.HotkeyId) {
                            try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($cmd.Id)" } catch {}
                            try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey $data.InstanceId } catch {}
                        }
                    }
                    
                    $data.SavedCommands = $imported
                    
                    # Re-register imported hotkeys
                    foreach ($cmd in $data.SavedCommands) { 
                        if ($null -eq $cmd.PSObject.Properties['Hotkey']) { $cmd | Add-Member -MemberType NoteProperty -Name 'Hotkey' -Value $null -Force }
                        if ($null -eq $cmd.PSObject.Properties['HotkeyId']) { $cmd | Add-Member -MemberType NoteProperty -Name 'HotkeyId' -Value $null -Force }
                        
                        $cmd.HotkeyId = $null # Reset ID as it's from file

                                    if ($cmd.Hotkey) {
                                         try {
                                            $scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($data.InstanceId)' -CommandId '$($cmd.Id)'")
                                            $ownerKey = "quickcmd_$($data.InstanceId)_$($cmd.Id)"
                                            # Friendly label for UI
                                            $ownerLabel = $null
                                            if ($cmd.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($cmd.Name)) { $ownerLabel = $cmd.Name }
                                            elseif ($cmd.PSObject.Properties['CommandText'] -and -not [string]::IsNullOrEmpty($cmd.CommandText)) { $ownerLabel = $cmd.CommandText }
                                            else { $ownerLabel = "Invoke Command: $($cmd.Id)" }
                                            $cmd.HotkeyId = SetHotkey -KeyCombinationString $cmd.Hotkey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel
                                        } catch {
                                            $cmd.HotkeyId = $null
                                        }
                        }
                    } 
                    
                    RefreshSavedCommandsUI $form
                    UpdateChatCommanderSettings $data -forceWrite
                    Show-DarkMessageBox "Commands imported successfully." "Import" "Ok" "Information"
                }
            } catch {
                Show-DarkMessageBox "Failed to import commands: $_" "Import Error" "Ok" "Error"
            }
        }
    })

    $formData.BtnInstanceHotkeyToggle.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        
        $currentHotkeyText = $data.GlobalHotkey
        $oldHotkeyIdToUnregister = $data.GlobalHotkeyId
        $ownerKey = "global_toggle_$($data.InstanceId)"
        
        $newHotkey = Show-KeyCaptureDialog $currentHotkeyText -OwnerForm $form
        
        if ($newHotkey -and $newHotkey -ne $currentHotkeyText)
        {
            $data.GlobalHotkey = $newHotkey
            try
            {
                $script = @"
if (`$global:DashboardConfig.Resources.FtoolForms.Contains('$($data.InstanceId)')) {
    `$f = `$global:DashboardConfig.Resources.FtoolForms['$($data.InstanceId)']
    if (`$f -and -not `$f.IsDisposed -and `$f.Tag) {
        `$toggle = `$f.Tag.BtnHotkeyToggle
        if (`$toggle) {
            if (`$toggle.InvokeRequired) { `$toggle.Invoke([System.Action]{ `$toggle.Checked = -not `$toggle.Checked }) } else { `$toggle.Checked = -not `$toggle.Checked }
            ToggleChatCommanderHotkeys -InstanceId '$($data.InstanceId)' -ToggleState `$toggle.Checked
        }
    }
}
"@
                $scriptBlock = [scriptblock]::Create($script)
                # Friendly label for UI
                $ownerLabel = $null
                if ($data.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($data.Name)) { $ownerLabel = $data.Name }
                elseif ($data.PSObject.Properties['WindowTitle'] -and -not [string]::IsNullOrEmpty($data.WindowTitle)) { $ownerLabel = $data.WindowTitle }
                else { $ownerLabel = "Chat Commander: $($data.InstanceId)" }
                $data.GlobalHotkeyId = SetHotkey -KeyCombinationString $data.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel -OldHotkeyId $oldHotkeyIdToUnregister
            }
            catch { $data.GlobalHotkeyId = $null; $data.GlobalHotkey = $currentHotkeyText }
            UpdateChatCommanderSettings $data -forceWrite
        }
        elseif (-not $newHotkey -and $oldHotkeyIdToUnregister)
        {
            try { UnregisterHotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $ownerKey } catch {}
            try { UnregisterHotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $data.InstanceId } catch {}
            $data.GlobalHotkeyId = $null
            $data.GlobalHotkey = $null
            UpdateChatCommanderSettings $data -forceWrite
        }
    })

    $formData.BtnAddInstance.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        $row = $data.Row

        if (-not $row -or -not $row.Tag) { return }

        $baseInstanceId = "ChatCommander_" + $row.Cells[2].Value.ToString()

        $nextIndex = 1
        $foundGap = $false
        
        while (-not $foundGap)
        {
            $testId = "${baseInstanceId}_${nextIndex}"
            if (-not $global:DashboardConfig.Resources.FtoolForms.Contains($testId) -or 
                $global:DashboardConfig.Resources.FtoolForms[$testId].IsDisposed)
            {
                $foundGap = $true
            }
            else
            {
                $nextIndex++
            }
        }
        
        $newInstanceId = "${baseInstanceId}_${nextIndex}"

        $anchorForm = $form
        
        for ($i = $nextIndex - 1; $i -ge 0; $i--)
        {
            $prevId = if ($i -eq 0) { $baseInstanceId } else { "${baseInstanceId}_${i}" }
            
            if ($global:DashboardConfig.Resources.FtoolForms.Contains($prevId))
            {
                $candidate = $global:DashboardConfig.Resources.FtoolForms[$prevId]
                if ($candidate -and -not $candidate.IsDisposed -and $candidate.Visible)
                {
                    $anchorForm = $candidate
                    break
                }
            }
        }

        $newForm = CreateChatCommanderForm -instanceId $newInstanceId -targetWindowRect $null -windowTitle $data.WindowTitle -row $row
        
        if ($newForm -and $newForm -is [System.Windows.Forms.Form])
        {
            $newForm.Left = $anchorForm.Right + 10
            $newForm.Top = $anchorForm.Top
            $global:DashboardConfig.Resources.FtoolForms[$newInstanceId] = $newForm
            $newForm.Show()
            $newForm.BringToFront()
        }
        else
        {
            Write-Verbose "Error: CreateChatCommanderForm returned unexpected type: $($newForm.GetType().FullName)"
        }
    })

    $formData.BtnHotkeyToggle.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        $toggleOn = $this.Checked
        ToggleChatCommanderHotkeys -InstanceId $data.InstanceId -ToggleState $toggleOn
        UpdateChatCommanderSettings $data -forceWrite
    })

    # Removed: Interval.Add_TextChanged

    $formData.Name.Add_TextChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        # This is for the new command input name, not critical to save on every change
        # UpdateChatCommanderSettings $form.Tag -forceWrite
    })

    # Removed: BtnLoopToggle.Add_Click, BtnWaitToggle.Add_Click, BtnHoldKeyToggle.Add_Click, TxtHoldKeyInterval.Add_TextChanged

    $formData.PositionSliderX.Add_ValueChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        UpdateChatCommanderSettings $form.Tag -forceWrite
    })

    $formData.PositionSliderY.Add_ValueChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        UpdateChatCommanderSettings $form.Tag -forceWrite
    })

    $formData.txtSearchCommands.Add_TextChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        # Reset to page 1 on search
        $form.Tag.CurrentPage = 1
        RefreshSavedCommandsUI $form
    })

    $formData.BtnShowAll.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        
        $data.ShowAllMode = -not $data.ShowAllMode
        $data.BtnShowAll.Text = if ($data.ShowAllMode) { "LIST" } else { "ALL" }
        $data.BtnShowAll.BackColor = if ($data.ShowAllMode) { [System.Drawing.Color]::FromArgb(40, 120, 40) } else { [System.Drawing.Color]::FromArgb(60, 60, 100) }
        
        RefreshSavedCommandsUI $form
    })

    $formData.BtnPrevPage.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        if ($data.CurrentPage -gt 1) {
            $data.CurrentPage--
            RefreshSavedCommandsUI $form
        }
    })

    $formData.BtnNextPage.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        # Max page check is done in RefreshSavedCommandsUI, but we can check here too if we want
        $data.CurrentPage++
        RefreshSavedCommandsUI $form
    })

    $formData.GridAllCommands.Add_CellContentClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        
        $form = $s.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        
        $row = $s.Rows[$e.RowIndex]
        $cmd = $row.Tag
        
        if ($e.ColumnIndex -eq 0) { # Visibility Checkbox
            $newVal = -not ($row.Cells[0].Value -eq $true)
            $row.Cells[0].Value = $newVal
            
            if ($newVal) {
                # Add to visible
                $exists = $false
                foreach ($c in $data.SavedCommands) { if ($c.Id -eq $cmd.Id) { $exists = $true; break } }
				if (-not $exists) { 
				                    $data.SavedCommands.Add($cmd) | Out-Null 
				
				                    if ($cmd.Hotkey) {
				                        try {
				                            $scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($data.InstanceId)' -CommandId '$($cmd.Id)'")
				                            $ownerKey = "quickcmd_$($data.InstanceId)_$($cmd.Id)"
				                            $ownerLabel = if ($cmd.Name) { $cmd.Name } else { $cmd.CommandText }
				                            $cmd.HotkeyId = SetHotkey -KeyCombinationString $cmd.Hotkey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel
				                        } catch {
				                            $cmd.HotkeyId = $null
				                        }
				                    }
				                }
				} else {
                # Remove from visible
                $toRemove = $null
                foreach ($c in $data.SavedCommands) { if ($c.Id -eq $cmd.Id) { $toRemove = $c; break } }
                if ($toRemove) { 
                    if ($toRemove.HotkeyId) { try { UnregisterHotkeyInstance -Id $toRemove.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($toRemove.Id)" } catch {} }
                    $data.SavedCommands.Remove($toRemove) 
                }
            }
            UpdateChatCommanderSettings $data -forceWrite
            if (Get-Command RefreshHotkeysList -ErrorAction SilentlyContinue) { try { RefreshHotkeysList } catch {} }
        }
        elseif ($e.ColumnIndex -eq 3) { # Run
            if ($data.RunningTimer) { StopChatCommanderSequence $data }
            $job = Invoke-ChatCommanderAction -formData $data -commandString $cmd.CommandText
            if ($job) {
                $data.ActiveJobs.Add($job) | Out-Null
                $job.AsyncResult.AsyncWaitHandle.WaitOne() | Out-Null
                try { $job.PowerShell.EndInvoke($job.AsyncResult); $job.PowerShell.Dispose() } catch {}
                $data.ActiveJobs.Remove($job)
            }
        }
        elseif ($e.ColumnIndex -eq 4) { # Edit
            # Switch back to list mode to edit? Or just load into input fields?
            # Loading into input fields is standard behavior here
            $data.TxtCommandText.Text = $cmd.CommandText
            $data.Name.Text = $cmd.Name
            $data.EditingCommandId = $cmd.Id
            $data.BtnAdd.Text = ([char]0x2705) + ' Save'
            $data.BtnAdd.BackColor = [System.Drawing.Color]::FromArgb(40, 120, 40)
            
            # Optionally switch back to list mode if user wants to see the item they are editing in context?
            # For now, stay in table mode but allow editing via top panel
        }
        elseif ($e.ColumnIndex -eq 5) { # Delete
            if ($cmd.HotkeyId) { try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($cmd.Id)" } catch {} }
            $data.SavedCommands.Remove($cmd)
            $data.GlobalCommandList.Remove($cmd)
            RefreshSavedCommandsUI $form
            UpdateChatCommanderSettings $data -forceWrite
        }
    })

    $formData.Form.Add_FormClosed({
        param($src, $e)
        $instanceId = $src.Tag.InstanceId
        if ($instanceId) { CleanupChatCommanderResources $instanceId }
    })
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion