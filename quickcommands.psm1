<# quickcommands.psm1 #>

#region Helper Functions

function InitChatCommanderRunspace
{
    param($InstanceId)
    
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
    
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
        param($Handle, $CommandText) 

        $WM_KEYDOWN = 0x0100; $WM_KEYUP = 0x0101
        $VK_RETURN = 0x0D 
        $VK_SHIFT = 0x10 
        $VK_BACK = 0x08 
        $WM_CHAR = 0x0102

        
        
        [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, $VK_RETURN, 1)
        Start-Sleep -Milliseconds 50 
        [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYUP, $VK_RETURN, 0xC0000000)
        Start-Sleep -Milliseconds 50 

        
        for ($i = 0; $i -lt 10; $i++) {
            [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, $VK_BACK, 1)
            Start-Sleep -Milliseconds 5
            [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYUP, $VK_BACK, 0xC0000000)
            Start-Sleep -Milliseconds 5
        }

        
        foreach ($char in $CommandText.ToCharArray())
        {
            [Custom.Ftool]::fnPostMessage($Handle, $WM_CHAR, [int]$char, 1)
            Start-Sleep -Milliseconds 10
        }
        
        

        
        [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, $VK_RETURN, 1)
        Start-Sleep -Milliseconds 50 
        [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYUP, $VK_RETURN, 0xC0000000)
    }
}

function Invoke-ChatCommanderAction
{
    param($formData, $commandString) 
    
    
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
            $formData.RunspacePool = $null 
            return $null
        }
    }

    
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
        ClientRect   = (New-Object Custom.Native+RECT)
    }
        
    $positionTimer.Add_Tick({
        param($s, $e)
        try
        {
            if (-not $s -or -not $s.Tag) { $s.Stop(); $s.Dispose(); return }
            $timerData = $s.Tag
            if (-not $timerData -or -not $timerData.ContainsKey('Form') -or -not $timerData['Form'] -or $timerData['Form'].IsDisposed) { $s.Stop(); $s.Dispose(); return }

            $form     = $timerData.Form
            $formData = $timerData.FormData
            if (-not $formData -or -not $formData.PositionSliderX -or -not $formData.PositionSliderY) { return }

            $isMinimized = $false
            if ($timerData.ContainsKey('WindowHandle') -and $timerData.WindowHandle -ne [IntPtr]::Zero) {
                $isMinimized = [Custom.Native]::IsWindowMinimized($timerData.WindowHandle)
            }

            if ($isMinimized) {
                $virtual = [System.Windows.Forms.SystemInformation]::VirtualScreen
                $sliderValueX = $formData.PositionSliderX.Value
                $sliderValueY = $formData.PositionSliderY.Value
                $minLeft = $virtual.Left + 8
                $maxLeft = $virtual.Right - $form.Width - 8
                $timerData.RealLeft = [Math]::Max($minLeft, [Math]::Min($maxLeft, $sliderValueX))
                $minTop = $virtual.Top + 8
                $maxTop = $virtual.Bottom - $form.Height - 8
                $timerData.RealTop = [Math]::Max($minTop, [Math]::Min($maxTop, $sliderValueY))
                if ($form.Visible) { $form.Hide() }
                return
            } else {
                if (-not $form.Visible) { $form.Show() }
            }

            $virtual = [System.Windows.Forms.SystemInformation]::VirtualScreen
            try {
                if ($timerData.ContainsKey('WindowHandle') -and $timerData.WindowHandle -and $timerData.WindowHandle -ne [IntPtr]::Zero) {
                    $currRect = New-Object Custom.Native+RECT
                    if ([Custom.Native]::GetWindowRect($timerData.WindowHandle, [ref]$currRect)) {
                        if ($timerData.ClientRect -and ($timerData.ClientRect.Left -ne 0 -or $timerData.ClientRect.Top -ne 0 -or $timerData.ClientRect.Right -ne 0 -or $timerData.ClientRect.Bottom -ne 0)) {
                            $dx = $currRect.Left - $timerData.ClientRect.Left
                            $dy = $currRect.Top - $timerData.ClientRect.Top
                            if (($dx -ne 0 -or $dy -ne 0) -and -not ($formData.IsDragging)) {
                                try {
                                    $minX = $formData.PositionSliderX.Minimum
                                    $maxX = $formData.PositionSliderX.Maximum
                                    $minX = [Math]::Max($virtual.Left, [Math]::Min($virtual.Right, $minX))
                                    $maxX = [Math]::Max($virtual.Left, [Math]::Min($virtual.Right, $maxX))
                                    if ($minX -gt $maxX) { $tmp = $minX; $minX = $maxX; $maxX = $tmp }
                                    $newX = [int]([Math]::Max($minX, [Math]::Min($maxX, $formData.PositionSliderX.Value + $dx)))
                                    $formData.SuppressPersist = $true
                                    try {
                                        $formData.PositionSliderX.Value = $newX
                                        $minY = $formData.PositionSliderY.Minimum
                                        $maxY = $formData.PositionSliderY.Maximum
                                        $minY = [Math]::Max($virtual.Top, [Math]::Min($virtual.Bottom, $minY))
                                        $maxY = [Math]::Max($virtual.Top, [Math]::Min($virtual.Bottom, $maxY))
                                        if ($minY -gt $maxY) { $tmp = $minY; $minY = $maxY; $maxY = $tmp }
                                        $newY = [int]([Math]::Max($minY, [Math]::Min($maxY, $formData.PositionSliderY.Value + $dy)))
                                        $formData.PositionSliderY.Value = $newY
                                    }
                                    finally { $formData.SuppressPersist = $false }



                                    $timerData.RealLeft = $timerData.RealLeft + $dx
                                    $timerData.RealTop  = $timerData.RealTop  + $dy
                                } catch {}
                            }
                        }
                        $timerData.ClientRect = $currRect
                    }
                }
            } catch {}

            $virtual = [System.Windows.Forms.SystemInformation]::VirtualScreen
            $sliderValueX = $formData.PositionSliderX.Value
            $sliderValueY = $formData.PositionSliderY.Value

            $minLeft = $virtual.Left + 8
            $maxLeft = $virtual.Right - $form.Width - 8
            $targetLeft = [Math]::Max($minLeft, [Math]::Min($maxLeft, $sliderValueX))

            $minTop = $virtual.Top + 8
            $maxTop = $virtual.Bottom - $form.Height - 8
            $targetTop = [Math]::Max($minTop, [Math]::Min($maxTop, $sliderValueY))

            $newRealLeft = $timerData.RealLeft + ($targetLeft - $timerData.RealLeft) * 0.2
            $newRealTop  = $timerData.RealTop  + ($targetTop  - $timerData.RealTop)  * 0.2
            $timerData.RealLeft = $newRealLeft
            $timerData.RealTop  = $newRealTop

            $form.Left = [int]$newRealLeft
            $form.Top  = [int]$newRealTop

            $formHandle = $form.Handle
            $flags = [Custom.Native]::SWP_NOMOVE -bor [Custom.Native]::SWP_NOSIZE -bor [Custom.Native]::SWP_NOACTIVATE

            if ($timerData.ContainsKey('WindowHandle') -and $timerData.WindowHandle -and $timerData.WindowHandle -ne [IntPtr]::Zero) {
                $linkedHandle = $timerData.WindowHandle
                $foregroundWindow = [Custom.Native]::GetForegroundWindow()
                $shouldBeTopMost = $false
                if ($foregroundWindow -eq $linkedHandle) { $shouldBeTopMost = $true }
                else {
                    $activeCtrl = [System.Windows.Forms.Control]::FromHandle($foregroundWindow)
                    if ($activeCtrl -and -not $activeCtrl.IsDisposed -and $activeCtrl.Tag) {
                        if ($activeCtrl.Tag.PSObject.Properties['SelectedWindow'] -and $activeCtrl.Tag.SelectedWindow -eq $linkedHandle) { $shouldBeTopMost = $true }
                    }
                }

                if ($shouldBeTopMost) {
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
                    if ($timerData.ZState -ne 'topmost' -or $forceUpdate) {
                        if ($forceUpdate) { [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::HWND_NOTOPMOST, 0, 0, 0, 0, $flags) | Out-Null }
                        [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::HWND_TOPMOST, 0, 0, 0, 0, $flags) | Out-Null
                        $timerData.ZState = 'topmost'
                    }
                } else {
                    if ($timerData.ZState -ne 'standard') { [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::HWND_NOTOPMOST, 0, 0, 0, 0, $flags) | Out-Null; $timerData.ZState = 'standard' }
                    $next = [Custom.Native]::GetWindow($formHandle, 2)
                    $amIAboveGame = $false; $loopCount = 0
                    while ($next -ne [IntPtr]::Zero -and $loopCount -lt 50) { if ($next -eq $linkedHandle) { $amIAboveGame = $true; break }; $next = [Custom.Native]::GetWindow($next, 2); $loopCount++ }
                    if (-not $amIAboveGame) { $bottomOfCluster = [Custom.Native]::GetWindow($linkedHandle, 3); if ($bottomOfCluster -ne [IntPtr]::Zero) { [Custom.Native]::PositionWindow($formHandle, $bottomOfCluster, 0, 0, 0, 0, $flags) | Out-Null } else { [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::TopWindowHandle, 0, 0, 0, 0, $flags) | Out-Null } }
                }
            } else {
                if ($timerData.ZState -ne 'standard') { [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::HWND_NOTOPMOST, 0, 0, 0, 0, $flags) | Out-Null; $timerData.ZState = 'standard' }
            }
        }
        catch {
            Write-Verbose ('CHATCOMMANDER: Position timer critical error: {0}' -f $_.Exception.Message)
        }
    })
    
    $positionTimer.Start()
    $global:DashboardConfig.Resources.Timers["ChatCommanderPosition_$($formData.InstanceId)"] = $positionTimer
}

function StartChatCommanderSequence
{
    param($formData, [switch]$RunSavedSequence, [switch]$AutoStart)

    $steps = [System.Collections.ArrayList]::new()

    if ($RunSavedSequence)
    {
        
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

    if (-not $formData.PSObject.Properties['CloseAfterRunPending'])
    {
        $formData | Add-Member -MemberType NoteProperty -Name 'CloseAfterRunPending' -Value $false
    }
    $closeAfterRun = $false
    if ($AutoStart -and $formData.PSObject.Properties['ChkCloseAfterRun'] -and $formData.ChkCloseAfterRun)
    {
        $closeAfterRun = [bool]$formData.ChkCloseAfterRun.Checked
    }
    $formData.CloseAfterRunPending = $closeAfterRun

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
            
            
            $shouldStop = ($nextIdx -ge $fData.ActiveSteps.Count)

            
            $pollTimer = New-Object System.Windows.Forms.Timer
            $pollTimer.Interval = 2
            $pollTimer.Tag = @{ 
                Job = $job; 
                MainTimer = $s; 
                NextInterval = 50; 
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

    if ($formData.PSObject.Properties['CloseAfterRunPending'])
    {
        if ($NaturalEnd -and $formData.CloseAfterRunPending)
        {
            $formData.CloseAfterRunPending = $false
            $formToClose = $null
            if ($formData.PSObject.Properties['Form']) { $formToClose = $formData.Form }
            if ($formToClose -and -not $formToClose.IsDisposed -and $formToClose.IsHandleCreated)
            {
                try
                {
                    $closeAction = { StopChatCommanderForm $formToClose }.GetNewClosure()
                    $formToClose.BeginInvoke([System.Action]$closeAction) | Out-Null
                    Write-Verbose "ChatCommander: close-after-run queued for $($formData.InstanceId)"
                }
                catch { Write-Verbose "ChatCommander: close-after-run failed to queue: $_" }
            }
        }
        elseif (-not $NaturalEnd)
        {
            $formData.CloseAfterRunPending = $false
        }
    }
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
        if ($global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId))
        {
            $form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
            if ($form -and -not $form.IsDisposed -and $form.Tag -and $form.Tag.SavedCommands)
            {
                foreach ($cmd in $form.Tag.SavedCommands)
                {
                    if ($cmd.HotkeyId)
                    {
                        $ownerKey = "quickcmd_$($InstanceId)_$($cmd.Id)"
                        if ($ToggleState)
                        {
                            try { ResumeHotkeysForOwner -OwnerKey $ownerKey } catch {}
                        }
                        else
                        {
                            try { PauseHotkeysForOwner -OwnerKey $ownerKey } catch {}
                        }
                    }
                }
            }
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
            
            
            try {
                if (Test-DashboardCommand 'UnregisterHotkeysByOwnerPattern') {
                    UnregisterHotkeysByOwnerPattern -OwnerPattern "quickcmd_${instanceId}_*"
                }
            } catch {}
        }
    }
    
    if (Test-DashboardCommand 'RefreshHotkeysList') { try { RefreshHotkeysList } catch {} }
}

function LoadChatCommanderSettings
{
    param($formData)
    $formData.IsLoading = $true
    
    if (-not $global:DashboardConfig.Config.Contains('ChatCommander')) { $formData.IsLoading = $false; return }
    
    $profilePrefix = $null
    if (Test-DashboardCommand 'FindOrCreateProfile')
    {
        $profilePrefix = FindOrCreateProfile $formData.WindowTitle -ProfileHint (Get-RowProfileHint -formData $formData)
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

        $runOnOpenName = "run_on_open_$p"
        if ($cfg.Contains($runOnOpenName))
        {
            $runOnOpenValue = $cfg[$runOnOpenName]
            try
            {
                $runOnOpenText = "$runOnOpenValue".Trim()
                if ($runOnOpenText -eq '1' -or $runOnOpenText -eq '0') { $formData.ChkRunOnStart.Checked = ($runOnOpenText -eq '1') }
                else { $formData.ChkRunOnStart.Checked = [bool]::Parse($runOnOpenText) }
            }
            catch { $formData.ChkRunOnStart.Checked = $true }
        }
        else { $formData.ChkRunOnStart.Checked = $true }

        $closeAfterRunName = "close_after_run_$p"
        if ($cfg.Contains($closeAfterRunName))
        {
            $closeAfterRunValue = $cfg[$closeAfterRunName]
            try
            {
                $closeAfterRunText = "$closeAfterRunValue".Trim()
                if ($closeAfterRunText -eq '1' -or $closeAfterRunText -eq '0') { $formData.ChkCloseAfterRun.Checked = ($closeAfterRunText -eq '1') }
                else { $formData.ChkCloseAfterRun.Checked = [bool]::Parse($closeAfterRunText) }
            }
            catch { $formData.ChkCloseAfterRun.Checked = $false }
        }
        else { $formData.ChkCloseAfterRun.Checked = $false }

        if ($cfg.Contains("PosX_$p")) {
            $positionValue = [int]$cfg["PosX_$p"]
            $min = $formData.PositionSliderX.Minimum
            $max = $formData.PositionSliderX.Maximum
            if ($min -gt $max) { $tmp = $min; $min = $max; $max = $tmp }
            $formData.PositionSliderX.Value = [Math]::Max($min, [Math]::Min($max, $positionValue))
        } else {
            $min = $formData.PositionSliderX.Minimum
            $max = $formData.PositionSliderX.Maximum
            $formData.PositionSliderX.Value = [Math]::Max($min, [Math]::Min($max, $formData.Form.Left))
        }
        if ($cfg.Contains("PosY_$p")) {
            $positionValue = [int]$cfg["PosY_$p"]
            $min = $formData.PositionSliderY.Minimum
            $max = $formData.PositionSliderY.Maximum
            if ($min -gt $max) { $tmp = $min; $min = $max; $max = $tmp }
            $formData.PositionSliderY.Value = [Math]::Max($min, [Math]::Min($max, $positionValue))
        } else {
            $min = $formData.PositionSliderY.Minimum
            $max = $formData.PositionSliderY.Maximum
            $formData.PositionSliderY.Value = [Math]::Max($min, [Math]::Min($max, $formData.Form.Top))
        }

        
        $globalCommands = [System.Collections.ArrayList]::new()
        if ($cfg.Contains("GlobalCommands")) {
            try {
                $raw = $cfg["GlobalCommands"]
                if ($raw -is [System.Collections.IEnumerable] -and $raw -isnot [string]) { $raw = $raw -join '' }
                
                
                if ($raw -is [string]) {
                    
                    $cleanRaw = $raw -replace '\\"', '"'
                    $imported = ConvertFrom-Json $cleanRaw -ErrorAction SilentlyContinue
                    
                    if (-not $imported) {
                        
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

        
        $visibleIds = [System.Collections.Generic.HashSet[string]]::new()
        if ($cfg.Contains($visibleKey)) {
            try {
                $vRaw = $cfg[$visibleKey]
                if ($vRaw -is [System.Collections.IEnumerable] -and $vRaw -isnot [string]) { $vRaw = $vRaw -join '' }

                $vList = $null
                
                if ($vRaw -is [string]) {
                    $cleanVRaw = $vRaw -replace '\\"', '"'
                    $vList = ConvertFrom-Json $cleanVRaw -ErrorAction SilentlyContinue
                    
                    if (-not $vList) {
                        $vList = ConvertFrom-Json $vRaw -ErrorAction SilentlyContinue
                    }
                }

                if ($vList) {
                    
                    if ($vList -is [string]) {
                        $visibleIds.Add($vList) | Out-Null
                    }
                    elseif ($vList -is [System.Collections.IEnumerable]) {
                        foreach ($id in $vList) { $visibleIds.Add($id) | Out-Null }
                    }
                }
            } catch {}
        } elseif ($cfg.Contains("SavedCommands_$p")) {
             
             foreach ($cmd in $globalCommands) { if ($cmd.Id) { $visibleIds.Add($cmd.Id) | Out-Null } }
        }

        
        $formData.SavedCommands = [System.Collections.ArrayList]::new()
        foreach ($cmd in $formData.GlobalCommandList) {
            
            if ($cmd.Id -and ($visibleIds.Contains($cmd.Id) -or $visibleIds.Contains($cmd.Id.ToString()))) {
                $formData.SavedCommands.Add($cmd) | Out-Null
                
                
                if ($null -eq $cmd.PSObject.Properties['Hotkey']) { $cmd | Add-Member -MemberType NoteProperty -Name 'Hotkey' -Value $null -Force }
                if ($null -eq $cmd.PSObject.Properties['HotkeyId']) { $cmd | Add-Member -MemberType NoteProperty -Name 'HotkeyId' -Value $null -Force }
                if ($cmd.Hotkey) {
                    try {
                            $scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($formData.InstanceId)' -CommandId '$($cmd.Id)'")
                            
                            $ownerKey = "quickcmd_$($formData.InstanceId)_$($cmd.Id)"
                            
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
        $formData.ChkRunOnStart.Checked = $true
        $formData.ChkCloseAfterRun.Checked = $false
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
    if (Test-DashboardCommand 'FindOrCreateProfile')
    {
        $profilePrefix = FindOrCreateProfile $formData.WindowTitle -ProfileHint (Get-RowProfileHint -formData $formData)
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
        
        $cfg["Name_$p"] = $formData.Name.Text
        $cfg["PosX_$p"] = $formData.PositionSliderX.Value
        $cfg["PosY_$p"] = $formData.PositionSliderY.Value
        
        $cfg["GlobalHotkey_$p"] = $formData.GlobalHotkey
        $cfg["hotkeys_enabled_$p"] = $formData.BtnHotkeyToggle.Checked
        if ($formData.PSObject.Properties['ChkRunOnStart'] -and $formData.ChkRunOnStart)
        {
            $cfg["run_on_open_$p"] = $formData.ChkRunOnStart.Checked
        }
        if ($formData.PSObject.Properties['ChkCloseAfterRun'] -and $formData.ChkCloseAfterRun)
        {
            $cfg["close_after_run_$p"] = $formData.ChkCloseAfterRun.Checked
        }

        
        $cfg["GlobalCommands"] = (ConvertTo-Json -InputObject $formData.GlobalCommandList -Compress -Depth 5)
        
        
        $visibleIds = [System.Collections.ArrayList]::new()
        foreach ($cmd in $formData.SavedCommands) { if ($cmd.Id) { $visibleIds.Add($cmd.Id) | Out-Null } }
        $cfg[$visibleKey] = (ConvertTo-Json -InputObject $visibleIds -Compress -Depth 2)
        
        
        if ($cfg.Contains("SavedCommands_$p")) { $cfg.Remove("SavedCommands_$p") }
        
        
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
            if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
        }
        else
        {
            if (Test-DashboardCommand 'Request-DashboardConfigWrite') { Request-DashboardConfigWrite }
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
    
    $windowTitle = if ($row.Tag.PSObject.Properties['MainWindowTitle']) { $row.Tag.MainWindowTitle }
    elseif ($row.Tag.PSObject.Properties['Title']) { $row.Tag.Title }
    else { $row.Cells[1].Value.ToString() }
    
    $ChatCommanderForm = CreateChatCommanderForm $instanceId $targetWindowRect $windowTitle $row
    
    $global:DashboardConfig.Resources.FtoolForms[$instanceId] = $ChatCommanderForm
    
    Register-OpenForm $ChatCommanderForm
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
    $ChatCommanderForm.Height = 400 
    $ChatCommanderForm.Top = if ($targetWindowRect) { ($targetWindowRect.Top + 30) } else { 200 }
    $ChatCommanderForm.Left = if ($targetWindowRect) { ($targetWindowRect.Right - $ChatCommanderForm.Width - 10) } else { 200 }
    $ChatCommanderForm.BackColor = Get-Theme 'MainBg'
    $ChatCommanderForm.ForeColor = Get-Theme 'PrimaryText'
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
    $ChatCommanderToolTip | Add-Member -MemberType NoteProperty -Name 'TipFont' -Value (Get-UIFont 'Segoe UI' 9)
    $ChatCommanderToolTip.Add_Draw({
            $g, $b = $_.Graphics, $_.Bounds
            $tc = Get-Theme
            $bgBrush = New-Object System.Drawing.SolidBrush $tc['ToolTipBg']
            $borderPen = New-Object System.Drawing.Pen $tc['TooltipBorderColor']
            $fgBrush = New-Object System.Drawing.SolidBrush $tc['ToolTipFg']
            try
            {
            	$g.FillRectangle($bgBrush, $b)
            	$g.DrawRectangle($borderPen, $b.X, $b.Y, $b.Width-1, $b.Height-1)
            	$g.DrawString($_.ToolTipText, $this.TipFont, $fgBrush, 3, 3, [System.Drawing.StringFormat]::GenericTypographic)
            }
            finally
            {
            	$bgBrush.Dispose(); $borderPen.Dispose(); $fgBrush.Dispose()
            }
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

    $headerPanel = SetUIElement -type 'Panel' -visible $true -width 248 -height 24 -top 1 -left 1 -tag 'TopBar'
    $ChatCommanderForm.Controls.Add($headerPanel)

    $headerPanel.Add_Paint({
            param($s, $e)
            try {
                $tc = Get-Theme
                $accent = $null
                try { $accent = $tc['TopBarAccent'] } catch {}
                if ($accent -is [System.Drawing.Color]) {
                    $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(170, $accent.R, $accent.G, $accent.B))
                    $e.Graphics.FillRectangle($brush, 0, ($s.ClientSize.Height - 1), $s.ClientSize.Width, 1)
                    $brush.Dispose()
                }
            }
            catch {}
        })

    $labelText = if ($row.Cells[1].Value) { $row.Cells[1].Value.ToString() } else { "" }
    $labelWinTitle = SetUIElement -type 'Label' -visible $true -width 170 -height 30 -top 2 -left 2 -text ($labelText) -font (Get-UIFont 'Segoe UI' 6 'Regular')
    $headerPanel.Controls.Add($labelWinTitle)

    $btnInstanceHotkeyToggle = SetUIElement -type 'Label' -visible $true -width 16 -height 16 -top 30 -left 2 -text ([char]0x2328) -font (Get-UIFont 'Segoe UI' 10) -tooltip "Set Master Hotkey`nAssign a global hotkey to toggle all hotkeys for this instance."
    $ChatCommanderForm.Controls.Add($btnInstanceHotkeyToggle)

    $btnHotkeyToggle = SetUIElement -type 'Toggle' -visible $true -width 32 -height 16 -top 4 -left 177 -text '' -fs 'Flat' -font (Get-UIFont 'Segoe UI' 8) -checked $true -tooltip "Toggle Hotkeys`nEnable or disable all hotkeys for this specific ChatCommander instance."
    $headerPanel.Controls.Add($btnHotkeyToggle)

    $btnImport = SetUIElement -type 'Button' -visible $true -width 38 -height 18 -top 30 -left 23 -text 'Import' -fs 'Flat' -font (Get-UIFont 'Segoe UI' 7) -tag 'HeaderButton' -tooltip "Import Settings`nLoad settings from a file."
    $ChatCommanderForm.Controls.Add($btnImport)

    $btnExport = SetUIElement -type 'Button' -visible $true -width 38 -height 18 -top 30 -left 64 -text 'Export' -fs 'Flat' -font (Get-UIFont 'Segoe UI' 7) -tag 'HeaderButton' -tooltip "Export Settings`nSave current settings to a file."
    $ChatCommanderForm.Controls.Add($btnExport)

    $btnAddInstance = SetUIElement -type 'Button' -visible $true -width 76 -height 18 -top 30 -left 105 -text 'New Instance' -fs 'Flat' -font (Get-UIFont 'Segoe UI' 7) -tag 'HeaderButton' -tooltip "Add Instance`nCreates another ChatCommander window for this specific client."
    $ChatCommanderForm.Controls.Add($btnAddInstance)
        
    $btnShowHide = SetUIElement -type 'Button' -visible $true -width 16 -height 16 -top 4 -left 211 -text ([char]0x25B2) -fs 'Flat' -font (Get-UIFont 'Segoe UI' 11) -tag 'HeaderButton' -tooltip "Minimize/Expand`nCollapse or expand the ChatCommander window to save screen space."
    $headerPanel.Controls.Add($btnShowHide)
        
    $btnReset = SetUIElement -type 'Button' -visible $true -width 38 -height 18 -top 30 -left 184 -text 'Reset' -fs 'Flat' -font (Get-UIFont 'Segoe UI' 7) -tag 'HeaderButton' -tooltip "Reset`nReset all settings for this instance to default."
    $ChatCommanderForm.Controls.Add($btnReset)

    $btnClose = SetUIElement -type 'Button' -visible $true -width 16 -height 16 -top 4 -left 229 -text ([char]0x166D) -fs 'Flat' -font (Get-UIFont 'Segoe UI' 11) -tag 'Danger' -tooltip "Close`nStops the ChatCommander and closes this window."
    $headerPanel.Controls.Add($btnClose)

    
    $chkRunOnStart = SetUIElement -type 'CheckBox' -visible $true -width 240 -height 16 -top 52 -left 5 -text 'Run commands on open' -font (Get-UIFont 'Segoe UI' 7) -checked $true -tooltip "Run on Open`nTicked: the saved commands are fired automatically when this window is opened for a client after login (QuickCmd auto-start).`nUnticked: the window is opened but no command is sent.`nManual runs (Run All Visible, the ▶ button and hotkeys) always work."
    $ChatCommanderForm.Controls.Add($chkRunOnStart)

    $chkCloseAfterRun = SetUIElement -type 'CheckBox' -visible $true -width 240 -height 16 -top 70 -left 5 -text 'Close after run' -font (Get-UIFont 'Segoe UI' 7) -checked $false -tooltip "Close after Run`nTicked: this window closes itself automatically once the run-on-open sequence has finished.`nOnly the automatic startup run is affected - manual runs (Run All Visible, the ▶ buttons and hotkeys) keep the window open.`nRequires 'Run commands on open' to be ticked to have any effect."
    $ChatCommanderForm.Controls.Add($chkCloseAfterRun)

    $panelNewCommand = SetUIElement -type 'Panel' -visible $true -width 240 -height 115 -top 88 -left 5
    $ChatCommanderForm.Controls.Add($panelNewCommand)
    
    $lblCmd = SetUIElement -type 'Label' -visible $true -width 100 -height 15 -top 0 -left 3 -text "Command:" -font (Get-UIFont 'Segoe UI' 7)
    $panelNewCommand.Controls.Add($lblCmd)

    $txtCommandText = SetUIElement -type 'TextBox' -visible $true -width 234 -height 24 -top 15 -left 3 -text '' -font (Get-UIFont 'Segoe UI' 8) -tooltip "Chat Command`nEnter the chat command to send."
    $panelNewCommand.Controls.Add($txtCommandText)
    
    $lblName = SetUIElement -type 'Label' -visible $true -width 100 -height 15 -top 40 -left 3 -text "Name:" -font (Get-UIFont 'Segoe UI' 7)
    $panelNewCommand.Controls.Add($lblName)

    $name = SetUIElement -type 'TextBox' -visible $true -width 170 -height 17 -top 55 -left 3 -text 'New' -font (Get-UIFont 'Segoe UI' 8 'Regular') -tooltip "Name`nGive this command a descriptive name for easier identification."
    $panelNewCommand.Controls.Add($name)

    $lblVar = SetUIElement -type 'Label' -visible $true -width 100 -height 15 -top 75 -left 3 -text "Variable:" -font (Get-UIFont 'Segoe UI' 7)
    $panelNewCommand.Controls.Add($lblVar)

    $txtVariable = SetUIElement -type 'TextBox' -visible $true -width 170 -height 17 -top 90 -left 3 -text '' -font (Get-UIFont 'Segoe UI' 8 'Regular') -tooltip "Variable`nReplaces [var] in the command text with this value."
    $panelNewCommand.Controls.Add($txtVariable)

    
    
    $btnAdd = SetUIElement -type 'Button' -visible $true -width 55 -height 20 -top 88 -left 178 -text ([char]0x2795) + ' Add' -fs 'Flat' -font (Get-UIFont 'Segoe UI' 8) -tooltip "Add Command`nAdds the current chat command to the saved list."
    $panelNewCommand.Controls.Add($btnAdd)

    

    $btnRunSequence = SetUIElement -type 'Button' -visible $true -width 240 -height 25 -top ($panelNewCommand.Bottom + 5) -left 5 -text "Run All Visible" -fs 'Flat' -font (Get-UIFont 'Segoe UI' 8) -tag 'Primary' -tooltip "Run Sequence`nExecutes all currently visible commands in the list from top to bottom."
    $ChatCommanderForm.Controls.Add($btnRunSequence)

    
    $lblSearch = SetUIElement -type 'Label' -visible $true -width 100 -height 15 -top ($btnRunSequence.Bottom + 5) -left 5 -text "Search:" -font (Get-UIFont 'Segoe UI' 7)
    $ChatCommanderForm.Controls.Add($lblSearch)

    $txtSearchCommands = SetUIElement -type 'TextBox' -visible $true -width 195 -height 24 -top ($lblSearch.Bottom + 0) -left 5 -text '' -font (Get-UIFont 'Segoe UI' 8) -tooltip "Search Commands`nType to filter saved chat commands across all profiles."
    $ChatCommanderForm.Controls.Add($txtSearchCommands)

    $btnShowAll = SetUIElement -type 'Button' -visible $true -width 40 -height 24 -top ($lblSearch.Bottom + 0) -left 205 -text "ALL" -fs 'Flat' -font (Get-UIFont 'Segoe UI' 7) -tooltip "Show All`nToggle table view of all commands."
    $ChatCommanderForm.Controls.Add($btnShowAll)

    
    $panelSavedCommands = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelSavedCommands.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $panelSavedCommands.WrapContents = $false
    $panelSavedCommands.AutoSize = $true
    $panelSavedCommands.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $panelSavedCommands.Width = 240
    $panelSavedCommands.Top = ($txtSearchCommands.Bottom + 5)
    $panelSavedCommands.Left = 5
    $panelSavedCommands.BackColor = Get-Theme 'SecondaryBg'
    $ChatCommanderForm.Controls.Add($panelSavedCommands)

    
    $panelPagination = SetUIElement -type 'Panel' -visible $true -width 240 -height 30 -top 0 -left 5
    $ChatCommanderForm.Controls.Add($panelPagination)

    $btnPrevPage = SetUIElement -type 'Button' -visible $true -width 30 -height 20 -top 5 -left 5 -text "<" -fs 'Flat' -font (Get-UIFont 'Segoe UI' 8)
    $panelPagination.Controls.Add($btnPrevPage)

    $lblPageInfo = SetUIElement -type 'Label' -visible $true -width 160 -height 20 -top 8 -left 40 -text "1 / 1" -font (Get-UIFont 'Segoe UI' 8)
    $lblPageInfo.TextAlign = 'MiddleCenter'
    $panelPagination.Controls.Add($lblPageInfo)

    $btnNextPage = SetUIElement -type 'Button' -visible $true -width 30 -height 20 -top 5 -left 205 -text ">" -fs 'Flat' -font (Get-UIFont 'Segoe UI' 8)
    $panelPagination.Controls.Add($btnNextPage)

    
    $gridAllCommands = SetUIElement -type 'DataGridView' -visible $false -width 240 -height 300 -top ($txtSearchCommands.Bottom + 5) -left 5
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
    $colDel.DefaultCellStyle.ForeColor = Get-Theme 'ErrorColor'

    $gridAllCommands.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colVis, $colName, $colCmd, $colExec, $colEdit, $colDel))
    $ChatCommanderForm.Controls.Add($gridAllCommands)

    
    $gridAllCommands.Add_CellContentClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -eq 0) { 
            $grid = $s
            $form = $grid.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag
            
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            $isChecked = $grid.Rows[$e.RowIndex].Cells[0].Value
            $cmd = $grid.Rows[$e.RowIndex].Tag
            
            if ($isChecked) {
                
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
                            Write-Verbose "QuickCommand: Registered hotkey $($cmd.Hotkey) for command $($cmd.Name)"
                        } catch { 
                            $cmd.HotkeyId = $null 
                            Write-Warning "QuickCommand: Failed to register hotkey $($cmd.Hotkey) for command $($cmd.Name): $_"
                        }
                    }
                }
            } else {
                
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
            if (Test-DashboardCommand 'RefreshHotkeysList') { try { RefreshHotkeysList } catch {} }
        }
    })

    $virtual = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $maxX = $virtual.Right
    $maxY = $virtual.Bottom
    $minX = $virtual.Left
    $minX = [Math]::Max($virtual.Left, [Math]::Min($virtual.Right, $minX))
    $maxX = [Math]::Max($virtual.Left, [Math]::Min($virtual.Right, $maxX))
    if ($minX -gt $maxX) { $tmp = $minX; $minX = $maxX; $maxX = $tmp }
    $minY = $virtual.Top
    $minY = [Math]::Max($virtual.Top, [Math]::Min($virtual.Bottom, $minY))
    $maxY = [Math]::Max($virtual.Top, [Math]::Min($virtual.Bottom, $maxY))
    if ($minY -gt $maxY) { $tmp = $minY; $minY = $maxY; $maxY = $tmp }

    $positionSliderY = New-Object System.Windows.Forms.TrackBar
    $positionSliderY.Orientation = 'Vertical'
    $positionSliderY.Minimum = $minY
    $positionSliderY.Maximum = $maxY
    $positionSliderY.TickFrequency = 300
    $initialY = [Math]::Max($minY, [Math]::Min($maxY, $ChatCommanderForm.Top))
    $positionSliderY.Value = $initialY
    $positionSliderY.Size = New-Object System.Drawing.Size(0, 0)
    $positionSliderY.Location = New-Object System.Drawing.Point(0, 0)
        
    $positionSliderX = New-Object System.Windows.Forms.TrackBar
    $positionSliderX.Minimum = $minX
    $positionSliderX.Maximum = $maxX
    $positionSliderX.TickFrequency = 300
    $initialX = [Math]::Max($minX, [Math]::Min($maxX, $ChatCommanderForm.Left))
    $positionSliderX.Value = $initialX
    $positionSliderX.Size = New-Object System.Drawing.Size(0, 0)
    $positionSliderX.Location = New-Object System.Drawing.Point(0, 0)

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
        LabelWinTitle   = $labelWinTitle
        PanelNewCommand = $panelNewCommand
        TxtVariable     = $txtVariable
        TxtCommandText  = $txtCommandText
        
        Name            = $name 
        
        BtnHotkeyToggle = $btnHotkeyToggle
        BtnInstanceHotkeyToggle = $btnInstanceHotkeyToggle
        ChkRunOnStart    = $chkRunOnStart
        ChkCloseAfterRun = $chkCloseAfterRun
        BtnAdd          = $btnAdd 
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
        OriginalHeight  = $ChatCommanderForm.Height 
        HotkeyId        = $null
        GlobalHotkeyId  = $null
        GlobalHotkey    = $null
        SavedCommands   = [System.Collections.ArrayList]::new() 
        GlobalCommandList = [System.Collections.ArrayList]::new() 
        txtSearchCommands = $txtSearchCommands 
        panelSavedCommands = $panelSavedCommands 
        ActiveSteps     = @()
        CurrentStepIndex= 0
        ActiveJobs      = [System.Collections.ArrayList]::new()
        RunspacePool    = (InitChatCommanderRunspace -InstanceId $instanceId)
        IsLoading       = $false
        EditingCommandId = $null 
        CurrentPage     = 1
        ItemsPerPage    = 30
        ShowAllMode     = $false
    }
    
    $ChatCommanderForm.Tag = $formData
    
    CreateChatCommanderPositionTimer $formData | Out-Null
    LoadChatCommanderSettings $formData | Out-Null
    AddChatCommanderDragHandlers $formData | Out-Null
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

    if ($formData.IsCollapsed) {
        $ChatCommanderForm.Height = 26
        $formData.BtnShowHide.Text = [char]0x25BC
    } else {
        $ChatCommanderForm.Height = $formData.OriginalHeight
        $formData.BtnShowHide.Text = [char]0x25B2
    }
    
    return $ChatCommanderForm
}

function CreateSavedCommandEntryUI
{
    param($formData, $cmd)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(230, 26)
    $panel.BackColor = Get-Theme 'TertiaryBg'
    $panel.Tag = $cmd 
    $panel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = if ($cmd.Name) { $cmd.Name } else { $cmd.CommandText }
    $lbl.ForeColor = Get-Theme 'WhiteText'
    $lbl.Location = New-Object System.Drawing.Point(3, 4)
    $lbl.Size = New-Object System.Drawing.Size(125, 18)
    $lbl.Font = Get-UIFont 'Segoe UI' 8
    $lbl.AutoEllipsis = $true
    $lbl.Cursor = 'Hand'
    $panel.Controls.Add($lbl)
    
    $toolTip = $formData.ToolTipFtool
    if ($toolTip) { $toolTip.SetToolTip($lbl, "$($cmd.Name)`n$($cmd.CommandText)`nClick to Edit") }

    
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text =  [char]0x25B6
    $btnRun.FlatStyle = 'Flat'
    $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = Get-Theme 'SuccessColorAlt'
    $btnRun.ForeColor = Get-Theme 'WhiteText'
    $btnRun.Location = New-Object System.Drawing.Point(130, 1)
    $btnRun.Size = New-Object System.Drawing.Size(24, 24)
    $btnRun.Cursor = 'Hand'
    if ($toolTip) { $toolTip.SetToolTip($btnRun, "Execute Command") }
    $panel.Controls.Add($btnRun)

    
    $btnKey = New-Object System.Windows.Forms.Button
    $btnKey.Text = if ($cmd.Hotkey) { $cmd.Hotkey } else { "Key" }
    $btnKey.FlatStyle = 'Flat'
    $btnKey.FlatAppearance.BorderSize = 0
    $btnKey.BackColor = Get-Theme 'AccentBlue'
    $btnKey.ForeColor = Get-Theme 'WhiteText'
    $btnKey.Location = New-Object System.Drawing.Point(156, 1)
    $btnKey.Size = New-Object System.Drawing.Size(45, 24)
    $btnKey.Font = Get-UIFont 'Segoe UI' 6
    $btnKey.Cursor = 'Hand'
    if ($toolTip) { $toolTip.SetToolTip($btnKey, "Assign Hotkey") }
    $panel.Controls.Add($btnKey)

    
    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = "X"
    $btnDel.FlatStyle = 'Flat'
    $btnDel.FlatAppearance.BorderSize = 0
    $btnDel.BackColor = Get-Theme 'ErrorColorAlt'
    $btnDel.ForeColor = Get-Theme 'WhiteText'
    $btnDel.Location = New-Object System.Drawing.Point(203, 1)
    $btnDel.Size = New-Object System.Drawing.Size(24, 24)
    $btnDel.Cursor = 'Hand'
    if ($toolTip) { $toolTip.SetToolTip($btnDel, "Delete Command") }
    $panel.Controls.Add($btnDel)

    
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
                    
                    try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($cmd.Id)" } catch {}
                }

                $cmd.Hotkey = $newKey
                $cmd.HotkeyId = $null

                if ($newKey) {
                    try {
                        $scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($data.InstanceId)' -CommandId '$($cmd.Id)'")
                        $ownerKey = "quickcmd_$($data.InstanceId)_$($cmd.Id)"
                        
                        $ownerLabel = $null
                        if ($cmd.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($cmd.Name)) { $ownerLabel = $cmd.Name }
                        elseif ($cmd.PSObject.Properties['CommandText'] -and -not [string]::IsNullOrEmpty($cmd.CommandText)) { $ownerLabel = $cmd.CommandText }
                        else { $ownerLabel = "Invoke Command: $($cmd.Id)" }
                        $cmd.HotkeyId = SetHotkey -KeyCombinationString $newKey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel
                    } catch {}
                }
                
                $this.Text = if ($newKey) { $newKey } else { "Key" }
                UpdateChatCommanderSettings $data -forceWrite
                if (Test-DashboardCommand 'RefreshHotkeysList') { try { RefreshHotkeysList } catch {} }
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
            if (Test-DashboardCommand 'RefreshHotkeysList') { try { RefreshHotkeysList } catch {} }
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
            $data.BtnAdd.BackColor = Get-Theme 'SuccessColorAlt'
        }
    }.GetNewClosure())

    return $panel
}

function GetAllSavedCommands
{
    param($currentInstanceId)
    
    
    
    
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

    
    $sourceList = if ($formData.ShowAllMode) { $formData.GlobalCommandList } else { $formData.SavedCommands }
    
    
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
        
        
        AdjustChatCommanderFormHeight $form
    } else {
        
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
        
        
        $formData.LblPageInfo.Text = "$($formData.CurrentPage) / $totalPages"
        $formData.BtnPrevPage.Enabled = ($formData.CurrentPage -gt 1)
        $formData.BtnNextPage.Enabled = ($formData.CurrentPage -lt $totalPages)
        
        
        AdjustChatCommanderFormHeight $form
        
        
        $formData.PanelPagination.Top = $formData.panelSavedCommands.Bottom + 5
    }
}

function AdjustChatCommanderFormHeight 
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
            $data.HeaderPanel.Width = 498
            
            
            $data.BtnHotkeyToggle.Left = 427
            $data.BtnShowHide.Left = 461
            $data.BtnClose.Left = 479
            
            
            
            
            $data.PanelNewCommand.Width = 490
            $data.ChkRunOnStart.Width = 490
            $data.ChkCloseAfterRun.Width = 490
            $data.TxtCommandText.Width = 484
            $data.BtnAdd.Left = 428
            $data.Name.Width = 420
            
            $data.BtnRunSequence.Width = 490
            $data.BtnShowAll.Left = 455
            $data.txtSearchCommands.Width = 445
            
            $data.GridAllCommands.Width = 490
            $data.GridAllCommands.Height = 300 
            
            
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

        
        $form.Width = 250
        $data.HeaderPanel.Width = 248
        
        
        $data.BtnHotkeyToggle.Left = 177
        $data.BtnShowHide.Left = 211
        $data.BtnClose.Left = 229
        
        
        $data.PanelNewCommand.Width = 240
        $data.ChkRunOnStart.Width = 240
        $data.ChkCloseAfterRun.Width = 240
        $data.TxtCommandText.Width = 234
        $data.BtnAdd.Left = 178
        $data.Name.Width = 170
        
        $data.BtnRunSequence.Width = 240
        $data.BtnShowAll.Left = 205
        $data.txtSearchCommands.Width = 195
        
        $data.GridAllCommands.Width = 240

        
        $baseHeight = 313 
        $panel = $data.panelSavedCommands
        
        
        $panel.AutoSize = $true
        $panel.PerformLayout()
        $preferredPanelHeight = $panel.Height
        
        
        $paginationHeight = 0

        $newHeight = $baseHeight + $preferredPanelHeight + $paginationHeight + 10 
        
        $screen = [System.Windows.Forms.Screen]::FromPoint($form.Location)
        $workingArea = $screen.WorkingArea
        $distToBottom = $workingArea.Bottom - $form.Top
        $maxHeight = $distToBottom - 20
        if ($maxHeight -lt 285) { $maxHeight = 285 } 

        if ($newHeight -gt $maxHeight) {
            $newHeight = $maxHeight
            
            $panel.AutoSize = $false
            $panel.Height = $maxHeight - $baseHeight - $paginationHeight - 10
            $panel.AutoScroll = $true
        } else {
            
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
        
        if ($data.EditingCommandId) {
            $data.BtnAdd.Text = ([char]0x2705) + ' Save'
            $data.BtnAdd.BackColor = Get-Theme 'SuccessColorAlt'
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

        
        $existing = $data.GlobalCommandList | Where-Object { $_.Name -eq $newCommandName } | Select-Object -First 1
        
        if ($existing -and ($null -eq $data.EditingCommandId -or $existing.Id -ne $data.EditingCommandId)) {
             $newName = ShowInputBox -Title "Duplicate Name" -Prompt "A command with this name already exists.`nPlease enter a unique name:" -DefaultText "$newCommandName (1)"
             if ([string]::IsNullOrWhiteSpace($newName)) { return }
             $newCommandName = $newName
             
             $existing2 = $data.GlobalCommandList | Where-Object { $_.Name -eq $newCommandName } | Select-Object -First 1
             if ($existing2) {
                Show-DarkMessageBox "Name still duplicate. Action cancelled." "Error" "OK" "Error"
                return
             }
        }

        if ($data.EditingCommandId)
        {
            
            $commandToUpdate = $data.GlobalCommandList | Where-Object { $_.Id -eq $data.EditingCommandId } | Select-Object -First 1
            if ($commandToUpdate) {
                $commandToUpdate.CommandText = $newCommandText
                $commandToUpdate.Name = $newCommandName
            }
            
            $savedCmd = $data.SavedCommands | Where-Object { $_.Id -eq $data.EditingCommandId } | Select-Object -First 1
            if ($savedCmd) { $savedCmd.CommandText = $newCommandText; $savedCmd.Name = $newCommandName }

            $data.EditingCommandId = $null 
            $data.BtnAdd.Text = ([char]0x2795) + ' Add' 
            $data.BtnAdd.BackColor = Get-Theme 'ButtonBg' 
        }
        else
        {
            
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
        
        $data.TxtCommandText.Text = "" 
        $data.Name.Text = "New" 

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
            $this.Text = ([char]0x25B2) 
        }
        else
        {
            $data.OriginalHeight = $form.Height
            $form.Height = 26
            $data.IsCollapsed = $true
            $this.Text = ([char]0x25BC) 
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
            
            if ($data.RunningTimer) {
                StopChatCommanderSequence $data
            }

            
            foreach ($cmd in $data.SavedCommands) {
                if ($cmd.HotkeyId) {
                    try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($cmd.Id)" } catch {}
                    try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey $data.InstanceId } catch {}
                }
            }
            $data.SavedCommands.Clear()
            

            
            if ($data.GlobalHotkeyId) {
                 $ownerKey = "global_toggle_$($data.InstanceId)"
                 try { UnregisterHotkeyInstance -Id $data.GlobalHotkeyId -OwnerKey $ownerKey } catch {}
                 $data.GlobalHotkeyId = $null
                 $data.GlobalHotkey = $null
            }
            RefreshSavedCommandsUI $form

            
            $data.TxtCommandText.Text = ""
            
            $data.Name.Text = "New"
            $data.BtnHotkeyToggle.Checked = $true
            $data.ChkRunOnStart.Checked = $true
            $data.ChkCloseAfterRun.Checked = $false
            $minX = $data.PositionSliderX.Minimum
            $maxX = $data.PositionSliderX.Maximum
            $minY = $data.PositionSliderY.Minimum
            $maxY = $data.PositionSliderY.Maximum
            $data.PositionSliderX.Value = [Math]::Max($minX, [Math]::Min($maxX, $data.Form.Left))
            $data.PositionSliderY.Value = [Math]::Max($minY, [Math]::Min($maxY, $data.Form.Top))
            $data.EditingCommandId = $null
            $data.BtnAdd.Text = ([char]0x2795) + ' Add'
            $data.BtnAdd.BackColor = Get-Theme 'ButtonBg'

            
            $profilePrefix = $null
            if (Test-DashboardCommand 'FindOrCreateProfile')
            {
                $profilePrefix = FindOrCreateProfile $data.WindowTitle -ProfileHint (Get-RowProfileHint -formData $data)
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
                    
                    foreach ($cmd in $data.SavedCommands) {
                        if ($cmd.HotkeyId) {
                            try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($cmd.Id)" } catch {}
                            try { UnregisterHotkeyInstance -Id $cmd.HotkeyId -OwnerKey $data.InstanceId } catch {}
                        }
                    }
                    
                    $data.SavedCommands = $imported
                    
                    
                    foreach ($cmd in $data.SavedCommands) { 
                        if ($null -eq $cmd.PSObject.Properties['Hotkey']) { $cmd | Add-Member -MemberType NoteProperty -Name 'Hotkey' -Value $null -Force }
                        if ($null -eq $cmd.PSObject.Properties['HotkeyId']) { $cmd | Add-Member -MemberType NoteProperty -Name 'HotkeyId' -Value $null -Force }
                        
                        $cmd.HotkeyId = $null 

                                    if ($cmd.Hotkey) {
                                         try {
                                            $scriptBlock = [scriptblock]::Create("Invoke-SpecificChatCommand -InstanceId '$($data.InstanceId)' -CommandId '$($cmd.Id)'")
                                            $ownerKey = "quickcmd_$($data.InstanceId)_$($cmd.Id)"
                                            
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
            Register-OpenForm $newForm
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

    $formData.ChkRunOnStart.Add_CheckedChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        if ($data.IsLoading) { return }
        UpdateChatCommanderSettings $data -forceWrite
    })

    $formData.ChkCloseAfterRun.Add_CheckedChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        if ($data.IsLoading) { return }
        UpdateChatCommanderSettings $data -forceWrite
    })

    

    $formData.Name.Add_TextChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        UpdateChatCommanderSettings $form.Tag -forceWrite
    })

    

    $formData.PositionSliderX.Add_ValueChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        if (-not $data.PSObject.Properties['IsDragging'] -or -not $data.IsDragging) {
            if ($data.PSObject.Properties['SuppressPersist'] -and $data.SuppressPersist) {
                UpdateChatCommanderSettings $data
            }
            else {
                UpdateChatCommanderSettings $data -forceWrite
            }
        }
    })

    $formData.PositionSliderY.Add_ValueChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        if (-not $data.PSObject.Properties['IsDragging'] -or -not $data.IsDragging) {
            if ($data.PSObject.Properties['SuppressPersist'] -and $data.SuppressPersist) {
                UpdateChatCommanderSettings $data
            }
            else {
                UpdateChatCommanderSettings $data -forceWrite
            }
        }
    })

    $formData.txtSearchCommands.Add_TextChanged({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        
        $form.Tag.CurrentPage = 1
        RefreshSavedCommandsUI $form
    })

    $formData.BtnShowAll.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag
        
        $data.ShowAllMode = -not $data.ShowAllMode
        $data.BtnShowAll.Text = if ($data.ShowAllMode) { "LIST" } else { "ALL" }
        $data.BtnShowAll.BackColor = if ($data.ShowAllMode) { Get-Theme 'SuccessColorAlt' } else { Get-Theme 'AccentBlue' }
        
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
        
        if ($e.ColumnIndex -eq 0) { 
            $newVal = -not ($row.Cells[0].Value -eq $true)
            $row.Cells[0].Value = $newVal
            
            if ($newVal) {
                
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
                
                $toRemove = $null
                foreach ($c in $data.SavedCommands) { if ($c.Id -eq $cmd.Id) { $toRemove = $c; break } }
                if ($toRemove) { 
                    if ($toRemove.HotkeyId) { try { UnregisterHotkeyInstance -Id $toRemove.HotkeyId -OwnerKey "quickcmd_$($data.InstanceId)_$($toRemove.Id)" } catch {} }
                    $data.SavedCommands.Remove($toRemove) 
                }
            }
            UpdateChatCommanderSettings $data -forceWrite
            if (Test-DashboardCommand 'RefreshHotkeysList') { try { RefreshHotkeysList } catch {} }
        }
        elseif ($e.ColumnIndex -eq 3) { 
            if ($data.RunningTimer) { StopChatCommanderSequence $data }
            $job = Invoke-ChatCommanderAction -formData $data -commandString $cmd.CommandText
            if ($job) {
                $data.ActiveJobs.Add($job) | Out-Null
                [System.Threading.Tasks.Task]::Run([System.Action]{
                    $job.AsyncResult.AsyncWaitHandle.WaitOne() | Out-Null
                    try { $job.PowerShell.EndInvoke($job.AsyncResult); $job.PowerShell.Dispose() } catch {}
                    $form = $this.FindForm()
                    if ($form -and -not $form.IsDisposed) {
                        $form.BeginInvoke([System.Action]{
                            if ($data.ActiveJobs) { $data.ActiveJobs.Remove($job) }
                        })
                    }
                }) | Out-Null
            }
        }
        elseif ($e.ColumnIndex -eq 4) { 
            
            
            $data.TxtCommandText.Text = $cmd.CommandText
            $data.Name.Text = $cmd.Name
            $data.EditingCommandId = $cmd.Id
            $data.BtnAdd.Text = ([char]0x2705) + ' Save'
            $data.BtnAdd.BackColor = Get-Theme 'SuccessColorAlt'
            
            
            
        }
        elseif ($e.ColumnIndex -eq 5) { 
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

function UpdateChatCommanderSliderValuesFromFormPosition {
    param($formData, $windowRect)

    if (-not $formData -or -not $formData.PSObject.Properties['PositionSliderX'] -or -not $formData.PSObject.Properties['PositionSliderY']) { return }
    if (-not $windowRect) { return }

    try {
        $form = $formData.Form
        if (-not $form -or $form.IsDisposed) { return }

        $virtual = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $maxX = $virtual.Right
        $maxY = $virtual.Bottom
        $minX = $virtual.Left
        $minY = $virtual.Top

        $sliderX = [Math]::Max($minX, [Math]::Min($maxX, $form.Left))
        $formData.PositionSliderX.Value = $sliderX

        $sliderY = [Math]::Max($minY, [Math]::Min($maxY, $form.Top))
        $formData.PositionSliderY.Value = $sliderY
    }
    catch {
        Write-Verbose ('CHATCOMMANDER: Error updating slider values from form position: {0}' -f $_.Exception.Message)
    }
}

function AddChatCommanderDragHandlers {
    param($formData)

    if (-not $formData -or -not $formData.PSObject.Properties['Form']) { return }
    $form = $formData.Form
    if ($form.IsDisposed) { return }

    if (-not $formData.PSObject.Properties['IsDragging']) {
        $formData | Add-Member -MemberType NoteProperty -Name 'IsDragging' -Value $false
    }
    if (-not $formData.PSObject.Properties['DragStart']) {
        $formData | Add-Member -MemberType NoteProperty -Name 'DragStart' -Value @{ X = 0; Y = 0; FormX = 0; FormY = 0 }
    }
    if (-not $formData.PSObject.Properties['LastUpdateTime']) {
        $formData | Add-Member -MemberType NoteProperty -Name 'LastUpdateTime' -Value ([DateTime]::MinValue)
    }

    $mouseDown = {
        $f = $this.FindForm()
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            if ($f.Tag -and -not $f.Tag.IsDragging) {
                $f.Tag.IsDragging = $true
                $f.Tag.DragStart.X = [System.Windows.Forms.Cursor]::Position.X
                $f.Tag.DragStart.Y = [System.Windows.Forms.Cursor]::Position.Y
                $f.Tag.DragStart.FormX = $f.Left
                $f.Tag.DragStart.FormY = $f.Top
            }
        }
    }

    $mouseMove = {
        $f = $this.FindForm()
        if ($f.Tag -and $f.Tag.IsDragging) {
            try {
                $currentTime = [DateTime]::Now
                if (($currentTime - $f.Tag.LastUpdateTime).TotalMilliseconds -lt 20) { return }
                $f.Tag.LastUpdateTime = $currentTime

                $cursorX = [System.Windows.Forms.Cursor]::Position.X
                $cursorY = [System.Windows.Forms.Cursor]::Position.Y

                $deltaX = $cursorX - $f.Tag.DragStart.X
                $deltaY = $cursorY - $f.Tag.DragStart.Y

                $newLeft = $f.Tag.DragStart.FormX + $deltaX
                $newTop  = $f.Tag.DragStart.FormY + $deltaY

                $f.Left = $newLeft
                $f.Top  = $newTop

                if ($f.Tag) {
                    $timerKey = "ChatCommanderPosition_$($f.Tag.InstanceId)"
                    if ($global:DashboardConfig.Resources.Timers.Contains($timerKey)) {
                        $timer = $global:DashboardConfig.Resources.Timers[$timerKey]
                        if ($timer -and $timer.Tag) {
                            $timer.Tag.RealLeft = $newLeft
                            $timer.Tag.RealTop  = $newTop
                        }
                    }
                }

                if ($f.Tag.SelectedWindow -and $f.Tag.SelectedWindow -ne [IntPtr]::Zero) {
                    $rect = New-Object Custom.Native+RECT
                    if ([Custom.Native]::GetWindowRect($f.Tag.SelectedWindow, [ref]$rect)) {
                        UpdateChatCommanderSliderValuesFromFormPosition $f.Tag $rect
                    }
                }
            }
            catch {
                Write-Verbose ('CHATCOMMANDER: Error during mouse move: {0}' -f $_.Exception.Message)
            }
        }
    }

    $mouseUp = {
        $f = $this.FindForm()
        if ($f.Tag) {
            $f.Tag.IsDragging = $false
            if ($f.Tag.SelectedWindow -and $f.Tag.SelectedWindow -ne [IntPtr]::Zero) {
                try {
                    $rect = New-Object Custom.Native+RECT
                    if ([Custom.Native]::GetWindowRect($f.Tag.SelectedWindow, [ref]$rect)) {
                        UpdateChatCommanderSliderValuesFromFormPosition $f.Tag $rect
                        UpdateChatCommanderSettings $f.Tag -forceWrite
                    }
                } catch {}
            }
        }
    }

    $form.Add_MouseDown($mouseDown)
    $form.Add_MouseMove($mouseMove)
    $form.Add_MouseUp($mouseUp)

    if ($formData.HeaderPanel) {
        $formData.HeaderPanel.Add_MouseDown($mouseDown)
        $formData.HeaderPanel.Add_MouseMove($mouseMove)
        $formData.HeaderPanel.Add_MouseUp($mouseUp)
    }

    if ($formData.PSObject.Properties['LabelWinTitle'] -and $formData.LabelWinTitle) {
        $formData.LabelWinTitle.Add_MouseDown($mouseDown)
        $formData.LabelWinTitle.Add_MouseMove($mouseMove)
        $formData.LabelWinTitle.Add_MouseUp($mouseUp)
    }
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion