<# macro.psm1 #>

$script:_MacroKeyScriptStr = $null

#region Helper Functions

function InitMacroRunspace {
    param($InstanceId)
    
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
    $assemblies = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { 
        $_.FullName -match 'Custom.Native' -or 
        $_.FullName -match 'Custom.Ftool' -or 
        $_.FullName -match 'Custom.Win32'
    }
    
    foreach ($asm in $assemblies) {
        if (-not [string]::IsNullOrEmpty($asm.Location)) {
            $iss.Assemblies.Add((New-Object System.Management.Automation.Runspaces.SessionStateAssemblyEntry($asm.Location)))
        }
    }

    $maxThreads = [Math]::Max(15, [Environment]::ProcessorCount * 2)
    $pool = [runspacefactory]::CreateRunspacePool(1, $maxThreads, $iss, $Host)
    $pool.ThreadOptions = 'ReuseThread'
    $pool.Open()
    
    return $pool
}

function Get-MacroKeyScriptBlock {
    return {
        param($Handle, $Vk, $HoldTime, $IsMouse, $Mods, $IsAlt, $IsCtrl)

        $WM_KEYDOWN = 0x0100; $WM_KEYUP = 0x0101
        $WM_LBUTTONDOWN = 0x0201; $WM_LBUTTONUP = 0x0202
        $WM_RBUTTONDOWN = 0x0204; $WM_RBUTTONUP = 0x0205
        $WM_MBUTTONDOWN = 0x0207; $WM_MBUTTONUP = 0x0208
        $WM_XBUTTONDOWN = 0x020B; $WM_XBUTTONUP = 0x020C

        try {
            $point = New-Object Custom.Win32Point
            [Custom.Win32MouseUtils]::GetCursorPos([ref]$point) | Out-Null
            [Custom.Win32MouseUtils]::ScreenToClient($Handle, [ref]$point) | Out-Null
            $lParam = [Custom.Win32MouseUtils]::MakeLParam($point.X, $point.Y)

            if (-not $IsMouse) {
                $lParam = 1
            }

            $msgDown = 0; $msgUp = 0; $wParamDown = [IntPtr]::Zero; $wParamUp = [IntPtr]::Zero

            if ($IsMouse) {
                $MK_LBUTTON = 0x0001; $MK_RBUTTON = 0x0002; $MK_SHIFT = 0x0004
                $MK_CONTROL = 0x0008; $MK_MBUTTON = 0x0010; $MK_XBUTTON1 = 0x0020; $MK_XBUTTON2 = 0x0040
                
                $wDown = 0
                if ($Mods -contains 0x11) { $wDown = $wDown -bor $MK_CONTROL }
                if ($Mods -contains 0x10) { $wDown = $wDown -bor $MK_SHIFT }

                $targetFlag = 0; $xData = 0
                switch ($Vk) {
                    0x01 { $targetFlag = $MK_LBUTTON; $msgDown = $WM_LBUTTONDOWN; $msgUp = $WM_LBUTTONUP }
                    0x02 { $targetFlag = $MK_RBUTTON; $msgDown = $WM_RBUTTONDOWN; $msgUp = $WM_RBUTTONUP }
                    0x04 { $targetFlag = $MK_MBUTTON; $msgDown = $WM_MBUTTONDOWN; $msgUp = $WM_MBUTTONUP }
                    0x05 { $targetFlag = $MK_XBUTTON1; $xData = 0x00010000; $msgDown = $WM_XBUTTONDOWN; $msgUp = $WM_XBUTTONUP }
                    0x06 { $targetFlag = $MK_XBUTTON2; $xData = 0x00020000; $msgDown = $WM_XBUTTONDOWN; $msgUp = $WM_XBUTTONUP }
                }
                
                $wDown = $wDown -bor $targetFlag
                $wUp = $wDown -bxor $targetFlag
                if ($xData -ne 0) { $wDown = $wDown -bor $xData; $wUp = $wUp -bor $xData }
                
                $wParamDown = [IntPtr]$wDown
                $wParamUp = [IntPtr]$wUp
            }
            else {
                $msgDown = $WM_KEYDOWN
                $msgUp = $WM_SYSKEYUP
                $wParamDown = [IntPtr]$Vk
                $wParamUp = [IntPtr]$Vk
            }

            if ($IsAlt) { [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, 0x12, 0) }
            if ($IsCtrl) { [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, 0x11, 0x001D0001) }
            if ($Mods -contains 0x10) { [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, 0x10, 0) }

            foreach ($m in $Mods) {
                [Custom.Native]::keybd_event([byte]$m, 0, 0, [IntPtr]::Zero)
                [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, $m, 1)
            }

            [Custom.Ftool]::fnPostMessage($Handle, $msgDown, $wParamDown, $lParam)

            if ($HoldTime -gt 0) {
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while ($sw.ElapsedMilliseconds -lt $HoldTime) {
                    Start-Sleep -Milliseconds 5
                    if ($sw.ElapsedMilliseconds -lt $HoldTime -and -not $IsMouse) {
                        [Custom.Ftool]::fnPostMessage($Handle, $msgDown, $wParamDown, 0x40000000)
                    }
                }
                $sw.Stop()
            }
            else {
                Start-Sleep -Milliseconds 5
            }

            $finalLParam = $lParam
            
            if (-not $IsMouse) {
                $finalLParam = $lParam -bor 0xC0000000
            }

            [Custom.Ftool]::fnPostMessage($Handle, $msgUp, $wParamUp, $finalLParam)
        }
        finally {
            foreach ($m in $Mods | Sort-Object -Descending) {
                [Custom.Native]::keybd_event([byte]$m, 0, 2, [IntPtr]::Zero)
            }
            if ($IsCtrl) { [Custom.Ftool]::fnPostMessage($Handle, $WM_SYSKEYUP, 0x11, 0x001D0001) }
            if ($IsAlt) { [Custom.Ftool]::fnPostMessage($Handle, $WM_SYSKEYUP, 0x12, 0xC0000000) }
            if ($Mods -contains 0x10) { [Custom.Ftool]::fnPostMessage($Handle, $WM_SYSKEYUP, 0x10, 0xC0000000) }
            
            if ($Vk -and $Vk -gt 0) {
                if ($IsMouse) {
                    $msgUp = 0
                    switch ($Vk) {
                        0x01 { $msgUp = $WM_LBUTTONUP }
                        0x02 { $msgUp = $WM_RBUTTONUP }
                        0x04 { $msgUp = $WM_MBUTTONUP }
                        0x05 { $msgUp = $WM_XBUTTONUP }
                        0x06 { $msgUp = $WM_XBUTTONUP }
                    }
                    if ($msgUp -ne 0) {
                        [Custom.Ftool]::fnPostMessage($Handle, $msgUp, [IntPtr]0, $lParam)
                    }
                } else {
                    [Custom.Native]::keybd_event([byte]$Vk, 0, 2, [IntPtr]::Zero)
                }
            }
        }
    }
}

function Invoke-MacroKeyAction {
    param(
        $formData,
        $keyString = $null,
        [int]$holdTime = 50,
        $cachedVk = $null,
        $cachedIsMouse = $null,
        $cachedMods = $null,
        $cachedIsAlt = $null,
        $cachedIsCtrl = $null
    )

    if ($formData.SelectedWindow -eq [IntPtr]::Zero) { return $null }

    if (-not $formData.RunspacePool -or $formData.RunspacePool.RunspacePoolStateInfo.State -ne 'Opened') {
        Write-Verbose "Macro Runspace closed/broken. Attempting to restart..."
        try {
            $scriptBlock = [scriptblock]::Create("ToggleMacroInstance -InstanceId '$($formData.InstanceId)'")
            $ownerKey = "macroinst_$($formData.InstanceId)"
            $ownerLabel = $null
            if ($formData.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($formData.Name)) { $ownerLabel = $formData.Name }
            elseif ($formData.PSObject.Properties['WindowTitle'] -and -not [string]::IsNullOrEmpty($formData.WindowTitle)) { $ownerLabel = $formData.WindowTitle }
            else { $ownerLabel = "Macro: $($formData.InstanceId)" }
            $formData.HotkeyId = SetHotkey -KeyCombinationString $hk -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel
        } 
        catch {
            Write-Verbose "Failed to restart Runspace: $_"
            return $null
        }
    }

    if ($null -ne $cachedVk) {
        $virtualKeyCode = $cachedVk
        $isMouse        = $cachedIsMouse
        $mods           = $cachedMods
        $isAlt          = $cachedIsAlt
        $isCtrl         = $cachedIsCtrl
    }
    else {
        $parsedKey = ParseKeyString -KeyCombinationString $keyString
        if (-not $parsedKey -or -not $parsedKey.Primary) { return $null }
        $keyMappings    = GetVirtualKeyMappings
        $virtualKeyCode = $keyMappings[$parsedKey.Primary]
        if (-not $virtualKeyCode) { return $null }
        $isMouse = ($virtualKeyCode -ge 0x01 -and $virtualKeyCode -le 0x06)
        $mods = @(); $isAlt = $false; $isCtrl = $false
        foreach ($modName in $parsedKey.Modifiers) {
            switch ($modName.ToUpper()) {
                'CTRL'  { $mods += 0x11; $isCtrl = $true }
                'ALT'   { $mods += 0x12; $isAlt  = $true }
                'SHIFT' { $mods += 0x10 }
            }
        }
        if (-not $isMouse) { $isAlt = $false; $isCtrl = $false }
    }

    if (-not $script:_MacroKeyScriptStr) {
        $script:_MacroKeyScriptStr = (Get-MacroKeyScriptBlock).ToString()
    }

    $ps = [System.Management.Automation.PowerShell]::Create()
    try {
        $ps.RunspacePool = $formData.RunspacePool
        $ps.AddScript($script:_MacroKeyScriptStr) | Out-Null
        $ps.AddArgument($formData.SelectedWindow) | Out-Null
        $ps.AddArgument($virtualKeyCode) | Out-Null
        $ps.AddArgument($holdTime) | Out-Null
        $ps.AddArgument($isMouse) | Out-Null
        $ps.AddArgument($mods) | Out-Null
        $ps.AddArgument($isAlt) | Out-Null
        $ps.AddArgument($isCtrl) | Out-Null

        return @{ 
            PowerShell  = $ps
            AsyncResult = $ps.BeginInvoke()
        }
    }
    catch {
        Write-Verbose "Macro Invoke Error: $_"
        if ($ps) { $ps.Dispose() }
        return $null
    }
}

function CreateMacroPositionTimer {
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
            try {
                if (-not $s -or -not $s.Tag) {
                    $s.Stop(); $s.Dispose()
                    return
                }

                $timerData = $s.Tag
                if (-not $timerData -or -not $timerData.ContainsKey('Form') -or -not $timerData['Form'] -or $timerData['Form'].IsDisposed) { $s.Stop(); $s.Dispose(); return }

                $form = $timerData.Form
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
                }
                else {
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
                                        $formData.PositionSliderX.Value = $newX

                                        $minY = $formData.PositionSliderY.Minimum
                                        $maxY = $formData.PositionSliderY.Maximum
                                        $minY = [Math]::Max($virtual.Top, [Math]::Min($virtual.Bottom, $minY))
                                        $maxY = [Math]::Max($virtual.Top, [Math]::Min($virtual.Bottom, $maxY))
                                        if ($minY -gt $maxY) { $tmp = $minY; $minY = $maxY; $maxY = $tmp }
                                        $newY = [int]([Math]::Max($minY, [Math]::Min($maxY, $formData.PositionSliderY.Value + $dy)))
                                        $formData.PositionSliderY.Value = $newY

                                        $timerData.RealLeft = $timerData.RealLeft + $dx
                                        $timerData.RealTop = $timerData.RealTop + $dy
                                    }
                                    catch { }
                                }
                            }
                            $timerData.ClientRect = $currRect
                        }
                    }
                }
                catch { }

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
                $newRealTop = $timerData.RealTop + ($targetTop - $timerData.RealTop) * 0.2
                $timerData.RealLeft = $newRealLeft
                $timerData.RealTop = $newRealTop

                $targetIntLeft = [int]$newRealLeft
                $targetIntTop = [int]$newRealTop
                if ($form.Left -ne $targetIntLeft) { $form.Left = $targetIntLeft }
                if ($form.Top -ne $targetIntTop) { $form.Top = $targetIntTop }

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
                    }
                    else {
                        if ($timerData.ZState -ne 'standard') { [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::HWND_NOTOPMOST, 0, 0, 0, 0, $flags) | Out-Null; $timerData.ZState = 'standard' }
                        $next = [Custom.Native]::GetWindow($formHandle, 2)
                        $amIAboveGame = $false; $loopCount = 0
                        while ($next -ne [IntPtr]::Zero -and $loopCount -lt 50) { if ($next -eq $linkedHandle) { $amIAboveGame = $true; break }; $next = [Custom.Native]::GetWindow($next, 2); $loopCount++ }
                        if (-not $amIAboveGame) { $bottomOfCluster = [Custom.Native]::GetWindow($linkedHandle, 3); if ($bottomOfCluster -ne [IntPtr]::Zero) { [Custom.Native]::PositionWindow($formHandle, $bottomOfCluster, 0, 0, 0, 0, $flags) | Out-Null } else { [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::TopWindowHandle, 0, 0, 0, 0, $flags) | Out-Null } }
                    }
                }
                else {
                    if ($timerData.ZState -ne 'standard') { [Custom.Native]::PositionWindow($formHandle, [Custom.Native]::HWND_NOTOPMOST, 0, 0, 0, 0, $flags) | Out-Null; $timerData.ZState = 'standard' }
                }

            }
            catch {
                Write-Verbose ('MACRO: Position timer critical error: {0}' -f $_.Exception.Message)
            }
        })
    
    $positionTimer.Start()
    $global:DashboardConfig.Resources.Timers["macroPosition_$($formData.InstanceId)"] = $positionTimer
}

function RepositionSequences {
    param($form)
    
    if (-not $form -or $form.IsDisposed) { return }
    
    $form.SuspendLayout()
    try {
        $baseHeight = 140 
        $rowHeight = 70
        
        $newHeight = $baseHeight
        $position = 0
        
        $formData = $form.Tag
        if ($formData.SequencePanels) {
            foreach ($panel in $formData.SequencePanels) {
                if ($panel -and -not $panel.IsDisposed) {
                    $newTop = $baseHeight + ($position * $rowHeight)
                    $panel.Top = $newTop
                    $position++
                }
            }
            if ($position -gt 0) {
                $newHeight = $baseHeight + ($position * $rowHeight)
            }
        }
        
        if (-not $formData.IsCollapsed) {
            $finalHeight = $newHeight + 10
            $form.Height = $finalHeight
            $formData.OriginalHeight = $finalHeight
            $formData.PositionSliderY.Height = $finalHeight - 20
        }
    }
    finally {
        $form.ResumeLayout()
    }
}

function StartMacroSequence {
    param($formData)

    $steps = @()

    $mainKey = $formData.BtnKeySelect.Text
    $mainInterval = 1000
    if ([int]::TryParse($formData.Interval.Text, [ref]$mainInterval)) { if ($mainInterval -lt 1) { $mainInterval = 1 } }
    
    $mainHoldEnabled = $formData.BtnHoldKeyToggle.Checked
    $mainHoldInterval = $formData.TxtHoldKeyInterval.Text
    $mainWaitEnabled = $formData.BtnWaitToggle.Checked

    if (-not [string]::IsNullOrWhiteSpace($mainKey) -and $mainKey -ne 'none') {
        $steps += [PSCustomObject]@{ 
            Key          = $mainKey; 
            Interval     = $mainInterval; 
            HoldEnabled  = $mainHoldEnabled; 
            HoldInterval = $mainHoldInterval;
            WaitEnabled  = $mainWaitEnabled
        }
    }

    if ($formData.SequencePanels) {
        foreach ($panel in $formData.SequencePanels) {
            if ($panel -and -not $panel.IsDisposed) {
                $seqData = $panel.Tag
                $seqKey = $seqData.BtnKeySelect.Text
                $seqInterval = 1000
                if ([int]::TryParse($seqData.Interval.Text, [ref]$seqInterval)) { if ($seqInterval -lt 1) { $seqInterval = 1 } }
                $seqHoldEnabled = $seqData.BtnHoldKeyToggle.Checked
                $seqHoldInterval = $seqData.TxtHoldKeyInterval.Text
                $seqWaitEnabled = $seqData.BtnWaitToggle.Checked

                if (-not [string]::IsNullOrWhiteSpace($seqKey) -and $seqKey -ne 'none') {
                    $steps += [PSCustomObject]@{ 
                        Key          = $seqKey; 
                        Interval     = $seqInterval; 
                        HoldEnabled  = $seqHoldEnabled; 
                        HoldInterval = $seqHoldInterval;
                        WaitEnabled  = $seqWaitEnabled
                    }
                }
            }
        }
    }

    if ($steps.Count -eq 0) {
        Show-DarkMessageBox 'No valid keys defined in the macro sequence.' 'Macro Error' 'Ok' 'Warning'
        return
    }

    $keyMappings = GetVirtualKeyMappings
    $stepDataList = @()
    foreach ($step in $steps) {
        $parsedKey = ParseKeyString -KeyCombinationString $step.Key
        if ($parsedKey -and $parsedKey.Primary) {
            $vk     = $keyMappings[$parsedKey.Primary]
            $mouse  = ($vk -ge 0x01 -and $vk -le 0x06)
            $mds    = @(); $alt = $false; $ctrl = $false
            foreach ($modName in $parsedKey.Modifiers) {
                switch ($modName.ToUpper()) {
                    'CTRL'  { $mds += 0x11; $ctrl = $true }
                    'ALT'   { $mds += 0x12; $alt  = $true }
                    'SHIFT' { $mds += 0x10 }
                }
            }
            if (-not $mouse) { $alt = $false; $ctrl = $false }
            
            $holdTime = 0
            if ($step.HoldEnabled) {
                if (-not [int]::TryParse($step.HoldInterval, [ref]$holdTime)) { $holdTime = 5 }
            }

            $reverseMods = @($mds | Sort-Object -Descending)
            
            $stepDataList += @{
                Vk          = [int]$vk
                IsMouse     = [bool]$mouse
                Mods        = [int[]]$mds
                ReverseMods = [int[]]$reverseMods
                IsAlt       = [bool]$alt
                IsCtrl      = [bool]$ctrl
                HoldEnabled = [bool]$step.HoldEnabled
                HoldTime    = [int]$holdTime
                WaitEnabled = [bool]$step.WaitEnabled
                Interval    = [int][Math]::Max(1, $step.Interval)
            }
        }
    }

    if ($stepDataList.Count -eq 0) {
        Show-DarkMessageBox 'No valid key mappings could be resolved.' 'Macro Error' 'Ok' 'Warning'
        return
    }

    $formData.ActiveSteps = $steps
    $formData.CurrentStepIndex = 0
    
    if (-not $formData.PSObject.Properties['ActiveJobs']) {
        $formData | Add-Member -MemberType NoteProperty -Name 'ActiveJobs' -Value ([System.Collections.ArrayList]::new())
    }
    else {
        foreach ($job in $formData.ActiveJobs) {
            try {
                if ($job.PowerShell -and $job.PowerShell.InvocationStateInfo.State -eq 'Running') {
                    $job.PowerShell.Stop()
                }
                if ($job.PowerShell) { $job.PowerShell.Dispose() }
            }
            catch {}
        }
        $formData.ActiveJobs.Clear()
    }

    if (-not $formData.RunspacePool -or $formData.RunspacePool.RunspacePoolStateInfo.State -ne 'Opened') {
        if ($formData.RunspacePool) { try { $formData.RunspacePool.Dispose() } catch {} }
        $formData.RunspacePool = InitMacroRunspace -InstanceId $formData.InstanceId
    }
    
    $formData.BtnStart.Enabled = $false
    $formData.BtnStart.Visible = $false
    $formData.BtnStop.Enabled = $true
    $formData.BtnStop.Visible = $true

    if (-not $formData.PSObject.Properties['IsMacroRunning']) {
        $formData | Add-Member -MemberType NoteProperty -Name 'IsMacroRunning' -Value $true
    }
    $formData.IsMacroRunning = $true

    if (-not $formData.PSObject.Properties['Stopping']) {
        $formData | Add-Member -MemberType NoteProperty -Name 'Stopping' -Value $false
    }
    $formData.Stopping = $false

    $isLoop = $formData.BtnLoopToggle.Checked

    $seqTimer = New-Object System.Windows.Forms.Timer
    $seqTimer.Interval = 10
    $seqTimer.Tag = @{
        StepDataList = $stepDataList
        StepIndex    = 0
        State        = 'Delay'
        DelaySW      = [System.Diagnostics.Stopwatch]::StartNew()
        CurrentJob   = $null
        FormData     = $formData
        IsLoop       = $isLoop
    }

    $seqTimer.Add_Tick({
        param($s, $e)
        $td    = $s.Tag
        $fData = $td.FormData

        if (-not $fData.IsMacroRunning -or $fData.Stopping) { $s.Stop(); return }
        if (-not $fData.Form -or $fData.Form.IsDisposed) { $s.Stop(); $s.Dispose(); return }

        if ($fData.ActiveJobs -and $fData.ActiveJobs.Count -gt 0) {
            for ($i = $fData.ActiveJobs.Count - 1; $i -ge 0; $i--) {
                $j = $fData.ActiveJobs[$i]
                if ($j -and $j.AsyncResult -and $j.AsyncResult.IsCompleted) {
                    try {
                        DisposeManagedRunspace @{ PowerShell = $j.PowerShell; AsyncResult = $j.AsyncResult }
                    } catch {}
                    $fData.ActiveJobs.RemoveAt($i)
                }
            }
        }

        $step = $td.StepDataList[$td.StepIndex]

        if ($td.State -eq 'Delay') {
            if ($td.DelaySW.ElapsedMilliseconds -ge $step['Interval']) {
                $job = Invoke-MacroKeyAction `
                    -formData      $fData `
                    -holdTime      $step['HoldTime'] `
                    -cachedVk      $step['Vk'] `
                    -cachedIsMouse $step['IsMouse'] `
                    -cachedMods    $step['Mods'] `
                    -cachedIsAlt   $step['IsAlt'] `
                    -cachedIsCtrl  $step['IsCtrl']

                if ($job) { $fData.ActiveJobs.Add($job) | Out-Null }
                $td.CurrentJob = $job

                if ($step['WaitEnabled']) {
                    $td.State = 'Holding'
                } else {
                    $td.StepIndex++
                    if ($td.StepIndex -ge $td.StepDataList.Count) {
                        if ($td.IsLoop) { $td.StepIndex = 0 }
                        else {
                            $td.State = 'Done'
                            return
                        }
                    }
                    $td.State   = 'Delay'
                    $td.DelaySW = [System.Diagnostics.Stopwatch]::StartNew()
                }
            }
        } elseif ($td.State -eq 'Holding') {
            if (-not $td.CurrentJob -or $td.CurrentJob.AsyncResult.IsCompleted) {
                $td.StepIndex++
                if ($td.StepIndex -ge $td.StepDataList.Count) {
                    if ($td.IsLoop) { $td.StepIndex = 0 }
                    else { $s.Stop(); StopMacroSequence $fData -NaturalEnd; return }
                }
                $td.State   = 'Delay'
                $td.DelaySW = [System.Diagnostics.Stopwatch]::StartNew()
            }
        } elseif ($td.State -eq 'Done') {
            if (-not $td.CurrentJob -or $td.CurrentJob.AsyncResult.IsCompleted) {
                $s.Stop()
                StopMacroSequence $fData -NaturalEnd
                return
            }
        }
    })

    $seqTimer.Start()
    $formData.RunningTimer = $seqTimer
    $global:DashboardConfig.Resources.Timers["MacroTimer_$($formData.InstanceId)"] = $seqTimer
}

function StopMacroSequence {
    param($formData, [switch]$NaturalEnd)

    if ($formData.PSObject.Properties['IsMacroRunning']) {
        $formData.IsMacroRunning = $false
    }

    if ($formData.RunningTimer) {
        $formData.RunningTimer.Stop()
        $formData.RunningTimer.Dispose()
        $formData.RunningTimer = $null
    }
    
    if ($global:DashboardConfig.Resources.Timers.Contains("MacroTimer_$($formData.InstanceId)")) {
        $t = $global:DashboardConfig.Resources.Timers["MacroTimer_$($formData.InstanceId)"]
        if ($t) { $t.Stop(); $t.Dispose() }
        $global:DashboardConfig.Resources.Timers.Remove("MacroTimer_$($formData.InstanceId)")
    }

    if ($formData.ActiveJobs) {
        foreach ($job in $formData.ActiveJobs) {
            try {
                if (-not $NaturalEnd -and $job.PowerShell -and $job.PowerShell.InvocationStateInfo.State -eq 'Running') {
                    $job.PowerShell.Stop()
                }
                DisposeManagedRunspace @{ PowerShell = $job.PowerShell; AsyncResult = $job.AsyncResult }
            }
            catch {}
        }
        $formData.ActiveJobs.Clear()
    }

    $formData.BtnStart.Enabled = $true
    $formData.BtnStart.Visible = $true
    $formData.BtnStop.Enabled = $false
    $formData.BtnStop.Visible = $false
}

function ToggleMacroInstance {
    param($InstanceId)

    if (-not $global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId)) { return }
    $form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
    
    if ($form -and -not $form.IsDisposed -and $form.Tag.Type -eq 'Macro') {
        if ($form.InvokeRequired) {
            try {
                $form.BeginInvoke([System.Action] { ToggleMacroInstance -InstanceId $InstanceId }) | Out-Null
            }
            catch {}
        }

        try {
            $data = $form.Tag
            if ($data -and $data.RunningTimer) {
                StopMacroSequence $data
            }
            else {
                StartMacroSequence $data
            }
        }
        catch {
            Write-Verbose "MACRO: Error in ToggleMacroInstance: $_"
        }
    }
}

function ToggleMacroHotkeys {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstanceId,
        [Parameter(Mandatory = $true)]
        [bool]$ToggleState
    )
    
    if ($global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId)) {
        $form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
        if ($form -and -not $form.IsDisposed -and $form.InvokeRequired) {
            $form.BeginInvoke([System.Action] { ToggleMacroHotkeys -InstanceId $InstanceId -ToggleState $ToggleState }) | Out-Null
            return
        }
    }

    if (-not $global:DashboardConfig.Resources.InstanceHotkeysPaused) { $global:DashboardConfig.Resources.InstanceHotkeysPaused = @{} }
    
    $global:DashboardConfig.Resources.InstanceHotkeysPaused[$InstanceId] = (-not $ToggleState)

    try {
        $ownerKey = "macroinst_$InstanceId"
        if ($ToggleState) {
            try { ResumeHotkeysForOwner -OwnerKey $ownerKey } catch {}
        }
        else {
            try { PauseHotkeysForOwner -OwnerKey $ownerKey } catch {}
        }
    }
    catch {}

    if ($global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId)) {
        $form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
        if ($form -and -not $form.IsDisposed) {
            UpdateMacroSettings -formData $form.Tag -forceWrite
        }
        if (Get-Command RefreshHotkeysList -ErrorAction SilentlyContinue -Verbose:$False) {
            try { RefreshHotkeysList } catch {}
        }
    }
}

function CleanupMacroResources {
    param($instanceId)
    
    $timerKeysToRemove = @()
    foreach ($key in $global:DashboardConfig.Resources.Timers.Keys) {
        if ($key -eq "MacroTimer_$instanceId" -or $key -eq "macroPosition_$instanceId") {
            $timer = $global:DashboardConfig.Resources.Timers[$key]
            if ($timer) {
                $timer.Stop()
                $timer.Dispose()
            }
            $timerKeysToRemove += $key
        }
    }
    foreach ($key in $timerKeysToRemove) {
        $global:DashboardConfig.Resources.Timers.Remove($key)
    }

    if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId)) {
        $form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
        if ($form -and $form.Tag) {
            if ($form.Tag.ToolTipFtool) {
                $form.Tag.ToolTipFtool.Dispose()
            }
            
            if ($form.Tag.BtnHotkeyToggle -and $form.Tag.BtnHotkeyToggle.Checked) {
                try {
                    Write-Verbose "MACRO: Pausing hotkeys for instance $instanceId before cleanup (global toggle was active)"
                    PauseHotkeysForOwner -OwnerKey "macroinst_$instanceId"
                }
                catch {
                    Write-Warning "MACRO: Failed to pause hotkeys for instance $instanceId during cleanup. Error: $_"
                }
            }
            
            if ($form.Tag.ActiveJobs) {
                foreach ($job in $form.Tag.ActiveJobs) {
                    try {
                        if ($job.PowerShell) {
                            if ($job.PowerShell.InvocationStateInfo.State -eq 'Running') {
                                $job.PowerShell.Stop()
                            }
                            $job.PowerShell.Dispose()
                        }
                    }
                    catch {
                        Write-Verbose "Error disposing Macro Job: $_"
                    }
                }
                $form.Tag.ActiveJobs.Clear()
            }
            if ($form.Tag.RunspacePool) {
                try {
                    if ($form.Tag.RunspacePool.RunspacePoolStateInfo.State -ne 'Closed') {
                        $form.Tag.RunspacePool.Close()
                    }
                    $form.Tag.RunspacePool.Dispose()
                }
                catch {
                    Write-Verbose "Error disposing Macro Runspace: $_"
                }
                $form.Tag.RunspacePool = $null
            }

            if ($form.Tag.HotkeyId) {
                try { UnregisterHotkeyInstance -Id $form.Tag.HotkeyId -OwnerKey "macroinst_$($form.Tag.InstanceId)" } catch {}
            }
            if ($form.Tag.GlobalHotkeyId) { try { $globalOwnerKey = "global_toggle_$($form.Tag.InstanceId)"; UnregisterHotkeyInstance -Id $form.Tag.GlobalHotkeyId -OwnerKey $globalOwnerKey } catch {} }
        }
    }

    if (Get-Command RefreshHotkeysList -ErrorAction SilentlyContinue) { try { RefreshHotkeysList } catch {} }
}

function LoadMacroSettings {
    param($formData)
    $formData.IsLoading = $true
    
    if (-not $global:DashboardConfig.Config.Contains('Macro')) { $formData.IsLoading = $false; return }
    
    $profilePrefix = $null
    if (Get-Command FindOrCreateProfile -ErrorAction SilentlyContinue) {
        $profilePrefix = FindOrCreateProfile $formData.WindowTitle
    }
    
    if ($profilePrefix) {
        $configSuffix = ""
        $parts = $formData.InstanceId -split '_'
        if ($parts.Count -gt 2) {
            $configSuffix = "_Sub" + $parts[-1]
        }

        $cfg = $global:DashboardConfig.Config['Macro']
        
        $p = "$profilePrefix$configSuffix"

        if ($cfg.Contains("Key_$p")) { $formData.BtnKeySelect.Text = $cfg["Key_$p"] }
        if ($cfg.Contains("Interval_$p")) { $formData.Interval.Text = $cfg["Interval_$p"] }
        if ($cfg.Contains("Name_$p")) { $formData.Name.Text = $cfg["Name_$p"] }
        
        $globalHotkeyName = "GlobalHotkey_$p"
        if ($cfg.Contains("LoopEnabled_$p")) { try { $formData.BtnLoopToggle.Checked = [bool]::Parse($cfg["LoopEnabled_$p"]) } catch {} }
        if ($cfg.Contains("HoldEnabled_$p")) { try { $formData.BtnHoldKeyToggle.Checked = [bool]::Parse($cfg["HoldEnabled_$p"]) } catch {} }
        if ($cfg.Contains("HoldInterval_$p")) { $formData.TxtHoldKeyInterval.Text = $cfg["HoldInterval_$p"] }
        if ($cfg.Contains("WaitEnabled_$p")) { try { $formData.BtnWaitToggle.Checked = [bool]::Parse($cfg["WaitEnabled_$p"]) } catch {} }

        if ($cfg.Contains($globalHotkeyName)) { $formData.GlobalHotkey = $cfg[$globalHotkeyName] }

        $hotkeysEnabledName = "hotkeys_enabled_$p"
        if ($cfg.Contains($hotkeysEnabledName)) {
            $value = $cfg[$hotkeysEnabledName]
            try { $formData.BtnHotkeyToggle.Checked = [bool]::Parse($value) } catch { $formData.BtnHotkeyToggle.Checked = $true }
        }
        else { $formData.BtnHotkeyToggle.Checked = $true }

        if (-not $global:DashboardConfig.Resources.InstanceHotkeysPaused) { $global:DashboardConfig.Resources.InstanceHotkeysPaused = @{} }
        $global:DashboardConfig.Resources.InstanceHotkeysPaused[$formData.InstanceId] = (-not $formData.BtnHotkeyToggle.Checked)

        if ($cfg.Contains("PosX_$p")) {
            $positionValue = [int]$cfg["PosX_$p"]
            $min = $formData.PositionSliderX.Minimum
            $max = $formData.PositionSliderX.Maximum
            if ($min -gt $max) { $tmp = $min; $min = $max; $max = $tmp }
            $formData.PositionSliderX.Value = [Math]::Max($min, [Math]::Min($max, $positionValue))
        }
        if ($cfg.Contains("PosY_$p")) {
            $positionValue = [int]$cfg["PosY_$p"]
            $min = $formData.PositionSliderY.Minimum
            $max = $formData.PositionSliderY.Maximum
            if ($min -gt $max) { $tmp = $min; $min = $max; $max = $tmp }
            $formData.PositionSliderY.Value = [Math]::Max($min, [Math]::Min($max, $positionValue))
        }

        if ($cfg.Contains("collapsed_$p")) {
            $formData.IsCollapsed = [bool]::Parse($cfg["collapsed_$p"])
        }
        else {
            $formData.IsCollapsed = $false
        }

        if ($cfg.Contains("origHeight_$p")) {
            $formData.OriginalHeight = [int]$cfg["origHeight_$p"]
        }
        else {
            $formData.OriginalHeight = $formData.Form.Height
        }

        if ($cfg.Contains("Hotkey_$p")) {
            $hk = $cfg["Hotkey_$p"]
            $formData.Hotkey = $hk
            $formData.BtnHotKey.Text = $hk
            
            try {
                $scriptBlock = [scriptblock]::Create("ToggleMacroInstance -InstanceId '$($formData.InstanceId)'")
                $ownerLabel = $null
                if ($formData.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($formData.Name)) { $ownerLabel = $formData.Name }
                elseif ($formData.PSObject.Properties['WindowTitle'] -and -not [string]::IsNullOrEmpty($formData.WindowTitle)) { $ownerLabel = $formData.WindowTitle }
                else { $ownerLabel = "Macro: $($formData.InstanceId)" }
                $ownerKey = "macroinst_$($formData.InstanceId)"
                $formData.HotkeyId = SetHotkey -KeyCombinationString $hk -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel
            }
            catch {}
        }

        if ($cfg.Contains("SeqCount_$p")) {
            $count = [int]$cfg["SeqCount_$p"]
            for ($i = 1; $i -le $count; $i++) {
                $seqData = CreateSequencePanel $formData.Form
                if ($cfg.Contains("Seq${i}_Key_$p")) { $seqData.BtnKeySelect.Text = $cfg["Seq${i}_Key_$p"] }
                if ($cfg.Contains("Seq${i}_Interval_$p")) { $seqData.Interval.Text = $cfg["Seq${i}_Interval_$p"] }
                if ($cfg.Contains("Seq${i}_Name_$p")) { $seqData.Name.Text = $cfg["Seq${i}_Name_$p"] }
                if ($cfg.Contains("Seq${i}_HoldEnabled_$p")) { try { $seqData.BtnHoldKeyToggle.Checked = [bool]::Parse($cfg["Seq${i}_HoldEnabled_$p"]) } catch {} }
                if ($cfg.Contains("Seq${i}_HoldInterval_$p")) { $seqData.TxtHoldKeyInterval.Text = $cfg["Seq${i}_HoldInterval_$p"] }
                if ($cfg.Contains("Seq${i}_WaitEnabled_$p")) { try { $seqData.BtnWaitToggle.Checked = [bool]::Parse($cfg["Seq${i}_WaitEnabled_$p"]) } catch {} }
            }
            RepositionSequences $formData.Form
        }
    }
    else {
        $formData.Hotkey = $null
        $formData.BtnHotKey.Text = 'Hotkey'
        $formData.BtnHotkeyToggle.Checked = $true
    }
    $formData.IsLoading = $false
}

function UpdateMacroSettings {
    param($formData, [switch]$forceWrite)
    
    if ($formData.IsLoading) { return }
    
    if (-not $global:DashboardConfig.Config.Contains('Macro')) { 
        $global:DashboardConfig.Config['Macro'] = [ordered]@{} 
    }
    
    $profilePrefix = $null
    if (Get-Command FindOrCreateProfile -ErrorAction SilentlyContinue) {
        $profilePrefix = FindOrCreateProfile $formData.WindowTitle
    }

    if ($profilePrefix) {
        $configSuffix = ""
        $parts = $formData.InstanceId -split '_'
        if ($parts.Count -gt 2) {
            $configSuffix = "_Sub" + $parts[-1]
        }

        $cfg = $global:DashboardConfig.Config['Macro']
        
        $p = "$profilePrefix$configSuffix"
        
        $cfg["Key_$p"] = $formData.BtnKeySelect.Text
        $cfg["Interval_$p"] = $formData.Interval.Text
        $cfg["Name_$p"] = $formData.Name.Text
        $cfg["PosX_$p"] = $formData.PositionSliderX.Value
        $cfg["PosY_$p"] = $formData.PositionSliderY.Value
        $cfg["LoopEnabled_$p"] = $formData.BtnLoopToggle.Checked
        $cfg["HoldEnabled_$p"] = $formData.BtnHoldKeyToggle.Checked
        $cfg["HoldInterval_$p"] = $formData.TxtHoldKeyInterval.Text
        $cfg["WaitEnabled_$p"] = $formData.BtnWaitToggle.Checked
        $cfg["GlobalHotkey_$p"] = $formData.GlobalHotkey
        $cfg["hotkeys_enabled_$p"] = $formData.BtnHotkeyToggle.Checked
        $cfg["collapsed_$p"] = $formData.IsCollapsed
        $cfg["origHeight_$p"] = $formData.OriginalHeight

        if ($formData.Hotkey) { $cfg["Hotkey_$p"] = $formData.Hotkey } else { if ($cfg.Contains("Hotkey_$p")) { $cfg.Remove("Hotkey_$p") } }

        $seqs = $formData.SequencePanels
        $cfg["SeqCount_$p"] = $seqs.Count
        
        for ($i = 0; $i -lt $seqs.Count; $i++) {
            $n = $i + 1
            $sData = $seqs[$i].Tag
            $cfg["Seq${n}_Key_$p"] = $sData.BtnKeySelect.Text
            $cfg["Seq${n}_Interval_$p"] = $sData.Interval.Text
            $cfg["Seq${n}_Name_$p"] = $sData.Name.Text
            $cfg["Seq${n}_HoldEnabled_$p"] = $sData.BtnHoldKeyToggle.Checked
            $cfg["Seq${n}_HoldInterval_$p"] = $sData.TxtHoldKeyInterval.Text
            $cfg["Seq${n}_WaitEnabled_$p"] = $sData.BtnWaitToggle.Checked
        }
        
        $k = $seqs.Count + 1
        while ($cfg.Contains("Seq${k}_Key_$p")) {
            $cfg.Remove("Seq${k}_Key_$p")
            $cfg.Remove("Seq${k}_Interval_$p")
            $cfg.Remove("Seq${k}_Name_$p")
            $cfg.Remove("Seq${k}_HoldEnabled_$p")
            $cfg.Remove("Seq${k}_HoldInterval_$p")
            $cfg.Remove("Seq${k}_WaitEnabled_$p")
            $k++
        }

        if ($forceWrite) {
            if (Get-Command WriteConfig -ErrorAction SilentlyContinue) { WriteConfig }
        }
    }
}

function Import-MgpMacro {
    param(
        $formData,
        [string]$filePath
    )

    if (-not (Test-Path $filePath)) { return }

    $lines = Get-Content $filePath -ErrorAction SilentlyContinue
    if (-not $lines) { return }

    $sections = [ordered]@{}
    $currentSection = $null

    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith(';')) { continue }

        if ($trimmed -match '^\[(.*)\]$') {
            $currentSection = $matches[1].Trim()
            $sections[$currentSection] = [System.Collections.ArrayList]::new()
        }
        elseif ($null -ne $currentSection) {
            $sections[$currentSection].Add($trimmed) | Out-Null
        }
    }

    if ($sections.Count -eq 0) { return }

    $sectionKeys = @($sections.Keys)
    for ($secIdx = 0; $secIdx -lt $sectionKeys.Count; $secIdx++) {
        $secName = $sectionKeys[$secIdx]
        $secLines = $sections[$secName]

        $targetData = $null
        if ($secIdx -eq 0) {
            $targetData = $formData
        }
        else {
            $row = $formData.Row
            if ($row -and $row.Tag) {
                $rowIdentity = Get-RowIdentity -Row $row
                $safeIdentity = $rowIdentity -replace '[^a-zA-Z0-9_\-]', '_'
                $baseInstanceId = "Macro_$safeIdentity"

                $nextIndex = 1
                while ($global:DashboardConfig.Resources.FtoolForms.Contains("${baseInstanceId}_${nextIndex}") -and -not $global:DashboardConfig.Resources.FtoolForms["${baseInstanceId}_${nextIndex}"].IsDisposed) {
                    $nextIndex++
                }
                $newInstanceId = "${baseInstanceId}_${nextIndex}"
                $newForm = CreateMacroForm -instanceId $newInstanceId -targetWindowRect $null -windowTitle $formData.WindowTitle -row $row
                if ($newForm) {
                    $newForm.Left = $formData.Form.Right + 10 + (($secIdx - 1) * 20)
                    $newForm.Top = $formData.Form.Top
                    $global:DashboardConfig.Resources.FtoolForms[$newInstanceId] = $newForm
                    $newForm.Show()
                    $newForm.BringToFront()
                    $targetData = $newForm.Tag
                }
            }
        }

        if (-not $targetData) { continue }

        if ($targetData.RunningTimer) { StopMacroSequence $targetData }

        $bindedKey = $null
        $repeatVal = $false
        $parsedSteps = [System.Collections.ArrayList]::new()

        $currentStep = $null
        $expectingInterval = $false

        foreach ($l in $secLines) {
            if ($l -match '(?i)^bindedkey=\s*\{?([^}]+)\}?') {
                $bindedKey = $matches[1].Trim()
            }
            elseif ($l -match '(?i)^loop=\s*(1|true)') {
                $repeatVal = $true
            }
            elseif ($l -match '(?i)^repeat=\s*(\d+)') {
                $repeatVal = ([int]$matches[1] -gt 1)
            }
            elseif ($l -match '(?i)^e=\s*\{([^}]+)\s+down\}') {
                $kName = $matches[1].Trim()
                $currentStep = [PSCustomObject]@{
                    Key          = $kName
                    HoldEnabled  = $true
                    HoldInterval = '50'
                    Interval     = '1000'
                    WaitEnabled  = $true
                    Name         = "Step $($parsedSteps.Count + 1)"
                }
                $expectingInterval = $false
            }
            elseif ($l -match '(?i)^e=\s*\{([^}]+)\s+up\}') {
                $expectingInterval = $true
            }
            elseif ($l -match '(?i)^e=\s*(\d+)\s+D') {
                $ms = $matches[1]
                if ($null -ne $currentStep) {
                    if ($expectingInterval) {
                        $currentStep.Interval = $ms
                        $parsedSteps.Add($currentStep) | Out-Null
                        $currentStep = $null
                        $expectingInterval = $false
                    }
                    else {
                        $currentStep.HoldInterval = $ms
                        if ([int]$ms -gt 0) {
                            $currentStep.HoldEnabled = $true
                        }
                    }
                }
            }
            elseif ($l -match '(?i)^e=wait\s*(1|true)') {
                if ($null -ne $currentStep) { $currentStep.WaitEnabled = $true }
            }
            elseif ($l -match '(?i)^e=hold\s*(1|true)') {
                if ($null -ne $currentStep) { $currentStep.HoldEnabled = $true }
            }
        }

        if ($null -ne $currentStep) {
            $parsedSteps.Add($currentStep) | Out-Null
        }

        $targetData.Name.Text = $secName
        if ($bindedKey) {
            $targetData.BtnHotKey.Text = $bindedKey
            $targetData.Hotkey = $bindedKey
            try {
                $scriptBlock = [scriptblock]::Create("ToggleMacroInstance -InstanceId '$($targetData.InstanceId)'")
                $ownerKey = "macroinst_$($targetData.InstanceId)"
                $ownerLabel = if ($targetData.Name.Text) { $targetData.Name.Text } else { "Macro: $($targetData.InstanceId)" }
                $targetData.HotkeyId = SetHotkey -KeyCombinationString $bindedKey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel -OldHotkeyId $targetData.HotkeyId
            } catch {}
        }

        $targetData.BtnLoopToggle.Checked = $repeatVal

        while ($targetData.SequencePanels.Count -gt 0) {
            $p = $targetData.SequencePanels[0]
            $targetData.Form.Controls.Remove($p)
            $targetData.SequencePanels.RemoveAt(0)
            $p.Dispose()
        }

        if ($parsedSteps.Count -gt 0) {
            $first = $parsedSteps[0]
            $targetData.BtnKeySelect.Text = $first.Key
            $targetData.Interval.Text = $first.Interval
            $targetData.BtnHoldKeyToggle.Checked = $first.HoldEnabled
            $targetData.TxtHoldKeyInterval.Text = $first.HoldInterval
            $targetData.BtnWaitToggle.Checked = $first.WaitEnabled

            for ($s = 1; $s -lt $parsedSteps.Count; $s++) {
                $st = $parsedSteps[$s]
                $seqData = CreateSequencePanel $targetData.Form
                $seqData.BtnKeySelect.Text = $st.Key
                $seqData.Interval.Text = $st.Interval
                $seqData.Name.Text = $st.Name
                $seqData.BtnHoldKeyToggle.Checked = $st.HoldEnabled
                $seqData.TxtHoldKeyInterval.Text = $st.HoldInterval
                $seqData.BtnWaitToggle.Checked = $st.WaitEnabled
            }
        }
        else {
            $targetData.BtnHoldKeyToggle.Checked = $false
            $targetData.BtnWaitToggle.Checked = $false
        }

        RepositionSequences $targetData.Form
        UpdateMacroSettings $targetData -forceWrite
    }
}

#endregion

#region Core Functions

function MacroSelectedRow {
    param($row, [switch]$ForceNew, [Parameter(Mandatory = $false)]$Config, [int]$InstanceIndex = 0)
    
    if (-not $row) {
        Write-Verbose "MACRO: Invalid row data passed to MacroSelectedRow."
        return
    }
    
    $rowIdentity = Get-RowIdentity -Row $row
    $safeIdentity = $rowIdentity -replace '[^a-zA-Z0-9_\-]', '_'
    $instanceId = "Macro_$safeIdentity"
    
    if ($ForceNew) {
        if ($InstanceIndex -gt 0) {
            $instanceId = "${instanceId}_${InstanceIndex}"
        }
        else {
            $index = 1
            while ($global:DashboardConfig.Resources.FtoolForms.Contains("${instanceId}_${index}")) {
                $index++
            }
            $instanceId = "${instanceId}_${index}"
        }
        Write-Verbose "Macro Debug: Generated stable unique instance ID $instanceId for restoration."
    }

    if (-not $row.Tag -or -not $row.Tag.PSObject.Properties['MainWindowHandle'] -or -not $row.Tag.MainWindowHandle) {
        Write-Verbose "MACRO: Row is missing a valid MainWindowHandle."
        return
    }
    
    $windowHandle = $row.Tag.MainWindowHandle
    
    if (-not $ForceNew -and $global:DashboardConfig.Resources.FtoolForms.Contains($instanceId)) {
        $existingForm = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
        if ($existingForm -and -not $existingForm.IsDisposed) {
            try {
                if ($existingForm.InvokeRequired) {
                    $existingForm.Invoke([System.Action] { $existingForm.BringToFront() })
                }
                else {
                    $existingForm.BringToFront()
                }
            }
            catch {}
            return
        }
        else {
            $global:DashboardConfig.Resources.FtoolForms.Remove($instanceId)
        }
    }
    
    $targetWindowRect = New-Object Custom.Native+RECT
    [Custom.Native]::GetWindowRect($windowHandle, [ref]$targetWindowRect)
    
    $windowTitle = "Macro: $instanceId"
    if ($row.Tag.PSObject.Properties['MainWindowTitle'] -and -not [string]::IsNullOrEmpty($row.Tag.MainWindowTitle)) {
        $windowTitle = $row.Tag.MainWindowTitle
    }
    elseif ($row.Cells.Count -gt 1 -and $row.Cells[1].Value) {
        $windowTitle = $row.Cells[1].Value.ToString()
    }
    
    $macroForm = CreateMacroForm $instanceId $targetWindowRect $windowTitle $row -Config $Config
    
    if ($macroForm) {
        $global:DashboardConfig.Resources.FtoolForms[$instanceId] = $macroForm
        $macroForm.Show()
        $macroForm.BringToFront()
    }
}

function StopMacroForm {
    param($Form)
    
    if (-not $Form -or $Form.IsDisposed) { return }
    
    try {
        if ($Form.Tag -and $Form.Tag.InstanceId) {
            $instanceId = $Form.Tag.InstanceId
            CleanupMacroResources $instanceId
            if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId)) {
                $global:DashboardConfig.Resources.FtoolForms.Remove($instanceId)
            }
        }
        $Form.Close()
    }
    catch {}
}

function CreateMacroForm {
    param($instanceId, $targetWindowRect, $windowTitle, $row, [Parameter(Mandatory = $false)]$Config)
    
    if (-not $row -or -not $row.Cells -or $row.Cells.Count -lt 2) {
        Write-Verbose "MACRO: Cannot create form due to invalid row data."
        return $null
    }
    
    if ($null -ne $Config -and $Config -is [string]) {
        try {
            $Config = $Config | ConvertFrom-Json
        }
        catch {
            Write-Verbose "MACRO: Failed to parse Config string as JSON. Ignoring."
            $Config = $null
        }
    }

    $macroForm = New-Object Custom.FtoolFormWindow
    $macroForm.Width = 235
    $macroForm.Height = 145 
    $macroForm.Top = if ($targetWindowRect) { ($targetWindowRect.Top + 30) } else { 200 }
    $macroForm.Left = if ($targetWindowRect) { ($targetWindowRect.Left + 10) } else { 200 }
    $macroForm.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $macroForm.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $macroForm.Text = "Macro - $instanceId"
    $macroForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $macroForm.Opacity = 1.0 

    if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon)) {
        try { $macroForm.Icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon) } catch {}
    }

    $macroToolTip = New-Object System.Windows.Forms.ToolTip
    $macroToolTip.AutoPopDelay = 5000
    $macroToolTip.InitialDelay = 100
    $macroToolTip.ReshowDelay = 10
    $macroToolTip.ShowAlways = $true
    $macroToolTip.OwnerDraw = $true
    $macroToolTip | Add-Member -MemberType NoteProperty -Name 'TipFont' -Value (New-Object System.Drawing.Font('Segoe UI', 9))
    $macroToolTip.Add_Draw({
            $g, $b, $c = $_.Graphics, $_.Bounds, [System.Drawing.Color]
            $g.FillRectangle((New-Object System.Drawing.SolidBrush $c::FromArgb(30, 30, 30)), $b)
            $g.DrawRectangle((New-Object System.Drawing.Pen $c::FromArgb(100, 100, 100)), $b.X, $b.Y, $b.Width - 1, $b.Height - 1)
            $g.DrawString($_.ToolTipText, $this.TipFont, (New-Object System.Drawing.SolidBrush $c::FromArgb(240, 240, 240)), 3, 3, [System.Drawing.StringFormat]::GenericTypographic)
        })
    $macroToolTip.Add_Popup({
            $g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
            $s = $g.MeasureString($this.GetToolTip($_.AssociatedControl), $this.TipFont, [System.Drawing.PointF]::new(0, 0), [System.Drawing.StringFormat]::GenericTypographic)
            $g.Dispose(); $_.ToolTipSize = [System.Drawing.Size]::new($s.Width + 12, $s.Height + 8)
        })

    $previousToolTip = $global:DashboardConfig.UI.ToolTipFtool
    $global:DashboardConfig.UI.ToolTipFtool = $macroToolTip

    try {

        $headerPanel = SetUIElement -type 'Panel' -visible $true -width 250 -height 20 -top 0 -left 0 -bg @(40, 40, 40)
        $macroForm.Controls.Add($headerPanel)

        $labelText = if ($row.Cells[1].Value) { $row.Cells[1].Value.ToString() } else { "" }
        $labelWinTitle = SetUIElement -type 'Label' -visible $true -width 90 -height 20 -top 5 -left 5 -bg @(40, 40, 40, 0) -fg @(255, 255, 255) -text $labelText -font (New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Regular))
        $headerPanel.Controls.Add($labelWinTitle)

        $btnInstanceHotkeyToggle = SetUIElement -type 'Label' -visible $true -width 15 -height 15 -top 2 -left 90 -bg @(40, 40, 40) -fg @(255, 255, 255) -text ([char]0x2328) -font (New-Object System.Drawing.Font('Segoe UI', 10)) -tooltip "Set Master Hotkey`nAssign a global hotkey to toggle all hotkeys for this instance."
        $headerPanel.Controls.Add($btnInstanceHotkeyToggle)

        $btnHotkeyToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 15 -top 3 -left 105 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true -tooltip "Toggle Hotkeys`nEnable or disable all hotkeys for this specific macro instance."
        $headerPanel.Controls.Add($btnHotkeyToggle)

        $btnImport = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 138 -bg @(40, 80, 80) -fg @(255, 255, 255) -text ([char]0x2193) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Import MGP`nSelect a .mgp file to import macro sequences."
        $headerPanel.Controls.Add($btnImport)

        $btnAddInstance = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 153 -bg @(40, 40, 40) -fg @(255, 255, 255) -text ([char]0x2398) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Add Instance`nCreates another macro window for this specific client."
        $headerPanel.Controls.Add($btnAddInstance)
        
        $btnAdd = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 168 -bg @(40, 80, 80) -fg @(255, 255, 255) -text ([char]0x2795) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Add Sequence`nAdds a new key sequence step to the macro list."
        $headerPanel.Controls.Add($btnAdd)
    
        $btnShowHide = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 183 -bg @(60, 60, 100) -fg @(255, 255, 255) -text ([char]0x25B2) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Minimize/Expand`nCollapse or expand the macro window to save screen space."
        $headerPanel.Controls.Add($btnShowHide)
        
        $btnReset = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 198 -bg @(100, 100, 100) -fg @(255, 255, 255) -text ([char]0x21BB) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Reset`nReset all settings for this instance to default."
        $headerPanel.Controls.Add($btnReset)

        $btnClose = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 215 -bg @(150, 20, 20) -fg @(255, 255, 255) -text ([char]0x166D) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Close`nStops the macro and closes this window."
        $headerPanel.Controls.Add($btnClose)

        $panelSettings = SetUIElement -type 'Panel' -visible $true -width 190 -height 75 -top 60 -left 40 -bg @(50, 50, 50)
        $macroForm.Controls.Add($panelSettings)
    
        $btnKeySelect = SetUIElement -type 'Button' -visible $true -width 50 -height 24 -top 3 -left 3 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Bind Key`nClick here, then press a key to assign it to this action."
        $panelSettings.Controls.Add($btnKeySelect)
    
        $interval = SetUIElement -type 'TextBox' -visible $true -width 40 -height 15 -top 4 -left 56 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '1000' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Interval (ms)`nTime in milliseconds to wait before repeating this action."
        $panelSettings.Controls.Add($interval)
    
        $name = SetUIElement -type 'TextBox' -visible $true -width 40 -height 17 -top 4 -left 99 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'Main' -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular)) -tooltip "Name`nGive this sequence a descriptive name for easier identification."
        $panelSettings.Controls.Add($name)

        $btnHotKey = SetUIElement -type 'Button' -visible $true -width 40 -height 20 -top 3 -left 145 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'Hotkey' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 6)) -tooltip "Global Hotkey`nAssign a global hotkey to Start/Stop this specific macro sequence."
        $panelSettings.Controls.Add($btnHotKey)

        $lblLoop = SetUIElement -type 'Label' -visible $true -width 35 -height 10 -top 28 -left 45 -bg @(50, 50, 50) -fg @(180, 180, 180) -text 'Loop' -font (New-Object System.Drawing.Font('Segoe UI', 6))
        $panelSettings.Controls.Add($lblLoop)

        $lblWait = SetUIElement -type 'Label' -visible $true -width 35 -height 10 -top 28 -left 80 -bg @(50, 50, 50) -fg @(180, 180, 180) -text 'Wait' -font (New-Object System.Drawing.Font('Segoe UI', 6))
        $panelSettings.Controls.Add($lblWait)

        $lblHold = SetUIElement -type 'Label' -visible $true -width 35 -height 10 -top 28 -left 115 -bg @(50, 50, 50) -fg @(180, 180, 180) -text 'Hold' -font (New-Object System.Drawing.Font('Segoe UI', 6))
        $panelSettings.Controls.Add($lblHold)

        $btnStart = SetUIElement -type 'Button' -visible $true -width 40 -height 20 -top 40 -left 3 -bg @(0, 120, 215) -fg @(255, 255, 255) -text 'Start' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Start`nBegin executing the macro sequence."
        $panelSettings.Controls.Add($btnStart)

        $btnStop = SetUIElement -type 'Button' -visible $true -width 40 -height 20 -top 40 -left 3 -bg @(200, 50, 50) -fg @(255, 255, 255) -text 'Stop' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Stop`nStop the currently running macro sequence."
        $btnStop.Enabled = $false
        $btnStop.Visible = $false
        $panelSettings.Controls.Add($btnStop)

        $btnLoopToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 16 -top 42 -left 45 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true -tooltip "Loop`nToggle looping. If enabled, the macro repeats until stopped.`nLoops below 100ms can block the Mouse UI, stop with hotkey instead."
        $panelSettings.Controls.Add($btnLoopToggle)

        $btnWaitToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 16 -top 42 -left 80 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true -tooltip "Wait Mode`nIf enabled, the macro waits for the key hold duration to finish before proceeding."
        $panelSettings.Controls.Add($btnWaitToggle)

        $btnHoldKeyToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 16 -top 42 -left 115 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $false -tooltip "Hold Key`nEnable to hold the key down for a specific duration instead of a single press."
        $panelSettings.Controls.Add($btnHoldKeyToggle)

        $txtHoldKeyInterval = SetUIElement -type 'TextBox' -visible $true -width 30 -height 15 -top 42 -left 150 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '50' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Hold Duration (ms)`nHow long the key should be held down in milliseconds."
        $panelSettings.Controls.Add($txtHoldKeyInterval)
    

        try {
            if ($targetWindowRect -and $targetWindowRect -is [Custom.Native+RECT]) {
                $rectForScreen = [System.Drawing.Rectangle]::FromLTRB($targetWindowRect.Left, $targetWindowRect.Top, $targetWindowRect.Right, $targetWindowRect.Bottom)
                $screen = [System.Windows.Forms.Screen]::FromRectangle($rectForScreen)
            }
            elseif ($row -and $row.Tag -and $row.Tag.MainWindowHandle) {
                $screen = [System.Windows.Forms.Screen]::FromHandle([IntPtr]$row.Tag.MainWindowHandle)
            }
            else {
                $screen = [System.Windows.Forms.Screen]::PrimaryScreen
            }
        }
        catch { $screen = [System.Windows.Forms.Screen]::PrimaryScreen }
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
        $initialY = [Math]::Max($minY, [Math]::Min($maxY, $macroForm.Top))
        $positionSliderY.Value = $initialY
        $positionSliderY.Size = New-Object System.Drawing.Size(15, 110)
        $positionSliderY.Location = New-Object System.Drawing.Point(5, 20)
        $macroForm.Controls.Add($positionSliderY)
        
        $positionSliderX = New-Object System.Windows.Forms.TrackBar
        $positionSliderX.Minimum = $minX
        $positionSliderX.Maximum = $maxX
        $positionSliderX.TickFrequency = 300
        $initialX = [Math]::Max($minX, [Math]::Min($maxX, $macroForm.Left))
        $positionSliderX.Value = $initialX
        $positionSliderX.Size = New-Object System.Drawing.Size(190, 15)
        $positionSliderX.Location = New-Object System.Drawing.Point(45, 25)
        $macroForm.Controls.Add($positionSliderX)

    }
    finally {
        $global:DashboardConfig.UI.ToolTipFtool = $previousToolTip
    }

    $formData = [PSCustomObject]@{
        Type                    = 'Macro'
        InstanceId              = $instanceId
        HeaderPanel             = $headerPanel
        LabelWinTitle           = $labelWinTitle
		OriginalTitle 			= $labelText
        SelectedWindow          = $row.Tag.MainWindowHandle
        BtnKeySelect            = $btnKeySelect
        Interval                = $interval
        Name                    = $name
        BtnStart                = $btnStart
        BtnStop                 = $btnStop
        BtnLoopToggle           = $btnLoopToggle
        BtnWaitToggle           = $btnWaitToggle
        BtnHoldKeyToggle        = $btnHoldKeyToggle
        TxtHoldKeyInterval      = $txtHoldKeyInterval
        BtnHotKey               = $btnHotKey
        BtnInstanceHotkeyToggle = $btnInstanceHotkeyToggle
        BtnHotkeyToggle         = $btnHotkeyToggle
        BtnImport               = $btnImport
        BtnAdd                  = $btnAdd
        BtnAddInstance          = $btnAddInstance
        BtnClose                = $btnClose
        BtnShowHide             = $btnShowHide
        BtnReset                = $btnReset
        PositionSliderX         = $positionSliderX
        PositionSliderY         = $positionSliderY
        Form                    = $macroForm
        DebugLabel              = $labelDebug
        RunningTimer            = $null
        WindowTitle             = $windowTitle
        Row                     = $row
        ToolTipFtool            = $macroToolTip
        IsCollapsed             = $false
        OriginalHeight          = 145
        HotkeyId                = $null
        GlobalHotkeyId          = $null
        GlobalHotkey            = $null
        Hotkey                  = $null
        SequencePanels          = [System.Collections.ArrayList]::new()
        ActiveSteps             = @()
        CurrentStepIndex        = 0
        ActiveJobs              = [System.Collections.ArrayList]::new()
        RunspacePool            = (InitMacroRunspace -InstanceId $instanceId)
        IsLoading               = $false
    }
    
    $macroForm.Tag = $formData
    
    $global:DashboardConfig.Resources.FtoolForms[$instanceId] = $macroForm

    LoadMacroSettings $formData | Out-Null
    
    CreateMacroPositionTimer $formData | Out-Null
    ToggleMacroHotkeys -InstanceId $formData.InstanceId -ToggleState $formData.BtnHotkeyToggle.Checked | Out-Null

    if (-not [string]::IsNullOrEmpty($formData.GlobalHotkey)) {
        try {
            $ownerKey = "global_toggle_$($formData.InstanceId)"
            $script = @"
if (`$global:DashboardConfig.Resources.FtoolForms.Contains('$($formData.InstanceId)')) {
    `$f = `$global:DashboardConfig.Resources.FtoolForms['$($formData.InstanceId)']
    if (`$f -and -not `$f.IsDisposed -and `$f.Tag) {
        `$toggle = `$f.Tag.BtnHotkeyToggle
        if (`$toggle) {
            if (`$toggle.InvokeRequired) { `$toggle.Invoke([System.Action]{ `$toggle.Checked = -not `$toggle.Checked }) } else { `$toggle.Checked = -not `$toggle.Checked }
            ToggleMacroHotkeys -InstanceId '$($formData.InstanceId)' -ToggleState `$toggle.Checked
        }
    }
}
"@
            $scriptBlock = [scriptblock]::Create($script)
            $ownerLabel = "Macro: $($formData.InstanceId)"
            $formData.GlobalHotkeyId = SetHotkey -KeyCombinationString $formData.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel -OldHotkeyId $null
        }
        catch {
            $formData.GlobalHotkeyId = $null
        }
    }

    AddMacroEventHandlers $formData | Out-Null
    AddMacroDragHandlers $formData | Out-Null

    if ($formData.IsCollapsed) {
        $macroForm.Height = 26
        $formData.BtnShowHide.Text = [char]0x25BC
        
        $namePrefix = if ([string]::IsNullOrWhiteSpace($name.Text)) { "" } else { "$($name.Text) - " }
        $labelWinTitle.Text = "$namePrefix$labelText"
    }
    else {
        $macroForm.Height = $formData.OriginalHeight
        $formData.BtnShowHide.Text = [char]0x25B2
        $labelWinTitle.Text = $labelText
    }
    
    return $macroForm
}

function CreateSequencePanel {
    param($form)
    $formData = $form.Tag

    $previousToolTip = $global:DashboardConfig.UI.ToolTipFtool
    if ($formData.ToolTipFtool) {
        $global:DashboardConfig.UI.ToolTipFtool = $formData.ToolTipFtool
    }

    try {
    
        $panelSeq = SetUIElement -type 'Panel' -visible $true -width 190 -height 65 -top 0 -left 40 -bg @(50, 50, 50)
        $form.Controls.Add($panelSeq)
        $panelSeq.BringToFront()

        $btnKeySelect = SetUIElement -type 'Button' -visible $true -width 50 -height 24 -top 3 -left 3 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -tooltip "Bind Key`nClick here, then press a key to assign it to this action."
        $panelSeq.Controls.Add($btnKeySelect)

        $interval = SetUIElement -type 'TextBox' -visible $true -width 40 -height 15 -top 4 -left 56 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '1000' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Interval (ms)`nTime in milliseconds to wait before repeating this action."
        $panelSeq.Controls.Add($interval)

        $name = SetUIElement -type 'TextBox' -visible $true -width 60 -height 17 -top 4 -left 99 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'Name' -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular)) -tooltip "Name`nGive this sequence a descriptive name for easier identification."
        $panelSeq.Controls.Add($name)

        $btnRemove = SetUIElement -type 'Button' -visible $true -width 22 -height 20 -top 3 -left 163 -bg @(150, 50, 50) -fg @(255, 255, 255) -text 'X' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)) -tooltip "Remove Step`nDelete this key sequence from the macro."
        $panelSeq.Controls.Add($btnRemove)

        $lblWait = SetUIElement -type 'Label' -visible $true -width 35 -height 10 -top 28 -left 3 -bg @(50, 50, 50) -fg @(180, 180, 180) -text 'Wait' -font (New-Object System.Drawing.Font('Segoe UI', 6))
        $panelSeq.Controls.Add($lblWait)

        $lblHold = SetUIElement -type 'Label' -visible $true -width 35 -height 10 -top 28 -left 40 -bg @(50, 50, 50) -fg @(180, 180, 180) -text 'Hold' -font (New-Object System.Drawing.Font('Segoe UI', 6))
        $panelSeq.Controls.Add($lblHold)

        $btnWaitToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 16 -top 40 -left 3 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true -tooltip "Wait Mode`nIf enabled, the macro waits for the key hold duration to finish before proceeding."
        $panelSeq.Controls.Add($btnWaitToggle)

        $btnHoldKeyToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 16 -top 40 -left 40 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $false -tooltip "Hold Key`nEnable to hold the key down for a specific duration instead of a single press."
        $panelSeq.Controls.Add($btnHoldKeyToggle)

        $txtHoldKeyInterval = SetUIElement -type 'TextBox' -visible $true -width 40 -height 15 -top 42 -left 75 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '50' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Hold Duration (ms)`nHow long the key should be held down in milliseconds."
        $panelSeq.Controls.Add($txtHoldKeyInterval)

    }
    finally {
        $global:DashboardConfig.UI.ToolTipFtool = $previousToolTip
    }

    $seqData = [PSCustomObject]@{
        Panel              = $panelSeq
        BtnKeySelect       = $btnKeySelect
        Interval           = $interval
        Name               = $name
        BtnWaitToggle      = $btnWaitToggle
        BtnHoldKeyToggle   = $btnHoldKeyToggle
        TxtHoldKeyInterval = $txtHoldKeyInterval
        BtnRemove          = $btnRemove
    }
    $panelSeq.Tag = $seqData
    
    $formData.SequencePanels.Add($panelSeq) | Out-Null

    $btnKeySelect.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $formData = $form.Tag

            $currentKey = $this.Text
            $newKey = Show-KeyCaptureDialog $currentKey -OwnerForm $form
            if ($newKey -and $newKey -ne $currentKey) {
                $this.Text = $newKey
                if ($formData.RunningTimer) { StopMacroSequence $formData }
                UpdateMacroSettings $formData -forceWrite
            }
        })

    $btnRemove.Add_Click({
            $btn = $this
            $form = $btn.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $formData = $form.Tag
        
            $panelToRemove = $btn.Parent 
            if ($panelToRemove) {
                $form.Controls.Remove($panelToRemove)
                $formData.SequencePanels.Remove($panelToRemove)
                $panelToRemove.Dispose()
            
                RepositionSequences $form
                if ($formData.RunningTimer) { StopMacroSequence $formData }
                UpdateMacroSettings $formData -forceWrite
            }
        })

    $btnWaitToggle.Add_Click({ UpdateMacroSettings $this.FindForm().Tag -forceWrite })
    $btnHoldKeyToggle.Add_Click({ UpdateMacroSettings $this.FindForm().Tag -forceWrite })
    $interval.Add_TextChanged({ UpdateMacroSettings $this.FindForm().Tag -forceWrite })
    $name.Add_TextChanged({ UpdateMacroSettings $this.FindForm().Tag -forceWrite })
    $txtHoldKeyInterval.Add_TextChanged({ UpdateMacroSettings $this.FindForm().Tag -forceWrite })

    return $seqData
}

function UpdateMacroSliderValuesFromFormPosition {
    param($formData, $windowRect)
	
    if (-not $formData -or -not $formData.PSObject.Properties['PositionSliderX'] -or -not $formData.PSObject.Properties['PositionSliderY']) {
        return
    }
	
    if (-not $windowRect) {
        return
    }
	
    try {
        $form = $formData.Form
        if (-not $form -or $form.IsDisposed) {
            return
        }
		
        $formLeft = $form.Left
        $formTop = $form.Top
        try {
            if ($windowRect -and $windowRect -is [Custom.Native+RECT]) {
                $rectForScreen = [System.Drawing.Rectangle]::FromLTRB($windowRect.Left, $windowRect.Top, $windowRect.Right, $windowRect.Bottom)
                $screen = [System.Windows.Forms.Screen]::FromRectangle($rectForScreen)
            }
            elseif ($formData.SelectedWindow -and $formData.SelectedWindow -ne [IntPtr]::Zero) {
                $screen = [System.Windows.Forms.Screen]::FromHandle([IntPtr]$formData.SelectedWindow)
            }
            else {
                $screen = [System.Windows.Forms.Screen]::PrimaryScreen
            }
        }
        catch { $screen = [System.Windows.Forms.Screen]::PrimaryScreen }
        $virtual = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $maxX = $virtual.Right
        $maxY = $virtual.Bottom

        $minX = $virtual.Left
        $sliderX = [Math]::Max($minX, [Math]::Min($maxX, $formLeft))
        $formData.PositionSliderX.Value = $sliderX

        $minY = $virtual.Top
        $sliderY = [Math]::Max($minY, [Math]::Min($maxY, $formTop))
        $formData.PositionSliderY.Value = $sliderY

    }
    catch {
        Write-Verbose ('MACRO: Error updating slider values from form position: {0}' -f $_.Exception.Message)
    }
}

function AddMacroDragHandlers {
    param($formData)
	
    if (-not $formData -or -not $formData.PSObject.Properties['Form']) {
        return
    }
	
    $form = $formData.Form
    if ($form.IsDisposed) {
        return
    }
	
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
                if (($currentTime - $f.Tag.LastUpdateTime).TotalMilliseconds -lt 20) {
                    return
                }
                $f.Tag.LastUpdateTime = $currentTime
				
                $cursorX = [System.Windows.Forms.Cursor]::Position.X
                $cursorY = [System.Windows.Forms.Cursor]::Position.Y
				
                $deltaX = $cursorX - $f.Tag.DragStart.X
                $deltaY = $cursorY - $f.Tag.DragStart.Y
				
                $newLeft = $f.Tag.DragStart.FormX + $deltaX
                $newTop = $f.Tag.DragStart.FormY + $deltaY
				
                $f.Left = $newLeft
                $f.Top = $newTop
				
                if ($f.Tag) {
                    $timerKey = "macroPosition_$($f.Tag.InstanceId)"
                    if ($global:DashboardConfig.Resources.Timers.Contains($timerKey)) {
                        $timer = $global:DashboardConfig.Resources.Timers[$timerKey]
                        if ($timer -and $timer.Tag) {
                            $timer.Tag.RealLeft = $newLeft
                            $timer.Tag.RealTop = $newTop
                        }
                    }
                }
				
                if ($f.Tag.SelectedWindow -and $f.Tag.SelectedWindow -ne [IntPtr]::Zero) {
                    $rect = New-Object Custom.Native+RECT
                    if ([Custom.Native]::GetWindowRect($f.Tag.SelectedWindow, [ref]$rect)) {
                        UpdateMacroSliderValuesFromFormPosition $f.Tag $rect
                    }
                }
            }
            catch {
                Write-Verbose ('MACRO: Error during mouse move: {0}' -f $_.Exception.Message)
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
                        UpdateMacroSliderValuesFromFormPosition $f.Tag $rect
                        UpdateMacroSettings $f.Tag -forceWrite
                    }
                }
                catch {}
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

    if ($formData.LabelWinTitle) {
        $formData.LabelWinTitle.Add_MouseDown($mouseDown)
        $formData.LabelWinTitle.Add_MouseMove($mouseMove)
        $formData.LabelWinTitle.Add_MouseUp($mouseUp)
    }
}

function AddMacroEventHandlers {
    param($formData)

    $formData.BtnKeySelect.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag

            $currentKey = $this.Text
            $newKey = Show-KeyCaptureDialog $currentKey -OwnerForm $form
            if ($newKey -and $newKey -ne $currentKey) {
                $this.Text = $newKey
                if ($data.RunningTimer) { StopMacroSequence $data }
                UpdateMacroSettings $data -forceWrite
            }
        })

    $formData.BtnStart.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            StartMacroSequence $form.Tag
        })

    $formData.BtnStop.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            StopMacroSequence $form.Tag
        })

    $formData.BtnAdd.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag

            if ($data.IsCollapsed) { return }
            if ($data.SequencePanels.Count -ge 15) { return }
        
            CreateSequencePanel $form
            RepositionSequences $form
            if ($data.RunningTimer) { StopMacroSequence $data }
            UpdateMacroSettings $data -forceWrite
        })

    $formData.BtnShowHide.Add_Click({
        $form = $this.FindForm()
        if (-not $form -or -not $form.Tag) { return }
        $data = $form.Tag

        if ($data.IsCollapsed) {
            $form.Height = $data.OriginalHeight
            $data.IsCollapsed = $false
            $this.Text = [char]0x25B2

            if ($data.LabelWinTitle) {
                $data.LabelWinTitle.Text = $data.OriginalTitle
            }
        }
        else {
            $data.OriginalHeight = $form.Height
            $form.Height = 26
            $data.IsCollapsed = $true
            $this.Text = [char]0x25BC

            if ($data.LabelWinTitle -and $data.Name) {
                $namePrefix = if ([string]::IsNullOrWhiteSpace($data.Name.Text)) { "" } else { "$($data.Name.Text) - " }
                $data.LabelWinTitle.Text = "$namePrefix$($data.OriginalTitle)"
            }
        }
    })

    $formData.BtnClose.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $null = $data; $data = $form.Tag

            StopMacroForm $form
        })

    $formData.BtnReset.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag
        
            if ((Show-DarkMessageBox "Are you sure you want to reset this Macro instance to default settings? This will remove all sequences and clear keys." "Reset Defaults" "YesNo" "Warning") -eq 'Yes') {
                if ($data.RunningTimer) {
                    StopMacroSequence $data
                }

                while ($data.SequencePanels.Count -gt 0) {
                    $p = $data.SequencePanels[0]
                    $form.Controls.Remove($p)
                    $data.SequencePanels.RemoveAt(0)
                    $p.Dispose()
                }

                if ($data.GlobalHotkeyId) {
                    $ownerKey = "global_toggle_$($data.InstanceId)"
                    try { UnregisterHotkeyInstance -Id $data.GlobalHotkeyId -OwnerKey $ownerKey } catch {}
                    $data.GlobalHotkeyId = $null
                    $data.GlobalHotkey = $null
                }

                if ($data.HotkeyId) {
                    try { UnregisterHotkeyInstance -Id $data.HotkeyId -OwnerKey "macroinst_$($data.InstanceId)" } catch {}
                    $data.HotkeyId = $null
                    $data.Hotkey = $null
                }

                $data.BtnKeySelect.Text = ""
                $data.Interval.Text = "1000"
                $data.Name.Text = "Main"
                $data.BtnHotKey.Text = "Hotkey"
                $data.BtnHotkeyToggle.Checked = $true
                $data.BtnLoopToggle.Checked = $true
                $data.BtnWaitToggle.Checked = $true
                $data.BtnHoldKeyToggle.Checked = $false
                $data.TxtHoldKeyInterval.Text = "50"
                $minX = $data.PositionSliderX.Minimum
                $maxX = $data.PositionSliderX.Maximum
                $minY = $data.PositionSliderY.Minimum
                $maxY = $data.PositionSliderY.Maximum
                $data.PositionSliderX.Value = [Math]::Max($minX, [Math]::Min($maxX, $data.Form.Left))
                $data.PositionSliderY.Value = [Math]::Max($minY, [Math]::Min($maxY, $data.Form.Top))

                $profilePrefix = $null
                if (Get-Command FindOrCreateProfile -ErrorAction SilentlyContinue) {
                    $profilePrefix = FindOrCreateProfile $data.WindowTitle
                }

                if ($profilePrefix) {
                    $configSuffix = ""
                    $parts = $data.InstanceId -split '_'
                    if ($parts.Count -gt 2) {
                        $configSuffix = "_Sub" + $parts[-1]
                    }
                    $p = "$profilePrefix$configSuffix"
                
                    $keysToClear = [System.Collections.ArrayList]@()
                    foreach ($k in $global:DashboardConfig.Config['Macro'].Keys) {
                        if ($k -match "_$p$") {
                            $keysToClear.Add($k)
                        }
                    }
                    foreach ($k in $keysToClear) {
                        $global:DashboardConfig.Config['Macro'].Remove($k)
                    }
                }

                UpdateMacroSettings $data -forceWrite
                RepositionSequences $form
            }
        })

    $formData.BtnAddInstance.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag
            $row = $data.Row

            if (-not $row -or -not $row.Tag) { return }

            $rowIdentity = Get-RowIdentity -Row $row
            $safeIdentity = $rowIdentity -replace '[^a-zA-Z0-9_\-]', '_'
            $baseInstanceId = "Macro_$safeIdentity"

            $nextIndex = 1
            $foundGap = $false
        
            while (-not $foundGap) {
                $testId = "${baseInstanceId}_${nextIndex}"
                if (-not $global:DashboardConfig.Resources.FtoolForms.Contains($testId) -or 
                    $global:DashboardConfig.Resources.FtoolForms[$testId].IsDisposed) {
                    $foundGap = $true
                }
                else {
                    $nextIndex++
                }
            }
        
            $newInstanceId = "${baseInstanceId}_${nextIndex}"

            $anchorForm = $form
        
            for ($i = $nextIndex - 1; $i -ge 0; $i--) {
                $prevId = if ($i -eq 0) { $baseInstanceId } else { "${baseInstanceId}_${i}" }
            
                if ($global:DashboardConfig.Resources.FtoolForms.Contains($prevId)) {
                    $candidate = $global:DashboardConfig.Resources.FtoolForms[$prevId]
                    if ($candidate -and -not $candidate.IsDisposed -and $candidate.Visible) {
                        $anchorForm = $candidate
                        break
                    }
                }
            }

            $newForm = CreateMacroForm -instanceId $newInstanceId -targetWindowRect $null -windowTitle $data.WindowTitle -row $row
        
            if ($newForm -and $newForm -is [System.Windows.Forms.Form]) {
                $newForm.Left = $anchorForm.Right + 10
                $newForm.Top = $anchorForm.Top
                $global:DashboardConfig.Resources.FtoolForms[$newInstanceId] = $newForm
                $newForm.Show()
                $newForm.BringToFront()
                if (Get-Command RegisterConfiguredHotkeys -ErrorAction SilentlyContinue) { try { RegisterConfiguredHotkeys } catch {} }
                if (Get-Command RefreshHotkeysList -ErrorAction SilentlyContinue) { try { RefreshHotkeysList } catch {} }
            }
            else {
                Write-Verbose "Error: CreateMacroForm returned unexpected type: $($newForm.GetType().FullName)"
            }
        })


    $formData.BtnHotKey.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag

            $currentHotkeyText = $this.Text
            if ($currentHotkeyText -eq 'Hotkey') { $currentHotkeyText = $null }
            $oldId = $data.HotkeyId

            $newHotkey = Show-KeyCaptureDialog $currentHotkeyText -OwnerForm $form
        
            if ($newHotkey -and $newHotkey -ne $currentHotkeyText) {
                $data.Hotkey = $newHotkey
                try {
                    $scriptBlock = [scriptblock]::Create("ToggleMacroInstance -InstanceId '$($data.InstanceId)'")
                    $ownerKey = "macroinst_$($data.InstanceId)"
                    $ownerLabel = $null
                    if ($data.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($data.Name)) { $ownerLabel = $data.Name }
                    elseif ($data.PSObject.Properties['WindowTitle'] -and -not [string]::IsNullOrEmpty($data.WindowTitle)) { $ownerLabel = $data.WindowTitle }
                    else { $ownerLabel = "Macro: $($data.InstanceId)" }
                    $data.HotkeyId = SetHotkey -KeyCombinationString $newHotkey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel -OldHotkeyId $oldId
                    $this.Text = $newHotkey
                }
                catch {
                    $data.HotkeyId = $null
                    $this.Text = 'Hotkey'
                }
                UpdateMacroSettings $data -forceWrite
            }
            elseif (-not $newHotkey -and $oldId) {
                try { UnregisterHotkeyInstance -Id $oldId -OwnerKey "macroinst_$($data.InstanceId)" } catch {}
                $data.HotkeyId = $null
                $this.Text = 'Hotkey'
                UpdateMacroSettings $data -forceWrite
            }
        })

    $formData.BtnHotkeyToggle.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag
            $toggleOn = $this.Checked
            ToggleMacroHotkeys -InstanceId $data.InstanceId -ToggleState $toggleOn
            UpdateMacroSettings $data -forceWrite
        })

    $formData.BtnInstanceHotkeyToggle.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag
        
            $currentHotkeyText = $data.GlobalHotkey
            $oldHotkeyIdToUnregister = $data.GlobalHotkeyId
            $ownerKey = "global_toggle_$($data.InstanceId)"
        
            $newHotkey = Show-KeyCaptureDialog $currentHotkeyText -OwnerForm $form
        
            if ($newHotkey -and $newHotkey -ne $currentHotkeyText) {
                $data.GlobalHotkey = $newHotkey
                try {
                    $script = @"
if (`$global:DashboardConfig.Resources.FtoolForms.Contains('$($data.InstanceId)')) {
    `$f = `$global:DashboardConfig.Resources.FtoolForms['$($data.InstanceId)']
    if (`$f -and -not `$f.IsDisposed -and `$f.Tag) {
        `$toggle = `$f.Tag.BtnHotkeyToggle
        if (`$toggle) {
            if (`$toggle.InvokeRequired) { `$toggle.Invoke([System.Action]{ `$toggle.Checked = -not `$toggle.Checked }) } else { `$toggle.Checked = -not `$toggle.Checked }
            ToggleMacroHotkeys -InstanceId '$($data.InstanceId)' -ToggleState `$toggle.Checked
        }
    }
}
"@
                    $scriptBlock = [scriptblock]::Create($script)
                    $ownerLabel = $null
                    if ($data.PSObject.Properties['Name'] -and -not [string]::IsNullOrEmpty($data.Name)) { $ownerLabel = $data.Name }
                    elseif ($data.PSObject.Properties['WindowTitle'] -and -not [string]::IsNullOrEmpty($data.WindowTitle)) { $ownerLabel = $data.WindowTitle }
                    else { $ownerLabel = "Macro: $($data.InstanceId)" }
                    $data.GlobalHotkeyId = SetHotkey -KeyCombinationString $data.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey -OwnerLabel $ownerLabel -OldHotkeyId $oldHotkeyIdToUnregister
                }
                catch { $data.GlobalHotkeyId = $null; $data.GlobalHotkey = $currentHotkeyText }
                UpdateMacroSettings $data -forceWrite
            }
            elseif (-not $newHotkey -and $oldHotkeyIdToUnregister) {
                try { UnregisterHotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $ownerKey } catch {}
                try { UnregisterHotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $data.InstanceId } catch {}
                $data.GlobalHotkeyId = $null
                $data.GlobalHotkey = $null
                UpdateMacroSettings $data -forceWrite
            }
        })

    $formData.Interval.Add_TextChanged({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            UpdateMacroSettings $form.Tag -forceWrite
        })

    $formData.Name.Add_TextChanged({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            UpdateMacroSettings $form.Tag -forceWrite
        })

    $formData.BtnLoopToggle.Add_Click({ UpdateMacroSettings $this.FindForm().Tag -forceWrite })
    $formData.BtnWaitToggle.Add_Click({ UpdateMacroSettings $this.FindForm().Tag -forceWrite })
    $formData.BtnHoldKeyToggle.Add_Click({ UpdateMacroSettings $this.FindForm().Tag -forceWrite })
    $formData.TxtHoldKeyInterval.Add_TextChanged({ UpdateMacroSettings $this.FindForm().Tag -forceWrite })

    $formData.PositionSliderX.Add_ValueChanged({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag
            if (-not $data.PSObject.Properties['IsDragging'] -or -not $data.IsDragging) {
                UpdateMacroSettings $data -forceWrite
            }
        })

    $formData.PositionSliderY.Add_ValueChanged({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag
            if (-not $data.PSObject.Properties['IsDragging'] -or -not $data.IsDragging) {
                UpdateMacroSettings $data -forceWrite
            }
        })

	$formData.BtnImport.Add_Click({
            $form = $this.FindForm()
            if (-not $form -or -not $form.Tag) { return }
            $data = $form.Tag

            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = "MGP Files (*.mgp)|*.mgp|All Files (*.*)|*.*"
            $ofd.Title = "Select MGP Macro File to Import"
            if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                Import-MgpMacro -formData $data -filePath $ofd.FileName
            }
        })
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion