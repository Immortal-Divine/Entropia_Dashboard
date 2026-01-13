<# macro.psm1 #>

#region Helper Functions

function InitMacroRunspace
{
    param($InstanceId)
    
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
    $assemblies = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { 
        $_.FullName -match 'Custom.Native' -or 
        $_.FullName -match 'Custom.Ftool' -or 
        $_.FullName -match 'Custom.Win32'
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

function Get-MacroKeyScriptBlock
{
    return {
        param($Handle, $Vk, $HoldTime, $IsMouse, $Mods, $IsAlt)

        $WM_KEYDOWN = 0x0100; $WM_KEYUP = 0x0101
        $WM_LBUTTONDOWN = 0x0201; $WM_LBUTTONUP = 0x0202
        $WM_RBUTTONDOWN = 0x0204; $WM_RBUTTONUP = 0x0205
        $WM_MBUTTONDOWN = 0x0207; $WM_MBUTTONUP = 0x0208
        $WM_XBUTTONDOWN = 0x020B; $WM_XBUTTONUP = 0x020C

        $point = New-Object Custom.Win32Point
        [Custom.Win32MouseUtils]::GetCursorPos([ref]$point) | Out-Null
        [Custom.Win32MouseUtils]::ScreenToClient($Handle, [ref]$point) | Out-Null
        $lParam = [Custom.Win32MouseUtils]::MakeLParam($point.X, $point.Y)

		if (-not $IsMouse) {
			$lParam = 1
		}

        $msgDown = 0; $msgUp = 0; $wParamDown = [IntPtr]::Zero; $wParamUp = [IntPtr]::Zero

        if ($IsMouse)
        {
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
        else
        {
            $msgDown = $WM_KEYDOWN
            $msgUp = $WM_KEYUP
            $wParamDown = [IntPtr]$Vk
            $wParamUp = [IntPtr]$Vk
        }

        if ($IsMouse -and $IsAlt) { [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, 0x12, 0) }
        foreach ($m in $Mods) { [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYDOWN, $m, 0) }

        [Custom.Ftool]::fnPostMessage($Handle, $msgDown, $wParamDown, $lParam)

        if ($HoldTime -gt 0)
        {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            while ($sw.ElapsedMilliseconds -lt $HoldTime)
            {
                Start-Sleep -Milliseconds 5
                if ($sw.ElapsedMilliseconds -lt $HoldTime -and -not $IsMouse)
                {
                    [Custom.Ftool]::fnPostMessage($Handle, $msgDown, $wParamDown, 0x40000000)
                }
            }
            $sw.Stop()
        }
        else
        {
            Start-Sleep -Milliseconds 5
        }

        $finalLParam = $lParam
        
        if (-not $IsMouse) {
            $finalLParam = $lParam -bor 0xC0000000
        }

        [Custom.Ftool]::fnPostMessage($Handle, $msgUp, $wParamUp, $finalLParam)

        foreach ($m in $Mods) { [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYUP, $m, 0xC0000000) }
        if ($IsMouse -and $IsAlt) { [Custom.Ftool]::fnPostMessage($Handle, $WM_KEYUP, 0x12, 0xC0000000) }
    }
}

function Invoke-MacroKeyAction
{
    param($formData, $keyString, [int]$holdTime = 5)

    if ($formData.SelectedWindow -eq [IntPtr]::Zero) { return $null }

    if (-not $formData.RunspacePool -or $formData.RunspacePool.RunspacePoolStateInfo.State -ne 'Opened')
    {
        Write-Verbose "Macro Runspace closed/broken. Attempting to restart..."
        try {
            if ($formData.RunspacePool) { $formData.RunspacePool.Dispose() }
            $formData.RunspacePool = InitMacroRunspace -InstanceId $formData.InstanceId
        }
        catch {
            Write-Verbose "Failed to restart Runspace: $_"
            return $null
        }
    }

    $parsedKey = ParseKeyString -KeyCombinationString $keyString
    if (-not $parsedKey -or -not $parsedKey.Primary) { return $null }

    $primaryKeyName = $parsedKey.Primary
    $modifierNames = $parsedKey.Modifiers
    $keyMappings = GetVirtualKeyMappings
    $virtualKeyCode = $keyMappings[$primaryKeyName]

    if (-not $virtualKeyCode) { return $null }

    $isMouse = ($virtualKeyCode -ge 0x01 -and $virtualKeyCode -le 0x06)
    $mods = @()
    $isAlt = $false

    foreach ($modName in $modifierNames) {
        switch ($modName.ToUpper()) {
            'CTRL' { $mods += 0x11 }
            'ALT' { $mods += 0x12; $isAlt = $true }
            'SHIFT' { $mods += 0x10 }
        }
    }
    
    if (-not $isMouse) { if ($isAlt) { $isAlt = $false } }

    $ps = [System.Management.Automation.PowerShell]::Create()
    try {
        $ps.RunspacePool = $formData.RunspacePool
        $ps.AddScript((Get-MacroKeyScriptBlock).ToString()) | Out-Null
        $ps.AddArgument($formData.SelectedWindow) | Out-Null
        $ps.AddArgument($virtualKeyCode) | Out-Null
        $ps.AddArgument($holdTime) | Out-Null
        $ps.AddArgument($isMouse) | Out-Null
        $ps.AddArgument($mods) | Out-Null
        $ps.AddArgument($isAlt) | Out-Null

        return @{ 
            PowerShell = $ps
            AsyncResult = $ps.BeginInvoke()
        }
    } catch {
        Write-Verbose "Macro Invoke Error: $_"
        if ($ps) { $ps.Dispose() }
        return $null
    }
}

function CreateMacroPositionTimer
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

						$currentLeft = $timerData['Form'].Left
						$newLeft = $currentLeft + ($targetLeft - $currentLeft) * 0.2 
						$timerData['Form'].Left = [int]$newLeft

						$currentTop = $timerData['Form'].Top
						$newTop = $currentTop + ($targetTop - $currentTop) * 0.2 
						$timerData['Form'].Top = [int]$newTop
                                                
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
					$global:DashboardConfig.Resources.Timers.Remove("macroPosition_$($timerData.InstanceId)")
				}
			}
			catch {}
		})
    
	$positionTimer.Start()
	$global:DashboardConfig.Resources.Timers["macroPosition_$($formData.InstanceId)"] = $positionTimer
}

function RepositionSequences
{
	param($form)
    
	if (-not $form -or $form.IsDisposed) { return }
    
	$form.SuspendLayout()
	try
	{
		$baseHeight = 140 
		$rowHeight = 70
		
		$newHeight = $baseHeight
		$position = 0
        
		$formData = $form.Tag
		if ($formData.SequencePanels)
		{
			foreach ($panel in $formData.SequencePanels)
			{
				if ($panel -and -not $panel.IsDisposed)
				{
					$newTop = $baseHeight + ($position * $rowHeight)
					$panel.Top = $newTop
					$position++
				}
			}
			if ($position -gt 0)
			{
				$newHeight = $baseHeight + ($position * $rowHeight)
			}
		}
        
		if (-not $formData.IsCollapsed)
		{
			$finalHeight = $newHeight + 10
			$form.Height = $finalHeight
			$formData.OriginalHeight = $finalHeight
			$formData.PositionSliderY.Height = $finalHeight - 20
		}
	}
	finally
	{
		$form.ResumeLayout()
	}
}

function StartMacroSequence
{
	param($formData)

	$steps = @()

	$mainKey = $formData.BtnKeySelect.Text
	$mainInterval = 1000
	if ([int]::TryParse($formData.Interval.Text, [ref]$mainInterval)) { if ($mainInterval -lt 1) { $mainInterval = 1 } }
	
	$mainHoldEnabled = $formData.BtnHoldKeyToggle.Checked
	$mainHoldInterval = $formData.TxtHoldKeyInterval.Text
	$mainWaitEnabled = $formData.BtnWaitToggle.Checked

	if (-not [string]::IsNullOrWhiteSpace($mainKey) -and $mainKey -ne 'none')
	{
		$steps += [PSCustomObject]@{ 
			Key = $mainKey; 
			Interval = $mainInterval; 
			HoldEnabled = $mainHoldEnabled; 
			HoldInterval = $mainHoldInterval;
			WaitEnabled = $mainWaitEnabled
		}
	}

	if ($formData.SequencePanels)
	{
		foreach ($panel in $formData.SequencePanels)
		{
			if ($panel -and -not $panel.IsDisposed)
			{
				$seqData = $panel.Tag
				$seqKey = $seqData.BtnKeySelect.Text
				$seqInterval = 1000
				if ([int]::TryParse($seqData.Interval.Text, [ref]$seqInterval)) { if ($seqInterval -lt 1) { $seqInterval = 1 } }
				$seqHoldEnabled = $seqData.BtnHoldKeyToggle.Checked
				$seqHoldInterval = $seqData.TxtHoldKeyInterval.Text
				$seqWaitEnabled = $seqData.BtnWaitToggle.Checked

				if (-not [string]::IsNullOrWhiteSpace($seqKey) -and $seqKey -ne 'none')
				{
					$steps += [PSCustomObject]@{ 
						Key = $seqKey; 
						Interval = $seqInterval; 
						HoldEnabled = $seqHoldEnabled; 
						HoldInterval = $seqHoldInterval;
						WaitEnabled = $seqWaitEnabled
					}
				}
			}
		}
	}

	if ($steps.Count -eq 0)
	{
		Show-DarkMessageBox 'No valid keys defined in the macro sequence.' 'Macro Error' 'Ok' 'Warning'
		return
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
        $formData.RunspacePool = InitMacroRunspace -InstanceId $formData.InstanceId
    }
	
	$formData.BtnStart.Enabled = $false
	$formData.BtnStart.Visible = $false
	$formData.BtnStop.Enabled = $true
	$formData.BtnStop.Visible = $true

    if (-not $formData.PSObject.Properties['IsMacroRunning'])
    {
        $formData | Add-Member -MemberType NoteProperty -Name 'IsMacroRunning' -Value $true
    }
    $formData.IsMacroRunning = $true

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

        if (-not $fData.IsMacroRunning) { return }

        if ($fData.Stopping)
        {
            StopMacroSequence $fData -NaturalEnd
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
			
			$holdTime = $null
			if ($step.HoldEnabled)
			{
				if (-not [int]::TryParse($step.HoldInterval, [ref]$holdTime)) { $holdTime = 5 }
			}

            $job = Invoke-MacroKeyAction -formData $fData -keyString $step.Key -holdTime $holdTime
            
            if ($job)
            {
                $fData.ActiveJobs.Add($job) | Out-Null
            }

			$nextIdx = ($idx + 1) % $fData.ActiveSteps.Count
			$fData.CurrentStepIndex = $nextIdx
			
			$isSingleRun = -not $fData.BtnLoopToggle.Checked
            $shouldStop = ($isSingleRun -and $nextIdx -eq 0)

            if ($step.WaitEnabled -and $job)
            {
                $pollTimer = New-Object System.Windows.Forms.Timer
                $pollTimer.Interval = 2
                $pollTimer.Tag = @{ 
                    Job = $job; 
                    MainTimer = $s; 
                    NextInterval = $step.Interval; 
                    ShouldStop = $shouldStop;
                    FormData = $fData
                }
                
                $pollTimer.Add_Tick({
                    $pt = $this
                    $d = $pt.Tag
                    
                    if (-not $d.FormData.IsMacroRunning)
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

                        if ($d.FormData.IsMacroRunning)
                        {
                            if ($d.ShouldStop)
                            {
                                StopMacroSequence $d.FormData -NaturalEnd
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
            else
            {
                if ($fData.IsMacroRunning)
                {
                    if ($shouldStop)
                    {
                        $fData.Stopping = $true
                        $s.Interval = $step.Interval
                        $s.Start()
                        return
                    }
                    $s.Interval = $step.Interval
                    $s.Start()
                }
            }
		}
	}
    
	$timer.Add_Tick($tickAction)

	$formData.RunningTimer = $timer
	$global:DashboardConfig.Resources.Timers["MacroTimer_$($formData.InstanceId)"] = $timer

    & $tickAction $timer $null
}

function StopMacroSequence
{
	param($formData, [switch]$NaturalEnd)

    if ($formData.PSObject.Properties['IsMacroRunning'])
    {
        $formData.IsMacroRunning = $false
    }

	if ($formData.RunningTimer)
	{
		$formData.RunningTimer.Stop()
		$formData.RunningTimer.Dispose()
		$formData.RunningTimer = $null
	}
	
	if ($global:DashboardConfig.Resources.Timers.Contains("MacroTimer_$($formData.InstanceId)"))
	{
        $t = $global:DashboardConfig.Resources.Timers["MacroTimer_$($formData.InstanceId)"]
        if ($t) { $t.Stop(); $t.Dispose() }
		$global:DashboardConfig.Resources.Timers.Remove("MacroTimer_$($formData.InstanceId)")
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
	$formData.BtnStart.Enabled = $true
	$formData.BtnStart.Visible = $true
	$formData.BtnStop.Enabled = $false
	$formData.BtnStop.Visible = $false
}

function ToggleMacroInstance
{
	param($InstanceId)

	if (-not $global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId)) { return }
	$form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
	
	if ($form -and -not $form.IsDisposed -and $form.Tag.Type -eq 'Macro')
	{
		if ($form.InvokeRequired)
		{
			$form.BeginInvoke([System.Action]{ ToggleMacroInstance -InstanceId $InstanceId }) | Out-Null
			return
		}

		$data = $form.Tag
		if ($data.RunningTimer)
		{
			StopMacroSequence $data
		}
		else
		{
			StartMacroSequence $data
		}
	}
}

function ToggleMacroHotkeys
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
			$form.BeginInvoke([System.Action]{ ToggleMacroHotkeys -InstanceId $InstanceId -ToggleState $ToggleState }) | Out-Null
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
			UpdateMacroSettings -formData $form.Tag -forceWrite
		}
	}
}

function CleanupMacroResources
{
	param($instanceId)
    
	$timerKeysToRemove = @()
	foreach ($key in $global:DashboardConfig.Resources.Timers.Keys)
	{
		if ($key -eq "MacroTimer_$instanceId" -or $key -eq "macroPosition_$instanceId")
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
                        Write-Verbose "Error disposing Macro Job: $_"
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
                    Write-Verbose "Error disposing Macro Runspace: $_"
                }
                $form.Tag.RunspacePool = $null
            }

			if ($form.Tag.HotkeyId) { try { UnregisterHotkeyInstance -Id $form.Tag.HotkeyId -OwnerKey $form.Tag.InstanceId } catch {} }
			if ($form.Tag.GlobalHotkeyId) { try { $globalOwnerKey = "global_toggle_$($form.Tag.InstanceId)"; UnregisterHotkeyInstance -Id $form.Tag.GlobalHotkeyId -OwnerKey $globalOwnerKey } catch {} }
		}
	}
}

function LoadMacroSettings
{
	param($formData)
    $formData.IsLoading = $true
    
	if (-not $global:DashboardConfig.Config.Contains('Macro')) { $formData.IsLoading = $false; return }
    
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
		if ($cfg.Contains($hotkeysEnabledName))
		{
			$value = $cfg[$hotkeysEnabledName]
			try { $formData.BtnHotkeyToggle.Checked = [bool]::Parse($value) } catch { $formData.BtnHotkeyToggle.Checked = $true }
		}
		else { $formData.BtnHotkeyToggle.Checked = $true }

		if ($cfg.Contains("PosX_$p")) { $formData.PositionSliderX.Value = [int]$cfg["PosX_$p"] }
		if ($cfg.Contains("PosY_$p")) { $formData.PositionSliderY.Value = [int]$cfg["PosY_$p"] }

		if ($cfg.Contains("Hotkey_$p"))
		{
			$hk = $cfg["Hotkey_$p"]
			$formData.Hotkey = $hk
			$formData.BtnHotKey.Text = $hk
            
			try
			{
				$scriptBlock = [scriptblock]::Create("ToggleMacroInstance -InstanceId '$($formData.InstanceId)'")
				$formData.HotkeyId = SetHotkey -KeyCombinationString $hk -Action $scriptBlock -OwnerKey $formData.InstanceId
			}
			catch {}
		}

		if ($cfg.Contains("SeqCount_$p"))
		{
			$count = [int]$cfg["SeqCount_$p"]
			for ($i = 1; $i -le $count; $i++)
			{
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
	else
	{
		$formData.Hotkey = $null
		$formData.BtnHotKey.Text = 'Hotkey'
		$formData.BtnHotkeyToggle.Checked = $true
	}
    $formData.IsLoading = $false
}

function UpdateMacroSettings
{
	param($formData, [switch]$forceWrite)
    
    if ($formData.IsLoading) { return }
    
	if (-not $global:DashboardConfig.Config.Contains('Macro'))
	{ 
		$global:DashboardConfig.Config['Macro'] = [ordered]@{} 
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

		if ($formData.Hotkey) { $cfg["Hotkey_$p"] = $formData.Hotkey } else { if ($cfg.Contains("Hotkey_$p")) { $cfg.Remove("Hotkey_$p") } }

		$seqs = $formData.SequencePanels
		$cfg["SeqCount_$p"] = $seqs.Count
        
		for ($i = 0; $i -lt $seqs.Count; $i++)
		{
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
		while ($cfg.Contains("Seq${k}_Key_$p"))
		{
			$cfg.Remove("Seq${k}_Key_$p")
			$cfg.Remove("Seq${k}_Interval_$p")
			$cfg.Remove("Seq${k}_Name_$p")
			$cfg.Remove("Seq${k}_HoldEnabled_$p")
			$cfg.Remove("Seq${k}_HoldInterval_$p")
			$cfg.Remove("Seq${k}_WaitEnabled_$p")
			$k++
		}

		if ($forceWrite)
		{
			if (Get-Command WriteConfig -ErrorAction SilentlyContinue) { WriteConfig }
		}
	}
}

#endregion

#region Core Functions

function MacroSelectedRow
{
	param($row)
    
	if (-not $row -or -not $row.Cells -or $row.Cells.Count -lt 3) { return }
    
    $instanceId = "Macro_" + $row.Cells[2].Value.ToString()

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
    
	$macroForm = CreateMacroForm $instanceId $targetWindowRect $windowTitle $row
    
	$global:DashboardConfig.Resources.FtoolForms[$instanceId] = $macroForm
    
	$macroForm.Show()
	$macroForm.BringToFront()
}

function StopMacroForm
{
	param($Form)
    
	if (-not $Form -or $Form.IsDisposed) { return }
    
	try
	{
		$instanceId = $Form.Tag.InstanceId
		if ($instanceId)
		{
			CleanupMacroResources $instanceId
			if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId)) { $global:DashboardConfig.Resources.FtoolForms.Remove($instanceId) }
		}
		$Form.Close()
		$Form.Dispose()
	}
	catch {}
}

function CreateMacroForm
{
	param($instanceId, $targetWindowRect, $windowTitle, $row)
    
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

	if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon))
	{
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
			$g.FillRectangle((New-Object System.Drawing.SolidBrush $c::FromArgb(30,30,30)), $b)
			$g.DrawRectangle((New-Object System.Drawing.Pen $c::FromArgb(100,100,100)), $b.X, $b.Y, $b.Width-1, $b.Height-1)
			$g.DrawString($_.ToolTipText, $this.TipFont, (New-Object System.Drawing.SolidBrush $c::FromArgb(240,240,240)), 3, 3, [System.Drawing.StringFormat]::GenericTypographic)
		})
	$macroToolTip.Add_Popup({
			$g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
			$s = $g.MeasureString($this.GetToolTip($_.AssociatedControl), $this.TipFont, [System.Drawing.PointF]::new(0,0), [System.Drawing.StringFormat]::GenericTypographic)
			$g.Dispose(); $_.ToolTipSize = [System.Drawing.Size]::new($s.Width+12, $s.Height+8)
		})

	$previousToolTip = $global:DashboardConfig.UI.ToolTipFtool
	$global:DashboardConfig.UI.ToolTipFtool = $macroToolTip

	try
	{

	$headerPanel = SetUIElement -type 'Panel' -visible $true -width 250 -height 20 -top 0 -left 0 -bg @(40, 40, 40)
	$macroForm.Controls.Add($headerPanel)

	$labelWinTitle = SetUIElement -type 'Label' -visible $true -width 120 -height 20 -top 5 -left 5 -bg @(40, 40, 40, 0) -fg @(255, 255, 255) -text $row.Cells[1].Value -font (New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Regular))
	$headerPanel.Controls.Add($labelWinTitle)

	$btnAddInstance = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 165 -bg @(40, 40, 40) -fg @(255, 255, 255) -text ([char]0x2398) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Add Instance`nCreates another macro window for this specific client."
	$headerPanel.Controls.Add($btnAddInstance)

	$btnInstanceHotkeyToggle = SetUIElement -type 'Label' -visible $true -width 15 -height 15 -top 2 -left 120 -bg @(40, 40, 40) -fg @(255, 255, 255) -text ([char]0x2328) -font (New-Object System.Drawing.Font('Segoe UI', 10)) -tooltip "Set Master Hotkey`nAssign a global hotkey to toggle all hotkeys for this instance."
	$headerPanel.Controls.Add($btnInstanceHotkeyToggle)

	$btnHotkeyToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 15 -top 3 -left 135 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true -tooltip "Toggle Hotkeys`nEnable or disable all hotkeys for this specific macro instance."
	$headerPanel.Controls.Add($btnHotkeyToggle)
        
	$btnAdd = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 180 -bg @(40, 80, 80) -fg @(255, 255, 255) -text ([char]0x2795) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Add Sequence`nAdds a new key sequence step to the macro list."
	$headerPanel.Controls.Add($btnAdd)
        
	$btnShowHide = SetUIElement -type 'Button' -visible $true -width 15 -height 15 -top 3 -left 195 -bg @(60, 60, 100) -fg @(255, 255, 255) -text ([char]0x25B2) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11)) -tooltip "Minimize/Expand`nCollapse or expand the macro window to save screen space."
	$headerPanel.Controls.Add($btnShowHide)
        
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

	$btnLoopToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 16 -top 42 -left 45 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true -tooltip "Loop`nToggle looping. If enabled, the macro repeats until stopped."
	$panelSettings.Controls.Add($btnLoopToggle)

	$btnWaitToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 16 -top 42 -left 80 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true -tooltip "Wait Mode`nIf enabled, the macro waits for the key hold duration to finish before proceeding."
	$panelSettings.Controls.Add($btnWaitToggle)

	$btnHoldKeyToggle = SetUIElement -type 'Toggle' -visible $true -width 30 -height 16 -top 42 -left 115 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $false -tooltip "Hold Key`nEnable to hold the key down for a specific duration instead of a single press."
	$panelSettings.Controls.Add($btnHoldKeyToggle)

	$txtHoldKeyInterval = SetUIElement -type 'TextBox' -visible $true -width 30 -height 15 -top 42 -left 150 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '50' -font (New-Object System.Drawing.Font('Segoe UI', 9)) -tooltip "Hold Duration (ms)`nHow long the key should be held down in milliseconds."
	$panelSettings.Controls.Add($txtHoldKeyInterval)
    
	$positionSliderY = New-Object System.Windows.Forms.TrackBar
	$positionSliderY.Orientation = 'Vertical'
	$positionSliderY.Minimum = -18
	$positionSliderY.Maximum = 118
	$positionSliderY.TickFrequency = 300
	$positionSliderY.Value = 0
	$positionSliderY.Size = New-Object System.Drawing.Size(1, 110)
	$positionSliderY.Location = New-Object System.Drawing.Point(5, 20)
	$macroForm.Controls.Add($positionSliderY)
        
	$positionSliderX = New-Object System.Windows.Forms.TrackBar
	$positionSliderX.Minimum = -25
	$positionSliderX.Maximum = 125
	$positionSliderX.TickFrequency = 300
	$positionSliderX.Value = 0
	$positionSliderX.Size = New-Object System.Drawing.Size(190, 15)
	$positionSliderX.Location = New-Object System.Drawing.Point(45, 25)
	$macroForm.Controls.Add($positionSliderX)

	}
	finally
	{
		$global:DashboardConfig.UI.ToolTipFtool = $previousToolTip
	}

	$formData = [PSCustomObject]@{
		Type            = 'Macro'
		InstanceId      = $instanceId
		SelectedWindow  = $row.Tag.MainWindowHandle
		BtnKeySelect    = $btnKeySelect
		Interval        = $interval
		Name            = $name
		BtnStart        = $btnStart
		BtnStop         = $btnStop
		BtnLoopToggle   = $btnLoopToggle
		BtnWaitToggle   = $btnWaitToggle
		BtnHoldKeyToggle = $btnHoldKeyToggle
		TxtHoldKeyInterval = $txtHoldKeyInterval
		BtnHotKey       = $btnHotKey
		BtnInstanceHotkeyToggle = $btnInstanceHotkeyToggle
		BtnHotkeyToggle = $btnHotkeyToggle
		BtnAdd          = $btnAdd
		BtnAddInstance  = $btnAddInstance
		BtnClose        = $btnClose
		BtnShowHide     = $btnShowHide
		PositionSliderX = $positionSliderX
		PositionSliderY = $positionSliderY
		Form            = $macroForm
		RunningTimer    = $null
		WindowTitle     = $windowTitle
		Row             = $row
		ToolTipFtool    = $macroToolTip
		IsCollapsed     = $false
		OriginalHeight  = 145
		HotkeyId        = $null
		GlobalHotkeyId  = $null
		GlobalHotkey    = $null
		Hotkey          = $null
		SequencePanels  = [System.Collections.ArrayList]::new()
		ActiveSteps     = @()
		CurrentStepIndex= 0
        ActiveJobs      = [System.Collections.ArrayList]::new()
        RunspacePool    = (InitMacroRunspace -InstanceId $instanceId)
        IsLoading       = $false
	}
    
	$macroForm.Tag = $formData
    
	CreateMacroPositionTimer $formData | Out-Null
	LoadMacroSettings $formData | Out-Null
	ToggleMacroHotkeys -InstanceId $formData.InstanceId -ToggleState $formData.BtnHotkeyToggle.Checked | Out-Null

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
            ToggleMacroHotkeys -InstanceId '$($formData.InstanceId)' -ToggleState `$toggle.Checked
        }
    }
}
"@
			$scriptBlock = [scriptblock]::Create($script)
			$formData.GlobalHotkeyId = SetHotkey -KeyCombinationString $formData.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey
		}
		catch
		{
			$formData.GlobalHotkeyId = $null
		}
	}

	AddMacroEventHandlers $formData | Out-Null
    
	return $macroForm
}

function CreateSequencePanel
{
	param($form)
	$formData = $form.Tag

	$previousToolTip = $global:DashboardConfig.UI.ToolTipFtool
	if ($formData.ToolTipFtool)
	{
		$global:DashboardConfig.UI.ToolTipFtool = $formData.ToolTipFtool
	}

	try
	{
	
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
	finally
	{
		$global:DashboardConfig.UI.ToolTipFtool = $previousToolTip
	}

	$seqData = [PSCustomObject]@{
		Panel        = $panelSeq
		BtnKeySelect = $btnKeySelect
		Interval     = $interval
		Name         = $name
		BtnWaitToggle = $btnWaitToggle
		BtnHoldKeyToggle = $btnHoldKeyToggle
		TxtHoldKeyInterval = $txtHoldKeyInterval
		BtnRemove    = $btnRemove
	}
	$panelSeq.Tag = $seqData
	
	$formData.SequencePanels.Add($panelSeq) | Out-Null

	$btnKeySelect.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag) { return }
		$formData = $form.Tag

		$currentKey = $this.Text
		$newKey = Show-KeyCaptureDialog $currentKey -OwnerForm $form
		if ($newKey -and $newKey -ne $currentKey)
		{
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

function AddMacroEventHandlers
{
	param($formData)

	$formData.BtnKeySelect.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag) { return }
		$data = $form.Tag

		$currentKey = $this.Text
		$newKey = Show-KeyCaptureDialog $currentKey -OwnerForm $form
		if ($newKey -and $newKey -ne $currentKey)
		{
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

		if ($data.IsCollapsed)
		{
			$form.Height = $data.OriginalHeight
			$data.IsCollapsed = $false
			$this.Text = [char]0x25B2
		}
		else
		{
			$data.OriginalHeight = $form.Height
			$form.Height = 26
			$data.IsCollapsed = $true
			$this.Text = [char]0x25BC
		}
	})

	$formData.BtnClose.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag) { return }
		$null = $data; $data = $form.Tag

		StopMacroForm $form
	})

	$formData.BtnAddInstance.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag) { return }
		$data = $form.Tag
		$row = $data.Row

        if (-not $row -or -not $row.Tag) { return }

        $baseInstanceId = "Macro_" + $row.Cells[2].Value.ToString()

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

		$newForm = CreateMacroForm -instanceId $newInstanceId -targetWindowRect $null -windowTitle $data.WindowTitle -row $row
		
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
		
		if ($newHotkey -and $newHotkey -ne $currentHotkeyText)
		{
			$data.Hotkey = $newHotkey
			try
			{
				$scriptBlock = [scriptblock]::Create("ToggleMacroInstance -InstanceId '$($data.InstanceId)'")
				$data.HotkeyId = SetHotkey -KeyCombinationString $newHotkey -Action $scriptBlock -OwnerKey $data.InstanceId -OldHotkeyId $oldId
				$this.Text = $newHotkey
			}
			catch
			{
				$data.HotkeyId = $null
				$this.Text = 'Hotkey'
			}
			UpdateMacroSettings $data -forceWrite
		}
		elseif (-not $newHotkey -and $oldId)
		{
			UnregisterHotkeyInstance -Id $oldId -OwnerKey $data.InstanceId
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
            ToggleMacroHotkeys -InstanceId '$($data.InstanceId)' -ToggleState `$toggle.Checked
        }
    }
}
"@
				$scriptBlock = [scriptblock]::Create($script)
				$data.GlobalHotkeyId = SetHotkey -KeyCombinationString $data.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey -OldHotkeyId $oldHotkeyIdToUnregister
			}
			catch { $data.GlobalHotkeyId = $null; $data.GlobalHotkey = $currentHotkeyText }
			UpdateMacroSettings $data -forceWrite
		}
		elseif (-not $newHotkey -and $oldHotkeyIdToUnregister)
		{
			try { UnregisterHotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $ownerKey } catch {}
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
		UpdateMacroSettings $form.Tag -forceWrite
	})

	$formData.PositionSliderY.Add_ValueChanged({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag) { return }
		UpdateMacroSettings $form.Tag -forceWrite
	})
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion