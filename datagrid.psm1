<# datagrid.psm1 #>

#region Global Configuration
$script:UpdateInterval = 150

$script:States = @{
	Checking  = 'Checking...'
	Loading   = 'Loading...'
	Minimized = 'Min'
	Normal    = 'Normal'
}

$script:ProcessCache = @{}

$script:PendingChecks = [System.Collections.Generic.HashSet[int]]::new()

#endregion

#region Helper Functions
function TestValidParameters
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.DataGridView]$Grid
	)
	if (-not $Grid) { return $false }
	if (-not $global:DashboardConfig -or -not $global:DashboardConfig.Config -or -not $global:DashboardConfig.Config.Contains('ProcessName')) { return $false }
	$processName = $global:DashboardConfig.Config['ProcessName']['ProcessName']
	if ([string]::IsNullOrEmpty($processName)) { return $false }
	return $true
}

function RestoreWindowStyles
{
	[CmdletBinding()]
	param()
	foreach ($processId in $script:ProcessCache.Keys)
	{
		try
		{
			$hWnd = [IntPtr]::Zero
			if ($script:ProcessCache[$processId].hWnd) { $hWnd = $script:ProcessCache[$processId].hWnd }
			if ($hWnd -eq [IntPtr]::Zero -and ('Custom.SafeWindowCore' -as [Type])) { $hWnd = [Custom.SafeWindowCore]::FindBestWindow($processId) }
			if ($hWnd -ne [IntPtr]::Zero)
			{
				if ([Custom.Native]::IsWindow($hWnd)) { SetWindowToolStyle -hWnd $hWnd -Hide $false }
			}
		}
		catch {}
	}
}

function GetProcessList
{
	[CmdletBinding()]
	[OutputType([System.Diagnostics.Process[]])]
	param()
	$processName = $global:DashboardConfig.Config['ProcessName']['ProcessName']
	try
	{
		$allProcesses = Get-Process -Name $processName -ErrorAction SilentlyContinue
		$validProcesses = @()
		foreach ($proc in $allProcesses) { try { if (-not $proc.HasExited) { $validProcesses += $proc } } catch {} }
		return $validProcesses
	}
 catch { return @() }
}

function RemoveTerminatedProcesses
{
	[CmdletBinding()]
	param([System.Windows.Forms.DataGridView]$Grid, [System.Diagnostics.Process[]]$CurrentProcesses = @())
	if ($Grid.Rows.Count -eq 0) { return }
	$rowsToRemove = [System.Collections.Generic.List[System.Windows.Forms.DataGridViewRow]]::new()
	$processIdLookup = @{}
	foreach ($process in $CurrentProcesses) { if ($null -ne $process -and $null -ne $process.Id) { $processIdLookup[$process.Id] = $true } }

	foreach ($row in $Grid.Rows)
	{
		if (-not $row.Tag -or $null -eq $row.Tag.Id) { continue }
		if (-not $processIdLookup.ContainsKey($row.Tag.Id))
		{
			$rowsToRemove.Add($row)
			if ($script:ProcessCache.ContainsKey($row.Tag.Id)) { $script:ProcessCache.Remove($row.Tag.Id) }
			if ($script:PendingChecks.Contains($row.Tag.Id)) { $script:PendingChecks.Remove($row.Tag.Id) | Out-Null }
		}
	}
	for ($i = $rowsToRemove.Count - 1; $i -ge 0; $i--)
	{
		if ($rowsToRemove[$i].Index -ge 0 -and $rowsToRemove[$i].Index -lt $Grid.Rows.Count) { $Grid.Rows.Remove($rowsToRemove[$i]) }
	}
}

function NewRowLookupDictionary
{
	[CmdletBinding()]
	[OutputType([hashtable])]
	param([System.Windows.Forms.DataGridView]$Grid)
	$rowLookup = @{}
	foreach ($row in $Grid.Rows) { if ($row.Tag -and $null -ne $row.Tag.Id) { $rowLookup[$row.Tag.Id] = $row } }
	return $rowLookup
}

function UpdateExistingRow
{
	[CmdletBinding()]
	param([System.Windows.Forms.DataGridViewRow]$Row, [int]$Index, [string]$ProcessTitle, [int]$ProcessId, [System.Diagnostics.Process]$Process)
	if ($Row.Cells[0].Value -ne $Index) { $Row.Cells[0].Value = $Index }
	if ($Row.Cells[1].Value -ne $ProcessTitle) { $Row.Cells[1].Value = $ProcessTitle }
	if ($Row.Cells[2].Value -ne $ProcessId) { $Row.Cells[2].Value = $ProcessId }
	$Row.Tag = $Process
}

function UpdateRowIndices
{
	[CmdletBinding()]
	param([System.Windows.Forms.DataGridView]$Grid)
	try
	{
		if ($Grid.Rows.Count -eq 0) { return }
		$profileCounters = @{}
		for ($i = 0; $i -lt $Grid.Rows.Count; $i++)
		{
			$row = $Grid.Rows[$i]
			$processTitle = $row.Cells[1].Value
			$profileName = ''
			if ($processTitle -match '^\[([^\]]+)\]') { $profileName = $Matches[1] }
			if ([string]::IsNullOrEmpty($profileName)) { $profileName = 'NoProfile' }
			if (-not $profileCounters.ContainsKey($profileName)) { $profileCounters[$profileName] = 0 }
			$profileCounters[$profileName]++
			$newIndex = $profileCounters[$profileName]
			if ($row.Cells[0].Value -ne $newIndex) { $row.Cells[0].Value = $newIndex }
		}
	}
 catch {}
}

function AddNewProcessRow
{
	[CmdletBinding()]
	[OutputType([System.Windows.Forms.DataGridViewRow])]
	param([System.Windows.Forms.DataGridView]$Grid, [System.Diagnostics.Process]$Process, [int]$Index, [string]$ProcessTitle)
	try
	{
		$rowIndex = $Grid.Rows.Add($Index, $ProcessTitle, $Process.Id, $script:States.Checking)
		$Grid.Rows[$rowIndex].Tag = $Process
		if (-not $script:ProcessCache.ContainsKey($Process.Id)) { $script:ProcessCache[$Process.Id] = @{} }
		$script:ProcessCache[$Process.Id]['LastSeen'] = [DateTime]::Now
		return $Grid.Rows[$rowIndex]
	}
 catch { return $null }
}

function StartWindowStateCheck
{
	[CmdletBinding()]
	param([System.Diagnostics.Process]$Process, [System.Windows.Forms.DataGridViewRow]$Row, [System.Windows.Forms.DataGridView]$Grid, [IntPtr]$CachedHandle = [IntPtr]::Zero)

	if ($script:PendingChecks.Contains($Process.Id)) { return }
	$targetHandle = $CachedHandle
	if ($targetHandle -eq [IntPtr]::Zero)
	{
		return
	}

	try
	{
		$null = $script:PendingChecks.Add($Process.Id)
		$task = [Custom.Native]::ResponsiveAsync($targetHandle)
		$stateData = @{ ProcessId = $Process.Id; RowIndex = $Row.Index; hWnd = $targetHandle }
		$scheduler = [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext()

		$continuationAction = [System.Action[System.Threading.Tasks.Task, Object]] {
			param($completedTask, $state)
			$script:PendingChecks.Remove($state.ProcessId) | Out-Null
			try
			{
				$currentGrid = $global:DashboardConfig.UI.DataGridMain
				$processId = $state.ProcessId; $rowIndex = $state.RowIndex; $hWnd = $state.hWnd
				if (-not $currentGrid -or $currentGrid.IsDisposed -or $rowIndex -lt 0 -or $rowIndex -ge $currentGrid.Rows.Count) { return }

				$targetRow = $currentGrid.Rows[$rowIndex]
				if ($null -eq $targetRow -or $null -eq $targetRow.Tag -or $targetRow.Tag.Id -ne $processId)
				{
					$targetRow = FindTargetRow -Grid $currentGrid -ProcessId $processId -RowIndex -1
					if ($null -eq $targetRow) { return }
				}

				$isResponsive = $null
				if ($completedTask.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion) { $isResponsive = $completedTask.Result }
				$isMinimized = $false
				if ('Custom.SafeWindowCore' -as [Type]) { $isMinimized = [Custom.SafeWindowCore]::IsMinimized($hWnd) }
				else { $isMinimized = [Custom.Native]::IsMinimized($hWnd) }

				$finalState = if (-not $isResponsive) { $script:States.Loading }
				elseif ($isMinimized) { $script:States.Minimized }
				else { $script:States.Normal }

				if ($isResponsive -and $hWnd -ne 0 -and $hWnd -ne [IntPtr]::Zero)
				{
					$shouldHide = $false
					$rowTitle = $targetRow.Cells[1].Value
					if ($rowTitle -match '^\[(.*?)\]')
					{
						$profileName = $matches[1]
						if ($global:DashboardConfig.Config['HideProfiles'] -and $global:DashboardConfig.Config['HideProfiles'].Contains($profileName)) { $shouldHide = $true }
					}
					if ($shouldHide)
					{
						if ($isMinimized) { SetWindowToolStyle -hWnd $hWnd -Hide $true }
						else { SetWindowToolStyle -hWnd $hWnd -Hide $false }
					}
					else { SetWindowToolStyle -hWnd $hWnd -Hide $false }
				}

				$currentState = $targetRow.Cells[3].Value
				
				if ($currentState -in @('Reconnecting...', 'Queued...', 'Cancelled', 'Error') -or $currentState -like 'Client *')
				{
					
					
					return
				}

				if ($currentState -ne $finalState) { $targetRow.Cells[3].Value = $finalState }
			}
			catch { }
		}
		$task.ContinueWith($continuationAction, $stateData, [System.Threading.CancellationToken]::None, [System.Threading.Tasks.TaskContinuationOptions]::None, $scheduler)
	}
 catch
	{
		$script:PendingChecks.Remove($Process.Id) | Out-Null
	}
}

function FindTargetRow
{
	[CmdletBinding()]
	[OutputType([System.Windows.Forms.DataGridViewRow])]
	param([System.Windows.Forms.DataGridView]$Grid, [int]$ProcessId, [int]$RowIndex, [IntPtr]$WindowHandle = [IntPtr]::Zero)
	$row = $null
	if ($RowIndex -ge 0 -and $RowIndex -lt $Grid.Rows.Count)
	{
		$potentialRow = $Grid.Rows[$RowIndex]
		if ($potentialRow.Tag -and $potentialRow.Tag.Id -eq $ProcessId)
		{
			if ($WindowHandle -eq [IntPtr]::Zero -or ($potentialRow.Tag.MainWindowHandle -eq $WindowHandle)) { $row = $potentialRow }
		}
	}
	if ($null -eq $row)
	{
		foreach ($r in $Grid.Rows)
		{
			if ($r.Tag -and $r.Tag.Id -eq $ProcessId)
			{
				if ($WindowHandle -eq [IntPtr]::Zero -or ($r.Tag.MainWindowHandle -eq $WindowHandle)) { $row = $r; break }
			}
		}
	}
	return $row
}

function ClearOldProcessCache
{
	[CmdletBinding()]
	param()
	$now = [DateTime]::Now; $keysToRemove = [System.Collections.Generic.List[object]]::new(); $thresholdMinutes = 5
	$cacheKeys = @($script:ProcessCache.Keys)
	foreach ($key in $cacheKeys)
	{
		if ($script:ProcessCache.ContainsKey($key))
		{
			$entry = $script:ProcessCache[$key]
			if ($entry -and $entry.ContainsKey('LastSeen') -and $entry.LastSeen -is [DateTime])
			{
				$age = ($now - $entry.LastSeen).TotalMinutes
				if ($age -gt $thresholdMinutes) { $keysToRemove.Add($key) }
			}
			else { $keysToRemove.Add($key) }
		}
	}
	foreach ($key in $keysToRemove) { if ($script:ProcessCache.ContainsKey($key)) { $script:ProcessCache.Remove($key) } }
}

function GetProcessProfile
{
	[CmdletBinding()]
	[OutputType([string])]
	param([System.Diagnostics.Process]$Process)
	[string]$matchedProfile = ''; [string]$processExePath = ''
	if ('Custom.Native' -as [Type]) { try { $processExePath = [Custom.Native]::GetProcessPathById($Process.Id) } catch {} }
	if ([string]::IsNullOrEmpty($processExePath)) { try { if ($Process.MainModule) { $processExePath = $Process.MainModule.FileName } } catch {} }
	if ([string]::IsNullOrEmpty($processExePath)) { return $matchedProfile }

	if (-not [string]::IsNullOrEmpty($processExePath) -and $global:DashboardConfig -and $global:DashboardConfig.Config -and
		$global:DashboardConfig.Config.Contains('Profiles') -and $global:DashboardConfig.Config.Profiles)
	{
		$processDirectory = [System.IO.Path]::GetDirectoryName($processExePath)
		$normalizedProcessDirectory = $processDirectory.TrimEnd('\/') + '\'
		foreach ($profileEntry in $global:DashboardConfig.Config.Profiles.GetEnumerator())
		{
			[string]$profileName = $profileEntry.Key; [string]$profilePath = $profileEntry.Value
			$normalizedProfilePath = $profilePath.TrimEnd('\/') + '\'
			if ($normalizedProcessDirectory.StartsWith($normalizedProfilePath, [System.StringComparison]::OrdinalIgnoreCase))
			{
				$matchedProfile = $profileName; break
			}
		}
	}
	return $matchedProfile
}

function SetWindowToolStyle
{
	param([IntPtr]$hWnd, [bool]$Hide = $true)
	$GWL_EXSTYLE = -20; $WS_EX_TOOLWINDOW = 0x00000080; $WS_EX_APPWINDOW = 0x00040000
	try
	{
		if ('Custom.TaskbarTool' -as [Type]) { $null = [Custom.TaskbarTool]::SetTaskbarState($hWnd, -not $Hide) }
		[IntPtr]$currentExStylePtr = [Custom.Native]::GetWindowLongPtr($hWnd, $GWL_EXSTYLE)
		[long]$currentExStyle = $currentExStylePtr.ToInt64(); $newExStyle = $currentExStyle
		if ($Hide) { $newExStyle = ($currentExStyle -bor $WS_EX_TOOLWINDOW) -band (-bnot $WS_EX_APPWINDOW) }
		else { $newExStyle = ($currentExStyle -band (-bnot $WS_EX_TOOLWINDOW)) -bor $WS_EX_APPWINDOW }
		if ($newExStyle -ne $currentExStyle)
		{
			$isMinimized = [Custom.Native]::IsMinimized($hWnd)
			[Custom.Native]::ShowWindow($hWnd, [Custom.Native]::SW_HIDE)
			[IntPtr]$newExStylePtr = [IntPtr]$newExStyle
			[Custom.Native]::SetWindowLongPtr($hWnd, $GWL_EXSTYLE, $newExStylePtr) | Out-Null
			if ($isMinimized) { [Custom.Native]::ShowWindow($hWnd, [Custom.Native]::SW_SHOWMINNOACTIVE) }
			else { [Custom.Native]::ShowWindow($hWnd, [Custom.Native]::SW_SHOWNA) }
		}
	}
 catch {}
}

function Save-WindowPositions
{
	[CmdletBinding()]
	param()
	
	$grid = $global:DashboardConfig.UI.DataGridMain
	if (-not $grid) { return }
	
	$selectedRows = $grid.SelectedRows
	if ($selectedRows.Count -eq 0) { return }

	$savedCount = 0
	foreach ($row in $selectedRows)
	{
		if ($row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
		{
			$hWnd = $row.Tag.MainWindowHandle
			if ([Custom.Native]::IsMinimized($hWnd)) { continue }

			$rect = New-Object Custom.Native+RECT
			if ([Custom.Native]::GetWindowRect($hWnd, [ref]$rect))
			{
				if ($rect.Left -le -32000 -or $rect.Top -le -32000) { continue }

				$width = $rect.Right - $rect.Left
				$height = $rect.Bottom - $rect.Top
				$posString = "$($rect.Left),$($rect.Top),$width,$height"
				$identity = $row.Cells[1].Value.ToString()
				$global:DashboardConfig.Config['SavedWindowPositions'][$identity] = $posString
				$savedCount++
			}
		}
	}
	
	if ($savedCount -gt 0)
	{
		if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
	}
	if (Get-Command Show-DarkMessageBox -ErrorAction SilentlyContinue -Verbose:$False) { Show-DarkMessageBox "Window positions saved for $($savedCount) of $($selectedRows.Count) client(s).`nMinimized windows were skipped." "Position Saved" "OK" "Information" "success" }
}

function Restore-WindowPositions
{
	[CmdletBinding()]
	param()

	$grid = $global:DashboardConfig.UI.DataGridMain
	if (-not $grid -or -not $global:DashboardConfig.Config.Contains('SavedWindowPositions')) { return }
	
	foreach ($row in $grid.SelectedRows)
	{
		$identity = $row.Cells[1].Value.ToString()
		if ($global:DashboardConfig.Config['SavedWindowPositions'].Contains($identity))
		{
			$pos = $global:DashboardConfig.Config['SavedWindowPositions'][$identity] -split ','
			if ($pos.Count -eq 4 -and $row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
			{
				[Custom.Native]::PositionWindow($row.Tag.MainWindowHandle, [IntPtr]::Zero, [int]$pos[0], [int]$pos[1], [int]$pos[2], [int]$pos[3], 0x0014) 
			}
		}
	}
}
#endregion

#region Core Functions
function UpdateDataGrid
{
	[CmdletBinding()]
	param([System.Windows.Forms.DataGridView]$Grid)
	if (-not (TestValidParameters -Grid $Grid)) { return }
	$Grid.SuspendLayout()
	try
	{
		$currentProcesses = GetProcessList
		RemoveTerminatedProcesses -Grid $Grid -CurrentProcesses $currentProcesses
		if ($currentProcesses.Count -eq 0) { return }

		$rowLookup = NewRowLookupDictionary -Grid $Grid
		$processedProcesses = [System.Collections.Generic.List[PSObject]]::new()

		foreach ($process in $currentProcesses)
		{
			$pidVal = $process.Id
			if (-not $script:ProcessCache.ContainsKey($pidVal)) { $script:ProcessCache[$pidVal] = @{} }

			
			$cachedHWnd = [IntPtr]::Zero
			if ($script:ProcessCache[$pidVal].ContainsKey('hWnd')) { $cachedHWnd = $script:ProcessCache[$pidVal].hWnd }
			$lastWindowCheck = [DateTime]::MinValue
			if ($script:ProcessCache[$pidVal].ContainsKey('LastWindowCheck')) { $lastWindowCheck = $script:ProcessCache[$pidVal]['LastWindowCheck'] }
			if ($cachedHWnd -eq [IntPtr]::Zero -or -not [Custom.Native]::IsWindow($cachedHWnd))
			{
				if ([DateTime]::Now -gt $lastWindowCheck.AddSeconds(5))
				{
					if ('Custom.SafeWindowCore' -as [Type])
					{
						$cachedHWnd = [Custom.SafeWindowCore]::FindBestWindow($pidVal)
						$script:ProcessCache[$pidVal]['hWnd'] = $cachedHWnd
						$script:ProcessCache[$pidVal]['LastWindowCheck'] = [DateTime]::Now
					}
				}
			}
			$processTitle = $process.ProcessName
			$shouldUpdateTitle = $true
			if ($script:ProcessCache[$pidVal].ContainsKey('CachedTitle') -and $script:ProcessCache[$pidVal].ContainsKey('LastTitleCheck'))
			{
				if ([DateTime]::Now -lt $script:ProcessCache[$pidVal]['LastTitleCheck'].AddSeconds(2))
				{
					$shouldUpdateTitle = $false; $processTitle = $script:ProcessCache[$pidVal]['CachedTitle']
				}
			}
			if ($cachedHWnd -ne [IntPtr]::Zero -and $shouldUpdateTitle)
			{
				if ('Custom.SafeWindowCore' -as [Type])
				{
					$safeTitle = [Custom.SafeWindowCore]::GetText($cachedHWnd)
					if (-not [string]::IsNullOrEmpty($safeTitle))
					{
						$processTitle = $safeTitle; $script:ProcessCache[$pidVal]['CachedTitle'] = $safeTitle
					}
					$script:ProcessCache[$pidVal]['LastTitleCheck'] = [DateTime]::Now
				}
			}
			elseif ($cachedHWnd -ne [IntPtr]::Zero -and -not $shouldUpdateTitle)
			{
				if ($script:ProcessCache[$pidVal].ContainsKey('CachedTitle')) { $processTitle = $script:ProcessCache[$pidVal]['CachedTitle'] }
			}
			$processProfile = $null
			if ($script:ProcessCache[$pidVal].ContainsKey('Profile')) { $processProfile = $script:ProcessCache[$pidVal]['Profile'] }
			else { $processProfile = GetProcessProfile -Process $process; $script:ProcessCache[$pidVal]['Profile'] = $processProfile }
			$pStartTime = [DateTime]::MinValue
			if ($script:ProcessCache[$pidVal].ContainsKey('StartTime')) { $pStartTime = $script:ProcessCache[$pidVal]['StartTime'] }
			else { try { $pStartTime = $process.StartTime; $script:ProcessCache[$pidVal]['StartTime'] = $pStartTime } catch {} }

			$processedProcesses.Add([PSCustomObject]@{ Process = $process; ProcessTitle = $processTitle; ProcessProfile = $processProfile; StartTime = $pStartTime; CachedHWnd = $cachedHWnd })
		}

		
		$processedProcesses.Sort({
				param($a, $b)
				$res = [string]::Compare($a.ProcessProfile, $b.ProcessProfile, [StringComparison]::OrdinalIgnoreCase)
				if ($res -ne 0) { return $res }
				return [DateTime]::Compare($a.StartTime, $b.StartTime)
			})
		$sortedProcesses = $processedProcesses

		$processIndex = 0
		$previousProfile = $null; $previousRow = $null

		foreach ($processedProcess in $sortedProcesses)
		{
			$processIndex++
			$process = $processedProcess.Process; $processTitle = $processedProcess.ProcessTitle; $processProfile = $processedProcess.ProcessProfile; $hWnd = $processedProcess.CachedHWnd

			if (-not [string]::IsNullOrEmpty($processProfile)) { $processTitle = "[$processProfile]$processTitle" }
			if ($script:ProcessCache.ContainsKey($process.Id)) { $script:ProcessCache[$process.Id].LastSeen = [DateTime]::Now }

			$existingRow = $null
			if ($rowLookup.ContainsKey($process.Id)) { $existingRow = $rowLookup[$process.Id] }

			if ($existingRow) { UpdateExistingRow -Row $existingRow -Index $processIndex -ProcessTitle $processTitle -ProcessId $process.Id -Process $process }
			else { $newRow = AddNewProcessRow -Grid $Grid -Process $process -Index $processIndex -ProcessTitle $processTitle; $existingRow = $newRow; if ($newRow) { $rowLookup[$process.Id] = $newRow } }

			
			if ($existingRow)
			{
				if ($existingRow.DividerHeight -ne 0) { $existingRow.DividerHeight = 0 }
				if ($null -ne $previousRow -and $processProfile -ne $previousProfile) { $previousRow.DividerHeight = 4 }

				

				
				$isDisconnected = ($global:DashboardConfig.State.FlashingPids -and $global:DashboardConfig.State.FlashingPids.ContainsKey($process.Id))

				
				$isWindowFlashing = $false
				if (-not $isDisconnected -and $hWnd -ne [IntPtr]::Zero -and ('Custom.SafeWindowCore' -as [Type]))
				{
					$isWindowFlashing = [Custom.SafeWindowCore]::IsWindowFlashing($hWnd)
				}

				if ($isDisconnected)
				{
					
					if ($existingRow.DefaultCellStyle.BackColor -ne [System.Drawing.Color]::DarkRed)
					{
						$existingRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::DarkRed
						$existingRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
						$existingRow.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::DarkRed
						$existingRow.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
					}
				}
				elseif ($isWindowFlashing)
				{
					
					if ($existingRow.DefaultCellStyle.BackColor -ne [System.Drawing.Color]::Orange)
					{
						$existingRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::Orange
						$existingRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
						$existingRow.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::DarkOrange
						$existingRow.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
					}
				}
				else
				{
					

					if ($existingRow.DefaultCellStyle.BackColor -ne [System.Drawing.Color]::Empty)
					{
						$existingRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::Empty
						$existingRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Empty
						$existingRow.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::Empty
						$existingRow.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Empty
					}
				}

				$previousProfile = $processProfile
				$previousRow = $existingRow
				StartWindowStateCheck -Process $process -Row $existingRow -Grid $Grid -CachedHandle $hWnd
			}
		}
		UpdateRowIndices -Grid $Grid
		ClearOldProcessCache
	}
	finally { $Grid.ResumeLayout() }
}

function StartDataGridUpdateTimer
{
    [CmdletBinding()]
    param()

    if ($global:DashboardConfig.Resources.LaunchResources['DataGridUpdater'] -and 
        $global:DashboardConfig.Resources.LaunchResources['DataGridUpdater'].PowerShell.InvocationStateInfo.State -eq 'Running')
    {
        Write-Verbose "DATAGRID: Background update timer is already running."
        return
    }

    if ($global:DashboardConfig.Resources.LaunchResources['DataGridUpdater'])
    {
        try { 
            $global:DashboardConfig.Resources.LaunchResources['DataGridUpdater'].PowerShell.Stop()
            $global:DashboardConfig.Resources.LaunchResources['DataGridUpdater'].PowerShell.Dispose()
        } catch {}
        $global:DashboardConfig.Resources.LaunchResources.Remove('DataGridUpdater')
    }
    if ($global:DashboardConfig.Resources.Timers['dataGridUpdateTimer'])
    {
        try {
            $global:DashboardConfig.Resources.Timers['dataGridUpdateTimer'].Stop()
            $global:DashboardConfig.Resources.Timers['dataGridUpdateTimer'].Dispose()
        } catch {}
        $global:DashboardConfig.Resources.Timers.Remove('dataGridUpdateTimer')
    }

    $grid = $global:DashboardConfig.UI.DataGridMain
    if (-not $grid) { return }
    
    
    $prop = $grid.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance, NonPublic')
    if ($prop) { $prop.SetValue($grid, $true, $null) }

    
    
    
    $updateAction = [Action]{
        if ($global:DashboardConfig.UI.DataGridMain -and -not $global:DashboardConfig.UI.DataGridMain.IsDisposed) {
            try {
                UpdateDataGrid -Grid $global:DashboardConfig.UI.DataGridMain
            } catch {
                Write-Verbose "Error in timer-invoked UpdateDataGrid: $_"
            }
        }
    }

    
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = "STA"
    $rs.ThreadOptions = "ReuseThread"
    $rs.Open()
    
    
    $rs.SessionStateProxy.SetVariable('TargetGrid', $grid)
    $rs.SessionStateProxy.SetVariable('UpdateAction', $updateAction)
    $rs.SessionStateProxy.SetVariable('Interval', $script:UpdateInterval)

    
    
    $pulseScript = {
        while ($true)
        {
            if (-not $TargetGrid -or $TargetGrid.IsDisposed) { break }
            
            try {
                
                
                $TargetGrid.BeginInvoke($UpdateAction) | Out-Null
            }
            catch {
                break
            }
            
            Start-Sleep -Milliseconds $Interval
        }
    }
    
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    $ps.AddScript($pulseScript) | Out-Null
    
    $handle = $ps.BeginInvoke()
    
    if (-not $global:DashboardConfig.Resources.ContainsKey('LaunchResources')) {
        $global:DashboardConfig.Resources['LaunchResources'] = @{}
    }
    $global:DashboardConfig.Resources.LaunchResources['DataGridUpdater'] = @{
        PowerShell = $ps
        Runspace = $rs
        Handle = $handle
    }
    Write-Verbose "DATAGRID: Background pulse timer started."
}
#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion
