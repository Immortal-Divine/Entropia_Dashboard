<# datagrid.psm1 #>

#region Global Configuration
$script:UpdateInterval = 750

$script:ResponsivenessProbeTimeout = 20

$script:States = @{
	Checking  = 'Checking...'
	Loading   = 'Loading...'
	Minimized = 'Min'
	Normal    = 'Normal'
}

$script:ProcessCache = @{}

$script:PendingChecks = [System.Collections.Generic.HashSet[int]]::new()

$script:UpdatingGrid = $false

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

function RemoveTerminatedProcesses
{
	[CmdletBinding()]
	param([System.Windows.Forms.DataGridView]$Grid, [object[]]$CurrentProcesses = @())
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
	param([System.Windows.Forms.DataGridViewRow]$Row, [int]$Index, [string]$ProcessTitle, [int]$ProcessId, [object]$Process)
	if ($Row.Cells[0].Value -ne $Index) { $Row.Cells[0].Value = $Index }
	if ($Row.Cells[1].Value -ne $ProcessTitle) { $Row.Cells[1].Value = $ProcessTitle }
	if ($Row.Cells[2].Value -ne $ProcessId) { $Row.Cells[2].Value = $ProcessId }
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
			if ($null -eq $row.Cells -or $row.Cells.Count -lt 2) { continue }
			$processTitle = $row.Cells[1].Value
			$profileName = ''
			if ($processTitle -is [string])
			{
				$start = $processTitle.IndexOf('[')
				if ($start -eq 0)
				{
					$end = $processTitle.IndexOf(']', 1)
					if ($end -gt 1) { $profileName = $processTitle.Substring(1, $end - 1) }
				}
			}
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
	param([System.Windows.Forms.DataGridView]$Grid, [object]$Process, [int]$Index, [string]$ProcessTitle)
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

function Test-ClientWindowHidden
{
	[CmdletBinding()]
	[OutputType([bool])]
	param([string]$RowTitle)

	try
	{
		$config = $global:DashboardConfig.Config
		if (-not $config -or [string]::IsNullOrWhiteSpace($RowTitle)) { return $false }

		$profileName = $RowTitle
		$clientName = ''
		if ($RowTitle -match '^\[(?<pName>[^\]]+)\](?<rest>.*)')
		{
			$profileName = $Matches['pName']
			$clientName = ([string]$Matches['rest']).Trim()
		}
		elseif ($RowTitle -match '\[(?<pName>.*?)\]')
		{
			$profileName = $Matches['pName']
			$clientName = ([string]($RowTitle -replace '^\[.*?\]', '')).Trim()
		}
		else
		{
			return $false
		}

		if (-not [string]::IsNullOrWhiteSpace($clientName) -and $config.Contains('RememberedClients'))
		{
			$section = $config['RememberedClients']
			if ($section)
			{
				foreach ($key in @($section.Keys))
				{
					$parts = [string]$key -split '\|', 2
					if ($parts.Count -ne 2) { continue }
					if (-not ($parts[0] -ieq $profileName -and $parts[1] -ieq $clientName)) { continue }

					$tokens = @(([string]$section[$key]) -split ',')
					if ($tokens.Count -gt 4) { return ($tokens[4].Trim() -eq '1') }
					break
				}
			}
		}

		return (Test-HideProfilesDefault)
	}
	catch { }

	return $false
}

function StartWindowStateCheck
{
	[CmdletBinding()]
	param([object]$Process, [System.Windows.Forms.DataGridViewRow]$Row, [System.Windows.Forms.DataGridView]$Grid, [IntPtr]$CachedHandle = [IntPtr]::Zero)

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
					if ($null -ne $targetRow.Cells -and $targetRow.Cells.Count -gt 1)
					{
						$shouldHide = Test-ClientWindowHidden -RowTitle ([string]$targetRow.Cells[1].Value)
					}
					if ($shouldHide)
					{
						if ($isMinimized) { SetWindowToolStyle -hWnd $hWnd -Hide $true }
						else { SetWindowToolStyle -hWnd $hWnd -Hide $false }
					}
					else { SetWindowToolStyle -hWnd $hWnd -Hide $false }
				}

				$currentState = $targetRow.Cells[3].Value
				
				if ($currentState -in @('Error') -or $currentState -like 'Client *')
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

function Get-DataGridState
{
	[CmdletBinding()]
	[OutputType([System.Collections.Specialized.OrderedDictionary])]
	param()

	$state = [ordered]@{}
	foreach ($processId in @($script:ProcessCache.Keys))
	{
		$entry = $script:ProcessCache[$processId]
		if ($null -eq $entry) { continue }
		$state[$processId] = @{
			hWnd          = $entry['hWnd']
			LastSeen      = $entry['LastSeen']
			LastFlashSeen = $entry['LastFlashSeen']
		}
	}
	return $state
}

function Sync-DataGridState
{
	[CmdletBinding()]
	[OutputType([int])]
	param()

	$grid = $global:DashboardConfig.UI.DataGridMain
	$live = [System.Collections.Generic.HashSet[int]]::new()
	if ($grid -and -not $grid.IsDisposed)
	{
		foreach ($row in $grid.Rows)
		{
			try
			{
				if ($row.Tag -and $null -ne $row.Tag.Id) { [void]$live.Add([int]$row.Tag.Id) }
			}
			catch {}
		}
	}

	$removed = 0
	foreach ($processId in @($script:ProcessCache.Keys))
	{
		$id = 0
		if (-not [int]::TryParse([string]$processId, [ref]$id)) { continue }
		if (-not $live.Contains($id))
		{
			$script:ProcessCache.Remove($processId)
			$removed++
		}
	}
	return $removed
}

function GetProcessProfile
{
	[CmdletBinding()]
	[OutputType([string])]
	param($Process)
	[string]$matchedProfile = ''; [string]$processExePath = ''
	if ('Custom.Native' -as [Type]) { try { $processExePath = [Custom.Native]::GetProcessPathById($Process.Id) } catch {} }
	if ([string]::IsNullOrEmpty($processExePath) -and $Process.PSObject.Properties['MainModule']) { try { if ($Process.MainModule) { $processExePath = $Process.MainModule.FileName } } catch {} }
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
				if ($null -ne $row.Cells -and $row.Cells.Count -gt 1 -and $null -ne $row.Cells[1].Value)
				{
					$identity = $row.Cells[1].Value.ToString()
					$global:DashboardConfig.Config['SavedWindowPositions'][$identity] = $posString
					$savedCount++
				}
			}
		}
	}
	
	if ($savedCount -gt 0)
	{
		if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
	}
	if (Test-DashboardCommand 'Show-DarkMessageBox') { Show-DarkMessageBox "Window positions saved for $($savedCount) of $($selectedRows.Count) client(s).`nMinimized windows were skipped." "Position Saved" "OK" "Information" "success" }
}

function Restore-WindowPositions
{
	[CmdletBinding()]
	param()

	$grid = $global:DashboardConfig.UI.DataGridMain
	if (-not $grid -or -not $global:DashboardConfig.Config.Contains('SavedWindowPositions')) { return }
	
	foreach ($row in $grid.SelectedRows)
	{
		if ($null -ne $row.Cells -and $row.Cells.Count -gt 1 -and $null -ne $row.Cells[1].Value)
		{
			$identity = $row.Cells[1].Value.ToString()
			if ($global:DashboardConfig.Config['SavedWindowPositions'].Contains($identity))
			{
				$pos = $global:DashboardConfig.Config['SavedWindowPositions'][$identity] -split ','
				if ($pos.Count -eq 4 -and $row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
				{
					if (-not [Custom.Native]::IsMinimized($row.Tag.MainWindowHandle))
					{
						[Custom.Native]::PositionWindow($row.Tag.MainWindowHandle, [IntPtr]::Zero, [int]$pos[0], [int]$pos[1], [int]$pos[2], [int]$pos[3], 0x0014) 
					}
				}
			}
		}
	}
}

function Reapply-GridStatusStyling
{
	[CmdletBinding()]
	param()

	$grid = $global:DashboardConfig.UI.DataGridMain
	if (-not $grid -or $grid.IsDisposed) { return }
	if (-not $global:DashboardConfig.State -or -not $global:DashboardConfig.State.UIInitialized) { return }

	try
	{
			$now = [DateTime]::Now

		$cDisconnected = Get-Theme 'StatusDisconnected'
		$cDisconnectedText = Get-Theme 'StatusDisconnectedText'
		$cFlashing = Get-Theme 'StatusFlashing'
		$cFlashingSelect = Get-Theme 'StatusFlashingSelect'
		$cWhiteText = Get-Theme 'WhiteText'
		$cEmpty = Get-Theme 'Empty'

		foreach ($row in $grid.Rows)
		{
			if (-not $row.Tag -or $null -eq $row.Tag.Id) { continue }

			$pid = $row.Tag.Id
			$isDisconnected = ($global:DashboardConfig.State.FlashingPids -and $global:DashboardConfig.State.FlashingPids.ContainsKey($pid))
			$isWindowFlashing = $false

			if (-not $isDisconnected -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero -and ('Custom.SafeWindowCore' -as [Type]))
			{
				$isWindowFlashing = [Custom.SafeWindowCore]::IsWindowFlashing($row.Tag.MainWindowHandle)
				if ($isWindowFlashing)
				{
					if (-not $script:ProcessCache.ContainsKey($pid)) { $script:ProcessCache[$pid] = @{} }
					$script:ProcessCache[$pid].LastFlashSeen = $now
				}
				elseif ($script:ProcessCache.ContainsKey($pid) -and $script:ProcessCache[$pid].LastFlashSeen -and ($now - $script:ProcessCache[$pid].LastFlashSeen).TotalSeconds -lt 2.5)
				{
					$isWindowFlashing = $true
				}
			}

			if ($isDisconnected)
			{
				$row.DefaultCellStyle.BackColor = $cDisconnected
				$row.DefaultCellStyle.ForeColor = $cDisconnectedText
				$row.DefaultCellStyle.SelectionBackColor = $cDisconnected
				$row.DefaultCellStyle.SelectionForeColor = $cWhiteText
			}
			elseif ($isWindowFlashing)
			{
				$row.DefaultCellStyle.BackColor = $cFlashing
				$row.DefaultCellStyle.ForeColor = $cDisconnectedText
				$row.DefaultCellStyle.SelectionBackColor = $cFlashingSelect
				$row.DefaultCellStyle.SelectionForeColor = $cWhiteText
			}
			else
			{
				$row.DefaultCellStyle.BackColor = $cEmpty
				$row.DefaultCellStyle.ForeColor = $cEmpty
				$row.DefaultCellStyle.SelectionBackColor = $cEmpty
				$row.DefaultCellStyle.SelectionForeColor = $cEmpty
			}
		}
	}
	catch { }
}
#endregion

#region Core Functions
function UpdateDataGrid
{
	[CmdletBinding()]
	param(
		[System.Windows.Forms.DataGridView]$Grid,
		[System.Collections.Generic.List[object]]$CurrentProcesses = $null
	)

	if ($script:UpdatingGrid) { return }
	$script:UpdatingGrid = $true
	try { UpdateDataGridCore -Grid $Grid -CurrentProcesses $CurrentProcesses }
	finally { $script:UpdatingGrid = $false }
}

function UpdateDataGridCore
{
	[CmdletBinding()]
	param(
		[System.Windows.Forms.DataGridView]$Grid,
		[System.Collections.Generic.List[object]]$CurrentProcesses = $null
	)
	if (-not (TestValidParameters -Grid $Grid)) { return }
	
	$processes = $CurrentProcesses
	if ($null -eq $processes)
	{
		$processName = $global:DashboardConfig.Config['ProcessName']['ProcessName']
		$profileMap = $global:DashboardConfig.Config.Profiles
		$processes = [Custom.ProcessHelper]::GetProcessDataList($processName, $profileMap)
	}

	$syncWork = [System.Collections.Generic.List[object]]::new()

	$cDisconnected = Get-Theme 'StatusDisconnected'
	$cDisconnectedText = Get-Theme 'StatusDisconnectedText'
	$cFlashing = Get-Theme 'StatusFlashing'
	$cFlashingSelect = Get-Theme 'StatusFlashingSelect'
	$cWhiteText = Get-Theme 'WhiteText'
	$cEmpty = Get-Theme 'Empty'

	$Grid.SuspendLayout()
	try
	{
		RemoveTerminatedProcesses -Grid $Grid -CurrentProcesses $processes
		if ($processes.Count -eq 0) { return }

		$rowLookup = NewRowLookupDictionary -Grid $Grid
		
		$profileCounters = @{}
		$previousProfile = $null; $previousRow = $null

		foreach ($process in $processes)
		{
			$pidVal = $process.Id
			if (-not $script:ProcessCache.ContainsKey($pidVal)) { $script:ProcessCache[$pidVal] = @{} }

			$processProfile = $process.Profile
			if ([string]::IsNullOrEmpty($processProfile)) { $processProfile = 'NoProfile' }
			if (-not $profileCounters.ContainsKey($processProfile)) { $profileCounters[$processProfile] = 0 }
			$profileCounters[$processProfile]++
			$perProfileIndex = $profileCounters[$processProfile]

			$processTitle = $process.Title
			if ($process.Profile) { $processTitle = "[$($process.Profile)]$processTitle" }
			if ($script:ProcessCache.ContainsKey($process.Id)) { $script:ProcessCache[$process.Id].LastSeen = [DateTime]::Now }

			$existingRow = $null
			if ($rowLookup.ContainsKey($process.Id)) { $existingRow = $rowLookup[$process.Id] }

			if ($existingRow) 
			{ 
				UpdateExistingRow -Row $existingRow -Index $perProfileIndex -ProcessTitle $processTitle -ProcessId $process.Id -Process $process
				$existingRow.Tag = $process
			}
			else 
			{ 
				$newRow = AddNewProcessRow -Grid $Grid -Process $process -Index $perProfileIndex -ProcessTitle $processTitle
				$existingRow = $newRow
				if ($newRow) { $rowLookup[$process.Id] = $newRow } 
			}

			if ($existingRow)
			{
				if ($existingRow.DividerHeight -ne 0) { $existingRow.DividerHeight = 0 }
				if ($null -ne $previousRow -and $process.Profile -ne $previousProfile) { $previousRow.DividerHeight = 4 }

				$isDisconnected = ($global:DashboardConfig.State.FlashingPids -and $global:DashboardConfig.State.FlashingPids.ContainsKey($process.Id))
				$isWindowFlashing = $false
				$cachedHWnd = $process.MainWindowHandle
				if (-not $isDisconnected -and $cachedHWnd -ne [IntPtr]::Zero -and ('Custom.SafeWindowCore' -as [Type]))
				{
					$isWindowFlashing = [Custom.SafeWindowCore]::IsWindowFlashing($cachedHWnd)

					if ($isWindowFlashing) { $script:ProcessCache[$process.Id].LastFlashSeen = [DateTime]::Now }
					elseif ($script:ProcessCache[$process.Id].LastFlashSeen -and ([DateTime]::Now - $script:ProcessCache[$process.Id].LastFlashSeen).TotalSeconds -lt 2.5) {
						$isWindowFlashing = $true
					}
				}

				if ($isDisconnected)
				{
					if ($existingRow.DefaultCellStyle.BackColor -ne $cDisconnected -or $existingRow.DefaultCellStyle.SelectionBackColor -ne $cDisconnected)
					{
						$existingRow.DefaultCellStyle.BackColor = $cDisconnected
						$existingRow.DefaultCellStyle.ForeColor = $cDisconnectedText
						$existingRow.DefaultCellStyle.SelectionBackColor = $cDisconnected
						$existingRow.DefaultCellStyle.SelectionForeColor = $cWhiteText
					}
				}
				elseif ($isWindowFlashing)
				{
					if ($existingRow.DefaultCellStyle.BackColor -ne $cFlashing -or $existingRow.DefaultCellStyle.SelectionBackColor -ne $cFlashingSelect)
					{
						$existingRow.DefaultCellStyle.BackColor = $cFlashing
						$existingRow.DefaultCellStyle.ForeColor = $cDisconnectedText
						$existingRow.DefaultCellStyle.SelectionBackColor = $cFlashingSelect
						$existingRow.DefaultCellStyle.SelectionForeColor = $cWhiteText
					}
				}
				else
				{
					$isHighlighted = ($existingRow.DefaultCellStyle.BackColor -eq $cDisconnected -or 
									  $existingRow.DefaultCellStyle.BackColor -eq $cFlashing -or
									  $existingRow.DefaultCellStyle.SelectionBackColor -eq $cDisconnected -or
									  $existingRow.DefaultCellStyle.SelectionBackColor -eq $cFlashingSelect)

					if ($isHighlighted)
					{
						$script:ProcessCache[$process.Id].LastFlashSeen = $null
						$existingRow.DefaultCellStyle.BackColor = $cEmpty
						$existingRow.DefaultCellStyle.ForeColor = $cEmpty
						$existingRow.DefaultCellStyle.SelectionBackColor = $cEmpty
						$existingRow.DefaultCellStyle.SelectionForeColor = $cEmpty
					}
				}

				$hWnd = $process.MainWindowHandle

				if ($hWnd -ne [IntPtr]::Zero) { $syncWork.Add(@{ Row = $existingRow; hWnd = $hWnd }) }

				$previousProfile = $process.Profile
				$previousRow = $existingRow
				StartWindowStateCheck -Process $process -Row $existingRow -Grid $Grid -CachedHandle $hWnd
			}
		}
		UpdateRowIndices -Grid $Grid
		ClearOldProcessCache
		Sync-DataGridState | Out-Null
	}
	finally { $Grid.ResumeLayout() }

	foreach ($item in $syncWork)
	{
		$row = $item.Row
		$hWnd = $item.hWnd
		if ($Grid.IsDisposed -or $null -eq $row -or $row.DataGridView -ne $Grid) { continue }

		try
		{
			$isResponsiveNow = [Custom.Native]::Responsive($hWnd, $script:ResponsivenessProbeTimeout)

			$isMinSync = $false
			if ($isResponsiveNow)
			{
				if ('Custom.SafeWindowCore' -as [Type]) { $isMinSync = [Custom.SafeWindowCore]::IsMinimized($hWnd) }
				else { $isMinSync = [Custom.Native]::IsMinimized($hWnd) }
			}

			if ($row.Cells.Count -gt 3)
			{
				$syncState = if (-not $isResponsiveNow) { $script:States.Loading }
				elseif ($isMinSync) { $script:States.Minimized }
				else { $script:States.Normal }
				$existingState = $row.Cells[3].Value
				if ($existingState -ne $syncState -and $existingState -notin @('Error') -and $existingState -notlike 'Client *')
				{
					$row.Cells[3].Value = $syncState
				}
			}

			if (-not $isResponsiveNow) { continue }

			$shouldHide = $false
			if ($null -ne $row.Cells -and $row.Cells.Count -gt 1)
			{
				$shouldHide = Test-ClientWindowHidden -RowTitle ([string]$row.Cells[1].Value)
			}

			if ($shouldHide) { SetWindowToolStyle -hWnd $hWnd -Hide $isMinSync }
			else { SetWindowToolStyle -hWnd $hWnd -Hide $false }
		}
		catch { }
	}
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

    $grid = $global:DashboardConfig.UI.DataGridMain
    if (-not $grid) { return }
    
    $prop = $grid.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance, NonPublic')
    if ($prop) { $prop.SetValue($grid, $true, $null) }

    $updateAction = [System.Action[System.Object]]{
        param($data)
        if ($global:DashboardConfig.UI.DataGridMain -and -not $global:DashboardConfig.UI.DataGridMain.IsDisposed) {
            try {
                UpdateDataGrid -Grid $global:DashboardConfig.UI.DataGridMain -CurrentProcesses $data
            } catch {
                Write-Verbose "Error in timer-invoked UpdateDataGrid: $_"
            }
        }
    }

    $processName = $global:DashboardConfig.Config['ProcessName']['ProcessName']
    $profileMap = $global:DashboardConfig.Config.Profiles

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = "STA"
    $rs.ThreadOptions = "ReuseThread"
    $rs.Open()
    
    $rs.SessionStateProxy.SetVariable('TargetGrid', $grid)
    $rs.SessionStateProxy.SetVariable('UpdateAction', $updateAction)
    $rs.SessionStateProxy.SetVariable('Interval', $script:UpdateInterval)
    $rs.SessionStateProxy.SetVariable('ProcessName', $processName)
    $rs.SessionStateProxy.SetVariable('ProfileMap', $profileMap)

    $pulseScript = {
        while ($true)
        {
            if (-not $TargetGrid -or $TargetGrid.IsDisposed) { break }
            try {
                $data = [Custom.ProcessHelper]::GetProcessDataList($ProcessName, $ProfileMap)
                
                $TargetGrid.BeginInvoke($UpdateAction, $data) | Out-Null
            }
            catch { break }
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