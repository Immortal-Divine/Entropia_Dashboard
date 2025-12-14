<# datagrid.psm1 #>

#region Global Configuration
    #region Step: Define Update Interval
        $script:UpdateInterval = 100
    #endregion Step: Define Update Interval

    #region Step: Define Process Window States
        $script:States = @{
            Checking  = 'Checking...'
            Loading   = 'Loading...'
            Minimized = 'Min'
            Normal    = 'Normal'
        }
    #endregion Step: Define Process Window States

    #region Step: Initialize Process Cache
        $script:ProcessCache = @{}
    #endregion Step: Initialize Process Cache

    #region Step: Initialize Last Log Message Cache
        $script:LastLogMessages = @{}
    #endregion Step: Initialize Last Log Message Cache
#endregion Global Configuration

#region Helper Functions
    #region Function: Test-ValidParameters
        function Test-ValidParameters
        {
            [CmdletBinding()]
            [OutputType([bool])]
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridView]$Grid
            )

            #region Step: Validate Grid Parameter
                if (-not $Grid)
                {
                    return $false
                }
            #endregion Step: Validate Grid Parameter

            #region Step: Validate Global Configuration Existence and Structure
                if (-not $global:DashboardConfig -or -not $global:DashboardConfig.Config -or -not $global:DashboardConfig.Config.Contains('ProcessName'))
                {
                    Write-Verbose 'DATAGRID: Global settings missing. Define DashboardConfig before importing.' -ForegroundColor Red
                    return $false
                }
            #endregion Step: Validate Global Configuration Existence and Structure

            #region Step: Validate Process Name Configuration
                $processName = $global:DashboardConfig.Config['ProcessName']['ProcessName']
                if ([string]::IsNullOrEmpty($processName))
                {
                    return $false
                }
            #endregion Step: Validate Process Name Configuration

            #region Step: Return Validation Success
                return $true
            #endregion Step: Return Validation Success
        }
    #endregion Function: Test-ValidParameters

    #region Function: Get-ProcessList
        function Get-ProcessList
        {
            [CmdletBinding()]
            [OutputType([System.Diagnostics.Process[]])]
            param()

            #region Step: Get Configured Process Name
                $processName = $global:DashboardConfig.Config['ProcessName']['ProcessName']
            #endregion Step: Get Configured Process Name

            #region Step: Retrieve and Filter Processes Safely
                try
                {
                    #region Step: Define Parameters for Get-Process
                        $getProcessParams = @{
                            Name        = $processName
                            ErrorAction = 'SilentlyContinue'
                        }
                    #endregion Step: Define Parameters for Get-Process

                    #region Step: Get All Matching Processes
                        $allProcesses = Get-Process @getProcessParams
                    #endregion Step: Get All Matching Processes

                    #region Step: Filter for Valid (Non-Exited) Processes
                        $validProcesses = @()
                        foreach ($proc in $allProcesses)
                        {
                            try
                            {
                                # Check if process is still running by accessing HasExited property
                                $exited = $proc.HasExited
                                if (-not $exited)
                                {
                                    $validProcesses += $proc
                                }
                            }
                            catch
                            {
                                # Skip this process if we can't access its properties (e.g., access denied, already exited)
                                continue
                            }
                        }
                    #endregion Step: Filter for Valid (Non-Exited) Processes

                    #region Step: Filter for Processes with Windows and Sort Safely
                        $processes = @()
                        $windowProcesses = @()
                        try
                        {
                            # Filter for processes with windows
                            $windowProcesses = @($validProcesses |
                            Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowHandle -ne [IntPtr]::Zero })

                            # Sort by StartTime with error handling
                            $processes = @($windowProcesses | Sort-Object {
                                    try
                                    {
                                        $_.StartTime
                                    }
                                    catch
                                    {
                                        # Return a default date for processes where StartTime can't be accessed
                                        [DateTime]::MinValue
                                    }
                                })

                            # If no processes with windows found, try to use all valid processes
                            if ($processes.Count -eq 0 -and $validProcesses.Count -gt 0)
                            {
                                # Sort valid processes with error handling
                                $processes = @($validProcesses | Sort-Object {
                                        try
                                        {
                                            $_.StartTime
                                        }
                                        catch
                                        {
                                            [DateTime]::MinValue
                                        }
                                    })
                            }
                        }
                        catch
                        {
                            # Fall back to unsorted list if sorting fails
                            $processes = if ($windowProcesses.Count -gt 0)
                            {
                                $windowProcesses
                            }
                            else
                            {
                                $validProcesses
                            }
                        }
                    #endregion Step: Filter for Processes with Windows and Sort Safely

                    #region Step: Return Process List
                        return $processes
                    #endregion Step: Return Process List
                }
                catch
                {
                    #region Step: Handle Process Retrieval Errors
                        Write-Verbose "DASHBOARD: Error getting process list: $_" -ForegroundColor Red
                        return @()
                    #endregion Step: Handle Process Retrieval Errors
                }
            #endregion Step: Retrieve and Filter Processes Safely
        }
    #endregion Function: Get-ProcessList

    #region Function: Remove-TerminatedProcesses
        function Remove-TerminatedProcesses
        {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridView]$Grid,

                [Parameter(Mandatory = $false)]
                [System.Diagnostics.Process[]]$CurrentProcesses = @()
            )

            #region Step: Handle Empty Grid
                if ($Grid.Rows.Count -eq 0)
                {
                    return
                }
            #endregion Step: Handle Empty Grid

            #region Step: Initialize List for Rows to Remove
                $rowsToRemove = [System.Collections.Generic.List[System.Windows.Forms.DataGridViewRow]]::new()
            #endregion Step: Initialize List for Rows to Remove

            #region Step: Create Process ID Lookup for Efficient Check
                $processIdLookup = @{}
                foreach ($process in $CurrentProcesses)
                {
                    # Ensure process object is valid before accessing Id
                    if ($null -ne $process -and $null -ne $process.Id) {
                        $processIdLookup[$process.Id] = $true
                    }
                }
            #endregion Step: Create Process ID Lookup for Efficient Check

            #region Step: Identify Rows Corresponding to Terminated Processes
                foreach ($row in $Grid.Rows)
                {
                    # Skip rows without valid process data in Tag
                    if (-not $row.Tag -or $null -eq $row.Tag.Id)
                    {
                        continue
                    }

                    # Check if process still exists in current process list using lookup
                    $processExists = $processIdLookup.ContainsKey($row.Tag.Id)

                    # If process doesn't exist, mark row for removal and clear from cache
                    if (-not $processExists)
                    {
                        $rowsToRemove.Add($row)

                        # Also remove from process cache
                        if ($script:ProcessCache.ContainsKey($row.Tag.Id))
                        {
                            $script:ProcessCache.Remove($row.Tag.Id)
                        }
                    }
                }
            #endregion Step: Identify Rows Corresponding to Terminated Processes

            #region Step: Remove Identified Rows from Grid
                # Remove rows in reverse order to avoid index shifting issues
                for ($i = $rowsToRemove.Count - 1; $i -ge 0; $i--)
                {
                    # Check if row still exists in grid before removing
                    if ($rowsToRemove[$i].Index -ge 0 -and $rowsToRemove[$i].Index -lt $Grid.Rows.Count) {
                       $Grid.Rows.Remove($rowsToRemove[$i])
                    }
                }
            #endregion Step: Remove Identified Rows from Grid
        }
    #endregion Function: Remove-TerminatedProcesses

    #region Function: New-RowLookupDictionary
        function New-RowLookupDictionary
        {
            [CmdletBinding()]
            [OutputType([hashtable])]
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridView]$Grid
            )

            #region Step: Initialize Lookup Hashtable
                $rowLookup = @{}
            #endregion Step: Initialize Lookup Hashtable

            #region Step: Populate Lookup Hashtable from Grid Rows
                foreach ($row in $Grid.Rows)
                {
                    # Only add rows with valid process data (Tag exists and has an Id)
                    if ($row.Tag -and $null -ne $row.Tag.Id)
                    {
                        $rowLookup[$row.Tag.Id] = $row
                    }
                }
            #endregion Step: Populate Lookup Hashtable from Grid Rows

            #region Step: Return Lookup Hashtable
                return $rowLookup
            #endregion Step: Return Lookup Hashtable
        }
    #endregion Function: New-RowLookupDictionary

    #region Function: Update-ExistingRow
        function Update-ExistingRow
		{
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridViewRow]$Row,

                [Parameter(Mandatory = $true)]
                [int]$Index,

                [Parameter(Mandatory = $true)]
                [string]$ProcessTitle,

                [Parameter(Mandatory = $true)]
                [int]$ProcessId,

                [Parameter(Mandatory = $true)]
                [System.Diagnostics.Process]$Process
            )

            #region Step: Update Index Cell (Column 0) if Changed
                if ($Row.Cells[0].Value -ne $Index)
                {
                    $Row.Cells[0].Value = $Index
                }
            #endregion Step: Update Index Cell (Column 0) if Changed

            #region Step: Update Title Cell (Column 1) if Changed
                if ($Row.Cells[1].Value -ne $ProcessTitle)
                {
                    $Row.Cells[1].Value = $ProcessTitle
                }
			#endregion Step: Update Index Cell (Column 0) if Changed

			#region Step: Update PID Cell (Column 2) if Changed
			if ($Row.Cells[2].Value -ne $ProcessId)
			{
				$Row.Cells[2].Value = $ProcessId
			}
			#endregion Step: Update PID Cell (Column 2) if Changed

            #region Step: Update Row Tag with Latest Process Object
                $Row.Tag = $Process
            #endregion Step: Update Row Tag with Latest Process Object
		}
    #endregion Function: Update-ExistingRow

    #region Function: UpdateRowIndices
        function UpdateRowIndices
        {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridView]$Grid
            )

            #region Step: Update Row Indices Safely
                try
                {
                    #region Step: Handle Empty Grid
                        if ($Grid.Rows.Count -eq 0)
                        {
                            return
                        }
                    #endregion Step: Handle Empty Grid

                    #region Step: Iterate Through Rows and Update Index Cell (Column 0)
                        for ($i = 0; $i -lt $Grid.Rows.Count; $i++)
                        {
                            $row = $Grid.Rows[$i]
                            $newIndex = $i + 1  # Start index numbering from 1

                            # Only update if the index has changed
                            if ($row.Cells[0].Value -ne $newIndex)
                            {
                                $row.Cells[0].Value = $newIndex
                            }
                        }
                    #endregion Step: Iterate Through Rows and Update Index Cell (Column 0)
                }
                catch
                {
                    #region Step: Handle Index Update Errors
                        Write-Verbose "DATAGRID: Failed to update row indices: $_" -ForegroundColor Red
                    #endregion Step: Handle Index Update Errors
                }
            #endregion Step: Update Row Indices Safely
        }
    #endregion Function: UpdateRowIndices

    #region Function: Add-NewProcessRow
        function Add-NewProcessRow
        {
            [CmdletBinding()]
            [OutputType([System.Windows.Forms.DataGridViewRow])]
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridView]$Grid,

                [Parameter(Mandatory = $true)]
                [System.Diagnostics.Process]$Process,

                [Parameter(Mandatory = $true)]
                [int]$Index,

                [Parameter(Mandatory = $true)]
                [string]$ProcessTitle
            )

            #region Step: Add Row and Update Cache Safely
                try
                {
                    #region Step: Add New Row to Grid with Initial Data
                        # Adds Index, Title, PID, and initial State ('Checking...')
                        $rowIndex = $Grid.Rows.Add($Index, $ProcessTitle, $Process.Id, $script:States.Checking)
                    #endregion Step: Add New Row to Grid with Initial Data

                    #region Step: Store Process Object in Row Tag
                        # Store process object in row's Tag property for later reference (e.g., state checks)
                        $Grid.Rows[$rowIndex].Tag = $Process
                    #endregion Step: Store Process Object in Row Tag

                    #region Step: Add Process to Cache
                        # Add to process cache with current timestamp and window status
                        $script:ProcessCache[$Process.Id] = @{
                            LastSeen  = Get-Date
                            HasWindow = ($Process.MainWindowHandle -ne 0 -and $Process.MainWindowHandle -ne [IntPtr]::Zero)
                        }
                    #endregion Step: Add Process to Cache

                    #region Step: Return Newly Added Row
                        return $Grid.Rows[$rowIndex]
                    #endregion Step: Return Newly Added Row
                }
                catch
                {
                    #region Step: Handle Row Addition Errors
                        Write-Verbose "DATAGRID: Error adding new process row: $_" -ForegroundColor Red
                        return $null
                    #endregion Step: Handle Row Addition Errors
                }
            #endregion Step: Add Row and Update Cache Safely
        }
    #endregion Function: Add-NewProcessRow

    #region Function: Start-WindowStateCheck
        function Start-WindowStateCheck
        {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.Diagnostics.Process]$Process,

                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridViewRow]$Row,

                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridView]$Grid
            )

            #region Step: Handle Processes Without Windows
                # If the process has no main window handle, assume 'Normal' state and skip async check.
                if ($Process.MainWindowHandle -eq 0 -or $Process.MainWindowHandle -eq [IntPtr]::Zero)
                {
                    # Check if row still exists and update state if needed
                    if ($Row -ne $null -and $Row.Index -ge 0 -and $Row.Cells[3].Value -ne $script:States.Normal) {
                       $Row.Cells[3].Value = $script:States.Normal
                    }
                    return
                }
            #endregion Step: Handle Processes Without Windows

            #region Step: Perform Asynchronous Window State Check
                try
                {
                    #region Step: Start Asynchronous Responsiveness Check
                        # Call native helper to check responsiveness without blocking
                        $task = [Custom.Native]::ResponsiveAsync($Process.MainWindowHandle)
                    #endregion Step: Start Asynchronous Responsiveness Check

                    #region Step: Prepare State Data for Callback
                        # Package necessary data to pass to the continuation (callback)
                        $stateData = @{
                            ProcessId = $Process.Id
                            RowIndex  = $Row.Index # Store index as row object might become invalid
                            hWnd      = $Process.MainWindowHandle # Use handle from process object directly
                        }
                    #endregion Step: Prepare State Data for Callback

                    #region Step: Get UI Thread Scheduler
                        # Get the scheduler associated with the current UI thread for safe updates
                        $scheduler = [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext()
                    #endregion Step: Get UI Thread Scheduler

                    #region Step: Define Continuation Action (Callback for UI Update)
                        # This action runs on the UI thread after the async task completes
                        $continuationAction = [System.Action[System.Threading.Tasks.Task, Object]] {
                            param($completedTask, $state) # Renamed stateData parameter to state

                            #region Step: Callback - Handle UI Update Safely
                                try
                                {
                                    #region Step: Callback - Get Context Data
                                        # Retrieve grid reference and state data within the callback context
                                        $currentGrid = $global:DashboardConfig.UI.DataGridFiller
                                        $processId = $state.ProcessId
                                        $rowIndex = $state.RowIndex
                                        $hWnd = $state.hWnd
                                    #endregion Step: Callback - Get Context Data

                                    #region Step: Callback - Validate Grid and Row Existence
                                        # Exit if grid is disposed or row index is invalid
                                        if (-not $currentGrid -or $currentGrid.IsDisposed -or $rowIndex -lt 0 -or $rowIndex -ge $currentGrid.Rows.Count)
                                        {
                                            return
                                        }
                                        # Find the target row using the stored index, verify it's the correct process
                                        $targetRow = $currentGrid.Rows[$rowIndex]
                                        if ($null -eq $targetRow -or $null -eq $targetRow.Tag -or $targetRow.Tag.Id -ne $processId) {
                                            # Row might have been removed or changed, try finding by ID as fallback
                                            $targetRow = Find-TargetRow -Grid $currentGrid -ProcessId $processId -RowIndex -1 # Use -1 to force search
                                            if ($null -eq $targetRow) {
                                                return # Row not found, cannot update
                                            }
                                        }
                                    #endregion Step: Callback - Validate Grid and Row Existence

                                    #region Step: Callback - Determine Window State
                                        # Check minimized state using native helper
                                        $isMinimized = [Custom.Native]::IsMinimized($hWnd)
                                        $isResponsive = $null

                                        # Get responsiveness result from the completed async task
                                        if ($completedTask.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion)
                                        {
                                            $isResponsive = $completedTask.Result
                                        }

                                        # Determine the final display state based on responsiveness and minimized status
                                        $finalState = if (-not $isResponsive)
                                        {
                                            $script:States.Loading # Unresponsive maps to Loading
                                        }
                                        elseif ($isMinimized)
                                        {
                                            $script:States.Minimized
                                        }
                                        else
                                        {
                                            $script:States.Normal
                                        }

                                         # Hide or show the window from Alt+Tab based on its minimized state and the setting
                                         if ($hWnd -ne 0 -and $hWnd -ne [IntPtr]::Zero)
                                         {
                                             $hideMinimizedOption = $global:DashboardConfig.Config['Options']['HideMinimizedWindows'] -eq '1'
                                            
                                             if ($hideMinimizedOption) {
                                                 # If hiding is enabled AND the window is minimized, apply tool style
                                                 if ($isMinimized) {
                                                     Set-WindowToolStyle -hWnd $hWnd -Hide $true
                                                 }
                                                 else {
                                                     # If hiding is enabled but window is NOT minimized, ensure it's visible (user restored it)
                                                     Set-WindowToolStyle -hWnd $hWnd -Hide $false
                                                 }
                                             }
                                             else {
                                                 # If hiding is NOT enabled, ensure it's visible regardless of minimized state
                                                 Set-WindowToolStyle -hWnd $hWnd -Hide $false
                                             }
                                         }
                                    #endregion Step: Callback - Determine Window State

                                    #region Step: Callback - Update Row State Cell if Changed
                                        # Update the state cell (column 3) only if the state has changed
                                        $currentState = $targetRow.Cells[3].Value
                                        if ($currentState -ne $finalState)
                                        {
                                            $targetRow.Cells[3].Value = $finalState
                                        }
                                    #endregion Step: Callback - Update Row State Cell if Changed
                                }
                                catch
                                {
                                    # Silent error handling for UI update callback to prevent crashes
                                }
                            #endregion Step: Callback - Handle UI Update Safely
                        }
                    #endregion Step: Define Continuation Action (Callback for UI Update)

                    #region Step: Schedule Continuation Action to Run on UI Thread
                        # Schedule the continuation action to run when the async task finishes,
                        # passing the state data and ensuring it executes on the UI thread via the scheduler.
                        $task.ContinueWith(
                            $continuationAction,
                            $stateData,
                            [System.Threading.CancellationToken]::None,
                            [System.Threading.Tasks.TaskContinuationOptions]::None,
                            $scheduler
                        )
                    #endregion Step: Schedule Continuation Action to Run on UI Thread
                }
                catch
                {
                    #region Step: Handle Asynchronous Check Initiation Errors
                        # If starting the async check fails, default the row state to 'Normal'
                         if ($Row -ne $null -and $Row.Index -ge 0 -and $Row.Cells[3].Value -ne $script:States.Normal) {
                            $Row.Cells[3].Value = $script:States.Normal
                         }
                    #endregion Step: Handle Asynchronous Check Initiation Errors
                }
            #endregion Step: Perform Asynchronous Window State Check
        }
    #endregion Function: Start-WindowStateCheck

    #region Function: Find-TargetRow
        function Find-TargetRow
        {
            [CmdletBinding()]
            [OutputType([System.Windows.Forms.DataGridViewRow])]
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridView]$Grid,

                [Parameter(Mandatory = $true)]
                [int]$ProcessId,

                [Parameter(Mandatory = $true)]
                [int]$RowIndex,

                [Parameter(Mandatory = $false)]
                [IntPtr]$WindowHandle = [IntPtr]::Zero
            )

            #region Step: Attempt Direct Index Access (if RowIndex is valid)
                $row = $null
                if ($RowIndex -ge 0 -and $RowIndex -lt $Grid.Rows.Count)
                {
                    $potentialRow = $Grid.Rows[$RowIndex]

                    # Verify the row at the index contains the correct process ID
                    if ($potentialRow.Tag -and $potentialRow.Tag.Id -eq $ProcessId)
                    {
                        # Optional: Verify window handle if provided
                        if ($WindowHandle -eq [IntPtr]::Zero -or ($potentialRow.Tag.MainWindowHandle -eq $WindowHandle)) {
                           $row = $potentialRow
                        }
                    }
                }
            #endregion Step: Attempt Direct Index Access (if RowIndex is valid)

            #region Step: Fallback to Linear Search if Direct Access Failed or Skipped
                if ($null -eq $row)
                {
                    foreach ($r in $Grid.Rows)
                    {
                        # Check if row has tag and matching process ID
                        if ($r.Tag -and $r.Tag.Id -eq $ProcessId)
                        {
                            # Optional: Verify window handle if provided
                            if ($WindowHandle -eq [IntPtr]::Zero -or ($r.Tag.MainWindowHandle -eq $WindowHandle))
                            {
                                $row = $r
                                break # Found the row, exit loop
                            }
                        }
                    }
                }
            #endregion Step: Fallback to Linear Search if Direct Access Failed or Skipped

            #region Step: Return Found Row (or null)
                return $row
            #endregion Step: Return Found Row (or null)
        }
    #endregion Function: Find-TargetRow

    #region Function: Clear-OldProcessCache
        function Clear-OldProcessCache
        {
            [CmdletBinding()]
            param()

            #region Step: Initialize Variables
                $now = Get-Date
                $keysToRemove = [System.Collections.Generic.List[object]]::new()
                $thresholdMinutes = 5 # Define cleanup threshold
            #endregion Step: Initialize Variables

            #region Step: Identify Cache Keys Older Than Threshold
                # Iterate safely over a copy of the keys in case the collection is modified
                $cacheKeys = @($script:ProcessCache.Keys)
                foreach ($key in $cacheKeys)
                {
                    # Check if key still exists before accessing
                    if ($script:ProcessCache.ContainsKey($key)) {
                        $entry = $script:ProcessCache[$key]
                        # Ensure LastSeen property exists and is a DateTime object
                        if ($entry -and $entry.ContainsKey('LastSeen') -and $entry.LastSeen -is [DateTime]) {
                            $age = ($now - $entry.LastSeen).TotalMinutes

                            # Mark entries older than the threshold for removal
                            if ($age -gt $thresholdMinutes)
                            {
                                $keysToRemove.Add($key)
                            }
                        } else {
                             # Mark invalid entries for removal
                             $keysToRemove.Add($key)
                        }
                    }
                }
            #endregion Step: Identify Cache Keys Older Than Threshold

            #region Step: Remove Identified Keys from Cache
                foreach ($key in $keysToRemove)
                {
                    # Check again if key exists before removing
                    if ($script:ProcessCache.ContainsKey($key)) {
                        $script:ProcessCache.Remove($key)
                    }
                }
            #endregion Step: Remove Identified Keys from Cache
        }
    #endregion Function: Clear-OldProcessCache

    #region Function: Set-WindowToolStyle
	function Set-WindowToolStyle
	{
		<#
		.SYNOPSIS
			Hides/Shows window from Alt-Tab (via Style) and Taskbar (via ITaskbarList) with Fallback.
		#>
		param(
			[Parameter(Mandatory = $true)]
			[IntPtr]$hWnd,
			[Parameter(Mandatory = $false)]
			[bool]$Hide = $true
		)

		#region Step: Define Constants
		$GWL_EXSTYLE      = -20
		$WS_EX_TOOLWINDOW = 0x00000080
		$WS_EX_APPWINDOW  = 0x00040000
		#endregion

		#region Step: Apply Logic
		try
		{
			# 1. Primary Method: ITaskbarList (COM)
			# We try this first. If it succeeds, $comSuccess is $true. 
			# If it fails (Class not registered/32-bit issue), $comSuccess is $false.
			$comSuccess = $false
			
            # Check for "Custom.TaskbarTool" (correct namespace)
			if ("Custom.TaskbarTool" -as [Type]) {
				# Inverted Logic: If $Hide is true, Visible is false (-not $Hide)
				$comSuccess = [Custom.TaskbarTool]::SetTaskbarState($hWnd, -not $Hide)
			}

			# 2. Logic for Window Styles (Alt-Tab & Fallback)
			[IntPtr]$currentExStylePtr = [Custom.Native]::GetWindowLongPtr($hWnd, $GWL_EXSTYLE)
			[long]$currentExStyle = $currentExStylePtr.ToInt64()
			$newExStyle = $currentExStyle

			$hideMinimized = $global:DashboardConfig.Config['Options']['HideMinimizedWindows'] -eq '1'

			# FALLBACK LOGIC:
			# If COM failed ($comSuccess -eq $false) AND we want to Hide ($Hide -eq $true),
			# we MUST force the TOOLWINDOW style, otherwise the window remains visible in the Taskbar.
			# We effectively override '$hideMinimized' only if the primary method broke.
			$shouldUseToolWindow = ($Hide -and $hideMinimized) -or ($Hide -and -not $comSuccess)

			if ($shouldUseToolWindow)
			{
				# Add TOOLWINDOW (Hides from Alt-Tab + Taskbar Fallback), Remove APPWINDOW
				$newExStyle = ($currentExStyle -bor $WS_EX_TOOLWINDOW) -band (-bnot $WS_EX_APPWINDOW)
			}
			else
			{
				# Remove TOOLWINDOW, Add APPWINDOW
				$newExStyle = ($currentExStyle -band (-bnot $WS_EX_TOOLWINDOW)) -bor $WS_EX_APPWINDOW
			}

			# 3. Apply Style Update if needed
			if ($newExStyle -ne $currentExStyle)
			{
                # Check state BEFORE hiding
                $isMinimized = [Custom.Native]::IsMinimized($hWnd)

                # 1. Hide the window to force Taskbar update
                [Custom.Native]::ShowWindow($hWnd, [Custom.Native]::SW_HIDE)
                
                # 2. Apply the new style
				[IntPtr]$newExStylePtr = [IntPtr]$newExStyle
				[Custom.Native]::SetWindowLongPtr($hWnd, $GWL_EXSTYLE, $newExStylePtr) | Out-Null
				
                # 3. Show Back (Conditional on State)
                if ($isMinimized) {
                    # If it was minimized, keep it minimized without activating or restoring.
                    # This prevents the "resize/down movement" glitch.
                    [Custom.Native]::ShowWindow($hWnd, [Custom.Native]::SW_SHOWMINNOACTIVE)
                } else {
                    # If it was normal, show it in current state without activating.
                    [Custom.Native]::ShowWindow($hWnd, [Custom.Native]::SW_SHOWNA)
                }
			}
		}
		catch
		{
			Write-Warning "DATAGRID: Failed to set window tool style for handle $hWnd. $($_.Exception.Message)"
		}
		#endregion
	}
    #endregion Function: Set-WindowToolStyle
#endregion Helper Functions

#region Core Functions
    #region Function: Update-DataGrid
        function Update-DataGrid
        {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.DataGridView]$Grid
            )

            #region Step: Validate Parameters and Configuration
                if (-not (Test-ValidParameters -Grid $Grid))
                {
                    return
                }
            #endregion Step: Validate Parameters and Configuration

            #region Step: Get Current List of Target Processes
                $currentProcesses = Get-ProcessList
            #endregion Step: Get Current List of Target Processes

            #region Step: Remove Rows for Terminated Processes
                Remove-TerminatedProcesses -Grid $Grid -CurrentProcesses $currentProcesses
            #endregion Step: Remove Rows for Terminated Processes

            #region Step: Handle Case of No Running Target Processes
                if ($currentProcesses.Count -eq 0)
                {
                    return
                }
            #endregion Step: Handle Case of No Running Target Processes

            #region Step: Create Efficient Row Lookup Dictionary
                # Build a hashtable for fast lookup of existing rows by Process ID
                $rowLookup = New-RowLookupDictionary -Grid $Grid
            #endregion Step: Create Efficient Row Lookup Dictionary

            #region Step: Process Each Running Target Process
                $processIndex = 0
                foreach ($process in $currentProcesses)
                {
                    #region Step: Increment Process Index (for display)
                        $processIndex++  # Start display index from 1
                    #endregion Step: Increment Process Index (for display)

                    #region Step: Determine Process Title (Window Title or Name)
                        # Use MainWindowTitle if available, otherwise fall back to ProcessName
                        $processTitle = if (-not [string]::IsNullOrEmpty($process.MainWindowTitle))
                        {
                            $process.MainWindowTitle
                        }
                        else
                        {
                            $process.ProcessName
                        }
                    #endregion Step: Determine Process Title (Window Title or Name)

                    #region Step: Check if Process Row Already Exists in Grid
                        $existingRow = $null
                        if ($rowLookup.ContainsKey($process.Id)) {
                            $existingRow = $rowLookup[$process.Id]
                        }
                    #endregion Step: Check if Process Row Already Exists in Grid

                    #region Step: Update Process Cache Entry (LastSeen, HasWindow)
                        # Update the cache entry for this process ID
                        if ($script:ProcessCache.ContainsKey($process.Id))
                        {
                            $script:ProcessCache[$process.Id].LastSeen = Get-Date
                            $script:ProcessCache[$process.Id].HasWindow = ($process.MainWindowHandle -ne 0 -and $process.MainWindowHandle -ne [IntPtr]::Zero)
                        }
                        # If not in cache, it will be added when the row is added (if applicable)
                    #endregion Step: Update Process Cache Entry (LastSeen, HasWindow)

                    #region Step: Update Existing Row or Add New Row
                        if ($existingRow)
                        {
                            #region Step: Update Existing Row Data
                                Update-ExistingRow -Row $existingRow -Index $processIndex -ProcessTitle $processTitle -ProcessId $process.Id -Process $process
                            #endregion Step: Update Existing Row Data
                        }
                        else
                        {
                            #region Step: Determine if New Row Should Be Added (Cache Logic)
                                # Add if not in cache, OR if it's in cache but previously had no window and now does.
                                $shouldAdd = $true
                                if ($script:ProcessCache.ContainsKey($process.Id))
                                {
                                    # If process is in cache but didn't have a window before, and still doesn't, don't add again.
                                    # This prevents adding processes during brief moments they might lose their window handle.
                                    $hasWindowNow = ($process.MainWindowHandle -ne 0 -and $process.MainWindowHandle -ne [IntPtr]::Zero)
                                    if (-not $script:ProcessCache[$process.Id].HasWindow -and -not $hasWindowNow)
                                    {
                                        $shouldAdd = $false
                                    }
                                }
                            #endregion Step: Determine if New Row Should Be Added (Cache Logic)

                            #region Step: Add New Row if Required
                                if ($shouldAdd)
                                {
                                    $newRow = Add-NewProcessRow -Grid $Grid -Process $process -Index $processIndex -ProcessTitle $processTitle
                                    # Use the newly added row for the state check below
                                    $existingRow = $newRow
                                    # Add the new row to the lookup if it was created successfully
                                    if ($newRow) {
                                        $rowLookup[$process.Id] = $newRow
                                    }
                                }
                            #endregion Step: Add New Row if Required
                        }
                    #endregion Step: Update Existing Row or Add New Row

                    #region Step: Start Asynchronous Window State Check for the Row
                        # If a row exists for this process (either updated or newly added), start the state check.
                        if ($existingRow)
                        {
                            Start-WindowStateCheck -Process $process -Row $existingRow -Grid $Grid
                        }
                    #endregion Step: Start Asynchronous Window State Check for the Row
                                    }
            #endregion Step: Process Each Running Target Process

            #region Step: Update All Row Indices for Sequential Numbering
                # Ensure the '#' column is correctly numbered after additions/removals
                UpdateRowIndices -Grid $Grid
            #endregion Step: Update All Row Indices for Sequential Numbering

            #region Step: Clean Up Old Process Cache Entries Periodically
                # Remove stale entries from the cache
                Clear-OldProcessCache
            #endregion Step: Clean Up Old Process Cache Entries Periodically
        }
    #endregion Function: Update-DataGrid

    #region Function: Start-DataGridUpdateTimer
        function Start-DataGridUpdateTimer
        {
            [CmdletBinding()]
            param()

            #region Step: Clean Up Existing Timer (if any)
                # Ensure the path to the timer exists before trying to access it
                if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and $global:DashboardConfig.Resources.Timers -and $global:DashboardConfig.Resources.Timers['dataGridUpdateTimer'])
                {
                    try {
                        $global:DashboardConfig.Resources.Timers['dataGridUpdateTimer'].Stop()
                        $global:DashboardConfig.Resources.Timers['dataGridUpdateTimer'].Dispose()
                        $global:DashboardConfig.Resources.Timers.Remove('dataGridUpdateTimer') # Remove reference
                    } catch {
                        Write-Verbose "DATAGRID: Error cleaning up existing timer: $_" -ForegroundColor Red
                    }
                }
            #endregion Step: Clean Up Existing Timer (if any)

            #region Step: Create and Configure New Timer
                $dataGridUpdateTimer = New-Object System.Windows.Forms.Timer
                $dataGridUpdateTimer.Interval = $script:UpdateInterval
            #endregion Step: Create and Configure New Timer

            #region Step: Define Timer Tick Event Handler (Calls Update-DataGrid)
                $tickAction = {
                    # Check if UI components are valid before attempting update
                    if ($global:DashboardConfig -and $global:DashboardConfig.UI -and $global:DashboardConfig.UI.DataGridFiller -and -not $global:DashboardConfig.UI.DataGridFiller.IsDisposed)
                    {
                        try
                        {
                            # Use BeginInvoke to run Update-DataGrid on the UI thread asynchronously
                            $global:DashboardConfig.UI.DataGridFiller.BeginInvoke([Action] {
                                    # Nested try/catch for the action itself
                                    try {
                                        Update-DataGrid -Grid $global:DashboardConfig.UI.DataGridFiller
                                    } catch {
                                        # Log error occurring within the invoked action
                                        Write-Verbose "DATAGRID: Error during invoked Update-DataGrid: $_" -ForegroundColor Red
                                    }
                                })
                        }
                        catch
                        {
                            # Log error occurring during BeginInvoke call
                            Write-Verbose "DATAGRID: Error invoking DataGrid update: $_" -ForegroundColor Red
                            # Consider stopping the timer if BeginInvoke fails repeatedly
                            # $this.Stop()
                        }
                    }
                    else
                    {
                        # UI no longer valid, stop the timer
                        Write-Verbose "DATAGRID: UI or DataGrid disposed, stopping update timer." -ForegroundColor Yellow
                        $this.Stop() # $this refers to the timer object inside the event handler
                    }
                }
                $dataGridUpdateTimer.Add_Tick($tickAction)
            #endregion Step: Define Timer Tick Event Handler (Calls Update-DataGrid)

            #region Step: Start the Timer
                $dataGridUpdateTimer.Start()
            #endregion Step: Start the Timer

            #region Step: Store Timer Reference Globally
                # Ensure the Timers hashtable exists
                if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and -not $global:DashboardConfig.Resources.ContainsKey('Timers')) {
                    $global:DashboardConfig.Resources['Timers'] = @{}
                }
                # Store the timer reference
                if ($global:DashboardConfig -and $global:DashboardConfig.Resources) {
                   $global:DashboardConfig.Resources.Timers['dataGridUpdateTimer'] = $dataGridUpdateTimer
                }
            #endregion Step: Store Timer Reference Globally
        }
    #endregion Function: Start-DataGridUpdateTimer
#endregion Core Functions

#region Function: Restore-WindowStyles
        function Restore-WindowStyles
        {
            [CmdletBinding()]
            param()

            Write-Verbose "DATAGRID: Restoring window styles for managed clients..." -ForegroundColor Yellow

            # Iterate through all process IDs currently in the cache
            foreach ($processId in $script:ProcessCache.Keys)
            {
                try
                {
                    # Attempt to get the process by ID
                    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue

                    # If process exists and has a main window handle, restore its style
                    if ($process -and $process.MainWindowHandle -ne 0 -and $process.MainWindowHandle -ne [IntPtr]::Zero)
                    {
                        Set-WindowToolStyle -hWnd $process.MainWindowHandle -Hide $false
                        Write-Verbose "DATAGRID: Restored style for PID $($process.Id)."
                    }
                }
                catch
                {
                    Write-Verbose "DATAGRID: Error restoring window style for PID $($processId). $_" -ForegroundColor Red
                }
            }
            Write-Verbose "DATAGRID: Finished restoring window styles." -ForegroundColor Yellow
        }
#endregion Function: Restore-WindowStyles

#region Module Exports
    #region Step: Export Public Functions
        # Export the functions intended for public use by other modules or scripts.
        Export-ModuleMember -Function Start-DataGridUpdateTimer, Update-DataGrid, Restore-WindowStyles
    #endregion Step: Export Public Functions
#endregion Module Exports