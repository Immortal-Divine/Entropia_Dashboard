<# login.psm1 #>

#region Configuration

$global:LoginCancellation = [hashtable]::Synchronized(@{
		IsCancelled         = $false
		ScriptInitiatedMove = $false
	})

if (-not $global:LoginInputGate) { $global:LoginInputGate = [System.Threading.SemaphoreSlim]::new(1, 1) }
if (-not $global:LoginLauncherGate) { $global:LoginLauncherGate = [System.Threading.SemaphoreSlim]::new(1, 1) }
if (-not $global:LoginJobs) { $global:LoginJobs = [hashtable]::Synchronized(@{}) }
$TOTAL_STEPS_PER_CLIENT = 15
$script:LastLoginToastTick = 0

if (-not $global:LoginResources)
{
	$global:LoginResources = @{
		PowerShellInstance  = $null
		Runspace            = $null
		EventSubscriptionId = $null
		EventSubscriber     = $null
		InfoSubscription    = $null
		AsyncResult         = $null
		IsStopping          = $false
		IsMouseHookActive   = $false
	}
}

if (-not $global:DashboardConfig) { $global:DashboardConfig = @{} }
if (-not $global:DashboardConfig.State) { $global:DashboardConfig.State = @{} }
if ($global:DashboardConfig.State -is [System.Collections.IDictionary])
{
	if (-not $global:DashboardConfig.State.Contains('LoginNotificationMap')) { $global:DashboardConfig.State['LoginNotificationMap'] = @{} }
}
else
{
	if (-not $global:DashboardConfig.State.LoginNotificationMap) { $global:DashboardConfig.State.LoginNotificationMap = @{} }
}
if (-not $global:LoginNotificationStack) { $global:LoginNotificationStack = [System.Collections.ArrayList]::new() }


#endregion

#region Helper Functions

function ValidateAndRefreshProcessReference
{
	param(
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.DataGridViewRow]$Row,
		[Parameter(Mandatory = $false)]
		[int]$MaxRetries = 3,
		[Parameter(Mandatory = $false)]
		[int]$RetryDelayMs = 100
	)
	
	$process = $Row.Tag
	
	if ($null -ne $process -and $process.Id -gt 0)
	{
		try
		{
			$testProc = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
			if ($null -ne $testProc -and -not $testProc.HasExited)
			{
				return $process
			}
		}
		catch {}
	}
	
	for ($attempt = 0; $attempt -lt $MaxRetries; $attempt++)
	{
		try
		{
			if ($Row.Cells -and $Row.Cells.Count -gt 1)
			{
				$cellTitle = $Row.Cells[1].Value
				if ($null -ne $cellTitle)
				{
					$titleStr = [string]$cellTitle
					
					$allProcesses = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -and $_.MainWindowHandle -ne 0 }
					foreach ($proc in $allProcesses)
					{
						if ($proc.MainWindowTitle -like "*$titleStr*")
						{
							$Row.Tag = $proc
							return $proc
						}
					}
				}
			}
		}
		catch {}
		
		if ($attempt -lt ($MaxRetries - 1))
		{
			Start-Sleep -Milliseconds $RetryDelayMs
		}
	}
	
	return $null
}

function ResolveLoginProfileName
{
	[CmdletBinding()]
	[OutputType([string])]
	param(
		[Parameter()]
		[string]$Title = '',
		[Parameter()]
		$Process = $null
	)

	$profileName = ''
	if (-not [string]::IsNullOrWhiteSpace($Title) -and $Title -match '^\[(?<pName>[^\]]+)\]') { $profileName = $Matches['pName'] }

	if ([string]::IsNullOrWhiteSpace($profileName) -and $Process -and $Process.Id)
	{
		try { $profileName = GetProcessProfile -Process $Process } catch { $profileName = '' }
	}

	if ([string]::IsNullOrWhiteSpace($profileName)) { return 'Default' }
	return $profileName
}

function GetLoginProfileConfig
{
	[CmdletBinding()]
	[OutputType([System.Collections.IDictionary])]
	param(
		[Parameter()]
		[string]$ProfileName = 'Default'
	)

	$loginConfig = $null
	if ($global:DashboardConfig -and $global:DashboardConfig.Config) { $loginConfig = $global:DashboardConfig.Config['LoginConfig'] }

	if ($loginConfig -is [System.Collections.IDictionary])
	{
		if (-not [string]::IsNullOrWhiteSpace($ProfileName) -and $loginConfig.Contains($ProfileName) -and $loginConfig[$ProfileName] -is [System.Collections.IDictionary])
		{
			return $loginConfig[$ProfileName]
		}
		if ($loginConfig.Contains('Default') -and $loginConfig['Default'] -is [System.Collections.IDictionary])
		{
			return $loginConfig['Default']
		}
	}

	return @{}
}

function GetClientRootFolder
{
	[CmdletBinding()]
	[OutputType([string])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$ExeFolder
	)

	$candidates = @($ExeFolder)
	$folderName = Split-Path -Leaf -Path $ExeFolder
	if ($folderName -match '^(?i)(bin\d*|win\d*|x86|x64)$')
	{
		$parentFolder = Split-Path -Parent -Path $ExeFolder
		if (-not [string]::IsNullOrEmpty($parentFolder)) { $candidates = @($parentFolder) + $candidates }
	}

	foreach ($candidate in $candidates)
	{
		if (Test-Path -LiteralPath (Join-Path -Path $candidate -ChildPath 'Log')) { return $candidate }
	}

	return $candidates[0]
}

function GetClientLogPath
{
	param(
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.DataGridViewRow]$Row
	)
	$determinedLogBaseFolder = ''
	$process = ValidateAndRefreshProcessReference -Row $Row -MaxRetries 2 -RetryDelayMs 50
	$profileName = ''
	$processExeBaseFolder = ''
	if ($null -eq $Row.Cells) { return '' }
	$processTitle = $Row.Cells[1].Value
	if ($null -eq $processTitle) { $processTitle = '' }
	$null = $entryNum; $entryNum = $Row.Cells[0].Value

	$profileName = ResolveLoginProfileName -Title ([string]$processTitle) -Process $process

	if ($process -and $process.Id)
	{
		$processExePath = [Custom.Native]::GetProcessPathById($process.Id)
		if (-not [string]::IsNullOrEmpty($processExePath))
		{
			$processExeBaseFolder = Split-Path -Parent -Path $processExePath
		}
	}

	if ([string]::IsNullOrEmpty($determinedLogBaseFolder) -and
		$global:DashboardConfig -and $global:DashboardConfig.Config -and
		$global:DashboardConfig.Config.Contains('Profiles'))
	{
		$foundProfileKey = $null
		foreach ($key in $global:DashboardConfig.Config.Profiles.Keys)
		{
			if ($key.ToLowerInvariant() -eq $profileName.ToLowerInvariant())
			{
				$foundProfileKey = $key
				break
			}
		}
		if ($foundProfileKey)
		{
			$profilePath = $global:DashboardConfig.Config.Profiles[$foundProfileKey]
			if (-not [string]::IsNullOrEmpty($profilePath))
			{
				$determinedLogBaseFolder = $profilePath
			}
		}
	}

	if ([string]::IsNullOrEmpty($determinedLogBaseFolder) -and -not [string]::IsNullOrEmpty($processExeBaseFolder))
	{
		$determinedLogBaseFolder = GetClientRootFolder -ExeFolder $processExeBaseFolder
	}

	if ([string]::IsNullOrEmpty($determinedLogBaseFolder))
	{
		if ($global:DashboardConfig.Config)
		{
			$launcherPathConfig = $global:DashboardConfig.Config['LauncherPath']
			if ($launcherPathConfig -and $launcherPathConfig.Contains('LauncherPath'))
			{
				$launcherPath = $launcherPathConfig['LauncherPath']
				if (-not [string]::IsNullOrEmpty($launcherPath))
				{
					$determinedLogBaseFolder = Split-Path -Path $launcherPath -Parent
				}
			}
		}
	}

	if ([string]::IsNullOrEmpty($determinedLogBaseFolder))
	{
		$determinedLogBaseFolder = Join-Path -Path $env:APPDATA -ChildPath 'Entropia_Dashboard'
	}

	$actualLogPath = Join-Path -Path $determinedLogBaseFolder -ChildPath "Log\network_$([DateTime]::Now.ToString('yyyyMMdd')).log"
	return $actualLogPath
}

#endregion

#region Core Function

function LoginSelectedRow
{
	param(
		[Parameter(Mandatory = $false)]
		$RowInput,
		[Parameter(Mandatory = $false)]
		[Object]$WindowHandle = [IntPtr]::Zero,
		[string]$LogFilePath
	)
    
	
	if ($WindowHandle -isnot [IntPtr])
	{
		if ($WindowHandle)
		{
			try { $WindowHandle = [IntPtr]$WindowHandle } catch { $WindowHandle = [IntPtr]::Zero }
		}
		else
		{
			$WindowHandle = [IntPtr]::Zero
		}
	}

	$global:DashboardConfig.State['LoginActive'] = $true

	try 
	{
		$null = $rowsToProcess; $rowsToProcess = @()
		$rawSelection = @()

		if ($RowInput)
		{
			if ($RowInput -is [System.Collections.IEnumerable] -and $RowInput -isnot [string] -and $RowInput -isnot [System.Windows.Forms.DataGridViewRow])
			{
				foreach ($item in $RowInput)
				{
					$rawSelection += $item
				}
			}
			else
			{
				$rawSelection += $RowInput
			}
		}
		elseif ($global:DashboardConfig.UI.DataGridMain.SelectedRows.Count -gt 0)
		{
			$rawSelection = $global:DashboardConfig.UI.DataGridMain.SelectedRows | Sort-Object Index
		}
		else
		{
			if ($global:LoginJobs.Count -eq 0) { $global:DashboardConfig.State['LoginActive'] = $false }
			return
		}

		$enrichedData = $rawSelection | ForEach-Object {
			$item = $_
			$actualRow = $null
			$entryNum = 0

			if ($null -eq $item)
			{
				return
			}
			
			if ($item.PSTypeNames -contains 'LoginOverrideWrapper')
			{
				$actualRow = $item.Row
				$entryNum = $item.OverrideAccountID -as [int]
			} 
			elseif ($item.Cells -and $item.Cells.Count -gt 0)
			{
				$actualRow = $item
				$entryNum = $actualRow.Cells[0].Value -as [int]
			}

			if ($null -ne $actualRow -and $null -ne $actualRow.Cells -and $actualRow.Cells.Count -gt 1)
			{
				$title = [string]$actualRow.Cells[1].Value
				$profileName = ResolveLoginProfileName -Title $title
                
				[PSCustomObject]@{ 
					OriginalRow = $actualRow
					EntryNum    = $entryNum 
					Profile     = $profileName 
				}
			}
		}
		
		$profilePriorityMap = [ordered]@{}
		$priorityCounter = 0
		foreach ($item in $enrichedData)
		{
			$pKey = $item.Profile.ToString()
			if (-not $profilePriorityMap.Contains($pKey)) { $profilePriorityMap[$pKey] = $priorityCounter++ }
		}
		
		$sortedData = $enrichedData | Sort-Object @{Expression = {$profilePriorityMap[$_.Profile.ToString()]}; Ascending = $true}, @{Expression = {$_.EntryNum}; Ascending = $true}
		
		$loginTasks = New-Object System.Collections.Generic.LinkedList[PSObject]
		foreach ($dataItem in $sortedData)
		{
			$row = $dataItem.OriginalRow
			$entryNum = [int]$dataItem.EntryNum 
            
			if ($null -eq $row.Cells -or $null -eq $row.Cells[1].Value) { continue }
			$processTitle = $row.Cells[1].Value.ToString()
			$profileName = ResolveLoginProfileName -Title $processTitle

			$process = ValidateAndRefreshProcessReference -Row $row -MaxRetries 5 -RetryDelayMs 200
			if ($null -eq $process)
			{
				Write-Verbose "LOGIN: Skipping entry $entryNum ($profileName) - process reference could not be established"
				continue
			}

			$profileName = ResolveLoginProfileName -Title $processTitle -Process $process

			$actualLogPath = $LogFilePath
			if ([string]::IsNullOrEmpty($actualLogPath))
			{
				$actualLogPath = GetClientLogPath -Row $row
			}
			if ($null -eq $actualLogPath) { $actualLogPath = '' }

			$jobData = [PSCustomObject]@{
				EntryNumber    = $entryNum
				ProcessId      = $process.Id
				ExplicitHandle = if (($WindowHandle -ne [IntPtr]::Zero) -and ($RowInput -eq $row -or ($RowInput.Row -eq $row))) { $WindowHandle } else { [IntPtr]::Zero }
				LogPath        = $actualLogPath
				Config         = (GetLoginProfileConfig -ProfileName $profileName)
				ProfileName    = $profileName
				ClientTitle    = $processTitle
			}
			
			$loginTasks.AddLast([PSCustomObject]@{ JobData = $jobData; RetryCount = 0 })
		}
		
		ShowToast -Title 'Login Process' -Message 'Starting Login Process...' -Type 'Info' -Key 9999 -TimeoutSeconds 0

		$global:DashboardConfig.UI.LoginButton.BackColor = Get-Theme 'ButtonBg'
		$global:DashboardConfig.UI.LoginButton.Text = 'Running...'

		if ($global:LoginJobs.Count -eq 0)
		{
			$global:LoginCancellation.IsCancelled = $false
			$global:LoginCancellation.ScriptInitiatedMove = $false
		}
		$global:LoginResources['IsStopping'] = $false

		$hookCallback = [Custom.MouseHookManager+HookProc] {
			param($nCode, $wParam, $lParam)
			if ($nCode -ge 0 -and $wParam -eq 0x0200)
			{
				if (-not $global:LoginCancellation.ScriptInitiatedMove)
				{
					
					if ($global:LoginJobs -and -not $global:LoginResources['IsStopping'])
					{
						$runningPipelines = $false
						$hookJobs = @()
						try
						{
							[System.Threading.Monitor]::Enter($global:LoginJobs.SyncRoot)
							try { $hookJobs = @($global:LoginJobs.Values) } finally { [System.Threading.Monitor]::Exit($global:LoginJobs.SyncRoot) }
						}
						catch {}
						foreach ($loginJob in $hookJobs)
						{
							$psInstance = $loginJob.PowerShell
							if ($psInstance -and $psInstance.InvocationStateInfo.State -eq 'Running')
							{
								$runningPipelines = $true
								try { $psInstance.Stop() } catch {}
							}
						}
						if ($runningPipelines)
						{
							Write-Verbose 'LOGIN: Mouse movement detected, stopping login pipelines.'
							$global:LoginCancellation.IsCancelled = $true
						}
					}
				}
			}
			return [Custom.MouseHookManager]::CallNextHookEx([Custom.MouseHookManager]::HookId, $nCode, $wParam, $lParam)
		}
		if (-not $global:LoginResources.IsMouseHookActive)
		{
			try
			{
				[Custom.MouseHookManager]::Start($hookCallback)
				$global:LoginResources.IsMouseHookActive = $true
				Write-Verbose 'LOGIN: Mouse hook started.'
			}
			catch
			{
				Write-Warning "LOGIN: Failed to start mouse hook: $_"
			}
		}

		
		$varsToInject = @{
			'DashboardConfig'        = $global:DashboardConfig
			'TOTAL_STEPS_PER_CLIENT' = $TOTAL_STEPS_PER_CLIENT
			'CancellationContext'    = $global:LoginCancellation
			'LoginResources'         = $global:LoginResources
			'LoginInputGate'         = $global:LoginInputGate
			'LoginLauncherGate'      = $global:LoginLauncherGate
		}

		$assembliesToAdd = [System.Collections.Generic.HashSet[string]]::new()
		$customTypesToLoad = @('Custom.Native', 'Custom.Ftool', 'Custom.MouseHookManager', 'Custom.ColorWriter')
        
		foreach ($typeName in $customTypesToLoad)
		{
			$type = $typeName -as [Type]
			if ($type -and -not [string]::IsNullOrEmpty($type.Assembly.Location))
			{
				$assembliesToAdd.Add($type.Assembly.Location) | Out-Null
			}
		}
        
		
		if ($assembliesToAdd.Count -eq 0)
		{
			$loadedAssemblies = [AppDomain]::CurrentDomain.GetAssemblies()
			foreach ($asm in $loadedAssemblies)
			{
				try
				{
					if ($asm.IsDynamic) { continue }
					if ([string]::IsNullOrWhiteSpace($asm.Location)) { continue }

					foreach ($t in $asm.GetExportedTypes())
					{
						if ($customTypesToLoad -contains $t.FullName)
						{
							$assembliesToAdd.Add($asm.Location) | Out-Null
							break
						}
					}
				}
				catch {}
			}
		}

		
		$localRunspace = $global:LoginResources.Runspace
		if ($null -eq $localRunspace -or $localRunspace.IsClosed)
		{
			$localRunspace = NewManagedRunspace `
				-Name 'LoginRunspacePool' `
				-MinRunspaces 1 `
				-MaxRunspaces 10 `
				-SessionVariables $varsToInject `
				-Assemblies ($assembliesToAdd | Select-Object -Unique)

			$global:LoginResources = @{
				PowerShellInstance   = $null
				Runspace             = $localRunspace
				EventSubscriptionId  = $null
				EventSubscriber      = $null
				ProgressSubscription = $null
				AsyncResult          = $null
				IsStopping           = $false
				IsMouseHookActive    = $global:LoginResources.IsMouseHookActive
			}
		}
		else
		{
			$global:LoginResources['Runspace'] = $localRunspace
		}

		
		$loginScript = {
			param($Jobs, $TotalStepsPerClient, $GlobalOptions, $LoginConfig, $CancellationContext, $State, $JobId, $JobPidSet, $JobEntry)
			$LoginTasks = $Jobs

			if (-not $TotalStepsPerClient -or $TotalStepsPerClient -eq 0) { $TotalStepsPerClient = 15 }

			if (-not ('Custom.Native' -as [Type]))
			{
				throw 'CRITICAL ERROR: [Custom.Native] type is missing in Background Runspace. Automation cannot proceed.'
			}
            
			$Global:CurrentActiveProcessId = 0
			$Global:CurrentActiveWindowHandle = [IntPtr]::Zero
			$InformationPreference = 'SilentlyContinue'

			function CheckCancel
			{
				if ($CancellationContext.IsCancelled) { throw 'LoginCancelled' }
			}

			function Wait-Gate
			{
				param([System.Threading.SemaphoreSlim]$Gate)
				while (-not $Gate.Wait(100))
				{
					CheckCancel
				}
			}

			function Enter-InputGate
			{
				if ($global:LoginInputGateHeld) { return $false }
				Wait-Gate -Gate $global:LoginInputGate
				$global:LoginInputGateHeld = $true
				return $true
			}

			function Exit-InputGate
			{
				param([bool]$Acquired)
				if (-not $Acquired) { return }
				$global:LoginInputGateHeld = $false
				try { [void]$global:LoginInputGate.Release() } catch {}
			}

			function Register-LogClaim
			{
				param([string]$LogPath, [string]$Marker, [int]$Required = 1)
				if ([string]::IsNullOrWhiteSpace($LogPath) -or [string]::IsNullOrWhiteSpace($Marker)) { return $null }
				try { return [Custom.LogClaimLedger]::Register($LogPath, $Marker, $Required) } catch { return $null }
			}

			function Wait-LogClaim
			{
				param($Claim, [string]$LogPath, [int]$TimeoutSeconds)
				if ($null -eq $Claim) { return $false }
				if ($TimeoutSeconds -le 0) { $TimeoutSeconds = 60 }
				$waitSw = [System.Diagnostics.Stopwatch]::StartNew()
				try
				{
					while ($waitSw.Elapsed.TotalSeconds -lt $TimeoutSeconds)
					{
						CheckCancel
						try { [Custom.LogClaimLedger]::Poll($LogPath) } catch {}
						if ([Custom.LogClaimLedger]::IsDone($Claim)) { return $true }
						SleepWithCancel -Milliseconds 100
					}
					return $false
				}
				finally
				{
					if (-not [Custom.LogClaimLedger]::IsDone($Claim)) { try { [Custom.LogClaimLedger]::Cancel($Claim) } catch {} }
				}
			}

			function Invoke-LoggedKeyPress
			{
				param([int]$VirtualKeyCode, [string]$LogPath, [string]$Marker, [int]$Required = 1)
				$gateAcquired = Enter-InputGate
				$claim = $null
				try
				{
					$claim = Register-LogClaim -LogPath $LogPath -Marker $Marker -Required $Required
					$null = Invoke-KeyPress -VirtualKeyCode $VirtualKeyCode
				}
				catch
				{
					if ($null -ne $claim) { try { [Custom.LogClaimLedger]::Cancel($claim) } catch {} }
					throw
				}
				finally { Exit-InputGate $gateAcquired }
				return $claim
			}

			function SleepWithCancel
			{
				param([int]$Milliseconds)
				$sw = [System.Diagnostics.Stopwatch]::StartNew()
				while ($sw.Elapsed.TotalMilliseconds -lt $Milliseconds)
				{
					if ($CancellationContext.IsCancelled) { throw 'LoginCancelled' }
					Start-Sleep -Milliseconds 50
				}
			}

			function EnsureWindowResponsive
			{
				if ($Global:CurrentActiveProcessId -eq 0) { return }
				$proc = Get-Process -Id $Global:CurrentActiveProcessId -ErrorAction SilentlyContinue
				if (-not $proc) { throw "Process with ID $Global:CurrentActiveProcessId has terminated." }

				$sw = [System.Diagnostics.Stopwatch]::StartNew()
				while (-not $proc.Responding)
				{
					if ($sw.Elapsed.TotalSeconds -gt 60) { throw 'Timeout: Window is Not Responding (Hung).' }
					CheckCancel
					$proc.Refresh()
					Start-Sleep -Milliseconds 10
				}
				try { $proc.WaitForInputIdle(20) | Out-Null } catch {}
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

			function EnsureWindowStateCore
			{
				$hWnd = $Global:CurrentActiveWindowHandle
				if ($hWnd -eq [IntPtr]::Zero) { return }
                
				SetWindowToolStyle -hWnd $hWnd -Hide $false
				$changed = $false
				$CancellationContext.ScriptInitiatedMove = $true
				try
				{
					if ([Custom.Native]::IsWindowMinimized($hWnd))
					{
						[Custom.Native]::ShowWindow($hWnd, 6)
						[Custom.Native]::ShowWindow($hWnd, 9)
						$changed = $true
					}
                    
					$fg = [Custom.Native]::GetForegroundWindow()
					if ($fg -ne $hWnd)
					{
						if (-not [Custom.Native]::SetForegroundWindow($hWnd))
						{
							[Custom.Native]::keybd_event(0x12, 0, 0, [IntPtr]::Zero)
							Start-Sleep -Milliseconds 10
							[Custom.Native]::SetForegroundWindow($hWnd)
							[Custom.Native]::keybd_event(0x12, 0, 2, [IntPtr]::Zero)
						}
						$changed = $true
					}
				}
				finally
				{
					$CancellationContext.ScriptInitiatedMove = $false
				}
                
				if ($changed)
				{
					SleepWithCancel -Milliseconds 500
				}

				$fg = [Custom.Native]::GetForegroundWindow()
				if ($fg -ne $hWnd)
				{
					[Custom.Native]::ShowWindow($hWnd, 9)
					[Custom.Native]::SetForegroundWindow($hWnd)
				}
			}

			function RestoreWindowForInput
			{
				$hWnd = $Global:CurrentActiveWindowHandle
				if ($hWnd -eq [IntPtr]::Zero) { return }

				SetWindowToolStyle -hWnd $hWnd -Hide $false
				try
				{
					if ([Custom.Native]::IsWindowMinimized($hWnd))
					{
						[void][Custom.Native]::ShowWindow($hWnd, 4)
						SleepWithCancel -Milliseconds 60
					}
				}
				catch {}
			}

			function Invoke-MouseClickCore
			{
				param([int]$X, [int]$Y, [switch]$DoubleClick)
				CheckCancel; EnsureWindowResponsive; CheckCancel
                
				$hWnd = $Global:CurrentActiveWindowHandle
				if ($hWnd -eq [IntPtr]::Zero) { $hWnd = [Custom.Native]::GetForegroundWindow() }

				RestoreWindowForInput

				$pt = New-Object Custom.Win32Point
				$pt.X = $X; $pt.Y = $Y
				[void][Custom.Win32MouseUtils]::ScreenToClient($hWnd, [ref]$pt)
				$lParam = [Custom.Win32MouseUtils]::MakeLParam($pt.X, $pt.Y).ToInt32()

				$CancellationContext.ScriptInitiatedMove = $true
				try
				{
					[void][Custom.Native]::SetCursorPos(($X - 3), ($Y - 3)); Start-Sleep -Milliseconds 20
					[void][Custom.Native]::SetCursorPos($X, $Y); Start-Sleep -Milliseconds 30

					$WM_MOUSEMOVE = 0x0200; $WM_LBUTTONDOWN = 0x0201; $WM_LBUTTONUP = 0x0202; $WM_LBUTTONDBLCLK = 0x0203
					[Custom.Ftool]::fnPostMessage($hWnd, $WM_MOUSEMOVE, 0, $lParam); Start-Sleep -Milliseconds 15
					[Custom.Ftool]::fnPostMessage($hWnd, $WM_LBUTTONDOWN, 1, $lParam); Start-Sleep -Milliseconds 25
					[Custom.Ftool]::fnPostMessage($hWnd, $WM_LBUTTONUP, 0, $lParam)
					if ($DoubleClick)
					{
						Start-Sleep -Milliseconds 25
						[Custom.Ftool]::fnPostMessage($hWnd, $WM_LBUTTONDBLCLK, 1, $lParam); Start-Sleep -Milliseconds 25
						[Custom.Ftool]::fnPostMessage($hWnd, $WM_LBUTTONUP, 0, $lParam)
					}

					Start-Sleep -Milliseconds 100
				}
				catch { Write-Verbose "Mouse click failed: $_" } finally
				{ 
					$CancellationContext.ScriptInitiatedMove = $false; CheckCancel
				}
			}

			function Invoke-KeyPressCore
			{
				param([int]$VirtualKeyCode)
				CheckCancel; EnsureWindowResponsive; CheckCancel
                
				RestoreWindowForInput

				try
				{
					$hWnd = $Global:CurrentActiveWindowHandle
					if ($hWnd -eq [IntPtr]::Zero) { $hWnd = [Custom.Native]::GetForegroundWindow() }
					[Custom.Ftool]::fnPostMessage($hWnd, 0x0100, $VirtualKeyCode, 0); SleepWithCancel -Milliseconds 25
					[Custom.Ftool]::fnPostMessage($hWnd, 0x0101, $VirtualKeyCode, 0); SleepWithCancel -Milliseconds 25
				}
				catch { if ($_.Exception.Message -eq 'LoginCancelled') { throw }; throw "KeyPress Failed: $($_.Exception.Message)" }
			}

			function EnsureWindowState
			{
				$gateAcquired = Enter-InputGate
				try { EnsureWindowStateCore }
				finally { Exit-InputGate $gateAcquired }
			}

			function Invoke-MouseClick
			{
				param([int]$X, [int]$Y)
				CheckCancel; EnsureWindowResponsive; CheckCancel
				$gateAcquired = Enter-InputGate
				try { Invoke-MouseClickCore -X $X -Y $Y }
				finally { Exit-InputGate $gateAcquired }
			}

			function Invoke-MouseDoubleClick
			{
				param([int]$X, [int]$Y)
				CheckCancel; EnsureWindowResponsive; CheckCancel
				$gateAcquired = Enter-InputGate
				try { Invoke-MouseClickCore -X $X -Y $Y -DoubleClick }
				finally { Exit-InputGate $gateAcquired }
			}

			function Invoke-KeyPress
			{
				param([int]$VirtualKeyCode)
				CheckCancel; EnsureWindowResponsive; CheckCancel
				$gateAcquired = Enter-InputGate
				try { Invoke-KeyPressCore -VirtualKeyCode $VirtualKeyCode }
				finally { Exit-InputGate $gateAcquired }
			}

			function ParseCoordinates
			{
				param([string]$ConfigString)
				if ([string]::IsNullOrWhiteSpace($ConfigString) -or $ConfigString -notmatch ',') { return $null }
				$parts = $ConfigString.Split(',')
				if ($parts.Count -eq 2) { return @{ X = [int]$parts[0].Trim(); Y = [int]$parts[1].Trim() } }
				return $null
			}

			function Wait-ForLogEntry
			{
				param($LogPath, $SearchStrings, $TimeoutSeconds)
				if ($TimeoutSeconds -le 0) { $TimeoutSeconds = 60 }
				$waitSw = [System.Diagnostics.Stopwatch]::StartNew()
				while ($waitSw.Elapsed.TotalSeconds -lt $TimeoutSeconds)
				{
					CheckCancel
					if (-not [string]::IsNullOrWhiteSpace($LogPath) -and (Test-Path $LogPath))
					{
						try
						{
							$lastLines = Get-Content $LogPath -Tail 20 -ErrorAction SilentlyContinue
							if ($lastLines) { foreach ($str in $SearchStrings) { if ($lastLines -match [regex]::Escape($str)) { return $true } } }
						}
						catch {}
					}
					SleepWithCancel -Milliseconds 250
				}
				return $false
			}

			function Wait-UntilWorldLoaded
			{
				param($LogPath, $Config, $TotalSteps, $CurrentStep, $ClientIdx, $ClientCount, $EntryNum, $ProfileName, $ServerID, $ChannelID, $ClientTitle)
				$threshold = 3; $searchStr = '13 - CACHE_ACK_JOIN'
				if ($Config['WorldLoadLogThreshold']) { $threshold = [int]$Config['WorldLoadLogThreshold'] }
				if ($Config['WorldLoadLogEntry']) { $searchStr = $Config['WorldLoadLogEntry'] }
				$foundCount = 0; $timeout = New-TimeSpan -Minutes 2; $sw = [System.Diagnostics.Stopwatch]::StartNew()
				$CurrentStep++; ReportLoginProgress -Action 'Waiting for World Load...' -Step $CurrentStep -TotalSteps $TotalSteps -ClientIdx $ClientIdx -ClientCount $ClientCount -EntryNum $EntryNum -ProfileName $ProfileName -ServerID $ServerID -ChannelID $ChannelID -ClientTitle $ClientTitle
				while ($foundCount -lt $threshold)
				{
					if ($sw.Elapsed -gt $timeout) { throw 'World load timeout' }
					CheckCancel
					if (-not [string]::IsNullOrWhiteSpace($LogPath) -and (Test-Path $LogPath))
					{
						try
						{
							$lines = Get-Content $LogPath -Tail 50 -ErrorAction SilentlyContinue
							if ($lines) { $foundCount = ($lines | Select-String -SimpleMatch $searchStr).Count }
						}
						catch {}
					}
					SleepWithCancel -Milliseconds 250
				}
			}

			function Write-LogWithRetry
			{
				param([string]$FilePath, [string]$Value)
				if ([string]::IsNullOrWhiteSpace($FilePath)) { return }
				for ($i = 0; $i -lt 5; $i++) { try { Set-Content -Path $FilePath -Value $Value -Force -ErrorAction Stop; return } catch { Start-Sleep -Milliseconds 50 } }
			}

			function Write-Verbose
			{
				[CmdletBinding()]
				param([string]$Message, [string]$ForegroundColor = 'DarkGray')
                
				$dateStr = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
				$callerName = 'LoginSelectedRow'
				$bracketedCaller = "[$callerName]"
				$paddedCaller = $bracketedCaller.PadRight(35)
				$prefix = " | $dateStr - $paddedCaller - "
				$finalMsg = "$prefix$Message"

				if ('Custom.ColorWriter' -as [Type])
				{
					[Custom.ColorWriter]::WriteColored($finalMsg, $ForegroundColor)
				}
				else
				{
					Write-Host $finalMsg -ForegroundColor $ForegroundColor
				}
			}

			function ReportLoginProgress
			{
				param($Action, $Step, $TotalSteps, $ClientIdx, $ClientCount, $EntryNum, $ProfileName, $ServerID, $ChannelID, $ClientTitle)
				
				$finalClientTitle = $ClientTitle
				$setupName = $DashboardConfig.Resources['CurrentLaunchSetupName']
				if ((-not [string]::IsNullOrEmpty($setupName)) -and $DashboardConfig.Config)
				{
					$setupKey = "Setup_$setupName"
					if ($DashboardConfig.Config.Contains($setupKey))
					{
						$setupData = $DashboardConfig.Config[$setupKey]
						foreach ($clientSetupKey in $setupData.Keys)
						{
							$parts = $setupData[$clientSetupKey] -split ',', 5
							if ($parts.Length -ge 4)
							{
								if (($parts[1].Trim() -eq $ProfileName) -and (($parts[2].Trim() -as [int]) -eq $EntryNum))
								{
									$finalClientTitle = $parts[3].Trim()
									break
								}
							}
						}
					}
				}

				$pct = 0; if ($TotalSteps -gt 0) { $pct = [int](($Step / $TotalSteps) * 100) }
				
				$accDisplay = "($EntryNum)"
				if (-not [string]::IsNullOrEmpty($finalClientTitle))
				{
					$cleanTitle = $finalClientTitle
					if ($cleanTitle -match '^\[.*?\]\s*(.*)') { $cleanTitle = $Matches[1] }
					$accDisplay = "$cleanTitle ($EntryNum)"
				}
				
				$msg = "Profile: $ProfileName | Server ($ServerID) | Channel ($ChannelID)`nAccount: $accDisplay`n$Action"
				Write-Verbose -Message $msg
				if ($JobEntry) {
					$JobEntry['Status'] = $msg
					$JobEntry['Percent'] = $pct
					$JobEntry['LastUpdateTick'] = [DateTime]::Now.Ticks
					$JobEntry['ProfileName'] = $ProfileName
					$JobEntry['EntryNum'] = $EntryNum
				}
			}

			try
			{
				$totalClients = $Jobs.Count
				$totalGlobalSteps = $totalClients * $TotalStepsPerClient
				$currentClientIndex = 0
				$clientsLoggedInCount = 0

				if ($State -and -not $State['LoginGracePids']) { $State['LoginGracePids'] = [System.Collections.Hashtable]::Synchronized(@{}) }

				while ($LoginTasks.Count -gt 0)
				{
					$task = $LoginTasks.First.Value
					$LoginTasks.RemoveFirst()
					$job = $task.JobData

					$Global:CurrentActiveProcessId = 0
					$Global:CurrentActiveWindowHandle = [IntPtr]::Zero

					$currentClientIndex++
					$loginAttemptDisplay = $clientsLoggedInCount + 1

					try
					{
						CheckCancel; SleepWithCancel -Milliseconds 50
						$entryNumber = $job.EntryNumber
						$processId = $job.ProcessId
						$explicitHandle = if ($job.ExplicitHandle -and $job.ExplicitHandle -ne 0) { $job.ExplicitHandle } else { [IntPtr]::Zero }
						$Global:CurrentActiveProcessId = $processId
						$profileName = $job.ProfileName
						$clientTitle = $job.ClientTitle

						if ($processId -gt 0) {
							if ($JobPidSet) { [void]$JobPidSet.Add([int]$processId) }
							if ($State -and $State['ActiveLoginPids']) { [void]$State['ActiveLoginPids'].Add($processId) }
						}
					
						$logPath = $job.LogPath
						$config = $job.Config
						$stepBase = ($loginAttemptDisplay - 1) * $TotalStepsPerClient
						$currentStep = $stepBase
					
						if ($processId -eq 0) { throw "Process ID not found for Client $entryNumber" }
						$process = Get-Process -Id $processId -ErrorAction SilentlyContinue
						if (-not $process) { throw "Process $processId is gone." }

						$serverID = '1'; $channelID = '1'; $charSlot = '1'; $startCollector = 'No'
						$settingKey = "Client${entryNumber}_Settings"
						if ($config[$settingKey])
						{
							$parts = $config[$settingKey] -split ','
							if ($parts.Count -eq 4) { $serverID = $parts[0]; $channelID = $parts[1]; $charSlot = $parts[2]; $startCollector = $parts[3] }
						}
						
						$currentStep++; ReportLoginProgress -Action 'Starting Login Process...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle

						Write-LogWithRetry -FilePath $logPath -Value ''
						$workingHwnd = if ($explicitHandle -ne [IntPtr]::Zero) { $explicitHandle } else { $process.MainWindowHandle }
						$Global:CurrentActiveWindowHandle = $workingHwnd
						$currentStep++; ReportLoginProgress -Action 'Getting Client ready...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle

						RestoreWindowForInput; CheckCancel; EnsureWindowResponsive; CheckCancel; SleepWithCancel -Milliseconds 50

						$rect = New-Object Custom.Native+RECT
						[Custom.Native]::GetWindowRect($workingHwnd, [ref]$rect)
						$x = $rect.Left + 100; $y = $rect.Top + 100
					
						$currentStep++; ReportLoginProgress -Action "Log into Account $($entryNumber)..." -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle

						$firstNickCoords = ParseCoordinates $config['FirstNickCoords']
						$scrollDownCoords = ParseCoordinates $config['ScrollDownCoords']
						$scrollbaseX = if ($scrollDownCoords) { $rect.Left + $scrollDownCoords.X } else { [int](($rect.Left + $rect.Right) / 2) + 160 }
						$scrollbaseY = if ($scrollDownCoords) { $rect.Top + $scrollDownCoords.Y } else { [int](($rect.Top + $rect.Bottom) / 2) + 46 }

						if ($firstNickCoords)
						{
							$baseX = $rect.Left + $firstNickCoords.X
							$baseY = $rect.Top + $firstNickCoords.Y
							if ($entryNumber -ge 6 -and $entryNumber -le 10)
							{
								CheckCancel; EnsureWindowResponsive; CheckCancel
								Invoke-MouseClick -X ($scrollbaseX + 5) -Y ($scrollbaseY - 5); SleepWithCancel -Milliseconds 50
								$targetY = $baseY + (($entryNumber - 6) * 18)
								Invoke-MouseDoubleClick -X $baseX -Y $targetY; SleepWithCancel -Milliseconds 50
								Invoke-MouseDoubleClick -X $baseX -Y $targetY
							}
							elseif ($entryNumber -ge 1 -and $entryNumber -le 5)
							{
								$targetY = $baseY + (($entryNumber - 1) * 18)
								CheckCancel; EnsureWindowResponsive; CheckCancel
								Invoke-MouseDoubleClick -X $baseX -Y $targetY; SleepWithCancel -Milliseconds 50
								Invoke-MouseDoubleClick -X $baseX -Y $targetY
							}
						}
						else
						{
							$centerX = [int](($rect.Left + $rect.Right) / 2) + 25
							$centerY = [int](($rect.Top + $rect.Bottom) / 2) + 18
							$adjustedY = $centerY
							if ($entryNumber -ge 6 -and $entryNumber -le 10)
							{
								$yOffset = ($entryNumber - 8) * 18
								$adjustedY = $centerY + $yOffset
								Invoke-MouseClick -X ($centerX + 145) -Y ($centerY + 28); SleepWithCancel -Milliseconds 50
							}
							elseif ($entryNumber -ge 1 -and $entryNumber -le 5)
							{
								$yOffset = ($entryNumber - 3) * 18
								$adjustedY = $centerY + $yOffset
							}
							elseif ($entryNumber -ge 11) { return }
							CheckCancel; EnsureWindowResponsive; CheckCancel
							Invoke-MouseDoubleClick -X $centerX -Y $adjustedY; SleepWithCancel -Milliseconds 50
							Invoke-MouseDoubleClick -X $centerX -Y $adjustedY
							SleepWithCancel -Milliseconds 50
						}
						SleepWithCancel -Milliseconds 50
					
						
						CheckCancel; EnsureWindowResponsive; CheckCancel
						if ($config["Server${serverID}Coords"])
						{
							$coords = ParseCoordinates $config["Server${serverID}Coords"]
							if ($coords)
							{
								$currentStep++; ReportLoginProgress -Action "Selecting Server $serverID..." -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
								Invoke-MouseDoubleClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
								SleepWithCancel -Milliseconds 50
							}
						}
						SleepWithCancel -Milliseconds 50; CheckCancel; EnsureWindowResponsive; CheckCancel
					
						
						if ($config["Channel${channelID}Coords"])
						{
							$coords = ParseCoordinates $config["Channel${channelID}Coords"]
							if ($coords)
							{
								$currentStep++; ReportLoginProgress -Action "Selecting Channel $channelID..." -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
								Invoke-MouseDoubleClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
								SleepWithCancel -Milliseconds 50
							}
						}
						SleepWithCancel -Milliseconds 50; CheckCancel; EnsureWindowResponsive; CheckCancel
						$currentStep++; ReportLoginProgress -Action 'Entering Character Selection...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
						$loginClaim = Invoke-LoggedKeyPress -VirtualKeyCode 0x0D -LogPath $logPath -Marker '6 - LOGIN_PLAYER_LIST' -Required 1

						$currentStep++
						$loginReady = Wait-LogClaim -Claim $loginClaim -LogPath $logPath -TimeoutSeconds 25
						if (-not $loginReady) { throw "CERT Login Timeout for Client $entryNumber" }
						SleepWithCancel -Milliseconds 50; CheckCancel; EnsureWindowResponsive; CheckCancel
					
						
						$slotCoordKey = "Char${charSlot}Coords"
						if ($config[$slotCoordKey] -and $config[$slotCoordKey] -ne '0,0')
						{
							$coordsSlot = ParseCoordinates $config[$slotCoordKey]
							if ($coordsSlot)
							{
							if ($charSlot -ge 1 -and $charSlot -le 3)
							{
								if ($config["SwitchLeftCoords"] -and $config["SwitchLeftCoords"] -ne '0,0')
								{
									$coordsSwitch = ParseCoordinates $config["SwitchLeftCoords"]
									if ($coordsSwitch)
									{
										$currentStep++; ReportLoginProgress -Action 'Switching to Page 1...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
										Invoke-MouseClick -X ($rect.Left + $coordsSwitch.X) -Y ($rect.Top + $coordsSwitch.Y)
										SleepWithCancel -Milliseconds 200
									}
								}
							}
							elseif ($charSlot -ge 4 -and $charSlot -le 6)
								{
								if ($config["SwitchRightCoords"] -and $config["SwitchRightCoords"] -ne '0,0')
									{
									$coordsSwitch = ParseCoordinates $config["SwitchRightCoords"]
										if ($coordsSwitch)
										{
										$currentStep++; ReportLoginProgress -Action 'Switching to Page 2...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
											Invoke-MouseClick -X ($rect.Left + $coordsSwitch.X) -Y ($rect.Top + $coordsSwitch.Y)
											SleepWithCancel -Milliseconds 200
										}
									}
								}
								$currentStep++; ReportLoginProgress -Action "Selecting Character Slot $charSlot..." -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
								Invoke-MouseClick -X ($rect.Left + $coordsSlot.X) -Y ($rect.Top + $coordsSlot.Y)
								SleepWithCancel -Milliseconds 100
							}
						}
						CheckCancel; EnsureWindowResponsive; CheckCancel
						$worldMarker = '13 - CACHE_ACK_JOIN'; $worldRequired = 3
						if ($config['WorldLoadLogThreshold']) { $worldRequired = [int]$config['WorldLoadLogThreshold'] }
						if ($config['WorldLoadLogEntry']) { $worldMarker = $config['WorldLoadLogEntry'] }
						$worldClaim = Invoke-LoggedKeyPress -VirtualKeyCode 0x0D -LogPath $logPath -Marker $worldMarker -Required $worldRequired
						SleepWithCancel -Milliseconds 50

						$currentStep++; ReportLoginProgress -Action 'Waiting for World Load...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
						$cacheJoin = Wait-LogClaim -Claim $worldClaim -LogPath $logPath -TimeoutSeconds 60
						if (-not $cacheJoin) { throw "Main Login Timeout for Client $entryNumber" }
						CheckCancel; EnsureWindowResponsive; CheckCancel
					
						$delay = if ($config['PostLoginDelay']) { [int]$config['PostLoginDelay'] } else { 1 }
						$currentStep++; ReportLoginProgress -Action "Post Login Delay ($delay s)" -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
						SleepWithCancel -Milliseconds ($delay * 1000)
						CheckCancel; EnsureWindowResponsive; CheckCancel

						
						if ($startCollector -eq 'Yes' -and $config['CollectorStartCoords'] -and $config['CollectorStartCoords'] -ne '0,0')
						{
							$coords = ParseCoordinates $config['CollectorStartCoords']
							if ($coords)
							{
								$currentStep++; ReportLoginProgress -Action 'Starting Collector...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
								Invoke-MouseClick -X ($rect.Left + $coords.X) -Y ($rect.Top + $coords.Y)
								SleepWithCancel -Milliseconds 1000
							}
						}

						$currentStep++; ReportLoginProgress -Action 'Minimizing...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
						[Custom.Native]::SendMessage($workingHwnd, 0x0112, 0xF020, 0)
						$currentStep++; ReportLoginProgress -Action 'Optimizing...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
						[Custom.Native]::EmptyWorkingSet($process.Handle)
						$currentStep++; ReportLoginProgress -Action 'Finished...' -Step $currentStep -TotalSteps $totalGlobalSteps -ClientIdx $loginAttemptDisplay -ClientCount $totalClients -EntryNum $entryNumber -ProfileName $profileName -ServerID $serverID -ChannelID $channelID -ClientTitle $clientTitle
						
						$clientsLoggedInCount++
					}
					catch
					{
						if ($_.Exception.Message -eq 'LoginCancelled' -or $_.ToString() -eq 'LoginCancelled') { throw }

						Write-Verbose "LOGIN: Client $($job.EntryNumber) failed. Error: $($_.Exception.Message)"
						$processIdToKill = $job.ProcessId
						if ($processIdToKill -gt 0)
						{
							Write-Verbose "LOGIN: Killing unresponsive client PID $processIdToKill"
							try { Stop-Process -Id $processIdToKill -Force -ErrorAction SilentlyContinue } catch {}
						}
						
						$task.RetryCount++
						$relaunchSuccess = $false

						if ($task.RetryCount -le 2)
						{
							$relaunchGateAcquired = Enter-InputGate
							try
							{
								Write-Verbose "LOGIN: Attempting to relaunch client (Attempt $($task.RetryCount + 1))..."
								$lPath = $DashboardConfig.Config['LauncherPath']['LauncherPath']
								$pName = $DashboardConfig.Config['ProcessName']['ProcessName']
								
								if ($job.ProfileName -and $job.ProfileName -ne 'Default' -and $DashboardConfig.Config['Profiles'] -and $DashboardConfig.Config['Profiles'][$job.ProfileName])
								{
									$profPath = $DashboardConfig.Config['Profiles'][$job.ProfileName]
									if (Test-Path $profPath)
									{
										$exeName = [System.IO.Path]::GetFileName($lPath)
										$profLPath = Join-Path $profPath $exeName
										if (Test-Path $profLPath) { $lPath = $profLPath }
									}
								}

								if (Test-Path $lPath)
								{
									$existingPids = @(Get-Process -Name $pName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
									$lDir = [System.IO.Path]::GetDirectoryName($lPath)
									$lProc = Start-Process -FilePath $lPath -WorkingDirectory $lDir -PassThru -WindowStyle Normal
								
									if ($lProc) {
										$playCoords = $null
										if ($job.Config -and $job.Config['LauncherPlayCoords']) { $playCoords = $job.Config['LauncherPlayCoords'] }
										
										if ($playCoords -and $playCoords -ne '0,0') {
											$logDir = Join-Path $lDir "Log"
											$initialCount = 0
											if (Test-Path $logDir) {
												$logFile = Get-ChildItem -Path $logDir -Filter "launcher*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
												if ($logFile) { try { $s=[System.IO.File]::Open($logFile.FullName,'Open','Read','ReadWrite'); $r=New-Object System.IO.StreamReader($s); $c=$r.ReadToEnd(); $r.Close(); $s.Close(); $initialCount=([regex]::Matches($c, "Status: Update completed successfully!")).Count } catch {} }
											}
											
											$patchTimeout = [DateTime]::Now.AddSeconds(45)
											while ([DateTime]::Now -lt $patchTimeout) {
												if ($lProc.HasExited) { break }
												$success = $false
												if (Test-Path $logDir) {
													$logFile = Get-ChildItem -Path $logDir -Filter "launcher*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
													if ($logFile) { try { $s=[System.IO.File]::Open($logFile.FullName,'Open','Read','ReadWrite'); $r=New-Object System.IO.StreamReader($s); $c=$r.ReadToEnd(); $r.Close(); $s.Close(); if (([regex]::Matches($c, "Status: Update completed successfully!")).Count -gt $initialCount) { $success = $true } } catch {} }
												}
												
												if ($success) {
													try {
														$hWnd = $lProc.MainWindowHandle
														if ($hWnd -ne [IntPtr]::Zero -and ('Custom.Native' -as [Type])) {
															[Custom.Native]::ShowWindow($hWnd, 1); [Custom.Native]::SetForegroundWindow($hWnd); Start-Sleep -Milliseconds 500
															$rect = New-Object Custom.Native+RECT; [Custom.Native]::GetWindowRect($hWnd, [ref]$rect)
															$xy = $playCoords -split ','; 
															if ($xy.Count -eq 2) {
																$x = [int]$xy[0] + $rect.Left; $y = [int]$xy[1] + $rect.Top
																[Custom.Native]::SetCursorPos($x, $y); [Custom.Native]::mouse_event(0x02, 0, 0, 0, 0); [Custom.Native]::mouse_event(0x04, 0, 0, 0, 0)
															}
														}
													} catch {}
													break
												}
												Start-Sleep -Milliseconds 1000
											}
										}
									}
									$timeout = [DateTime]::Now.AddSeconds(60)
									while ([DateTime]::Now -lt $timeout)
									{
										CheckCancel
										$currentPids = @(Get-Process -Name $pName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
										$diff = $currentPids | Where-Object { $_ -notin $existingPids }
										if ($diff)
										{
											$newPid = $diff[0]
											$job.ProcessId = $newPid
											$job.ExplicitHandle = [IntPtr]::Zero
											
											$wTimeout = [DateTime]::Now.AddSeconds(30)
											while ([DateTime]::Now -lt $wTimeout)
											{
												CheckCancel
												try {
													$p = Get-Process -Id $newPid -ErrorAction SilentlyContinue
													if ($p -and $p.MainWindowHandle -ne [IntPtr]::Zero)
													{
														$job.ExplicitHandle = $p.MainWindowHandle
														$relaunchSuccess = $true
														break
													}
												} catch {}
												Start-Sleep -Milliseconds 500
											}
											break
										}
										Start-Sleep -Milliseconds 500
									}
								}
							}
							catch { Write-Verbose "LOGIN: Relaunch failed: $_" }
							finally { Exit-InputGate $relaunchGateAcquired }
						}

						if ($relaunchSuccess)
						{
							Write-Verbose "LOGIN: Relaunch successful. Retrying login..."
							if ($task.RetryCount -eq 1) { $LoginTasks.AddFirst($task) }
							else { $LoginTasks.AddLast($task) }
						}
						elseif ($task.RetryCount -eq 1)
						{
							Write-Verbose "LOGIN: Retrying client immediately (Attempt 2)"
							$LoginTasks.AddFirst($task)
						}
						elseif ($task.RetryCount -eq 2)
						{
							Write-Verbose "LOGIN: Retrying client at end of sequence (Attempt 3)"
							$LoginTasks.AddLast($task)
						}
						else
						{
							Write-Verbose "LOGIN: Client failed 3 times. Skipping."
						}
					}
				}
			}
			catch
			{
				if ($Global:CurrentActiveWindowHandle -ne [IntPtr]::Zero)
				{
					try { [Custom.Native]::ShowWindow($Global:CurrentActiveWindowHandle, 6) } catch {}
				}
				if ($_.Exception.Message -eq 'LoginCancelled' -or $_.ToString() -eq 'LoginCancelled')
				{
					Write-Verbose 'LOGIN: Background processing cancelled.'
				}
				else
				{
					throw $_
				}
			}
		}

		$completionEvent = {
			param($s, $e)
			$state = $e.InvocationStateInfo.State
			$mainForm = $global:DashboardConfig.UI.MainForm
			$completedJobId = $null
			try { if ($event -and $event.MessageData) { $completedJobId = [string]$event.MessageData } } catch {}
			if (-not $completedJobId) { try { if ($e -and $e.MessageData) { $completedJobId = [string]$e.MessageData } } catch {} }
			if ($state -eq 'Failed') { Write-Verbose "LOGIN EXCEPTION: $($e.InvocationStateInfo.Reason)" }
			if ($state -match 'Completed|Failed|Stopped')
			{
				if ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated)
				{
					try
					{
						$cleanupDelegate = [Action[object]] { param($id) CleanUpLoginResources -JobId ([string]$id) }
						$mainForm.BeginInvoke($cleanupDelegate, @([object]$completedJobId)) | Out-Null
					}
					catch { CleanUpLoginResources -JobId $completedJobId }
				}
				else { CleanUpLoginResources -JobId $completedJobId }
			}
		}

		
		$stepsVal = if ($TOTAL_STEPS_PER_CLIENT) { $TOTAL_STEPS_PER_CLIENT } else { 15 }
		$options = if ($global:DashboardConfig.Config) { $global:DashboardConfig.Config['Options'] } else { @{} }
		$loginCfg = if ($global:DashboardConfig.Config) { $global:DashboardConfig.Config['LoginConfig'] } else { @{} }

		$startedJobs = 0
		$localLoginJobIds = @()
		foreach ($taskToStart in @($loginTasks))
		{
			$jobId = [Guid]::NewGuid().ToString('N')
			$jobPidSet = [System.Collections.Generic.HashSet[int]]::new()
			$jobEntry = [hashtable]::Synchronized(@{
				JobId          = $jobId
				PowerShell     = $null
				AsyncResult    = $null
				EventSubId     = $null
				EventSub       = $null
				PidSet         = $jobPidSet
				JobData        = $taskToStart.JobData
				Status         = 'Starting Login Process...'
				Percent        = 0
				LastUpdateTick = [DateTime]::Now.Ticks
				StartTime      = [DateTime]::Now
			})

			$jobTasks = New-Object System.Collections.Generic.LinkedList[PSObject]
			$jobTasks.AddLast($taskToStart)

			$result = InvokeInManagedRunspace -RunspacePool $localRunspace -ScriptBlock $loginScript -AsJob -ArgumentList $jobTasks, $stepsVal, $options, $loginCfg, $global:LoginCancellation, $global:DashboardConfig.State, $jobId, $jobPidSet, $jobEntry
			if (-not ($result -and $result.PowerShell)) { throw "Failed to start login background job for $($taskToStart.JobData.ProfileName)" }

			$loginPS = $result.PowerShell
			$jobEventName = 'LoginOp_' + $jobId
			$eventSub = Register-ObjectEvent -InputObject $loginPS -EventName InvocationStateChanged -SourceIdentifier $jobEventName -MessageData $jobId -Action $completionEvent

			$jobEntry['PowerShell'] = $loginPS
			$jobEntry['AsyncResult'] = $result.AsyncResult
			$jobEntry['EventSubId'] = $jobEventName
			$jobEntry['EventSub'] = $eventSub
			$global:LoginJobs[$jobId] = $jobEntry
			$localLoginJobIds += $jobId
			$startedJobs++

			Write-Verbose "LOGIN: Started simultaneous login job for $($taskToStart.JobData.ProfileName) ($($taskToStart.JobData.ClientTitle))."

			if ($loginPS.InvocationStateInfo.State -match 'Completed|Failed|Stopped')
			{
				CleanUpLoginResources -JobId $jobId
			}
		}

		if ($startedJobs -eq 0) { throw 'Failed to start login background jobs' }

		if (-not $global:DashboardConfig.Resources.Timers.Contains('loginUITimer'))
		{
			$uiUpdateTimer = New-Object System.Windows.Forms.Timer
			$uiUpdateTimer.Interval = 100
			$uiUpdateTimer.Add_Tick({
				try {
					$activeLogins = @()
					[System.Threading.Monitor]::Enter($global:LoginJobs.SyncRoot)
					try { $activeLogins = @($global:LoginJobs.Values) } finally { [System.Threading.Monitor]::Exit($global:LoginJobs.SyncRoot) }
					if ($activeLogins.Count -gt 0)
					{
						$latest = $activeLogins | Sort-Object @{ Expression = { [long]$_.LastUpdateTick } } -Descending | Select-Object -First 1
						$tick = [long]$latest.LastUpdateTick
						if ($tick -gt $script:LastLoginToastTick)
						{
							$script:LastLoginToastTick = $tick
							$percentSum = 0
							foreach ($activeLogin in $activeLogins) { $percentSum += [int]$activeLogin.Percent }
							$percentAvg = [int]($percentSum / [Math]::Max(1, $activeLogins.Count))
							$headline = if ($activeLogins.Count -gt 1) { "Logging in $($activeLogins.Count) clients..." } else { 'Login Progress' }
							ShowToast -Title 'Login Progress' -Message "$headline`n$($latest.Status)" -Type 'Info' -Key 9999 -TimeoutSeconds 0 -Progress $percentAvg
						}
					}
				} catch {}
			})
			$uiUpdateTimer.Start()
			$global:DashboardConfig.Resources.Timers['loginUITimer'] = $uiUpdateTimer
		}

		Write-Verbose 'LOGIN: Background processes started.'

	}
 catch
	{
		Write-Verbose "LOGIN: Error starting login sequence: $_"
		if ($localLoginJobIds -and $localLoginJobIds.Count -gt 0)
		{
			CleanUpLoginResources -globalLoginResourcesRef $global:DashboardConfig -JobId $localLoginJobIds
		}
		elseif ($global:LoginJobs.Count -eq 0)
		{
			$global:DashboardConfig.State['LoginActive'] = $false
		}
	}
}

function CleanUpLoginResources
{
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		$globalLoginResourcesRef,
		[Parameter(Mandatory = $false)]
		[string[]]$JobId,
		[Parameter(Mandatory = $false)]
		[switch]$All,
		[Parameter(Mandatory = $false)]
		[switch]$Sync
	)

	
	if (-not $global:LoginJobs) { $global:LoginJobs = [hashtable]::Synchronized(@{}) }

	$jobIdsToClean = @()
	if ($JobId -and $JobId.Count -gt 0) { $jobIdsToClean = @($JobId) }
	else
	{
		try
		{
			$jobsLock = $global:LoginJobs.SyncRoot
			[System.Threading.Monitor]::Enter($jobsLock)
			try { $jobIdsToClean = @($global:LoginJobs.Keys) }
			finally { [System.Threading.Monitor]::Exit($jobsLock) }
		}
		catch { $jobIdsToClean = @($global:LoginJobs.Keys) }
	}

	$jobSnapshots = @()
	$gracedPids = @()
	foreach ($cleanJobId in $jobIdsToClean)
	{
		$jobSnapshot = $null
		try
		{
			$jobsLock = $global:LoginJobs.SyncRoot
			[System.Threading.Monitor]::Enter($jobsLock)
			try
			{
				if ($global:LoginJobs.ContainsKey($cleanJobId)) { $jobSnapshot = $global:LoginJobs[$cleanJobId] }
				$global:LoginJobs.Remove($cleanJobId) | Out-Null
			}
			finally { [System.Threading.Monitor]::Exit($jobsLock) }
		}
		catch {}

		if ($null -eq $jobSnapshot) { continue }
		$jobSnapshots += $jobSnapshot
		if ($jobSnapshot.PidSet) { $gracedPids += @($jobSnapshot.PidSet) }
		elseif ($jobSnapshot.JobData -and $jobSnapshot.JobData.ProcessId -gt 0) { $gracedPids += [int]$jobSnapshot.JobData.ProcessId }
	}

	Write-Verbose "LOGIN: Cleaning up $($jobSnapshots.Count) finished login job(s); $($global:LoginJobs.Count) still running."

	if ($global:LoginJobs.Count -gt 0)
	{
		Write-Verbose 'LOGIN: Other client logins are still running. Shared resources stay active.'
		return
	}

	$global:LoginResources['IsStopping'] = $true
	Write-Verbose 'LOGIN: Cleaning up resources...'

	if (-not $globalLoginResourcesRef) { $globalLoginResourcesRef = $global:DashboardConfig }

	
	if (-not $global:DashboardConfig.State.LoginGracePids) { $global:DashboardConfig.State.LoginGracePids = [System.Collections.Hashtable]::Synchronized(@{}) }
	foreach ($pidToGrace in @($gracedPids | Select-Object -Unique))
	{
		if ($null -eq $pidToGrace) { continue }
		$pidValue = [int]$pidToGrace
		if ($pidValue -le 0) { continue }
		try { $global:DashboardConfig.State.LoginGracePids[$pidValue] = [DateTime]::Now.AddSeconds(120) } catch {}
		if ($global:DashboardConfig.State.ActiveLoginPids) { try { [void]$global:DashboardConfig.State.ActiveLoginPids.Remove($pidValue) } catch {} }
	}

	
	if ($global:LoginResources.IsMouseHookActive)
	{
		try
		{
			[Custom.MouseHookManager]::Stop()
			$global:LoginResources.IsMouseHookActive = $false
			Write-Verbose 'LOGIN: Mouse hook stopped.'
		}
		catch
		{
			Write-Error "LOGIN: Failed to stop mouse hook: $_" 
		}
	}

	
	if (-not ($global:DashboardConfig.State.ReconnectQueue -and $global:DashboardConfig.State.ReconnectQueue.Count -gt 0))
	{
		CloseToast -Key 9999
	}

	if ($global:DashboardConfig.Resources.Timers.Contains('loginUITimer'))
	{
		try
		{
			$global:DashboardConfig.Resources.Timers['loginUITimer'].Stop()
			$global:DashboardConfig.Resources.Timers['loginUITimer'].Dispose()
			$global:DashboardConfig.Resources.Timers.Remove('loginUITimer')
		}
		catch {}
	}
	
	$uiCleanupAction = {
		$ui = $global:DashboardConfig.UI
		if ($ui.LoginButton)
		{
			$ui.LoginButton.BackColor = Get-Theme 'ButtonBg'
			$ui.LoginButton.Text = 'Login'
		}
        
		if ($global:DashboardConfig.State)
		{
			$global:DashboardConfig.State['LoginActive'] = $false
		}
	}
	
	$mainForm = $global:DashboardConfig.UI.MainForm
    
	
	
	if ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated -and $mainForm.InvokeRequired)
	{
		try
		{ 
			
			$mainForm.BeginInvoke($uiCleanupAction)
		}
		catch
		{
			Write-Warning "LOGIN: Failed to BeginInvoke UI cleanup action: $_" 
			
			& $uiCleanupAction $globalLoginResourcesRef
		}
	}
 	elseif ($mainForm -and -not $mainForm.IsDisposed -and $mainForm.IsHandleCreated)
	{
		
		& $uiCleanupAction $globalLoginResourcesRef
	}
 	else
	{
		
		Write-Warning 'LOGIN: UI form not available for cleanup, skipping UI updates.'
		$global:DashboardConfig.State['LoginActive'] = $false 
	}

	
	$localRes = $global:LoginResources.Clone()
	$localRes.JobSnapshots = $jobSnapshots
	$localRes.EventSubscriptionId = @($jobSnapshots | ForEach-Object { $_.EventSubId } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })

	
	if ($localRes.EventSubscriptionId)
	{
		foreach ($subscriptionId in @($localRes.EventSubscriptionId))
		{
			if ([string]::IsNullOrWhiteSpace([string]$subscriptionId)) { continue }
			Unregister-Event -SourceIdentifier ([string]$subscriptionId) -ErrorAction SilentlyContinue
		}
		Write-Verbose "LOGIN: Unregistered InvocationStateChanged event: $($localRes.EventSubscriptionId -join ', ')"
	}
	if ($localRes.InfoSubscription)
	{
		Unregister-Event -SubscriptionId $localRes.InfoSubscription.Id -ErrorAction SilentlyContinue
		Write-Verbose "LOGIN: Unregistered Information.DataAdded event (ID: $($localRes.InfoSubscription.Id))" 
	}

	$disposeAction = {
		try
		{
			foreach ($snapshot in @($localRes.JobSnapshots))
			{
				if ($null -eq $snapshot) { continue }
				$ps = $snapshot.PowerShell; $ar = $snapshot.AsyncResult
				if ($ps)
				{
					try { if ($ps.InvocationStateInfo.State -eq 'Running') { $ps.Stop() } } catch {}
					if ($ar) { try { $ps.EndInvoke($ar) } catch {} }
					try { $ps.Dispose() } catch {}
				}
			}
			$rs = $localRes.Runspace
			if ($rs) { try { $rs.Dispose() } catch {} }
		} catch {}
	}

	if ($Sync) { & $disposeAction }
	else { [System.Threading.Tasks.Task]::Run([Action]$disposeAction) | Out-Null }
	
	$global:LoginResources = @{
		PowerShellInstance  = $null
		Runspace            = $null
		EventSubscriptionId = $null
		EventSubscriber     = $null
		AsyncResult         = $null
		IsStopping          = $false
		InfoSubscription    = $null
		IsMouseHookActive   = $false
	}
}
#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion