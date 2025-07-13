<# launch.psm1 
	.SYNOPSIS
		Client Launcher Module for Entropia Dashboard.

	.DESCRIPTION
		This module provides functionality to launch and manage Entropia Universe game clients:
		- Launches multiple game clients based on configuration
		- Monitors client processes and window states
		- Provides thread-safe logging of launch operations
		- Handles cleanup of resources when operations complete

	.NOTES
		Author: Immortal / Divine
		Version: 1.1
		Requires: PowerShell 5.1, .NET Framework 4.5+, classes.psm1
#>

#region Configuration and Constants

# Default launcher timeout in seconds
$script:LauncherTimeout = 30

# Default delay between client launches in seconds
$script:LaunchDelay = 5

# Maximum retry attempts
$script:MaxRetryAttempts = 3

$script:ProcessConfig = @{
	MaxRetries = 3
	RetryDelay = 500 # milliseconds
	Timeout    = 30000 # milliseconds
}

#endregion Configuration and Constants

#region Launch Management Functions

function Start-ClientLaunch
{
	<#
	.SYNOPSIS
	Initializes and starts the client launch process.
	
	.DESCRIPTION
	Prepares and starts the asynchronous launch operation for Entropia clients.
	#>
	[CmdletBinding()]
	param()
	
	# Prevent multiple launches
	if ($global:DashboardConfig.State.LaunchActive)
	{
		[System.Windows.Forms.MessageBox]::Show('Launch operation already in progress', 'Information',
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
		return
	}
	
	$global:DashboardConfig.State.LaunchActive = $true
	
	# Create a new ordered dictionary for settings
	$settingsDict = [ordered]@{}
	
	# Copy settings from global config
	foreach ($section in $global:DashboardConfig.Config.Keys)
	{
		$settingsDict[$section] = [ordered]@{}
		foreach ($key in $global:DashboardConfig.Config[$section].Keys)
		{
			$settingsDict[$section][$key] = $global:DashboardConfig.Config[$section][$key]
		}
	}
	
	# Ensure latest config is loaded
	Read-Config
	
	# Get required settings
	$neuzName = $settingsDict['ProcessName']['ProcessName']
	$launcherPath = $settingsDict['LauncherPath']['LauncherPath']
	
	Write-Verbose "LAUNCH: Using ProcessName: $neuzName" -ForegroundColor DarkGray
	Write-Verbose "LAUNCH: Using LauncherPath: $launcherPath" -ForegroundColor DarkGray
	
	# Validate settings
	if ([string]::IsNullOrEmpty($neuzName) -or [string]::IsNullOrEmpty($launcherPath) -or -not (Test-Path $launcherPath))
	{
		Write-Verbose 'LAUNCH: Invalid launch settings' -ForegroundColor Yellow
		$global:DashboardConfig.State.LaunchActive = $false
		return
	}
	
	# Get max clients setting
	$maxClients = 1
	if ($settingsDict['MaxClients'].Contains('MaxClients') -and -not [string]::IsNullOrEmpty($settingsDict['MaxClients']['MaxClients']))
	{
		$maxClients = [int]($settingsDict['MaxClients']['MaxClients'])
	}
	
	# Prepare runspace for background processing
	$localRunspace = $null
	try
	{
		# Create a runspace with ApartmentState.STA to ensure proper UI interaction
		$localRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 1)
		$localRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
		$localRunspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
		$localRunspace.Open()
	}
	catch
	{
		Write-Verbose "LAUNCH: Error creating runspace: $_" -ForegroundColor Red
		$global:DashboardConfig.State.LaunchActive = $false
		return
	}
	
	$launchPS = [PowerShell]::Create()
	$launchPS.RunspacePool = $localRunspace
	
	# Store references globally for cleanup
	$global:LaunchResources = @{
		PowerShellInstance  = $launchPS
		Runspace            = $localRunspace
		EventSubscriptionId = $null
		EventSubscriber     = $null
		AsyncResult         = $null
		StartTime           = [DateTime]::Now
	}
	
	$launchPS.AddScript({
			param(
				$Settings,
				$LauncherPath,
				$NeuzName,
				$MaxClients,
				$IniPath
			)
		
			try
			{
				Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
			
				$classes = @'
				using System;
			
				public class ColorWriter
				{
					public static void WriteColored(string message, string color)
					{
						ConsoleColor originalColor = Console.ForegroundColor;
						
						switch(color.ToLower())
						{
							case "darkgray":
								Console.ForegroundColor = ConsoleColor.DarkGray;
								break;
							case "yellow":
								Console.ForegroundColor = ConsoleColor.Yellow;
								break;
							case "red":
								Console.ForegroundColor = ConsoleColor.Red;
								break;
							case "cyan":
								Console.ForegroundColor = ConsoleColor.Cyan;
								break;
							case "green":
								Console.ForegroundColor = ConsoleColor.Green;
								break;
							default:
								Console.ForegroundColor = ConsoleColor.DarkGray;
								break;
						}
						
						Console.WriteLine(message);
						Console.ForegroundColor = originalColor;
					}
				}
'@
			
				Add-Type -TypeDefinition $classes -Language 'CSharp'
			
				# Create a function that overrides the built-in Write-Verbose
				function Write-Verbose
				{
					[CmdletBinding()]
					param(
						[Parameter(Mandatory = $true, Position = 0)]
						[string]$Object,
				
						[Parameter()]
						[ValidateSet('darkgray', 'yellow', 'red', 'cyan', 'green')]
						[string]$ForegroundColor = 'darkgray'
					)
				
					[ColorWriter]::WriteColored($Object, $ForegroundColor)
				}
			
				$launcherDir = [System.IO.Path]::GetDirectoryName($LauncherPath)
				$launcherName = [System.IO.Path]::GetFileNameWithoutExtension($LauncherPath)
			
				Write-Verbose 'LAUNCH: Checking clients...' -ForegroundColor DarkGray
				$currentClients = @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue)
				$pCount = $currentClients.Count
			
				if ($pCount -gt 0)
				{
					Write-Verbose "LAUNCH: Found $pCount client(s)" -ForegroundColor DarkGray
				}
				else
				{
					Write-Verbose 'LAUNCH: No clients found' -ForegroundColor DarkGray
				}
			
				if ($pCount -ge $MaxClients)
				{
					Write-Verbose "LAUNCH: Max reached: $MaxClients" -ForegroundColor DarkGray
					return
				}
			
				$clientsToLaunch = $MaxClients - $pCount
				Write-Verbose "LAUNCH: Launching $clientsToLaunch more" -ForegroundColor Cyan
			
				# Store only the IDs, not the process objects
				$existingPIDs = $currentClients | Select-Object -ExpandProperty Id
				if ($existingPIDs.Count -gt 0)
				{
					Write-Verbose "LAUNCH: Tracking PIDs: $($existingPIDs -join ',')" -ForegroundColor DarkGray
				}
			
				# Clear process objects to avoid keeping references
				$currentClients = $null
			
				for ($attempt = 1; $attempt -le $clientsToLaunch; $attempt++)
				{
				
					Write-Verbose "LAUNCH: Client $attempt/$clientsToLaunch" -ForegroundColor Cyan
				
					# Check if launcher is already running
					$launcherRunning = $null -ne (Get-Process -Name $launcherName -ErrorAction SilentlyContinue)
					if ($launcherRunning)
					{
						Write-Verbose 'LAUNCH: Launcher running' -ForegroundColor DarkGray
						Write-Verbose 'LAUNCH: Waiting (30s timeout)' -ForegroundColor DarkGray
						$launcherTimeout = New-TimeSpan -Seconds 30
						$launcherStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
						$progressReported = @(5, 10, 15, 20, 25)
					
						while ($null -ne (Get-Process -Name $launcherName -ErrorAction SilentlyContinue))
						{
							$elapsedSeconds = [int]$launcherStopwatch.Elapsed.TotalSeconds
						
							if ($elapsedSeconds -in $progressReported)
							{
								$progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
								Write-Verbose "LAUNCH: Waiting. ($elapsedSeconds s / 30)" -ForegroundColor DarkGray
							}
						
							if ($launcherStopwatch.Elapsed -gt $launcherTimeout)
							{
								Write-Verbose 'LAUNCH: Timeout - killing' -ForegroundColor Yellow
								try
								{
									Stop-Process -Name $launcherName -Force -ErrorAction SilentlyContinue
									Write-Verbose 'LAUNCH: Terminated' -ForegroundColor DarkGray
								}
								catch
								{
									Write-Verbose "LAUNCH: $($_.Exception.Message)" -ForegroundColor Red
								}
								Start-Sleep -Seconds 1
								break
							}
							Start-Sleep -Milliseconds 500
						}
						Write-Verbose 'LAUNCH: Launcher closed' -ForegroundColor DarkGray
					}
				
					# Start launcher
					Write-Verbose 'LAUNCH: Starting launcher' -ForegroundColor DarkGray
					$launcherProcess = Start-Process -FilePath $LauncherPath -WorkingDirectory $launcherDir -PassThru
					$launcherPID = $launcherProcess.Id
					Write-Verbose "LAUNCH: PID: $launcherPID" -ForegroundColor DarkGray
				
					# Release the process object immediately
					$launcherProcess = $null
				
					Write-Verbose 'LAUNCH: Initializing...' -ForegroundColor DarkGray
					Start-Sleep -Seconds 1
				
					# Monitor launcher process
					$timeout = New-TimeSpan -Minutes 2
					$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
					$launcherClosed = $false
					$launcherClosedNormally = $false
					$progressReported = @(1, 5, 15, 30, 60, 90, 120)
				
					Write-Verbose 'LAUNCH: Monitoring (2min)' -ForegroundColor DarkGray
					while (-not $launcherClosed -and $stopwatch.Elapsed -lt $timeout)
					{
						$elapsedSeconds = [int]$stopwatch.Elapsed.TotalSeconds
						if ($elapsedSeconds -in $progressReported)
						{
							$progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
							Write-Verbose "LAUNCH: Close dashboard if patching! Long Patching operations may be canceled by the dashboard ($elapsedSeconds s / 120)" -ForegroundColor Yellow
						}
					
						# Get process info without storing the object
						$launcherExists = $false
						$launcherResponding = $true
					
						try
						{
							$tempProcess = Get-Process -Id $launcherPID -ErrorAction SilentlyContinue
							if ($tempProcess)
							{
								$launcherExists = $true
								$launcherResponding = $tempProcess.Responding
								# Release the reference immediately
								$tempProcess = $null
							}
						}
						catch
						{
							$launcherExists = $false
						}
					
						if (-not $launcherExists)
						{
							$launcherClosed = $true
							$launcherClosedNormally = $true
						}
						else
						{
							if (-not $launcherResponding)
							{
								$elapsedTime = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
								Write-Verbose "LAUNCH: Not responding $elapsedTime" -ForegroundColor Yellow
								try
								{
									Stop-Process -Id $launcherPID -Force -ErrorAction SilentlyContinue
									Write-Verbose 'LAUNCH: Terminated' -ForegroundColor Red
								}
								catch
								{
									Write-Verbose "LAUNCH: $($_.Exception.Message)" -ForegroundColor Red
								}
								$launcherClosed = $true
								$launcherClosedNormally = $false
							}
						}
					
						if (-not $launcherClosed)
						{
							Start-Sleep -Milliseconds 500
						}
					}
				
					if (-not $launcherClosed)
					{
						Write-Verbose 'LAUNCH: Timeout - killing' -ForegroundColor Red
						try
						{
							Stop-Process -Id $launcherPID -Force -ErrorAction SilentlyContinue
							Write-Verbose 'LAUNCH: Terminated' -ForegroundColor Red
						}
						catch
						{
							Write-Verbose "LAUNCH: $($_.Exception.Message)" -ForegroundColor Red
						}
						$launcherClosedNormally = $false
					}
				
					# Wait for client to start
					$clientStarted = $false
					$newClientPID = 0
					$stopwatch.Restart()
					$clientDetectionTimeout = New-TimeSpan -Seconds 30
					$progressReported = @(5, 10, 15, 20, 25)
				
					Write-Verbose 'LAUNCH: Waiting for client' -ForegroundColor DarkGray
					while (-not $clientStarted -and $stopwatch.Elapsed -lt $clientDetectionTimeout)
					{
						$elapsedSeconds = [int]$stopwatch.Elapsed.TotalSeconds
						if ($elapsedSeconds -in $progressReported)
						{
							$progressReported = $progressReported | Where-Object { $_ -ne $elapsedSeconds }
							Write-Verbose "LAUNCH: Waiting. ($elapsedSeconds s / 30)" -ForegroundColor DarkGray
						}
					
						# Get current client PIDs without keeping process objects
						$currentPIDs = @(Get-Process -Name $NeuzName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
						$newPIDs = $currentPIDs | Where-Object { $_ -notin $existingPIDs }
					
						if ($newPIDs.Count -gt 0)
						{
							try
							{
								# Get the newest client by getting process info and sorting, but don"t keep references
								$tempNewClients = @(Get-Process -Id $newPIDs -ErrorAction SilentlyContinue)
								$tempNewClient = $tempNewClients | Sort-Object StartTime -Descending | Select-Object -First 1
							
								if ($tempNewClient)
								{
									$newClientPID = $tempNewClient.Id
									$clientStarted = $true
									$existingPIDs += $newClientPID
									$elapsedTime = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
									Write-Verbose "LAUNCH: Client started PID: $newClientPID" -ForegroundColor DarkGray
								
									# Release references
									$tempNewClient = $null
									$tempNewClients = $null
								}
							}
							catch
							{
								Write-Verbose "LAUNCH: Error sorting clients: $($_.Exception.Message)" -ForegroundColor Red
								# Try without sorting
								if ($newPIDs.Count -gt 0)
								{
									$newClientPID = $newPIDs[0]
									$clientStarted = $true
									$existingPIDs += $newClientPID
									Write-Verbose "LAUNCH: Using client PID: $newClientPID (fallback)" -ForegroundColor Yellow
								}
							}
						
							if ($clientStarted)
							{
								# Wait for window to be ready
								$windowReady = $false
								$innerTimeout = New-TimeSpan -Seconds 30
								$innerStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
								$innerProgressReported = @(5, 10, 15, 20, 25)
							
								Write-Verbose 'LAUNCH: Waiting for window' -ForegroundColor DarkGray
								while (-not $windowReady -and $innerStopwatch.Elapsed -lt $innerTimeout)
								{
									$innerElapsedSeconds = [int]$innerStopwatch.Elapsed.TotalSeconds
									if ($innerElapsedSeconds -in $innerProgressReported)
									{
										$innerProgressReported = $innerProgressReported | Where-Object { $_ -ne $innerElapsedSeconds }
										Write-Verbose "LAUNCH: Waiting. ($innerElapsedSeconds s / 30)" -ForegroundColor DarkGray
									}
								
									# Check client without keeping references
									$clientExists = $false
									$clientResponding = $false
									$clientWindowHandle = [IntPtr]::Zero
									$clientHandle = [IntPtr]::Zero
								
									try
									{
										$tempClient = Get-Process -Id $newClientPID -ErrorAction SilentlyContinue
										if ($tempClient)
										{
											$clientExists = $true
											$clientResponding = $tempClient.Responding
											$clientWindowHandle = $tempClient.MainWindowHandle
											$clientHandle = $tempClient.Handle
										
											# Store window handle in a separate variable and release process object
											$tempClient = $null
										}
									}
									catch
									{
										$clientExists = $false
									}
								
									if (-not $clientExists)
									{
										$elapsedTime = [math]::Round($innerStopwatch.Elapsed.TotalSeconds, 1)
										Write-Verbose 'LAUNCH: Client terminated' -ForegroundColor Red
										$clientStarted = $false
										$clientResponding = $false
										$windowReady = $false
										break
									}
								
									if ($clientResponding -and $clientWindowHandle -ne [IntPtr]::Zero)
									{
										$windowReady = $true
										$elapsedTime = [math]::Round($innerStopwatch.Elapsed.TotalSeconds, 1)
										Write-Verbose 'LAUNCH: Window ready' -ForegroundColor DarkGray
									
										Write-Verbose 'LAUNCH: Minimizing...' -ForegroundColor DarkGray
										Start-Sleep -Milliseconds 500
										[Native]::ShowWindow($clientWindowHandle, [Native]::SW_MINIMIZE)
									
										Write-Verbose 'LAUNCH: Optimizing...' -ForegroundColor Cyan
										try
										{
											[Native]::EmptyWorkingSet($clientHandle)
										}
										catch
										{
											Write-Verbose 'LAUNCH: Opt failed' -ForegroundColor Red
										}
									
										Write-Verbose "LAUNCH: Client ready: $newClientPID" -ForegroundColor Green
									}
								
									Start-Sleep -Milliseconds 500
								}
							
								if (-not $windowReady)
								{
									Write-Verbose 'LAUNCH: Window not responsive' -ForegroundColor Red
									$clientStarted = $false
									break
								}
							}
						}
					
						Start-Sleep -Milliseconds 500
					}
				
					if (-not $clientStarted)
					{
						if (-not $launcherClosedNormally)
						{
							Write-Verbose 'LAUNCH: Launcher failed' -ForegroundColor Red
						}
						else
						{
							Write-Verbose 'LAUNCH: No client detected' -ForegroundColor Yellow
						}
					}
				
					# Check client count without keeping references
					$pCount = (Get-Process -Name $NeuzName -ErrorAction SilentlyContinue | Measure-Object).Count
					if ($pCount -ge $MaxClients)
					{
						Write-Verbose "LAUNCH: Max reached: $MaxClients" -ForegroundColor Yellow
						break
					}
				
					Start-Sleep -Seconds 2
				}
			
			}
			catch
			{
				Write-Verbose "LAUNCH: $($_.Exception.Message)" -ForegroundColor Red
			}
			finally
			{
			
				# Clear any remaining references to processes
				$existingPIDs = $null
				[System.GC]::Collect()
			}
		}).AddArgument($settingsDict).AddArgument($launcherPath).AddArgument($neuzName).AddArgument($maxClients).AddArgument($global:DashboardConfig.Paths.Ini)
	
	# Event registration and cleanup section
	try
	{
		# First, define the completion action that will be called when the operation completes
		$completionScriptBlock = {
			Write-Verbose 'LAUNCH: Launch operation completed, processing messages' -ForegroundColor Green
			
			# Call cleanup function
			Write-Verbose 'LAUNCH: Calling Stop-ClientLaunch' -ForegroundColor Cyan
			Stop-ClientLaunch
		}
		
		# Store the completion action in a script-level variable
		$script:LaunchCompletionAction = $completionScriptBlock
		
		# Create a unique event name
		$eventName = 'LaunchOperation_' + [Guid]::NewGuid().ToString('N')
		Write-Verbose "LAUNCH: Creating event with name: $eventName" -ForegroundColor DarkGray
		
		# Create a simple scriptblock for the event that doesn"t reference any variables
		$simpleEventAction = {
			param($src, $e)
			
			# Only process completed, failed, or stopped states
			$state = $e.InvocationStateInfo.State
			if ($state -eq 'Completed' -or $state -eq 'Failed' -or $state -eq 'Stopped')
			{
				Write-Verbose "LAUNCH: PowerShell operation state changed to: $state" -ForegroundColor DarkGray
				
				# Call cleanup directly - don"t reference any global variables
				if (Get-Command -Name Stop-ClientLaunch -ErrorAction SilentlyContinue)
				{
					Write-Verbose 'LAUNCH: Calling Stop-ClientLaunch from event handler' -ForegroundColor Cyan
					Stop-ClientLaunch
				}
				else
				{
					Write-Verbose 'LAUNCH: Stop-ClientLaunch function not found' -ForegroundColor Red
				}
			}
		}
		
		# Register the event with minimal dependencies
		Write-Verbose 'LAUNCH: Registering event handler' -ForegroundColor Cyan
		$eventSub = Register-ObjectEvent -InputObject $launchPS -EventName InvocationStateChanged -SourceIdentifier $eventName -Action $simpleEventAction
		
		# Verify event registration
		if ($null -eq $eventSub)
		{
			throw 'Failed to register event subscriber'
		}
		
		# Store event information
		$global:LaunchResources.EventSubscriptionId = $eventName
		$global:LaunchResources.EventSubscriber = $eventSub
		
		# Set up a safety timer
		Write-Verbose 'LAUNCH: Setting up safety timer' -ForegroundColor DarkGray
		$safetyTimer = New-Object System.Timers.Timer
		$safetyTimer.Interval = 300000  # 5 minutes
		$safetyTimer.AutoReset = $false
		
		# Create a simple timer elapsed handler
		$safetyTimer.Add_Elapsed({
				Write-Verbose 'LAUNCH: Safety timer elapsed' -ForegroundColor DarkGray
			
				# Call cleanup directly - don"t reference any global variables
				if (Get-Command -Name Stop-ClientLaunch -ErrorAction SilentlyContinue)
				{
					Write-Verbose 'LAUNCH: Calling Stop-ClientLaunch from safety timer' -ForegroundColor DarkGray
					Stop-ClientLaunch
				}
				else
				{
					Write-Verbose 'LAUNCH: Stop-ClientLaunch function not found' -ForegroundColor Red
				}
			})
		
		# Start the timer
		$safetyTimer.Start()
		$global:DashboardConfig.Resources.Timers['launchSafetyTimer'] = $safetyTimer
		
		# Start the async operation
		Write-Verbose 'LAUNCH: Starting async operation' -ForegroundColor Cyan
		$asyncResult = $launchPS.BeginInvoke()
		
		# Verify async operation
		if ($null -eq $asyncResult)
		{
			throw 'Failed to start async operation'
		}
		
		# Store the async result
		$global:LaunchResources.AsyncResult = $asyncResult

		# Simple property change handler that updates button appearance
		if ($global:DashboardConfig.State.LaunchActive) {
			$global:DashboardConfig.UI.Launch.FlatStyle = 'Popup'
		} else {
			$global:DashboardConfig.UI.Launch.FlatStyle = 'Flat'
		}
		
		Write-Verbose 'LAUNCH: Launch operation started successfully' -ForegroundColor Green
	}
	catch
	{
		Write-Verbose "LAUNCH: Error in launch setup: $_" -ForegroundColor Red
		
		# Call cleanup
		Stop-ClientLaunch
	}
}

function Stop-ClientLaunch
{
	<#
	.SYNOPSIS
	Cleans up resources used by the launch operation.
	
	.DESCRIPTION
	Ensures proper cleanup of all resources used during the launch operation.
	#>
	[CmdletBinding()]
	param()
	
	Write-Verbose 'LAUNCH: Cleaning up launch resources' -ForegroundColor Cyan
	
	try
	{
		# Set launch in progress flag to false
		$global:DashboardConfig.State.LaunchActive = $false
		
		# Check if we have resources to clean up
		if ($null -eq $global:LaunchResources)
		{
			Write-Verbose '  LAUNCH: No launch resources to clean up' -ForegroundColor DarkGray
			return
		}
		
		# Unregister event subscription if it exists
		if ($null -ne $global:LaunchResources.EventSubscriptionId)
		{
			try
			{
				Write-Verbose "LAUNCH: Unregistering event subscription: $($global:LaunchResources.EventSubscriptionId)" -ForegroundColor DarkGray
				Unregister-Event -SourceIdentifier $global:LaunchResources.EventSubscriptionId -Force -ErrorAction SilentlyContinue
			}
			catch
			{
				Write-Verbose "LAUNCH: Failed to unregister event: $_" -ForegroundColor Red
			}
		}
		
		# Remove event subscriber if it exists
		if ($null -ne $global:LaunchResources.EventSubscriber)
		{
			try
			{
				Write-Verbose 'LAUNCH: Removing event subscriber' -ForegroundColor DarkGray
				# Alternative approach: Use Unregister-Event with the subscription ID
				if ($null -ne $global:LaunchResources.EventSubscriptionId)
				{
					Unregister-Event -SourceIdentifier $global:LaunchResources.EventSubscriptionId -ErrorAction SilentlyContinue
				}
			}
			catch
			{
				Write-Verbose "LAUNCH: Failed to remove event subscriber: $_" -ForegroundColor Red
			}
		}
		
		# Stop and dispose PowerShell instance if it exists
		if ($null -ne $global:LaunchResources.PowerShellInstance)
		{
			try
			{
				Write-Verbose 'LAUNCH: Stopping PowerShell instance' -ForegroundColor DarkGray
				if ($global:LaunchResources.PowerShellInstance.InvocationStateInfo.State -eq 'Running')
				{
					$global:LaunchResources.PowerShellInstance.Stop()
				}
				$global:LaunchResources.PowerShellInstance.Dispose()
			}
			catch
			{
				Write-Verbose "LAUNCH: Failed to stop PowerShell instance: $_" -ForegroundColor Red
			}
		}
		
		# Close runspace if it exists
		if ($null -ne $global:LaunchResources.Runspace)
		{
			try
			{
				Write-Verbose 'LAUNCH: Closing runspace' -ForegroundColor DarkGray
				$global:LaunchResources.Runspace.Close()
				$global:LaunchResources.Runspace.Dispose()
			}
			catch
			{
				Write-Verbose "LAUNCH: Failed to close runspace: $_" -ForegroundColor Red
			}
		}
		
		# Stop the launch timer if it exists
		if ($global:DashboardConfig.Resources.Timers.Contains('launchTimer'))
		{
			try
			{
				Write-Verbose 'LAUNCH: Stopping launch timer' -ForegroundColor DarkGray
				$global:DashboardConfig.Resources.Timers['launchTimer'].Stop()
			}
			catch
			{
				Write-Verbose "LAUNCH: Failed to stop launch timer: $_" -ForegroundColor Red
			}
		}
		
		# Stop the safety timer if it exists
		if ($global:DashboardConfig.Resources.Timers.Contains('launchSafetyTimer'))
		{
			try
			{
				Write-Verbose 'LAUNCH: Stopping safety timer' -ForegroundColor DarkGray
				$global:DashboardConfig.Resources.Timers['launchSafetyTimer'].Stop()
				$global:DashboardConfig.Resources.Timers['launchSafetyTimer'].Dispose()
				$global:DashboardConfig.Resources.Timers.Remove('launchSafetyTimer')
			}
			catch
			{
				Write-Verbose "LAUNCH: Failed to stop safety timer: $_" -ForegroundColor Red
			}
		}
		
		# Clear the resources
		$global:LaunchResources = $null
		
		# Force garbage collection
		[System.GC]::Collect()
		
		Write-Verbose 'LAUNCH: Launch cleanup completed' -ForegroundColor Green
		# Simple property change handler that updates button appearance
		if ($global:DashboardConfig.State.LaunchActive -eq $true) {
			$global:DashboardConfig.UI.Launch.FlatStyle = 'Popup'
		} else {
			$global:DashboardConfig.UI.Launch.FlatStyle = 'Flat'
		}
	}
	catch
	{
		Write-Verbose "LAUNCH: Error during launch cleanup: $_" -ForegroundColor Red
		# Ensure the flag is reset even if cleanup fails
		$global:DashboardConfig.State.LaunchActive = $false
	}
}


#endregion Launch Management Functions

#region Module Exports

# Export module functions
Export-ModuleMember -Function Start-ClientLaunch, Stop-ClientLaunch

#endregion Module Exports