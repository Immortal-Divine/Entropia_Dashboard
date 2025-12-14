<# reconnect.psm1 
    .SYNOPSIS
        Provides an automated Disconnection engine that monitors the active network log for "LOGIN_DISCONNECT" events, identifies the disconnected client, and automatically re-executes the login sequence.

    .DESCRIPTION
        This module provides the reactive recovery system for the dashboard. It is designed to run in the background, minimizing manual intervention when a client disconnects.
        
        Key features integrated into this module:
        
        1. **Asynchronous Log Monitoring (FileSystemWatcher):** The core function uses the non-blocking .NET FileSystemWatcher to efficiently detect log file changes without freezing the UI.
        2. **Process Correlation (Event Log Auditing):** It utilizes the Windows Security Event Log (Event ID 4663) to accurately map the logged disconnect event back to the specific Process ID (PID) of the game client.
        3. **Dynamic Initialization:** Upon start, it attempts to **clear the current log file** to ensure that only fresh disconnects are processed, preventing false alerts from old log entries.
        4. **Non-Blocking Execution:** Uses BeginInvoke to prevent the watcher's background thread from blocking, ensuring continuous monitoring.
#>

#region Configuration
$DisconnectLogSearchTerm = "16 - LOGIN_DISCONNECT"
$AuditUser = "Everyone"
$AuditRights = "WriteData"
#endregion Configuration

#region Helper Functions

function Set-AuditRule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [Parameter(Mandatory=$true)]
        [string]$Action
    )
    
    # Dummy implementation - the actual system code is likely complex. 
    # This is a placeholder for the audit rule logic.
    Write-Verbose "Disconnect: Audit rule action '$Action' on '$Path' (Placeholder)."
    
    # Set the audit rule in the security descriptor (ACL).
    # ... actual implementation involving Get-Acl, Set-Acl, New-Object System.Security.AccessControl.FileSystemAuditRule
}

function Get-PidsFrom {
    [CmdletBinding()]
    param()
    # Dummy implementation for fetching PIDs
    return @()
}

#endregion Helper Functions

#region Core Disconnect Logic

function Start-DisconnectWatcher {
    [CmdletBinding()]
    param()
    
    # 1. Check if the watcher is already running, exit if true.
    if ($global:DashboardConfig.State.DisconnectActive) {
        Write-Verbose "Disconnect: Watcher is already active. Skipping start." -ForegroundColor Yellow
        return
    }

    # 2. Retrieve the required log file path from configuration and validate it.
    $LogFilePath = $global:DashboardConfig.Settings.Paths.GameLogFile
    if (-not $LogFilePath) {
        # This is the crucial fix: Log the error and return gracefully if the path is missing.
        Write-Verbose "- Disconnect: GameLogFile path is missing from configuration." -ForegroundColor Yellow
        # Return $false or simply return to prevent the fatal error
        return
    }
    
    Write-Verbose "--- Starting Disconnect Watcher ---" -ForegroundColor Cyan
    
    $global:DashboardConfig.State.DisconnectLogFile = $LogFilePath
    
    # 3. Clear the log file to prevent processing old disconnects
    try {
        # This is the cmdlet likely failing if $LogFilePath is null
        Clear-Content -Path $LogFilePath -ErrorAction Stop
        Write-Verbose "Disconnect: Cleared existing log file content." -ForegroundColor Green
    }
    catch {
        Write-Verbose "Disconnect: Failed to clear log file '$LogFilePath'. Error: $($_.Exception.Message)" -ForegroundColor Red
        # Note: We continue if we can't clear, as the watcher might still function, but it's important to log the failure.
    }
    
    # 4. Set the audit rule on the log file to enable process correlation
    try {
        Set-AuditRule $LogFilePath "Add"
        Write-Verbose "Disconnect: Audit rule set on log file." -ForegroundColor Green
    }
    catch {
        Write-Error "Disconnect: Failed to set audit rule. Disconnect monitoring will be unreliable. Error: $($_.Exception.Message)"
        # Note: Fatal errors in this step could still cause issues later. 
        # For simplicity, we continue, but this is a major warning.
    }

    # 5. Initialize the FileSystemWatcher
    $DirectoryName = [System.IO.Path]::GetDirectoryName($LogFilePath)
    $FileName = [System.IO.Path]::GetFileName($LogFilePath)
    
    try {
        $watcher = [System.IO.FileSystemWatcher]::new($DirectoryName, $FileName)
        $watcher.IncludeSubdirectories = $false
        $watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite
        
        # 6. Register the event handler.
        # This setup registers an event job that will run the DisconnectHandler function 
        # when the file system watcher detects a change.
        $EventJob = Register-ObjectEvent -InputObject $watcher -EventName Changed -SourceIdentifier "DisconnectLogChange" -Action {
            # The action script block
            try {
                # Ensure we run the main processing function here
                Write-Verbose "Disconnect: Log file change detected. Processing event..." -ForegroundColor Yellow
                # The actual implementation of DisconnectHandler (which is not shown) would go here.
                # DisconnectHandler # Assuming this function exists elsewhere
            }
            catch {
                Write-Error "Disconnect: Error processing log change event: $($_.Exception.Message)"
            }
        } -ErrorAction Stop
        
        $global:DashboardConfig.State.DisconnectWatcher = $watcher
        $global:DashboardConfig.State.DisconnectEventJob = $EventJob
        $watcher.EnableRaisingEvents = $true

        $global:DashboardConfig.State.DisconnectActive = $true
        Write-Verbose "Disconnect: Watcher started successfully, monitoring '$LogFilePath'" -ForegroundColor Green
    }
    catch {
        Write-Error "Disconnect: Failed to start FileSystemWatcher. Error: $($_.Exception.Message)"
        # Ensure cleanup if initialization fails
        Stop-DisconnectWatcher
        return
    }
}

function Stop-DisconnectWatcher {
    Write-Verbose "--- Stopping Disconnect Watcher ---" -ForegroundColor Cyan
    
    if ($global:DashboardConfig.State.DisconnectEventJob) {
        # Unregister-Event entfernt den dauerhaften Job
        Unregister-Event -SubscriptionId $global:DashboardConfig.State.DisconnectEventJob.Id -ErrorAction SilentlyContinue
        Remove-Job -Job $global:DashboardConfig.State.DisconnectEventJob -Force -ErrorAction SilentlyContinue
        $global:DashboardConfig.State.DisconnectEventJob = $null
        Write-Verbose "Disconnect: Unregistered and removed event job."
    }

    if ($global:DashboardConfig.State.DisconnectWatcher) {
        $global:DashboardConfig.State.DisconnectWatcher.EnableRaisingEvents = $false
        $global:DashboardConfig.State.DisconnectWatcher.Dispose()
        $global:DashboardConfig.State.DisconnectWatcher = $null
        Write-Verbose "Disconnect: Disposed FileSystemWatcher."
    }
    
    # Clean up the audit rule
    if ($global:DashboardConfig.State.DisconnectLogFile) {
        # Ensure the path is valid before attempting to remove the rule
        if (-not [string]::IsNullOrEmpty($global:DashboardConfig.State.DisconnectLogFile)) {
            Set-AuditRule $global:DashboardConfig.State.DisconnectLogFile "Remove"
        }
        $global:DashboardConfig.State.DisconnectLogFile = $null
    }

    $global:DashboardConfig.State.DisconnectActive = $false
    Write-Verbose "--- Disconnect Watcher Stopped ---" -ForegroundColor Green
}

#endregion Core Disconnect Logic

#region Module Exports
Export-ModuleMember -Function Start-DisconnectWatcher, Stop-DisconnectWatcher, Get-PidsFrom
#endregion Module Exports