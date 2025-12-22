<# start.ps1
	.SYNOPSIS
		Initializes and launches the Entropia Dashboard application.

	.DESCRIPTION
		This script is the main entry point for the Entropia Dashboard application. 
		It performs a sequence of critical actions to ensure the application starts correctly, remains stable, and shuts down cleanly.

		Its primary responsibilities include:

		1.  **Environment Verification and Elevation:**
			*   It first checks if it's running with Administrator privileges, in a 32-bit PowerShell process, and with the 'Bypass' execution policy.
			*   If these requirements are not met, it automatically attempts to relaunch itself with the correct permissions, prompting the user for elevation via UAC if necessary.

		2.  **Configuration and Directory Setup:**
			*   It establishes a global configuration structure (`$global:DashboardConfig`) to manage all application settings, state, paths, and resources.
			*   It ensures that the required application data directories exist in the user's `%APPDATA%\Entropia_Dashboard` folder and that it has write permissions.

		3.  **Module Management and Loading:**
			*   It defines a set of modules with specific priorities (Critical, Important, Optional), dependencies, and a defined loading order.
			*   It intelligently deploys modules from source files or embedded Base64 strings into the AppData directory. To optimize startup, it uses SHA256 hash checks to avoid rewriting files that haven't changed.
			*   It uses a robust, multi-attempt strategy to import PowerShell modules (.psm1), gracefully falling back to alternative methods (including reflection-based loading and, as a last resort, Invoke-Expression) to handle different execution contexts (script vs. compiled EXE) and potential loading issues. Only failures of 'Critical' modules will prevent the application from starting.

		4.  **UI Initialization and Lifecycle Management:**
			*   After successfully loading the critical modules, it initializes the main graphical user interface (UI) by calling functions from the 'ui.psm1' module.
			*   It starts and manages the Windows Forms message loop, which keeps the application's UI responsive to user input and other events. It prioritizes an efficient, low-CPU P/Invoke-based message loop if available.
			*   The entire execution is wrapped in comprehensive error handling. Critical failures are displayed to the user in a dialog box, and a final cleanup routine is executed upon exit to stop timers and dispose of UI resources properly.

	.NOTES
		Author: Immortal / Divine
		Version: 2.1
		Requires: PowerShell 5.1+, .NET Framework 4.5+, Administrator privileges, Bypassed 32-bit PowerShell execution.

		Documentation Standards Followed:
		- Module Level Documentation: Synopsis, Description, Notes.
		- Function Level Documentation: Synopsis, Parameter Descriptions, Output Specifications.
		- Code Organization: Logical grouping using #region Description/ #endregion Description. Functions organized by workflow.
		- Step Documentation: Code blocks enclosed in '#region Step: Description' / '#endregion Step: Description'.
		- Error Handling: Comprehensive try/catch/finally blocks with verbose logging and user notification on critical failure.

		Execution Policy Note: This script requires and attempts to set the execution policy to 'Bypass' for the *current process*.
		This is necessary for its dynamic module loading and execution features but reduces script execution security restrictions.
		Ensure you understand the implications before running this script in sensitive environments.

		Invoke-Expression Note: The fallback module import methods uses Invoke-Expression. This cmdlet
		can execute arbitrary code and poses a security risk if the module content is compromised.
#>

#region Custom Write-Verbose

[CmdletBinding()]
param()

if ($args -contains '-Verbose') {
	$VerbosePreference = "Continue"
	Write-Verbose "-Verbose argument detected, enabling verbose preference."
}

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop

function Write-Verbose {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)][string]$Message,
        [string]$ForegroundColor = 'DarkGray'
    )
    if ($VerbosePreference -eq "Continue") {
        $dateStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $callStack = Get-PSCallStack
        $caller = if ($callStack.Count -gt 1) { $callStack[1] } else { $callStack[0] }
        $callerName = if ($caller.Command) { $caller.Command } else { "Script" }
        
        $bracketedCaller = "[$callerName]"
        $paddedCaller = $bracketedCaller.PadRight(35)
        $prefix = " | $dateStr - $paddedCaller - "
        $indentation = " " * $prefix.Length + "| "
        
        $consoleWidth = [Math]::Max($Host.UI.RawUI.WindowSize.Width, 800)

        # Word Wrapping Logic
        $lines = $Message -split "`r`n"
        $formattedLines = @()
        $availableWidth = $consoleWidth - $prefix.Length - 5

        foreach ($line in $lines) {
            if ($line.Length -le $availableWidth) {
                $formattedLines += $line
                continue
            }
            $currentLine = ""
            $line.Split(' ') | ForEach-Object {
                if (($currentLine.Length + $_.Length + 1) -le $availableWidth) {
                    $currentLine += (if ($currentLine) { " $_" } else { $_ })
                } else {
                    $formattedLines += $currentLine
                    $currentLine = $_
                }
            }
            if ($currentLine) { $formattedLines += $currentLine }
        }
        
        # Line Joining Logic
        $formattedMessage = ""
		if ($formattedLines.Count -gt 0) {
			$formattedMessage = $formattedLines[0]
			for ($i = 1; $i -lt $formattedLines.Count; $i++) {
				$formattedMessage += "`r`n$indentation$($formattedLines[$i])"
			}
		}

		$logPath = $global:DashboardConfig.Paths.Verbose 
 
		$logLine = "$prefix$formattedMessage"
         
		try {
			Add-Content -Path $logPath -Value $logLine -ErrorAction SilentlyContinue
		} catch {}
        
        
        # Color Mapping Logic
        $color = switch ($ForegroundColor.ToLower()) {
            'red' {[ConsoleColor]::Red}
            'yellow' {[ConsoleColor]::Yellow}
            'green' {[ConsoleColor]::Green}
            'cyan' {[ConsoleColor]::Cyan}
            default {[ConsoleColor]::DarkGray}
        }

        # Colored Output Logic
        $orig = $host.UI.RawUI.ForegroundColor
        try {
            $host.UI.RawUI.ForegroundColor = $color
            [Console]::Error.WriteLine("$prefix$formattedMessage")
        } finally {
            $host.UI.RawUI.ForegroundColor = $orig
        }
        
        # Steppable Pipeline/Write-Verbose Logic
        $wrappedCmdlet = $ExecutionContext.InvokeCommand.GetCommand("Microsoft.PowerShell.Utility\Write-Verbose", [System.Management.Automation.CommandTypes]::Cmdlet)
        $script = { & $wrappedCmdlet "$prefix$formattedMessage" }
        $pipe = $script.GetSteppablePipeline()
        $pipe.Begin($true)
        $pipe.End()
    }
}

try {
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing
    Write-Verbose "INFO: Assemblies loaded" 
} catch {
    $msg = $_.Exception.Message
    Write-Verbose "ERROR: Failed to load assemblies: $msg" -ForegroundColor Red
    throw "Failed to initialize application. Required assemblies could not be loaded: $msg"
}

#endregion Custom Write-Verbose

#region Global Configuration

$global:DashboardConfig = [hashtable]::Synchronized(@{
    Paths = @{
        Source = Join-Path $env:USERPROFILE 'Entropia_Dashboard\.main'
        App = Join-Path $env:APPDATA 'Entropia_Dashboard\'
		Profiles = Join-Path $env:APPDATA 'Entropia_Dashboard\Profiles'
        Modules = Join-Path $env:APPDATA 'Entropia_Dashboard\modules'
        Icon = Join-Path $env:APPDATA 'Entropia_Dashboard\modules\icon.ico'
        FtoolDLL = Join-Path $env:APPDATA 'Entropia_Dashboard\modules\ftool.dll'
        Ini = Join-Path $env:APPDATA 'Entropia_Dashboard\config.ini'
        Verbose = Join-Path $env:APPDATA 'Entropia_Dashboard\log\verbose1.log'
    }
    State = @{
        ConfigInitialized = $false
        UIInitialized = $false
        LoginActive = $false
        LaunchActive = $false
        PreviousLaunchState = $false
        PreviousLoginState = $false
        IsRunningAsExe = $false
        IsDragging = $false
		DisconnectActive = $false
		LastNotifyHwnd = $false
		ReconnectActive = $false
    }
    Resources = @{
        Timers = [ordered]@{}
        FtoolForms = [ordered]@{}
        LastEventTimes = @{}
        ExtensionData = @{}
        ExtensionTracking = @{}
        LoadedModuleContent = @{}
        LaunchResources = @{}
        DragSourceGrid = $null
    }
    UI = @{
        Login = @{}
    }
	DefaultConfig = [ordered]@{
            'LauncherPath' = [ordered]@{ 'LauncherPath' = 'Select Launcher.exe' }
            'ProcessName' = [ordered]@{ 'ProcessName' = 'neuz' }
            'MaxClients' = [ordered]@{ 'MaxClients' = '1' }
            'Login' = [ordered]@{ 'Login' = '1,1,1,1,1,1,1,1,1,1'; 'FinalizeCollectorLogin' = '0'; 'NeverRestartingCollectorLogin' = '0' }
            'Ftool' = [ordered]@{}
            'Options' = [ordered]@{ 'HideMinimizedWindows' = '0' }
            'Paths' = [ordered]@{ 'JunctionTarget' = (Join-Path $env:APPDATA 'Entropia_Dashboard\Profiles') }
            'ReconnectProfiles' = @{} 
            'LoginConfig' = [ordered]@{ 'PostLoginDelay' = '1'; 'Server1Coords' = '0,0'; 'Server2Coords' = '0,0'; 'Channel1Coords' = '0,0'; 'Channel2Coords' = '0,0'; 'FirstNickCoords' = '0,0'; 'ScrollDownCoords' = '0,0'; 'Char1Coords' = '0,0'; 'Char2Coords' = '0,0'; 'Char3Coords' = '0,0'; 'CollectorStartCoords' = '0,0'; 'DisconnectOKCoords' = '0,0'; 'LoginDetailsOKCoords' = '0,0' }
        }
    Config = [ordered]@{}
    ConfigWriteTimer = @{}
    LoadedModules = @{}
})

#region Step: Define Module Metadata
$global:DashboardConfig.Modules = @{
		'ftool.dll' = @{ 
			Priority = 'Critical'; Order = 1; Dependencies = @(); 
			Base64Content = '
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S0VSTkVMMzIuZGxsAAByAV9lbmNvZGVfcG9pbnRlcgCTAl9tYWxsb2NfY3J0APQEZnJlZQAAcwFfZW5jb2RlZF9udWxsAGgBX2RlY29kZV9wb2ludGVyABACX2luaXR0ZXJtABECX2luaXR0ZXJtX2UAHQFfYW1zZ19leGl0AAATAV9hZGp1c3RfZmRpdgAAbQBfX0NwcFhjcHRGaWx0ZXIAUwFfY3J0X2RlYnVnZ2VyX2hvb2sAAI8AX19jbGVhbl90eXBlX2luZm9fbmFtZXNfaW50ZXJuYWwAAPMDX3VubG9jawCZAF9fZGxsb25leGl0AIICX2xvY2sAKANfb25leGl0AE1TVkNSODAuZGxsAHsBX2V4Y2VwdF9oYW5kbGVyNF9jb21tb24AKQJJbnRlcmxvY2tlZEV4Y2hhbmdlAFYDU2xlZXAAJgJJbnRlcmxvY2tlZENvbXBhcmVFeGNoYW5nZQAAXgNUZXJtaW5hdGVQcm9jZXNzAABCAUdldEN1cnJlbnRQcm9jZXNzAG4DVW5oYW5kbGVkRXhjZXB0aW9uRmlsdGVyAABKA1NldFVuaGFuZGxlZEV4Y2VwdGlvbkZpbHRlcgA5AklzRGVidWdnZXJQcmVzZW50AKMCUXVlcnlQZXJmb3JtYW5jZUNvdW50ZXIA3wFHZXRUaWNrQ291bnQAAEYBR2V0Q3VycmVudFRocmVhZElkAABDAUdldEN1cnJlbnRQcm9jZXNzSWQAygFHZXRTeXN0ZW1UaW1lQXNGaWxlVGltZQAAAAAAAAAAAAAAAAAAAAAAAACkGltIAAAAAFQmAAABAAAABgAAAAYAAAAYJgAAMCYAAEgmAAAAEAAAEBAAACAQAAAwEAAAUBAAAEAQAABiJgAAcSYAAIAmAACPJgAAniYAALAmAAAAAAEAAgADAAQABQBGdW5jdGlvbnMuZGxsAGZuUG9zdE1lc3NhZ2VBAGZuUG9zdE1lc3NhZ2VXAGZuU2VuZE1lc3NhZ2VBAGZuU2VuZE1lc3NhZ2VXAGZuU2V0QWN0aXZlV2luZG93AGZuU2V0Q3Vyc29yUG9zAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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
				'
		}
		'icon.ico' = @{ 
			Priority = 'Critical'; Order = 2; Dependencies = @(); 
			Base64Content = '
AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAACMuAAAjLgAAAAAAAAAAAAAAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AgIC/wICAv8CAgL/BAQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wQEAv8CAgL/AgIC/wICAv8DAwP/AwMD/wMDA/8HBgP/BwYD/wcGA/8HBgP/BwYD/wcGA/8HBgP/BwYD/wcGA/8GBgP/BQUD/wUEA/8NDgP/DQ4D/wUEA/8FBQP/BgYD/wcGA/8HBgP/BwYD/wcGA/8HBgP/BwYD/wcGA/8HBgP/BgYD/wMDA/8DAwP/AwMD/wQEBP8EBAT/BAQE/wMDBP8DAwT/AwME/wMDBP8DAwT/AwME/wQFBP8CAwT/BQQE/xIPA/8jHQP/LiUD/zo0A/86NAP/LiUD/yMdA/8SDwP/BQQE/wIDBP8EBQT/AwME/wMDBP8DAwT/AwME/wMDBP8DAwT/BAQE/wQEBP8EBAT/BAQE/wQEBP8EBAT/BAQE/wQEBP8EBAT/BAQE/wQEBP8DAwT/EBME/ycmBP84LQP/OC0E/ykhBP8cFwT/FxME/xcTBP8cFwT/KSEE/zgtBP84LQP/JyYE/xATBP8DAwT/BAQE/wQEBP8EBAT/BAQE/wQEBP8EBAT/BAQE/wQEBP8FBQX/BQUF/wUFBf8FBQX/BQUF/wUFBf8FBQX/BAQF/w0MBf83LgT/OjIE/xYSBf8JCgX/Cw0F/w4QBf8PEQX/DhEF/w4QBf8LDQX/CQoF/xYSBf86MgT/OC4E/w0MBf8EBAX/BQUF/wUFBf8FBQX/BQUF/wUFBf8FBQX/BQUF/wYGBv8GBgb/BgYG/wYGBv8GBgb/BgYG/wUFBv8SEAb/PjIF/yMdBf8JCQb/DxIG/xETBv8MDgb/CQoG/wgJBv8ICQb/CQoG/wwOBv8REwb/DxIG/wkJBv8jHQX/PjIF/xIQBv8FBQb/BgYG/wYGBv8GBgb/BgYG/wYGBv8GBgb/BwcH/wcHB/8HBwf/BwcH/wcHB/8GBgf/Dw0H/z4zBf8dGAb/CwwH/xIVBv8MDQf/BwcH/wYGB/8HBwf/BwcH/wcHB/8HBwf/BgYH/wcHB/8MDQf/EhUG/wsMB/8dGAb/PjMF/w8NB/8GBgf/BwcH/wcHB/8HBwf/BwcH/wcHB/8ICAj/CAgI/wgICP8ICAj/CQkI/xMWB/85MAb/JR4H/wsNCP8TFgf/CQoI/wgHCP8ICAj/CAgI/wgICP8ICAj/CAgI/wgICP8ICAj/CAgI/wgHCP8JCgj/ExYH/wsNCP8lHgf/OTAG/xMWB/8JCQj/CAgI/wgICP8ICAj/CAgI/wkJCf8JCQn/CQkJ/wkJCf8ICAn/KykH/z01B/8LDAn/FBcI/woKCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8KCgn/FBcI/wsMCf89NQf/KykH/wgICf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/woKCf88MQf/GRYI/xIVCf8ODwn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8ODwn/EhUJ/xkWCP88MQf/CgoJ/wkJCf8JCQn/CQkJ/wkJCf8KCgr/CgoK/woKCv8JCQr/GBUJ/zwyCP8ODgr/FRcK/woKCv8KCgr/CgoK/woKCv8KCgr/CgoK/woKCv8KCgr/CgoK/woKCv8KCgr/CgoK/woKCv8KCgr/CgoK/woKCv8UFwr/Dg4K/zwyCP8YFQn/CQkK/woKCv8KCgr/CgoK/wsLC/8LCwv/CwsL/wkKC/8qIwr/LiYJ/xATC/8REwv/CwsL/wsLC/8LCwv/CwsL/wsLC/8LCwv/CwsL/wsLC/8LCwv/CwsL/wsLC/8LCwv/CwsL/wsLC/8LCwv/CwsL/xETC/8QEgv/LiYJ/yojCv8JCgv/CwsL/wsLC/8LCwv/DAwM/wwMDP8MDAz/CgoM/zUsCv8jHgv/FBcL/w8QDP8MDAz/DAwM/wwMDP8MDAz/DAwM/wwMDP8MDAz/CwoM/wsKDP8MDAz/DAwM/wwMDP8MDAz/DAwM/wwMDP8MDAz/DxAM/xQXC/8jHgv/NSwK/woKDP8MDAz/DAwM/wwMDP8NDQ3/DQ0N/w0NDf8SFQ3/QTsK/x4bDP8WGQz/Dw8N/w0NDf8NDQ3/DQ0N/w0NDf8NDQ3/DQ0N/wwMDf9JWQn/SVkJ/wwMDf8NDQ3/DQ0N/w0NDf8NDQ3/DQ0N/w0NDf8PEA3/FxoM/x4bDP9BOwr/EhUN/w0NDf8NDQ3/DQ0N/w0NDf8NDQ3/DQ0N/xMVDf9COwr/HhsM/xYZDP8PDw3/DQ0N/w0NDf8NDQ3/DQ0N/w0NDf8NDQ3/CwoN/1ltCP9ZbQj/CwoN/w0NDf8NDQ3/DQ0N/w0NDf8NDQ3/DQ0N/w8QDf8XGgz/HhsM/0I7Cv8TFQ3/DQ0N/w0NDf8NDQ3/Dg4O/w4ODv8ODg7/DAwO/zYuDP8lIA3/FhkN/xESDv8ODg7/Dg4O/w4ODv8ODg7/Dg4O/w4ODv8MCw7/Mz0L/zM9C/8MCw7/Dg4O/w4ODv8ODg7/Dg4O/w4ODv8ODg7/ERIO/xYZDf8lIA3/Ni4M/wwMDv8ODg7/Dg4O/w4ODv8PDw//Dw8P/w8PD/8NDg//LScN/zEqDf8UFg//FRcO/w8PD/8PDw//Dw8P/w8PD/8PDw//Dw8P/w0MD/81Pwz/NT8M/w0MD/8PDw//Dw8P/w8PD/8PDw//Dw8P/w8PD/8VFw//FBYP/zEqDf8tJw3/DQ4P/w8PD/8PDw//Dw8P/xAQEP8QEBD/EBAQ/w8PEP8eGw//QTYN/xMUEP8aHQ//EBAQ/xAQEP8QEBD/EBAQ/xAQEP8QEBD/Dg0Q/zZADf82QA3/Dg0Q/xAQEP8QEBD/EBAQ/xAQEP8QEBD/EBAQ/xodD/8TFBD/QTYN/x4bD/8PDxD/EBAQ/xAQEP8QEBD/ERER/xEREf8RERH/ERER/xISEf9CNw3/IB0Q/xkcEP8VFhD/ERER/xEREf8RERH/ERER/xEREf8PDhH/NkAO/zZADv8PDhH/ERER/xEREf8RERH/ERER/xEREf8VFhD/GRwQ/yAdEP9CNw3/EhIR/xEREf8RERH/ERER/xEREf8RERH/ERER/xEREf8RERH/EBAR/zIxD/9EPA7/FBUR/xwfEP8TExH/ERER/xEREf8RERH/ERER/w8OEf82QA7/NkAO/w8OEf8RERH/ERER/xEREf8RERH/ExMR/xwfEP8UFBH/RDwO/zIxD/8QEBH/ERER/xEREf8RERH/ERER/xISEv8SEhL/EhIS/xISEv8TExL/HSAR/0I5D/8uKBD/FhgS/x0gEf8UFBL/EhIS/xISEv8SEhL/EA8S/zdBDv83QQ7/EA8S/xISEv8SEhL/EhIS/xQUEv8dIBH/FhcS/y4oEP9COQ//HSAR/xMTEv8SEhL/EhIS/xISEv8SEhL/ExMT/xMTE/8TExP/ExMT/xMTE/8SEhP/GxkS/0g8D/8oIxH/FxgT/x4hEv8XGRP/ExMT/xMTE/8REBP/OEIP/zhCD/8REBP/ExMT/xMTE/8XGRP/HiES/xcYE/8oIxH/SDwP/xsZEv8SEhP/ExMT/xMTE/8TExP/ExMT/xMTE/8UFBT/FBQU/xQUFP8UFBT/FBQU/xQUFP8TExT/IB0T/0k9D/8vKRL/FhcU/xwfE/8eIRP/GhwT/xUVFP87RRD/O0UQ/xUVFP8aHBP/HiAT/xwfE/8WFxT/LykS/0k9D/8gHRP/ExMU/xQUFP8UFBT/FBQU/xQUFP8UFBT/FBQU/xUVFf8VFRX/FRUV/xUVFf8VFRX/FRUV/xUVFf8UFBX/HBsU/0Q7EP9GPhD/JCAT/xgZFP8aHBT/Gx0U/zpFEP87RhD/Gx0U/xocFP8YGRT/JCAT/0Y+EP9EOxD/HBsU/xQUFf8VFRX/FRUV/xUVFf8VFRX/FRUV/xUVFf8VFRX/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8VFBb/ICMV/zY0Ev9FOxH/RTsR/zcvEv8rJhT/KSYU/ykmFP8rJhT/Ny8S/0U7Ef9FOxH/NjQS/yAjFP8VFBb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8XFxb/FRYW/xcXFv8kIRX/NC0U/z01E/9JQhH/SUIR/z41E/80LRT/JCEV/xcXFv8VFhb/FxcW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xYXF/8VFhf/FRUX/xwfFv8cHxb/FRUX/xUWF/8WFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=
				'
		}
		'classes.psm1' = @{ 
			Priority = 'Critical'; Order = 3; Dependencies = @(); 
			FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'classes.psm1'); 
			Base64Content = '

			'
		}
		'ini.psm1' = @{ 
			Priority = 'Critical'; Order = 4; Dependencies = @('classes.psm1'); 
			FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'ini.psm1'); 
			Base64Content = '

			'
		}
		'ui.psm1' = @{ 
			Priority = 'Critical'; Order = 5; Dependencies = @('classes.psm1', 'ini.psm1'); 
			FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'ui.psm1'); 
			Base64Content = '

			'
		}
		'datagrid.psm1' = @{ 
			Priority = 'Important'; Order = 6; Dependencies = @('classes.psm1', 'ui.psm1', 'ini.psm1'); 
			FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'datagrid.psm1'); 
			Base64Content = '

			'
		}
		'launch.psm1' = @{ 
			Priority = 'Optional'; Order = 7; Dependencies = @('classes.psm1', 'ui.psm1', 'ini.psm1', 'datagrid.psm1'); 
			FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'launch.psm1'); 
			Base64Content = '

			'
		}
		'login.psm1' = @{ 
			Priority = 'Optional'; Order = 8; Dependencies = @('classes.psm1', 'ui.psm1', 'ini.psm1', 'datagrid.psm1'); 
			FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'login.psm1'); 
			Base64Content = '

			'
		}
		'ftool.psm1' = @{ 
			Priority = 'Optional'; Order = 9; Dependencies = @('classes.psm1', 'ui.psm1', 'ini.psm1', 'datagrid.psm1'); 
			FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'ftool.psm1'); 
			Base64Content = '

			'
		}
		'reconnect.psm1' = @{ 
			Priority = 'Optional'; Order = 10; Dependencies = @('ftool.dll', 'datagrid.psm1', 'ini.psm1', 'ui.psm1'); 
			FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'reconnect.psm1'); 
			Base64Content = '

			'
		}
}
#endregion Step: Define Module Metadata

#endregion Global Configuration

#region Environment Initialization and Checks

#region Function: Show-ErrorDialog
function Show-ErrorDialog {
    param([Parameter(Mandatory=$true)][string]$Message)
    try {
        if (-not ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))) {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        }
        [System.Windows.Forms.MessageBox]::Show($Message, 'Entropia Dashboard Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } catch {
        Write-Verbose "Failed to display error dialog: `"$Message`". Dialog Display Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}
#endregion Function: Show-ErrorDialog

#region Function: Request-Elevation
function Request-Elevation {
    param()
    Write-Verbose "Checking environment (Admin, 32-bit, Policy)..." -ForegroundColor Cyan
    [bool]$needsRestart = $false
    [System.Collections.ArrayList]$reason = @()

    [bool]$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        $needsRestart = $true; $null = $reason.Add('Administrator privileges required.')
    }

    [bool]$is32Bit = [IntPtr]::Size -eq 4
    if (-not $is32Bit) {
        $needsRestart = $true; $null = $reason.Add('32-bit execution required.')
    }

    [string]$currentPolicy = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue
    if ($currentPolicy -ne 'Bypass') {
        $needsRestart = $true
        $effectivePolicy = if ($currentPolicy -ne '') { $currentPolicy } else { Get-ExecutionPolicy }
        $null = $reason.Add("Execution Policy 'Bypass' required (Current effective: '$effectivePolicy').")
    }

    if ($needsRestart) {
        Write-Verbose "Restarting script needed: $($reason -join ' ')" -ForegroundColor Yellow
        [string]$psExe = Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
        if (-not (Test-Path $psExe -PathType Leaf)) {
            Show-ErrorDialog "FATAL: Required 32-bit PowerShell executable not found at '$psExe'. Cannot continue."
            exit 1
        }
        
        $encodedCommand = @"

"@
        try {
            $decodedCommand = [System.Text.Encoding]::UTF8.GetString(([System.Convert]::FromBase64String($encodedCommand)))
        } catch {
            Show-ErrorDialog "FATAL: Failed to decode the embedded command. Error: $($_.Exception.Message)"
            exit 1
        }

        $tempScriptPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ([System.Guid]::NewGuid().ToString() + ".ps1"))

        try {
            [System.IO.File]::WriteAllText($tempScriptPath, $decodedCommand, [System.Text.Encoding]::UTF8)
            [string]$psArgs = "-WindowStyle Hidden -noexit -ExecutionPolicy Bypass -File `"$tempScriptPath`""
            
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $psExe
            $psi.Arguments = $psArgs
            $psi.UseShellExecute = $true
            $psi.Verb = 'RunAs'
            
            Write-Verbose "Attempting restart: `"$psExe`" $psArgs (Temp: $tempScriptPath)" -ForegroundColor Cyan
            [System.Diagnostics.Process]::Start($psi) | Out-Null
            Write-Verbose "Success. Exiting current process." -ForegroundColor Green
            exit 0
        } catch {
            Show-ErrorDialog "FATAL: Failed to restart script (Admin/32-bit/Bypass). Error: $($_.Exception.Message)"
            if (Test-Path $tempScriptPath) {
                try { Remove-Item $tempScriptPath -ErrorAction SilentlyContinue } catch {}
            }
            exit 1
        }
    } else {
        Write-Verbose "Script already running with required environment settings." -ForegroundColor Green
    }
}
#endregion Function: Request-Elevation

#region Function: Initialize-ScriptEnvironment
	function Initialize-ScriptEnvironment
	{
		<#
		.SYNOPSIS
			Verifies that the script environment meets all requirements *after* any potential restart attempt by Request-Elevation.
		
		.DESCRIPTION
			This function performs final checks to ensure the script is operating in the correct environment before proceeding with core logic.
			It re-validates:
			1. Administrator Privileges: Confirms the script is now running elevated.
			2. 32-bit Mode: Confirms the script is now running in a 32-bit PowerShell process.
			3. Execution Policy: Confirms the process scope execution policy is 'Bypass'. If not (which shouldn't happen if Request-Elevation worked),
			it makes a final attempt to set it using Set-ExecutionPolicy.
			
			If any check fails, it displays a specific error message using Show-ErrorDialog and returns $false.
		
		.OUTPUTS
			[bool] Returns $true if all environment checks pass successfully, otherwise returns $false.
		
		.NOTES
			- This function should be called *after* Request-Elevation. It acts as a final safeguard.
			- Failure here is typically fatal for the application, as indicated by the error messages and the return value.
			- The attempt to set ExecutionPolicy within this function is a fallback; ideally, Request-Elevation should have ensured this.
		#>
		[CmdletBinding()]
		[OutputType([bool])] 
		param()
		
		Write-Verbose "Verifying final script environment settings..." -ForegroundColor Cyan
		try
		{
			#region Step: Verify Administrator Privileges
				# $isAdmin - Flag ($true/$false), $true if the current user is an Admin.
				[bool]$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
				if (-not $isAdmin)
				{
					# Show error and return $false if not running as Admin.
					Show-ErrorDialog 'FATAL: Application requires administrator privileges to run.'
					return $false
				}
				Write-Verbose "[OK] Running with administrator privileges." -ForegroundColor Green
			#endregion Step: Verify Administrator Privileges
			
			#region Step: Verify 32-bit Execution Mode
				# $is32Bit - Flag ($true/$false), $true if the process is 32-bit.
				[bool]$is32Bit = [IntPtr]::Size -eq 4
				if (-not $is32Bit)
				{
					# Show error and return $false if not running in 32-bit mode.
					Show-ErrorDialog 'FATAL: Application must run in 32-bit PowerShell mode.'
					return $false
				}
				Write-Verbose "[OK] Running in 32-bit mode." -ForegroundColor Green
			#endregion Step: Verify 32-bit Execution Mode
			
			#region Step: Verify Process Execution Policy
				# $currentPolicy - Text, the execution policy for this process.
				[string]$currentPolicy = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue
				if ($currentPolicy -ne 'Bypass')
				{
					# This is a backup. Ideally, Request-Elevation already set 'Bypass'.
					Write-Verbose "  Process Execution Policy is not 'Bypass' (Current: '$(if (-not [string]::IsNullOrEmpty($currentPolicy)) { $currentPolicy } else { Get-ExecutionPolicy })'). Attempting final Set..." -ForegroundColor Yellow
					try
					{
						# Try to force the policy to Bypass for this process.
						Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
						# Check again after trying.
						$currentPolicy = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue
						if ($currentPolicy -ne 'Bypass')
						{
							# If it still didn't work, report a major error.
							Show-ErrorDialog "FATAL: Failed to set required PowerShell Execution Policy to 'Bypass'.  (Current: '$(if (-not [string]::IsNullOrEmpty($currentPolicy)) { $currentPolicy } else { Get-ExecutionPolicy })')."
							return $false
						}
						Write-Verbose "[OK] Execution policy successfully forced to Bypass for this process." -ForegroundColor Green
					}
					catch
					{
						# Catch errors during the last Set-ExecutionPolicy try.
						Show-ErrorDialog "FATAL: Error setting PowerShell Execution Policy to 'Bypass'.  (Current: '$(if (-not [string]::IsNullOrEmpty($currentPolicy)) { $currentPolicy } else { Get-ExecutionPolicy })'). Error: $($_.Exception.Message)"
						return $false
					}
				}
				else
				{
					Write-Verbose "[OK] Execution policy is '$currentPolicy'." -ForegroundColor Green
				}
			#endregion Step: Verify Process Execution Policy
			
			# If all checks passed:
			Write-Verbose "  Environment verification successful." -ForegroundColor Green
			return $true
		}
		catch
		{
			# Catch any surprise errors during the check itself.
			Show-ErrorDialog "FATAL: An unexpected error occurred during environment verification: $($_.Exception.Message)"
			return $false
		}
	}
#endregion Function: Initialize-ScriptEnvironment

#region Function: Initialize-BaseConfig
	function Initialize-BaseConfig
	{
		<#
		.SYNOPSIS
			Ensures necessary application data directories exist in %APPDATA%, are writable, and clears log files.
		
		.DESCRIPTION
			This function is responsible for setting up the fundamental directory structure required by the application
			within the user's roaming application data folder (%APPDATA%). It specifically checks for and, if necessary, creates:
			1. The main application directory: %APPDATA%\Entropia_Dashboard
			2. The modules subdirectory: %APPDATA%\Entropia_Dashboard\modules
			3. The log subdirectory: %APPDATA%\Entropia_Dashboard\log
			
			After ensuring the directories exist, it performs a quick write test in each directory to verify permissions.
			Finally, it clears the contents of the main log files (`verbose1.log` and `verbose2.log`) to ensure a clean session.
		
		.OUTPUTS
			[bool] Returns $true if all directories exist (or were created successfully), are verified as writable, and log clearing is attempted.
			Returns $false if directory creation fails or if a directory is found to be non-writable.
		
		.NOTES
			- Upon successful completion (returning $true), it sets the global state flag '$global:DashboardConfig.State.ConfigInitialized' to $true.
			- Errors during directory creation or the write test are logged to the error stream and presented to the user via Show-ErrorDialog,
			as these are typically fatal issues preventing the application from functioning correctly.
			- Log file clearing errors are reported as non-critical warnings.
		#>
		[CmdletBinding()]
		[OutputType([bool])]
		param()

		Write-Verbose "Initializing base configuration directories in %APPDATA%..." -ForegroundColor Cyan
		try
		{
			# Define log directory path from the configured log file path.
			# $logDir - Text, path to the log directory.
			[string]$logDir = Split-Path -Path $global:DashboardConfig.Paths.Verbose -Parent

			# List of essential folders that must exist and be writable.
			# $directories - List of text paths for the required folders.
			[string[]]$directories = @(
				$global:DashboardConfig.Paths.App,     # e.g., C:\Users\User\AppData\Roaming\Entropia_Dashboard\
				$global:DashboardConfig.Paths.Modules, # e.g., C:\Users\User\AppData\Roaming\Entropia_Dashboard\modules\
				$logDir                                # e.g., C:\Users\User\AppData\Roaming\Entropia_Dashboard\log\
			)
			
			# Go through each needed folder path.
			foreach ($dir in $directories)
			{
				#region Step: Ensure Directory Exists
					# Check if the path exists and is actually a folder (Container).
					if (-not (Test-Path -Path $dir -PathType Container))
					{
						Write-Verbose "  Directory not found. Creating: '$dir'" -ForegroundColor DarkGray
						try
						{
							# Create the folder. -Force makes parent folders too. -ErrorAction Stop stops if it fails.
							$null = New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop
						}
						catch
						{
							# Handle errors when creating the folder (like permissions, bad path).
							$errorMsg = "  Failed to create required directory '$dir'. Please check permissions or path validity. Error: $($_.Exception.Message)"
							Write-Verbose $errorMsg -ForegroundColor Red
							Show-ErrorDialog $errorMsg
							return $false # Can't continue if creating the folder fails.
						}
					}
					else
					{
						Write-Verbose "  Directory exists: '$dir'" -ForegroundColor DarkGray
					}
				#endregion Step: Ensure Directory Exists
				
				#region Step: Test Directory Writability
					# Make a temporary file path in the current folder for a write test.
					# $testFile - Text, path for the temporary test file.
					[string]$testFile = Join-Path -Path $dir -ChildPath 'write_test.tmp'
					try
					{
						# Try writing a small bit of text to the test file.
						[System.IO.File]::WriteAllText($testFile, 'TestWriteAccess')
						# If writing works, delete the test file right away. -Force skips asking.
						Remove-Item -Path $testFile -Force -ErrorAction Stop
						Write-Verbose "  Directory is writable: '$dir'" -ForegroundColor DarkGray
					}
					catch
					{
						# Handle errors during writing or deleting (probably bad permissions).
						$errorMsg = "  Cannot write to directory '$dir'. Please check permissions. Error: $($_.Exception.Message)"
						Write-Verbose $errorMsg -ForegroundColor Red
						Show-ErrorDialog $errorMsg
						# Try cleaning up the test file just in case it was made but couldn't be deleted.
						if (Test-Path -Path $testFile -PathType Leaf)
						{
							Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
						}
						return $false # Can't continue if the folder isn't writable.
					}
				#endregion Step: Test Directory Writability
			} # End of the loop for each directory.
			
			#region Step: Clear Log Files
				# This step was moved to the beginning of the main execution block to prevent clearing logs after they have been written to.
			#endregion Step: Clear Log Files
			
			# If the loop finishes without returning false, all folders are ready.
			Write-Verbose "  Base configuration directories initialized and verified successfully." -ForegroundColor Green
			# Set the main state flag.
			$global:DashboardConfig.State.ConfigInitialized = $true
			return $true
		}
		catch
		{
			# Catch any surprise errors during the whole setup process.
			$errorMsg = "  An unexpected error occurred during base configuration directory initialization: $($_.Exception.Message)"
			Write-Verbose $errorMsg -ForegroundColor Red
			Show-ErrorDialog $errorMsg
			return $false
		}
	}
#endregion Function: Initialize-BaseConfig

#endregion Environment Initialization and Checks

#region Module Handling Functions

#region Function: Write-Module
	function Write-Module
	{
		<#
		.SYNOPSIS
			Writes module content (from a source file or Base64 string) to the designated modules directory in %APPDATA%, performing hash checks to avoid redundant writes.
		
		.DESCRIPTION
			This function handles the deployment of module files (e.g., .psm1, .dll, .ico) from their source location or embedded Base64 representation
			to the application's 'modules' directory under %APPDATA% (defined in $global:DashboardConfig.Paths.Modules).
			
			Key operations:
			1. Ensures the target 'modules' directory exists, attempting to create it if necessary.
			2. Retrieves the module content as a byte array, either by reading the source file specified by the -Content parameter or by decoding the Base64 string provided via -ContentBase64.
			3. If the target file already exists in the 'modules' directory:
			a. Compares the file size of the existing file with the size of the new content. If different, an update is needed.
			b. If sizes match, calculates the SHA256 hash of both the existing file and the new content in memory.
			c. If the hashes match, the function logs that no update is needed and returns the path to the existing file, avoiding an unnecessary write operation.
			d. If hashes differ, an update is needed.
			4. If the target file does not exist or an update is required (sizes/hashes differ), the function attempts to write the new content (byte array) to the target path.
			5. Includes a simple retry mechanism (up to 5 seconds) with short delays (100ms) specifically for System.IO.IOException errors during the write attempt, which often indicate temporary file locks.
		
		.PARAMETER ModuleName
			[string] The destination filename for the module in the target directory (e.g., 'ui.psm1', 'ftool.dll', 'icon.ico'). (Mandatory)
		
		.PARAMETER Content
			[string] Used in the 'FilePath' parameter set. The full path to the source file containing the module content to be copied. (Mandatory, ParameterSetName='FilePath')
		
		.PARAMETER ContentBase64
			[string] Used in the 'Base64Content' parameter set. A Base64 encoded string containing the module content to be decoded and written. (Mandatory, ParameterSetName='Base64Content')
		
		.OUTPUTS
			[string] Returns the full path to the successfully written (or verified existing and matching) module file in the target 'modules' directory.
			Returns $null if any critical operation fails (e.g., directory creation, source file reading, Base64 decoding, final write attempt after retries).
		
		.NOTES
			- Uses SHA256 hash comparison for efficient and reliable detection of unchanged files.
			- Error handling is implemented for directory creation, file reading, Base64 decoding, hash calculation, and file writing.
			- The write retry loop is basic and may not handle all concurrent access scenarios perfectly but addresses common temporary locks.
			- Uses [System.IO.File]::ReadAllBytes and ::WriteAllBytes for potentially better performance with binary files (.dll, .ico) compared to Get-Content/Set-Content.
		#>
		[CmdletBinding(DefaultParameterSetName = 'FilePath')] # Default to FilePath if only unnamed inputs are used.
		[OutputType([string])]
		param (
			[Parameter(Mandatory = $true, Position = 0)]
			[string]$ModuleName, # e.g., 'ui.psm1'
		
			[Parameter(Mandatory = $true, ParameterSetName = 'FilePath', Position = 1)]
			[ValidateScript({ Test-Path $_ -PathType Leaf })] # Basic check: make sure path exists and is a file.
			[string]$Content, # Source file path, e.g., 'C:\path\to\source\ui.psm1'
		
			[Parameter(Mandatory = $true, ParameterSetName = 'Base64Content')]
			[string]$ContentBase64 # Base64 encoded content text
		)
		
		# Get the destination folder path from the main config.
		# $modulesDir - Text, destination folder for modules.
		[string]$modulesDir = $global:DashboardConfig.Paths.Modules
		# Build the full path for the destination file.
		# $finalPath - Text, full destination path for the module file.
		[string]$finalPath = Join-Path -Path $modulesDir -ChildPath $ModuleName
		
		Write-Verbose "Executing Write-Module for '$ModuleName' to '$finalPath'" -ForegroundColor Cyan
		try
		{
			#region Step: Ensure Target Directory Exists
				# Check if the destination folder exists; try creating it if not.
				if (-not (Test-Path -Path $modulesDir -PathType Container))
				{
					Write-Verbose "Target module directory not found, attempting creation: '$modulesDir'" -ForegroundColor DarkGray
					try
					{
						$null = New-Item -Path $modulesDir -ItemType Directory -Force -ErrorAction Stop
						Write-Verbose "Target module directory created successfully: '$modulesDir'" -ForegroundColor Green
					}
					catch
					{
						# Major error if folder cannot be created.
						Write-Verbose "Failed to create target module directory '$modulesDir': $($_.Exception.Message)" -ForegroundColor Red
						return $null # Cannot continue.
					}
				}
			#endregion Step: Ensure Target Directory Exists
			
			#region Step: Get Content Bytes from Source (File or Base64)
				# $bytes - Array of bytes that will hold the module content.
				[byte[]]$bytes = $null
				Write-Verbose "  ParameterSetName: $($PSCmdlet.ParameterSetName)" -ForegroundColor DarkGray
				
				# Handle Base64 input
				if ($PSCmdlet.ParameterSetName -eq 'Base64Content')
				{
					if ([string]::IsNullOrEmpty($ContentBase64))
					{
						Write-Verbose "  ModuleName '$ModuleName': ContentBase64 parameter was provided but is empty." -ForegroundColor Yellow
						return $null
					}
					try
					{
						$bytes = [System.Convert]::FromBase64String($ContentBase64)
						Write-Verbose "  Decoded Base64 content for '$ModuleName' ($($bytes.Length) bytes)." -ForegroundColor DarkGray
					}
					catch
					{
						# Major error if Base64 decoding fails.
						Write-Verbose "  Failed to decode Base64 content for '$ModuleName': $($_.Exception.Message)" -ForegroundColor Red
						return $null
					}
				}
				# Handle FilePath input
				elseif ($PSCmdlet.ParameterSetName -eq 'FilePath')
				{
					# File existence already checked by ValidateScript, but double-check path is valid.
					if ([string]::IsNullOrEmpty($Content) -or -not ([System.IO.File]::Exists($Content)) )
					{
						Write-Verbose "  ModuleName '$ModuleName': Source file path '$Content' is invalid or does not exist." -ForegroundColor Red
						return $null # Shouldn't happen with ValidateScript, but good safety check.
					}
					try
					{
						$bytes = [System.IO.File]::ReadAllBytes($Content)
						Write-Verbose "  Read source file content for '$ModuleName' from '$Content' ($($bytes.Length) bytes)." -ForegroundColor DarkGray
					}
					catch
					{
						# Major error if source file cannot be read.
						Write-Verbose "  Failed to read source file '$Content' for '$ModuleName': $($_.Exception.Message)" -ForegroundColor Red
						return $null
					}
				}
				else # Shouldn't get here because of parameter sets
				{
					Write-Verbose "  ModuleName '$ModuleName': Invalid parameter combination or missing content." -ForegroundColor Red
					return $null
				}
				
				# Final check if the byte array got filled.
				if ($null -eq $bytes)
				{
					Write-Verbose "  Failed to obtain content bytes for '$ModuleName'. Source data might be empty or invalid." -ForegroundColor Red
					return $null
				}
			#endregion Step: Get Content Bytes from Source (File or Base64)
			
			#region Step: Check if File Needs Updating (Size and Hash Comparison)
				# $updateNeeded - Flag ($true/$false), decides if the file needs writing.
				[bool]$updateNeeded = $true
				if (Test-Path -Path $finalPath -PathType Leaf) # Check if the destination file exists.
				{
					Write-Verbose "  Target file exists: '$finalPath'. Comparing size and hash..." -ForegroundColor DarkGray
					try
					{
						# Get info about the existing file.
						# $fileInfo - File info object for the existing file.
						$fileInfo = Get-Item -LiteralPath $finalPath -Force -ErrorAction Stop
						
						# 1. Compare file sizes first (quick check).
						if ($fileInfo.Length -eq $bytes.Length)
						{
							Write-Verbose "  File sizes match ($($bytes.Length) bytes). Comparing SHA256 hashes..." -ForegroundColor DarkGray
							# 2. If sizes match, compare SHA256 hashes.
							# $existingHash - Text, SHA256 hash of the file on disk.
							[string]$existingHash = (Get-FileHash -LiteralPath $finalPath -Algorithm SHA256 -ErrorAction Stop).Hash
							
							# Calculate hash of the new content (bytes) in memory.
							# $memStream - Memory stream to feed bytes to Get-FileHash. 'Using' cleans it up.
							$newHash = try
							{
								$memStream = New-Object System.IO.MemoryStream(,$bytes)
								(Get-FileHash -InputStream $memStream -Algorithm SHA256 -ErrorAction Stop).Hash
							}
							finally
							{
								if ($memStream)
								{
									$memStream.Dispose() 
								}
							}
							
							Write-Verbose " - Existing Hash: $existingHash" -ForegroundColor DarkGray
							Write-Verbose " - New Hash:    - $newHash" -ForegroundColor DarkGray
							
							# If hashes match, no update needed.
							if ($existingHash -eq $newHash)
							{
								Write-Verbose "  Hashes match for '$ModuleName'. No update needed." -ForegroundColor DarkGray
								$updateNeeded = $false
								# Return path to the existing, checked file.
								return $finalPath
							}
							else
							{
								Write-Verbose "  Hashes differ for '$ModuleName'. Update required." -ForegroundColor Yellow 
							}
						}
						else
						{
							Write-Verbose "  File sizes differ (Existing: $($fileInfo.Length), New: $($bytes.Length)). Update required." -ForegroundColor Yellow 
						}
					}
					catch
					{
						# Handle errors during size/hash compare (like file locked, permissions).
						# Log a warning and assume an update is needed.
						Write-Verbose "  Could not compare size/hash for '$ModuleName' (Path: '$finalPath'). Will attempt to overwrite. Error: $($_.Exception.Message)" -ForegroundColor Yellow
						$updateNeeded = $true
					}
				}
				else
				{
					Write-Verbose "  Target file does not exist: '$finalPath'. Writing new file." -ForegroundColor DarkGray 
					$updateNeeded = $true
				}
			#endregion Step: Check if File Needs Updating (Size and Hash Comparison)
			
			#region Step: Write File to Target Path (with Retry on IO Exception)
				if ($updateNeeded)
				{
					# Set up retry settings.
					# $timeoutMilliseconds - Number, max time (ms) to spend retrying the write.
					[int]$timeoutMilliseconds = 5000  # 5 seconds total retry time.
					# $retryDelayMilliseconds - Number, delay (ms) between retries.
					[int]$retryDelayMilliseconds = 100 # Wait 100ms before trying again.
					# $startTime - DateTime, when the retry loop started.
					[datetime]$startTime = Get-Date
					# $fileWritten - Flag ($true/$false) if file was written okay within the time limit.
					[bool]$fileWritten = $false
					# $attempts - Number, counts how many times we tried writing.
					[int]$attempts = 0
					
					Write-Verbose "  Attempting to write file: '$finalPath'" -ForegroundColor DarkGray
					while (((Get-Date) - $startTime).TotalMilliseconds -lt $timeoutMilliseconds)
					{
						$attempts++
						try
						{
							# Try writing all bytes to the final path using a .NET method.
							[System.IO.File]::WriteAllBytes($finalPath, $bytes)
							$fileWritten = $true
							Write-Verbose "  Successfully wrote '$ModuleName' to '$finalPath' on attempt $attempts." -ForegroundColor Green
							break # Exit the retry loop if write worked.
						}
						catch [System.IO.IOException]
						{
							# Catch IO errors specifically (probably file lock). Log warning and retry after delay.
							Write-Verbose "  Attempt $($attempts): IO Error writing '$finalPath' (Retrying in $retryDelayMilliseconds ms): $($_.Exception.Message)" -ForegroundColor Red
							# Check if time is almost up before waiting.
							if (((Get-Date) - $startTime).TotalMilliseconds + $retryDelayMilliseconds -ge $timeoutMilliseconds)
							{
								Write-Verbose "  Timeout nearing, breaking retry loop for '$finalPath'." -ForegroundColor Yellow
								break # Don't wait longer than the timeout.
							}
							Start-Sleep -Milliseconds $retryDelayMilliseconds
						}
						catch
						{
							# Catch other surprise, non-retryable errors during write. Log error and stop loop.
							Write-Verbose "  Attempt $($attempts): Non-IO Error writing '$finalPath': $($_.Exception.Message)" -ForegroundColor Red
							$fileWritten = $false # Make sure flag is false.
							break # Exit loop on non-retryable error.
						}
					} # End of while retry loop
					
					# Check if the file was written okay after the loop.
					if (-not $fileWritten)
					{
						Write-Verbose "  Failed to write module '$ModuleName' to '$finalPath' after $attempts attempts within $timeoutMilliseconds ms timeout." -ForegroundColor Red
						return $null # Return null to show it failed.
					}
				} # End if($updateNeeded)
			#endregion Step: Write File to Target Path (with Retry on IO Exception)
			
			# If we get here, the file exists and is current, or it was just written successfully.
			return $finalPath
		}
		catch
		{
			# Catch any surprise errors in the main function part (like input check failed earlier).
			Write-Verbose "  An unexpected error occurred in Write-Module for '$ModuleName': $($_.Exception.Message)" -ForegroundColor Red
			return $null
		}
	}
#endregion Function: Write-Module


#region Function: Import-ModuleUsingReflection
	# ... (Keep Import-ModuleUsingReflection function as it was) ...
	function Import-ModuleUsingReflection
    {
        [CmdletBinding()]
        [OutputType([bool])]
        param(
            [Parameter(Mandatory = $true)]
            [ValidateScript({ Test-Path $_ -PathType Leaf })]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [string]$ModuleName
        )

        Write-Verbose "Attempting reflection-style import (Global Scope Injection) for '$ModuleName'." -ForegroundColor Cyan

        try
        {
            if (-not (Test-Path -Path $Path -PathType Leaf)) { return $false }

            # 1. Read Content
            [string]$moduleContent = [System.IO.File]::ReadAllText($Path)
            if ([string]::IsNullOrWhiteSpace($moduleContent)) { return $true }
            
            # 2. Store for debugging/resources
            if ($global:DashboardConfig) {
                if (-not $global:DashboardConfig.Resources) { $global:DashboardConfig.Resources = @{} }
                if (-not $global:DashboardConfig.Resources.LoadedModuleContent) { $global:DashboardConfig.Resources['LoadedModuleContent'] = @{} }
                $global:DashboardConfig.Resources.LoadedModuleContent[$ModuleName] = $moduleContent
            }

            # 3. GLOBAL SCOPE INJECTION (The Critical Fix)
            # We use Regex to find "function Name" and replace it with "function Global:Name"
            # This ensures the function survives when Import-DashboardModules returns.
            # Pattern explanation: 
            # (?m) = Multi-line mode
            # ^\s* = Start of line, optional whitespace
            # function\s+ = The word 'function' followed by space
            # ([\w-]+) = Capture group 1: The function name (letters, numbers, hyphens)
            $globalizedContent = $moduleContent -replace '(?m)^\s*function\s+([\w-]+)', 'function Global:$1'

            # 4. Prepend the dummy Export-ModuleMember just in case
            $noOpExportFunc = "function Global:Export-ModuleMember { param([string]`$Function, [string]`$Variable, [string]`$Alias, [string]`$Cmdlet) }"
            $finalContent = "$noOpExportFunc`n`n$globalizedContent"

            # 5. Execute via Dot-Sourcing
            # Even though we are dot-sourcing inside a function, the 'Global:' prefix we added
            # to the text will ensure the functions land in the Global scope.
            Write-Verbose "Executing globalized content for '$ModuleName'..." -ForegroundColor DarkGray
            
            $scriptBlock = [ScriptBlock]::Create($finalContent)
            . $scriptBlock

            Write-Verbose "Successfully executed content for '$ModuleName'." -ForegroundColor Green
            return $true
        }
        catch
        {
            Write-Verbose "FATAL error during global import for '$ModuleName': $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
#endregion Function: Import-ModuleUsingReflection


#region Function: Import-DashboardModules
	function Import-DashboardModules
	{
		<#
		.SYNOPSIS
			Loads all defined dashboard modules according to priority, dependencies, and execution context (Script vs EXE).
		
		.DESCRIPTION
			This crucial function orchestrates the loading of all modules specified in '$global:DashboardConfig.Modules'.
			It performs the following steps:
			1. Initializes tracking variables for loaded and failed modules.
			2. Determines if the script is running as a compiled EXE or a standard .ps1 script, storing the result in '$global:DashboardConfig.State.IsRunningAsExe'. This influences the import strategy.
			3. Sorts the modules based on the 'Order' property defined in their metadata to ensure correct loading sequence.
			4. Iterates through the sorted modules:
			a. Checks if all dependencies listed for the current module are already present in '$global:DashboardConfig.LoadedModules'. If not, skips the module and records the failure. Critical module dependency failures trigger a critical failure flag.
			b. Calls 'Write-Module' to ensure the module file (or resource like .dll, .ico) exists in the %APPDATA%\modules directory, handling source file paths or Base64 content, and using hash checks for efficiency. If Write-Module fails, records the failure. Critical module write failures trigger the critical failure flag.
			c. If Write-Module succeeds, adds the module name and its written path to '$global:DashboardConfig.LoadedModules'. This satisfies dependency checks for subsequent modules, including non-PSM1 files like DLLs or icons.
			d. If the module is a PowerShell module (.psm1):
			i. Attempts multiple import strategies in sequence until one succeeds:
			- Attempt 1 (Preferred): Standard `Import-Module`. If running as EXE, it first modifies the content in memory to prepend a no-op `Export-ModuleMember`, writes this to a temporary file, imports the temp file, and then deletes it. If running as a script, it imports the written module path directly. Success is verified by checking `Get-Module`.
			- Attempt 2 (Alternative): Calls `Import-ModuleUsingReflection` function (InvokeCommand in global scope). **Crucially, after this attempt returns true, this function now performs an additional verification step using `Get-Command` for key functions expected from the module.** If key functions are missing, Attempt 2 is marked as failed, and the process proceeds to Attempt 3.
			- Attempt 3 (Last Resort): Uses `Invoke-Expression` on the module content after attempting to remove/comment out `Export-ModuleMember` calls using string replacement. This attempt includes its own verification and global re-definition of functions. **(Security Risk)**
			ii. If all import attempts fail for a .psm1 module, records the failure, removes the module from '$global:DashboardConfig.LoadedModules' (as it was written but not imported), and triggers the critical failure flag if the module was critical.
			5. After processing all modules, checks the critical failure flag. If set, returns a status object indicating failure.
			6. Logs warnings for any 'Important' modules that failed and informational messages for 'Optional' module failures.
			7. If no critical failures occurred, returns a status object indicating overall success (though non-critical modules may have failed).
		
		.OUTPUTS
			[PSCustomObject] Returns an object with the following properties:
			- Status [bool]: $true if all 'Critical' modules were successfully written and (if applicable) imported without fatal errors. $false if any 'Critical' module failed or if an unhandled exception occurred.
			- LoadedModules [hashtable]: A hashtable containing {ModuleName = Path} entries for all modules that were successfully written to the AppData directory by Write-Module (includes .psm1, .dll, .ico, etc.). Note that for .psm1, inclusion here doesn't guarantee successful *import*, only successful writing/verification. Check FailedModules for import status.
			- FailedModules [hashtable]: A hashtable containing {ModuleName = ErrorMessage} entries for modules that failed during dependency check, writing (Write-Module), or importing (for .psm1 files).
			- CriticalFailure [bool]: $true if a module marked with Priority='Critical' failed at any stage (dependency, write, or import). $false otherwise.
			- Exception [string]: (Optional) Included only if an unexpected, unhandled exception occurred within the Import-DashboardModules function itself. Contains the exception message.
		
		.NOTES
			- The multi-attempt import strategy for .psm1 files adds complexity but aims for robustness, especially in potentially problematic EXE execution environments.
			- Attempt 2 now includes verification. If it passes, Attempt 3 (Invoke-Expression) is skipped.
			- The use of `Invoke-Expression` (Attempt 3) remains a significant security risk and should ideally be avoided by refactoring modules to work with Attempt 1 or a reliable Attempt 2.
			- Dependency checking relies on modules being added to `$global:DashboardConfig.LoadedModules` *after* successful execution of `Write-Module`.
			- Error reporting distinguishes between Critical, Important, and Optional module failures. Only Critical failures halt the application startup process.
		#>
		[CmdletBinding()]
		[OutputType([PSCustomObject])]
		param()
		
		Write-Verbose "Initializing module import process..." -ForegroundColor Cyan
		
		# Set up the return object structure and internal tracking variables.
		# $result - PSCustomObject to return. Start with default failure state.
		$result = [PSCustomObject]@{
			Status          = $false # Default to failure until proven successful.
			LoadedModules   = $global:DashboardConfig.LoadedModules # Use global directly, shows state during the run.
			FailedModules   = @{}    # List to store {ModuleName = ErrorMessage}.
			CriticalFailure = $false # Flag for critical module failures.
			Exception       = $null  # Placeholder for errors we didn't handle.
		}
		# $failedModules - Local reference to the list inside the result object for easier updates.
		[hashtable]$failedModules = $result.FailedModules
		
		try
		{
			#region Step: Determine Execution Context (EXE vs. Script)
				# Get info about the current running process.
				# $currentProcess - Process object for the current PowerShell instance.
				$currentProcess = Get-Process -Id $PID -ErrorAction Stop # Use Get-Process instead of GetCurrentProcess() for consistent MainModule access.

				# $processPath - Text, the full path of the program file for the current process. Use Path property.
				[string]$processPath = $currentProcess.Path # Use Path property, usually more reliable

				# $isRunningAsExe - Flag ($true/$false).
				[bool]$isRunningAsExe = $processPath -like '*.exe' -and ($processPath -notlike '*powershell.exe' -and $processPath -notlike '*pwsh.exe')
				
				# Ensure State exists before setting IsRunningAsExe
				if ($global:DashboardConfig -and -not $global:DashboardConfig.ContainsKey('State')) {
					$global:DashboardConfig['State'] = @{}
				}
				if ($global:DashboardConfig -and $global:DashboardConfig.State) {
					$global:DashboardConfig.State.IsRunningAsExe = $isRunningAsExe # Store globally.
				}
				Write-Verbose "  Execution context detected: $(if($isRunningAsExe){'Compiled EXE'} else {'PowerShell Script'}) (Process Path: '$processPath')" -ForegroundColor DarkGray
			#endregion Step: Determine Execution Context (EXE vs. Script)
			
			#region Step: Sort Modules by Defined 'Order' Property
				Write-Verbose "  Sorting modules based on 'Order' property..." -ForegroundColor DarkGray
				# $sortedModules - A list of module entries (Key/Value pairs) sorted by the 'Order' value in the module's info.
				# Need to handle errors if module config is messed up.
				$sortedModules = $global:DashboardConfig.Modules.GetEnumerator() |
				Where-Object {
					# Basic check: Make sure key exists and value is a hashtable with an 'Order' property.
					$_.Value -is [hashtable] -and $_.Value.ContainsKey('Order') -and $_.Value.Order -is [int]
				} |
				Sort-Object { $_.Value.Order } -ErrorAction SilentlyContinue # Sort based on the number 'Order' value.
				
				if (-not $sortedModules -or $sortedModules.Count -ne $global:DashboardConfig.Modules.Count)
				{
					# Check if sorting failed or if some modules were skipped due to bad structure.
					$invalidModules = $global:DashboardConfig.Modules.GetEnumerator() | Where-Object { -not ($_.Value -is [hashtable] -and $_.Value.ContainsKey('Order') -and $_.Value.Order -is [int]) }
					$errorMessage = "  Failed to sort modules or found invalid module configurations. Check structure in `$global:DashboardConfig.Modules."
					if ($invalidModules)
					{
						$errorMessage += " Invalid modules: $($invalidModules.Key -join ', ')"
					}
					Write-Verbose $errorMessage -ForegroundColor Red
					$result.Status = $false
					$result.CriticalFailure = $true # Treat sorting/config errors as critical.
					$failedModules['Module Sorting/Validation'] = $errorMessage
					return $result # Return failure right away.
				}
				Write-Verbose "  Processing $($sortedModules.Count) modules in defined order." -ForegroundColor DarkGray
			#endregion Step: Sort Modules by Defined 'Order' Property
			
			#region Step: Process Each Module in Sorted Order
				foreach ($entry in $sortedModules)
				{
					# $moduleName - Text, the key/filename of the module (e.g., 'ui.psm1').
					[string]$moduleName = $entry.Key
					# $moduleInfo - Hashtable holding info for this module (Priority, Order, Dependencies, FilePath/Base64Content).
					$moduleInfo = $entry.Value # Already checked as a hashtable during sorting.
					
					Write-Verbose "Processing Module: '$moduleName' (Priority: $($moduleInfo.Priority), Order: $($moduleInfo.Order))" -ForegroundColor Cyan
					
					#region SubStep: Check Dependencies
						Write-Verbose "- Checking dependencies..." -ForegroundColor DarkGray
						# $dependenciesMet - Flag ($true/$false), assume true until a missing dependency found.
						[bool]$dependenciesMet = $true
						# Check if Dependencies key exists, is an array, and has items.
						if ($moduleInfo.Dependencies -and $moduleInfo.Dependencies -is [array] -and $moduleInfo.Dependencies.Count -gt 0)
						{
							Write-Verbose "  - Required: $($moduleInfo.Dependencies -join ', ')" -ForegroundColor DarkGray
							foreach ($dependency in $moduleInfo.Dependencies)
							{
								# Check if the dependency is a key in the *global* loaded modules list.
								if (-not $global:DashboardConfig.LoadedModules.ContainsKey($dependency))
								{
									$errorMessage = "- Dependency NOT MET: Module '$dependency' must be loaded before '$moduleName'."
									Write-Verbose "- $errorMessage" -ForegroundColor Yellow
									$failedModules[$moduleName] = $errorMessage
									$dependenciesMet = $false
									# Check if this failure is critical.
									if ($moduleInfo.Priority -eq 'Critical')
									{
										Write-Verbose "- CRITICAL FAILURE: Critical module '$moduleName' cannot load due to missing dependency '$dependency'." -ForegroundColor Red
										$result.CriticalFailure = $true
									}
									break # No need to check more dependencies for this module.
								}
								else
								{
									Write-Verbose "  - Dependency satisfied: '$dependency' is loaded." -ForegroundColor DarkGray
								}
							}
						}
						else
						{
							Write-Verbose "  - No dependencies listed for '$moduleName'." -ForegroundColor DarkGray
						}
						
						# If dependencies aren't met, skip the rest of this module.
						if (-not $dependenciesMet)
						{
							continue
						} # Go to the next module in the loop.
					
					#endregion SubStep: Check Dependencies
					
					#region SubStep: Write Module to AppData Directory (Using Write-Module)
						# $modulePath - Text, path where module was written/checked. $null on failure.
						[string]$modulePath = $null
						Write-Verbose "- Ensuring module file exists in AppData via Write-Module for '$moduleName'..." -ForegroundColor DarkGray
						
						# Call Write-Module, giving inputs based on module's config (FilePath or Base64Content).
						try
						{
							if ($moduleInfo.ContainsKey('FilePath'))
							{
								[string]$sourceFilePath = $moduleInfo.FilePath
								# --- Add check for source file path ---
								if (-not (Test-Path $sourceFilePath -PathType Leaf)) {
									throw "Source FilePath specified in config does not exist or is not a file: '$sourceFilePath'"
								}
								Write-Verbose "Calling Write-Module with source FilePath: '$sourceFilePath'" -ForegroundColor Cyan
								$modulePath = Write-Module -ModuleName $moduleName -Content $sourceFilePath -ErrorAction Stop # Use Stop to catch errors here.
							}
							elseif ($moduleInfo.ContainsKey('Base64Content'))
							{
								[string]$base64Content = $moduleInfo.Base64Content
								Write-Verbose "Calling Write-Module with Base64Content (Length: $($base64Content.Length))" -ForegroundColor Cyan
								# Make sure content isn't null/empty before passing
								if ([string]::IsNullOrEmpty($base64Content))
								{
									throw "Base64Content for module '$moduleName' is empty."
								}
								$modulePath = Write-Module -ModuleName $moduleName -ContentBase64 $base64Content -ErrorAction Stop
							}
							else
							{
								# Shouldn't get here if sorting check worked.
								throw "Invalid module configuration format for '$moduleName' - missing FilePath or Base64Content."
							}
								
							# Check if Write-Module returned a valid path.
							if ([string]::IsNullOrEmpty($modulePath))
							{
								# Write-Module should ideally error out on failure with ErrorAction Stop, but double-check.
								throw "Write-Module returned null or empty path for '$moduleName', indicating write failure."
							}
								
							Write-Verbose "- [OK] Module file ready/verified: '$modulePath'" -ForegroundColor Green
							# Add/Update path in global loaded modules list. Happens for ALL written files (.psm1, .dll, .ico).
							# This is key for checking dependencies of non-PSM1 files.
							# Ensure LoadedModules hashtable exists
							if ($global:DashboardConfig -and -not $global:DashboardConfig.ContainsKey('LoadedModules')) {
								$global:DashboardConfig['LoadedModules'] = @{}
							}
							if ($global:DashboardConfig -and $global:DashboardConfig.LoadedModules) {
								$global:DashboardConfig.LoadedModules[$moduleName] = $modulePath
							}
								
						}
						catch
						{
							# Catch errors from Write-Module call or the code block above.
							$errorMessage = "- Failed to write or verify module file for '$moduleName'. Error: $($_.Exception.Message)"
							Write-Verbose "- $errorMessage" -ForegroundColor Red
							$failedModules[$moduleName] = $errorMessage
							# Check if this failure is critical.
							if ($moduleInfo.Priority -eq 'Critical')
							{
								Write-Verbose "- CRITICAL FAILURE: Failed to write critical module '$moduleName'." -ForegroundColor Red
								$result.CriticalFailure = $true
							}
							continue # Go to the next module.
						}
					#endregion SubStep: Write Module to AppData Directory (Using Write-Module)
						
					#region SubStep: Import PowerShell Modules (.psm1)
						# Only try PowerShell import steps if the module is a .psm1 file.
						if ($moduleName -like '*.psm1')
						{
							Write-Verbose "Attempting to import PowerShell module '$moduleName' from '$modulePath'..." -ForegroundColor Cyan
							# $importSuccess - Flag ($true/$false) for successful import of this specific PSM1 module.
							[bool]$importSuccess = $false
							# $importErrorDetails - Text to store failure details if all tries fail.
							[string]$importErrorDetails = 'All import attempts failed.'
							[string]$moduleBaseName = [System.IO.Path]::GetFileNameWithoutExtension($moduleName)

							# --- Import Try 1: Standard Import-Module (with EXE changes if needed) ---
							if (-not $importSuccess)
							{
								Write-Verbose "- Attempt 1: Using standard Import-Module..." -ForegroundColor Cyan
								try
								{
									# $effectiveModulePath - Path for Import-Module (might be temp path for EXE).
									[string]$effectiveModulePath = $modulePath
									# $tempModulePath - Path to temp changed file if running as EXE.
									[string]$tempModulePath = $null
										
									if ($isRunningAsExe)
									{
										Write-Verbose "  - (Running as EXE: Prepending no-op Export-ModuleMember to temporary file for import)" -ForegroundColor DarkGray
										# Create a unique temporary file path.
										$tempModulePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ('{0}_{1}.psm1' -f $moduleBaseName, [System.Guid]::NewGuid().ToString('N')))
										Write-Verbose "  - Temp file path: $tempModulePath" -ForegroundColor DarkGray
											
										# Read original content (already in global config or re-read to be safe).
										# $originalContent - Text, content of the module.
										# Ensure Resources and LoadedModuleContent exist
										[string]$originalContent = $null
										if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and $global:DashboardConfig.Resources.LoadedModuleContent -and $global:DashboardConfig.Resources.LoadedModuleContent.ContainsKey($moduleName)) {
											$originalContent = $global:DashboardConfig.Resources.LoadedModuleContent[$moduleName]
										}
										if ($null -eq $originalContent)
										{
											$originalContent = [System.IO.File]::ReadAllText($modulePath)
										} # Re-read if not found
											
										# Define the dummy function text.
										$noOpExportFunc = "function Export-ModuleMember { param([Parameter(ValueFromPipeline=`$true)][string[]]`$Function, [string[]]`$Variable, [string[]]`$Alias, [string[]]`$Cmdlet) { Write-Verbose ""Ignoring Export-ModuleMember (EXE Mode Import for '$($using:moduleName)')"" -ForegroundColor Cyan} }"
										# Add to beginning and write to temp file using UTF8 encoding.
										Set-Content -Path $tempModulePath -Value "$noOpExportFunc`n`n# --- Original module content ($moduleName) follows ---`n$originalContent" -Encoding UTF8 -Force -ErrorAction Stop
										$effectiveModulePath = $tempModulePath # Use the temp path for import.
									}
										
									# Run Import-Module. -Force re-imports if already loaded (good for dev/debug).
									Import-Module -Name $effectiveModulePath -Force -ErrorAction Stop
										
									# Check module loaded okay by using Get-Module with the base name.
									if (Get-Module -Name $moduleBaseName -ErrorAction SilentlyContinue)
									{
										$importSuccess = $true
										Write-Verbose "- [OK] Attempt 1: SUCCESS (Standard Import-Module verified for '$moduleBaseName')." -ForegroundColor Green
									}
									else
									{
										# This might happen if Import-Module finishes but the module somehow doesn't show up right.
										Write-Verbose "- Attempt 1: FAILED (Standard Import-Module) - Module '$moduleBaseName' not found via Get-Module after import call." -ForegroundColor Yellow
										$importErrorDetails = "Standard Import-Module completed but module '$moduleBaseName' could not be verified via Get-Module."
										# If import failed, make sure any existing module state is removed before trying next way.
										Remove-Module -Name $moduleBaseName -Force -ErrorAction SilentlyContinue
									}
										
								}
								catch
								{
									Write-Verbose "- Attempt 1: FAILED (Standard Import-Module Error): $($_.Exception.Message)" -ForegroundColor Yellow
									$importErrorDetails = "Standard Import-Module Error: $($_.Exception.Message)"
									# Make sure any partial/failed module state is removed.
									Remove-Module -Name $moduleBaseName -Force -ErrorAction SilentlyContinue
								}
								finally
								{
									# Clean up temp file if one was made for EXE mode.
									if ($tempModulePath -and (Test-Path $tempModulePath))
									{
										Write-Verbose "  - Cleaning up temporary file: $tempModulePath" -ForegroundColor DarkGray
										Remove-Item -Path $tempModulePath -Force -ErrorAction SilentlyContinue
									}
								}
							} # End Try 1
																
							# --- Import Try 2: Import-ModuleUsingReflection (WITH VERIFICATION) ---
							if (-not $importSuccess)
							{
								# Check if the alternative function exists.
								if (Get-Command Import-ModuleUsingReflection -ErrorAction SilentlyContinue)
								{
									Write-Verbose "- Attempt 2: Using alternative Import-ModuleUsingReflection (InvokeCommand)..." -ForegroundColor Yellow
									try
									{
										# Call the reflection import function
										if (Import-ModuleUsingReflection -Path $modulePath -ModuleName $moduleName -ErrorAction Stop)
										{
											# --- Verification Step Added ---
											Write-Verbose "  - Attempt 2: InvokeCommand finished. Verifying key functions globally for '$moduleName'..." -ForegroundColor Magenta

											$moduleFileName = (Split-Path $moduleName -Leaf).Trim().ToLower()
											Write-Verbose "DEBUG: moduleFileName is '$moduleFileName' - Length: $($moduleFileName.Length)" -ForegroundColor Magenta
											# Get the list of expected functions for this module (reuse from Attempt 3 logic)
											$keyFunctionsToCapture = @()
											if ($moduleFileName -in @('ini.psm1', 'ini')) { 
												$keyFunctionsToCapture = @('Copy-OrderedDictionary','Get-IniFileContent','LoadDefaultConfigOnError','Initialize-IniConfig','Read-Config','Write-Config') 
											}
											elseif ($moduleFileName -in @('ui.psm1', 'ui')) { 
												$keyFunctionsToCapture = @('Initialize-UI','Register-UIEventHandlers','Show-SettingsForm','Hide-SettingsForm','Set-UIElement','Sync-ConfigToUI','Sync-UIToConfig','Show-InputBox','RefreshLoginProfileSelector','Sync-ProfilesToConfig') 
											}
											elseif ($moduleFileName -in @('datagrid.psm1', 'datagrid')) { 
												$keyFunctionsToCapture = @('Test-ValidParameters','Restore-WindowStyles','Get-ProcessList','Remove-TerminatedProcesses','New-RowLookupDictionary','Update-ExistingRow','UpdateRowIndices','Add-NewProcessRow','Start-WindowStateCheck','Find-TargetRow','Clear-OldProcessCache','Get-ProcessProfile','Set-WindowToolStyle','Update-DataGrid','Start-DataGridUpdateTimer' ) 
											}
											elseif ($moduleFileName -in @('launch.psm1', 'launch')) { 
												$keyFunctionsToCapture = @('Start-ClientLaunch','Write-Verbose','Invoke-SavedLaunchSequence','Stop-ClientLaunch') 
											}
											elseif ($moduleFileName -in @('login.psm1', 'login')) { 
												$keyFunctionsToCapture = @('Get-ClientLogPath','Update-Progress','LoginSelectedRow','CleanUpLoginResources') 
											}
											elseif ($moduleFileName -in @('ftool.psm1', 'ftool')) { 
												$keyFunctionsToCapture = @('Set-Hotkey','Resume-AllHotkeys','Resume-PausedKeys','Resume-HotkeysForOwner','Remove-AllHotkeys','Test-HotkeyConflict','Invoke-FtoolAction','PauseAllHotkeys','PauseHotkeysForOwner','Unregister-HotkeyInstance','ToggleSpecificFtoolInstance','ToggleInstanceHotkeys','Get-VirtualKeyMappings','NormalizeKeyString','ParseKeyString','Get-KeyCombinationString','IsModifierKeyCode','Show-KeyCaptureDialog','LoadFtoolSettings','FindOrCreateProfile','InitializeExtensionTracking','GetNextExtensionNumber','FindExtensionKeyByControl','LoadExtensionSettings','UpdateSettings','IsWindowBelow','CreatePositionTimer','RepositionExtensions','CreateSpammerTimer','ToggleButtonState','CheckRateLimit','AddFormCleanupHandler','CleanupInstanceResources','Stop-FtoolForm','RemoveExtension','FtoolSelectedRow','CreateFtoolForm','AddFtoolEventHandlers','CreateExtensionPanel','AddExtensionEventHandlers') 
											}
											elseif ($moduleFileName -in @('reconnect.psm1', 'reconnect')) { 
												$keyFunctionsToCapture = @('CheckCancel','SleepWithCancel','Invoke-GuardedAction','Update-NotificationPositions','Close-Notification','Show-InteractiveNotification','ClearGameLogs','Start-DisconnectWatcher','Stop-DisconnectWatcher','Invoke-ReconnectionSequence') 
											}
											else {
												Write-Verbose "WARNING: No verification list found for module '$moduleFileName' (Length: $($moduleFileName.Length)). verification skipped." -ForegroundColor Yellow
											}

											[bool]$attempt2VerificationPassed = $true # Assume success until proven otherwise
											[string]$missingFunction = $null

											if ($keyFunctionsToCapture.Count -gt 0) {
												foreach ($funcName in $keyFunctionsToCapture) {
													if (-not (Get-Command -Name $funcName -CommandType Function -ErrorAction SilentlyContinue)) {
														$attempt2VerificationPassed = $false
														$missingFunction = $funcName
														Write-Verbose "  - Attempt 2: VERIFICATION FAILED. Function '$funcName' not found globally after InvokeCommand." -ForegroundColor Red
														$importErrorDetails = "Attempt 2 (InvokeCommand) completed but verification failed: Function '$funcName' not found globally."
														break # Stop checking if one is missing
													}
												}
											} else {
												Write-Verbose "  - Attempt 2: No specific key functions listed for verification for '$moduleName'. Assuming success based on InvokeCommand completion." -ForegroundColor DarkGray
												# If no functions to verify, trust the $true return from Import-ModuleUsingReflection
												$attempt2VerificationPassed = $true 
											}

											# Set final import success based on verification
											if ($attempt2VerificationPassed) {
												Write-Verbose "- [OK] Attempt 2: SUCCESS (InvokeCommand completed AND key functions verified for '$moduleName')." -ForegroundColor Green
												$importSuccess = $true
											} else {
												# Failure already logged above
												$importSuccess = $false
											}
											# --- End Verification Step ---
										}
										else # Import-ModuleUsingReflection returned false (fatal error during its execution)
										{
											Write-Verbose "- Attempt 2: FAILED (Import-ModuleUsingReflection returned false)." -ForegroundColor Yellow
											$importErrorDetails = 'Import-ModuleUsingReflection returned false (fatal execution error).'
											$importSuccess = $false # Ensure flag is false
										}
									}
									catch # Catch errors *calling* Import-ModuleUsingReflection
									{
										Write-Verbose "- Attempt 2: FAILED (Error calling Import-ModuleUsingReflection): $($_.Exception.Message)" -ForegroundColor Yellow
										$importErrorDetails = "Error calling Import-ModuleUsingReflection: $($_.Exception.Message)"
										$importSuccess = $false # Ensure flag is false
									}
								}
								else # Import-ModuleUsingReflection command not found
								{
									Write-Verbose "- Attempt 2: SKIPPED (Import-ModuleUsingReflection function not found)." -ForegroundColor Yellow
								}
							} # End Try 2
								
							# --- Import Try 3: Direct Invoke-Expression (Last Resort - Security Risk!) ---
							# This only runs if $importSuccess is still $false after Attempt 1 and Attempt 2 (including verification)
							if (-not $importSuccess)
							{
								Write-Verbose "- Attempt 3: Using LAST RESORT Invoke-Expression (Security Risk!)..." -ForegroundColor Yellow
								# === Add a variable to track functions caught just in *this* try ===
								$functionsCapturedInThisAttempt = @{}
								try
								{
									# Read module content (might be saved in global config).
									[string]$invokeContent = $null
									if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and $global:DashboardConfig.Resources.LoadedModuleContent -and $global:DashboardConfig.Resources.LoadedModuleContent.ContainsKey($moduleName)) {
										$invokeContent = $global:DashboardConfig.Resources.LoadedModuleContent[$moduleName]
									}
									if ($null -eq $invokeContent)
									{
										$invokeContent = [System.IO.File]::ReadAllText($modulePath)
									} # Re-read if needed.

									# Basic try to disable Export-ModuleMember calls using multi-line regex replace.
									$invokeContent = $invokeContent -replace '(?m)^\s*Export-ModuleMember.*', "# Export-ModuleMember call disabled by Invoke-Expression wrapper for $moduleName"

									# Run the (maybe changed) content directly in the global space.
									Invoke-Expression -Command $invokeContent -ErrorAction Stop

									# First check if IEX finished without MAJOR error
									$iexCompletedWithoutTerminatingError = $?
									$moduleFileName = (Split-Path $moduleName -Leaf).Trim().ToLower()
									Write-Verbose "DEBUG: moduleFileName is '$moduleFileName' - Length: $($moduleFileName.Length)" -ForegroundColor Magenta
									# Check key functions right away AND grab them if found
									$keyFunctionsToCapture = @()
									# --- LIST ALL EXPECTED EXPORTED/USED FUNCTIONS FOR EACH MODULE ---
									if ($moduleFileName -in @('ini.psm1', 'ini')) { 
										$keyFunctionsToCapture = @('Copy-OrderedDictionary','Get-IniFileContent','LoadDefaultConfigOnError','Initialize-IniConfig','Read-Config','Write-Config') 
									}
									elseif ($moduleFileName -in @('ui.psm1', 'ui')) { 
										$keyFunctionsToCapture = @('Initialize-UI','Register-UIEventHandlers','Show-SettingsForm','Hide-SettingsForm','Set-UIElement','Sync-ConfigToUI','Sync-UIToConfig','Show-InputBox','RefreshLoginProfileSelector','Sync-ProfilesToConfig') 
									}
									elseif ($moduleFileName -in @('datagrid.psm1', 'datagrid')) { 
										$keyFunctionsToCapture = @('Test-ValidParameters','Restore-WindowStyles','Get-ProcessList','Remove-TerminatedProcesses','New-RowLookupDictionary','Update-ExistingRow','UpdateRowIndices','Add-NewProcessRow','Start-WindowStateCheck','Find-TargetRow','Clear-OldProcessCache','Get-ProcessProfile','Set-WindowToolStyle','Update-DataGrid','Start-DataGridUpdateTimer' ) 
									}
									elseif ($moduleFileName -in @('launch.psm1', 'launch')) { 
										$keyFunctionsToCapture = @('Start-ClientLaunch','Write-Verbose','Invoke-SavedLaunchSequence','Stop-ClientLaunch') 
									}
									elseif ($moduleFileName -in @('login.psm1', 'login')) { 
										$keyFunctionsToCapture = @('Get-ClientLogPath','Update-Progress','LoginSelectedRow','CleanUpLoginResources') 
									}
									elseif ($moduleFileName -in @('ftool.psm1', 'ftool')) { 
										$keyFunctionsToCapture = @('Set-Hotkey','Resume-AllHotkeys','Resume-PausedKeys','Resume-HotkeysForOwner','Remove-AllHotkeys','Test-HotkeyConflict','Invoke-FtoolAction','PauseAllHotkeys','PauseHotkeysForOwner','Unregister-HotkeyInstance','ToggleSpecificFtoolInstance','ToggleInstanceHotkeys','Get-VirtualKeyMappings','NormalizeKeyString','ParseKeyString','Get-KeyCombinationString','IsModifierKeyCode','Show-KeyCaptureDialog','LoadFtoolSettings','FindOrCreateProfile','InitializeExtensionTracking','GetNextExtensionNumber','FindExtensionKeyByControl','LoadExtensionSettings','UpdateSettings','IsWindowBelow','CreatePositionTimer','RepositionExtensions','CreateSpammerTimer','ToggleButtonState','CheckRateLimit','AddFormCleanupHandler','CleanupInstanceResources','Stop-FtoolForm','RemoveExtension','FtoolSelectedRow','CreateFtoolForm','AddFtoolEventHandlers','CreateExtensionPanel','AddExtensionEventHandlers') 
									}
									elseif ($moduleFileName -in @('reconnect.psm1', 'reconnect')) { 
										$keyFunctionsToCapture = @('CheckCancel','SleepWithCancel','Invoke-GuardedAction','Update-NotificationPositions','Close-Notification','Show-InteractiveNotification','ClearGameLogs','Start-DisconnectWatcher','Stop-DisconnectWatcher','Invoke-ReconnectionSequence') 
									}
									else {
										Write-Verbose "WARNING: No verification list found for module '$moduleFileName' (Length: $($moduleFileName.Length)). verification skipped." -ForegroundColor Yellow
									}
									# Create captured functions storage if it doesn't exist
									if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and -not $global:DashboardConfig.Resources.ContainsKey('CapturedFunctions')) {
										$global:DashboardConfig.Resources['CapturedFunctions'] = @{}
									}

									$captureSuccess = $true # Assume capture worked at first
									$criticalFunctionMissing = $false

									if ($keyFunctionsToCapture.Count -gt 0) {
										Write-Verbose "- Attempt 3: Verifying and capturing key functions for '$moduleName' immediately after IEX..." -ForegroundColor Magenta
										foreach ($funcName in $keyFunctionsToCapture) {
											$funcInfo = Get-Command -Name $funcName -CommandType Function -ErrorAction SilentlyContinue
											if ($funcInfo) {
												$capturedScriptBlock = $funcInfo.ScriptBlock
												Write-Verbose "  - Found and capturing ScriptBlock for '$funcName'." -ForegroundColor Magenta
												# Store globally for possible later use (though direct global definition is main now)
												if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and $global:DashboardConfig.Resources.CapturedFunctions) {
													$global:DashboardConfig.Resources.CapturedFunctions[$funcName] = $capturedScriptBlock
												}
												# === Store locally for immediate global definition ===
												$functionsCapturedInThisAttempt[$funcName] = $capturedScriptBlock
											} else {
												Write-Verbose "  - WARNING: Could not find/capture function '$funcName' immediately after IEX for '$moduleName'." -ForegroundColor Yellow
												$captureSuccess = $false
												# Check if the missing function is critical FOR STARTUP
												# --- Adjusted Critical Function Check ---
												$isCriticalModule = $moduleInfo.Priority -eq 'Critical' 
												# Consider a function critical if it's in a Critical module AND in the key function list
												if ($isCriticalModule) { 
													$criticalFunctionMissing = $true
													$importErrorDetails += "; Critical function '$funcName' not found after IEX in Critical module '$moduleName'"
													Write-Verbose "    - Missing function '$funcName' is considered critical for module '$moduleName'." -ForegroundColor Red
												} else {
													 $importErrorDetails += "; Non-critical function '$funcName' not found after IEX for module '$moduleName'"
												}
												# --- End Adjusted Critical Function Check ---
											}
										}
									}

									# Decide overall success for Try 3
									# Success means IEX didn't have non-terminating errors ($?),
									# capture succeeded, AND no *critical* functions were missing.
									if ($iexCompletedWithoutTerminatingError -and $captureSuccess -and (-not $criticalFunctionMissing)) {
										Write-Verbose "  - Attempt 3: IEX completed and key functions captured/verified for '$moduleName'." -ForegroundColor DarkGreen

										# === Define captured functions globally RIGHT AWAY ===
										Write-Verbose "  - Defining captured functions globally for '$moduleName'..." -ForegroundColor Magenta
										$definitionSuccess = $true # Track success of this small step
										foreach ($kvp in $functionsCapturedInThisAttempt.GetEnumerator()) {
											$funcNameToDefine = $kvp.Key
											$scriptBlockToDefine = $kvp.Value
											try {
												# Define in global function space
												Set-Item -Path "Function:\global:$funcNameToDefine" -Value $scriptBlockToDefine -Force -ErrorAction Stop
												Write-Verbose "    - Defined Function:\global:$funcNameToDefine" -ForegroundColor DarkMagenta
											} catch {
												Write-Verbose "    - FAILED to define Function:\global:$funcNameToDefine globally: $($_.Exception.Message)" -ForegroundColor Red
												$definitionSuccess = $false
												$importErrorDetails += "; Failed to define captured function '$funcNameToDefine' globally."
												# If defining a critical function fails, mark critical failure for the whole import process
												# --- Adjusted Critical Function Check ---
												if ($moduleInfo.Priority -eq 'Critical') {
													$result.CriticalFailure = $true
													Write-Verbose "    - Defining critical function '$funcNameToDefine' failed. Marking import as critical failure." -ForegroundColor Red
												}
												 # --- End Adjusted Critical Function Check ---
												break # Stop trying to define others for this module if one fails
											}
										}

										# Only mark the whole import successful if definitions also worked
										if ($definitionSuccess) {
											$importSuccess = $true
											Write-Verbose "- [OK] Attempt 3: SUCCESS (Invoke-Expression completed, key functions captured AND globally defined for $moduleName)." -ForegroundColor Green
										} else {
											$importSuccess = $false # Failed during definition
											Write-Verbose "- Attempt 3: FAILED during global definition phase for $moduleName." -ForegroundColor Red
										}

									} else { # IEX failed, capture failed, or critical function missing
										Write-Verbose "- Attempt 3: FAILED (IEX completed=$iexCompletedWithoutTerminatingError, CaptureSuccess=$captureSuccess, CriticalMissing=$criticalFunctionMissing) for $moduleName." -ForegroundColor Red
										if (-not $iexCompletedWithoutTerminatingError) { $importErrorDetails += "; IEX failed with non-terminating error detected by `$?."}
										if ($criticalFunctionMissing) { $importErrorDetails += "; Critical function missing prevented Attempt 3 success." }
										if (-not $captureSuccess) { $importErrorDetails += "; Function capture failed during Attempt 3." }
										$importSuccess = $false # Make sure import is marked as failed
									}
								}
								catch # Catch MAJOR errors from Invoke-Expression itself
								{
									Write-Verbose "- Attempt 3: FAILED (Invoke-Expression Error): $($_.Exception.Message)" -ForegroundColor Red
									$importErrorDetails = "Invoke-Expression Error: $($_.Exception.Message)"
									$importSuccess = $false # Make sure success is false if IEX throws major error
								}
							} # End Try 3

							# --- Final Check for PSM1 Import Success --- 
							if ($importSuccess)
							{
								Write-Verbose "- [OK] Successfully imported PSM1 module: '$moduleName'." -ForegroundColor Green
								# Module already added to $global:DashboardConfig.LoadedModules after Write-Module step.
							}
							else
							{
								$errorMessage = "All import methods FAILED for PSM1 module: '$moduleName'. Last error detail: $importErrorDetails"
								Write-Verbose "- $errorMessage" -ForegroundColor Red
								$failedModules[$moduleName] = $errorMessage
								# Critical: Remove from LoadedModules list if import failed after writing okay, as it's not really usable.
								if ($global:DashboardConfig -and $global:DashboardConfig.LoadedModules -and $global:DashboardConfig.LoadedModules.ContainsKey($moduleName))
								{
									Write-Verbose "- Removing '$moduleName' from LoadedModules list due to import failure." -ForegroundColor Yellow
									$global:DashboardConfig.LoadedModules.Remove($moduleName)
								}
								# Check if this failure is critical.
								if ($moduleInfo.Priority -eq 'Critical')
								{
									Write-Verbose "- CRITICAL FAILURE: Failed to import critical PSM1 module '$moduleName'." -ForegroundColor Red
									$result.CriticalFailure = $true
								}
							}
						}
					#endregion SubStep: Import PowerShell Modules (.psm1)
				} # End foreach ($entry in $sortedModules)
			#endregion Step: Process Each Module in Sorted Order
				
			#region Step: Final Status Check and Result Construction
				Write-Verbose "Module import check..." -ForegroundColor Cyan
					
				# Check for Critical Failures gathered during the loop.
				if ($result.CriticalFailure)
				{
					Write-Verbose "  CRITICAL FAILURE: One or more critical modules failed to load or write. Application cannot continue." -ForegroundColor Red
					# Find which critical modules exactly failed.
					$criticalModules = $global:DashboardConfig.Modules.GetEnumerator() | Where-Object { $_.Value.Priority -eq 'Critical' }
					$failedCritical = $criticalModules | Where-Object { $failedModules.ContainsKey($_.Key) }
					if ($failedCritical)
					{
						Write-Verbose "  Failed critical modules: $($failedCritical.Key -join ', ')" -ForegroundColor Red
						$failedCritical | ForEach-Object { Write-Verbose "  - $($_.Key): $($failedModules[$_.Key])" -ForegroundColor Red } 
					}
					$result.Status = $false # Make sure status is false.
					# Return the result object showing critical failure.
					return $result
				}
					
				# Report Important Module Failures (as Warnings).
				$importantModules = $global:DashboardConfig.Modules.GetEnumerator() | Where-Object { $_.Value.Priority -eq 'Important' }
				$failedImportant = $importantModules | Where-Object { $failedModules.ContainsKey($_.Key) }
				if ($failedImportant.Count -gt 0)
				{
					Write-Verbose "  IMPORTANT module failures detected: $($failedImportant.Key -join ', '). Application may have limited functionality." -ForegroundColor Yellow
					# Log details of failures for important modules.
					$failedImportant | ForEach-Object { Write-Verbose "  - $($_.Key): $($failedModules[$_.Key])" -ForegroundColor Yellow } 
				}
					
				# Report Optional Module Failures (as Info/DarkYellow).
				$optionalModules = $global:DashboardConfig.Modules.GetEnumerator() | Where-Object { $_.Value.Priority -eq 'Optional' }
				$failedOptional = $optionalModules | Where-Object { $failedModules.ContainsKey($_.Key) }
				if ($failedOptional.Count -gt 0)
				{
					Write-Verbose "  Optional module failures detected: $($failedOptional.Key -join ', '). Non-essential features might be unavailable." -ForegroundColor DarkYellow
					# Log details.
					$failedOptional | ForEach-Object { Write-Verbose "  - $($_.Key): $($failedModules[$_.Key])" -ForegroundColor DarkGray }
				}
					
				# If no critical failures happened, the whole process is seen as successful for startup.
				$successCount = 0
				if ($global:DashboardConfig -and $global:DashboardConfig.LoadedModules) {
					$successCount = $global:DashboardConfig.LoadedModules.Count
				}
				$failCount = $failedModules.Count
				Write-Verbose "  Module loading phase complete. Modules written/verified: $successCount. Failures (any type): $failCount." -ForegroundColor DarkGray
				if ($successCount -gt 0)
				{
					Write-Verbose "  Successfully written/verified modules: $($global:DashboardConfig.LoadedModules.Keys -join ', ')" -ForegroundColor DarkGray
				}
				if ($failCount -gt 0)
				{
					Write-Verbose "  Failed modules logged above." -ForegroundColor Yellow
				}
					
				# Set final status to true as no critical failures happened.
				$result.Status = $true
				$result.CriticalFailure = $false # Explicitly set false.
				# Return the final result object.
				return $result
			#endregion Step: Final Status Check and Result Construction
		}
		catch
		{
			# Catch surprise, unhandled errors within the main Import-DashboardModules function body.
			$errorMessage = "  FATAL UNHANDLED EXCEPTION in Import-DashboardModules: $($_.Exception.Message)"
			Write-Verbose $errorMessage -ForegroundColor Red
			Write-Verbose "  Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
			# Fill and return the result object showing critical failure due to the error.
			$result.Status = $false
			$result.CriticalFailure = $true
			$result.Exception = $_.Exception.Message # Store error message.
			$failedModules['Unhandled Exception'] = $errorMessage # Add to failed modules list.
			return $result
		}
	}
#endregion Function: Import-DashboardModules

#endregion Module Handling Functions

#region UI and Application Lifecycle Functions

#region Function: Start-Dashboard
	function Start-Dashboard
	{
		<#
			.SYNOPSIS
				Initializes and displays the main dashboard user interface (UI) form.
			
			.DESCRIPTION
				This function orchestrates the startup of the application's graphical user interface. It performs these actions:
				1. Checks if the 'Initialize-UI' function, expected to be loaded from the 'ui.psm1' module, exists using `Get-Command`. If not found, it throws a terminating error as the UI cannot be built.
				2. Calls the `Initialize-UI` function. It assumes this function is responsible for creating all UI elements (forms, controls) and populating the '$global:DashboardConfig.UI' hashtable, including setting '$global:DashboardConfig.UI.MainForm'.
				3. Checks the return value of `Initialize-UI`. If it returns $false or null (interpreted as failure), it throws a terminating error.
				4. Verifies that '$global:DashboardConfig.UI.MainForm' exists and is a valid '[System.Windows.Forms.Form]' object after `Initialize-UI` returns successfully. If not, it throws a terminating error.
				5. If the MainForm is valid, it calls the `.Show()` method to make the main window visible and `.Activate()` to bring it to the foreground.
				6. Sets the global state flag '$global:DashboardConfig.State.UIInitialized' to $true.
			
			.OUTPUTS
				[bool] Returns $true if the UI is successfully initialized, the main form is found, shown, and activated.
				Returns $false if any step fails (missing function, initialization failure, missing main form), typically after throwing an error that gets caught by the main execution block.
			
			.NOTES
				- This function has a strong dependency on the 'ui.psm1' module being loaded correctly and functioning as expected (defining `Initialize-UI` and creating `MainForm`).
				- Errors encountered during this process are considered fatal for the application and are thrown to be caught by the main script's try/catch block, which should then display an error using `Show-ErrorDialog`.
		#>
		[CmdletBinding()]
		[OutputType([bool])]
		param()

		Write-Verbose "Starting Dashboard User Interface..." -ForegroundColor Cyan
		try
		{
			#region Step: Check for and Call Initialize-UI Function
				Write-Verbose "- Checking for required Initialize-UI function (from ui.psm1)..." -ForegroundColor DarkGray
				# Check that the Initialize-UI command (function) is available now.
				if (-not (Get-Command Initialize-UI -ErrorAction SilentlyContinue ))
				{
					# Throw a major error if the function is missing.
					throw "FATAL: Initialize-UI function not found. Ensure 'ui.psm1' module loaded correctly and defines this function."
				}

				Write-Verbose "- Calling Initialize-UI function..." -ForegroundColor DarkGray
				# Run the UI setup function. Save its return value.
				Initialize-UI # Call directly now

				Write-Verbose "- [OK] Initialize-UI function executed successfully." -ForegroundColor Green
			#endregion Step: Check for and Call Initialize-UI Function

			#region Step: Verify, Show, and Activate Main Form
				Write-Verbose "- Verifying presence and type of UI.MainForm object..." -ForegroundColor DarkGray
				# Check if MainForm property exists in UI config and is a valid Form object.
				if ($null -eq $global:DashboardConfig.UI.MainForm -or -not ($global:DashboardConfig.UI.MainForm -is [System.Windows.Forms.Form]))
				{
					# Throw a major error if main form is missing or invalid after successful Initialize-UI call.
					throw 'FATAL: UI.MainForm object not found or is not a valid System.Windows.Forms.Form in $global:DashboardConfig after successful Initialize-UI call.'
				}

				Write-Verbose "- [OK] UI.MainForm found and is valid. Showing and activating window..." -ForegroundColor Green
				# Make the main app window visible.
				$global:DashboardConfig.UI.MainForm.Show() 
				
				# Update the global state flag to show the UI is now set up and running.
				$global:DashboardConfig.State.UIInitialized = $true
				Write-Verbose "  Dashboard UI started successfully." -ForegroundColor Green
			#endregion Step: Verify, Show, and Activate Main Form

			# Return true showing successful UI startup.
			return $true
		}
		catch
		{
			$errorMsg = "  FATAL: Failed to start dashboard UI. Error: $($_.Exception.Message)"
			Write-Verbose $errorMsg -ForegroundColor Red
			# Throw the error again to send it up to the main run block's catch.
			throw $_ # Use throw $_ to keep original error details.
		}
	}
#endregion Function: Start-Dashboard
	
#region Function: Start-MessageLoop
	function Start-MessageLoop
	{
		<#
			.SYNOPSIS
				Runs the Windows Forms message loop to keep the UI responsive until the main form is closed.
			
			.DESCRIPTION
				This function implements the core message processing loop required for a Windows Forms application. It keeps the UI alive and responsive to user interactions, window events, and timer ticks.
				
				The function first performs pre-checks:
				1. Verifies that the UI has been initialized (`$global:DashboardConfig.State.UIInitialized`).
				2. Verifies that the main form object (`$global:DashboardConfig.UI.MainForm`) exists, is a valid Form, and is not already disposed.
				
				If checks pass, it determines the loop method:
				- Preferred Native Loop: If the 'Native' class (expected from 'classes.psm1') and its required P/Invoke methods (`AsyncExecution`, `PeekMessage`, `TranslateMessage`, `DispatchMessage`) are detected, it uses an efficient loop based on `MsgWaitForMultipleObjectsEx` (wrapped in `AsyncExecution`). This waits for messages or a timeout, processing messages only when they arrive, thus minimizing CPU usage when idle.
				- Fallback DoEvents Loop: If the Native methods are unavailable, it falls back to a loop using `[System.Windows.Forms.Application]::DoEvents()`. This processes all pending messages but does not wait efficiently, potentially consuming more CPU. A short `Start-Sleep` (e.g., 20ms) is added within this loop to prevent 100% CPU usage.
				
				The chosen loop runs continuously as long as the main form (`$global:DashboardConfig.UI.MainForm`) is visible and not disposed.
			
			.OUTPUTS
				[void] This function runs synchronously and blocks execution until the main UI form is closed or disposed. It does not return a value.
			
			.NOTES
				- Requires the main UI form (`$global:DashboardConfig.UI.MainForm`) to be successfully initialized and shown by `Start-Dashboard` before being called.
				- The efficiency of the UI heavily depends on the availability and correctness of the 'Native' class methods from 'classes.psm1'. The `DoEvents` fallback is less performant.
				- Includes basic error handling within the loop itself and a final `DoEvents` fallback attempt if the primary loop method encounters an unhandled exception.
				- Logs the chosen loop method and status messages during execution and upon exit.
		#>
		[CmdletBinding()]
		[OutputType([void])]
		param()
			
		Write-Verbose "`Starting UI message loop..." -ForegroundColor Cyan
			
		#region Step: Pre-Loop Checks for UI State and Main Form Validity
			Write-Verbose "  Checking UI state before starting message loop..." -ForegroundColor DarkGray
			# Check if UI setup flag is set.
			if (-not $global:DashboardConfig.State.UIInitialized)
			{
				Write-Verbose "  UI not marked as initialized ($global:DashboardConfig.State.UIInitialized is $false). Skipping message loop." -ForegroundColor Yellow
				return # Exit function if UI isn't ready.
			}
			# Check if MainForm object exists and is a valid, non-disposed Form.
			$mainForm = $global:DashboardConfig.UI.MainForm # Local variable to make things easier.
			if ($null -eq $mainForm -or -not ($mainForm -is [System.Windows.Forms.Form]))
			{
				Write-Verbose "  MainForm object ($global:DashboardConfig.UI.MainForm) is missing or not a valid Form object. Cannot start message loop." -ForegroundColor Yellow
				return # Exit function if MainForm is invalid.
			}
			if ($mainForm.IsDisposed)
			{
				Write-Verbose "  MainForm ($global:DashboardConfig.UI.MainForm) is already disposed. Cannot start message loop." -ForegroundColor Yellow
				return # Exit function if MainForm is already disposed (cleaned up).
			}
			Write-Verbose "  Pre-loop checks passed. MainForm is valid and UI is initialized." -ForegroundColor Green
		#endregion Step: Pre-Loop Checks for UI State and Main Form Validity
			
		# $loopMethod - Text showing which loop type is used ('Native' or 'DoEvents').
		[string]$loopMethod = 'Unknown'
		try
		{
			#region Step: Determine Loop Method (Efficient Native P/Invoke vs. Fallback DoEvents)
				# $useNativeLoop - Flag ($true/$false), $true if Native methods seem available.
				[bool]$useNativeLoop = $false
				Write-Verbose "Detecting availability of Native methods for efficient loop..." -ForegroundColor Cyan
				try
				{
					# Check if the 'Native' type exists and has the key methods we need.
					# Use GetType() which errors if type not found, unlike PSTypeName.
					$nativeType = [type]'Custom.Native' # Errors if 'Native' class not loaded.
					if (($nativeType.GetMethod('AsyncExecution')) -and
						($nativeType.GetMethod('PeekMessage')) -and
						($nativeType.GetMethod('TranslateMessage')) -and
						($nativeType.GetMethod('DispatchMessage')))
					{
						Write-Verbose "- [OK] Native P/Invoke methods found (requires 'classes.psm1'). Using efficient message loop." -ForegroundColor Green
						$useNativeLoop = $true
						$loopMethod = 'Custom.Native'
					}
					else
					{
						Write-Verbose "- Native class found, but required methods (AsyncExecution, PeekMessage, etc.) are missing. Falling back to DoEvents loop." -ForegroundColor Yellow
						$loopMethod = 'DoEvents'
					}
				}
				catch [System.Management.Automation.RuntimeException]
				{
					# Catch specific error for type not found.
					Write-Verbose "- Native class not found. Falling back to less efficient Application.DoEvents() loop." -ForegroundColor Red
					$loopMethod = 'DoEvents'
				}
				catch
				{
					Write-Verbose "- Error checking for Native methods: $($_.Exception.Message). Falling back to DoEvents loop." -ForegroundColor Red
					$loopMethod = 'DoEvents'
				}
					
				# Make sure WinForms part is loaded if using DoEvents backup.
				if (-not $useNativeLoop)
				{
					Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue # Load if not already loaded.
				}
			#endregion Step: Determine Loop Method (Efficient Native P/Invoke vs. Fallback DoEvents)
				
			#region Step: Run the Chosen Message Loop
				Write-Verbose "Entering message loop (Method: $loopMethod). Loop runs until main form is closed..." -ForegroundColor Cyan
				# Loop keeps going as long as MainForm is valid, visible, and not disposed.
				# Re-check $mainForm validity inside the loop to be safe.
				while ($mainForm -and $mainForm.Visible -and -not $mainForm.IsDisposed)
				{
					if ($useNativeLoop)
					{
						# --- Efficient Native P/Invoke Loop ---
						try
						{
							# Wait efficiently for window messages (QS_ALLINPUT) or a timeout (like 50ms).
							# $result - Return value from AsyncExecution (based on MsgWaitForMultipleObjectsEx).
							# WAIT_OBJECT_0 (0) means a message arrived. WAIT_TIMEOUT (0x102) means timeout.
							$result = [Custom.Native]::AsyncExecution(0, [IntPtr[]]@(), $false, 50, [Custom.Native]::QS_ALLINPUT) # Timeout 50ms
								
							# If a message arrived (result is not WAIT_TIMEOUT).
							if ($result -ne 0x102) # Compare with decimal value of WAIT_TIMEOUT.
							{
								# Handle all waiting messages currently in the queue.
								# $msg - Structure to hold message details (Custom.Native+MSG).
								$msg = New-Object Custom.Native+MSG
								# PeekMessage with PM_REMOVE gets and removes message. Loop while messages exist.
								while ([Custom.Native]::PeekMessage([ref]$msg, [IntPtr]::Zero, 0, 0, [Custom.Native]::PM_REMOVE))
								{
									# Turn virtual-key messages into character messages.
									$null = [Custom.Native]::TranslateMessage([ref]$msg)
									# Send the message to the right window handler.
									$null = [Custom.Native]::DispatchMessage([ref]$msg)
								}
							}
							# If it was a timeout ($result -eq 0x102), the loop just continues and waits again. Nothing needed.
						}
						catch
						{
							# Catch errors *inside* the native loop run (e.g., P/Invoke call failed).
							Write-Verbose "  Error during Native message loop iteration: $($_.Exception.Message). Attempting to fall back to DoEvents..." -ForegroundColor Red
							$useNativeLoop = $false # Switch to DoEvents for the next loops.
							$loopMethod = 'DoEvents'
							Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue # Make sure assembly loaded for backup.
							Start-Sleep -Milliseconds 50 # Add a small pause before maybe starting DoEvents loop next time.
						}
					}
					else # Use Backup DoEvents Loop
					{
						# --- Backup Application.DoEvents() Loop ---
						try
						{
							# Handle all messages currently in the queue. Less efficient as it does everything even if idle.
							[System.Windows.Forms.Application]::DoEvents()
							# Add a small pause to stop this backup loop from using 100% CPU if no messages.
							Start-Sleep -Milliseconds 20 # 20ms pause balances responsiveness and CPU use.
						}
						catch
						{
							# Catch errors during DoEvents() or Start-Sleep.
							Write-Verbose "  Error during DoEvents fallback loop iteration: $($_.Exception.Message). Loop may become unresponsive." -ForegroundColor Red
							# Maybe add longer pause or break if errors keep happening? For now, just log and continue loop.
							Start-Sleep -Milliseconds 100
						}
					}
				} # End while ($mainForm.Visible -and -not $mainForm.IsDisposed)
			#endregion Step: Run the Chosen Message Loop
		}
		catch
		{
			# Catch surprise errors setting up or during the main loop logic (outside the inner try/catch).
			Write-Verbose "  FATAL Error occurred within the UI message loop setup or main structure: $($_.Exception.Message)" -ForegroundColor Red
			# Try a very basic DoEvents loop as a last resort if the main loop structure failed.
			Write-Verbose "  Attempting basic DoEvents fallback loop after critical error..." -ForegroundColor Cyan
			try
			{
				if ($mainForm -and -not $mainForm.IsDisposed)
				{
					Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue 
				}
					
				while ($mainForm -and $mainForm.Visible -and -not $mainForm.IsDisposed)
				{
					[System.Windows.Forms.Application]::DoEvents()
					Start-Sleep -Milliseconds 50 # Slightly longer pause in emergency backup.
				}
			}
			catch
			{
				Write-Verbose "  Emergency fallback DoEvents loop also failed: $($_.Exception.Message)" -ForegroundColor Red
				[System.Windows.Forms.Application]::Run($mainForm)
			}
		}
		finally
		{
			# This block runs when the message loop stops, either normally (window closed) or due to an error caught above.
			# Log the final state of the main form. Use $? to check if $mainForm variable exists before using its properties.
			if ($mainForm -and ($mainForm -is [System.Windows.Forms.Form]))
			{
				Write-Verbose "UI message loop exited (Method: $loopMethod). Final Form State -> Visible: $($mainForm.Visible), Disposed: $($mainForm.IsDisposed)" -ForegroundColor Cyan
			}
			else
			{
				Write-Verbose "UI message loop exited (Method: $loopMethod). MainForm object appears invalid or null upon exit." -ForegroundColor Yellow
			}
			# Mark UI as not initialized anymore *after* the loop finishes.
			$global:DashboardConfig.State.UIInitialized = $false
		}
	}
#endregion Function: Start-MessageLoop
	
#region Function: Stop-Dashboard
	function Stop-Dashboard
	{
		<#
			.SYNOPSIS
				Performs comprehensive cleanup of application resources during shutdown.
			
			.DESCRIPTION
				This function is responsible for gracefully stopping and releasing all resources allocated by the application
				and its modules. It's designed to be called within the main script's `finally` block to ensure cleanup
				happens reliably, even if errors occurred during execution.
				
				Cleanup is performed in a specific order to minimize dependency issues and errors:
				1.  **Ftool Forms:** If the optional 'ftool.psm1' module was loaded and created forms (tracked in `$global:DashboardConfig.Resources.FtoolForms`), it attempts to close and dispose of them. It preferably calls a `Stop-FtoolForm` function (if defined by ftool.psm1) for module-specific cleanup before falling back to basic `.Close()` and `.Dispose()` calls.
				2.  **Timers:** Stops and disposes of all `System.Windows.Forms.Timer` objects registered in `$global:DashboardConfig.Resources.Timers`. Handles nested collections if necessary.
				3.  **Main UI Form:** Disposes of the main application window (`$global:DashboardConfig.UI.MainForm`) if it exists and isn't already disposed.
				4.  **Runspaces & Module Cleanup:**
				*   Disposes of known background runspaces (e.g., `$global:DashboardConfig.Resources.LaunchResources` if used by 'launch.psm1').
				*   Calls specific cleanup functions (e.g., `Stop-ClientLaunch`, `CleanupLogin`, `CleanupFtool`) if they exist (assumed to be defined by the respective modules). These functions are expected to handle module-specific resource release (e.g., closing handles, stopping threads).
				5.  **Application State:** Resets global state flags (`UIInitialized`, `LoginActive`, `LaunchActive`) to $false.
			
			.OUTPUTS
				[bool] Returns $true if all cleanup steps attempted completed without throwing *new* errors during the cleanup process itself.
				Returns $false if any cleanup step encountered an error (logged as a warning). The function attempts to continue subsequent cleanup steps even if one fails.
			
			.NOTES
				- Uses individual `try/catch` blocks around major cleanup sections (Ftool forms, Timers, Main Form, Runspaces/Modules) to ensure robustness. An error in one section should not prevent others from running.
				- Errors encountered *during cleanup* are logged using `Write-Verbose` and cause the function to return $false, but they do not typically halt the entire cleanup process.
				- Relies on modules potentially defining specific cleanup functions (`Cleanup<ModuleName>`) or resources (like `$global:DashboardConfig.Resources.LaunchResources`). These need to be implemented correctly within the modules themselves.
				- The order of operations is important (e.g., dispose child forms before main form, stop timers before disposing forms they might interact with).
			#>
		[CmdletBinding()]
		[OutputType([bool])]
		param()
			
		Write-Verbose "Stopping Dashboard and Cleaning Up Application Resources..." -ForegroundColor Cyan
		# $cleanupOverallSuccess - Flag ($true/$false) to track if *any* cleanup step had an error. Default to true.
		[bool]$cleanupOverallSuccess = $true

		if (Get-Command Restore-WindowStyles -ErrorAction SilentlyContinue) { Restore-WindowStyles } else {	Write-Verbose "  Restore-WindowStyles function not found. Restart Dashboard to enable hidden windows again." -ForegroundColor Yellow }

		#region Step 1: Clean Up launch recources
		Write-Verbose "Step 1: Cleaning up Launch..." -ForegroundColor Cyan
		if (Get-Command Stop-ClientLaunch -ErrorAction SilentlyContinue)
		{
			Stop-ClientLaunch
		}
		else
		{
			Write-Verbose "  Stop-ClientLaunch function not found. Skipping launch cleanup." -ForegroundColor Yellow
		}
		#endregion Step 1: Clean Up launch recources

		#region Step 2: Clean Up Disconnect Watcher
			Write-Verbose "Step 2: Cleaning up Disconnect Watcher..." -ForegroundColor Cyan
			if (Get-Command Stop-DisconnectWatcher -ErrorAction SilentlyContinue)
			{
				Stop-DisconnectWatcher
			}
			else
			{
				Write-Verbose "  Stop-DisconnectWatcher function not found. Skipping watcher cleanup." -ForegroundColor Yellow
			}
		#endregion Step 2: Clean Up Disconnect Watcher
			
		#region Step 3: Clean Up Ftool Forms (if Ftool module was loaded/used)
			Write-Verbose "Step 3: Cleaning up Ftool forms..." -ForegroundColor Cyan
			try
			{
				# Check if the FtoolForms list exists and has items. Use .PSObject.Properties to check safely.
				$ftoolForms = $global:DashboardConfig.Resources.FtoolForms
				if ($ftoolForms -and $ftoolForms.Count -gt 0)
				{
					# Check if the special cleanup function from ftool.psm1 exists.
					# $stopFtoolFormCmd - FunctionInfo object or null.
					$stopFtoolFormCmd = Get-Command -Name Stop-FtoolForm -ErrorAction SilentlyContinue
					# Make a copy of the keys to loop over, as we change the list during the loop.
					# $formKeys - List of text (form names).
					[string[]]$formKeys = @($ftoolForms.Keys)
					Write-Verbose "- Found $($formKeys.Count) Ftool form(s) registered. Attempting cleanup..." -ForegroundColor DarkGray
						
					foreach ($key in $formKeys)
					{
						# Get the form object safely.
						# $form - The Ftool form object, maybe null or disposed.
						$form = $ftoolForms[$key]
						# Check if it's a valid, non-disposed Windows Form.
						if ($form -and $form -is [System.Windows.Forms.Form] -and -not $form.IsDisposed)
						{
							$formText = try
							{
								$form.Text 
							}
							catch
							{
								'(Error getting text)' 
							} # Get form text safely.
							Write-Verbose "  - Stopping Ftool form '$formText' (Key: $key)." -ForegroundColor Cyan
							try
							{
								# Use the module's special cleanup function if available.
								if ($stopFtoolFormCmd)
								{
									Write-Verbose "  - Using Stop-FtoolForm function..." -ForegroundColor Cyan
									Stop-FtoolForm -Form $form -ErrorAction Stop # Call specific cleanup.
								}
								else # Basic backup cleanup.
								{
									Write-Verbose "  - Stop-FtoolForm command not found. Performing basic Close() for form '$formText'." -ForegroundColor Yellow
									# Ask the form to close nicely. This triggers FormClosing/FormClosed events.
									$form.Close()
									# Give a tiny moment for events to process, maybe not needed but can help sometimes.
									Start-Sleep -Milliseconds 20
								}
							}
							catch # Catch errors specifically from Stop-FtoolForm or Close().
							{
								Write-Verbose "  - Error during Stop-FtoolForm or Close() for form '$formText': $($_.Exception.Message)" -ForegroundColor Red
								# Mark overall cleanup as possibly failed, but continue to make sure Dispose() is called.
								$cleanupOverallSuccess = $false
							}
							finally # Always try to dispose the form directly, whether Close() worked or not.
							{
								Write-Verbose "  - Ensuring Dispose() is called for form '$formText'." -ForegroundColor Cyan
								try
								{
									if (-not $form.IsDisposed)
									{
										$form.Dispose() 
									}
								}
								catch
								{
									Write-Verbose "  - Error during final Dispose() for form '$formText': $($_.Exception.Message)" -ForegroundColor Red
									$cleanupOverallSuccess = $false
								}
							}
						}
						elseif ($form -and $form -is [System.Windows.Forms.Form] -and $form.IsDisposed)
						{
							Write-Verbose "  - Ftool form with Key '$key' was already disposed." -ForegroundColor DarkGray
						}
						else
						{
							Write-Verbose "  - Ftool form entry with Key '$key' is null or not a valid Form object." -ForegroundColor Yellow
							$cleanupOverallSuccess = $false
						}
							
						# Remove the entry from the tracking list after trying cleanup.
						$ftoolForms.Remove($key) | Out-Null
					} # End foreach form key
					Write-Verbose "- Finished Ftool form cleanup." -ForegroundColor Green
				}
				else
				{
					Write-Verbose "  No active Ftool forms found in configuration to clean up." -ForegroundColor DarkGray 
				}
			}
			catch # Catch errors in the Ftool cleanup part setup (e.g., accessing FtoolForms).
			{
				Write-Verbose "Error during Ftool form cleanup phase setup: $($_.Exception.Message)" -ForegroundColor Red
				$cleanupOverallSuccess = $false
			}
		#endregion Step 3: Clean Up Ftool Forms (if Ftool module was loaded/used)
			
		#region Step 4: Clean Up Application Timers
			Write-Verbose "Step 4: Cleaning up application timers..." -ForegroundColor Cyan
			try
			{
				# Check if the Timers list exists and has items.
				$timersCollection = $global:DashboardConfig.Resources.Timers
				if ($timersCollection -and $timersCollection.Count -gt 0)
				{
					Write-Verbose "- Found $($timersCollection.Count) timer registration(s). Stopping and disposing..." -ForegroundColor Cyan
					# Use a temporary list to gather all unique timer objects, handling possible nesting or duplicates.
					# $uniqueTimers - List of separate timer objects.
					[System.Collections.Generic.List[System.Windows.Forms.Timer]]$uniqueTimers = New-Object System.Collections.Generic.List[System.Windows.Forms.Timer]
						
					# Go through the registered items in the Timers list.
					# Items could be single timers, or nested lists (like hashtables) of timers.
					foreach ($item in $timersCollection.Values)
					{
						if ($item -is [System.Windows.Forms.Timer])
						{
							if (-not $uniqueTimers.Contains($item))
							{
								$uniqueTimers.Add($item) 
							}
						}
						elseif ($item -is [System.Collections.IDictionary])
						{
							# If item is another list, go through its values.
							$item.Values | Where-Object { $_ -is [System.Windows.Forms.Timer] } | ForEach-Object {
								if (-not $uniqueTimers.Contains($_))
								{
									$uniqueTimers.Add($_) 
								}
							}
						}
						# Add checks for other possible list types if used (like ArrayList).
					}
					Write-Verbose "- Found $($uniqueTimers.Count) unique System.Windows.Forms.Timer object(s) to dispose." -ForegroundColor Cyan
						
					# Go through the unique timer objects.
					foreach ($timer in $uniqueTimers)
					{
						try
						{
							# Check if timer object is valid and not already disposed.
							if ($timer -and -not $timer.IsDisposed) # Check IsDisposed before accessing properties like Enabled.
							{
								Write-Verbose "  - Disposing timer (Was Enabled: $($timer.Enabled))." -ForegroundColor Green
								# Stop the timer first if it's running now.
								if ($timer.Enabled)
								{
									$timer.Stop() 
								}
								# Dispose of the timer object to free up resources.
								$timer.Dispose()
							}
							else
							{
								Write-Verbose "  - Skipping already disposed or invalid timer object." -ForegroundColor DarkGray
							}
						}
						catch # Catch errors during individual timer Stop() or Dispose().
						{
							Write-Verbose "  - Error stopping or disposing a timer: $($_.Exception.Message)" -ForegroundColor Red
							$cleanupOverallSuccess = $false # Mark overall cleanup as possibly incomplete.
						}
					} # End foreach timer
						
					# Clear the main timer list in the global config after trying disposal.
					Write-Verbose "- Clearing global timer registration collection." -ForegroundColor Cyan
					$timersCollection.Clear()
					Write-Verbose "- Finished timer cleanup." -ForegroundColor Green
				}
				else
				{
					Write-Verbose "- No active timers found in configuration to clean up." -ForegroundColor DarkGray 
				}
			}
			catch # Catch errors in the timer cleanup part setup.
			{
				Write-Verbose "Error during timer cleanup phase setup: $($_.Exception.Message)" -ForegroundColor Red
				$cleanupOverallSuccess = $false
			}
		#endregion Step 4: Clean Up Application Timers
			
		#region Step 5: Clean Up Main UI Form
			Write-Verbose "Step 5: Cleaning up main UI form..." -ForegroundColor Cyan
			try
			{
				# Check if the main form object exists, is a Form, and is not already disposed.
				$mainForm = $global:DashboardConfig.UI.PSObject.Properties['MainForm']
				if ($mainForm -and $mainForm.Value -is [System.Windows.Forms.Form] -and -not $mainForm.Value.IsDisposed)
				{
					Write-Verbose "- Disposing MainForm object..." -ForegroundColor DarkGray
					# Dispose of the main form object. Should trigger its FormClosed event if not already closed.
					$mainForm.Value.Dispose()
					Write-Verbose "- [OK] MainForm disposed." -ForegroundColor Green
				}
				elseif ($mainForm -and $mainForm.Value -is [System.Windows.Forms.Form] -and $mainForm.Value.IsDisposed)
				{
					Write-Verbose "- MainForm was already disposed." -ForegroundColor Yellow
				}
				else
				{
					Write-Verbose "- MainForm object not found or invalid in configuration." -ForegroundColor Yellow
				}
			}
			catch # Catch errors during main form disposal.
			{
				Write-Verbose "Error disposing main UI form: $($_.Exception.Message)" -ForegroundColor Red
				$cleanupOverallSuccess = $false
			}
		#endregion Step 5: Clean Up Main UI Form
			
		#region Step 6: Reset Application State Flags
			Write-Verbose "Step 6: Resetting application state flags..." -ForegroundColor Cyan
			try
			{
				# Reset flags to show the app is no longer active/set up.
				$global:DashboardConfig.State.UIInitialized = $null
				$global:DashboardConfig.State.LoginActive = $null
				$global:DashboardConfig.State.LaunchActive = $null
				$global:DashboardConfig.State.ConfigInitialized = $null
				$global:DashboardConfig.LoadedModules = $null
				Write-Verbose "- State flags reset." -ForegroundColor Green
			}
			catch # Catch errors during state flag resetting.
			{
				Write-Verbose "  Error resetting global state flags: $($_.Exception.Message)" -ForegroundColor Red
				# Continue cleanup even with this small issue.
				$cleanupOverallSuccess = $false
			}
		#endregion Step 6: Reset Application State Flags
				
		#region Step 7: Final Log Message for Cleanup Status
			# Set log color based on overall cleanup success flag.
			# $finalColor - Text, 'Green' for success, 'Yellow' for partial success/warnings.
			[string]$finalColor = if ($cleanupOverallSuccess)
			{
				'Green' 
			}
			else
			{
				'Yellow' 
			}
			Write-Verbose "--- Dashboard Cleanup Finished. Overall Success: $cleanupOverallSuccess ---" -ForegroundColor $finalColor
		#endregion Step 7: Final Log Message
			
		# Return the overall success status of the cleanup actions.
		return $cleanupOverallSuccess
	}
#endregion Function: Stop-Dashboard

#endregion UI and Application Lifecycle Functions

#region Main Execution Block

# --- Initial Setup and Log Cleanup ---
try {
    # Ensure log directory exists
    $LogDir = Split-Path -Path $global:DashboardConfig.Paths.Verbose -Parent
    if (-not (Test-Path $LogDir -PathType Container)) { New-Item $LogDir -ItemType Directory -Force | Out-Null }
    # Clear old log contents (Verbose and VerboseLog)
    $global:DashboardConfig.Paths.Verbose | ForEach-Object {
        if (Test-Path $_ -PathType Leaf) { Clear-Content $_ -ErrorAction SilentlyContinue }
    }
    $VerbosePreference = 'Continue' # Ensure Write-Verbose messages are processed
} catch {
    Write-Warning "Initial log setup failed: $($_.Exception.Message)"
}

# --- Initialization Banner ---
$TS = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Verbose "=========================================" -ForegroundColor Cyan
Write-Verbose "=== Initializing Entropia Dashboard ===" -ForegroundColor Cyan
Write-Verbose "=== Timestamp: $TS ===" -ForegroundColor Cyan
Write-Verbose "=========================================" -ForegroundColor Cyan

# Main try/catch/finally block to manage the app life cycle.
try {
    # --- Splash Screen Initialization ---
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    $splashForm = New-Object System.Windows.Forms.Form -Property @{
        FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
        Size            = New-Object System.Drawing.Size(450, 120)
        Text            = 'Starting Entropia Dashboard...'
        TopMost         = $true
        Padding         = New-Object System.Windows.Forms.Padding(10)
		BackColor       = [System.Drawing.Color]::FromArgb(30, 30, 30)
        ForeColor       = [System.Drawing.Color]::White
    }
    
    $titleLabel = New-Object System.Windows.Forms.Label -Property @{
        Text      = 'Entropia Dashboard'
        Font      = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
        Dock      = 'Top'
        TextAlign = 'MiddleCenter'
    }

    $statusLabel = New-Object System.Windows.Forms.Label -Property @{
        Text      = 'Initializing...'
        Font      = New-Object System.Drawing.Font('Segoe UI', 9)
        Dock      = 'Top'
        TextAlign = 'MiddleCenter'
        Padding   = New-Object System.Windows.Forms.Padding(0, 5, 0, 5)
    }

    $progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{
        Style = 'Continuous'
        Dock  = 'Bottom'
    }
    
    $splashForm.Controls.AddRange(@($statusLabel, $progressBar, $titleLabel))
    $splashForm.Show()

    # Helper function to update the splash screen
    $updateSplash = {
        param([string]$Text, [int]$Percentage)
        if ($splashForm -and -not $splashForm.IsDisposed) {
            $splashForm.Invoke([Action]{
                $statusLabel.Text = "$Text..."
                $progressBar.Value = [Math]::Min(100, $Percentage)
            })
            $splashForm.Refresh()
        }
    }
    # --- End Splash Screen Initialization ---

    # 1. Environment Check
    Write-Verbose "--- Step 1: Ensuring Correct Execution Environment ---" -ForegroundColor Cyan
    & $updateSplash "Verifying execution environment" 10
    Request-Elevation # May restart and exit here
    if (-not (Initialize-ScriptEnvironment)) { throw 'Environment verification failed. Cannot continue.' }
    Write-Verbose "[OK] Environment verified." -ForegroundColor Green
        
    # 2. Base Configuration
    Write-Verbose "--- Step 2: Initializing Base Configuration (AppData Paths) ---" -ForegroundColor Cyan
    & $updateSplash "Initializing configuration" 20
    if (-not (Initialize-BaseConfig)) { throw 'Failed to initialize base application paths. Cannot continue.' }
    Write-Verbose "[OK] Base configuration paths initialized." -ForegroundColor Green
        
	# 2.5: Auto-update Check
	Write-Verbose "--- Step 2.5: Checking for script updates ---" -ForegroundColor Cyan
    & $updateSplash "Checking for updates" 30
	try {
			# Determine the path to the currently executing script or EXE.
			$executablePath = $null
			if (-not [string]::IsNullOrEmpty($PSCommandPath)) {
				$executablePath = $PSCommandPath
			} elseif (-not [string]::IsNullOrEmpty($MyInvocation.MyCommand.Path)) {
				$executablePath = $MyInvocation.MyCommand.Path
			} else {
				# Fallback for compiled EXEs where other variables might be empty
				try {
					$executablePath = (Get-Process -Id $PID -ErrorAction Stop).Path
				} catch {
					# This could fail in some very restricted environments.
				}
			}

			# If we couldn't find a path, we can't check the version.
			if ([string]::IsNullOrEmpty($executablePath)) {
				Write-Verbose "  Could not determine script/EXE path. Update check will be skipped." -ForegroundColor Yellow
			} else {
				Write-Verbose "  Checking version of: $executablePath" -ForegroundColor DarkGray
				# Read local content and extract Version from the .NOTES header
				$localVersion = '0.0'
				$localContent = Get-Content -Path $executablePath -Raw -ErrorAction SilentlyContinue
				if ($localContent) {
					$lm = [regex]::Match($localContent, '(?m)^\s*Version:\s*([0-9]+(?:\.[0-9]+)*)')
					if ($lm.Success) { $localVersion = $lm.Groups[1].Value }
				}
				# Fetch remote script raw content from GitHub to compare version
				$remoteUrl = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/version'
				$remoteContent = $null
				try {
					$remoteContent = (Invoke-WebRequest -Uri $remoteUrl -UseBasicParsing -ErrorAction Stop).Content
				} catch {
					Write-Verbose "  Update check failed to fetch remote script: $($_.Exception.Message)" -ForegroundColor Yellow
				}

				if ($remoteContent) {
					# 1. Parse Remote Version
					$rm = [regex]::Match($remoteContent, '(?m)^\s*Version:\s*([0-9]+(?:\.[0-9]+)*)')
					
					if ($rm.Success) {
						$remoteVersion = $rm.Groups[1].Value

						# 2. Parse Changelogs (Extracts everything after [Changelogs])
						$changelogDisplay = ""
						$cm = [regex]::Match($remoteContent, '(?is)\[Changelogs\]\s*(.*)$')
						if ($cm.Success) {
							$extractedLog = $cm.Groups[1].Value.Trim()
							if (-not [string]::IsNullOrWhiteSpace($extractedLog)) {
								# Limit length to prevent MessageBox overflow
								if ($extractedLog.Length -gt 2000) { 
									$extractedLog = $extractedLog.Substring(0, 2000) + "...(truncated)" 
								}
								$changelogDisplay = "`n`n=== What's New ===`n$extractedLog"
							}
						}

						try {
							if ([version]$localVersion -lt [version]$remoteVersion) {
								Write-Verbose "  Update available: local $localVersion < remote $remoteVersion" -ForegroundColor Yellow
								
								# Prompt user: Yes = Download & Restart, No = Continue
								$promptMsg = "A newer version ($remoteVersion) of Entropia Dashboard is available.`nYou have $localVersion.$changelogDisplay`n`nPress 'Yes' to download the update and restart automatically.`nPress 'No' to continue without updating."
								$caption = "Entropia Dashboard - Update Available"
								
								$resp = [System.Windows.Forms.MessageBox]::Show($promptMsg, $caption, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
								
								if ($resp -eq [System.Windows.Forms.DialogResult]::Yes) {
									Write-Verbose "  User chose to update. Starting download..." -ForegroundColor Cyan
									
									# --- Auto-Update Logic ---
									
									# 1. Define Paths
									# Note: Using 'raw/main' to get the file directly
									$updateUrl = 'https://github.com/Immortal-Divine/Entropia_Dashboard/raw/main/Entropia%20Dashboard.exe'
									$tempExePath = "$executablePath.new"
									$batchPath = Join-Path $env:TEMP ("dashboard_updater_" + [Guid]::NewGuid().ToString() + ".bat")

									# 2. Download the new file
									try {
                                        & $updateSplash "Downloading update" 35
										# Using .NET WebClient for synchronous download with simple error handling
										$webClient = New-Object System.Net.WebClient
										$webClient.DownloadFile($updateUrl, $tempExePath)
										Write-Verbose "  Download complete: $tempExePath" -ForegroundColor Green
									} catch {
										$errMsg = "Failed to download update. Please download manually from GitHub.`nError: $($_.Exception.Message)"
										Write-Verbose "  $errMsg" -ForegroundColor Red
										[System.Windows.Forms.MessageBox]::Show($errMsg, "Update Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
										# Fallback to opening browser if download fails
										Start-Process 'https://github.com/Immortal-Divine/Entropia_Dashboard'
										return 
									}

									# 3. Create a temporary batch file to handle the swap
									# We cannot overwrite the running EXE, so we run a batch file that waits for us to exit, then moves the file.
									$batchContent = @"
@echo off
title Entropia Dashboard Updater
echo Waiting for application to close...
timeout /t 2 /nobreak >nul

:RetryMove
echo Updating application file...
move /y "$tempExePath" "$executablePath" >nul 2>&1
if errorlevel 1 (
    echo File is locked. Retrying in 1 second...
    timeout /t 1 /nobreak >nul
    goto RetryMove
)

echo Update successful. Restarting...
start "" "$executablePath"
del "%~f0"
"@
									Set-Content -Path $batchPath -Value $batchContent -Encoding ASCII

									# 4. Execute the batch file hidden and Exit
									Write-Verbose "  Starting updater script and exiting..." -ForegroundColor Cyan
                                    & $updateSplash "Restarting to apply update" 40
									Start-Process -FilePath $batchPath -WindowStyle Hidden
									
									# Return from main try block to trigger 'finally' and exit
									return
								} else {
									Write-Verbose "  User chose to continue without updating." -ForegroundColor Cyan
								}
							} else {
								Write-Verbose "  Local script version ($localVersion) is up-to-date (remote $remoteVersion)." -ForegroundColor Green
							}
						} catch {
							Write-Verbose "  Version compare failed: $($_.Exception.Message)" -ForegroundColor Yellow
						}
					} else {
						Write-Verbose "  Remote script did not contain a version header." -ForegroundColor Yellow
					}
				}
			}
		} catch {
			Write-Verbose "  Unexpected error during update check: $($_.Exception.Message)" -ForegroundColor Yellow
	}

    # 3. Load Modules
    Write-Verbose "--- Step 3: Loading Dashboard Modules ---" -ForegroundColor Cyan
    & $updateSplash "Loading core modules" 50
    $importResult = Import-DashboardModules
    if (-not $importResult.Status) { throw 'Critical module loading failed. Cannot continue.' }
    Write-Verbose "[OK] Core modules loaded successfully." -ForegroundColor Green
        
    # 4. Load INI Config
    Write-Verbose "--- Step 4: Loading INI Configuration ---" -ForegroundColor Cyan
    & $updateSplash "Loading saved settings" 70
    if (Get-Command Initialize-IniConfig -ErrorAction SilentlyContinue) {
        try {
            if (-not (Initialize-IniConfig -ErrorAction Stop)) { Write-Verbose "Initialize-IniConfig failed. Defaults used." -ForegroundColor Yellow }
            else { Write-Verbose "[OK] INI configuration loaded successfully." -ForegroundColor Green }
        } catch { Write-Verbose "Error during Initialize-IniConfig: $($_.Exception.Message). Defaults used." -ForegroundColor Yellow }
    } else { Write-Verbose "Initialize-IniConfig not found. Skipping INI load." -ForegroundColor Yellow }
        
    # 5. Start UI
    Write-Verbose "--- Step 5: Starting Dashboard UI ---" -ForegroundColor Cyan
    & $updateSplash "Building main user interface" 90
    if (-not (Start-Dashboard)) { throw 'Start-Dashboard returned failure. Cannot continue.' }
    Write-Verbose "[OK] Dashboard UI started." -ForegroundColor Green
    if (Get-Command Start-DataGridUpdateTimer -ErrorAction SilentlyContinue) {
        Start-DataGridUpdateTimer; Write-Verbose "[OK] DataGrid update timer started." -ForegroundColor Green
    } else { Write-Verbose "[WARN] Start-DataGridUpdateTimer not found." -ForegroundColor Yellow }

    & $updateSplash "Finalizing" 100
    Start-Sleep -Milliseconds 250 # Give a moment for the user to see "Finalizing"
    if ($splashForm -and -not $splashForm.IsDisposed) { $splashForm.Close() }

    # 6. Run UI Message Loop (Foreground focus logic simplified)
    Write-Verbose "--- Step 6: Running UI Message Loop ---" -ForegroundColor Cyan
    $handle = $global:DashboardConfig.UI.MainForm.Handle
    if ($handle -ne [IntPtr]::Zero) {
        Start-Sleep -Milliseconds 200
        if ([Custom.Native]::IsWindowMinimized($handle)) { [Custom.Native]::ShowWindow($handle, 9) | Out-Null }
        if (-not ([Custom.Native]::SetForegroundWindow($handle))) {
            Write-Verbose "SetForegroundWindow failed. Trying Alt-key workaround..." -ForegroundColor Yellow
            try { 
                [Custom.Native]::keybd_event(0x12, 0, 1, 0); Start-Sleep -Milliseconds 50
                [Custom.Native]::keybd_event(0x12, 0, 3, 0); Start-Sleep -Milliseconds 100
                if (-not ([Custom.Native]::SetForegroundWindow($handle))) {
                    Write-Verbose "Alt workaround failed. Using Activate()." -ForegroundColor Yellow
                    $global:DashboardConfig.UI.MainForm.Activate()
                }
            } catch { Write-Verbose "Error in Alt-key simulation: $_" -ForegroundColor Red }
        }
    }

    # 7. Check active Reconnect Settings
	Write-Verbose "--- Step 7: Updating Reconnect Supervisor ---" -ForegroundColor Cyan
	if (Get-Command Start-DisconnectWatcher -ErrorAction SilentlyContinue) {
    Start-DisconnectWatcher
	} else {
		Write-Verbose "[WARN] Start-DisconnectWatcher not found." -ForegroundColor Yellow
	}
    Start-MessageLoop

    Write-Verbose "UI Message loop finished. Proceeding to final cleanup..." -ForegroundColor Green
}
catch {
    # --- Main Catch Block ---
    if ($splashForm -and -not $splashForm.IsDisposed) { $splashForm.Dispose() }
    $errorMessage = "`nFATAL UNHANDLED ERROR: $($_.Exception.Message)"
    Write-Verbose $errorMessage -ForegroundColor Red
    try { Show-ErrorDialog ($errorMessage + "`n`nStack Trace:`n" + $($_.ScriptStackTrace)) }
    catch { Write-Verbose "Failed to show final error dialog. The critical error was: $errorMessage" -ForegroundColor Red }
}
finally {
    # 7. Final Application Cleanup
    Write-Verbose "--- Step 7: Entering Final Application Cleanup ---" -ForegroundColor Cyan
    if ($splashForm -and -not $splashForm.IsDisposed) { $splashForm.Dispose() }
    if (Get-Command Remove-AllHotkeys -ErrorAction SilentlyContinue) { Remove-AllHotkeys; Write-Verbose "Hotkeys unregistered." -ForegroundColor Cyan }
    if (Get-Command Stop-Dashboard -ErrorAction SilentlyContinue) {
        $cleanupStatus = Stop-Dashboard; Write-Verbose "[OK] Stop-Dashboard completed (Success: $cleanupStatus)." -ForegroundColor Green
    } else {
        Write-Verbose "Stop-Dashboard not found. Attempting basic MainForm dispose..." -ForegroundColor Yellow
        try {
            $mainForm = $global:DashboardConfig.UI.MainForm
            if ($mainForm -is [System.Windows.Forms.Form] -and -not $mainForm.IsDisposed) { $mainForm.Dispose() }
        } catch { Write-Verbose "Fallback MainForm dispose failed: $($_.Exception.Message)" -ForegroundColor Red }
    }
    
    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')) {
        try { [System.Windows.Forms.Application]::ExitThread() }
        catch { Write-Verbose "Error calling ExitThread(): $($_.Exception.Message)" -ForegroundColor Red }
    }
    
    Write-Verbose "=========================================" -ForegroundColor Cyan
    Write-Verbose "=== Entropia Dashboard Exited ===" -ForegroundColor Cyan
    Write-Verbose "=========================================" -ForegroundColor Cyan

}

#endregion Main Execution Block