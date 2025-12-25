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

if ($args -contains '-Verbose')
{
	$VerbosePreference = 'Continue'
	Write-Verbose '-Verbose argument detected, enabling verbose preference.'
}

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop

function Write-Verbose
{
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true, Position = 0)][string]$Message,
		[string]$ForegroundColor = 'DarkGray'
	)
	if ($VerbosePreference -eq 'Continue')
	{
		$dateStr = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
		$callStack = Get-PSCallStack
		$caller = if ($callStack.Count -gt 1) { $callStack[1] } else { $callStack[0] }
		$callerName = if ($caller.Command) { $caller.Command } else { 'Script' }
        
		$bracketedCaller = "[$callerName]"
		$paddedCaller = $bracketedCaller.PadRight(35)
		$prefix = " | $dateStr - $paddedCaller - "
		$indentation = ' ' * $prefix.Length + '| '
        
		$consoleWidth = [Math]::Max($Host.UI.RawUI.WindowSize.Width, 800)

		
		$lines = $Message -split "`r`n"
		$formattedLines = @()
		$availableWidth = $consoleWidth - $prefix.Length - 5

		foreach ($line in $lines)
		{
			if ($line.Length -le $availableWidth)
			{
				$formattedLines += $line
				continue
			}
			$currentLine = ''
			$line.Split(' ') | ForEach-Object {
				if (($currentLine.Length + $_.Length + 1) -le $availableWidth)
				{
					$currentLine += (if ($currentLine) { " $_" } else { $_ })
				}
				else
				{
					$formattedLines += $currentLine
					$currentLine = $_
				}
			}
			if ($currentLine) { $formattedLines += $currentLine }
		}
        
		
		$formattedMessage = ''
		if ($formattedLines.Count -gt 0)
		{
			$formattedMessage = $formattedLines[0]
			for ($i = 1; $i -lt $formattedLines.Count; $i++)
			{
				$formattedMessage += "`r`n$indentation$($formattedLines[$i])"
			}
		}

		$logPath = $global:DashboardConfig.Paths.Verbose 
 
		$logLine = "$prefix$formattedMessage"
         
		try
		{
			Add-Content -Path $logPath -Value $logLine -ErrorAction SilentlyContinue
		}
		catch {}
        
        
		
		$color = switch ($ForegroundColor.ToLower())
		{
			'red' {[ConsoleColor]::Red}
			'yellow' {[ConsoleColor]::Yellow}
			'green' {[ConsoleColor]::Green}
			'cyan' {[ConsoleColor]::Cyan}
			default {[ConsoleColor]::DarkGray}
		}

		
		$orig = $host.UI.RawUI.ForegroundColor
		try
		{
			$host.UI.RawUI.ForegroundColor = $color
			[Console]::Error.WriteLine("$prefix$formattedMessage")
		}
		finally
		{
			$host.UI.RawUI.ForegroundColor = $orig
		}
        
		
		$wrappedCmdlet = $ExecutionContext.InvokeCommand.GetCommand('Microsoft.PowerShell.Utility\Write-Verbose', [System.Management.Automation.CommandTypes]::Cmdlet)
		$script = { & $wrappedCmdlet "$prefix$formattedMessage" }
		$pipe = $script.GetSteppablePipeline()
		$pipe.Begin($true)
		$pipe.End()
	}
}

#endregion Custom Write-Verbose

#region Global Configuration

$global:DashboardConfig = @{
	Paths            = @{
		Source   = Join-Path $env:USERPROFILE 'Entropia_Dashboard\.main'
		App      = Join-Path $env:APPDATA 'Entropia_Dashboard\'
		Profiles = Join-Path $env:APPDATA 'Entropia_Dashboard\Profiles'
		Modules  = Join-Path $env:APPDATA 'Entropia_Dashboard\modules'
		Icon     = Join-Path $env:APPDATA 'Entropia_Dashboard\modules\icon.ico'
		FtoolDLL = Join-Path $env:APPDATA 'Entropia_Dashboard\modules\ftool.dll'
		Ini      = Join-Path $env:APPDATA 'Entropia_Dashboard\config.ini'
		Verbose  = Join-Path $env:APPDATA 'Entropia_Dashboard\log\verbose1.log'
	}
	State            = @{
		ConfigInitialized    = $false
		UIInitialized        = $false
		LoginActive          = $false
		LaunchActive         = $false
		PreviousLaunchState  = $false
		PreviousLoginState   = $false
		IsRunningAsExe       = $false
		IsDragging           = $false
		DisconnectActive     = $false
		LastNotifyHwnd       = $false
		ReconnectActive      = $false
		LoginNotificationMap = @{}
	}
	Resources        = @{
		Timers              = [ordered]@{}
		FtoolForms          = [ordered]@{}
		LastEventTimes      = @{}
		ExtensionData       = @{}
		ExtensionTracking   = @{}
		LoadedModuleContent = @{}
		LaunchResources     = @{}
		DragSourceGrid      = $null
	}
	UI               = @{
		Login = @{}
	}
	DefaultConfig    = [ordered]@{
		'LauncherPath'      = [ordered]@{ 'LauncherPath' = 'Select Launcher.exe' }
		'ProcessName'       = [ordered]@{ 'ProcessName' = 'neuz' }
		'MaxClients'        = [ordered]@{ 'MaxClients' = '1' }
		'Login'             = [ordered]@{ 'NeverRestartingCollectorLogin' = '0' }
		'Ftool'             = [ordered]@{}
		'Options'           = [ordered]@{}
		'Paths'             = [ordered]@{ 'JunctionTarget' = (Join-Path $env:APPDATA 'Entropia_Dashboard\Profiles') }
		'ReconnectProfiles' = [ordered]@{} 
		'LoginConfig'       = [ordered]@{ 'PostLoginDelay' = '0'; 'Server1Coords' = '0,0'; 'Server2Coords' = '0,0'; 'Channel1Coords' = '0,0'; 'Channel2Coords' = '0,0'; 'FirstNickCoords' = '0,0'; 'ScrollDownCoords' = '0,0'; 'Char1Coords' = '0,0'; 'Char2Coords' = '0,0'; 'Char3Coords' = '0,0'; 'CollectorStartCoords' = '0,0'; 'DisconnectOKCoords' = '0,0'; 'LoginDetailsOKCoords' = '0,0' }
	}
	Config           = [ordered]@{}
	ConfigWriteTimer = @{}
	LoadedModules    = @{}
}

#region Step: Define Module Metadata
$global:DashboardConfig.Modules = @{
	'ftool.dll'            = @{ 
		Priority = 'Critical'; Order = 1; Dependencies = @() 
		Base64Content = '
TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA6AAAAA4fug4AtAnNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAADylR4YtvRwS7b0cEu29HBLkTILS7T0cEuRMg1Lt/RwS5EyHku09HBLkTIdS7v0cEt1+y1LtfRwS7b0cUuX9HBLkTICS7f0cEuRMgpLt/RwS5EyCEu39HBLUmljaLb0cEsAAAAAAAAAAFBFAABMAQUApBpbSAAAAAAAAAAA4AACIQsBCAAACgAAAA4AAAAAAAClFAAAABAAAAAgAAAAAAAQABAAAAACAAAEAAAAAAAAAAQAAAAAAAAAAGAAAAAEAAB7lQAAAgAAAAAAEAAAEAAAAAAQAAAQAAAAAAAAEAAAAPAlAADPAAAAnCIAADwAAAAAQAAArAEAAAAAAAAAAAAAAAAAAAAAAAAAUAAAbAEAAKAgAAAcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQCEAAEAAAAAAAAAAAAAAAAAgAACMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALnRleHQAAABoCQAAABAAAAAKAAAABAAAAAAAAAAAAAAAAAAAIAAAYC5yZGF0YQAAvwYAAAAgAAAACAAAAA4AAAAAAAAAAAAAAAAAAEAAAEAuZGF0YQAAAHwDAAAAMAAAAAIAAAAWAAAAAAAAAAAAAAAAAABAAADALnJzcmMAAACsAQAAAEAAAAACAAAAGAAAAAAAAAAAAAAAAAAAQAAAQC5yZWxvYwAApAEAAABQAAAAAgAAABoAAAAAAAAAAAAAAAAAAEAAAEIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIv/VYvs/yVcMwAQzMzMzMyL/1WL7P8lVDMAEMzMzMzMi/9Vi+z/JUwzABDMzMzMzIv/VYvs/yVYMwAQzMzMzMyL/1WL7P8lSDMAEMzMzMzMuPcRAAD/JWAzABDMzMzMzItEJAiD+ANWD4fsAAAA/ySFYBEAEGjIIAAQ/xUAIAAQgz1cMwAQAIs1CCAAEKNQMwAQdRVo4CAAEFD/1oPABaNcMwAQoVAzABCDPVQzABAAdRVo8CAAEFD/1oPABaNUMwAQoVAzABCDPUwzABAAdRVoACEAEFD/1oPABaNMMwAQoVAzABCDPVgzABAAdRVoECEAEFD/1oPABaNYMwAQoVAzABCDPUgzABAAdRVoICEAEFD/1oPABaNIMwAQoVAzABCDPWAzABAAdTBoMCEAEFD/1oPABaNgMwAQsAFewgwAoVAzABCFwHQRUP8VBCAAEMcFUDMAEAAAAACwAV7CDABAEQAQdRAAEHUQABBAEQAQOw0AMAAQdQLzw+lHAwAAVmiAAAAA/xV8IAAQi/BW/xWAIAAQhfZZWaN0MwAQo3AzABB1BTPAQF7DgyYA6NwEAABosRYAEOjABAAAxwQkyhUAEOi0BAAAWTPAXsOLRCQIVTPtO8V1DjktEDAAEH46/w0QMAAQg/gBiw1gIAAQiwlTVleJDWQzABAPhdQAAABkoRgAAACLcASLHTAgABCJbCQYv2wzABDrFjPA6WsBAAA7xnQWaOgDAAD/FTQgABBVVlf/0zvFdejrCMdEJBgBAAAAoWgzABCFwGoCXnQJah/o0wUAAOs8aJwgABBolCAAEMcFaDMAEAEAAADosgUAAIXAWVl0BzPA6QsBAABokCAAEGiMIAAQ6JAFAABZiTVoMwAQOWwkHFl1CFVX/xU4IAAQOS14MwAQdB5oeDMAEOisBAAAhcBZdA//dCQcVv90JBz/FXgzABD/BRAwABDpsgAAADvFD4WqAAAAizUwIAAQv2wzABDrC2joAwAA/xU0IAAQVWoBV//WhcB166FoMwAQg/gCdApqH+gaBQAAWet0/zV0MwAQix1wIAAQ/9OL6IXtWXRM/zVwMwAQ/9NZi/DrIIM+AHQbiwaJRCQY/xV0IAAQOUQkGHQJ/3QkGP/TWf/Qg+4EO/Vz2VX/FXggABBZ/xV0IAAQo3AzABCjdDMAEGoAV8cFaDMAEAAAAAD/FTggABAzwEBfXltdwgwAahBoOCIAEOiZBAAAi/mL8otdCDPAQIlF5DPJiU38iTUIMAAQiUX8O/F1EDkNEDAAEHUIiU3k6bcAAAA78HQFg/4CdS6hvCAAEDvBdAhXVlP/0IlF5IN95AAPhJMAAABXVlPo1v3//4lF5IXAD4SAAAAAV1ZT6Ff8//+JReSD/gF1JIXAdSBXUFPoQ/z//1dqAFPopv3//6G8IAAQhcB0BldqAFP/0IX2dAWD/gN1Q1dWU+iG/f//hcB1AyFF5IN95AB0LqG8IAAQhcB0JVdWU//QiUXk6xuLReyLCIsJiU3gUFHotwMAAFlZw4tl6INl5ACDZfwAx0X8/v///+gJAAAAi0Xk6OADAADDxwUIMAAQ/////8ODfCQIAXUF6P8DAAD/dCQEi0wkEItUJAzozf7//1nCDABVi+yB7CgDAACjIDEAEIkNHDEAEIkVGDEAEIkdFDEAEIk1EDEAEIk9DDEAEGaMFTgxABBmjA0sMQAQZowdCDEAEGaMBQQxABBmjCUAMQAQZowt/DAAEJyPBTAxABCLRQCjJDEAEItFBKMoMQAQjUUIozQxABCLheD8///HBXAwABABAAEAoSgxABCjJDAAEMcFGDAAEAkEAMDHBRwwABABAAAAoQAwABCJhdj8//+hBDAAEImF3Pz///8VHCAAEKNoMAAQagHoswMAAFlqAP8VICAAEGjAIAAQ/xUkIAAQgz1oMAAQAHUIagHojwMAAFloCQQAwP8VKCAAEFD/FSwgABDJw2hAMwAQ6HYDAABZw2oUaGAiABDoUgIAAP81dDMAEIs1cCAAEP/WWYlF5IP4/3UM/3UI/xVMIAAQWetnagjoUAMAAFmDZfwA/zV0MwAQ/9aJReT/NXAzABD/1llZiUXgjUXgUI1F5FD/dQiLNYAgABD/1llQ6BMDAACJRdz/deT/1qN0MwAQ/3Xg/9aDxBSjcDMAEMdF/P7////oCQAAAItF3OgIAgAAw2oI6NcCAABZw/90JAToUv////fYG8D32FlIw1ZXuCgiABC/KCIAEDvHi/BzD4sGhcB0Av/Qg8YEO/dy8V9ew1ZXuDAiABC/MCIAEDvHi/BzD4sGhcB0Av/Qg8YEO/dy8V9ew8zMzMzMzMzMzMzMi0wkBGaBOU1adAMzwMOLQTwDwYE4UEUAAHXwM8lmgXgYCwEPlMGLwcPMzMzMzMzMi0QkBItIPAPID7dBFFNWD7dxBjPShfZXjUQIGHYei3wkFItIDDv5cgmLWAgD2Tv7cgyDwgGDwCg71nLmM8BfXlvDzMzMzMzMzMzMzMzMzMxVi+xq/miAIgAQaI0YABBkoQAAAABQg+wIU1ZXoQAwABAxRfgzxVCNRfBkowAAAACJZejHRfwAAAAAaAAAABDoPP///4PEBIXAdFWLRQgtAAAAEFBoAAAAEOhS////g8QIhcB0O4tAJMHoH/fQg+ABx0X8/v///4tN8GSJDQAAAABZX15bi+Vdw4tF7IsIiwEz0j0FAADAD5TCi8LDi2Xox0X8/v///zPAi03wZIkNAAAAAFlfXluL5V3DzP8lbCAAEP8laCAAEP8lZCAAEP8lXCAAEGiNGAAQZP81AAAAAItEJBCJbCQQjWwkECvgU1ZXoQAwABAxRfwzxVCJZej/dfiLRfzHRfz+////iUX4jUXwZKMAAAAAw4tN8GSJDQAAAABZX19eW4vlXVHD/3QkEP90JBD/dCQQ/3QkEGhwEQAQaAAwABDotgAAAIPEGMNVi+yD7BChADAAEINl+ACDZfwAU1e/TuZAuzvHuwAA//90DYXDdAn30KMEMAAQ62BWjUX4UP8VPCAAEIt1/DN1+P8VDCAAEDPw/xUQIAAQM/D/FRQgABAz8I1F8FD/FRggABCLRfQzRfAz8Dv3dQe+T+ZAu+sLhfN1B4vGweAQC/CJNQAwABD31ok1BDAAEF5fW8nD/yVYIAAQ/yVUIAAQ/yVEIAAQ/yWEIAAQ/yVIIAAQ/yVQIAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkIwAAdCMAAIIjAACyJQAAnCUAAIwlAAByJQAAXiUAAEAlAAAkJQAAECUAAPwkAADeJAAA1iQAAMAkAADIJQAAAAAAAHAkAACIJAAAkCQAAKYkAABMJAAANiQAACQkAAAUJAAABiQAAPgjAADsIwAA2iMAAMojAADCIwAAtCMAAKIjAAB6JAAAAAAAAAAAAAAAAAAAAAAAAH8RABAAAAAAAAAAAKQaW0gAAAAAAgAAAIkAAACIIQAAiA8AAAAAAAAYMAAQcDAAEHUAcwBlAHIAMwAyAC4AZABsAGwAAAAAAFBvc3RNZXNzYWdlQQAAAABQb3N0TWVzc2FnZVcAAAAAU2VuZE1lc3NhZ2VBAAAAAFNlbmRNZXNzYWdlVwAAAABTZXRDdXJzb3JQb3MAAAAAU2V0QWN0aXZlV2luZG93AEgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwABAgIgAQAQAAAFJTRFNTUO3S8L97Q4yN0/DLuR2RBwAAAGM6XERvY3VtZW50cyBhbmQgU2V0dGluZ3NcQWxkZVxNeSBEb2N1bWVudHNcVmlzdWFsIFN0dWRpbyAyMDA1XFByb2plY3RzXEZseUZGIEFwcGxpY2F0aW9uc1xyZWxlYXNlXEZ1bmN0aW9ucy5wZGIAAAAAAAAAAAAAAAAAAAAAjRgAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/v///wAAAADQ////AAAAAP7///8AAAAAmhQAEAAAAABmFAAQehQAEP7///8AAAAAzP///wAAAAD+////AAAAAHIWABAAAAAA/v///wAAAADY////AAAAAP7////pFwAQ/RcAENgiAAAAAAAAAAAAAJQjAAAAIAAAHCMAAAAAAAAAAAAAmiQAAEQgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGQjAAB0IwAAgiMAALIlAACcJQAAjCUAAHIlAABeJQAAQCUAACQlAAAQJQAA/CQAAN4kAADWJAAAwCQAAMglAAAAAAAAcCQAAIgkAACQJAAApiQAAEwkAAA2JAAAJCQAABQkAAAGJAAA+CMAAOwjAADaIwAAyiMAAMIjAAC0IwAAoiMAAHokAAAAAAAAVQJMb2FkTGlicmFyeVcAAPgARnJlZUxpYnJhcnkAoAFHZXRQcm9jQWRkcmVzcwAAS0VSTkVMMzIuZGxsAAByAV9lbmNvZGVfcG9pbnRlcgCTAl9tYWxsb2NfY3J0APQEZnJlZQAAcwFfZW5jb2RlZF9udWxsAGgBX2RlY29kZV9wb2ludGVyABACX2luaXR0ZXJtABECX2luaXR0ZXJtX2UAHQFfYW1zZ19leGl0AAATAV9hZGp1c3RfZmRpdgAAbQBfX0NwcFhjcHRGaWx0ZXIAUwFfY3J0X2RlYnVnZ2VyX2hvb2sAAI8AX19jbGVhbl90eXBlX2luZm9fbmFtZXNfaW50ZXJuYWwAAPMDX3VubG9jawCZAF9fZGxsb25leGl0AIICX2xvY2sAKANfb25leGl0AE1TVkNSODAuZGxsAHsBX2V4Y2VwdF9oYW5kbGVyNF9jb21tb24AKQJJbnRlcmxvY2tlZEV4Y2hhbmdlAFYDU2xlZXAAJgJJbnRlcmxvY2tlZENvbXBhcmVFeGNoYW5nZQAAXgNUZXJtaW5hdGVQcm9jZXNzAABCAUdldEN1cnJlbnRQcm9jZXNzAG4DVW5oYW5kbGVkRXhjZXB0aW9uRmlsdGVyAABKA1NldFVuaGFuZGxlZEV4Y2VwdGlvbkZpbHRlcgA5AklzRGVidWdnZXJQcmVzZW50AKMCUXVlcnlQZXJmb3JtYW5jZUNvdW50ZXIA3wFHZXRUaWNrQ291bnQAAEYBR2V0Q3VycmVudFRocmVhZElkAABDAUdldEN1cnJlbnRQcm9jZXNzSWQAygFHZXRTeXN0ZW1UaW1lQXNGaWxlVGltZQAAAAAAAAAAAAAAAAAAAAAAAACkGltIAAAAAFQmAAABAAAABgAAAAYAAAAYJgAAMCYAAEgmAAAAEAAAEBAAACAQAAAwEAAAUBAAAEAQAABiJgAAcSYAAIAmAACPJgAAniYAALAmAAAAAAEAAgADAAQABQBGdW5jdGlvbnMuZGxsAGZuUG9zdE1lc3NhZ2VBAGZuUG9zdE1lc3NhZ2VXAGZuU2VuZE1lc3NhZ2VBAGZuU2VuZE1lc3NhZ2VXAGZuU2V0QWN0aXZlV2luZG93AGZuU2V0Q3Vyc29yUG9zAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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
				'
	}
	'icon.ico'             = @{ 
		Priority = 'Critical'; Order = 2; Dependencies = @() 
		Base64Content = '
AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAACMuAAAjLgAAAAAAAAAAAAAAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AQEB/wEBAf8BAQH/AgIC/wICAv8CAgL/BAQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wUEAv8FBAL/BQQC/wQEAv8CAgL/AgIC/wICAv8DAwP/AwMD/wMDA/8HBgP/BwYD/wcGA/8HBgP/BwYD/wcGA/8HBgP/BwYD/wcGA/8GBgP/BQUD/wUEA/8NDgP/DQ4D/wUEA/8FBQP/BgYD/wcGA/8HBgP/BwYD/wcGA/8HBgP/BwYD/wcGA/8HBgP/BgYD/wMDA/8DAwP/AwMD/wQEBP8EBAT/BAQE/wMDBP8DAwT/AwME/wMDBP8DAwT/AwME/wQFBP8CAwT/BQQE/xIPA/8jHQP/LiUD/zo0A/86NAP/LiUD/yMdA/8SDwP/BQQE/wIDBP8EBQT/AwME/wMDBP8DAwT/AwME/wMDBP8DAwT/BAQE/wQEBP8EBAT/BAQE/wQEBP8EBAT/BAQE/wQEBP8EBAT/BAQE/wQEBP8DAwT/EBME/ycmBP84LQP/OC0E/ykhBP8cFwT/FxME/xcTBP8cFwT/KSEE/zgtBP84LQP/JyYE/xATBP8DAwT/BAQE/wQEBP8EBAT/BAQE/wQEBP8EBAT/BAQE/wQEBP8FBQX/BQUF/wUFBf8FBQX/BQUF/wUFBf8FBQX/BAQF/w0MBf83LgT/OjIE/xYSBf8JCgX/Cw0F/w4QBf8PEQX/DhEF/w4QBf8LDQX/CQoF/xYSBf86MgT/OC4E/w0MBf8EBAX/BQUF/wUFBf8FBQX/BQUF/wUFBf8FBQX/BQUF/wYGBv8GBgb/BgYG/wYGBv8GBgb/BgYG/wUFBv8SEAb/PjIF/yMdBf8JCQb/DxIG/xETBv8MDgb/CQoG/wgJBv8ICQb/CQoG/wwOBv8REwb/DxIG/wkJBv8jHQX/PjIF/xIQBv8FBQb/BgYG/wYGBv8GBgb/BgYG/wYGBv8GBgb/BwcH/wcHB/8HBwf/BwcH/wcHB/8GBgf/Dw0H/z4zBf8dGAb/CwwH/xIVBv8MDQf/BwcH/wYGB/8HBwf/BwcH/wcHB/8HBwf/BgYH/wcHB/8MDQf/EhUG/wsMB/8dGAb/PjMF/w8NB/8GBgf/BwcH/wcHB/8HBwf/BwcH/wcHB/8ICAj/CAgI/wgICP8ICAj/CQkI/xMWB/85MAb/JR4H/wsNCP8TFgf/CQoI/wgHCP8ICAj/CAgI/wgICP8ICAj/CAgI/wgICP8ICAj/CAgI/wgHCP8JCgj/ExYH/wsNCP8lHgf/OTAG/xMWB/8JCQj/CAgI/wgICP8ICAj/CAgI/wkJCf8JCQn/CQkJ/wkJCf8ICAn/KykH/z01B/8LDAn/FBcI/woKCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8KCgn/FBcI/wsMCf89NQf/KykH/wgICf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/woKCf88MQf/GRYI/xIVCf8ODwn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8JCQn/CQkJ/wkJCf8ODwn/EhUJ/xkWCP88MQf/CgoJ/wkJCf8JCQn/CQkJ/wkJCf8KCgr/CgoK/woKCv8JCQr/GBUJ/zwyCP8ODgr/FRcK/woKCv8KCgr/CgoK/woKCv8KCgr/CgoK/woKCv8KCgr/CgoK/woKCv8KCgr/CgoK/woKCv8KCgr/CgoK/woKCv8UFwr/Dg4K/zwyCP8YFQn/CQkK/woKCv8KCgr/CgoK/wsLC/8LCwv/CwsL/wkKC/8qIwr/LiYJ/xATC/8REwv/CwsL/wsLC/8LCwv/CwsL/wsLC/8LCwv/CwsL/wsLC/8LCwv/CwsL/wsLC/8LCwv/CwsL/wsLC/8LCwv/CwsL/xETC/8QEgv/LiYJ/yojCv8JCgv/CwsL/wsLC/8LCwv/DAwM/wwMDP8MDAz/CgoM/zUsCv8jHgv/FBcL/w8QDP8MDAz/DAwM/wwMDP8MDAz/DAwM/wwMDP8MDAz/CwoM/wsKDP8MDAz/DAwM/wwMDP8MDAz/DAwM/wwMDP8MDAz/DxAM/xQXC/8jHgv/NSwK/woKDP8MDAz/DAwM/wwMDP8NDQ3/DQ0N/w0NDf8SFQ3/QTsK/x4bDP8WGQz/Dw8N/w0NDf8NDQ3/DQ0N/w0NDf8NDQ3/DQ0N/wwMDf9JWQn/SVkJ/wwMDf8NDQ3/DQ0N/w0NDf8NDQ3/DQ0N/w0NDf8PEA3/FxoM/x4bDP9BOwr/EhUN/w0NDf8NDQ3/DQ0N/w0NDf8NDQ3/DQ0N/xMVDf9COwr/HhsM/xYZDP8PDw3/DQ0N/w0NDf8NDQ3/DQ0N/w0NDf8NDQ3/CwoN/1ltCP9ZbQj/CwoN/w0NDf8NDQ3/DQ0N/w0NDf8NDQ3/DQ0N/w8QDf8XGgz/HhsM/0I7Cv8TFQ3/DQ0N/w0NDf8NDQ3/Dg4O/w4ODv8ODg7/DAwO/zYuDP8lIA3/FhkN/xESDv8ODg7/Dg4O/w4ODv8ODg7/Dg4O/w4ODv8MCw7/Mz0L/zM9C/8MCw7/Dg4O/w4ODv8ODg7/Dg4O/w4ODv8ODg7/ERIO/xYZDf8lIA3/Ni4M/wwMDv8ODg7/Dg4O/w4ODv8PDw//Dw8P/w8PD/8NDg//LScN/zEqDf8UFg//FRcO/w8PD/8PDw//Dw8P/w8PD/8PDw//Dw8P/w0MD/81Pwz/NT8M/w0MD/8PDw//Dw8P/w8PD/8PDw//Dw8P/w8PD/8VFw//FBYP/zEqDf8tJw3/DQ4P/w8PD/8PDw//Dw8P/xAQEP8QEBD/EBAQ/w8PEP8eGw//QTYN/xMUEP8aHQ//EBAQ/xAQEP8QEBD/EBAQ/xAQEP8QEBD/Dg0Q/zZADf82QA3/Dg0Q/xAQEP8QEBD/EBAQ/xAQEP8QEBD/EBAQ/xodD/8TFBD/QTYN/x4bD/8PDxD/EBAQ/xAQEP8QEBD/ERER/xEREf8RERH/ERER/xISEf9CNw3/IB0Q/xkcEP8VFhD/ERER/xEREf8RERH/ERER/xEREf8PDhH/NkAO/zZADv8PDhH/ERER/xEREf8RERH/ERER/xEREf8VFhD/GRwQ/yAdEP9CNw3/EhIR/xEREf8RERH/ERER/xEREf8RERH/ERER/xEREf8RERH/EBAR/zIxD/9EPA7/FBUR/xwfEP8TExH/ERER/xEREf8RERH/ERER/w8OEf82QA7/NkAO/w8OEf8RERH/ERER/xEREf8RERH/ExMR/xwfEP8UFBH/RDwO/zIxD/8QEBH/ERER/xEREf8RERH/ERER/xISEv8SEhL/EhIS/xISEv8TExL/HSAR/0I5D/8uKBD/FhgS/x0gEf8UFBL/EhIS/xISEv8SEhL/EA8S/zdBDv83QQ7/EA8S/xISEv8SEhL/EhIS/xQUEv8dIBH/FhcS/y4oEP9COQ//HSAR/xMTEv8SEhL/EhIS/xISEv8SEhL/ExMT/xMTE/8TExP/ExMT/xMTE/8SEhP/GxkS/0g8D/8oIxH/FxgT/x4hEv8XGRP/ExMT/xMTE/8REBP/OEIP/zhCD/8REBP/ExMT/xMTE/8XGRP/HiES/xcYE/8oIxH/SDwP/xsZEv8SEhP/ExMT/xMTE/8TExP/ExMT/xMTE/8UFBT/FBQU/xQUFP8UFBT/FBQU/xQUFP8TExT/IB0T/0k9D/8vKRL/FhcU/xwfE/8eIRP/GhwT/xUVFP87RRD/O0UQ/xUVFP8aHBP/HiAT/xwfE/8WFxT/LykS/0k9D/8gHRP/ExMU/xQUFP8UFBT/FBQU/xQUFP8UFBT/FBQU/xUVFf8VFRX/FRUV/xUVFf8VFRX/FRUV/xUVFf8UFBX/HBsU/0Q7EP9GPhD/JCAT/xgZFP8aHBT/Gx0U/zpFEP87RhD/Gx0U/xocFP8YGRT/JCAT/0Y+EP9EOxD/HBsU/xQUFf8VFRX/FRUV/xUVFf8VFRX/FRUV/xUVFf8VFRX/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8VFBb/ICMV/zY0Ev9FOxH/RTsR/zcvEv8rJhT/KSYU/ykmFP8rJhT/Ny8S/0U7Ef9FOxH/NjQS/yAjFP8VFBb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8XFxb/FRYW/xcXFv8kIRX/NC0U/z01E/9JQhH/SUIR/z41E/80LRT/JCEV/xcXFv8VFhb/FxcW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xYWFv8WFhb/FhYW/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xYXF/8VFhf/FRUX/xwfFv8cHxb/FRUX/xUWF/8WFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/FxcX/xcXF/8XFxf/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8YGBj/GBgY/xgYGP8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xkZGf8ZGRn/GRkZ/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/Ghoa/xoaGv8aGhr/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=
				'
	}
	'classes.psm1'         = @{ 
		Priority = 'Critical'; Order = 3; Dependencies = @() 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'classes.psm1') 
		Base64Content = '

			'
	}
	'ini.psm1'             = @{ 
		Priority = 'Critical'; Order = 4; Dependencies = @('classes.psm1') 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'ini.psm1') 
		Base64Content = '

			'
	}
	'ui.psm1'              = @{ 
		Priority = 'Critical'; Order = 5; Dependencies = @('classes.psm1', 'ini.psm1') 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'ui.psm1') 
		Base64Content = '

			'
	}
	'datagrid.psm1'        = @{ 
		Priority = 'Important'; Order = 6; Dependencies = @('classes.psm1', 'ui.psm1', 'ini.psm1') 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'datagrid.psm1') 
		Base64Content = '

			'
	}
	'notifications.psm1'   = @{ 
		Priority = 'Important'; Order = 7; Dependencies = @('ftool.dll', 'classes.psm1', 'ui.psm1', 'ini.psm1', 'datagrid.psm1') 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'notifications.psm1') 
		Base64Content = '

			'
	}
	'runspace-helper.psm1' = @{ 
		Priority = 'Important'; Order = 8; Dependencies = @('ftool.dll', 'classes.psm1', 'ui.psm1', 'ini.psm1', 'datagrid.psm1') 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'runspace-helper.psm1') 
		Base64Content = '

			'
	}
	'launch.psm1'          = @{ 
		Priority = 'Important'; Order = 9; Dependencies = @('classes.psm1', 'ui.psm1', 'ini.psm1', 'datagrid.psm1', 'notifications.psm1') 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'launch.psm1') 
		Base64Content = '

			'
	}
	'login.psm1'           = @{ 
		Priority = 'Important'; Order = 10; Dependencies = @('ftool.dll', 'classes.psm1', 'ui.psm1', 'ini.psm1', 'datagrid.psm1', 'launch.psm1', 'notifications.psm1') 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'login.psm1') 
		Base64Content = '

			'
	}
	'ftool.psm1'           = @{ 
		Priority = 'Important'; Order = 11; Dependencies = @('ftool.dll', 'classes.psm1', 'ui.psm1', 'ini.psm1', 'datagrid.psm1', 'launch.psm1', 'login.psm1') 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'ftool.psm1') 
		Base64Content = '

			'
	}
	'reconnect.psm1'       = @{ 
		Priority = 'Important'; Order = 12; Dependencies = @('ftool.dll', 'classes.psm1', 'ui.psm1', 'ini.psm1', 'datagrid.psm1', 'launch.psm1', 'login.psm1', 'notifications.psm1') 
		FilePath = (Join-Path $global:DashboardConfig.Paths.Source 'reconnect.psm1') 
		Base64Content = '

			'
	}
}
#endregion Step: Define Module Metadata

#endregion Global Configuration

#region Environment Initialization and Checks

#region Function: ShowErrorDialog
function ShowErrorDialog
{
	param([Parameter(Mandatory = $true)][string]$Message)
	try
	{
		if (-not ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')))
		{
			InitializeClassesModule
		}
		[System.Windows.Forms.MessageBox]::Show($Message, 'Entropia Dashboard Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
	}
 catch
	{
		Write-Verbose "Failed to display error dialog: `"$Message`". Dialog Display Error: $($_.Exception.Message)"
	}
}
#endregion Function: ShowErrorDialog

#region Function: RequestElevation
function RequestElevation
{
	param()
	Write-Verbose 'Checking environment (Admin, 32-bit, Policy)...'
	[bool]$needsRestart = $false
	[System.Collections.ArrayList]$reason = @()

	[bool]$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	if (-not $isAdmin)
	{
		$needsRestart = $true; $null = $reason.Add('Administrator privileges required.')
	}

	[bool]$is32Bit = [IntPtr]::Size -eq 4
	if (-not $is32Bit)
	{
		$needsRestart = $true; $null = $reason.Add('32-bit execution required.')
	}

	[string]$currentPolicy = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue
	if ($currentPolicy -ne 'Bypass')
	{
		$needsRestart = $true
		$effectivePolicy = if ($currentPolicy -ne '') { $currentPolicy } else { Get-ExecutionPolicy }
		$null = $reason.Add("Execution Policy 'Bypass' required (Current effective: '$effectivePolicy').")
	}

	if ($needsRestart)
	{
		Write-Verbose "Restarting script needed: $($reason -join ' ')"
		[string]$psExe = Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
		if (-not (Test-Path $psExe -PathType Leaf))
		{
			ShowErrorDialog "FATAL: Required 32-bit PowerShell executable not found at '$psExe'. Cannot continue."
			exit 1
		}
        
		$encodedCommand = @'

'@
		try
  {
			$decodedCommand = [System.Text.Encoding]::UTF8.GetString(([System.Convert]::FromBase64String($encodedCommand)))
		}
		catch
		{
			ShowErrorDialog "FATAL: Failed to decode the embedded command. Error: $($_.Exception.Message)"
			exit 1
		}

		$tempScriptPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ([System.Guid]::NewGuid().ToString() + '.ps1'))

		try
		{
			[System.IO.File]::WriteAllText($tempScriptPath, $decodedCommand, [System.Text.Encoding]::UTF8)
			[string]$psArgs = "-WindowStyle Hidden -noexit -ExecutionPolicy Bypass -File `"$tempScriptPath`""
            
			$psi = New-Object System.Diagnostics.ProcessStartInfo
			$psi.FileName = $psExe
			$psi.Arguments = $psArgs
			$psi.UseShellExecute = $true
			$psi.Verb = 'RunAs'
            
			Write-Verbose "Attempting restart: `"$psExe`" $psArgs (Temp: $tempScriptPath)"
			[System.Diagnostics.Process]::Start($psi) | Out-Null
			Write-Verbose 'Success. Exiting current process.'
			exit 0
		}
		catch
		{
			ShowErrorDialog "FATAL: Failed to restart script (Admin/32-bit/Bypass). Error: $($_.Exception.Message)"
			if (Test-Path $tempScriptPath)
			{
				try { Remove-Item $tempScriptPath -ErrorAction SilentlyContinue } catch {}
			}
			exit 1
		}
	}
 else
	{
		Write-Verbose 'Script already running with required environment settings.'
	}
}
#endregion Function: RequestElevation

#region Function: InitializeScriptEnvironment
function InitializeScriptEnvironment
{
	<#
		.SYNOPSIS
			Verifies that the script environment meets all requirements *after* any potential restart attempt by RequestElevation.
		
		.DESCRIPTION
			This function performs final checks to ensure the script is operating in the correct environment before proceeding with core logic.
			It re-validates:
			1. Administrator Privileges: Confirms the script is now running elevated.
			2. 32-bit Mode: Confirms the script is now running in a 32-bit PowerShell process.
			3. Execution Policy: Confirms the process scope execution policy is 'Bypass'. If not (which shouldn't happen if RequestElevation worked),
			it makes a final attempt to set it using Set-ExecutionPolicy.
			
			If any check fails, it displays a specific error message using ShowErrorDialog and returns $false.
		
		.OUTPUTS
			[bool] Returns $true if all environment checks pass successfully, otherwise returns $false.
		
		.NOTES
			- This function should be called *after* RequestElevation. It acts as a final safeguard.
			- Failure here is typically fatal for the application, as indicated by the error messages and the return value.
			- The attempt to set ExecutionPolicy within this function is a fallback; ideally, RequestElevation should have ensured this.
		#>
	[CmdletBinding()]
	[OutputType([bool])] 
	param()
		
	Write-Verbose 'Verifying final script environment settings...'
	try
	{
		#region Step: Verify Administrator Privileges
		
		[bool]$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
		if (-not $isAdmin)
		{
			
			ShowErrorDialog 'FATAL: Application requires administrator privileges to run.'
			return $false
		}
		Write-Verbose '[OK] Running with administrator privileges.'
		#endregion Step: Verify Administrator Privileges
			
		#region Step: Verify 32-bit Execution Mode
		
		[bool]$is32Bit = [IntPtr]::Size -eq 4
		if (-not $is32Bit)
		{
			
			ShowErrorDialog 'FATAL: Application must run in 32-bit PowerShell mode.'
			return $false
		}
		Write-Verbose '[OK] Running in 32-bit mode.'
		#endregion Step: Verify 32-bit Execution Mode
			
		#region Step: Verify Process Execution Policy
		
		[string]$currentPolicy = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue
		if ($currentPolicy -ne 'Bypass')
		{
			
			Write-Verbose "  Process Execution Policy is not 'Bypass' (Current: '$(if (-not [string]::IsNullOrEmpty($currentPolicy)) { $currentPolicy } else { Get-ExecutionPolicy })'). Attempting final Set..."
			try
			{
				
				Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
				
				$currentPolicy = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue
				if ($currentPolicy -ne 'Bypass')
				{
					
					ShowErrorDialog "FATAL: Failed to set required PowerShell Execution Policy to 'Bypass'.  (Current: '$(if (-not [string]::IsNullOrEmpty($currentPolicy)) { $currentPolicy } else { Get-ExecutionPolicy })')."
					return $false
				}
				Write-Verbose '[OK] Execution policy successfully forced to Bypass for this process.'
			}
			catch
			{
				
				ShowErrorDialog "FATAL: Error setting PowerShell Execution Policy to 'Bypass'.  (Current: '$(if (-not [string]::IsNullOrEmpty($currentPolicy)) { $currentPolicy } else { Get-ExecutionPolicy })'). Error: $($_.Exception.Message)"
				return $false
			}
		}
		else
		{
			Write-Verbose "[OK] Execution policy is '$currentPolicy'."
		}
		#endregion Step: Verify Process Execution Policy
			
		
		Write-Verbose '  Environment verification successful.'
		return $true
	}
	catch
	{
		
		ShowErrorDialog "FATAL: An unexpected error occurred during environment verification: $($_.Exception.Message)"
		return $false
	}
}
#endregion Function: InitializeScriptEnvironment

#region Function: InitializeBaseConfig
function InitializeBaseConfig
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
			- Errors during directory creation or the write test are logged to the error stream and presented to the user via ShowErrorDialog,
			as these are typically fatal issues preventing the application from functioning correctly.
			- Log file clearing errors are reported as non-critical warnings.
		#>
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	Write-Verbose 'Initializing base configuration directories in %APPDATA%...'
	try
	{
		
		
		[string]$logDir = Split-Path -Path $global:DashboardConfig.Paths.Verbose -Parent

		
		
		[string[]]$directories = @(
			$global:DashboardConfig.Paths.App,     
			$global:DashboardConfig.Paths.Modules, 
			$logDir                                
		)
			
		
		foreach ($dir in $directories)
		{
			#region Step: Ensure Directory Exists
			
			if (-not (Test-Path -Path $dir -PathType Container))
			{
				Write-Verbose "  Directory not found. Creating: '$dir'"
				try
				{
					
					$null = New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop
				}
				catch
				{
					
					$errorMsg = "  Failed to create required directory '$dir'. Please check permissions or path validity. Error: $($_.Exception.Message)"
					Write-Verbose $errorMsg
					ShowErrorDialog $errorMsg
					return $false 
				}
			}
			else
			{
				Write-Verbose "  Directory exists: '$dir'"
			}
			#endregion Step: Ensure Directory Exists
				
			#region Step: Test Directory Writability
			
			
			[string]$testFile = Join-Path -Path $dir -ChildPath 'write_test.tmp'
			try
			{
				
				[System.IO.File]::WriteAllText($testFile, 'TestWriteAccess')
				
				Remove-Item -Path $testFile -Force -ErrorAction Stop
				Write-Verbose "  Directory is writable: '$dir'"
			}
			catch
			{
				
				$errorMsg = "  Cannot write to directory '$dir'. Please check permissions. Error: $($_.Exception.Message)"
				Write-Verbose $errorMsg
				ShowErrorDialog $errorMsg
				
				if (Test-Path -Path $testFile -PathType Leaf)
				{
					Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
				}
				return $false 
			}
			#endregion Step: Test Directory Writability
		} 
			
		#region Step: Clear Log Files
		
		#endregion Step: Clear Log Files
			
		
		Write-Verbose '  Base configuration directories initialized and verified successfully.'
		
		$global:DashboardConfig.State.ConfigInitialized = $true
		return $true
	}
	catch
	{
		
		$errorMsg = "  An unexpected error occurred during base configuration directory initialization: $($_.Exception.Message)"
		Write-Verbose $errorMsg 
		ShowErrorDialog $errorMsg
		return $false
	}
}
#endregion Function: InitializeBaseConfig

#endregion Environment Initialization and Checks

#region Module Handling Functions

#region Function: WriteModule
function WriteModule
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
	[CmdletBinding(DefaultParameterSetName = 'FilePath')] 
	[OutputType([string])]
	param (
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$ModuleName, 
		
		[Parameter(Mandatory = $true, ParameterSetName = 'FilePath', Position = 1)]
		[ValidateScript({ Test-Path $_ -PathType Leaf })] 
		[string]$Content, 
		
		[Parameter(Mandatory = $true, ParameterSetName = 'Base64Content')]
		[string]$ContentBase64 
	)
		
	
	
	[string]$modulesDir = $global:DashboardConfig.Paths.Modules
	
	
	[string]$finalPath = Join-Path -Path $modulesDir -ChildPath $ModuleName
		
	Write-Verbose "Executing WriteModule for '$ModuleName' to '$finalPath'"
	try
	{
		#region Step: Ensure Target Directory Exists
		
		if (-not (Test-Path -Path $modulesDir -PathType Container))
		{
			Write-Verbose "Target module directory not found, attempting creation: '$modulesDir'"
			try
			{
				$null = New-Item -Path $modulesDir -ItemType Directory -Force -ErrorAction Stop
				Write-Verbose "Target module directory created successfully: '$modulesDir'"
			}
			catch
			{
				
				Write-Verbose "Failed to create target module directory '$modulesDir': $($_.Exception.Message)"
				return $null 
			}
		}
		#endregion Step: Ensure Target Directory Exists
			
		#region Step: Get Content Bytes from Source (File or Base64)
		
		[byte[]]$bytes = $null
		Write-Verbose "  ParameterSetName: $($PSCmdlet.ParameterSetName)"
				
		
		if ($PSCmdlet.ParameterSetName -eq 'Base64Content')
		{
			if ([string]::IsNullOrEmpty($ContentBase64))
			{
				Write-Verbose "  ModuleName '$ModuleName': ContentBase64 parameter was provided but is empty."
				return $null
			}
			try
			{
				$bytes = [System.Convert]::FromBase64String($ContentBase64)
				Write-Verbose "  Decoded Base64 content for '$ModuleName' ($($bytes.Length) bytes)."
			}
			catch
			{
				
				Write-Verbose "  Failed to decode Base64 content for '$ModuleName': $($_.Exception.Message)"
				return $null
			}
		}
		
		elseif ($PSCmdlet.ParameterSetName -eq 'FilePath')
		{
			
			if ([string]::IsNullOrEmpty($Content) -or -not ([System.IO.File]::Exists($Content)) )
			{
				Write-Verbose "  ModuleName '$ModuleName': Source file path '$Content' is invalid or does not exist."
				return $null 
			}
			try
			{
				$bytes = [System.IO.File]::ReadAllBytes($Content)
				Write-Verbose "  Read source file content for '$ModuleName' from '$Content' ($($bytes.Length) bytes)."
			}
			catch
			{
				
				Write-Verbose "  Failed to read source file '$Content' for '$ModuleName': $($_.Exception.Message)"
				return $null
			}
		}
		else 
		{
			Write-Verbose "  ModuleName '$ModuleName': Invalid parameter combination or missing content."
			return $null
		}
				
		
		if ($null -eq $bytes)
		{
			Write-Verbose "  Failed to obtain content bytes for '$ModuleName'. Source data might be empty or invalid."
			return $null
		}
		#endregion Step: Get Content Bytes from Source (File or Base64)
			
		#region Step: Check if File Needs Updating (Size and Hash Comparison)
		
		[bool]$updateNeeded = $true
		if (Test-Path -Path $finalPath -PathType Leaf) 
		{
			Write-Verbose "  Target file exists: '$finalPath'. Comparing size and hash..."
			try
			{
				
				
				$fileInfo = Get-Item -LiteralPath $finalPath -Force -ErrorAction Stop
						
				
				if ($fileInfo.Length -eq $bytes.Length)
				{
					Write-Verbose "  File sizes match ($($bytes.Length) bytes). Comparing SHA256 hashes..."
					
					
					[string]$existingHash = (Get-FileHash -LiteralPath $finalPath -Algorithm SHA256 -ErrorAction Stop).Hash
							
					
					
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
							
					Write-Verbose " - Existing Hash: $existingHash"
					Write-Verbose " - New Hash:    - $newHash"
							
					
					if ($existingHash -eq $newHash)
					{
						Write-Verbose "  Hashes match for '$ModuleName'. No update needed."
						$updateNeeded = $false
						
						return $finalPath
					}
					else
					{
						Write-Verbose "  Hashes differ for '$ModuleName'. Update required." 
					}
				}
				else
				{
					Write-Verbose "  File sizes differ (Existing: $($fileInfo.Length), New: $($bytes.Length)). Update required." 
				}
			}
			catch
			{
				
				
				Write-Verbose "  Could not compare size/hash for '$ModuleName' (Path: '$finalPath'). Will attempt to overwrite. Error: $($_.Exception.Message)"
				$updateNeeded = $true
			}
		}
		else
		{
			Write-Verbose "  Target file does not exist: '$finalPath'. Writing new file." 
			$updateNeeded = $true
		}
		#endregion Step: Check if File Needs Updating (Size and Hash Comparison)
			
		#region Step: Write File to Target Path (with Retry on IO Exception)
		if ($updateNeeded)
		{
			
			
			[int]$timeoutMilliseconds = 5000  
			
			[int]$retryDelayMilliseconds = 100 
			
			[datetime]$startTime = Get-Date
			
			[bool]$fileWritten = $false
			
			[int]$attempts = 0
					
			Write-Verbose "  Attempting to write file: '$finalPath'"
			while (((Get-Date) - $startTime).TotalMilliseconds -lt $timeoutMilliseconds)
			{
				$attempts++
				try
				{
					
					[System.IO.File]::WriteAllBytes($finalPath, $bytes)
					$fileWritten = $true
					Write-Verbose "  Successfully wrote '$ModuleName' to '$finalPath' on attempt $attempts."
					break 
				}
				catch [System.IO.IOException]
				{
					
					Write-Verbose "  Attempt $($attempts): IO Error writing '$finalPath' (Retrying in $retryDelayMilliseconds ms): $($_.Exception.Message)"
					
					if (((Get-Date) - $startTime).TotalMilliseconds + $retryDelayMilliseconds -ge $timeoutMilliseconds)
					{
						Write-Verbose "  Timeout nearing, breaking retry loop for '$finalPath'."
						break 
					}
					Start-Sleep -Milliseconds $retryDelayMilliseconds
				}
				catch
				{
					
					Write-Verbose "  Attempt $($attempts): Non-IO Error writing '$finalPath': $($_.Exception.Message)"
					$fileWritten = $false 
					break 
				}
			} 
					
			
			if (-not $fileWritten)
			{
				Write-Verbose "  Failed to write module '$ModuleName' to '$finalPath' after $attempts attempts within $timeoutMilliseconds ms timeout."
				return $null 
			}
		} 
		#endregion Step: Write File to Target Path (with Retry on IO Exception)
			
		
		return $finalPath
	}
	catch
	{
		
		Write-Verbose "  An unexpected error occurred in WriteModule for '$ModuleName': $($_.Exception.Message)"
		return $null
	}
}
#endregion Function: WriteModule


#region Function: ImportModuleUsingReflection
function ImportModuleUsingReflection
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

	Write-Verbose "Attempting reflection-style import (Global Scope Injection) for '$ModuleName'."

	try
	{
		if (-not (Test-Path -Path $Path -PathType Leaf)) { return $false }

		
		[string]$moduleContent = [System.IO.File]::ReadAllText($Path)
		if ($null -eq $moduleContent)
		{
			if (Test-Path -Path $Path -PathType Leaf)
			{
				Write-Verbose "Import-ModuleUsingReflection: Module file '$Path' is empty. Considering import successful (no-op)." -ForegroundColor Yellow
				$global:DashboardConfig.Resources.LoadedModuleContent[$ModuleName] = ''
				return $true
			}
			else
			{
				Write-Verbose "Import-ModuleUsingReflection: Failed to read module file '$Path'." -ForegroundColor Red
				return $false
			}
		}
		
		if ($global:DashboardConfig -and -not $global:DashboardConfig.ContainsKey('Resources'))
		{
			$global:DashboardConfig['Resources'] = @{}
		}
		
		if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and -not $global:DashboardConfig.Resources.ContainsKey('LoadedModuleContent'))
		{
			$global:DashboardConfig.Resources['LoadedModuleContent'] = @{}
		}
		$global:DashboardConfig.Resources.LoadedModuleContent[$ModuleName] = $moduleContent
		Write-Verbose "Read and stored original content for '$ModuleName'." -ForegroundColor DarkGray
		

		[string]$noOpExportFunc = @"
function Export-ModuleMember { 
param(
[string]`$Function, 
[string]`$Variable, 
[string]`$Alias, 
[string]`$Cmdlet
)
}
"@
		[string]$modifiedContent = @"
$noOpExportFunc


$moduleContent

"@

		Write-Verbose "Creating ScriptBlock and executing modified content globally via Invoke-Command for '$ModuleName'..." -ForegroundColor DarkGray
		[scriptblock]$scriptBlock = [ScriptBlock]::Create($modifiedContent)
		try
		{
			
			$null = Invoke-Command -ScriptBlock $scriptBlock
					
			
			
			if (-not $?)
			{
				Write-Verbose "Execution of '$ModuleName' content via Invoke-Command completed, but non-terminating errors occurred (see logs above). This is considered a failure." -ForegroundColor Yellow
				return $false
			}
					
			
			Write-Verbose "Successfully finished executing modified script block for '$ModuleName' via Invoke-Command." -ForegroundColor Green
			return $true
		}
		catch
		{
			Write-Verbose "FATAL error during global import for '$ModuleName': $($_.Exception.Message)"
			return $false
		}
	} 
	catch
	{
		Write-Verbose "FATAL error during reflection-style import (InvokeCommand) setup for '$ModuleName': $($_.Exception.Message)" -ForegroundColor Red
		return $false 
	}
}
#endregion Function: ImportModuleUsingReflection


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
			b. Calls 'WriteModule' to ensure the module file (or resource like .dll, .ico) exists in the %APPDATA%\modules directory, handling source file paths or Base64 content, and using hash checks for efficiency. If WriteModule fails, records the failure. Critical module write failures trigger the critical failure flag.
			c. If WriteModule succeeds, adds the module name and its written path to '$global:DashboardConfig.LoadedModules'. This satisfies dependency checks for subsequent modules, including non-PSM1 files like DLLs or icons.
			d. If the module is a PowerShell module (.psm1):
			i. Attempts multiple import strategies in sequence until one succeeds:
			- Attempt 1 (Preferred): Standard `Import-Module`. If running as EXE, it first modifies the content in memory to prepend a no-op `Export-ModuleMember`, writes this to a temporary file, imports the temp file, and then deletes it. If running as a script, it imports the written module path directly. Success is verified by checking `Get-Module`.
			- Attempt 2 (Alternative): Calls `ImportModuleUsingReflection` function (InvokeCommand in global scope). **Crucially, after this attempt returns true, this function now performs an additional verification step using `Get-Command` for key functions expected from the module.** If key functions are missing, Attempt 2 is marked as failed, and the process proceeds to Attempt 3.
			- Attempt 3 (Last Resort): Uses `Invoke-Expression` on the module content after attempting to remove/comment out `Export-ModuleMember` calls using string replacement. This attempt includes its own verification and global re-definition of functions. **(Security Risk)**
			ii. If all import attempts fail for a .psm1 module, records the failure, removes the module from '$global:DashboardConfig.LoadedModules' (as it was written but not imported), and triggers the critical failure flag if the module was critical.
			5. After processing all modules, checks the critical failure flag. If set, returns a status object indicating failure.
			6. Logs warnings for any 'Important' modules that failed and informational messages for 'Optional' module failures.
			7. If no critical failures occurred, returns a status object indicating overall success (though non-critical modules may have failed).
		
		.OUTPUTS
			[PSCustomObject] Returns an object with the following properties:
			- Status [bool]: $true if all 'Critical' modules were successfully written and (if applicable) imported without fatal errors. $false if any 'Critical' module failed or if an unhandled exception occurred.
			- LoadedModules [hashtable]: A hashtable containing {ModuleName = Path} entries for all modules that were successfully written to the AppData directory by WriteModule (includes .psm1, .dll, .ico, etc.). Note that for .psm1, inclusion here doesn't guarantee successful *import*, only successful writing/verification. Check FailedModules for import status.
			- FailedModules [hashtable]: A hashtable containing {ModuleName = ErrorMessage} entries for modules that failed during dependency check, writing (WriteModule), or importing (for .psm1 files).
			- CriticalFailure [bool]: $true if a module marked with Priority='Critical' failed at any stage (dependency, write, or import). $false otherwise.
			- Exception [string]: (Optional) Included only if an unexpected, unhandled exception occurred within the Import-DashboardModules function itself. Contains the exception message.
		
		.NOTES
			- The multi-attempt import strategy for .psm1 files adds complexity but aims for robustness, especially in potentially problematic EXE execution environments.
			- Attempt 2 now includes verification. If it passes, Attempt 3 (Invoke-Expression) is skipped.
			- The use of `Invoke-Expression` (Attempt 3) remains a significant security risk and should ideally be avoided by refactoring modules to work with Attempt 1 or a reliable Attempt 2.
			- Dependency checking relies on modules being added to `$global:DashboardConfig.LoadedModules` *after* successful execution of `WriteModule`.
			- Error reporting distinguishes between Critical, Important, and Optional module failures. Only Critical failures halt the application startup process.
		#>
	[CmdletBinding()]
	[OutputType([PSCustomObject])]
	param()
		
	Write-Verbose 'Initializing module import process...'
		
	
	
	$result = [PSCustomObject]@{
		Status          = $false 
		LoadedModules   = $global:DashboardConfig.LoadedModules 
		FailedModules   = @{}    
		CriticalFailure = $false 
		Exception       = $null  
	}
	
	[hashtable]$failedModules = $result.FailedModules
		
	try
	{
		#region Step: Determine Execution Context (EXE vs. Script)
		
		
		$currentProcess = Get-Process -Id $PID -ErrorAction Stop 

		
		[string]$processPath = $currentProcess.Path 

		
		[bool]$isRunningAsExe = $processPath -like '*.exe' -and ($processPath -notlike '*powershell.exe' -and $processPath -notlike '*pwsh.exe')
				
		
		if ($global:DashboardConfig -and -not $global:DashboardConfig.ContainsKey('State'))
		{
			$global:DashboardConfig['State'] = @{}
		}
		if ($global:DashboardConfig -and $global:DashboardConfig.State)
		{
			$global:DashboardConfig.State.IsRunningAsExe = $isRunningAsExe 
		}
		Write-Verbose "  Execution context detected: $(if($isRunningAsExe){'Compiled EXE'} else {'PowerShell Script'}) (Process Path: '$processPath')"
		#endregion Step: Determine Execution Context (EXE vs. Script)
			
		#region Step: Sort Modules by Defined 'Order' Property
		Write-Verbose "  Sorting modules based on 'Order' property..."
		
		
		$sortedModules = $global:DashboardConfig.Modules.GetEnumerator() |
				Where-Object {
			
			$_.Value -is [hashtable] -and $_.Value.ContainsKey('Order') -and $_.Value.Order -is [int]
		} |
				Sort-Object { $_.Value.Order } -ErrorAction SilentlyContinue 
				
		if (-not $sortedModules -or $sortedModules.Count -ne $global:DashboardConfig.Modules.Count)
		{
			
			$invalidModules = $global:DashboardConfig.Modules.GetEnumerator() | Where-Object { -not ($_.Value -is [hashtable] -and $_.Value.ContainsKey('Order') -and $_.Value.Order -is [int]) }
			$errorMessage = "  Failed to sort modules or found invalid module configurations. Check structure in `$global:DashboardConfig.Modules."
			if ($invalidModules)
			{
				$errorMessage += " Invalid modules: $($invalidModules.Key -join ', ')"
			}
			Write-Verbose $errorMessage
			$result.Status = $false
			$result.CriticalFailure = $true 
			$failedModules['Module Sorting/Validation'] = $errorMessage
			return $result 
		}
		Write-Verbose "  Processing $($sortedModules.Count) modules in defined order."
		#endregion Step: Sort Modules by Defined 'Order' Property
			
		#region Step: Process Each Module in Sorted Order
		foreach ($entry in $sortedModules)
		{
			
			[string]$moduleName = $entry.Key
			
			$moduleInfo = $entry.Value 
					
			Write-Verbose "Processing Module: '$moduleName' (Priority: $($moduleInfo.Priority), Order: $($moduleInfo.Order))"
					
			#region SubStep: Check Dependencies
			Write-Verbose '- Checking dependencies...'
			
			[bool]$dependenciesMet = $true
			
			if ($moduleInfo.Dependencies -and $moduleInfo.Dependencies -is [array] -and $moduleInfo.Dependencies.Count -gt 0)
			{
				Write-Verbose "  - Required: $($moduleInfo.Dependencies -join ', ')"
				foreach ($dependency in $moduleInfo.Dependencies)
				{
					
					if (-not $global:DashboardConfig.LoadedModules.ContainsKey($dependency))
					{
						$errorMessage = "- Dependency NOT MET: Module '$dependency' must be loaded before '$moduleName'."
						Write-Verbose "- $errorMessage"
						$failedModules[$moduleName] = $errorMessage
						$dependenciesMet = $false
						
						if ($moduleInfo.Priority -eq 'Critical')
						{
							Write-Verbose "- CRITICAL FAILURE: Critical module '$moduleName' cannot load due to missing dependency '$dependency'."
							$result.CriticalFailure = $true
						}
						break 
					}
					else
					{
						Write-Verbose "  - Dependency satisfied: '$dependency' is loaded."
					}
				}
			}
			else
			{
				Write-Verbose "  - No dependencies listed for '$moduleName'."
			}
						
			
			if (-not $dependenciesMet)
			{
				continue
			} 
					
			#endregion SubStep: Check Dependencies
					
			#region SubStep: Write Module to AppData Directory (Using WriteModule)
			
			[string]$modulePath = $null
			Write-Verbose "- Ensuring module file exists in AppData via WriteModule for '$moduleName'..."
						
			
			try
			{
				if ($moduleInfo.ContainsKey('FilePath'))
				{
					[string]$sourceFilePath = $moduleInfo.FilePath
					
					if (-not (Test-Path $sourceFilePath -PathType Leaf))
					{
						throw "Source FilePath specified in config does not exist or is not a file: '$sourceFilePath'"
					}
					Write-Verbose "Calling WriteModule with source FilePath: '$sourceFilePath'"
					$modulePath = WriteModule -ModuleName $moduleName -Content $sourceFilePath -ErrorAction Stop 
				}
				elseif ($moduleInfo.ContainsKey('Base64Content'))
				{
					[string]$base64Content = $moduleInfo.Base64Content
					Write-Verbose "Calling WriteModule with Base64Content (Length: $($base64Content.Length))"
					
					if ([string]::IsNullOrEmpty($base64Content))
					{
						throw "Base64Content for module '$moduleName' is empty."
					}
					$modulePath = WriteModule -ModuleName $moduleName -ContentBase64 $base64Content -ErrorAction Stop
				}
				else
				{
					
					throw "Invalid module configuration format for '$moduleName' - missing FilePath or Base64Content."
				}
								
				
				if ([string]::IsNullOrEmpty($modulePath))
				{
					
					throw "WriteModule returned null or empty path for '$moduleName', indicating write failure."
				}
								
				Write-Verbose "- [OK] Module file ready/verified: '$modulePath'"
				
				
				
				if ($global:DashboardConfig -and -not $global:DashboardConfig.ContainsKey('LoadedModules'))
				{
					$global:DashboardConfig['LoadedModules'] = @{}
				}
				if ($global:DashboardConfig -and $global:DashboardConfig.LoadedModules)
				{
					$global:DashboardConfig.LoadedModules[$moduleName] = $modulePath
				}
								
			}
			catch
			{
				
				$errorMessage = "- Failed to write or verify module file for '$moduleName'. Error: $($_.Exception.Message)"
				Write-Verbose "- $errorMessage"
				$failedModules[$moduleName] = $errorMessage
				
				if ($moduleInfo.Priority -eq 'Critical')
				{
					Write-Verbose "- CRITICAL FAILURE: Failed to write critical module '$moduleName'."
					$result.CriticalFailure = $true
				}
				continue 
			}
			#endregion SubStep: Write Module to AppData Directory (Using WriteModule)
						
			#region SubStep: Import PowerShell Modules (.psm1)
			
			if ($moduleName -like '*.psm1')
			{
				Write-Verbose "Attempting to import PowerShell module '$moduleName' from '$modulePath'..."
				
				[bool]$importSuccess = $false
				
				[string]$importErrorDetails = 'All import attempts failed.'
				[string]$moduleBaseName = [System.IO.Path]::GetFileNameWithoutExtension($moduleName)

				
				if (-not $importSuccess)
				{
					Write-Verbose '- Attempt 1: Using standard Import-Module...' 
					try
					{
						
						[string]$effectiveModulePath = $modulePath
						
						[string]$tempModulePath = $null
										
						if ($isRunningAsExe)
						{
							Write-Verbose '  - (Running as EXE: Prepending no-op Export-ModuleMember to temporary file for import)' 
							
							$tempModulePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ('{0}_{1}.psm1' -f $moduleBaseName, [System.Guid]::NewGuid().ToString('N')))
							Write-Verbose "  - Temp file path: $tempModulePath" 
											
							
							
							
							[string]$originalContent = $null
							if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and $global:DashboardConfig.Resources.LoadedModuleContent -and $global:DashboardConfig.Resources.LoadedModuleContent.ContainsKey($moduleName))
							{
								$originalContent = $global:DashboardConfig.Resources.LoadedModuleContent[$moduleName]
							}
							if ($null -eq $originalContent)
							{
								$originalContent = [System.IO.File]::ReadAllText($modulePath)
							} 
											
							
							$noOpExportFunc = "function Export-ModuleMember { param([Parameter(ValueFromPipeline=`$true)][string[]]`$Function, [string[]]`$Variable, [string[]]`$Alias, [string[]]`$Cmdlet) { Write-Verbose ""Ignoring Export-ModuleMember (EXE Mode Import for '$($using:moduleName)')"" } }"
							
							Set-Content -Path $tempModulePath -Value "$noOpExportFunc`n`n# --- Original module content ($moduleName) follows ---`n$originalContent" -Encoding UTF8 -Force -ErrorAction Stop
							$effectiveModulePath = $tempModulePath 
						}
										
						
						Import-Module -Name $effectiveModulePath -Force -ErrorAction Stop
										
						
						if (Get-Module -Name $moduleBaseName -ErrorAction SilentlyContinue)
						{
							$importSuccess = $true
							Write-Verbose "- [OK] Attempt 1: SUCCESS (Standard Import-Module verified for '$moduleBaseName')." 
						}
						else
						{
							
							Write-Verbose "- Attempt 1: FAILED (Standard Import-Module) - Module '$moduleBaseName' not found via Get-Module after import call." 
							$importErrorDetails = "Standard Import-Module completed but module '$moduleBaseName' could not be verified via Get-Module."
							
							Remove-Module -Name $moduleBaseName -Force -ErrorAction SilentlyContinue
						}
										
					}
					catch
					{
						Write-Verbose "- Attempt 1: FAILED (Standard Import-Module Error): $($_.Exception.Message)" 
						$importErrorDetails = "Standard Import-Module Error: $($_.Exception.Message)"
						
						Remove-Module -Name $moduleBaseName -Force -ErrorAction SilentlyContinue
					}
					finally
					{
						
						if ($tempModulePath -and (Test-Path $tempModulePath))
						{
							Write-Verbose "  - Cleaning up temporary file: $tempModulePath" 
							Remove-Item -Path $tempModulePath -Force -ErrorAction SilentlyContinue
						}
					}
				} 
																
				
				if (-not $importSuccess)
				{
					
					if (Get-Command ImportModuleUsingReflection -ErrorAction SilentlyContinue)
					{
						Write-Verbose '- Attempt 2: Using alternative ImportModuleUsingReflection (InvokeCommand)...' 
						try
						{
							
							if (ImportModuleUsingReflection -Path $modulePath -ModuleName $moduleName -ErrorAction Stop)
							{
								
								Write-Verbose "  - Attempt 2: InvokeCommand finished. Verifying key functions globally for '$moduleName'..." 

								$moduleFileName = (Split-Path $moduleName -Leaf).Trim().ToLower()
								Write-Verbose "DEBUG: moduleFileName is '$moduleFileName' - Length: $($moduleFileName.Length)" 
								
								$keyFunctionsToCapture = @()
								if ($moduleFileName -in @('ini.psm1', 'ini'))
								{ 
									$keyFunctionsToCapture = @('CopyOrderedDictionary','GetIniFileContent','LoadDefaultConfigOnError','InitializeIniConfig','ReadConfig','WriteConfig') 
								}
								elseif ($moduleFileName -in @('ui.psm1', 'ui'))
								{ 
									$keyFunctionsToCapture = @('InitializeUI','RegisterUIEventHandlers','ShowSettingsForm','HideSettingsForm','SetUIElement','SyncConfigToUI','SyncUIToConfig','ShowInputBox','RefreshLoginProfileSelector','SyncProfilesToConfig') 
								}
								elseif ($moduleFileName -in @('datagrid.psm1', 'datagrid'))
								{ 
									$keyFunctionsToCapture = @('TestValidParameters','RestoreWindowStyles','GetProcessList','RemoveTerminatedProcesses','NewRowLookupDictionary','UpdateExistingRow','UpdateRowIndices','AddNewProcessRow','StartWindowStateCheck','FindTargetRow','ClearOldProcessCache','GetProcessProfile','SetWindowToolStyle','UpdateDataGrid','StartDataGridUpdateTimer' ) 
								}
								elseif ($moduleFileName -in @('launch.psm1', 'launch'))
								{ 
									$keyFunctionsToCapture = @('StartClientLaunch','Write-Verbose','InvokeSavedLaunchSequence','StopClientLaunch') 
								}
								elseif ($moduleFileName -in @('login.psm1', 'login'))
								{ 
									$keyFunctionsToCapture = @('GetClientLogPath','Update-Progress','LoginSelectedRow','CleanUpLoginResources') 
								}
								elseif ($moduleFileName -in @('ftool.psm1', 'ftool'))
								{ 
									$keyFunctionsToCapture = @('SetHotkey','ResumeAllHotkeys','ResumePausedKeys','ResumeHotkeysForOwner','RemoveAllHotkeys','TestHotkeyConflict','InvokeFtoolAction','PauseAllHotkeys','PauseHotkeysForOwner','UnregisterHotkeyInstance','ToggleSpecificFtoolInstance','ToggleInstanceHotkeys','GetVirtualKeyMappings','NormalizeKeyString','ParseKeyString','GetKeyCombinationString','IsModifierKeyCode','Show-KeyCaptureDialog','LoadFtoolSettings','FindOrCreateProfile','InitializeExtensionTracking','GetNextExtensionNumber','FindExtensionKeyByControl','LoadExtensionSettings','UpdateSettings','IsWindowBelow','CreatePositionTimer','RepositionExtensions','CreateSpammerTimer','ToggleButtonState','CheckRateLimit','AddFormCleanupHandler','CleanupInstanceResources','StopFtoolForm','RemoveExtension','FtoolSelectedRow','CreateFtoolForm','AddFtoolEventHandlers','CreateExtensionPanel','AddExtensionEventHandlers') 
								}
								elseif ($moduleFileName -in @('reconnect.psm1', 'reconnect'))
								{ 
									$keyFunctionsToCapture = @('StartDisconnectWatcher','StopDisconnectWatcher','InvokeReconnectionSequence') 
								}
								elseif ($moduleFileName -in @('notifications.psm1', 'notifications'))
								{ 
									$keyFunctionsToCapture = @('UpdateNotificationPositions','CloseToast','ShowToast','ShowInteractiveNotification','ShowReconnectInteractiveNotification') 
								}
								elseif ($moduleFileName -in @('runspace-helper.psm1', 'runspace-helper'))
								{ 
									$keyFunctionsToCapture = @('NewManagedRunspace','InvokeInManagedRunspace','DisposeManagedRunspace') 
								}
								else
								{
									Write-Verbose "WARNING: No verification list found for module '$moduleFileName' (Length: $($moduleFileName.Length)). verification skipped." 
								}

								[bool]$attempt2VerificationPassed = $true 
								[string]$missingFunction = $null

								if ($keyFunctionsToCapture.Count -gt 0)
								{
									foreach ($funcName in $keyFunctionsToCapture)
									{
										if (-not (Get-Command -Name $funcName -CommandType Function -ErrorAction SilentlyContinue))
										{
											$attempt2VerificationPassed = $false
											$missingFunction = $funcName
											Write-Verbose "  - Attempt 2: VERIFICATION FAILED. Function '$funcName' not found globally after InvokeCommand." 
											$importErrorDetails = "Attempt 2 (InvokeCommand) completed but verification failed: Function '$funcName' not found globally."
											break 
										}
									}
								}
								else
								{
									Write-Verbose "  - Attempt 2: No specific key functions listed for verification for '$moduleName'. Assuming success based on InvokeCommand completion." 
									
									$attempt2VerificationPassed = $true 
								}

								
								if ($attempt2VerificationPassed)
								{
									Write-Verbose "- [OK] Attempt 2: SUCCESS (InvokeCommand completed AND key functions verified for '$moduleName')." 
									$importSuccess = $true
								}
								else
								{
									
									$importSuccess = $false
								}
								
							}
							else 
							{
								Write-Verbose '- Attempt 2: FAILED (ImportModuleUsingReflection returned false).' 
								$importErrorDetails = 'ImportModuleUsingReflection returned false (fatal execution error).'
								$importSuccess = $false 
							}
						}
						catch 
						{
							Write-Verbose "- Attempt 2: FAILED (Error calling ImportModuleUsingReflection): $($_.Exception.Message)" 
							$importErrorDetails = "Error calling ImportModuleUsingReflection: $($_.Exception.Message)"
							$importSuccess = $false 
						}
					}
					else 
					{
						Write-Verbose '- Attempt 2: SKIPPED (ImportModuleUsingReflection function not found).' 
					}
				} 
								
				
				
				if (-not $importSuccess)
				{
					Write-Verbose '- Attempt 3: Using LAST RESORT Invoke-Expression (Security Risk!)...' 
					
					$functionsCapturedInThisAttempt = @{}
					try
					{
						
						[string]$invokeContent = $null
						if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and $global:DashboardConfig.Resources.LoadedModuleContent -and $global:DashboardConfig.Resources.LoadedModuleContent.ContainsKey($moduleName))
						{
							$invokeContent = $global:DashboardConfig.Resources.LoadedModuleContent[$moduleName]
						}
						if ($null -eq $invokeContent)
						{
							$invokeContent = [System.IO.File]::ReadAllText($modulePath)
						} 

						
						$invokeContent = $invokeContent -replace '(?m)^\s*Export-ModuleMember.*', "# Export-ModuleMember call disabled by Invoke-Expression wrapper for $moduleName"

						
						Invoke-Expression -Command $invokeContent -ErrorAction Stop

						
						$iexCompletedWithoutTerminatingError = $?
						$moduleFileName = (Split-Path $moduleName -Leaf).Trim().ToLower()
						Write-Verbose "DEBUG: moduleFileName is '$moduleFileName' - Length: $($moduleFileName.Length)" 
						
						$keyFunctionsToCapture = @()
						
						if ($moduleFileName -in @('ini.psm1', 'ini'))
						{ 
							$keyFunctionsToCapture = @('CopyOrderedDictionary','GetIniFileContent','LoadDefaultConfigOnError','InitializeIniConfig','ReadConfig','WriteConfig') 
						}
						elseif ($moduleFileName -in @('ui.psm1', 'ui'))
						{ 
							$keyFunctionsToCapture = @('InitializeUI','RegisterUIEventHandlers','ShowSettingsForm','HideSettingsForm','SetUIElement','SyncConfigToUI','SyncUIToConfig','ShowInputBox','RefreshLoginProfileSelector','SyncProfilesToConfig') 
						}
						elseif ($moduleFileName -in @('datagrid.psm1', 'datagrid'))
						{ 
							$keyFunctionsToCapture = @('TestValidParameters','RestoreWindowStyles','GetProcessList','RemoveTerminatedProcesses','NewRowLookupDictionary','UpdateExistingRow','UpdateRowIndices','AddNewProcessRow','StartWindowStateCheck','FindTargetRow','ClearOldProcessCache','GetProcessProfile','SetWindowToolStyle','UpdateDataGrid','StartDataGridUpdateTimer' ) 
						}
						elseif ($moduleFileName -in @('launch.psm1', 'launch'))
						{ 
							$keyFunctionsToCapture = @('StartClientLaunch','Write-Verbose','InvokeSavedLaunchSequence','StopClientLaunch') 
						}
						elseif ($moduleFileName -in @('login.psm1', 'login'))
						{ 
							$keyFunctionsToCapture = @('GetClientLogPath','Update-Progress','LoginSelectedRow','CleanUpLoginResources') 
						}
						elseif ($moduleFileName -in @('ftool.psm1', 'ftool'))
						{ 
							$keyFunctionsToCapture = @('SetHotkey','ResumeAllHotkeys','ResumePausedKeys','ResumeHotkeysForOwner','RemoveAllHotkeys','TestHotkeyConflict','InvokeFtoolAction','PauseAllHotkeys','PauseHotkeysForOwner','UnregisterHotkeyInstance','ToggleSpecificFtoolInstance','ToggleInstanceHotkeys','GetVirtualKeyMappings','NormalizeKeyString','ParseKeyString','GetKeyCombinationString','IsModifierKeyCode','Show-KeyCaptureDialog','LoadFtoolSettings','FindOrCreateProfile','InitializeExtensionTracking','GetNextExtensionNumber','FindExtensionKeyByControl','LoadExtensionSettings','UpdateSettings','IsWindowBelow','CreatePositionTimer','RepositionExtensions','CreateSpammerTimer','ToggleButtonState','CheckRateLimit','AddFormCleanupHandler','CleanupInstanceResources','StopFtoolForm','RemoveExtension','FtoolSelectedRow','CreateFtoolForm','AddFtoolEventHandlers','CreateExtensionPanel','AddExtensionEventHandlers') 
						}
						elseif ($moduleFileName -in @('reconnect.psm1', 'reconnect'))
						{ 
							$keyFunctionsToCapture = @('StartDisconnectWatcher','StopDisconnectWatcher','InvokeReconnectionSequence') 
						}
						elseif ($moduleFileName -in @('notifications.psm1', 'notifications'))
						{ 
							$keyFunctionsToCapture = @('UpdateNotificationPositions','CloseToast','ShowToast','ShowInteractiveNotification','ShowReconnectInteractiveNotification') 
						}
						elseif ($moduleFileName -in @('runspace-helper.psm1', 'runspace-helper'))
						{ 
							$keyFunctionsToCapture = @('NewManagedRunspace','InvokeInManagedRunspace','DisposeManagedRunspace') 
						}
						else
						{
							Write-Verbose "WARNING: No verification list found for module '$moduleFileName' (Length: $($moduleFileName.Length)). verification skipped." 
						}
						
						if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and -not $global:DashboardConfig.Resources.ContainsKey('CapturedFunctions'))
						{
							$global:DashboardConfig.Resources['CapturedFunctions'] = @{}
						}

						$captureSuccess = $true 
						$criticalFunctionMissing = $false

						if ($keyFunctionsToCapture.Count -gt 0)
						{
							Write-Verbose "- Attempt 3: Verifying and capturing key functions for '$moduleName' immediately after IEX..." 
							foreach ($funcName in $keyFunctionsToCapture)
							{
								$funcInfo = Get-Command -Name $funcName -CommandType Function -ErrorAction SilentlyContinue
								if ($funcInfo)
								{
									$capturedScriptBlock = $funcInfo.ScriptBlock
									Write-Verbose "  - Found and capturing ScriptBlock for '$funcName'." 
									
									if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and $global:DashboardConfig.Resources.CapturedFunctions)
									{
										$global:DashboardConfig.Resources.CapturedFunctions[$funcName] = $capturedScriptBlock
									}
									
									$functionsCapturedInThisAttempt[$funcName] = $capturedScriptBlock
								}
								else
								{
									Write-Verbose "  - WARNING: Could not find/capture function '$funcName' immediately after IEX for '$moduleName'." 
									$captureSuccess = $false
									
									
									$isCriticalModule = $moduleInfo.Priority -eq 'Critical' 
									
									if ($isCriticalModule)
									{ 
										$criticalFunctionMissing = $true
										$importErrorDetails += "; Critical function '$funcName' not found after IEX in Critical module '$moduleName'"
										Write-Verbose "    - Missing function '$funcName' is considered critical for module '$moduleName'." 
									}
									else
									{
										$importErrorDetails += "; Non-critical function '$funcName' not found after IEX for module '$moduleName'"
									}
									
								}
							}
						}

						
						
						
						if ($iexCompletedWithoutTerminatingError -and $captureSuccess -and (-not $criticalFunctionMissing))
						{
							Write-Verbose "  - Attempt 3: IEX completed and key functions captured/verified for '$moduleName'." 

							
							Write-Verbose "  - Defining captured functions globally for '$moduleName'..." 
							$definitionSuccess = $true 
							foreach ($kvp in $functionsCapturedInThisAttempt.GetEnumerator())
							{
								$funcNameToDefine = $kvp.Key
								$scriptBlockToDefine = $kvp.Value
								try
								{
									
									Set-Item -Path "Function:\global:$funcNameToDefine" -Value $scriptBlockToDefine -Force -ErrorAction Stop
									Write-Verbose "    - Defined Function:\global:$funcNameToDefine" 
								}
								catch
								{
									Write-Verbose "    - FAILED to define Function:\global:$funcNameToDefine globally: $($_.Exception.Message)" 
									$definitionSuccess = $false
									$importErrorDetails += "; Failed to define captured function '$funcNameToDefine' globally."
									
									
									if ($moduleInfo.Priority -eq 'Critical')
									{
										$result.CriticalFailure = $true
										Write-Verbose "    - Defining critical function '$funcNameToDefine' failed. Marking import as critical failure." 
									}
									
									break 
								}
							}

							
							if ($definitionSuccess)
							{
								$importSuccess = $true
								Write-Verbose "- [OK] Attempt 3: SUCCESS (Invoke-Expression completed, key functions captured AND globally defined for $moduleName)." 
							}
							else
							{
								$importSuccess = $false 
								Write-Verbose "- Attempt 3: FAILED during global definition phase for $moduleName." 
							}

						}
						else
						{
							
							Write-Verbose "- Attempt 3: FAILED (IEX completed=$iexCompletedWithoutTerminatingError, CaptureSuccess=$captureSuccess, CriticalMissing=$criticalFunctionMissing) for $moduleName." 
							if (-not $iexCompletedWithoutTerminatingError) { $importErrorDetails += "; IEX failed with non-terminating error detected by `$?."}
							if ($criticalFunctionMissing) { $importErrorDetails += '; Critical function missing prevented Attempt 3 success.' }
							if (-not $captureSuccess) { $importErrorDetails += '; Function capture failed during Attempt 3.' }
							$importSuccess = $false 
						}
					}
					catch 
					{
						Write-Verbose "- Attempt 3: FAILED (Invoke-Expression Error): $($_.Exception.Message)" 
						$importErrorDetails = "Invoke-Expression Error: $($_.Exception.Message)"
						$importSuccess = $false 
					}
				} 

				
				if ($importSuccess)
				{
					Write-Verbose "- [OK] Successfully imported PSM1 module: '$moduleName'." 
					
				}
				else
				{
					$errorMessage = "All import methods FAILED for PSM1 module: '$moduleName'. Last error detail: $importErrorDetails"
					Write-Verbose "- $errorMessage" 
					$failedModules[$moduleName] = $errorMessage
					
					if ($global:DashboardConfig -and $global:DashboardConfig.LoadedModules -and $global:DashboardConfig.LoadedModules.ContainsKey($moduleName))
					{
						Write-Verbose "- Removing '$moduleName' from LoadedModules list due to import failure." 
						$global:DashboardConfig.LoadedModules.Remove($moduleName)
					}
					
					if ($moduleInfo.Priority -eq 'Critical')
					{
						Write-Verbose "- CRITICAL FAILURE: Failed to import critical PSM1 module '$moduleName'." 
						$result.CriticalFailure = $true
					}
				}
			}
			#endregion SubStep: Import PowerShell Modules (.psm1)
		} 
		#endregion Step: Process Each Module in Sorted Order
				
		#region Step: Final Status Check and Result Construction
		Write-Verbose 'Module import check...' 
					
		
		if ($result.CriticalFailure)
		{
			Write-Verbose '  CRITICAL FAILURE: One or more critical modules failed to load or write. Application cannot continue.' 
			
			$criticalModules = $global:DashboardConfig.Modules.GetEnumerator() | Where-Object { $_.Value.Priority -eq 'Critical' }
			$failedCritical = $criticalModules | Where-Object { $failedModules.ContainsKey($_.Key) }
			if ($failedCritical)
			{
				Write-Verbose "  Failed critical modules: $($failedCritical.Key -join ', ')" 
				$failedCritical | ForEach-Object { Write-Verbose "  - $($_.Key): $($failedModules[$_.Key])"  } 
			}
			$result.Status = $false 
			
			return $result
		}
					
		
		$importantModules = $global:DashboardConfig.Modules.GetEnumerator() | Where-Object { $_.Value.Priority -eq 'Important' }
		$failedImportant = $importantModules | Where-Object { $failedModules.ContainsKey($_.Key) }
		if ($failedImportant.Count -gt 0)
		{
			Write-Verbose "  IMPORTANT module failures detected: $($failedImportant.Key -join ', '). Application may have limited functionality." 
			
			$failedImportant | ForEach-Object { Write-Verbose "  - $($_.Key): $($failedModules[$_.Key])"  } 
		}
					
		
		$optionalModules = $global:DashboardConfig.Modules.GetEnumerator() | Where-Object { $_.Value.Priority -eq 'Optional' }
		$failedOptional = $optionalModules | Where-Object { $failedModules.ContainsKey($_.Key) }
		if ($failedOptional.Count -gt 0)
		{
			Write-Verbose "  Optional module failures detected: $($failedOptional.Key -join ', '). Non-essential features might be unavailable." 
			
			$failedOptional | ForEach-Object { Write-Verbose "  - $($_.Key): $($failedModules[$_.Key])"  }
		}
					
		
		$successCount = 0
		if ($global:DashboardConfig -and $global:DashboardConfig.LoadedModules)
		{
			$successCount = $global:DashboardConfig.LoadedModules.Count
		}
		$failCount = $failedModules.Count
		Write-Verbose "  Module loading phase complete. Modules written/verified: $successCount. Failures (any type): $failCount." 
		if ($successCount -gt 0)
		{
			Write-Verbose "  Successfully written/verified modules: $($global:DashboardConfig.LoadedModules.Keys -join ', ')" 
		}
		if ($failCount -gt 0)
		{
			Write-Verbose '  Failed modules logged above.' 
		}
					
		
		$result.Status = $true
		$result.CriticalFailure = $false 
		
		return $result
		#endregion Step: Final Status Check and Result Construction
	}
	catch
	{
		
		$errorMessage = "  FATAL UNHANDLED EXCEPTION in Import-DashboardModules: $($_.Exception.Message)"
		Write-Verbose $errorMessage 
		Write-Verbose "  Stack trace: $($_.ScriptStackTrace)" 
		
		$result.Status = $false
		$result.CriticalFailure = $true
		$result.Exception = $_.Exception.Message 
		$failedModules['Unhandled Exception'] = $errorMessage 
		return $result
	}
}
#endregion Function: Import-DashboardModules

#endregion Module Handling Functions

#region UI and Application Lifecycle Functions

#region Function: StartDashboard
function StartDashboard
{
	<#
			.SYNOPSIS
				Initializes and displays the main dashboard user interface (UI) form.
			
			.DESCRIPTION
				This function orchestrates the startup of the application's graphical user interface. It performs these actions:
				1. Checks if the 'InitializeUI' function, expected to be loaded from the 'ui.psm1' module, exists using `Get-Command`. If not found, it throws a terminating error as the UI cannot be built.
				2. Calls the `InitializeUI` function. It assumes this function is responsible for creating all UI elements (forms, controls) and populating the '$global:DashboardConfig.UI' hashtable, including setting '$global:DashboardConfig.UI.MainForm'.
				3. Checks the return value of `InitializeUI`. If it returns $false or null (interpreted as failure), it throws a terminating error.
				4. Verifies that '$global:DashboardConfig.UI.MainForm' exists and is a valid '[System.Windows.Forms.Form]' object after `InitializeUI` returns successfully. If not, it throws a terminating error.
				5. If the MainForm is valid, it calls the `.Show()` method to make the main window visible and `.Activate()` to bring it to the foreground.
				6. Sets the global state flag '$global:DashboardConfig.State.UIInitialized' to $true.
			
			.OUTPUTS
				[bool] Returns $true if the UI is successfully initialized, the main form is found, shown, and activated.
				Returns $false if any step fails (missing function, initialization failure, missing main form), typically after throwing an error that gets caught by the main execution block.
			
			.NOTES
				- This function has a strong dependency on the 'ui.psm1' module being loaded correctly and functioning as expected (defining `InitializeUI` and creating `MainForm`).
				- Errors encountered during this process are considered fatal for the application and are thrown to be caught by the main script's try/catch block, which should then display an error using `ShowErrorDialog`.
		#>
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	Write-Verbose 'Starting Dashboard User Interface...' 
	try
	{
		#region Step: Check for and Call InitializeUI Function
		Write-Verbose '- Checking for required InitializeUI function (from ui.psm1)...' 
		
		if (-not (Get-Command InitializeUI -ErrorAction SilentlyContinue ))
		{
			
			throw "FATAL: InitializeUI function not found. Ensure 'ui.psm1' module loaded correctly and defines this function."
		}

		Write-Verbose '- Calling InitializeUI function...' 
		
		InitializeUI 

		Write-Verbose '- [OK] InitializeUI function executed successfully.' 
		#endregion Step: Check for and Call InitializeUI Function

		#region Step: Verify, Show, and Activate Main Form
		Write-Verbose '- Verifying presence and type of UI.MainForm object...' 
		
		if ($null -eq $global:DashboardConfig.UI.MainForm -or -not ($global:DashboardConfig.UI.MainForm -is [System.Windows.Forms.Form]))
		{
			
			throw 'FATAL: UI.MainForm object not found or is not a valid System.Windows.Forms.Form in $global:DashboardConfig after successful InitializeUI call.'
		}

		Write-Verbose '- [OK] UI.MainForm found and is valid. Showing and activating window...' 
		
		$global:DashboardConfig.UI.MainForm.Show() 
				
		
		$global:DashboardConfig.State.UIInitialized = $true
		Write-Verbose '  Dashboard UI started successfully.' 
		#endregion Step: Verify, Show, and Activate Main Form

		
		return $true
	}
	catch
	{
		$errorMsg = "  FATAL: Failed to start dashboard UI. Error: $($_.Exception.Message)"
		Write-Verbose $errorMsg 
		
		throw $_ 
	}
}
#endregion Function: StartDashboard
	
#region Function: StartMessageLoop
function StartMessageLoop
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
				- Requires the main UI form (`$global:DashboardConfig.UI.MainForm`) to be successfully initialized and shown by `StartDashboard` before being called.
				- The efficiency of the UI heavily depends on the availability and correctness of the 'Native' class methods from 'classes.psm1'. The `DoEvents` fallback is less performant.
				- Includes basic error handling within the loop itself and a final `DoEvents` fallback attempt if the primary loop method encounters an unhandled exception.
				- Logs the chosen loop method and status messages during execution and upon exit.
		#>
	[CmdletBinding()]
	[OutputType([void])]
	param()
			
	Write-Verbose "`Starting UI message loop..." 
			
	#region Step: Pre-Loop Checks for UI State and Main Form Validity
	Write-Verbose '  Checking UI state before starting message loop...' 
	
	if (-not $global:DashboardConfig.State.UIInitialized)
	{
		Write-Verbose "  UI not marked as initialized ($global:DashboardConfig.State.UIInitialized is $false). Skipping message loop." 
		return 
	}
	
	$mainForm = $global:DashboardConfig.UI.MainForm 
	if ($null -eq $mainForm -or -not ($mainForm -is [System.Windows.Forms.Form]))
	{
		Write-Verbose "  MainForm object ($global:DashboardConfig.UI.MainForm) is missing or not a valid Form object. Cannot start message loop." 
		return 
	}
	if ($mainForm.IsDisposed)
	{
		Write-Verbose "  MainForm ($global:DashboardConfig.UI.MainForm) is already disposed. Cannot start message loop." 
		return 
	}
	Write-Verbose '  Pre-loop checks passed. MainForm is valid and UI is initialized.' 
	#endregion Step: Pre-Loop Checks for UI State and Main Form Validity
			
	
	[string]$loopMethod = 'Unknown'
	try
	{
		#region Step: Determine Loop Method (Efficient Native P/Invoke vs. Fallback DoEvents)
		
		[bool]$useNativeLoop = $false
		Write-Verbose 'Detecting availability of Native methods for efficient loop...' 
		try
		{
			
			
			$nativeType = [type]'Custom.Native' 
			if (($nativeType.GetMethod('AsyncExecution')) -and
				($nativeType.GetMethod('PeekMessage')) -and
				($nativeType.GetMethod('TranslateMessage')) -and
				($nativeType.GetMethod('DispatchMessage')))
			{
				Write-Verbose "- [OK] Native P/Invoke methods found (requires 'classes.psm1'). Using efficient message loop." 
				$useNativeLoop = $true
				$loopMethod = 'Custom.Native'
			}
			else
			{
				Write-Verbose '- Native class found, but required methods (AsyncExecution, PeekMessage, etc.) are missing. Falling back to DoEvents loop.' 
				$loopMethod = 'DoEvents'
			}
		}
		catch [System.Management.Automation.RuntimeException]
		{
			
			Write-Verbose '- Native class not found. Falling back to less efficient Application.DoEvents() loop.' 
			$loopMethod = 'DoEvents'
		}
		catch
		{
			Write-Verbose "- Error checking for Native methods: $($_.Exception.Message). Falling back to DoEvents loop." 
			$loopMethod = 'DoEvents'
		}
					
		
		if (-not $useNativeLoop)
		{
			
			InitializeClassesModule
		}
		#endregion Step: Determine Loop Method (Efficient Native P/Invoke vs. Fallback DoEvents)
				
		#region Step: Run the Chosen Message Loop
		Write-Verbose "Entering message loop (Method: $loopMethod). Loop runs until main form is closed..." 
		
		
		while ($mainForm -and $mainForm.Visible -and -not $mainForm.IsDisposed)
		{
			if ($useNativeLoop)
			{
				
				try
				{
					
					
					
					$result = [Custom.Native]::AsyncExecution(0, [IntPtr[]]@(), $false, 50, [Custom.Native]::QS_ALLINPUT) 
								
					
					if ($result -ne 0x102) 
					{
						
						
						$msg = New-Object Custom.Native+MSG
						
						while ([Custom.Native]::PeekMessage([ref]$msg, [IntPtr]::Zero, 0, 0, [Custom.Native]::PM_REMOVE))
						{
							
							$null = [Custom.Native]::TranslateMessage([ref]$msg)
							
							$null = [Custom.Native]::DispatchMessage([ref]$msg)
						}
					}
					
				}
				catch
				{
					
					Write-Verbose "  Error during Native message loop iteration: $($_.Exception.Message). Attempting to fall back to DoEvents..." 
					$useNativeLoop = $false 
					$loopMethod = 'DoEvents'
					
					InitializeClassesModule
					Start-Sleep -Milliseconds 50 
				}
			}
			else 
			{
				
				try
				{
					
					[System.Windows.Forms.Application]::DoEvents()
					
					Start-Sleep -Milliseconds 20 
				}
				catch
				{
					
					Write-Verbose "  Error during DoEvents fallback loop iteration: $($_.Exception.Message). Loop may become unresponsive." 
					
					Start-Sleep -Milliseconds 100
				}
			}
		} 
		#endregion Step: Run the Chosen Message Loop
	}
	catch
	{
		
		Write-Verbose "  FATAL Error occurred within the UI message loop setup or main structure: $($_.Exception.Message)" 
		
		Write-Verbose '  Attempting basic DoEvents fallback loop after critical error...' 
		try
		{
			if ($mainForm -and -not $mainForm.IsDisposed)
			{
				
				InitializeClassesModule
			}
					
			while ($mainForm -and $mainForm.Visible -and -not $mainForm.IsDisposed)
			{
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Milliseconds 50 
			}
		}
		catch
		{
			Write-Verbose "  Emergency fallback DoEvents loop also failed: $($_.Exception.Message)" 
			[System.Windows.Forms.Application]::Run($mainForm)
		}
	}
	finally
	{
		
		
		if ($mainForm -and ($mainForm -is [System.Windows.Forms.Form]))
		{
			Write-Verbose "UI message loop exited (Method: $loopMethod). Final Form State -> Visible: $($mainForm.Visible), Disposed: $($mainForm.IsDisposed)" 
		}
		else
		{
			Write-Verbose "UI message loop exited (Method: $loopMethod). MainForm object appears invalid or null upon exit." 
		}
		
		$global:DashboardConfig.State.UIInitialized = $false
	}
}
#endregion Function: StartMessageLoop
	
#region Function: StopDashboard
function StopDashboard
{
	<#
			.SYNOPSIS
				Performs comprehensive cleanup of application resources during shutdown.
			
			.DESCRIPTION
				This function is responsible for gracefully stopping and releasing all resources allocated by the application
				and its modules. It's designed to be called within the main script's `finally` block to ensure cleanup
				happens reliably, even if errors occurred during execution.
				
				Cleanup is performed in a specific order to minimize dependency issues and errors:
				1.  **Ftool Forms:** If the optional 'ftool.psm1' module was loaded and created forms (tracked in `$global:DashboardConfig.Resources.FtoolForms`), it attempts to close and dispose of them. It preferably calls a `StopFtoolForm` function (if defined by ftool.psm1) for module-specific cleanup before falling back to basic `.Close()` and `.Dispose()` calls.
				2.  **Timers:** Stops and disposes of all `System.Windows.Forms.Timer` objects registered in `$global:DashboardConfig.Resources.Timers`. Handles nested collections if necessary.
				3.  **Main UI Form:** Disposes of the main application window (`$global:DashboardConfig.UI.MainForm`) if it exists and isn't already disposed.
				4.  **Runspaces & Module Cleanup:**
				*   Disposes of known background runspaces (e.g., `$global:DashboardConfig.Resources.LaunchResources` if used by 'launch.psm1').
				*   Calls specific cleanup functions (e.g., `StopClientLaunch`, `CleanupLogin`, `CleanupFtool`) if they exist (assumed to be defined by the respective modules). These functions are expected to handle module-specific resource release (e.g., closing handles, stopping threads).
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
			
	Write-Verbose 'Stopping Dashboard and Cleaning Up Application Resources...' 
	
	[bool]$cleanupOverallSuccess = $true

	if (Get-Command RestoreWindowStyles -ErrorAction SilentlyContinue) { RestoreWindowStyles } else {	Write-Verbose '  RestoreWindowStyles function not found. Restart Dashboard to enable hidden windows again.'  }

	#region Step 1: Clean Up launch recources
	Write-Verbose 'Step 1: Cleaning up Launch...' 
	if (Get-Command StopClientLaunch -ErrorAction SilentlyContinue)
	{
		StopClientLaunch
	}
	else
	{
		Write-Verbose '  StopClientLaunch function not found. Skipping launch cleanup.' 
	}
	#endregion Step 1: Clean Up launch recources

	#region Step 2: Clean Up Disconnect Watcher
	Write-Verbose 'Step 2: Cleaning up Disconnect Watcher...' 
	if (Get-Command StopDisconnectWatcher -ErrorAction SilentlyContinue)
	{
		StopDisconnectWatcher
	}
	else
	{
		Write-Verbose '  StopDisconnectWatcher function not found. Skipping watcher cleanup.' 
	}
	#endregion Step 2: Clean Up Disconnect Watcher
			
	#region Step 3: Clean Up Ftool Forms (if Ftool module was loaded/used)
	Write-Verbose 'Step 3: Cleaning up Ftool forms...' 
	try
	{
		
		$ftoolForms = $global:DashboardConfig.Resources.FtoolForms
		if ($ftoolForms -and $ftoolForms.Count -gt 0)
		{
			
			
			$stopFtoolFormCmd = Get-Command -Name StopFtoolForm -ErrorAction SilentlyContinue
			
			
			[string[]]$formKeys = @($ftoolForms.Keys)
			Write-Verbose "- Found $($formKeys.Count) Ftool form(s) registered. Attempting cleanup..." 
						
			foreach ($key in $formKeys)
			{
				
				
				$form = $ftoolForms[$key]
				
				if ($form -and $form -is [System.Windows.Forms.Form] -and -not $form.IsDisposed)
				{
					$formText = try
					{
						$form.Text 
					}
					catch
					{
						'(Error getting text)' 
					} 
					Write-Verbose "  - Stopping Ftool form '$formText' (Key: $key)." 
					try
					{
						
						if ($stopFtoolFormCmd)
						{
							Write-Verbose '  - Using StopFtoolForm function...' 
							StopFtoolForm -Form $form -ErrorAction Stop 
						}
						else 
						{
							Write-Verbose "  - StopFtoolForm command not found. Performing basic Close() for form '$formText'." 
							
							$form.Close()
							
							Start-Sleep -Milliseconds 20
						}
					}
					catch 
					{
						Write-Verbose "  - Error during StopFtoolForm or Close() for form '$formText': $($_.Exception.Message)" 
						
						$cleanupOverallSuccess = $false
					}
					finally 
					{
						Write-Verbose "  - Ensuring Dispose() is called for form '$formText'." 
						try
						{
							if (-not $form.IsDisposed)
							{
								$form.Dispose() 
							}
						}
						catch
						{
							Write-Verbose "  - Error during final Dispose() for form '$formText': $($_.Exception.Message)" 
							$cleanupOverallSuccess = $false
						}
					}
				}
				elseif ($form -and $form -is [System.Windows.Forms.Form] -and $form.IsDisposed)
				{
					Write-Verbose "  - Ftool form with Key '$key' was already disposed." 
				}
				else
				{
					Write-Verbose "  - Ftool form entry with Key '$key' is null or not a valid Form object." 
					$cleanupOverallSuccess = $false
				}
							
				
				$ftoolForms.Remove($key) | Out-Null
			} 
			Write-Verbose '- Finished Ftool form cleanup.' 
		}
		else
		{
			Write-Verbose '  No active Ftool forms found in configuration to clean up.'  
		}
	}
	catch 
	{
		Write-Verbose "Error during Ftool form cleanup phase setup: $($_.Exception.Message)" 
		$cleanupOverallSuccess = $false
	}
	#endregion Step 3: Clean Up Ftool Forms (if Ftool module was loaded/used)
			
	#region Step 4: Clean Up Application Timers
	Write-Verbose 'Step 4: Cleaning up application timers...' 
	try
	{
		
		$timersCollection = $global:DashboardConfig.Resources.Timers
		if ($timersCollection -and $timersCollection.Count -gt 0)
		{
			Write-Verbose "- Found $($timersCollection.Count) timer registration(s). Stopping and disposing..." 
			
			
			[System.Collections.Generic.List[System.Windows.Forms.Timer]]$uniqueTimers = New-Object System.Collections.Generic.List[System.Windows.Forms.Timer]
						
			
			
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
					
					$item.Values | Where-Object { $_ -is [System.Windows.Forms.Timer] } | ForEach-Object {
						if (-not $uniqueTimers.Contains($_))
						{
							$uniqueTimers.Add($_) 
						}
					}
				}
				
			}
			Write-Verbose "- Found $($uniqueTimers.Count) unique System.Windows.Forms.Timer object(s) to dispose." 
						
			
			foreach ($timer in $uniqueTimers)
			{
				try
				{
					
					if ($timer -and -not $timer.IsDisposed) 
					{
						Write-Verbose "  - Disposing timer (Was Enabled: $($timer.Enabled))." 
						
						if ($timer.Enabled)
						{
							$timer.Stop() 
						}
						
						$timer.Dispose()
					}
					else
					{
						Write-Verbose '  - Skipping already disposed or invalid timer object.' 
					}
				}
				catch 
				{
					Write-Verbose "  - Error stopping or disposing a timer: $($_.Exception.Message)" 
					$cleanupOverallSuccess = $false 
				}
			} 
						
			
			Write-Verbose '- Clearing global timer registration collection.' 
			$timersCollection.Clear()
			Write-Verbose '- Finished timer cleanup.' 
		}
		else
		{
			Write-Verbose '- No active timers found in configuration to clean up.'  
		}
	}
	catch 
	{
		Write-Verbose "Error during timer cleanup phase setup: $($_.Exception.Message)" 
		$cleanupOverallSuccess = $false
	}
	#endregion Step 4: Clean Up Application Timers
			
	#region Step 5: Clean Up Main UI Form
	Write-Verbose 'Step 5: Cleaning up main UI form...' 
	try
	{
		
		$mainForm = $global:DashboardConfig.UI.PSObject.Properties['MainForm']
		if ($mainForm -and $mainForm.Value -is [System.Windows.Forms.Form] -and -not $mainForm.Value.IsDisposed)
		{
			Write-Verbose '- Disposing MainForm object...' 
			
			$mainForm.Value.Dispose()
			Write-Verbose '- [OK] MainForm disposed.' 
		}
		elseif ($mainForm -and $mainForm.Value -is [System.Windows.Forms.Form] -and $mainForm.Value.IsDisposed)
		{
			Write-Verbose '- MainForm was already disposed.' 
		}
		else
		{
			Write-Verbose '- MainForm object not found or invalid in configuration.' 
		}
	}
	catch 
	{
		Write-Verbose "Error disposing main UI form: $($_.Exception.Message)" 
		$cleanupOverallSuccess = $false
	}
	#endregion Step 5: Clean Up Main UI Form
			
	#region Step 6: Reset Application State Flags
	Write-Verbose 'Step 6: Resetting application state flags...' 
	try
	{
		
		$global:DashboardConfig.State.UIInitialized = $null
		$global:DashboardConfig.State.LoginActive = $null
		$global:DashboardConfig.State.LaunchActive = $null
		$global:DashboardConfig.State.ConfigInitialized = $null
		$global:DashboardConfig.LoadedModules = $null
		Write-Verbose '- State flags reset.' 
	}
	catch 
	{
		Write-Verbose "  Error resetting global state flags: $($_.Exception.Message)" 
		
		$cleanupOverallSuccess = $false
	}
	#endregion Step 6: Reset Application State Flags
				
	#region Step 7: Final Log Message for Cleanup Status
	Write-Verbose "--- Dashboard Cleanup Finished. Overall Success: $cleanupOverallSuccess ---"
	#endregion Step 7: Final Log Message
			
	
	return $cleanupOverallSuccess
}
#endregion Function: StopDashboard

#endregion UI and Application Lifecycle Functions

#region Main Execution Block


try
{
	
	$LogDir = Split-Path -Path $global:DashboardConfig.Paths.Verbose -Parent
	if (-not (Test-Path $LogDir -PathType Container)) { New-Item $LogDir -ItemType Directory -Force | Out-Null }
	
	$global:DashboardConfig.Paths.Verbose | ForEach-Object {
		if (Test-Path $_ -PathType Leaf) { Clear-Content $_ -ErrorAction SilentlyContinue }
	}
	$VerbosePreference = 'Continue' 
}
catch
{
	Write-Warning "Initial log setup failed: $($_.Exception.Message)"
}


$TS = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Verbose '=========================================' 
Write-Verbose '=== Initializing Entropia Dashboard ===' 
Write-Verbose "=== Timestamp: $TS ===" 
Write-Verbose '=========================================' 


try
{
	
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

	
	$updateSplash = {
		param([string]$Text, [int]$Percentage)
		if ($splashForm -and -not $splashForm.IsDisposed)
		{
			$splashForm.Invoke([Action] {
					$statusLabel.Text = "$Text..."
					$progressBar.Value = [Math]::Min(100, $Percentage)
				})
			$splashForm.Refresh()
		}
	}
	

	
	Write-Verbose '--- Step 1: Ensuring Correct Execution Environment ---' 
	& $updateSplash 'Verifying execution environment' 10
	RequestElevation 
	if (-not (InitializeScriptEnvironment)) { throw 'Environment verification failed. Cannot continue.' }
	Write-Verbose '[OK] Environment verified.' 
        
	
	Write-Verbose '--- Step 2: Initializing Base Configuration (AppData Paths) ---' 
	& $updateSplash 'Initializing configuration' 20
	if (-not (InitializeBaseConfig)) { throw 'Failed to initialize base application paths. Cannot continue.' }
	Write-Verbose '[OK] Base configuration paths initialized.' 
        
	
	Write-Verbose '--- Step 2.5: Checking for script updates ---' 
	& $updateSplash 'Checking for updates' 30
	try
	{
		
		$executablePath = $null
		if (-not [string]::IsNullOrEmpty($PSCommandPath))
		{
			$executablePath = $PSCommandPath
		}
		elseif (-not [string]::IsNullOrEmpty($MyInvocation.MyCommand.Path))
		{
			$executablePath = $MyInvocation.MyCommand.Path
		}
		else
		{
			
			try
			{
				$executablePath = (Get-Process -Id $PID -ErrorAction Stop).Path
			}
			catch
			{
				
			}
		}

		
		if ([string]::IsNullOrEmpty($executablePath))
		{
			Write-Verbose '  Could not determine script/EXE path. Update check will be skipped.' 
		}
		else
		{
			Write-Verbose "  Checking version of: $executablePath" 
			
			$localVersion = '0.0'
			$localContent = Get-Content -Path $executablePath -Raw -ErrorAction SilentlyContinue
			if ($localContent)
			{
				$lm = [regex]::Match($localContent, '(?m)^\s*Version:\s*([0-9]+(?:\.[0-9]+)*)')
				if ($lm.Success) { $localVersion = $lm.Groups[1].Value }
			}
			
			$remoteUrl = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/version'
			$remoteContent = $null
			try
			{
				$remoteContent = (Invoke-WebRequest -Uri $remoteUrl -UseBasicParsing -ErrorAction Stop).Content
			}
			catch
			{
				Write-Verbose "  Update check failed to fetch remote script: $($_.Exception.Message)" 
			}

			if ($remoteContent)
			{
				
				$rm = [regex]::Match($remoteContent, '(?m)^\s*Version:\s*([0-9]+(?:\.[0-9]+)*)')
					
				if ($rm.Success)
				{
					$remoteVersion = $rm.Groups[1].Value

					
					$changelogDisplay = ''
					$cm = [regex]::Match($remoteContent, '(?is)\[Changelogs\]\s*(.*)$')
					if ($cm.Success)
					{
						$extractedLog = $cm.Groups[1].Value.Trim()
						if (-not [string]::IsNullOrWhiteSpace($extractedLog))
						{
							
							if ($extractedLog.Length -gt 2000)
							{ 
								$extractedLog = $extractedLog.Substring(0, 2000) + '...(truncated)' 
							}
							$changelogDisplay = "`n`n=== What's New ===`n$extractedLog"
						}
					}

					try
					{
						if ([version]$localVersion -lt [version]$remoteVersion)
						{
							Write-Verbose "  Update available: local $localVersion < remote $remoteVersion" 
								
							
							$promptMsg = "A newer version ($remoteVersion) of Entropia Dashboard is available.`nYou have $localVersion.$changelogDisplay`n`nPress 'Yes' to download the update and restart automatically.`nPress 'No' to continue without updating."
							$caption = 'Entropia Dashboard - Update Available'
								
							$resp = [System.Windows.Forms.MessageBox]::Show($promptMsg, $caption, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
								
							if ($resp -eq [System.Windows.Forms.DialogResult]::Yes)
							{
								Write-Verbose '  User chose to update. Starting download...' 
									
								
									
								
								
								$updateUrl = 'https://github.com/Immortal-Divine/Entropia_Dashboard/raw/main/Entropia%20Dashboard.exe'
								$tempExePath = "$executablePath.new"
								$batchPath = Join-Path $env:TEMP ('dashboard_updater_' + [Guid]::NewGuid().ToString() + '.bat')

								
								try
								{
									& $updateSplash 'Downloading update' 35
									
									$webClient = New-Object System.Net.WebClient
									$webClient.DownloadFile($updateUrl, $tempExePath)
									Write-Verbose "  Download complete: $tempExePath" 
								}
								catch
								{
									$errMsg = "Failed to download update. Please download manually from GitHub.`nError: $($_.Exception.Message)"
									Write-Verbose "  $errMsg" 
									[System.Windows.Forms.MessageBox]::Show($errMsg, 'Update Failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
									
									Start-Process 'https://github.com/Immortal-Divine/Entropia_Dashboard'
									return 
								}

								
								
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

								
								Write-Verbose '  Starting updater script and exiting...' 
								& $updateSplash 'Restarting to apply update' 40
								Start-Process -FilePath $batchPath -WindowStyle Hidden
									
								
								return
							}
							else
							{
								Write-Verbose '  User chose to continue without updating.' 
							}
						}
						else
						{
							Write-Verbose "  Local script version ($localVersion) is up-to-date (remote $remoteVersion)." 
						}
					}
					catch
					{
						Write-Verbose "  Version compare failed: $($_.Exception.Message)" 
					}
				}
				else
				{
					Write-Verbose '  Remote script did not contain a version header.' 
				}
			}
		}
	}
	catch
	{
		Write-Verbose "  Unexpected error during update check: $($_.Exception.Message)" 
	}

	
	Write-Verbose '--- Step 3: Loading Dashboard Modules ---' 
	& $updateSplash 'Loading core modules' 50
	$importResult = Import-DashboardModules
	if (-not $importResult.Status) { throw 'Critical module loading failed. Cannot continue.' }
	Write-Verbose '[OK] Core modules loaded successfully.' 
        
	
	Write-Verbose '--- Step 4: Loading INI Configuration ---' 
	& $updateSplash 'Loading saved settings' 70
	if (Get-Command InitializeIniConfig -ErrorAction SilentlyContinue)
	{
		try
		{
			if (-not (InitializeIniConfig -ErrorAction Stop)) { Write-Verbose 'InitializeIniConfig failed. Defaults used.'  }
			else { Write-Verbose '[OK] INI configuration loaded successfully.'  }
		}
		catch { Write-Verbose "Error during InitializeIniConfig: $($_.Exception.Message). Defaults used."  }
	}
	else { Write-Verbose 'InitializeIniConfig not found. Skipping INI load.'  }
        
	
	Write-Verbose '--- Step 5: Starting Dashboard UI ---' 
	& $updateSplash 'Building main user interface' 90
	if (-not (StartDashboard)) { throw 'StartDashboard returned failure. Cannot continue.' }
	Write-Verbose '[OK] Dashboard UI started.' 
	if (Get-Command StartDataGridUpdateTimer -ErrorAction SilentlyContinue)
	{
		StartDataGridUpdateTimer; Write-Verbose '[OK] DataGrid update timer started.' 
	}
	else { Write-Verbose '[WARN] StartDataGridUpdateTimer not found.'  }

	& $updateSplash 'Finalizing' 100
	Start-Sleep -Milliseconds 250 
	if ($splashForm -and -not $splashForm.IsDisposed) { $splashForm.Close() }

	
	Write-Verbose '--- Step 6: Running UI Message Loop ---' 
	$handle = $global:DashboardConfig.UI.MainForm.Handle
	if ($handle -ne [IntPtr]::Zero)
	{
		Start-Sleep -Milliseconds 200
		if ([Custom.Native]::IsWindowMinimized($handle)) { [Custom.Native]::ShowWindow($handle, 9) | Out-Null }
		if (-not ([Custom.Native]::SetForegroundWindow($handle)))
		{
			Write-Verbose 'SetForegroundWindow failed. Trying Alt-key workaround...' 
			try
			{ 
				[Custom.Native]::keybd_event(0x12, 0, 1, 0); Start-Sleep -Milliseconds 50
				[Custom.Native]::keybd_event(0x12, 0, 3, 0); Start-Sleep -Milliseconds 100
				if (-not ([Custom.Native]::SetForegroundWindow($handle)))
				{
					Write-Verbose 'Alt workaround failed. Using Activate().' 
					$global:DashboardConfig.UI.MainForm.Activate()
				}
			}
			catch { Write-Verbose "Error in Alt-key simulation: $_"  }
		}
	}

	
	Write-Verbose '--- Step 7: Updating Reconnect Supervisor ---' 
	if (Get-Command StartDisconnectWatcher -ErrorAction SilentlyContinue)
	{
		StartDisconnectWatcher
	}
	else
	{
		Write-Verbose '[WARN] StartDisconnectWatcher not found.' 
	}
	StartMessageLoop

	Write-Verbose 'UI Message loop finished. Proceeding to final cleanup...' 
}
catch
{
	
	if ($splashForm -and -not $splashForm.IsDisposed) { $splashForm.Dispose() }
	$errorMessage = "`nFATAL UNHANDLED ERROR: $($_.Exception.Message)"
	Write-Verbose $errorMessage 
	try { ShowErrorDialog ($errorMessage + "`n`nStack Trace:`n" + $($_.ScriptStackTrace)) }
	catch { Write-Verbose "Failed to show final error dialog. The critical error was: $errorMessage"  }
}
finally
{
	
	Write-Verbose '--- Step 7: Entering Final Application Cleanup ---' 
	if ($splashForm -and -not $splashForm.IsDisposed) { $splashForm.Dispose() }
	if (Get-Command RemoveAllHotkeys -ErrorAction SilentlyContinue) { RemoveAllHotkeys; Write-Verbose 'Hotkeys unregistered.'  }
	if (Get-Command StopDashboard -ErrorAction SilentlyContinue)
	{
		$cleanupStatus = StopDashboard; Write-Verbose "[OK] StopDashboard completed (Success: $cleanupStatus)." 
	}
	else
	{
		Write-Verbose 'StopDashboard not found. Attempting basic MainForm dispose...' 
		try
		{
			$mainForm = $global:DashboardConfig.UI.MainForm
			if ($mainForm -is [System.Windows.Forms.Form] -and -not $mainForm.IsDisposed) { $mainForm.Dispose() }
		}
		catch { Write-Verbose "Fallback MainForm dispose failed: $($_.Exception.Message)"  }
	}
    
	if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
	{
		try { [System.Windows.Forms.Application]::ExitThread() }
		catch { Write-Verbose "Error calling ExitThread(): $($_.Exception.Message)"  }
	}
    
	Write-Verbose '=========================================' 
	Write-Verbose '=== Entropia Dashboard Exited ===' 
	Write-Verbose '=========================================' 

}

#endregion Main Execution Block