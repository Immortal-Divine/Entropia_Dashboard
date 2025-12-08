<# ui.psm1
    .SYNOPSIS
        User Interface Manager for Entropia Dashboard.
    .DESCRIPTION
        This module creates and manages the complete user interface for Entropia Dashboard:
        - Builds the main application window and all dialog forms
        - Creates interactive controls (buttons, panels, grids, text boxes)
        - Handles window dragging, resizing, and positioning
        - Manages client process monitoring display
        - Implements settings management through visual interface
        - Maintains responsive layout across different screen sizes
        - Provides Launch / Login / Ftool automation
    .NOTES
        Author: Immortal / Divine
        Version: 1.2.1
        Requires: PowerShell 5.1+, .NET Framework 4.5+, classes.psm1, ini.psm1, datagrid.psm1

        Documentation Standards Followed:
        - Module Level Documentation: Synopsis, Description, Notes.
        - Function Level Documentation: Synopsis, Parameter Descriptions, Output Specifications.
        - Code Organization: Logical grouping using #region / #endregion. Functions organized by workflow.
        - Step Documentation: Code blocks enclosed in '#region Step: Description' / '#endregion Step: Description'.
        - Variable Definitions: Inline comments describing the purpose of significant variables.
        - Error Handling: Comprehensive try/catch/finally blocks with error logging and user notification.

        This module relies heavily on the global $global:DashboardConfig object for state and configuration.
#>

#region Helper Functions
    #region Function: Sync-UIToConfig
	function Sync-UIToConfig
	{
		<#
		.SYNOPSIS
			Synchronizes the current state of UI input elements to the global configuration object.
		.OUTPUTS
			[bool] Returns $true if synchronization was successful, $false otherwise.
		.NOTES
			Reads values from UI controls (TextBoxes, ComboBoxes) and updates the corresponding
			keys in the $global:DashboardConfig.Config hashtable. Ensures necessary sections exist.
		#>
		[CmdletBinding()]
		[OutputType([bool])]
		param()

		#region Step: Attempt to sync UI state to global config
			try
			{
				#region Step: Log Sync Start
					Write-Verbose '  UI: Syncing UI to config' -ForegroundColor Cyan
				#endregion Step: Log Sync Start

				#region Step: Validate UI and Config objects
					$UI = $global:DashboardConfig.UI
					if (-not ($UI -and $global:DashboardConfig.Config))
					{
						Write-Verbose '  UI: UI or Config object not found, cannot sync.' -ForegroundColor Yellow
						return $false
					}
				#endregion Step: Validate UI and Config objects

				#region Step: Ensure required config sections exist
					# Ensure sections exist in the config hashtable before attempting to write to them.
					@('LauncherPath', 'ProcessName', 'MaxClients', 'Login') | ForEach-Object {
						if (-not $global:DashboardConfig.Config.Contains($_))
						{
							$global:DashboardConfig.Config[$_] = [ordered]@{}
						}
					}
				#endregion Step: Ensure required config sections exist

				#region Step: Sync basic UI control values to config
					# Read values from TextBoxes and update the config.
					$global:DashboardConfig.Config['LauncherPath']['LauncherPath'] = $UI.InputLauncher.Text
					$global:DashboardConfig.Config['ProcessName']['ProcessName'] = $UI.InputProcess.Text
					$global:DashboardConfig.Config['MaxClients']['MaxClients'] = $UI.InputMax.Text
				#endregion Step: Sync basic UI control values to config

				#region Step: Sync login position ComboBox selections to config
					# Collect selected items from login position ComboBoxes.
					$loginPos = @()
					$UI.Login.Keys | Sort-Object { [int]($_ -replace 'Login', '') } | ForEach-Object {
						$combo = $UI.Login[$_]
						$loginPos += if ($combo.SelectedItem)
						{
							$combo.SelectedItem
						}
						else
						{
							'1' # Default to '1' if nothing is selected
						}
					}
					# Store the collected positions as a comma-separated string in the config.
					$global:DashboardConfig.Config['Login']['Login'] = $loginPos -join ','
				#endregion Step: Sync login position ComboBox selections to config

				#region Step: Sync finalize collector login checkbox to config
					# Read checkbox state and store as string (0 for false, 1 for true).
					$finalizeLoginValue = if ($UI.FinalizeCollectorLogin.Checked) { '1' } else { '0' }
					$global:DashboardConfig.Config['Login']['FinalizeCollectorLogin'] = $finalizeLoginValue
				#endregion Step: Sync finalize collector login checkbox to config

				#region Step: Sync NeverRestarting collector login checkbox to config
					# Read checkbox state and store as string (0 for false, 1 for true).
					$NeverRestartingLoginValue = if ($UI.NeverRestartingCollectorLogin.Checked) { '1' } else { '0' }
					$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin'] = $NeverRestartingLoginValue
				#endregion Step: Sync NeverRestarting collector login checkbox to config

				#region Step: Sync HideMinimizedWindows checkbox to config
					# Read checkbox state and store as string (0 for false, 1 for true).
					$hideMinimizedWindowsValue = if ($UI.HideMinimizedWindows.Checked) { '1' } else { '0' }
					$global:DashboardConfig.Config['Options']['HideMinimizedWindows'] = $hideMinimizedWindowsValue
				#endregion Step: Sync HideMinimizedWindows checkbox to config

				#region Step: Log Sync Success
					Write-Verbose '  UI: UI synced to config' -ForegroundColor Green
					return $true
				#endregion Step: Log Sync Success
			}
			catch
			{
				#region Step: Handle errors during sync
					Write-Verbose "  UI: Failed to sync UI to config: $_" -ForegroundColor Red
					return $false
				#endregion Step: Handle errors during sync
			}
		#endregion Step: Attempt to sync UI state to global config
	}
#endregion Function: Sync-UIToConfig

#region Function: Sync-ConfigToUI
	function Sync-ConfigToUI
	{
		<#
		.SYNOPSIS
			Synchronizes the global configuration object values to the UI elements.
		.OUTPUTS
			[bool] Returns $true if synchronization was successful, $false otherwise.
		.NOTES
			Reads values from the $global:DashboardConfig.Config hashtable and updates the
			corresponding UI controls (TextBoxes, ComboBoxes). Handles cases where config values might be missing.
		#>
		[CmdletBinding()]
		[OutputType([bool])]
		param()

		#region Step: Attempt to sync global config to UI state
			try
			{
				#region Step: Log Sync Start
					Write-Verbose '  UI: Syncing config to UI' -ForegroundColor Cyan
				#endregion Step: Log Sync Start

				#region Step: Validate UI and Config objects
					$UI = $global:DashboardConfig.UI
					if (-not ($UI -and $global:DashboardConfig.Config))
					{
						Write-Verbose '  UI: UI or Config object not found, cannot sync.' -ForegroundColor Yellow
						return $false
					}
				#endregion Step: Validate UI and Config objects

				#region Step: Sync LauncherPath from config to UI
					# Update Launcher Path TextBox if the config value exists.
					if ($global:DashboardConfig.Config['LauncherPath']['LauncherPath'])
					{
						$UI.InputLauncher.Text = $global:DashboardConfig.Config['LauncherPath']['LauncherPath']
					}
				#endregion Step: Sync LauncherPath from config to UI

				#region Step: Sync ProcessName from config to UI
					# Update Process Name TextBox if the config value exists.
					if ($global:DashboardConfig.Config['ProcessName']['ProcessName'])
					{
						$UI.InputProcess.Text = $global:DashboardConfig.Config['ProcessName']['ProcessName']
					}
				#endregion Step: Sync ProcessName from config to UI

				#region Step: Sync MaxClients from config to UI
					# Update Max Clients TextBox if the config value exists.
					if ($global:DashboardConfig.Config['MaxClients']['MaxClients'])
					{
						$UI.InputMax.Text = $global:DashboardConfig.Config['MaxClients']['MaxClients']
					}
				#endregion Step: Sync MaxClients from config to UI

				#region Step: Sync login position config to ComboBox selections
					# Update Login Position ComboBoxes based on the comma-separated config string.
					if ($global:DashboardConfig.Config['Login']['Login'])
					{
						$positions = $global:DashboardConfig.Config['Login']['Login'] -split ','

						# Iterate through ComboBoxes and set selected item based on config.
						for ($i = 0; $i -lt [Math]::Min($UI.Login.Count, $positions.Count); $i++)
						{
							$key = "Login$($i+1)"
							$value = $positions[$i]
							$combo = $UI.Login[$key]

							# Set selected item only if the ComboBox exists and the value is valid.
							if ($combo -and $combo.Items.Contains($value))
							{
								$combo.SelectedItem = $value
							}
						}
					}
				#endregion Step: Sync login position config to ComboBox selections

				#region Step: Sync finalize collector login config to checkbox
					# Update FinalizeCollectorLogin CheckBox if the config value exists.
					if ($global:DashboardConfig.Config['Login']['FinalizeCollectorLogin'])
					{
						$UI.FinalizeCollectorLogin.Checked = ([int]$global:DashboardConfig.Config['Login']['FinalizeCollectorLogin']) -eq 1
					}
					else
					{
						# Default to unchecked if setting doesn't exist
						$UI.FinalizeCollectorLogin.Checked = $false
					}
				#endregion Step: Sync finalize collector login config to checkbox

				#region Step: Sync NeverRestarting collector login config to checkbox
					# Update NeverRestartingCollectorLogin CheckBox if the config value exists.
					if ($global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin'])
					{
						$UI.NeverRestartingCollectorLogin.Checked = ([int]$global:DashboardConfig.Config['Login']['NeverRestartingCollectorLogin']) -eq 1
					}
					else
					{
						# Default to unchecked if setting doesn't exist
						$UI.NeverRestartingCollectorLogin.Checked = $false
					}
				#endregion Step: Sync NeverRestarting collector login config to checkbox

				#region Step: Sync HideMinimizedWindows config to checkbox
					# Update HideMinimizedWindows CheckBox if the config value exists.
					if ($global:DashboardConfig.Config['Options']['HideMinimizedWindows'])
					{
						$UI.HideMinimizedWindows.Checked = ([int]$global:DashboardConfig.Config['Options']['HideMinimizedWindows']) -eq 1
					}
					else
					{
						# Default to unchecked if setting doesn't exist
						$UI.HideMinimizedWindows.Checked = $false
					}
				#endregion Step: Sync HideMinimizedWindows config to checkbox

				#region Step: Log Sync Success
					Write-Verbose '  UI: Config synced to UI' -ForegroundColor Green
					return $true
				#endregion Step: Log Sync Success
			}
			catch
			{
				#region Step: Handle errors during sync
					Write-Verbose "  UI: Failed to sync config to UI: $_" -ForegroundColor Red
					return $false
				#endregion Step: Handle errors during sync
			}
		#endregion Step: Attempt to sync global config to UI state
	}
#endregion Function: Sync-ConfigToUI
#endregion Helper Functions

#region Core UI Functions

#region Function: Initialize-UI
	function Initialize-UI
	{
		<#
		.SYNOPSIS
			Initializes all UI components for the dashboard application.
		.OUTPUTS
			[bool] Returns $true if UI initialization was successful, $false otherwise (though currently always returns $true or throws).
		.NOTES
			Creates the main form, settings form, all buttons, labels, text boxes, data grids,
			and context menus. Populates the global $global:DashboardConfig.UI object with references
			to these elements. Calls Register-UIEventHandlers at the end.
		#>
		[CmdletBinding()]
		param()

		#region Step: Log UI initialization start
			Write-Verbose '  UI: Initializing UI...' -ForegroundColor Cyan
		#endregion Step: Log UI initialization start

		#region Step: Create Main UI Elements
			#region Step: Create Main Application Form
				# $mainFormProps: Hashtable defining properties for the main application window.
				$mainFormProps = @{
					type            = 'Form'
					visible         = $false # Start hidden, shown later
					width           = 470
					height          = 440
					bg              = @(30, 30, 30)                                 # Dark background
					id              = 'MainForm'
					text            = 'Entropia Dashboard'
					startPosition   = 'CenterScreen' # Position controlled manually or by saved state
					formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None  # Borderless window
				}
				$mainForm = Set-UIElement @mainFormProps
			#endregion Step: Create Main Application Form

			#region Step: Create Settings Form
				# $settingsFormProps: Hashtable defining properties for the settings dialog window.
				$settingsFormProps = @{
					type            = 'Form'
					visible         = $false # Start hidden
					width           = 470
					height          = 440
					bg              = @(30, 30, 30)
					id              = 'SettingsForm'
					text            = 'Settings'
					startPosition   = 'Manual'
					formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
					opacity         = 0 # Start invisible for fade-in effect
					topMost         = $true # Always on top of main form when visible
				}
				$settingsForm = Set-UIElement @settingsFormProps
			#endregion Step: Create Settings Form

			#region Step: Load custom icon if specified and exists
				# Attempt to load a custom icon from the path defined in global config.
				if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon))
				{
					try
					{
						$icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon)
						$mainForm.Icon = $icon
						$settingsForm.Icon = $icon
					}
					catch
					{
						Write-Verbose "  UI: Failed to load icon from $($global:DashboardConfig.Paths.Icon): $_" -ForegroundColor Red
					}
				}
			#endregion Step: Load custom icon if specified and exists

			#region Step: Create Top Bar Panel
				# $topBarProps: Hashtable defining properties for the panel used as a custom title/drag bar.
				$topBarProps = @{
					type    = 'Panel'
					width   = 470
					height  = 30
					bg      = @(20, 20, 20) # Dark background
					id      = 'TopBar'
				}
				$topBar = Set-UIElement @topBarProps
			#endregion Step: Create Top Bar Panel

			#region Step: Create Title Label
				# $titleLabelProps: Hashtable defining properties for the application title label on the top bar.
				$titleLabelProps = @{
					type    = 'Label'
					width   = 140
					height  = 12
					top     = 5
					left    = 10
					fg      = @(240, 240, 240)
					id      = 'TitleLabel'
					text    = 'Entropia Dashboard'
					font    = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
				}
				$titleLabelForm = Set-UIElement @titleLabelProps
			#endregion Step: Create Title Label

			#region Step: Create Copyright Label
				# $copyrightLabelProps: Hashtable defining properties for the application copyright label on the top bar.
				$copyrightLabelProps = @{
					type    = 'Label'
					width   = 140
					height  = 10
					top     = 16
					left    = 10
					fg      = @(230, 230, 230)
					id      = 'CopyrightLabel'
					text    = [char]0x00A9 + ' Immortal / Divine 2025 - v1.2.1'
					font    = New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Italic)
				}
				$copyrightLabelForm = Set-UIElement @copyrightLabelProps
			#endregion Step: Create Copyright Label

			#region Step: Create Minimize Button
				# $minFormProps: Hashtable defining properties for the minimize window button.
				$minFormProps = @{
					type    = 'Button'
					width   = 30
					height  = 30
					left    = 410
					bg      = @(40, 40, 40)
					fg      = @(240, 240, 240)
					id      = 'MinForm'
					text    = '_'
					fs      = 'Flat' # Flat appearance
					font    = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
				}
				$btnMinimizeForm = Set-UIElement @minFormProps
			#endregion Step: Create Minimize Button

			#region Step: Create Close Button
				# $closeFormProps: Hashtable defining properties for the close window button.
				$closeFormProps = @{
					type    = 'Button'
					width   = 30
					height  = 30
					left    = 440
					bg      = @(210, 45, 45) # Red color for mouse over (in 1.1)
					fg      = @(240, 240, 240)
					id      = 'CloseForm'
					text    = [char]0x166D # 'X' symbol
					fs      = 'Flat'
					font    = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
				}
				$btnCloseForm = Set-UIElement @closeFormProps
			#endregion Step: Create Close Button
		#endregion Step: Create Main UI Elements

		#region Step: Create Main Form Action Buttons
			#region Step: Create Launch Button
				# $launchProps: Hashtable defining properties for the Launch button.
				$launchProps = @{
					type    = 'Button'
					width   = 125
					height  = 30
					top     = 40
					left    = 15
					bg      = @(35, 175, 75) # Green color for active status
					fg      = @(240, 240, 240)
					id      = 'Launch'
					text    = 'Launch'
					fs      = 'Flat'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$btnLaunch = Set-UIElement @launchProps
			#endregion Step: Create Launch Button

			#region Step: Create Login Button
				# $loginProps: Hashtable defining properties for the Login button.
				$loginProps = @{
					type    = 'Button'
					width   = 125
					height  = 30
					top     = 40
					left    = 150
					bg      = @(35, 175, 75) # Green color for active status
					fg      = @(240, 240, 240)
					id      = 'Login'
					text    = 'Login'
					fs      = 'Flat'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$btnLogin = Set-UIElement @loginProps
			#endregion Step: Create Login Button

			#region Step: Create Settings Button
				# $settingsProps: Hashtable defining properties for the Settings button.
				$settingsProps = @{
					type    = 'Button'
					width   = 80
					height  = 30
					top     = 40
					left    = 285
					bg      = @(255, 165, 0) # Orange color for invalid settings (in 1.1)
					fg      = @(240, 240, 240)
					id      = 'Settings'
					text    = 'Settings'
					fs      = 'Flat'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$btnSettings = Set-UIElement @settingsProps
			#endregion Step: Create Settings Button

			#region Step: Create Terminate Button (Terminate Selected)
				# $exitProps: Hashtable defining properties for the button to terminate selected processes.
				$exitProps = @{
					type    = 'Button'
					width   = 80
					height  = 30
					top     = 40
					left    = 375
					bg      = @(210, 45, 45) # Red color for mouse over (in 1.1)
					fg      = @(240, 240, 240)
					id      = 'Terminate'
					text    = 'Terminate'
					fs      = 'Flat'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$btnStop = Set-UIElement @exitProps
			#endregion Step: Create Terminate Button (Terminate Selected)

			#region Step: Create Ftool Button
				# $ftoolProps: Hashtable defining properties for the main Ftool button.
				$ftoolProps = @{
					type    = 'Button'
					width   = 440
					height  = 30
					top     = 75
					left    = 15
					bg      = @(40, 40, 40)
					fg      = @(240, 240, 240)
					id      = 'Ftool'
					text    = 'Ftool'
					fs      = 'Flat'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$btnFtool = Set-UIElement @ftoolProps
			#endregion Step: Create Ftool Button
		#endregion Step: Create Main Form Action Buttons

		#region Step: Create Settings Form Controls
			#region Step: Create Save Settings Button
				# $saveProps: Hashtable defining properties for the Save button on the settings form.
				$saveProps = @{
					type    = 'Button'
					width   = 120
					height  = 40
					top     = 340
					left    = 20
					bg      = @(35, 175, 75) # Green color for valid settings (in 1.1)
					fg      = @(240, 240, 240)
					id      = 'Save'
					text    = 'Save'
					fs      = 'Flat'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$btnSave = Set-UIElement @saveProps
			#endregion Step: Create Save Settings Button

			#region Step: Create Cancel Settings Button
				# $cancelProps: Hashtable defining properties for the Cancel button on the settings form.
				$cancelProps = @{
					type    = 'Button'
					width   = 120
					height  = 40
					top     = 340
					left    = 150
					bg      = @(210, 45, 45) # Red color for mouse over (in 1.1)
					fg      = @(240, 240, 240)
					id      = 'Cancel'
					text    = 'Cancel'
					fs      = 'Flat'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$btnCancel = Set-UIElement @cancelProps
			#endregion Step: Create Cancel Settings Button

			#region Step: Create Browse Launcher Path Button
				# $browseLauncherProps: Hashtable defining properties for the Browse button next to the launcher path input.
				$browseLauncherProps = @{
					type    = 'Button'
					width   = 55
					height  = 25
					top     = 20
					left    = 110
					bg      = @(40, 40, 40) 
					fg      = @(240, 240, 240)
					id      = 'Browse'
					text    = 'Browse'
					fs      = 'Flat'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$btnBrowseLauncher = Set-UIElement @browseLauncherProps
			#endregion Step: Create Browse Launcher Path Button

			#region Step: Create Launcher Path Label
				# $launcherLabelProps: Hashtable defining properties for the label associated with the launcher path input.
				$launcherLabelProps = @{
					type    = 'Label'
					width   = 85
					height  = 20
					top     = 25
					left    = 20
					bg      = @(40, 40, 40, 0) # Transparent background
					fg      = @(240, 240, 240)
					id      = 'LabelLauncher'
					text    = 'Launcher Path:'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$lblLauncher = Set-UIElement @launcherLabelProps
			#endregion Step: Create Launcher Path Label

			#region Step: Create Process Name Label
				# $processNameLabelProps: Hashtable defining properties for the label associated with the process name input.
				$processNameLabelProps = @{
					type    = 'Label'
					width   = 85
					height  = 20
					top     = 95
					left    = 20
					bg      = @(40, 40, 40, 0) # Transparent background
					fg      = @(240, 240, 240)
					id      = 'LabelProcess'
					text    = 'Process Name:'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$lblProcessName = Set-UIElement @processNameLabelProps
			#endregion Step: Create Process Name Label

			#region Step: Create Max Clients Label
				# $maxClientsLabelProps: Hashtable defining properties for the label associated with the max clients input.
				$maxClientsLabelProps = @{
					type    = 'Label'
					width   = 85
					height  = 20
					top     = 165
					left    = 20
					bg      = @(40, 40, 40, 0) # Transparent background
					fg      = @(240, 240, 240)
					id      = 'LabelMax'
					text    = 'Max Clients:'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$lblMaxClients = Set-UIElement @maxClientsLabelProps
			#endregion Step: Create Max Clients Label
		#endregion Step: Create Settings Form Controls

		#region Step: Create Settings Form Input Controls
			#region Step: Create Launcher Path TextBox
				# $launcherTextBoxProps: Hashtable defining properties for the TextBox to input the launcher path.
				$launcherTextBoxProps = @{
					type    = 'TextBox'
					width   = 150
					height  = 30
					top     = 50
					left    = 20
					bg      = @(40, 40, 40)
					fg      = @(240, 240, 240)
					id      = 'InputLauncher'
					text    = ''
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$txtLauncher = Set-UIElement @launcherTextBoxProps
			#endregion Step: Create Launcher Path TextBox

			#region Step: Create Process Name TextBox
				# $processNameTextBoxProps: Hashtable defining properties for the TextBox to input the target process name.
				$processNameTextBoxProps = @{
					type    = 'TextBox'
					width   = 150
					height  = 30
					top     = 120
					left    = 20
					bg      = @(40, 40, 40)
					fg      = @(240, 240, 240)
					id      = 'InputProcess'
					text    = ''
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$txtProcessName = Set-UIElement @processNameTextBoxProps
			#endregion Step: Create Process Name TextBox

			#region Step: Create Max Clients TextBox
				# $maxClientsTextBoxProps: Hashtable defining properties for the TextBox to input the maximum number of clients.
				$maxClientsTextBoxProps = @{
					type    = 'TextBox'
					width   = 150
					height  = 30
					top     = 190
					left    = 20
					bg      = @(40, 40, 40)
					fg      = @(240, 240, 240)
					id      = 'InputMax'
					text    = ''
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$txtMaxClients = Set-UIElement @maxClientsTextBoxProps
			#endregion Step: Create Max Clients TextBox

			#region Step: Define Slot Options for Login Positions
				# $slotOptions: Array defining the available choices for login position ComboBoxes.
				$slotOptions = @('1', '2', '3')
				# $LoginCombos: Ordered dictionary to store references to the created login ComboBoxes.
				$LoginCombos = [ordered]@{}
			#endregion Step: Define Slot Options for Login Positions

			#region Step: Dynamically Create Login Position Labels and ComboBoxes
				# Loop to create a label and ComboBox for each potential login position (1 to 10).
				1..10 | ForEach-Object {
					$i = $_
					$positionKey = "Login$i"

					#region Step: Create Label for Position $i
						# $lblProps: Hashtable defining properties for the label for the current login position.
						$lblProps = @{
							type    = 'Label'
							visible = $true
							width   = 110
							height  = 20
							top     = (25 + (($i - 1) * 30)) # Calculate vertical position
							left    = 180
							bg      = @(30, 30, 30, 0) # Transparent background
							fg      = @(240, 240, 240)
							id      = "LabelPos$i"
							text    = "Login Position $i`:"
							font    = New-Object System.Drawing.Font('Segoe UI', 9)
						}
						$lbl = Set-UIElement @lblProps
					#endregion Step: Create Label for Position $i

					#region Step: Create ComboBox for Position $i
						# $cmbProps: Hashtable defining properties for the ComboBox for the current login position.
						$cmbProps = @{
							type          = 'ComboBox'
							visible       = $true
							width         = 150
							height        = 25
							top           = (25 + (($i - 1) * 30)) # Calculate vertical position
							left          = 290
							bg            = @(40, 40, 40)
							fg            = @(240, 240, 240)
							fs            = 'Flat'
							id            = "Login$i"
							font          = New-Object System.Drawing.Font('Segoe UI', 9)
							dropdownstyle = 'DropDownList' # User cannot type custom values
						}
						$cmb = Set-UIElement @cmbProps
					#endregion Step: Create ComboBox for Position $i

					#region Step: Add Slot Options to ComboBox
						# Populate the created ComboBox with the defined slot options.
						$slotOptions | ForEach-Object {
							$cmb.Items.Add($_)
						}
					#endregion Step: Add Slot Options to ComboBox

					#region Step: Add Controls to Settings Form and Store ComboBox
						# Add the newly created label and ComboBox to the settings form's controls collection.
						$settingsForm.Controls.Add($lbl)
						$settingsForm.Controls.Add($cmb)
						# Store a reference to the ComboBox in the $LoginCombos dictionary.
						$LoginCombos[$positionKey] = $cmb
					#endregion Step: Add Controls to Settings Form and Store ComboBox
				}
			#endregion Step: Dynamically Create Login Position Labels and ComboBoxes

			#region Step: Create Finalize Collector Login CheckBox
				# $finalizeCheckBoxProps: Hashtable defining properties for the CheckBox to enable/disable finalize collector login.
				$finalizeCheckBoxProps = @{
					type    = 'CheckBox'
					width   = 200
					height  = 20
					top     = 310
					left    = 0
					bg      = @(30, 30, 30)
					fg      = @(240, 240, 240)
					id      = 'FinalizeCollectorLogin'
					text    = 'Start Collector after Login'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$chkFinalizeLogin = Set-UIElement @finalizeCheckBoxProps
				$settingsForm.Controls.Add($chkFinalizeLogin)
			#endregion Step: Create Finalize Collector Login CheckBox

			#region Step: Create NeverRestarting Collector Login CheckBox
				# $NeverRestartingCheckBoxProps: Hashtable defining properties for the CheckBox to enable/disable NeverRestarting collector login.
				$NeverRestartingCheckBoxProps = @{
					type    = 'CheckBox'
					width   = 200
					height  = 20
					top     = 290
					left    = 0
					bg      = @(30, 30, 30)
					fg      = @(240, 240, 240)
					id      = 'NeverRestartingCollectorLogin'
					text    = 'Fix to start disc. collector'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$chkNeverRestartingLogin = Set-UIElement @NeverRestartingCheckBoxProps
				$settingsForm.Controls.Add($chkNeverRestartingLogin)
			#endregion Step: Create NeverRestarting Collector Login CheckBox

			#region Step: Create Hide Minimized Windows CheckBox
				# $hideMinimizedWindowsCheckBoxProps: Hashtable defining properties for the CheckBox to enable/disable hiding minimized windows.
				$hideMinimizedWindowsCheckBoxProps = @{
					type    = 'CheckBox'
					width   = 200
					height  = 20
					top     = 270
					left    = 0
					bg      = @(30, 30, 30)
					fg      = @(240, 240, 240)
					id      = 'HideMinimizedWindows'
					text    = 'Hide minimized clients'
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$chkHideMinimizedWindows = Set-UIElement @hideMinimizedWindowsCheckBoxProps
				$settingsForm.Controls.Add($chkHideMinimizedWindows)
			#endregion Step: Create Hide Minimized Windows CheckBox

		#endregion Step: Create Settings Form Input Controls

		#region Step: Create Main Form DataGrid Display Controls
			#region Step: Create Main DataGrid (Process List) (in 1.1)
				# $dataGridMainProps: Hashtable defining properties for the primary DataGridView displaying process info.
				$dataGridMainProps = @{
					type    = 'DataGridView'
					visible = $false
					width   = 155
					height  = 320
					top     = 115
					left    = 5
					bg      = @(40, 40, 40)
					fg      = @(240, 240, 240)
					id      = 'DataGridMain'
					text    = ''
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$DataGridMain = Set-UIElement @dataGridMainProps
			#endregion Step: Create Main DataGrid (Process List) (in 1.1)

			#region Step: Create Filler DataGrid
				# $dataGridFillerProps: Hashtable defining properties for the secondary DataGridView (purpose might be specific).
				$dataGridFillerProps = @{
					type    = 'DataGridView'
					width   = 450
					height  = 320
					top     = 115
					left    = 10
					bg      = @(40, 40, 40)
					fg      = @(240, 240, 240)
					id      = 'DataGridFiller'
					text    = ''
					font    = New-Object System.Drawing.Font('Segoe UI', 9)
				}
				$DataGridFiller = Set-UIElement @dataGridFillerProps
			#endregion Step: Create Filler DataGrid
		#endregion Step: Create Main Form DataGrid Display Controls

		#region Step: Create Context Menu and Login Position Controls

			#region Step: Create Context Menu for DataGrids
				# Create the context menu strip and its items for DataGrid interactions.
				$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
				$itmFront = New-Object System.Windows.Forms.ToolStripMenuItem('Show')
				$itmBack = New-Object System.Windows.Forms.ToolStripMenuItem('Minimize')
				$itmResizeCenter = New-Object System.Windows.Forms.ToolStripMenuItem('Resize')
			#endregion Step: Create Context Menu for DataGrids

		#endregion Step: Create Context Menu and Login Position Controls

		#region Step: Set Up Control Hierarchy and Context Menus
			#region Step: Add Controls to Main Form
				# Add the primary controls to the main application form.
				$mainForm.Controls.AddRange(@($topBar, $btnLogin, $btnFtool, $btnLaunch, $btnSettings, $btnStop, $DataGridMain, $DataGridFiller))
			#endregion Step: Add Controls to Main Form

			#region Step: Add Controls to Top Bar
				# Add the title label and window control buttons to the top bar panel.
				$topBar.Controls.AddRange(@($titleLabelForm, $copyrightLabelForm, $btnMinimizeForm, $btnCloseForm))
			#endregion Step: Add Controls to Top Bar

			#region Step: Set Up Context Menu for DataGrids
				# Add the previously created items to the context menu strip.
				$ctxMenu.Items.AddRange(@($itmFront, $itmBack, $itmResizeCenter))
				# Assign the context menu to both DataGridView controls.
				$DataGridMain.ContextMenuStrip = $ctxMenu
				$DataGridFiller.ContextMenuStrip = $ctxMenu
			#endregion Step: Set Up Context Menu for DataGrids

			#region Step: Add Controls to Settings Form
				# Add the primary controls to the settings form. (Login position controls were added dynamically earlier).
				$settingsForm.Controls.AddRange(@($btnSave, $btnCancel, $lblLauncher, $txtLauncher, $btnBrowseLauncher, $lblProcessName, $txtProcessName, $lblMaxClients, $txtMaxClients))
			#endregion Step: Add Controls to Settings Form
		#endregion Step: Set Up Control Hierarchy and Context Menus

		#region Step: Create Global UI Object for Element Access
			# $global:DashboardConfig.UI: A central PSCustomObject holding references to all created UI elements for easy access throughout the application.
			$global:DashboardConfig.UI = [PSCustomObject]@{
				# Main form and containers
				MainForm                   = $mainForm
				SettingsForm               = $settingsForm
				TopBar                     = $topBar

				# Window control buttons
				CloseForm                  = $btnCloseForm
				MinForm                    = $btnMinimizeForm

				# Main display elements
				DataGridMain               = $DataGridMain
				DataGridFiller             = $DataGridFiller

				# Main action buttons
				LoginButton                = $btnLogin
				Ftool                      = $btnFtool
				Settings                   = $btnSettings
				Exit                       = $btnStop
				Launch                     = $btnLaunch

				# Settings form labels
				LabelLauncher              = $lblLauncher
				LabelProcess               = $lblProcessName
				LabelMax                   = $lblMaxClients

				# Settings form inputs
				InputLauncher              = $txtLauncher
				InputProcess               = $txtProcessName
				InputMax                   = $txtMaxClients
				Browse                     = $btnBrowseLauncher
				Save                       = $btnSave
				Cancel                     = $btnCancel

				# Login position controls
				PosRange                   = $slotOptions # Available position numbers
				Login                      = $LoginCombos # Dictionary of Login ComboBoxes
				FinalizeCollectorLogin     = $chkFinalizeLogin # Checkbox for finalize collector login
				NeverRestartingCollectorLogin = $chkNeverRestartingLogin # Checkbox for NeverRestarting collector login
				HideMinimizedWindows       = $chkHideMinimizedWindows # Checkbox for hiding minimized windows

				# Context menu items
				ContextMenu                = $ctxMenu
				ContextMenuFront           = $itmFront
				ContextMenuBack            = $itmBack
				ContextMenuResizeAndCenter = $itmResizeCenter
			}
		#endregion Step: Create Global UI Object for Element Access

		#region Step: Register All UI Event Handlers
			# Call the function to attach event handlers to the created UI elements.
			Register-UIEventHandlers
		#endregion Step: Register All UI Event Handlers

		#region Step: Return Success Status
			return $true
		#endregion Step: Return Success Status
	}
#endregion Function: Initialize-UI

#region Function: Register-UIEventHandlers
	function Register-UIEventHandlers
	{
		<#
		.SYNOPSIS
			Registers all necessary event handlers for the UI elements.
		.NOTES
			Defines a mapping of UI element names to their events and corresponding script blocks (actions).
			Uses Register-ObjectEvent to attach these handlers. Includes logic for form loading, closing,
			resizing, button clicks, context menu actions, etc. Ensures previous handlers with the same
			source identifier are unregistered first.
		#>
		[CmdletBinding()]
		param()

		#region Step: Validate Global UI Object Existence
			# Ensure the UI object has been initialized before attempting to register events.
			if ($null -eq $global:DashboardConfig.UI)
			{
				Write-Verbose '  UI: Global UI is null, exiting event registration' -ForegroundColor Red
				return
			}
		#endregion Step: Validate Global UI Object Existence

		#region Step: Define Event Handler Mappings
			# $eventMappings: Hashtable defining which script block to execute for specific events on specific UI elements.
			$eventMappings = @{
				# Main form events
				MainForm                   = @{
					#region Step: Handle MainForm Load Event
						# Initialize form on load
						Load        = {
							#region Step: Load Configuration from INI File on Form Load
								# Load settings from the INI file if the path is configured.
								if ($global:DashboardConfig.Paths.Ini)
								{
									# Check if INI file exists
									$iniExists = Test-Path -Path $global:DashboardConfig.Paths.Ini
									if ($iniExists)
									{
										Write-Verbose "  UI: INI file exists at: $($global:DashboardConfig.Paths.Ini)" -ForegroundColor DarkGray

										# Check file size (for debugging)
										$fileInfo = Get-Item -Path $global:DashboardConfig.Paths.Ini
										Write-Verbose "  UI: INI file size: $($fileInfo.Length) bytes" -ForegroundColor DarkGray

										# Try to read raw content for debugging
										try
										{
											$rawContent = Get-Content -Path $global:DashboardConfig.Paths.Ini -Raw
											Write-Verbose "  UI: INI content: `r`n$($rawContent.Substring(0, [Math]::Min(1000, $rawContent.Length))) `r`n..." -ForegroundColor DarkGray
										}
										catch
										{
											Write-Verbose "  UI: Could not read raw INI content: $_" -ForegroundColor Yellow
										}

										# Use Get-IniFileContent to read all settings at once
										$iniSettings = Get-IniFileContent -Ini $global:DashboardConfig.Paths.Ini

										# Check if we got any settings
										if ($iniSettings.Count -gt 0)
										{
											Write-Verbose '  UI: Successfully read settings from INI file' -ForegroundColor Green

											# Store settings in global variable
											$global:DashboardConfig.Config = $iniSettings

											# Log loaded settings (for debugging)
											foreach ($section in $global:DashboardConfig.Config.Keys)
											{
												if ($global:DashboardConfig.Config[$section] -and $global:DashboardConfig.Config[$section].Keys.Count -gt 0)
												{
													foreach ($key in $global:DashboardConfig.Config[$section].Keys)
													{
														Write-Verbose "  UI: Loaded setting $section.$key = $($global:DashboardConfig.Config[$section][$key])" -ForegroundColor DarkGrey
													}
												}
												else
												{
													Write-Verbose "  UI: Section $section has no keys" -ForegroundColor Yellow
												}
											}
										}
										else
										{
											Write-Verbose '  UI: No settings found in INI file' -ForegroundColor Yellow
										}
									}
									else
									{
										Write-Verbose "  UI: INI file does not exist at: $($global:DashboardConfig.Paths.Ini)" -ForegroundColor Yellow
									}

									# Update UI with loaded settings
									Sync-ConfigToUI
								}
							#endregion Step: Load Configuration from INI File on Form Load

							#region Step: Store Initial Control Properties for Resizing
								# Store initial dimensions and positions for dynamic resizing.
								$script:initialControlProps = @{}
								$script:initialFormWidth = $global:DashboardConfig.UI.MainForm.Width
								$script:initialFormHeight = $global:DashboardConfig.UI.MainForm.Height
							#endregion Step: Store Initial Control Properties for Resizing

							#region Step: Define Controls to Scale and Store Initial Properties
								# Define which controls should be scaled during form resize.
								$controlsToScale = @('TopBar', 'Login', 'Ftool', 'Settings', 'Exit', 'Launch', 'DataGridMain', 'DataGridFiller', 'MinForm', 'CloseForm')

								# Store initial properties for each scalable control.
								foreach ($controlName in $controlsToScale)
								{
									$control = $global:DashboardConfig.UI.$controlName
									if ($control)
									{
										$script:initialControlProps[$controlName] = @{
											Left             = $control.Left
											Top              = $control.Top
											Width            = $control.Width
											Height           = $control.Height
											IsScalableBottom = ($controlName -eq 'DataGridFiller' -or $controlName -eq 'DataGridMain') # Mark grids for vertical scaling
										}
									}
								}
							#endregion Step: Define Controls to Scale and Store Initial Properties
						}
					#endregion Step: Handle MainForm Load Event

					#region Step: Handle MainForm Shown Event
						# Start update timer when form is actually shown
						Shown       = {
							#region Step: Start DataGrid Update Timer When Form is Shown
								if ($global:DashboardConfig.UI.DataGridFiller)
								{
									try
									{
										Start-DataGridUpdateTimer
									}
									catch
									{
										# Silent error handling if timer start fails
										Write-Verbose "  UI: Failed to start DataGrid update timer: $_" -ForegroundColor Yellow
									}
								}
							#endregion Step: Start DataGrid Update Timer When Form is Shown
						}
					#endregion Step: Handle MainForm Shown Event

					#region Step: Handle MainForm FormClosing Event
						# Clean up resources when form is closing
						FormClosing = {
							param($src, $e)

							#region Step: Clean Up Resources on Form Closing
								Write-Verbose '  UI: Form closing - cleaning up resources' -ForegroundColor Cyan

								#region Step: Clean Up Ftool Instances
									# Clean up any running ftool forms and their associated resources.
									if ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Count -gt 0)
									{
										Write-Verbose "  UI: Cleaning up ftool instances: $($global:DashboardConfig.Resources.FtoolForms.Count) forms" -ForegroundColor Cyan

										# Get a copy of the keys to avoid collection modification issues during iteration.
										$instanceIds = @($global:DashboardConfig.Resources.FtoolForms.Keys)

										foreach ($instanceId in $instanceIds)
										{
											Write-Verbose "  UI: Cleaning up ftool instance: $instanceId" -ForegroundColor DarkGray
											$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]

											if ($form -and -not $form.IsDisposed)
											{
												# Use Stop-FtoolForm if available for proper cleanup.
												if (Get-Command -Name Stop-FtoolForm -ErrorAction SilentlyContinue)
												{
													Stop-FtoolForm -Form $form
												}
												else
												{
													# Fallback cleanup if Stop-FtoolForm is not found.
													$data = $form.Tag
													if ($data)
													{
														# Clean up running spammer timer if exists.
														if ($data.RunningSpammer)
														{
															$data.RunningSpammer.Stop()
															$data.RunningSpammer.Dispose()
														}

														# Clean up form-specific timers.
														if ($data.Timers -and $data.Timers.Count -gt 0)
														{
															foreach ($timerKey in @($data.Timers.Keys))
															{
																$timer = $data.Timers[$timerKey]
																if ($timer)
																{
																	$timer.Stop()
																	$timer.Dispose()
																}
															}
														}
													}

													# Close and dispose form.
													$form.Close()
													$form.Dispose()
													# This might be too aggressive here, consider if needed.
													# [System.Windows.Forms.Application]::Exit()
												}

												# Remove the form reference from the global collection.
												$global:DashboardConfig.Resources.FtoolForms.Remove($instanceId)
											}
										}
									}
								#endregion Step: Clean Up Ftool Instances

								#region Step: Clean Up Global Timers
									# Clean up all registered timers stored in the global resources.
									if ($global:DashboardConfig.Resources.Timers -and $global:DashboardConfig.Resources.Timers.Count -gt 0)
									{
										Write-Verbose "  UI: Cleaning up $($global:DashboardConfig.Resources.Timers.Count) timers" -ForegroundColor DarkGray

										# Handle nested timer collections first (e.g., timers within ftool data).
										foreach ($collectionKey in @($global:DashboardConfig.Resources.Timers.Keys))
										{
											$collection = $global:DashboardConfig.Resources.Timers[$collectionKey]

											if ($collection -is [System.Collections.Hashtable] -or $collection -is [System.Collections.Specialized.OrderedDictionary])
											{
												$nestedKeys = @($collection.Keys)
												foreach ($nestedKey in $nestedKeys)
												{
													$timer = $collection[$nestedKey]
													if ($timer -is [System.Windows.Forms.Timer])
													{
														if ($timer.Enabled) { $timer.Stop() }
														$timer.Dispose()
														$collection.Remove($nestedKey)
													}
												}
											}
										}

										# Handle direct timers stored in the main Timers collection.
										$timerKeys = @($global:DashboardConfig.Resources.Timers.Keys)
										foreach ($key in $timerKeys)
										{
											$timer = $global:DashboardConfig.Resources.Timers[$key]
											if ($timer -is [System.Windows.Forms.Timer])
											{
												if ($timer.Enabled) { $timer.Stop() }
												$timer.Dispose()
												$global:DashboardConfig.Resources.Timers.Remove($key)
											}
										}
									}
								#endregion Step: Clean Up Global Timers

								#region Step: Clean Up Background Jobs
									# Stop and remove any running PowerShell background jobs.
									try
									{
										$runningJobs = Get-Job -ErrorAction SilentlyContinue | Where-Object { $_.State -ne 'Completed' }
										if ($runningJobs -and $runningJobs.Count -gt 0)
										{
											Write-Verbose "  UI: Stopping $($runningJobs.Count) background jobs." -ForegroundColor DarkGray
											$runningJobs | Stop-Job -ErrorAction SilentlyContinue
											Get-Job -ErrorAction SilentlyContinue | Remove-Job -Force -ErrorAction SilentlyContinue
										}
									}
									catch
									{
										Write-Verbose "  UI: Error cleaning up background jobs: $_"-ForegroundColor Red
									}
								#endregion Step: Clean Up Background Jobs

								#region Step: Clean Up Runspaces
									# Dispose of any active runspaces.
									if ($global:runspaces -and $global:runspaces.Count -gt 0)
									{
										Write-Verbose "  UI: Disposing $($global:runspaces.Count) runspaces." -ForegroundColor DarkGray
										foreach ($rs in $global:runspaces)
										{
											try
											{
												if ($rs.Runspace.RunspaceStateInfo.State -ne 'Closed')
												{
													$rs.PowerShell.Dispose()
													$rs.Runspace.Dispose()
												}
											}
											catch
											{
												Write-Verbose "  UI: Error disposing runspace: $_"-ForegroundColor Red
											}
										}
										$global:runspaces.Clear()
									}
								#endregion Step: Clean Up Runspaces

								#region Step: Clean Up Launch Resources
									# Stop any ongoing client launch processes.
									Stop-ClientLaunch
								#endregion Step: Clean Up Launch Resources

								#region Step: Force Garbage Collection
									# Explicitly run garbage collection to release memory.
									[System.GC]::Collect()
									[System.GC]::WaitForPendingFinalizers()
								#endregion Step: Force Garbage Collection
							#endregion Step: Clean Up Resources on Form Closing
						}
					#endregion Step: Handle MainForm FormClosing Event

					#region Step: Handle MainForm Resize Event
						# Handle form resizing to dynamically adjust control positions and sizes
						Resize      = {
							#region Step: Handle Form Resizing and Scale Controls
								# Skip if initialization data is missing
								if (-not $script:initialControlProps -or -not $global:DashboardConfig.UI)
								{
									return
								}

								# Calculate scaling factors based on current vs initial size
								$currentFormWidth = $global:DashboardConfig.UI.MainForm.ClientSize.Width
								$currentFormHeight = $global:DashboardConfig.UI.MainForm.ClientSize.Height
								$scaleW = $currentFormWidth / $script:initialFormWidth

								# Define fixed areas (e.g., height of the top button bar area)
								$fixedTopHeight = 125
								$bottomMargin = 10

								# Resize and reposition each scalable control
								foreach ($controlName in $script:initialControlProps.Keys)
								{
									$control = $global:DashboardConfig.UI.$controlName
									if ($control)
									{
										$initialProps = $script:initialControlProps[$controlName]

										# Calculate new position and width based on horizontal scale
										$newLeft = [int]($initialProps.Left * $scaleW)
										$newWidth = [int]($initialProps.Width * $scaleW)

										# Handle special case for bottom-anchored controls (DataGrids)
										if ($initialProps.IsScalableBottom)
										{
											$control.Top = $fixedTopHeight
											# Adjust height based on remaining form height
											$control.Height = [Math]::Max(100, $currentFormHeight - $fixedTopHeight - $bottomMargin)
										}
										else
										{
											# Keep original top and height for non-vertically-scaling controls
											$control.Top = $initialProps.Top
											$control.Height = $initialProps.Height
										}

										# Apply new position and width
										$control.Left = $newLeft
										$control.Width = $newWidth
									}
								}
							#endregion Step: Handle Form Resizing and Scale Controls
						}
					#endregion Step: Handle MainForm Resize Event

				}

				# Settings form events
				SettingsForm               = @{
					#region Step: Handle SettingsForm Load Event
						# Initialize form on load
						Load = {
							#region Step: Sync Config to UI When Settings Form Loads
								try
								{
									# Populate UI controls with current config values when the form loads.
									Sync-ConfigToUI
								}
								catch
								{
									Write-Verbose "  UI: Error loading settings form: $_" -ForegroundColor Red
									[System.Windows.Forms.MessageBox]::Show("Failed to load settings: $($_.Exception.Message)", 'Error',
										[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
								}
							#endregion Step: Sync Config to UI When Settings Form Loads
						}
					#endregion Step: Handle SettingsForm Load Event
				}

				# Minimize button event
				MinForm                    = @{
					#region Step: Handle MinForm Click Event
						Click = {
							#region Step: Minimize Main Form
								# Minimize the main application window.
								$global:DashboardConfig.UI.MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
							#endregion Step: Minimize Main Form
						}
					#endregion Step: Handle MinForm Click Event
				}

				# Close button event
				CloseForm                  = @{
					#region Step: Handle CloseForm Click Event
						Click = {
							#region Step: Close Main Form and Exit Application
								try
								{
									# Close the main form, which triggers the FormClosing event for cleanup.
									$global:DashboardConfig.UI.MainForm.Close()
									# Attempt to exit the application message loop.
									[System.Windows.Forms.Application]::Exit()
									# Forcefully stop the current PowerShell process as a final measure.
									Stop-Process -Id $PID -Force
								}
								catch
								{
									[System.Windows.Forms.MessageBox]::Show("Failed to close the application: $($_.Exception.Message)", 'Error',
										[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
								}
							#endregion Step: Close Main Form and Exit Application
						}
					#endregion Step: Handle CloseForm Click Event
				}

				# Top bar drag event
				TopBar                     = @{
					#region Step: Handle TopBar MouseDown Event
						MouseDown = {
							param($src, $e)
							#region Step: Enable Form Dragging via Top Bar
								# Use native Windows messages to allow dragging the borderless form by its top bar.
								[Custom.Native]::ReleaseCapture()
								[Custom.Native]::SendMessage($global:DashboardConfig.UI.MainForm.Handle, 0xA1, 0x2, 0) # WM_NCLBUTTONDOWN, HTCAPTION
							#endregion Step: Enable Form Dragging via Top Bar
						}
					#endregion Step: Handle TopBar MouseDown Event
				}

				# Settings button event
				Settings                   = @{
					#region Step: Handle Settings Button Click Event
						Click = {
							#region Step: Show Settings Form
								# Call the function to display the settings form with a fade-in effect.
								Show-SettingsForm
							#endregion Step: Show Settings Form
						}
					#endregion Step: Handle Settings Button Click Event
				}

				# Save button event (in Settings Form)
				Save                       = @{
					#region Step: Handle Save Button Click Event
						Click = {
							#region Step: Save Settings from UI to Config File
								try
								{
									Write-Verbose '  UI: Updating settings from UI' -ForegroundColor Cyan

									# Sync UI values back to the global config object.
									Sync-UIToConfig
									# Write the updated config object to the INI file.
									$result = Write-Config

									# Log the settings being saved (for debugging).
									Write-Verbose '  UI: Settings to save:' -ForegroundColor DarkGray
									foreach ($section in $global:DashboardConfig.Config.Keys)
									{
										foreach ($key in $global:DashboardConfig.Config[$section].Keys)
										{
											Write-Verbose "  UI: $section.$key = $($global:DashboardConfig.Config[$section][$key])" -ForegroundColor DarkGray
										}
									}

									# Check if writing the config was successful.
									if (!($result))
									{
										Write-Verbose '  UI: Failed to save settings to INI file' -ForegroundColor Red
										[System.Windows.Forms.MessageBox]::Show("Failed to save settings.", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
										return
									}

									# Hide settings form with fade-out effect upon successful save.
									Hide-SettingsForm
								}
								catch
								{
									# Show error message if saving fails.
									Write-Verbose "  UI: Failed to save settings to INI file: $($_.Exception.Message)" -ForegroundColor Red
									[System.Windows.Forms.MessageBox]::Show("Failed to save settings: $($_.Exception.Message)", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
								}
							#endregion Step: Save Settings from UI to Config File
						}
					#endregion Step: Handle Save Button Click Event
				}

				# Cancel button event (in Settings Form)
				Cancel                     = @{
					#region Step: Handle Cancel Button Click Event
						Click = {
							#region Step: Hide Settings Form
								# Call the function to hide the settings form without saving changes.
								Hide-SettingsForm
							#endregion Step: Hide Settings Form
						}
					#endregion Step: Handle Cancel Button Click Event
				}

				# Browse button event (in Settings Form)
				Browse                     = @{
					#region Step: Handle Browse Button Click Event
						Click = {
							#region Step: Show OpenFileDialog for Launcher Path
								# Open a standard file dialog to select the game launcher executable.
								$d = New-Object System.Windows.Forms.OpenFileDialog
								$d.Filter = 'Executable Files (*.exe)|*.exe|All Files (*.*)|*.*'
								if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
								{
									# Update the launcher path TextBox with the selected file.
									$global:DashboardConfig.UI.InputLauncher.Text = $d.FileName
								}
							#endregion Step: Show OpenFileDialog for Launcher Path
						}
					#endregion Step: Handle Browse Button Click Event
				}

				# DataGrid events
				DataGridFiller = @{

					#region Step: Handle DataGrid DoubleClick Event
						# Handle double-click to bring the corresponding process window to the front
						DoubleClick = {
							param($src, $e)
							try {
								$grid = $src
								if (-not $grid) { return }
								# Determine which row was double-clicked.
								$hitTestInfo = $grid.HitTest($e.X, $e.Y)
								if ($hitTestInfo.RowIndex -ge 0) {
									$row = $grid.Rows[$hitTestInfo.RowIndex]
									# Check if the row has associated process info and a valid window handle.
									# Assuming $row.Tag holds a process object or similar with MainWindowHandle
									if ($row.Tag -and $row.Tag.GetType().GetProperty('MainWindowHandle') -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero) {
										# Use helper function/native methods if available
										[Custom.Native]::BringToFront($row.Tag.MainWindowHandle)
										Write-Verbose "  UI: DoubleClick - Bringing window handle $($row.Tag.MainWindowHandle) to front." -ForegroundColor DarkGray
									} elseif ($row.Tag -and $row.Tag.GetType().GetProperty('MainWindowHandle')) {
										Write-Verbose "  UI: DoubleClick - Row $($hitTestInfo.RowIndex) has tag, but MainWindowHandle is Zero." -ForegroundColor DarkGray
									} else {
										Write-Verbose "  UI: DoubleClick - Row $($hitTestInfo.RowIndex) does not have a valid Tag with MainWindowHandle." -ForegroundColor DarkGray
									}
								}
							} catch {
								Write-Verbose "  UI: Error in DataGridFiller DoubleClick: $_" -ForegroundColor Red
							}
						}
					#endregion Step: Handle DataGrid DoubleClick Event

					#region Step: Handle DataGrid MouseDown Event
						# Handle right-click for context menu, left-click for selection, Alt+Left-click for drag initiation
						MouseDown = {
							param($src, $e) # $src is the DataGridView control itself

							try {
								$grid = $src # Use the source control passed to the event
								if (-not $grid) {
									Write-Verbose "  UI: MouseDown - Source grid object is null." -ForegroundColor Yellow
									return
								}

								$hitTestInfo = $grid.HitTest($e.X, $e.Y)

								# --- Right-Click Handling ---
								if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
									if ($hitTestInfo.RowIndex -ge 0) {
										# Ensure the clicked row is selected before showing the context menu.
										if (-not ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control)) {
											if (-not $grid.Rows[$hitTestInfo.RowIndex].Selected) {
												$grid.ClearSelection()
											}
										}
										$grid.Rows[$hitTestInfo.RowIndex].Selected = $true
										Write-Verbose "  UI: Right-clicked row $($hitTestInfo.RowIndex), ensuring selection." -ForegroundColor DarkGray
									}
								}
								# --- Left-Click Handling ---
								elseif ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
									if ($hitTestInfo.RowIndex -ge 0) { # Clicked on a row
										$clickedRow = $grid.Rows[$hitTestInfo.RowIndex]

										# --- Normal Left Click for Selection ---
			
											# Standard behavior: Clear previous selection if Ctrl is NOT held.
											if (-not ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control)) {
												# Only clear if the clicked row isn't the *only* selected row
												if ($grid.SelectedRows.Count -ne 1 -or -not $clickedRow.Selected) {
													$grid.ClearSelection()
												}
											}
											# Toggle selection if Ctrl is pressed, otherwise just select.
											if (([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control)) {
												# If Ctrl is held, toggle selection

											}


											Write-Verbose "  UI: Left-clicked row $($hitTestInfo.RowIndex). Selected: $($clickedRow.Selected)" -ForegroundColor DarkGray
			
									}
									elseif ($hitTestInfo.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::None) { # Clicked on empty space
										$grid.ClearSelection()
										Write-Verbose "  UI: Clicked on empty DataGrid area, cleared selection." -ForegroundColor DarkGray
									}
								}
							} catch {
								Write-Verbose "  UI: Error in DataGridFiller MouseDown: $_" -ForegroundColor Red
							}
						}
					#endregion Step: Handle DataGrid MouseDown Event (Initiates Drag)

				}

				# Context menu item events
				ContextMenuFront           = @{
					#region Step: Handle ContextMenuFront Click Event
						Click = {
							#region Step: Bring Selected Process Windows to Front
								# Iterate through selected rows in the DataGrid.
								if ($global:DashboardConfig.UI.DataGridFiller.SelectedRows.Count -gt 0)
								{
									foreach ($row in $global:DashboardConfig.UI.DataGridFiller.SelectedRows)
									{
										# Check for valid process info and window handle.
										if ($row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
										{
											# Bring the window to the foreground.
											[Custom.Native]::BringToFront($row.Tag.MainWindowHandle)
										}
									}
								}
							#endregion Step: Bring Selected Process Windows to Front
						}
					#endregion Step: Handle ContextMenuFront Click Event
				}
				
				ContextMenuBack            = @{
					#region Step: Handle ContextMenuBack Click Event
						Click = {
							#region Step: Send Selected Process Windows to Back (Minimize)
								# Iterate through selected rows in the DataGrid.
								if ($global:DashboardConfig.UI.DataGridFiller.SelectedRows.Count -gt 0)
								{
									foreach ($row in $global:DashboardConfig.UI.DataGridFiller.SelectedRows)
									{
										# Check for valid process info and window handle.
										if ($row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
										{
											Write-Verbose 'UI: Minimizing...' -ForegroundColor Cyan
											[Custom.Native]::SendToBack($row.Tag.MainWindowHandle)
											Write-Verbose 'UI: Optimizing...' -ForegroundColor Cyan
											[Custom.Native]::EmptyWorkingSet($row.Tag.Handle)
										}
									}
								}
							#endregion Step: Send Selected Process Windows to Back (Minimize)
						}
					#endregion Step: Handle ContextMenuBack Click Event
				}

				ContextMenuResizeAndCenter = @{
					#region Step: Handle ContextMenuResizeAndCenter Click Event
						Click = {
							#region Step: Resize Selected Process Windows to Standard Size
								# Iterate through selected rows in the DataGrid.
								if ($global:DashboardConfig.UI.DataGridFiller.SelectedRows.Count -gt 0)
								{
									# Get screen dimensions and define standard window size.
									$scr = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
									$width = 1040
									$height = 807

									foreach ($row in $global:DashboardConfig.UI.DataGridFiller.SelectedRows)
									{
										# Check for valid process info and window handle.
										if ($row.Tag -and $row.Tag.MainWindowHandle -ne [IntPtr]::Zero)
										{
											# Use native function to resize and center the window.
											[Custom.Native]::PositionWindow(
												$row.Tag.MainWindowHandle,
												[Custom.Native]::TopWindowHandle,
												[int](($scr.Width - $width) / 2),  # Center X
												[int](($scr.Height - $height) / 2), # Center Y
												$width,
												$height,
												# Flags: Don't activate
												[Custom.Native+WindowPositionOptions]::DoNotActivate
											)
										}
									}
								}
							#endregion Step: Resize Selected Process Windows to Standard Size
						}
					#endregion Step: Handle ContextMenuResizeAndCenter Click Event
				}

				# Launch button event
				Launch                     = @{
					#region Step: Handle Launch Button Click Event
						Click = {
							#region Step: Initialize Client Launch Process
								try
								{
									# Call the function responsible for starting the client launch sequence.
									Start-ClientLaunch

								}
								catch
								{
									# Handle errors during launch initialization.
									$global:DashboardConfig.State.LaunchActive = $false
									Write-Verbose "  UI: Launch initialization failed: $($_.Exception.Message)" -ForegroundColor Red
									[System.Windows.Forms.MessageBox]::Show("An error occurred during launch: $($_.Exception.Message)", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
								}
							#endregion Step: Initialize Client Launch Process
						}
						
					#endregion Step: Handle Launch Button Click Event
				}

				# Login button event
				LoginButton                = @{
					#region Step: Handle Login Button Click Event
						Click = {
							#region Step: Initiate Login Process for Selected Clients
								try
								{
									# Check if the LoginSelectedRow function (likely from login.psm1) is available.
									if (Get-Command -Name LoginSelectedRow -ErrorAction SilentlyContinue)
									{
										# Ensure at least one client row is selected in the DataGrid.
										if ($global:DashboardConfig.UI.DataGridFiller.SelectedRows.Count -eq 0)
										{
											Write-Verbose '  UI: No clients selected for login' -ForegroundColor Yellow
											[System.Windows.Forms.MessageBox]::Show('Please select at least one client to log in.', 'Login', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
											return
										}

										# Determine log file path based on launcher path.
										$LogFolder = ($global:DashboardConfig.Config['LauncherPath']['LauncherPath'] -replace '\\Launcher\.exe$', '')
										$LogFilePath = Join-Path -Path $LogFolder -ChildPath "Log\network_$(Get-Date -Format 'yyyyMMdd').log"

										# Call the login function, passing the log file path.
										Write-Verbose '  UI: Starting login process for selected clients...' -ForegroundColor Cyan
										LoginSelectedRow -LogFilePath $LogFilePath

									}
									else
									{
										Write-Verbose '  UI: Login module (LoginSelectedRow command) not available' -ForegroundColor Red
										[System.Windows.Forms.MessageBox]::Show('Login functionality is not available.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
									}
								}
								catch
								{
									Write-Verbose "  UI: Error in login process: $_" -ForegroundColor Red
									[System.Windows.Forms.MessageBox]::Show("An error occurred during login: $($_.Exception.Message)", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
								}
							#endregion Step: Initiate Login Process for Selected Clients
						}
					#endregion Step: Handle Login Button Click Event
				}

				# Ftool button event
				Ftool                      = @{
					#region Step: Handle Ftool Button Click Event
						Click = {
							#region Step: Initiate Ftool Process for Selected Clients
							
								try
								{
									# Validate DataGrid exists.
									if (-not $global:DashboardConfig.UI.DataGridFiller)
									{
										Write-Verbose '  UI: DataGrid not found for Ftool action' -ForegroundColor Red
										return
									}

									# Ensure at least one row is selected.
									$selectedRows = $global:DashboardConfig.UI.DataGridFiller.SelectedRows
									Write-Verbose "  UI: Ftool button clicked, selected rows count: $($selectedRows.Count)" -ForegroundColor DarkGray

									if ($selectedRows.Count -eq 0)
									{
										[System.Windows.Forms.MessageBox]::Show('Please select at least one client row to use Ftool.', 'Ftool', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
										return
									}

									# Check if FtoolSelectedRow function (likely from ftool.psm1) is available.
									if (-not (Get-Command -Name FtoolSelectedRow -ErrorAction SilentlyContinue)) {
										Write-Verbose '  UI: Ftool module (FtoolSelectedRow command) not available' -ForegroundColor Red
										[System.Windows.Forms.MessageBox]::Show('Ftool functionality is not available.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
										return
									}

									# Process each selected row using the Ftool function.
									foreach ($row in $selectedRows)
									{
										FtoolSelectedRow $row
									}
								}
								catch
								{
									Write-Verbose "  UI: Error in Ftool click handler: $($_.Exception.Message)" -ForegroundColor Red
									[System.Windows.Forms.MessageBox]::Show("An error occurred initiating Ftool: $($_.Exception.Message)", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
								}
							#endregion Step: Initiate Ftool Process for Selected Clients
						}
					#endregion Step: Handle Ftool Button Click Event
				}

				# Exit button event (Terminate Selected)
				Exit                       = @{
					#region Step: Handle Exit Button Click Event
						Click = {
							#region Step: Terminate Selected Processes
								# Ensure at least one row is selected.
								if ($global:DashboardConfig.UI.DataGridFiller.SelectedRows.Count -gt 0)
								{
									# Confirm termination with the user.
									$result = [System.Windows.Forms.MessageBox]::Show(
										'Are you sure you want to terminate the selected processes?',
										'Confirm Termination',
										[System.Windows.Forms.MessageBoxButtons]::YesNo,
										[System.Windows.Forms.MessageBoxIcon]::Warning
									)

									if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
									{
										# Iterate through selected rows and attempt termination.
										foreach ($row in $global:DashboardConfig.UI.DataGridFiller.SelectedRows)
										{
											if ($row.Tag -and $row.Tag.Id)
											{
												$processId = $row.Tag.Id
												try
												{
													# Get the process object.
													$process = Get-Process -Id $processId -ErrorAction Stop

													# Try to close gracefully first.
													if (-not $process.HasExited)
													{
														# Restore if minimized before closing main window.
														if ([Custom.Native]::IsWindowMinimized($process.MainWindowHandle))
														{
															[Custom.Native]::ShowWindow($process.MainWindowHandle, [Custom.Native]::SW_RESTORE)
														}
														Start-Sleep -MilliSeconds 100

														Write-Verbose "  UI: Attempting graceful shutdown for PID $processId..." -ForegroundColor DarkGray
														$process.CloseMainWindow() | Out-Null

														# Wait briefly for graceful exit.
														if (-not $process.WaitForExit(1000)) {
															# If still running, force kill.
															Write-Verbose "  UI: Graceful shutdown failed for PID $processId. Forcing termination." -ForegroundColor Yellow
															$process.Kill()
														} else {
															Write-Verbose "  UI: Process PID $processId exited gracefully." -ForegroundColor Green
														}
													}
													Write-Verbose "  UI: Successfully terminated process ID $processId" -ForegroundColor Green
												}
												catch [System.ArgumentException] {
													# Process already exited or ID invalid
													Write-Verbose "  UI: Process ID $processId not found or already exited." -ForegroundColor Yellow
												}
												catch {
													Write-Verbose "  UI: Failed to terminate process ID $($processId): $_" -ForegroundColor Red

													# Try alternative termination as fallback.
													try
													{
														Write-Verbose "  UI: Attempting Stop-Process fallback for PID $processId..." -ForegroundColor DarkGray
														Stop-Process -Id $processId -Force -ErrorAction Stop
														Write-Verbose "  UI: Terminated process ID $processId using Stop-Process" -ForegroundColor Green
													}
													catch
													{
														Write-Verbose "  UI: Failed to terminate process ID $processId using Stop-Process: $_" -ForegroundColor Red
													}
												}
											}
										}

										# Refresh the grid after termination attempts.
										if ($global:DashboardConfig.UI.DataGridFiller.RefreshMethod)
										{
											Write-Verbose "  UI: Refreshing DataGrid after termination attempts." -ForegroundColor DarkGray
											$refresh = $global:DashboardConfig.UI.DataGridFiller.RefreshMethod
											try {
												if ($refresh -is [ScriptBlock]) { & $refresh }
												elseif ($refresh -is [System.Delegate] -or $refresh -is [System.Action]) { $refresh.Invoke() }
												elseif ($refresh -is [string]) {
													$cmd = Get-Command $refresh -ErrorAction SilentlyContinue
													if ($cmd) { & $cmd.Name } else { Invoke-Expression $refresh }
												} else {
													Write-Verbose ("UI: RefreshMethod unsupported type {0}" -f $refresh.GetType().FullName)
												}
											} catch {
												Write-Verbose ("UI: RefreshMethod invocation failed: {0}" -f $_.Exception.Message)
											}
										}
									}
								} else {
									[System.Windows.Forms.MessageBox]::Show('Please select at least one process to terminate.', 'Terminate Process', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
								}
							#endregion Step: Terminate Selected Processes
						}
					#endregion Step: Handle Exit Button Click Event
				}
			}
		#endregion Step: Define Event Handler Mappings

		#region Step: Register Defined Event Handlers for Each UI Element
			# Iterate through the event mappings and register each handler.
			foreach ($elementName in $eventMappings.Keys)
			{
				$element = $global:DashboardConfig.UI.$elementName
				if ($element)
				{
					foreach ($e in $eventMappings[$elementName].Keys)
					{
						# Create a unique source identifier for each event subscription.
						$sourceIdentifier = "EntropiaDashboard.$elementName.$e"

						# Unregister any existing event handler with the same identifier to prevent duplicates.
						Get-EventSubscriber -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue | Unregister-Event

						# Register the new event handler.
						Register-ObjectEvent -InputObject $element `
							-EventName $e `
							-Action $eventMappings[$elementName][$e] `
							-SourceIdentifier $sourceIdentifier `
							-ErrorAction SilentlyContinue # Continue if registration fails for some reason
					}
				}
				else {
					 Write-Verbose "  UI: Element '$elementName' not found in global UI object during event registration." -ForegroundColor Yellow
				}
			}
		#endregion Step: Register Defined Event Handlers for Each UI Element

		#region Step: Mark UI as Initialized
			# Set a flag indicating that UI initialization and event registration are complete.
			$global:DashboardConfig.State.UIInitialized = $true
			Write-Verbose '  UI: Event handlers registered.' -ForegroundColor Green
		#endregion Step: Mark UI as Initialized
	}
#endregion Function: Register-UIEventHandlers

#region Function: Show-SettingsForm
	function Show-SettingsForm
	{
		<#
		.SYNOPSIS
			Shows the settings form with a fade-in animation effect.
		.NOTES
			Makes the settings form visible, positions it relative to the main form,
			and initiates a timer-based fade-in by gradually increasing opacity.
			Prevents concurrent fade animations.
		#>
		[CmdletBinding()]
		param()

		#region Step: Prevent Concurrent Animations
			# Check if a fade-in or fade-out animation is already in progress.
			if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
				($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
			{
				return # Exit if an animation is active
			}
		#endregion Step: Prevent Concurrent Animations

		#region Step: Validate UI Objects
			# Ensure the necessary UI elements (main form, settings form) exist.
			if (-not ($global:DashboardConfig.UI -and $global:DashboardConfig.UI.SettingsForm -and $global:DashboardConfig.UI.MainForm))
			{
				Write-Verbose "  UI: Cannot show settings form - UI objects missing." -ForegroundColor Red
				return
			}
		#endregion Step: Validate UI Objects

		$settingsForm = $global:DashboardConfig.UI.SettingsForm

		#region Step: Position and Show Settings Form
			# Only proceed if the form is not already fully opaque (or nearly so).
			if ($settingsForm.Opacity -lt 0.95)
			{
				# Make the form visible before starting the fade.
				$settingsForm.Visible = $true

				# Calculate optimal position for settings form relative to the main form, keeping it on screen.
				$mainFormLocation = $global:DashboardConfig.UI.MainForm.Location
				$settingsFormWidth = $settingsForm.Width
				$settingsFormHeight = $settingsForm.Height
				$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
				$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

				# Attempt to center the settings form over the main form.
				$x = $mainFormLocation.X + (($global:DashboardConfig.UI.MainForm.Width - $settingsFormWidth) / 2)
				$y = $mainFormLocation.Y + (($global:DashboardConfig.UI.MainForm.Height - $settingsFormHeight) / 2)

				# Ensure form stays within screen bounds, adding a small margin.
				$margin = 0
				$x = [Math]::Max($margin, [Math]::Min($x, $screenWidth - $settingsFormWidth - $margin))
				$y = [Math]::Max($margin, [Math]::Min($y, $screenHeight - $settingsFormHeight - $margin))

				# Apply position and bring the settings form to the front.
				$settingsForm.Location = New-Object System.Drawing.Point($x, $y)
				$settingsForm.BringToFront()
				$settingsForm.Activate() # Give focus to the settings form
			}
		#endregion Step: Position and Show Settings Form

		#region Step: Create and Start Fade-In Animation Timer
			# Dispose previous timer if it exists
			if ($script:fadeInTimer) { $script:fadeInTimer.Dispose() }

			# Create a timer to handle the fade-in effect.
			$script:fadeInTimer = New-Object System.Windows.Forms.Timer
			$script:fadeInTimer.Interval = 15 # Interval for opacity steps (milliseconds)
			$script:fadeInTimer.Add_Tick({
					# Check if form still exists and hasn't been disposed.
					if (-not $global:DashboardConfig.UI.SettingsForm -or $global:DashboardConfig.UI.SettingsForm.IsDisposed)
					{
						$script:fadeInTimer.Stop()
						$script:fadeInTimer.Dispose()
						$script:fadeInTimer = $null
						return
					}

					# Increase opacity gradually until it reaches 1.
					if ($global:DashboardConfig.UI.SettingsForm.Opacity -lt 1)
					{
						$global:DashboardConfig.UI.SettingsForm.Opacity += 0.1
					}
					else
					{
						# Stop and dispose the timer once fully opaque.
						$global:DashboardConfig.UI.SettingsForm.Opacity = 1 # Ensure exactly 1
						$script:fadeInTimer.Stop()
						$script:fadeInTimer.Dispose()
						$script:fadeInTimer = $null
					}
				})
			# Start the fade-in timer.
			$script:fadeInTimer.Start()
			# Store timer reference for potential cleanup later.
			$global:DashboardConfig.Resources.Timers['fadeInTimer'] = $script:fadeInTimer
		#endregion Step: Create and Start Fade-In Animation Timer
	}
#endregion Function: Show-SettingsForm

#region Function: Hide-SettingsForm
	function Hide-SettingsForm
	{
		<#
		.SYNOPSIS
			Hides the settings form with a fade-out animation effect.
		.NOTES
			Initiates a timer-based fade-out by gradually decreasing opacity.
			Hides the form completely once opacity reaches zero.
			Prevents concurrent fade animations.
		#>
		[CmdletBinding()]
		param()

		#region Step: Prevent Concurrent Animations
			# Check if a fade-in or fade-out animation is already in progress.
			if (($script:fadeInTimer -and $script:fadeInTimer.Enabled) -or
				($global:fadeOutTimer -and $global:fadeOutTimer.Enabled))
			{
				return # Exit if an animation is active
			}
		#endregion Step: Prevent Concurrent Animations

		#region Step: Validate UI Objects
			# Ensure the settings form object exists.
			if (-not ($global:DashboardConfig.UI -and $global:DashboardConfig.UI.SettingsForm))
			{
				 Write-Verbose "  UI: Cannot hide settings form - UI object missing." -ForegroundColor Red
				return
			}
		#endregion Step: Validate UI Objects

		#region Step: Create and Start Fade-Out Animation Timer
			 # Dispose previous timer if it exists
			if ($global:fadeOutTimer) { $global:fadeOutTimer.Dispose() }

			# Create a timer to handle the fade-out effect.
			$global:fadeOutTimer = New-Object System.Windows.Forms.Timer
			$global:fadeOutTimer.Interval = 15 # Interval for opacity steps (milliseconds)
			$global:fadeOutTimer.Add_Tick({
					# Check if form still exists and hasn't been disposed.
					if (-not $global:DashboardConfig.UI.SettingsForm -or $global:DashboardConfig.UI.SettingsForm.IsDisposed)
					{
						$global:fadeOutTimer.Stop()
						$global:fadeOutTimer.Dispose()
						$global:fadeOutTimer = $null
						return
					}

					# Decrease opacity gradually until it reaches 0.
					if ($global:DashboardConfig.UI.SettingsForm.Opacity -gt 0)
					{
						$global:DashboardConfig.UI.SettingsForm.Opacity -= 0.1
					}
					else
					{
						# Stop the timer, ensure opacity is 0, and hide the form.
						$global:DashboardConfig.UI.SettingsForm.Opacity = 0 # Ensure exactly 0
						$global:fadeOutTimer.Stop()
						$global:fadeOutTimer.Dispose()
						$global:fadeOutTimer = $null
						$global:DashboardConfig.UI.SettingsForm.Hide()
					}
				})
			# Start the fade-out timer.
			$global:fadeOutTimer.Start()
			# Store timer reference for potential cleanup later.
			$global:DashboardConfig.Resources.Timers['fadeOutTimer'] = $global:fadeOutTimer
		#endregion Step: Create and Start Fade-Out Animation Timer
	}
#endregion Function: Hide-SettingsForm

#region Function: Set-UIElement
	function Set-UIElement
	{
		<#
		.SYNOPSIS
			Creates and configures various System.Windows.Forms UI elements based on provided parameters.
		.PARAMETER type
			[string] The type of UI element to create. Valid values: 'Form', 'Panel', 'Button', 'Label', 'DataGridView', 'TextBox', 'ComboBox', 'CheckBox', 'Toggle'. (Mandatory)
		.PARAMETER visible
			[bool] Sets the initial visibility of the element.
		.PARAMETER width
			[int] Sets the width of the element in pixels.
		.PARAMETER height
			[int] Sets the height of the element in pixels.
		.PARAMETER top
			[int] Sets the top position (Y-coordinate) of the element relative to its container.
		.PARAMETER left
			[int] Sets the left position (X-coordinate) of the element relative to its container.
		.PARAMETER bg
			[array] Sets the background color using an RGB or ARGB array (e.g., @(30,30,30) or @(255,0,0,128)).
		.PARAMETER fg
			[array] Sets the foreground (text) color using an RGB array (e.g., @(255,255,255)).
		.PARAMETER id
			[string] An identifier string (not directly used by WinForms, but useful for referencing in the $global:DashboardConfig.UI object).
		.PARAMETER text
			[string] Sets the text content or caption of the element (e.g., button text, label text, form title).
		.PARAMETER fs
			[System.Windows.Forms.FlatStyle] Sets the FlatStyle for elements like Buttons and ComboBoxes (e.g., 'Flat', 'Standard').
		.PARAMETER font
			[System.Drawing.Font] Sets the font for the element's text.
		.PARAMETER startPosition
			[string] For Forms, sets the initial starting position (e.g., 'Manual', 'CenterScreen').
		.PARAMETER formBorderStyle
			[int] For Forms, sets the border style using the System.Windows.Forms.FormBorderStyle enumeration value. Defaults to 'None'.
		.PARAMETER opacity
			[double] For Forms, sets the opacity level (0.0 to 1.0). Defaults to 1.0.
		.PARAMETER topMost
			[bool] For Forms, sets whether the form should stay on top of other windows.
        .PARAMETER checked
            [bool] For CheckBox or Toggle, sets the initial checked state.
		.PARAMETER multiline
			[switch] For TextBoxes, enables multi-line input.
		.PARAMETER readOnly
			[switch] For TextBoxes or DataGridViews, makes the content read-only.
		.PARAMETER scrollBars
			[switch] For TextBoxes, enables vertical scrollbars (if $multiline is also true).
		.PARAMETER dropDownStyle
			[string] For ComboBoxes, sets the style (e.g., 'Simple', 'DropDown', 'DropDownList'). Defaults to 'DropDownList'.
		.OUTPUTS
			[System.Windows.Forms.Control] Returns the created and configured UI element object.
		.NOTES
			Provides a standardized way to create common UI elements with consistent styling for the dark theme.
			Includes specific configurations for DataGridViews and custom drawing logic for Buttons, TextBoxes, and ComboBoxes
			to ensure visual consistency. Uses a DarkComboBox custom class for better ComboBox styling.
		#>
		[CmdletBinding()]
		param(
			[Parameter(Mandatory=$true)]
			[ValidateSet('Form', 'Panel', 'Button', 'Label', 'DataGridView', 'TextBox', 'ComboBox', 'CheckBox', 'Toggle')]
			[string]$type,
			[bool]$visible,
			[int]$width,
			[int]$height,
			[int]$top,
			[int]$left,
			[array]$bg,
			[array]$fg,
			[string]$id, # Used for referencing, not a direct WinForms property
			[string]$text,
			[System.Windows.Forms.FlatStyle]$fs,
			[System.Drawing.Font]$font,
			[string]$startPosition,
			[int]$formBorderStyle = [System.Windows.Forms.FormBorderStyle]::None,
			[double]$opacity = 1.0,
			[bool]$topMost,
            [bool]$checked,
			[switch]$multiline,
			[switch]$readOnly,
			[switch]$scrollBars,
			[ValidateSet('Simple', 'DropDown', 'DropDownList')]
			[string]$dropDownStyle = 'DropDownList'
		)

		#region Step: Create UI Element Based on Type
			# Create the appropriate .NET Windows Forms control object based on the $type parameter.
			$el = switch ($type)
			{
				'Form'         { New-Object System.Windows.Forms.Form }
				'Panel'        { New-Object System.Windows.Forms.Panel }
				'Button'       { New-Object System.Windows.Forms.Button }
				'Label'        { New-Object System.Windows.Forms.Label }
				'DataGridView' { New-Object System.Windows.Forms.DataGridView }
				'TextBox'      { New-Object System.Windows.Forms.TextBox }
				'ComboBox'     { New-Object System.Windows.Forms.ComboBox }
				'CheckBox'     { New-Object System.Windows.Forms.CheckBox }
				'Toggle'       { New-Object Custom.Toggle }
				default        { throw "Invalid element type specified: $type" }
			}
		#endregion Step: Create UI Element Based on Type

		#region Step: Configure DataGridView Specific Properties
			# Apply settings specific to DataGridView controls for appearance and behavior.
			if ($type -eq 'DataGridView')
			{
				$el.AllowUserToAddRows = $false          # Don't allow users to add new rows directly
				$el.ReadOnly = $false                    # Make the grid read-only
				$el.AllowUserToOrderColumns = $true      # Make the grid columns dragable
				$el.AllowUserToResizeColumns  = $false 	 # Make the grid columns size fixed
				$el.AllowUserToResizeRows = $false 		 # Make the grid rows size fixed
				$el.RowHeadersVisible = $false           # Hide the row header column
				$el.MultiSelect = $true                  # Allow selecting multiple rows
				$el.SelectionMode = 'FullRowSelect'      # Select entire rows instead of individual cells
				$el.AutoSizeColumnsMode = 'Fill'         # Make columns fill the available width
				$el.BorderStyle = 'FixedSingle'          # Adds the outer border
				$el.EnableHeadersVisualStyles = $false   # Allow custom header styling
				$el.CellBorderStyle = 'SingleHorizontal' # Horizontal lines between rows
				$el.ColumnHeadersBorderStyle = 'Single'  # No border around column headers
				$el.EditMode = 'EditProgrammatically'	 # Allows editing values on specific occasions
				$el.ColumnHeadersHeightSizeMode = 'DisableResizing'
				$el.RowHeadersWidthSizeMode = 'DisableResizing'
				$el.DefaultCellStyle.Alignment = 'MiddleCenter'
				$el.ColumnHeadersDefaultCellStyle.Alignment = 'MiddleCenter'

				# Set colors for better visibility in dark theme
				$el.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)    # Dark cell background
				$el.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 37)    # Dark cell background
				$el.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230) # Light text
				$el.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50) # Slightly darker header background
				$el.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240) # White header text
				$el.GridColor = [System.Drawing.Color]::FromArgb(70, 70, 70)                     # Color for grid lines
				$el.BackgroundColor = [System.Drawing.Color]::FromArgb(40, 40, 40)               # Background if grid is empty
				$el.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(60, 80, 180) # Selection background color (blueish)
				$el.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240) # White selected text

				# Add default columns expected by the application
				$el.Columns.AddRange(
					(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Index'; HeaderText = '#'; FillWeight = 8; SortMode = 'NotSortable';}),
					(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Titel'; HeaderText = 'Titel'; SortMode = 'NotSortable';}),
					(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'ID'; HeaderText = 'ID'; FillWeight = 20; SortMode = 'NotSortable';}),
					(New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'State'; HeaderText = 'State'; FillWeight = 40; SortMode = 'NotSortable';})
				)
			}
		#endregion Step: Configure DataGridView Specific Properties

		#region Step: Apply Common Control Properties
			# Apply properties common to most System.Windows.Forms.Control types.
			if ($el -is [System.Windows.Forms.Control])
			{
				if ($PSBoundParameters.ContainsKey('visible')) { $el.Visible = $visible }
				if ($PSBoundParameters.ContainsKey('width'))   { $el.Width = $width }
				if ($PSBoundParameters.ContainsKey('height'))  { $el.Height = $height }
				if ($PSBoundParameters.ContainsKey('top'))     { $el.Top = $top }
				if ($PSBoundParameters.ContainsKey('left'))    { $el.Left = $left }

				# Set background color from RGB or ARGB array
				if ($bg -is [array] -and $bg.Count -ge 3)
				{
					$el.BackColor = if ($bg.Count -eq 4) { [System.Drawing.Color]::FromArgb($bg[0], $bg[1], $bg[2], $bg[3]) }
									else                 { [System.Drawing.Color]::FromArgb($bg[0], $bg[1], $bg[2]) }
				}

				# Set foreground color from RGB array
				if ($fg -is [array] -and $fg.Count -ge 3)
				{
					$el.ForeColor = [System.Drawing.Color]::FromArgb($fg[0], $fg[1], $fg[2])
				}

				# Set font if provided
				if ($PSBoundParameters.ContainsKey('font')) { $el.Font = $font }
			}
		#endregion Step: Apply Common Control Properties

		#region Step: Apply Type-Specific Properties
			# Apply properties specific to the created element type.
			switch ($type)
			{
				'Form' {
					if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
					if ($PSBoundParameters.ContainsKey('startPosition')) {
						try { $el.StartPosition = [System.Windows.Forms.FormStartPosition]::$startPosition }
						catch { Write-Verbose "  UI: Invalid StartPosition value: $startPosition. Using default." -ForegroundColor Yellow }
					}
					if ($PSBoundParameters.ContainsKey('formBorderStyle')) {
						 try { $el.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]$formBorderStyle }
						 catch { Write-Verbose "  UI: Invalid FormBorderStyle value: $formBorderStyle. Using default." -ForegroundColor Yellow }
					}
					if ($PSBoundParameters.ContainsKey('opacity')) { $el.Opacity = [double]$opacity }
					if ($PSBoundParameters.ContainsKey('topMost')) { $el.TopMost = $topMost }
					if ($PSBoundParameters.ContainsKey('icon')) { $el.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSCommandPath) }
				}
				'Button' {
					if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
					if ($PSBoundParameters.ContainsKey('fs')) {
						$el.FlatStyle = $fs
						# Apply custom appearance for flat buttons to match dark theme
						$el.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60, 60, 60) # Subtle border
						$el.FlatAppearance.BorderSize = 1
						$el.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(70, 70, 70) # Slightly lighter on hover
						$el.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(90, 90, 90) # Even lighter when clicked

						# Custom Paint handler for more complex drawing.
						$el.Add_Paint({
							param($src, $e)
						
							# Only custom paint if we're using flat style
							if ($src.FlatStyle -eq [System.Windows.Forms.FlatStyle]::Flat)
							{
								# Draw the button background
								$bgBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(40, 40, 40))
								$e.Graphics.FillRectangle($bgBrush, 0, 0, $src.Width, $src.Height)
							
								# Draw text
								$textBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(240, 240, 240))
								$textFormat = [System.Drawing.StringFormat]::new()
								$textFormat.Alignment = [System.Drawing.StringAlignment]::Center
								$textFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
								$e.Graphics.DrawString($src.Text, $src.Font, $textBrush, 
									[System.Drawing.RectangleF]::new(0, 0, $src.Width, $src.Height), $textFormat)
							
								# Draw border
								$borderPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(60, 60, 60))
								$e.Graphics.DrawRectangle($borderPen, 0, 0, $src.Width, $src.Height)
							
								# Dispose resources
								$bgBrush.Dispose()
								$textBrush.Dispose()
								$borderPen.Dispose()
								$textFormat.Dispose()
							}
						})
					}
				}
				'Label' {
					if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
					# Ensure labels with transparent backgrounds are handled correctly
					if ($el.BackColor -eq [System.Drawing.Color]::Transparent) {
					   # May need additional handling depending on container if transparency issues arise
					}
				}
				'TextBox' {
					if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
					if ($PSBoundParameters.ContainsKey('multiline')) { $el.Multiline = $multiline }
					if ($PSBoundParameters.ContainsKey('readOnly')) { $el.ReadOnly = $readOnly }
					if ($PSBoundParameters.ContainsKey('scrollBars')) {
						$el.ScrollBars = if ($scrollBars -and $multiline) { [System.Windows.Forms.ScrollBars]::Vertical }
										 else { [System.Windows.Forms.ScrollBars]::None }
					}

					# Apply dark theme styling to TextBox
					$el.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
					$el.TextAlign = "Center"
					$el.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50) # Slightly lighter than background
					$el.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)

				}
				'ComboBox'
				{
					if ($null -ne $dropDownStyle)
					{
						try
						{
							$el.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::$dropDownStyle 
						}
						catch
						{
							Write-Verbose "UI: Invalid DropDownStyle value: $dropDownStyle. Using default." -ForegroundColor Yellow
						}
					}
					if ($null -ne $fs)
					{
						$el.FlatStyle = $fs
						$el.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
						
						# Set properties to show all items without scrollbar
						$el.IntegralHeight = $false
						
						# Store the original event handlers before creating the custom control
						$originalDrawItemScript = {
							param($src, $e)
							
							$e.DrawBackground()
							
							if ($e.Index -ge 0)
							{
								$brushBackground = if ($e.State -band [System.Windows.Forms.DrawItemState]::Selected)
								{
									[System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(40, 40, 40))
								}
								else
								{
									[System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(40, 40, 40))
								}
								
								$e.Graphics.FillRectangle($brushBackground, $e.bounds.Left, $e.bounds.Top, $e.bounds.Width, $e.bounds.Height)
								$e.Graphics.DrawString($src.Items[$e.Index].ToString(), $src.Font, [System.Drawing.Brushes]::FromArgb(240, 240, 240), $e.Bounds.Left, $e.Bounds.Top, $e.bounds.Width, $e.bounds.Height)
							}
							
							$e.DrawFocusRectangle()
						}
						
						$originalDropDownScript = {
							param($src, $e)
							
							# Calculate height needed for all items
							$itemHeight = $src.ItemHeight
							$totalItems = $src.Items.Count
							$requiredHeight = $itemHeight * $totalItems
							
							# Set dropdown height to show all items (max 300px to prevent extremely large dropdowns)
							$src.DropDownHeight = [Math]::Min($requiredHeight + 2, 300)
						}
						
						# Add the event handlers to the original control
						$el.Add_DrawItem($originalDrawItemScript)
						$el.Add_DropDown($originalDropDownScript)
													
						# Create a new instance of our custom ComboBox
						$customComboBox = New-Object Custom.DarkComboBox
						
						# Copy properties from the original ComboBox
						$customComboBox.Location = $el.Location
						$customComboBox.Size = $el.Size
						$customComboBox.Width = $el.Width - 20
						$customComboBox.DropDownStyle = $el.DropDownStyle
						$customComboBox.FlatStyle = $el.FlatStyle
						$customComboBox.DrawMode = $el.DrawMode
						$customComboBox.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
						$customComboBox.ForeColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
						$customComboBox.Font = $el.Font
						$customComboBox.IntegralHeight = $false
						$customComboBox.TabIndex = $el.TabIndex
						$customComboBox.Name = $el.Name
						
						# Copy any items from the original ComboBox
						foreach ($item in $el.Items)
						{
							$customComboBox.Items.Add($item)
						}
						
						# Add the same event handlers to the new control
						$customComboBox.Add_DropDown($originalDropDownScript)
						
						# Return the custom ComboBox
						$el = $customComboBox
					}
				}
                'CheckBox' {
                    if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
                    if ($PSBoundParameters.ContainsKey('checked')) { $el.Checked = $checked }
                    # Set modern flat style for dark theme
                    $el.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $el.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
                    $el.FlatAppearance.BorderSize = 1
                    $el.FlatAppearance.CheckedBackColor = [System.Drawing.Color]::FromArgb(0, 120, 215) # Windows blue
                    $el.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
                    $el.UseVisualStyleBackColor = $false
                    $el.CheckAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $el.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $el.Padding = [System.Windows.Forms.Padding]::new(20, 0, 0, 0) # Space between checkbox and text
                }
                'Toggle' {
                    if ($PSBoundParameters.ContainsKey('text')) { $el.Text = $text }
                    if ($PSBoundParameters.ContainsKey('checked')) { $el.Checked = $checked }
                }
			}
		#endregion Step: Apply Type-Specific Properties

		#region Step: Return Created UI Element
			return $el
		#endregion Step: Return Created UI Element
	}
#endregion Function: Set-UIElement

#endregion Core UI Functions

#region Module Exports
#region Step: Export Public Functions
	# Export the functions intended for use by other modules or the main script.
	Export-ModuleMember -Function Initialize-UI, Set-UIElement, Show-SettingsForm, Hide-SettingsForm, Sync-ConfigToUI, Sync-UIToConfig, Register-UIEventHandlers
#endregion Step: Export Public Functions
#endregion Module Exports