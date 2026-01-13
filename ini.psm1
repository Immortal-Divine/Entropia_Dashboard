<# ini.psm1 #>

#region Helper Functions

function CopyOrderedDictionary
{

	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.Collections.IDictionary]$Dictionary
	)

	$copy = [ordered]@{}
	try
	{
		foreach ($key in $Dictionary.Keys)
		{
			if ($Dictionary[$key] -is [System.Collections.IDictionary])
			{
				$copy[$key] = CopyOrderedDictionary -Dictionary $Dictionary[$key]
			}
			else
			{
				$copy[$key] = $Dictionary[$key]
			}
		}
	}
	catch
	{
		Write-Verbose "  Failed to copy dictionary element with key '$key': $_"

		throw 'Failed to complete dictionary copy.'
	}
	return $copy
}
function GetIniFileContent
{

	[CmdletBinding()]
	[OutputType([ordered])]
	param(
		[Parameter(Mandatory = $true)]
		[string]$IniPath
	)

	$result = [ordered]@{}
	try
	{
		if (-not (Test-Path -Path $IniPath -PathType Leaf))
		{
			Write-Verbose "  INI: (GetIniFileContent): File not found at '$IniPath'."
			return $result
		}


		$iniHandler = New-Object -TypeName IniFile -ArgumentList $IniPath

		$iniContent = $iniHandler.ReadIniFile()

		if ($null -ne $iniContent)
		{
			foreach ($section in $iniContent.Keys)
			{
				$result[$section] = [ordered]@{}
				$sectionDict = $iniContent[$section]

				if ($sectionDict -is [System.Collections.IDictionary])
				{
					foreach ($key in $sectionDict.Keys)
					{
						$result[$section][$key] = $sectionDict[$key]
					}
				}
				else
				{
					Write-Verbose "  INI: (GetIniFileContent): Section '[$section]' in '$IniPath' does not contain expected key-value pairs."
				}
			}
		}
	}
	catch
	{
		Write-Verbose "  INI: (GetIniFileContent): Failed to read INI file '$IniPath'. Error: $($_.Exception.Message)"
	}
	return $result
}
function LoadDefaultConfigOnError
{

	param(
		[Parameter(Mandatory = $true)]
		[string]$Reason
	)

	try
	{
		Write-Verbose "  INI: (ReadConfig): Loading default configuration because: $Reason"
		$global:DashboardConfig.Config = CopyOrderedDictionary -Dictionary $global:DashboardConfig.DefaultConfig -ErrorAction Stop
		return $false
	}
	catch
	{
		Write-Verbose "  INI: (ReadConfig): CRITICAL ERROR - Failed even to load default configuration! Error: $($_.Exception.Message)"

		$global:DashboardConfig.Config = [ordered]@{}
		return $false
	}
}

#endregion

#region Core Configuration Functions

function InitializeIniConfig
{
	[CmdletBinding()]
	[OutputType([bool])]
	param()

	Write-Verbose '  INI: (InitializeIniConfig): Initializing configuration...'

	if (-not $global:DashboardConfig.Paths.Ini)
	{
		Write-Verbose '  INI: (InitializeIniConfig): INI path not defined in global config.'
		return $false
	}

	$iniDir = Split-Path -Path $global:DashboardConfig.Paths.Ini -Parent
	if (-not (Test-Path -Path $iniDir -PathType Container))
	{
		try
		{
			New-Item -Path $iniDir -ItemType Directory -Force | Out-Null
			Write-Verbose "  INI: (InitializeIniConfig): Created config directory: $iniDir"
		}
		catch
		{
			Write-Verbose "  INI: (InitializeIniConfig): Failed to create config directory. Error: $($_.Exception.Message)"
			return $false
		}
	}

	$configLoaded = $false
	if (Test-Path -Path $global:DashboardConfig.Paths.Ini -PathType Leaf)
	{
		if (ReadConfig -ConfigPath $global:DashboardConfig.Paths.Ini)
		{
			Write-Verbose '  INI: (InitializeIniConfig): Successfully read existing config.'
			$configLoaded = $true
		}
		else
		{
			Write-Verbose '  INI: (InitializeIniConfig): Failed to read existing config. Loading defaults.'
		}
	}
 else
	{
		Write-Verbose '  INI: (InitializeIniConfig): Config file not found. Loading defaults.'
	}

	if (-not $configLoaded)
	{
		Write-Verbose "  INI: (InitializeIniConfig): No $global:DashboardConfig.Config"
		$global:DashboardConfig.Config = CopyOrderedDictionary -Dictionary $global:DashboardConfig.DefaultConfig

	}

	$configChanged = $false

	if (-not $global:DashboardConfig.DefaultConfig)
	{
		Write-Verbose '  INI: (InitializeIniConfig): DefaultConfig is missing. Skipping verification.'
		return $true
	}

	Write-Verbose '  INI: (InitializeIniConfig): Verifying config structure against defaults...'

	foreach ($sectionName in $global:DashboardConfig.DefaultConfig.Keys)
	{
		if (-not $global:DashboardConfig.Config.Contains($sectionName))
		{
			$global:DashboardConfig.Config[$sectionName] = [ordered]@{}
			Write-Verbose "  INI: (InitializeIniConfig): Adding missing section '[$sectionName]'."
			$configChanged = $true
		}

		$defaultSection = $global:DashboardConfig.DefaultConfig[$sectionName]
		if ($defaultSection -is [System.Collections.IDictionary])
		{
			foreach ($keyName in $defaultSection.Keys)
			{
				if ($sectionName -eq 'LoginConfig')
				{
					$existsInDefault = ($global:DashboardConfig.Config[$sectionName].Contains('Default') -and $global:DashboardConfig.Config[$sectionName]['Default'].Contains($keyName))

					if (-not $existsInDefault)
					{
						if (-not $global:DashboardConfig.Config[$sectionName].Contains('Default'))
						{
							$global:DashboardConfig.Config[$sectionName]['Default'] = [ordered]@{}
						}
						$global:DashboardConfig.Config[$sectionName]['Default'][$keyName] = $defaultSection[$keyName]
						Write-Verbose "  INI: (InitializeIniConfig): Adding missing key '$keyName' to 'Default' profile in section '[LoginConfig]'."
						$configChanged = $true
					}
				}
				elseif (-not $global:DashboardConfig.Config[$sectionName].Contains($keyName))
				{
					$global:DashboardConfig.Config[$sectionName][$keyName] = $defaultSection[$keyName]
					Write-Verbose "  INI: (InitializeIniConfig): Adding missing key '$keyName' in section '[$sectionName]' from defaults."
					$configChanged = $true
				}
			}
		}
	}

	if ($configChanged)
	{
		Write-Verbose '  INI: (InitializeIniConfig): Configuration updated with missing defaults. Writing changes...'
		WriteConfig -ConfigPath $global:DashboardConfig.Paths.Ini
	}

	return $true
}

function ReadConfig
{
	[CmdletBinding()]
	[OutputType([bool])]
	param(
		[Parameter()]
		[string]$ConfigPath
	)

	if ([string]::IsNullOrWhiteSpace($ConfigPath))
	{
		$ConfigPath = $global:DashboardConfig.Paths.Ini
	}

	if ([string]::IsNullOrWhiteSpace($ConfigPath))
	{
		Write-Verbose '  INI: (ReadConfig): Configuration path is not defined.'
		return (LoadDefaultConfigOnError -Reason 'Configuration path not defined')
	}
	Write-Verbose "  INI: (ReadConfig): Reading config from '$ConfigPath'"

	try
	{
		if (-not (Test-Path -Path $ConfigPath -PathType Leaf))
		{
			Write-Verbose "  INI: (ReadConfig): Config file not found at '$ConfigPath'."
			return (LoadDefaultConfigOnError -Reason 'Config file not found')
		}

		$iniHandler = [Custom.IniFile]::new($ConfigPath)
		$readConfig = $iniHandler.ReadIniFile()

		$global:DashboardConfig.Config = [ordered]@{}

		if ($null -ne $readConfig)
		{
			foreach ($section in $readConfig.Keys)
			{
				if ($section -eq 'LoginConfig')
				{
					$global:DashboardConfig.Config[$section] = [ordered]@{}
					$sourceSection = $readConfig[$section]

					if ($sourceSection -is [System.Collections.IDictionary])
					{
						foreach ($flattenedKey in $sourceSection.Keys)
						{
							if ($flattenedKey -match '^([^_]+)_(.+)$')
							{
								$profileName = $Matches[1]
								$settingName = $Matches[2]
								$settingValue = $sourceSection[$flattenedKey]

								if (-not $global:DashboardConfig.Config[$section].Contains($profileName))
								{
									$global:DashboardConfig.Config[$section][$profileName] = [ordered]@{}
								}
								$global:DashboardConfig.Config[$section][$profileName][$settingName] = $settingValue
							}
							else
							{
								$profileName = 'Default'
								$settingName = $flattenedKey
								$settingValue = $sourceSection[$flattenedKey]

								if (-not $global:DashboardConfig.Config[$section].Contains($profileName))
								{
									$global:DashboardConfig.Config[$section][$profileName] = [ordered]@{}
								}

								if (-not $global:DashboardConfig.Config[$section][$profileName].Contains($settingName))
								{
									$global:DashboardConfig.Config[$section][$profileName][$settingName] = $settingValue
								}
							}
						}
					}
				}
				else
				{
					$global:DashboardConfig.Config[$section] = [ordered]@{}
					if ($readConfig[$section] -is [System.Collections.IDictionary])
					{
						foreach ($key in $readConfig[$section].Keys)
						{
							$global:DashboardConfig.Config[$section][$key] = $readConfig[$section][$key]
						}
					}
				}
			}
		}
		else
		{
			Write-Verbose '  INI: (ReadConfig): Reading returned null.'
			return (LoadDefaultConfigOnError -Reason 'Reading config returned null')
		}

		Write-Verbose '  INI: (ReadConfig): Config loaded successfully.'
		return $true
	}
	catch
	{
		Write-Verbose "  INI: (ReadConfig): Failed to read/process config. Error: $($_.Exception.Message)"
		return (LoadDefaultConfigOnError -Reason "Exception: $($_.Exception.Message)")
	}
}

function WriteConfig
{
	[CmdletBinding(SupportsShouldProcess = $true)]
	[OutputType([bool])]
	param(
		[Parameter()]
		[System.Collections.IDictionary]$Config = $global:DashboardConfig.Config,

		[Parameter()]
		[string]$ConfigPath = $global:DashboardConfig.Paths.Ini
	)

	
	if ([string]::IsNullOrWhiteSpace($ConfigPath))
	{
		Write-Verbose '  INI: (WriteConfig): Configuration path is not defined.'
		return $false
	}
	if ($null -eq $Config)
	{
		Write-Verbose '  INI: (WriteConfig): Configuration data is null.'
		return $false
	}

	if (-not $pscmdlet.ShouldProcess($ConfigPath, 'Write Configuration to INI file'))
	{
		return $false
	}

	
	try
	{
		$configDir = Split-Path -Path $ConfigPath -Parent
		if (-not (Test-Path -Path $configDir -PathType Container))
		{
			Write-Verbose "  INI: (WriteConfig): Creating directory '$configDir'."
			$null = New-Item -ItemType Directory -Path $configDir -Force -ErrorAction Stop
		}
	}
	catch
	{
		Write-Verbose "  INI: (WriteConfig): Failed to create directory. Error: $($_.Exception.Message)"
		return $false
	}

	
	$configToWrite = New-Object System.Collections.Specialized.OrderedDictionary

	try
	{
		foreach ($sectionKey in $Config.Keys)
		{
			$sectionName = $sectionKey.ToString()
			$sectionValue = $Config[$sectionKey]

			if ($sectionValue -isnot [System.Collections.IDictionary])
			{
				Write-Verbose "  INI: (WriteConfig): Section '$sectionName' is not a dictionary, skipping."
				continue
			}

			$sectionData = New-Object System.Collections.Specialized.OrderedDictionary

			
			if ($sectionName -eq 'LoginConfig')
			{
				foreach ($profKey in $sectionValue.Keys)
				{
					$profileData = $sectionValue[$profKey]
					
					if ($profileData -is [System.Collections.IDictionary])
					{
						foreach ($settingKey in $profileData.Keys)
						{
							$flatKey = "${profKey}_$($settingKey)"
							$rawValue = $profileData[$settingKey]
							
							
							$preparedValue = if ($rawValue -is [Array]) { ($rawValue | Where-Object { $_ -ne $null }) -join ',' }
							else { if ($null -ne $rawValue) { $rawValue.ToString() } else { '' } }
							
							$sectionData.Add($flatKey, $preparedValue)
						}
					}
					else
					{
						
						$preparedValue = if ($profileData -is [Array]) { ($profileData | Where-Object { $_ -ne $null }) -join ',' }
						else { if ($null -ne $profileData) { $profileData.ToString() } else { '' } }
						$sectionData.Add($profKey.ToString(), $preparedValue)
					}
				}
			}
			
			else
			{
				foreach ($itemKey in $sectionValue.Keys)
				{
					$value = $sectionValue[$itemKey]
					
					
					$preparedValue = if ($value -is [Array]) { ($value | Where-Object { $_ -ne $null }) -join ',' }
					elseif ($value -is [System.Collections.IDictionary] -or $value -is [PSCustomObject]) { $value | ConvertTo-Json -Compress -Depth 5 }
					else { if ($null -ne $value) { $value.ToString() } else { '' } }
					
					$sectionData.Add($itemKey.ToString(), $preparedValue)
				}
			}

			$configToWrite.Add($sectionName, $sectionData)
		}
	}
	catch
	{
		Write-Verbose "  INI: (WriteConfig): Failed to prepare data. Error: $($_.Exception.Message)"
		return $false
	}

	
	try
	{
		$iniFile = [Custom.IniFile]::new($ConfigPath)
		$iniFile.WriteIniFile($configToWrite)

		Write-Verbose "  INI: (WriteConfig): Config written successfully to '$ConfigPath'."
		return $true
	}
	catch
	{
		Write-Verbose "  INI: (WriteConfig): C#Class failed to write file. Error: $($_.Exception.Message)"
		return $false
	}
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion