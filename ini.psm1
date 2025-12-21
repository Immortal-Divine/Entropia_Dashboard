<# ini.psm1
#>

#region Helper Functions

    function Copy-OrderedDictionary
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
                    $copy[$key] = Copy-OrderedDictionary -Dictionary $Dictionary[$key]
                }
                else
                {
                    $copy[$key] = $Dictionary[$key]
                }
            }
        }
        catch
        {
            Write-Verbose "  Failed to copy dictionary element with key '$key': $_" -ForegroundColor Red

            throw "Failed to complete dictionary copy."
        }
        return $copy
    }

    function Get-IniFileContent
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
                Write-Verbose "  INI: (Get-IniFileContent): File not found at '$IniPath'." -ForegroundColor Yellow
                return $result
            }

            if (-not ([System.Reflection.Assembly]::LoadWithPartialName('System.Collections.Specialized'))) {
                Add-Type -AssemblyName System.Collections.Specialized
            }
            if (-not ([System.Reflection.Assembly]::LoadWithPartialName('System.IO'))) {
                Add-Type -AssemblyName System.IO
            }


            $iniHandler = New-Object -TypeName IniFile -ArgumentList $IniPath

            $iniContent = $iniHandler.ReadIniFile()

            if ($null -ne $iniContent)
            {
                foreach ($section in $iniContent.Keys)
                {
                    $result[$section] = [ordered]@{}
                    $sectionDict = $iniContent[$section]

                    if ($sectionDict -is [System.Collections.IDictionary]) {
                        foreach ($key in $sectionDict.Keys)
                        {
                            $result[$section][$key] = $sectionDict[$key]
                        }
                    } else {
                        Write-Verbose "  INI: (Get-IniFileContent): Section '[$section]' in '$IniPath' does not contain expected key-value pairs." -ForegroundColor Yellow
                    }
                }
            }
        }
        catch
        {
            Write-Verbose "  INI: (Get-IniFileContent): Failed to read INI file '$IniPath'. Error: $($_.Exception.Message)" -ForegroundColor Red
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
            Write-Verbose "  INI: (Read-Config): Loading default configuration because: $Reason" -ForegroundColor Yellow
            $global:DashboardConfig.Config = Copy-OrderedDictionary -Dictionary $global:DashboardConfig.DefaultConfig -ErrorAction Stop
            return $false
        }
        catch
        {
            Write-Verbose "  INI: (Read-Config): CRITICAL ERROR - Failed even to load default configuration! Error: $($_.Exception.Message)" -ForegroundColor Red

            $global:DashboardConfig.Config = [ordered]@{}
            return $false
        }
    }

#endregion

#region Core Configuration Functions

    function Initialize-IniConfig
	{
		[CmdletBinding()]
		[OutputType([bool])]
		param()

		Write-Verbose "  INI: (Initialize-IniConfig): Initializing configuration..." -ForegroundColor Cyan

		if (-not $global:DashboardConfig.Paths.Ini) {
			Write-Verbose "  INI: (Initialize-IniConfig): INI path not defined in global config." -ForegroundColor Red
			return $false
		}

		$iniDir = Split-Path -Path $global:DashboardConfig.Paths.Ini -Parent
		if (-not (Test-Path -Path $iniDir -PathType Container)) {
			try {
				New-Item -Path $iniDir -ItemType Directory -Force | Out-Null
				Write-Verbose "  INI: (Initialize-IniConfig): Created config directory: $iniDir" -ForegroundColor Green
			} catch {
				Write-Verbose "  INI: (Initialize-IniConfig): Failed to create config directory. Error: $($_.Exception.Message)" -ForegroundColor Red
				return $false
			}
		}

		$configLoaded = $false
		if (Test-Path -Path $global:DashboardConfig.Paths.Ini -PathType Leaf) {
			if (Read-Config -ConfigPath $global:DashboardConfig.Paths.Ini) {
				Write-Verbose "  INI: (Initialize-IniConfig): Successfully read existing config." -ForegroundColor Green
				$configLoaded = $true
			} else {
				Write-Verbose "  INI: (Initialize-IniConfig): Failed to read existing config. Loading defaults." -ForegroundColor Yellow
			}
		} else {
			Write-Verbose "  INI: (Initialize-IniConfig): Config file not found. Loading defaults." -ForegroundColor Yellow
		}

		if (-not $configLoaded) {
			Write-Verbose "  INI: (Initialize-IniConfig): No $global:DashboardConfig.Config" -ForegroundColor Yellow
			$global:DashboardConfig.Config = Copy-OrderedDictionary -Dictionary $global:DashboardConfig.DefaultConfig

		}

		$configChanged = $false

		if (-not $global:DashboardConfig.DefaultConfig) {
			Write-Verbose "  INI: (Initialize-IniConfig): DefaultConfig is missing. Skipping verification." -ForegroundColor Yellow
			return $true
		}

		Write-Verbose "  INI: (Initialize-IniConfig): Verifying config structure against defaults..." -ForegroundColor DarkGray

		foreach ($sectionName in $global:DashboardConfig.DefaultConfig.Keys) {
			if (-not $global:DashboardConfig.Config.Contains($sectionName)) {
				$global:DashboardConfig.Config[$sectionName] = [ordered]@{}
				Write-Verbose "  INI: (Initialize-IniConfig): Adding missing section '[$sectionName]'." -ForegroundColor Yellow
				$configChanged = $true
			}

			$defaultSection = $global:DashboardConfig.DefaultConfig[$sectionName]
			if ($defaultSection -is [System.Collections.IDictionary]) {
				foreach ($keyName in $defaultSection.Keys) {
					if ($sectionName -eq 'LoginConfig') {
						$existsInDefault = ($global:DashboardConfig.Config[$sectionName].Contains('Default') -and $global:DashboardConfig.Config[$sectionName]['Default'].Contains($keyName))

						if (-not $existsInDefault) {
							if (-not $global:DashboardConfig.Config[$sectionName].Contains('Default')) {
								$global:DashboardConfig.Config[$sectionName]['Default'] = [ordered]@{}
							}
							$global:DashboardConfig.Config[$sectionName]['Default'][$keyName] = $defaultSection[$keyName]
							Write-Verbose "  INI: (Initialize-IniConfig): Adding missing key '$keyName' to 'Default' profile in section '[LoginConfig]'." -ForegroundColor Yellow
							$configChanged = $true
						}
					}
					elseif (-not $global:DashboardConfig.Config[$sectionName].Contains($keyName)) {
						$global:DashboardConfig.Config[$sectionName][$keyName] = $defaultSection[$keyName]
						Write-Verbose "  INI: (Initialize-IniConfig): Adding missing key '$keyName' in section '[$sectionName]' from defaults." -ForegroundColor Yellow
						$configChanged = $true
					}
				}
			}
		}

		if ($configChanged) {
			Write-Verbose "  INI: (Initialize-IniConfig): Configuration updated with missing defaults. Writing changes..." -ForegroundColor Cyan
			Write-Config -ConfigPath $global:DashboardConfig.Paths.Ini
		}

		return $true
	}

function Read-Config
{
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter()]
        [string]$ConfigPath
    )

    if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
        $ConfigPath = $global:DashboardConfig.Paths.Ini
    }

    if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
        Write-Verbose "  INI: (Read-Config): Configuration path is not defined." -ForegroundColor Red
        Return (LoadDefaultConfigOnError -Reason "Configuration path not defined")
    }
    Write-Verbose "  INI: (Read-Config): Reading config from '$ConfigPath'" -ForegroundColor Cyan

    try
    {
        if (-not (Test-Path -Path $ConfigPath -PathType Leaf)) {
            Write-Verbose "  INI: (Read-Config): Config file not found at '$ConfigPath'." -ForegroundColor Yellow
            Return (LoadDefaultConfigOnError -Reason "Config file not found")
        }

        $iniHandler = [Custom.IniFile]::new($ConfigPath)
        $readConfig = $iniHandler.ReadIniFile()

        $global:DashboardConfig.Config = [ordered]@{}

        if ($null -ne $readConfig) {
            foreach ($section in $readConfig.Keys)
            {
                if ($section -eq "LoginConfig") {
                    $global:DashboardConfig.Config[$section] = [ordered]@{}
                    $sourceSection = $readConfig[$section]

                    if ($sourceSection -is [System.Collections.IDictionary]) {
                        foreach ($flattenedKey in $sourceSection.Keys) {
                            if ($flattenedKey -match "^([^_]+)_(.+)$") {
                                $profileName = $Matches[1]
                                $settingName = $Matches[2]
                                $settingValue = $sourceSection[$flattenedKey]

                                if (-not $global:DashboardConfig.Config[$section].Contains($profileName)) {
                                    $global:DashboardConfig.Config[$section][$profileName] = [ordered]@{}
                                }
                                $global:DashboardConfig.Config[$section][$profileName][$settingName] = $settingValue
                                Write-Verbose "  INI: (Read-Config): Loaded LoginConfig - Profile '$profileName', Setting '$settingName'" -ForegroundColor DarkGray
                            }
                            else {
                                $profileName = "Default"
                                $settingName = $flattenedKey
                                $settingValue = $sourceSection[$flattenedKey]

                                if (-not $global:DashboardConfig.Config[$section].Contains($profileName)) {
                                    $global:DashboardConfig.Config[$section][$profileName] = [ordered]@{}
                                }

                                if (-not $global:DashboardConfig.Config[$section][$profileName].Contains($settingName)) {
                                    $global:DashboardConfig.Config[$section][$profileName][$settingName] = $settingValue
                                    Write-Verbose "  INI: (Read-Config): Mapped flat key '$settingName' to Profile 'Default'." -ForegroundColor DarkGray
                                }
                            }
                        }
                    }
                }
                else {
                    $global:DashboardConfig.Config[$section] = [ordered]@{}
                    if ($readConfig[$section] -is [System.Collections.IDictionary]) {
                        foreach ($key in $readConfig[$section].Keys)
                        {
                            $global:DashboardConfig.Config[$section][$key] = $readConfig[$section][$key]
                        }
                    }
                }
            }
        } else {
            Write-Verbose "  INI: (Read-Config): Reading returned null." -ForegroundColor Red
            Return (LoadDefaultConfigOnError -Reason "Reading config returned null")
        }

        Write-Verbose "  INI: (Read-Config): Config loaded successfully." -ForegroundColor Green
        return $true
    }
    catch
    {
        Write-Verbose "  INI: (Read-Config): Failed to read/process config. Error: $($_.Exception.Message)" -ForegroundColor Red
        Return (LoadDefaultConfigOnError -Reason "Exception: $($_.Exception.Message)")
    }
}

function Write-Config
{
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter()]
        [string]$ConfigPath
    )

    if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
        $ConfigPath = $global:DashboardConfig.Paths.Ini
    }

    Write-Verbose "  INI: (Write-Config): Saving configuration to '$ConfigPath'..." -ForegroundColor Cyan

    if (-not $global:DashboardConfig.Config) {
        Write-Verbose "  INI: (Write-Config): No configuration to save." -ForegroundColor Red
        return $false
    }

    try {
        $iniHandler = [Custom.IniFile]::new($ConfigPath)

        $configToSave = [ordered]@{}

        foreach ($section in $global:DashboardConfig.Config.Keys) {
            $configToSave[$section] = [ordered]@{}
            $memSection = $global:DashboardConfig.Config[$section]

            if ($section -eq 'LoginConfig') {
                foreach ($profKey in $memSection.Keys) {
                    $profileData = $memSection[$profKey]
                    if ($profileData -is [System.Collections.IDictionary]) {
                        foreach ($settingKey in $profileData.Keys) {
                            $flatKey = "${profKey}_${settingKey}"
                            $configToSave[$section][$flatKey] = $profileData[$settingKey]
                        }
                    }
                }
            }
            else {
                foreach ($key in $memSection.Keys) {
                    $configToSave[$section][$key] = $memSection[$key]
                }
            }
        }

        $iniHandler.WriteIniFile($configToSave)
        Write-Verbose "  INI: (Write-Config): Configuration saved successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Verbose "  INI: (Write-Config): Failed to save config. Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

#endregion

#region Module Exports
Export-ModuleMember -Function *
#endregion