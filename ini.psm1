<# ini.psm1
    .SYNOPSIS
        INI Configuration Management Module for Entropia Dashboard.
    .DESCRIPTION
        This PowerShell module provides functions to handle configuration file operations
        (reading and writing INI files) for the Entropia Dashboard application.
        It interacts with the global dashboard configuration state and utilizes the
        custom C# 'IniFile' class (defined in classes.psm1) for file parsing and writing.

        Key Functions:
        - Initialize-IniConfig: Ensures the main configuration file exists and is populated with necessary defaults.
        - Read-Config: Reads the configuration from the INI file into the global state.
        - Write-Config: Writes the current global configuration state back to the INI file.
        - Get-IniFileContent: Reads an arbitrary INI file into a PowerShell OrderedDictionary.
        - Copy-OrderedDictionary: Helper to deep-copy ordered dictionaries.
    .NOTES
        Author: Immortal / Divine
        Version: 1.2.1
        Requires:
        - PowerShell 5.1+
        - .NET Framework 4.5+
        - Entropia_Dashboard module 'classes.psm1' (for the IniFile C# class)
        - Entropia_Dashboard global variable '$global:DashboardConfig' (expected structure)

        Documentation Standards Followed:
        - Module Level Documentation: Synopsis, Description, Notes.
        - Function Level Documentation: Synopsis, Parameter Descriptions, Output Specifications.
        - Code Organization: Logical grouping using #region / #endregion. Functions organized by workflow.
        - Step Documentation: Code blocks enclosed in '#region Step: Description' / '#endregion Step: Description'.
        - Variable Definitions: Inline comments describing the purpose of significant variables.
        - Error Handling: Comprehensive try/catch/finally blocks with error logging and user notification.

        Relies heavily on the structure and availability of the $global:DashboardConfig variable.
        Ensure the IniFile C# class from 'classes.psm1' is loaded before using these functions.
#>

#region Helper Functions

    #region Function: Copy-OrderedDictionary
    function Copy-OrderedDictionary
    {
        <#
        .SYNOPSIS
            Recursively copies a System.Collections.Specialized.OrderedDictionary or a PowerShell [ordered] dictionary.
        .DESCRIPTION
            Creates a deep copy of an ordered dictionary, handling nested ordered dictionaries.
            This is useful for creating independent copies of configuration objects.
        .PARAMETER Dictionary
            [System.Collections.IDictionary] The ordered dictionary ([System.Collections.Specialized.OrderedDictionary] or [ordered]) to copy. (Mandatory)
        .OUTPUTS
            [ordered] A new PowerShell ordered dictionary containing a deep copy of the input.
        .NOTES
            Ensures nested dictionaries are also copied, not just referenced.
            Handles potential errors during the copy process.
        #>
        param (
            [Parameter(Mandatory = $true)]
            [ValidateNotNull()]
            [System.Collections.IDictionary]$Dictionary # Accepts both .NET and PS ordered dictionaries
        )

        #region Step: Initialize Output Dictionary
            # $copy: The new ordered dictionary to hold the copied data.
            $copy = [ordered]@{}
        #endregion Step: Initialize Output Dictionary

        #region Step: Iterate Through Keys and Copy Values
        try
        {
            foreach ($key in $Dictionary.Keys)
            {
                #region Step: Check for Nested Dictionary
                if ($Dictionary[$key] -is [System.Collections.IDictionary])
                {
                    #region Step: Recursively Copy Nested Dictionary
                    $copy[$key] = Copy-OrderedDictionary -Dictionary $Dictionary[$key]
                    #endregion Step: Recursively Copy Nested Dictionary
                }
                #endregion Step: Check for Nested Dictionary
                #region Step: Copy Simple Value
                else
                {
                    $copy[$key] = $Dictionary[$key]
                }
                #endregion Step: Copy Simple Value
            }
        }
        catch
        {
            #region Step: Error Handling - Failed to Copy Dictionary Element
            Write-Verbose "  Failed to copy dictionary element with key '$key': $_" -ForegroundColor Red
            # Consider re-throwing or returning null/empty depending on desired behavior
            throw "Failed to complete dictionary copy."
            #endregion Step: Error Handling - Failed to Copy Dictionary Element
        }
        #endregion Step: Iterate Through Keys and Copy Values

        #region Step: Return Copy
        return $copy
        #endregion Step: Return Copy
    }
    #endregion Function: Copy-OrderedDictionary

    #region Function: Get-IniFileContent
    function Get-IniFileContent
    {
        <#
        .SYNOPSIS
            Reads the content of a specified INI file into a PowerShell ordered dictionary.
        .DESCRIPTION
            Uses the C# 'IniFile' class (from classes.psm1) to read the structure and content of an INI file.
            It then converts the .NET OrderedDictionary returned by the class into a PowerShell [ordered] dictionary
            for easier manipulation within scripts.
        .PARAMETER IniPath
            [string] The full path to the INI file to be read. (Mandatory)
        .OUTPUTS
            [ordered] An ordered dictionary representing the INI file structure ([Section][Key] = Value).
            Returns an empty ordered dictionary if the file is not found or an error occurs during reading.
        .NOTES
            Requires the 'IniFile' class to be available (typically loaded via Add-Type from classes.psm1).
            Error messages are written to the host stream.
        #>
        [CmdletBinding()]
        [OutputType([ordered])]
        param(
            [Parameter(Mandatory = $true)]
            [string]$IniPath
        )

        #region Step: Initialize Result Dictionary
        # $result: The ordered dictionary to store the INI content.
        $result = [ordered]@{}
        #endregion Step: Initialize Result Dictionary

        #region Step: Read INI File using C# Class
        try
        {
            #region Step: Validate File Existence
            if (-not (Test-Path -Path $IniPath -PathType Leaf))
            {
                Write-Verbose "  INI: (Get-IniFileContent): File not found at '$IniPath'." -ForegroundColor Yellow
                return $result # Return empty dictionary
            }
            #endregion Step: Validate File Existence

            #region Step: Instantiate IniFile Handler
            # Ensure required .NET assemblies are loaded (though ideally done at module load)
            if (-not ([System.Reflection.Assembly]::LoadWithPartialName('System.Collections.Specialized'))) {
                Add-Type -AssemblyName System.Collections.Specialized
            }
            if (-not ([System.Reflection.Assembly]::LoadWithPartialName('System.IO'))) {
                Add-Type -AssemblyName System.IO
            }
            # Assuming IniFile class is loaded elsewhere (e.g., at module import)
            # $iniHandler: Instance of the C# class used to interact with the INI file.
            $iniHandler = New-Object -TypeName IniFile -ArgumentList $IniPath
            #endregion Step: Instantiate IniFile Handler

            #region Step: Read File Content
            # $iniContent: The raw content read from the INI file as a .NET OrderedDictionary.
            $iniContent = $iniHandler.ReadIniFile()
            #endregion Step: Read File Content

            #region Step: Convert .NET Dictionary to PowerShell Ordered Dictionary
            if ($null -ne $iniContent)
            {
                foreach ($section in $iniContent.Keys)
                {
                    $result[$section] = [ordered]@{}
                    $sectionDict = $iniContent[$section]
                    # Ensure the section's value is dictionary-like before iterating
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
            #endregion Step: Convert .NET Dictionary to PowerShell Ordered Dictionary
        }
        catch
        {
            #region Step: Error Handling - Failed to Read INI
            Write-Verbose "  INI: (Get-IniFileContent): Failed to read INI file '$IniPath'. Error: $($_.Exception.Message)" -ForegroundColor Red
            # Return empty dictionary on failure
            #endregion Step: Error Handling - Failed to Read INI
        }
        #endregion Step: Read INI File using C# Class

        #region Step: Return Result
        return $result
        #endregion Step: Return Result
    }
    #endregion Function: Get-IniFileContent

    #region Function: LoadDefaultConfigOnError
    function LoadDefaultConfigOnError
    {
        <#
        .SYNOPSIS
            Internal helper function to load default configuration when an error occurs during reading.
        .DESCRIPTION
            Called by Read-Config when it fails to read the specified INI file. It logs a warning
            with the reason and attempts to populate $global:DashboardConfig.Config with a copy
            of $global:DashboardConfig.DefaultConfig.
        .PARAMETER Reason
            [string] A description of the error that triggered loading defaults. (Mandatory)
        .OUTPUTS
            [bool] Always returns $false to indicate that the original read operation failed,
            even if loading defaults was successful. Returns $false also if loading defaults fails.
        .NOTES
            Modifies $global:DashboardConfig.Config.
            Uses Copy-OrderedDictionary.
            Logs errors if loading defaults also fails.
        #>
        param(
            [Parameter(Mandatory = $true)]
            [string]$Reason
        )

        #region Step: Log Warning and Attempt to Load Defaults
        try
        {
            Write-Verbose "  INI: (Read-Config): Loading default configuration because: $Reason" -ForegroundColor Yellow
            $global:DashboardConfig.Config = Copy-OrderedDictionary -Dictionary $global:DashboardConfig.DefaultConfig -ErrorAction Stop
            return $false # Indicate failure to read original, but defaults loaded
        }
        catch
        {
            #region Step: Error Handling - Critical Failure to Load Defaults
            Write-Verbose "  INI: (Read-Config): CRITICAL ERROR - Failed even to load default configuration! Error: $($_.Exception.Message)" -ForegroundColor Red
            # Ensure Config is at least an empty dictionary to prevent later errors
            $global:DashboardConfig.Config = [ordered]@{}
            return $false # Indicate critical failure
            #endregion Step: Error Handling - Critical Failure to Load Defaults
        }
        #endregion Step: Log Warning and Attempt to Load Defaults
    }
    #endregion Function: LoadDefaultConfigOnError

#endregion Helper Functions

#region Core Configuration Functions

    #region Function: Initialize-IniConfig
    function Initialize-IniConfig
    {
        <#
        .SYNOPSIS
            Initializes the global dashboard configuration ($global:DashboardConfig.Config).
        .DESCRIPTION
            This function ensures the dashboard's configuration is ready for use. It performs the following steps:
            1. Checks if the configuration INI file (defined in $global:DashboardConfig.Paths.Ini) exists.
            2. If the file doesn't exist, it creates a new one by copying the default configuration ($global:DashboardConfig.DefaultConfig) and writing it using Write-Config.
            3. If the file exists, it reads the configuration using Read-Config.
            4. It then verifies that all sections and keys defined in the default configuration are present in the loaded configuration. Missing sections or keys are added from the defaults.
            5. If any defaults were added, the updated configuration is written back to the file using Write-Config.
        .OUTPUTS
            [bool] Returns $true if initialization (including reading/writing the config file and verifying structure) completes successfully.
            Returns $false if any critical step fails (e.g., cannot write initial config, cannot read existing config, cannot write updated config).
        .NOTES
            - Modifies the $global:DashboardConfig.Config variable.
            - Relies heavily on the structure of $global:DashboardConfig (specifically .Paths.Ini, .Config, .DefaultConfig).
            - Depends on helper functions: Copy-OrderedDictionary, Read-Config, Write-Config.
            - Logs progress and outcomes to the host stream.
        #>
        [CmdletBinding()]
        [OutputType([bool])]
        param()

        #region Step: Log Initialization Start
        Write-Verbose '  INI: (Initialize-IniConfig): Initializing dashboard configuration...' -ForegroundColor Cyan
        #endregion Step: Log Initialization Start

        #region Step: Define Config Path
        # $configPath: The full path to the main INI configuration file.
        $configPath = $global:DashboardConfig.Paths.Ini
        if ([string]::IsNullOrWhiteSpace($configPath)) {
            Write-Verbose "  INI: (Initialize-IniConfig): Configuration path (\$global:DashboardConfig.Paths.Ini) is not defined." -ForegroundColor Red
            return $false
        }
        #endregion Step: Define Config Path

        #region Step: Check if Config File Exists
        if (-not (Test-Path -Path $configPath -PathType Leaf)) # Use Leaf to ensure it's a file, not a directory
        {
            #region Step: Config File Not Found - Create from Defaults
            Write-Verbose "  INI: (Initialize-IniConfig): Config file not found at '$configPath'. Creating from defaults." -ForegroundColor Yellow
            try
            {
                #region Step: Copy Default Config to Global Variable
                $global:DashboardConfig.Config = Copy-OrderedDictionary -Dictionary $global:DashboardConfig.DefaultConfig -ErrorAction Stop
                #endregion Step: Copy Default Config to Global Variable

                #region Step: Write Initial Config File
                $writeSuccess = Write-Config -ConfigPath $configPath -Config $global:DashboardConfig.Config
                if (-not $writeSuccess)
                {
                    Write-Verbose "  INI: (Initialize-IniConfig): Failed to write initial default configuration to '$configPath'." -ForegroundColor Red
                    # Ensure config is reset or nulled if write fails? For now, leave it as copied defaults.
                    return $false # Indicate initialization failure
                }
                Write-Verbose "  INI: (Initialize-IniConfig): Successfully created and wrote default config to '$configPath'." -ForegroundColor Green
                #endregion Step: Write Initial Config File
            }
            catch
            {
                #region Step: Error Handling - Failed to Create Default Config
                Write-Verbose "  INI: (Initialize-IniConfig): Error creating default configuration. Error: $($_.Exception.Message)" -ForegroundColor Red
                # Ensure config is at least an empty dictionary on catastrophic failure
                $global:DashboardConfig.Config = [ordered]@{}
                return $false # Indicate initialization failure
                #endregion Step: Error Handling - Failed to Create Default Config
            }
            #endregion Step: Config File Not Found - Create from Defaults
        }
        #endregion Step: Check if Config File Exists
        #region Step: Config File Exists - Read and Verify
        else
        {
            #region Step: Read Existing Config File
            Write-Verbose "  INI: (Initialize-IniConfig): Existing config file found at '$configPath'. Reading..." -ForegroundColor Cyan
            $readSuccess = Read-Config -ConfigPath $configPath # Pass path explicitly
            if (-not $readSuccess)
            {
                # Read-Config logs errors and attempts to load defaults on failure.
                Write-Verbose "  INI: (Initialize-IniConfig): Failed to read existing configuration file '$configPath'. Check previous errors. Initialization cannot continue reliably." -ForegroundColor Red
                # If Read-Config failed but loaded defaults, should we proceed? Assuming failure is critical here.
                return $false # Indicate initialization failure
            }
            Write-Verbose "  INI: (Initialize-IniConfig): Successfully read existing config." -ForegroundColor Green
            #endregion Step: Read Existing Config File

            #region Step: Verify Config Structure Against Defaults
            Write-Verbose "  INI: (Initialize-IniConfig): Verifying config structure against defaults..." -ForegroundColor Cyan
            # $needsUpdate: Flag indicating if the config file needs to be rewritten with added defaults.
            $needsUpdate = $false
            try
            {
                #region Step: Iterate Default Sections
                foreach ($section in $global:DashboardConfig.DefaultConfig.Keys)
                {
                    # Check if section exists in loaded config
                    if (-not $global:DashboardConfig.Config.Contains($section))
                    {
                        #region Step: Add Missing Section
                        Write-Verbose "  INI: (Initialize-IniConfig): Adding missing section '[$section]' from defaults." -ForegroundColor DarkGray
                        # Copy the entire default section
                        $global:DashboardConfig.Config[$section] = Copy-OrderedDictionary -Dictionary $global:DashboardConfig.DefaultConfig[$section] -ErrorAction Stop
                        $needsUpdate = $true
                        #endregion Step: Add Missing Section
                    }
                    else
                    {
                        # Section exists, check keys within the section
                        #region Step: Iterate Default Keys in Section
                        foreach ($key in $global:DashboardConfig.DefaultConfig[$section].Keys)
                        {
                            # Check if key exists in the loaded config's section
                            if (-not $global:DashboardConfig.Config[$section].Contains($key))
                            {
                                #region Step: Add Missing Key
                                Write-Verbose "  INI: (Initialize-IniConfig): Adding missing key '$key' in section '[$section]' from defaults." -ForegroundColor DarkGray
                                $global:DashboardConfig.Config[$section][$key] = $global:DashboardConfig.DefaultConfig[$section][$key]
                                $needsUpdate = $true
                                #endregion Step: Add Missing Key
                            }
                        }
                        #endregion Step: Iterate Default Keys in Section
                    }
                }
                #endregion Step: Iterate Default Sections
            }
            catch
            {
                #region Step: Error Handling - Failed Structure Verification
                Write-Verbose "  INI: (Initialize-IniConfig): Error verifying config structure against defaults. Error: $($_.Exception.Message)" -ForegroundColor Red
                return $false # Indicate initialization failure
                #endregion Step: Error Handling - Failed Structure Verification
            }
            #endregion Step: Verify Config Structure Against Defaults

            #region Step: Write Updated Config if Necessary
            if ($needsUpdate)
            {
                #region Step: Write Updated Config File
                Write-Verbose '  INI: (Initialize-IniConfig): Configuration updated with missing defaults. Writing changes...' -ForegroundColor DarkGray
                $writeSuccess = Write-Config -ConfigPath $configPath -Config $global:DashboardConfig.Config
                if (-not $writeSuccess)
                {
                    Write-Verbose "  INI: (Initialize-IniConfig): Failed to write updated configuration to '$configPath'." -ForegroundColor Red
                    return $false # Indicate initialization failure
                }
                Write-Verbose "  INI: (Initialize-IniConfig): Successfully wrote updated config to '$configPath'." -ForegroundColor Green
                #endregion Step: Write Updated Config File
            }
            else
            {
                #region Step: Log No Updates Needed
                Write-Verbose "  INI: (Initialize-IniConfig): Config structure verified. No updates needed." -ForegroundColor Green
                #endregion Step: Log No Updates Needed
            }
            #endregion Step: Write Updated Config if Necessary
        }
        #endregion Step: Config File Exists - Read and Verify

        #region Step: Final Success
        Write-Verbose "  INI: (Initialize-IniConfig): Configuration initialization completed successfully." -ForegroundColor Green
        return $true # Indicate successful initialization
        #endregion Step: Final Success
    }
    #endregion Function: Initialize-IniConfig

    #region Function: Read-Config
    function Read-Config
    {
        <#
        .SYNOPSIS
            Reads the configuration from the specified INI file into the global config variable.
        .DESCRIPTION
            Uses the C# 'IniFile' class to read the INI file specified by the ConfigPath parameter
            (or defaults to $global:DashboardConfig.Paths.Ini).
            Populates the $global:DashboardConfig.Config ordered dictionary with the content read from the file.
            If reading fails (e.g., file not found, parsing error), it logs an error and attempts to populate
            the global config with default values ($global:DashboardConfig.DefaultConfig) as a fallback,
            using the LoadDefaultConfigOnError helper function.
        .PARAMETER ConfigPath
            [string] The full path to the INI configuration file to read. Defaults to $global:DashboardConfig.Paths.Ini if not provided.
        .OUTPUTS
            [bool] Returns $true if the config file was read successfully and $global:DashboardConfig.Config was populated from the file.
            Returns $false if reading failed and defaults were loaded instead (or if loading defaults also failed).
        .NOTES
            - Modifies $global:DashboardConfig.Config.
            - Requires the 'IniFile' class (from classes.psm1).
            - Relies on $global:DashboardConfig structure (.Paths.Ini, .DefaultConfig).
            - Uses the internal LoadDefaultConfigOnError function for fallback.
            - Logs outcomes and errors to the host stream.
        #>
        [CmdletBinding()]
        [OutputType([bool])]
        param(
            [Parameter()]
            [string]$ConfigPath = $global:DashboardConfig.Paths.Ini
        )

        #region Step: Validate Config Path
        if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
            Write-Verbose "  INI: (Read-Config): Configuration path is not defined." -ForegroundColor Red
            # Attempt to load defaults as per original catch logic
            Return (LoadDefaultConfigOnError -Reason "Configuration path not defined")
        }
        Write-Verbose "  INI: (Read-Config): Reading config from '$ConfigPath'" -ForegroundColor Cyan
        #endregion Step: Validate Config Path

        #region Step: Read INI File
        try
        {
            #region Step: Check File Existence/Type
            if (-not (Test-Path -Path $ConfigPath -PathType Leaf)) {
                Write-Verbose "  INI: (Read-Config): Config file not found or is a directory at '$ConfigPath'." -ForegroundColor Yellow
                # Populate with defaults as per original logic
                Return (LoadDefaultConfigOnError -Reason "Config file not found at '$ConfigPath'")
            }
            #endregion Step: Check File Existence/Type

            #region Step: Use IniFile Class to Read
            # Assuming IniFile class is loaded
            # $iniHandler: Instance of the C# class used to interact with the INI file.
            $iniHandler = [Custom.IniFile]::new($ConfigPath)
            # $readConfig: The raw content read from the INI file as a .NET OrderedDictionary.
            $readConfig = $iniHandler.ReadIniFile()
            #endregion Step: Use IniFile Class to Read

            #region Step: Convert to PowerShell OrderedDictionary & Store Globally
            $global:DashboardConfig.Config = [ordered]@{}
            if ($null -ne $readConfig) {
                foreach ($section in $readConfig.Keys)
                {
                    $global:DashboardConfig.Config[$section] = [ordered]@{}
                    # Ensure the value associated with the section key is dictionary-like before iterating keys
                    if ($readConfig[$section] -is [System.Collections.IDictionary]) {
                        foreach ($key in $readConfig[$section].Keys)
                        {
                            $global:DashboardConfig.Config[$section][$key] = $readConfig[$section][$key]
                        }
                    }
                    else {
                        Write-Verbose "  INI: (Read-Config): Section '[$section]' in INI file '$ConfigPath' does not contain key-value pairs." -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Verbose "  INI: (Read-Config): Reading '$ConfigPath' returned null." -ForegroundColor Red
                # If null, treat as error and load defaults
                Return (LoadDefaultConfigOnError -Reason "Reading config file '$ConfigPath' returned null")
            }
            #endregion Step: Convert to PowerShell OrderedDictionary & Store Globally

            #region Step: Success
            Write-Verbose "  INI: (Read-Config): Config loaded successfully from '$ConfigPath'." -ForegroundColor Green
            return $true
            #endregion Step: Success
        }
        catch
        {
            #region Step: Error Handling - Failed to Read/Process
            Write-Verbose "  INI: (Read-Config): Failed to read/process config file '$ConfigPath'. Error: $($_.Exception.Message)" -ForegroundColor Red
            # Attempt to load defaults as a fallback
            Return (LoadDefaultConfigOnError -Reason "Error reading/processing '$ConfigPath': $($_.Exception.Message)")
            #endregion Step: Error Handling - Failed to Read/Process
        }
        #endregion Step: Read INI File
    }
    #endregion Function: Read-Config

    #region Function: Write-Config
    function Write-Config
    {
        <#
        .SYNOPSIS
            Writes a configuration dictionary to an INI file.
        .DESCRIPTION
            Takes a PowerShell ordered dictionary (typically $global:DashboardConfig.Config) and writes it
            to the specified INI file path using the C# 'IniFile' class.
            It handles converting array values within the dictionary to comma-separated strings,
            as INI files do not natively support arrays.
            It also ensures the target directory for the INI file exists, creating it if necessary.
        .PARAMETER Config
            [System.Collections.IDictionary] The configuration data ([ordered] dictionary) to write. Defaults to $global:DashboardConfig.Config if not provided.
        .PARAMETER ConfigPath
            [string] The full path to the INI file to write. Defaults to $global:DashboardConfig.Paths.Ini if not provided.
        .OUTPUTS
            [bool] Returns $true if the configuration was written successfully to the file.
            Returns $false if any error occurred during directory creation, data preparation, or file writing.
        .NOTES
            - Requires the 'IniFile' class (from classes.psm1).
            - Creates the destination directory if it doesn't exist.
            - Converts array values to comma-separated strings. Other complex types might not be handled correctly.
            - Supports -WhatIf and -Confirm through [CmdletBinding(SupportsShouldProcess=$true)].
            - Logs progress and errors to the host stream.
        #>
        [CmdletBinding(SupportsShouldProcess=$true)] # Added ShouldProcess support
        [OutputType([bool])]
        param(
            [Parameter()]
            [System.Collections.IDictionary]$Config = $global:DashboardConfig.Config, # Accept PS [ordered] or .NET

            [Parameter()]
            [string]$ConfigPath = $global:DashboardConfig.Paths.Ini
        )

        #region Step: Validate Inputs
        if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
            Write-Verbose "  INI: (Write-Config): Configuration path is not defined." -ForegroundColor Red
            return $false
        }
        if ($null -eq $Config) {
            Write-Verbose "  INI: (Write-Config): Configuration data to write is null." -ForegroundColor Red
            return $false
        }
        Write-Verbose "  INI: (Write-Config): Preparing to write config to '$ConfigPath'" -ForegroundColor Cyan

        if (-not $pscmdlet.ShouldProcess($ConfigPath, "Write Configuration")) {
            Write-Verbose "  INI: (Write-Config): Write operation cancelled by ShouldProcess." -ForegroundColor Yellow
            return $false
        }
        #endregion Step: Validate Inputs

        #region Step: Ensure Directory Exists
        try
        {
            $configDir = Split-Path -Path $ConfigPath -Parent
            if (-not (Test-Path -Path $configDir -PathType Container))
            {
                Write-Verbose "  INI: (Write-Config): Creating directory '$configDir'." -ForegroundColor DarkGray
                New-Item -ItemType Directory -Path $configDir -Force -ErrorAction Stop | Out-Null
            }
        }
        catch
        {
            Write-Verbose "  INI: (Write-Config): Failed to create directory '$configDir'. Error: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
        #endregion Step: Ensure Directory Exists

        #region Step: Prepare Data for Writing (Handle Arrays, Convert to .NET Dictionary)
        # Need to convert PS [ordered] to .NET OrderedDictionary for the C# class
        # $configToWrite: The .NET OrderedDictionary prepared for the C# IniFile class.
        $configToWrite = New-Object System.Collections.Specialized.OrderedDictionary

        try
        {
            foreach ($sectionKey in $Config.Keys)
            {
                $sectionName = $sectionKey.ToString() # Ensure string key
                # Ensure section value is a dictionary before proceeding
                if ($Config[$sectionKey] -is [System.Collections.IDictionary]) {
                    $sectionData = New-Object System.Collections.Specialized.OrderedDictionary
                    $sourceSection = $Config[$sectionKey]

                    foreach ($itemKey in $sourceSection.Keys)
                    {
                        $itemName = $itemKey.ToString() # Ensure string key
                        $value = $sourceSection[$itemKey]

                        # Process arrays: Convert them to comma-separated strings
                        if ($value -is [Array])
                        {
                            # Filter out potential $null elements before joining
                            $stringElements = $value | Where-Object { $_ -ne $null } | ForEach-Object { $_.ToString() }
                            $preparedValue = $stringElements -join ','
                            # Use Verbose stream for this detail
                            Write-Verbose "INI: (Write-Config): Converted array for [$sectionName]$itemName to '$preparedValue'"
                        }
                        elseif ($null -ne $value)
                        {
                            # Convert other non-null values to string just in case
                            $preparedValue = $value.ToString()
                        }
                        else
                        {
                            # Handle null value (write as empty string)
                            $preparedValue = ''
                        }
                        $sectionData.Add($itemName, $preparedValue)
                    }
                    $configToWrite.Add($sectionName, $sectionData)
                } else {
                    Write-Verbose "  INI: (Write-Config): Section '$sectionName' is not a dictionary, skipping write for this section." -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Verbose "  INI: (Write-Config): Failed to prepare configuration data for writing. Error: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
        #endregion Step: Prepare Data for Writing (Handle Arrays, Convert to .NET Dictionary)

        #region Step: Write using IniFile Class
        try
        {
            # Assuming IniFile class is loaded
            # $iniFile: Instance of the C# class used to write the INI file.
            $iniFile = [Custom.IniFile]::new($ConfigPath)
            # Pass the prepared .NET OrderedDictionary to the writing method
            $iniFile.WriteIniFile($configToWrite) # Assumes WriteIniFile handles potential IOExceptions

            Write-Verbose "  INI: (Write-Config): Config written successfully to '$ConfigPath'." -ForegroundColor Green
            return $true
        }
        catch # Catch errors specifically from the WriteIniFile call or instantiation
        {
            Write-Verbose "  INI: (Write-Config): Failed to write config using IniFile class to '$ConfigPath'. Error: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
        #endregion Step: Write using IniFile Class
    }
    #endregion Function: Write-Config

#endregion Core Configuration Functions

#region Module Exports

    #region Step: Export Public Functions
    # Export functions intended for external use by other modules or scripts.
    Export-ModuleMember -Function Initialize-IniConfig, Read-Config, Write-Config, Get-IniFileContent, Copy-OrderedDictionary, LoadDefaultConfigOnError
    #endregion Step: Export Public Functions

#endregion Module Exports