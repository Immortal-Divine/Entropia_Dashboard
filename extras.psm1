<# extras.psm1 #>

$ProxyUrl = 'https://worldboss.entropia-dashboard.workers.dev/'

if (-not $global:DashboardConfig.State.WorldBossListener)
{
	$global:DashboardConfig.State.WorldBossListener = @{
		LastTimestamp = ([DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds())
		Timer         = $null
		RunspacePool  = $null
		CurrentJob    = $null
		NextRunTime   = [DateTime]::MinValue
		IsFirstRun    = $true
		HiddenNotifications = [System.Collections.Generic.HashSet[string]]::new()
        BossStates    = @{}
	}
}
function Stop-WorldBossListener
{
	if ($global:DashboardConfig.State.WorldBossListener.Timer)
	{
		$global:DashboardConfig.State.WorldBossListener.Timer.Stop()
		$global:DashboardConfig.State.WorldBossListener.Timer.Dispose()
		$global:DashboardConfig.State.WorldBossListener.Timer = $null
	}
	if ($global:DashboardConfig.State.WorldBossListener.RunspacePool)
	{
        if ($global:DashboardConfig.State.WorldBossListener.CurrentJob) {
            $job = $global:DashboardConfig.State.WorldBossListener.CurrentJob
            try {
                if ($job.PowerShell) {
                    if ($job.PowerShell.InvocationStateInfo.State -eq 'Running') {
                        $job.PowerShell.Stop()
                    }
                    $job.PowerShell.Dispose()
                }
            } catch { Write-Verbose "Error stopping WorldBoss job: $_" }
        }
		if (Get-Command DisposeManagedRunspace -ErrorAction SilentlyContinue -Verbose:$False)
		{
			$jobToDispose = $global:DashboardConfig.State.WorldBossListener.CurrentJob
			if (-not $jobToDispose)
			{
				$jobToDispose = @{ RunspacePool = $global:DashboardConfig.State.WorldBossListener.RunspacePool }
			}
			DisposeManagedRunspace -JobResource $jobToDispose
		}
        try {
            if ($global:DashboardConfig.State.WorldBossListener.RunspacePool) {
                $global:DashboardConfig.State.WorldBossListener.RunspacePool.Dispose()
            }
        } catch {
            Write-Verbose "Error disposing WorldBossListener RunspacePool: $_"
        }
		$global:DashboardConfig.State.WorldBossListener.RunspacePool = $null
		$global:DashboardConfig.State.WorldBossListener.CurrentJob = $null
		Write-Verbose 'WorldBoss Listener Stopped.'
	}
}
function ParseTimeSpanString
{
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[string]$InputString
	)

	if ([string]::IsNullOrWhiteSpace($InputString))
	{
		return [TimeSpan]::Zero
	}

	$ts = [TimeSpan]::Zero
	$workString = $InputString -replace ',', '.'
    
	$pattern = '(?<colon>\d+(?:\:\d+)+(?:\.\d+)?)|(?<uval>\d+(?:\.\d+)?)\s*(?<unit>[a-zA-Z]+)|(?<lone>\d+(?:\.\d+)?)'
    
	$match = [regex]::Matches($workString, $pattern)
    
	foreach ($match in $match)
	{
		if ($match.Groups['colon'].Success)
		{
			$parts = $match.Groups['colon'].Value -split ':'
			$pCount = $parts.Count
			$s = 0; $m = 0; $h = 0; $d = 0
            
			if ($pCount -eq 2)
			{
				$h = [double]$parts[0]; $m = [double]$parts[1]
			}
			elseif ($pCount -eq 3)
			{
				$h = [double]$parts[0]; $m = [double]$parts[1]; $s = [double]$parts[2]
			}
			elseif ($pCount -ge 4)
			{
				$d = [double]$parts[0]; $h = [double]$parts[1]; $m = [double]$parts[2]; $s = [double]$parts[3]
			}
			$ts = $ts.Add([TimeSpan]::FromDays($d)).Add([TimeSpan]::FromHours($h)).Add([TimeSpan]::FromMinutes($m)).Add([TimeSpan]::FromSeconds($s))
		}
		elseif ($match.Groups['unit'].Success)
		{
			$val = [double]$match.Groups['uval'].Value
			$uStr = $match.Groups['unit'].Value
            
			if ($uStr -ceq 'M')
			{
				$ts = $ts.Add([TimeSpan]::FromDays($val * 30))
			}
			else
			{
				switch -Regex ($uStr.ToLower())
				{
					'^ms$|^milli(s|seconds?)?$' { $ts = $ts.Add([TimeSpan]::FromMilliseconds($val)); break }
					'^s$|^sec(s|onds?)?$' { $ts = $ts.Add([TimeSpan]::FromSeconds($val)); break }
					'^m$|^min(s|utes?)?$|^mn$|^mi$' { $ts = $ts.Add([TimeSpan]::FromMinutes($val)); break }
					'^h$|^hr(s)?$|^hours?$' { $ts = $ts.Add([TimeSpan]::FromHours($val)); break }
					'^d$|^days?$' { $ts = $ts.Add([TimeSpan]::FromDays($val)); break }
					'^w$|^wk(s)?$|^weeks?$' { $ts = $ts.Add([TimeSpan]::FromDays($val * 7)); break }
					'^mos?$|^mon(s|ths?)?$' { $ts = $ts.Add([TimeSpan]::FromDays($val * 30)); break }
					'^y$|^yr(s)?$|^years?$' { $ts = $ts.Add([TimeSpan]::FromDays($val * 365)); break }
				}
			}
		}
		elseif ($match.Groups['lone'].Success)
		{
			$val = [double]$match.Groups['lone'].Value
			$ts = $ts.Add([TimeSpan]::FromMinutes($val))
		}
	}

	return $ts
}

function Invoke-TestButtonAction
{
	param($Status, $Percent, $Type, $BossName, $ImageUrl, $Key)
    
	[System.Windows.Forms.MessageBox]::Show("Test: $Status Clicked")
    
	if ($Status -eq 'Died' -or $Status -eq 'False Alarm') {
		CloseToast -Key $Key
		ShowToast -Title "TEST: $BossName $Status" -Message "Local test update: $Status" -Type $Type -TimeoutSeconds 10 -IgnoreCancellation
	} else {
		if ($Status -eq 'Hide') {
			CloseToast -Key $Key
			return
		}

		$testButtons = @(
			@{ Text = "25%"; WidthPercent = 25; Action = { Invoke-TestButtonAction -Status "25%" -Percent 25 -Type "Warning" -BossName $BossName -ImageUrl $ImageUrl -Key $Key }.GetNewClosure() }
			@{ Text = "Died"; WidthPercent = 25; Action = { Invoke-TestButtonAction -Status "Died" -Percent 0 -Type "Info" -BossName $BossName -ImageUrl $ImageUrl -Key $Key }.GetNewClosure() },
			@{ Text = "False Alarm"; WidthPercent = 24; Action = { Invoke-TestButtonAction -Status "False Alarm" -Percent 0 -Type "Error" -BossName $BossName -ImageUrl $ImageUrl -Key $Key }.GetNewClosure() },
			@{ Text = "Hide"; WidthPercent = 24; Action = { Invoke-TestButtonAction -Status "Hide" -Percent 0 -Type "Info" -BossName $BossName -ImageUrl $ImageUrl -Key $Key }.GetNewClosure() }
		)

		ShowInteractiveNotification -Title "World Boss Detected (TEST) ($Status)" -Message "$BossName has just spawned! [$([DateTime]::Now.ToString('HH:mm'))]`nLocal Update: $Status" -Type "Warning" -Key $Key -TimeoutSeconds 0 -Progress $Percent -IgnoreCancellation -Buttons $testButtons -ImageUrl $ImageUrl
	}
}

function Send-BossUpdate
{
    param(
        $BossName,
        $Status,
        $Percent,
        [string]$DiscordMessageId
    )
    
    $user = $global:DashboardConfig.Config['User']['Username']
    if ([string]::IsNullOrWhiteSpace($user)) { $user = 'Dashboard User' }
    
    $targetUrl = $script:ProxyUrl 
    if ([string]::IsNullOrWhiteSpace($targetUrl)) {
        [System.Windows.Forms.MessageBox]::Show("Error: Proxy URL is missing. Cannot send update.")
        return
    }

    $code = $global:DashboardConfig.Config['Options']['AccessCode']

    $bossImg = ""
    $lookupName = $BossName -replace '\[RANDOM\] ', ''
    if ($global:DashboardConfig.Resources.BossData.ContainsKey($lookupName)) {
        $bossImg = $global:DashboardConfig.Resources.BossData[$lookupName].url
    }

    $payload = @{
        auth_code = $code
        type      = 'update_status'
        boss_name = $BossName
        status    = $Status
        percent   = $Percent
        username  = $user
        boss_img  = $bossImg
		discord_message_id = $DiscordMessageId
    } | ConvertTo-Json
    
    Add-Type -AssemblyName System.Net.Http
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

    try {
        $client = [System.Net.Http.HttpClient]::new()
        $client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        $content = [System.Net.Http.StringContent]::new($payload, [System.Text.Encoding]::UTF8, "application/json")
        
        $task = $client.PostAsync($targetUrl, $content)
        
        $continuation = [Action[System.Threading.Tasks.Task]]{
            param($t)
            try {
                if ($t.IsFaulted) {
                    Write-Verbose "Send-BossUpdate Failed: $($t.Exception.InnerException.Message)"
                }
            } finally {
                $client.Dispose()
            }
        }
        
        $scheduler = if ([System.Threading.SynchronizationContext]::Current) { [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext() } else { [System.Threading.Tasks.TaskScheduler]::Default }
        $task.ContinueWith($continuation, $scheduler) | Out-Null
    }
    catch {
        Write-Verbose "Send-BossUpdate Error: $_"
    }
}

function Send-Message
{
	param(
		$code,
		$bossName
	)

	$boss = $global:DashboardConfig.Resources.BossData[$bossName]

	if (-not $boss)
	{
		[Windows.Forms.MessageBox]::Show("Unknown boss: $bossName", 'Error')
		return
	}

	$bossImageUrl = $boss.url
	$bossRolePing = $boss.role
	$bossDisplay = $boss.name

	$userToSend = $global:DashboardConfig.Config['User']['Username']
	if ([string]::IsNullOrWhiteSpace($userToSend))
	{
		$userToSend = $null
	}

	if ($code -eq 'TEST')
	{
		$fileName = $bossImageUrl.Split('/')[-1]
		$localPath = Join-Path $global:DashboardConfig.Paths.Bosses $fileName
        
		$showImages = $true
		if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options'].Contains('ShowBossImages'))
		{
			$showImages = ([int]$global:DashboardConfig.Config['Options']['ShowBossImages']) -eq 1
		}
		$imgToUse = $null
		if ($showImages)
		{ 
			if (-not (Test-Path $localPath) -and $bossImageUrl)
			{
				$tempPath = "$localPath.$([Guid]::NewGuid()).tmp"
				try
				{
					[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
					if (-not $global:DashboardConfig.Resources.ActiveDownloads) { $global:DashboardConfig.Resources.ActiveDownloads = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new()) }
					$wc = New-Object System.Net.WebClient
					$global:DashboardConfig.Resources.ActiveDownloads.Add($wc) | Out-Null
					$wc.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)')
					try {
						$wc.DownloadFile($bossImageUrl, $tempPath)
					} finally {
						$wc.Dispose()
						$global:DashboardConfig.Resources.ActiveDownloads.Remove($wc)
					}
					if (Test-Path $localPath) { Remove-Item -LiteralPath $localPath -Force }
					Move-Item -LiteralPath $tempPath -Destination $localPath -Force
				}
				catch { Write-Verbose "Error downloading image in Send-Message: $_" }
				finally {
					if (Test-Path $tempPath) { try { Remove-Item -LiteralPath $tempPath -Force } catch {} }
				}
			}
			$imgToUse = if (Test-Path $localPath) { $localPath } else { $null } 
		}

		$testKey = $bossDisplay.GetHashCode()

		$testButtons = @(
			@{ Text = "25%"; WidthPercent = 25; Action = { Invoke-TestButtonAction -Status "25%" -Percent 25 -Type "Warning" -BossName $bossDisplay -ImageUrl $imgToUse -Key $testKey }.GetNewClosure() },
			@{ Text = "Died"; WidthPercent = 25; Action = { Invoke-TestButtonAction -Status "Died" -Percent 0 -Type "Info" -BossName $bossDisplay -ImageUrl $imgToUse -Key $testKey }.GetNewClosure() },
			@{ Text = "False Alarm"; WidthPercent = 24; Action = { Invoke-TestButtonAction -Status "False Alarm" -Percent 0 -Type "Error" -BossName $bossDisplay -ImageUrl $imgToUse -Key $testKey }.GetNewClosure() },
			@{ Text = "Hide"; WidthPercent = 24; Action = { Invoke-TestButtonAction -Status "Hide" -Percent 0 -Type "Info" -BossName $bossDisplay -ImageUrl $imgToUse -Key $testKey }.GetNewClosure() }
		)

		ShowInteractiveNotification `
			-Title "World Boss Detected (TEST)" `
			-Message "$bossDisplay has just spawned! [$([DateTime]::Now.ToString('HH:mm'))]" `
			-Type "Warning" `
			-Key $testKey `
			-TimeoutSeconds 0 `
			-Progress 100 `
			-IgnoreCancellation `
			-Buttons $testButtons `
			-ImageUrl $imgToUse `
		return
	}

	$footerText = "Sent via Entropia Dashboard by $($userToSend)`nhttps://immortal-divine.github.io/Entropia_Dashboard/"

	$embed = @{
		title       = "$bossDisplay spawned!"
		color       = 16724556
		description = ' '
		footer      = @{ text = $footerText } 
	}

	if (-not [string]::IsNullOrEmpty($bossImageUrl))
	{
		$embed.image = @{ url = $bossImageUrl }
		$embed.thumbnail = @{ url = $bossImageUrl }
	}

	$discordPayload = @{
		content = "<@&$bossRolePing> **$bossDisplay** spawned!"
		embeds  = @($embed)
	}

	$finalBody = @{
		auth_code      = $code
		message        = $discordPayload
		meta_boss_name = $bossDisplay
		meta_boss_img  = $bossImageUrl
		username       = $userToSend
	} | ConvertTo-Json -Depth 10

	try
	{
		[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
		Invoke-RestMethod -Uri $ProxyUrl -Method Post -Body $finalBody -ContentType 'application/json' -ErrorAction Stop
        
		ShowToast -Title 'Success' -Message "$bossDisplay reported!" -Type 'Info' -IgnoreCancellation
	}
	catch
	{
		$errorMsg = $_.Exception.Message
        
		if ($_.Exception.Response)
		{
			$reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
			$serverText = $reader.ReadToEnd()
			if (-not [string]::IsNullOrWhiteSpace($serverText))
			{
				$errorMsg = $serverText
			}
		}
	}
}

function ProcessSpawnEvent {
    param($Spawn, $LastTime, $State)

    $serverTime = [int64]$Spawn.timestamp
    $updateTime = if ($Spawn.lastUpdateTime) { [int64]$Spawn.lastUpdateTime } else { 0 }

    $bossName = $Spawn.name
    $trackingKey = if ($Spawn.id) { $Spawn.id } else { $bossName }
    $hpStatus = if ($Spawn.hpStatus) { $Spawn.hpStatus } else { '100%' }
    $hpPercent = if ($Spawn.hpPercent) { [int]$Spawn.hpPercent } else { 100 }

    if ($null -eq $State.BossStates) { $State.BossStates = @{} }

    $shouldNotify = $false


    if ($State.BossStates.ContainsKey($trackingKey)) {
        $cached = $State.BossStates[$trackingKey]

        if ("$($cached.hpStatus)" -ne "$hpStatus" -or [int]$cached.hpPercent -ne $hpPercent) {
            $shouldNotify = $true
        }
    }
    else {

        if ($hpStatus -eq 'Died' -or $hpStatus -eq 'False Alarm') {
            $shouldNotify = $false
        }
        else {
            $shouldNotify = $true
        }
    }

    $State.BossStates[$trackingKey] = @{
        hpPercent = $hpPercent;
        hpStatus = $hpStatus;
        lastUpdateTime = $updateTime
    }


    if ($shouldNotify) {
        $notifKey = $trackingKey.GetHashCode()

        if ($State.HiddenNotifications.Contains($trackingKey)) {
            if ($hpStatus -eq 'Died' -or $hpStatus -eq 'False Alarm') {
                $State.HiddenNotifications.Remove($trackingKey) | Out-Null
			}
            return
        }

        $bossImg = $null
        $showImages = $true
        if ($global:DashboardConfig.Config['Options'] -and $global:DashboardConfig.Config['Options'].Contains('ShowBossImages')) {
            $showImages = ([int]$global:DashboardConfig.Config['Options']['ShowBossImages']) -eq 1
        }
        if ($showImages) {
            $lookupName = $bossName -replace '\[RANDOM\] ', ''
            $bData = $global:DashboardConfig.Resources.BossData[$lookupName]
            if ($bData -and $bData.url) {
                $fileName = $bData.url.Split('/')[-1]
                $localPath = Join-Path $global:DashboardConfig.Paths.Bosses $fileName

                if (-not (Test-Path $localPath)) {
                    try {
                        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                        $wc = New-Object System.Net.WebClient
                        $wc.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)')
                        $tempPath = "$localPath.$([Guid]::NewGuid()).tmp"
                        $wc.DownloadFile($bData.url, $tempPath)
                        $wc.Dispose()
                        Move-Item -LiteralPath $tempPath -Destination $localPath -Force
                    } catch {}
                }

                if (Test-Path $localPath) { $bossImg = $localPath }
            }
        }

        $triggeredByText = ''
        if ($Spawn.triggeredBy) { $triggeredByText = "`nTriggered by: $($Spawn.triggeredBy)" }
        if ($Spawn.lastUpdateBy) { $triggeredByText = "`nLast update: $($Spawn.lastUpdateBy)" }

        $popTime = [DateTimeOffset]::FromUnixTimeMilliseconds($serverTime).LocalDateTime.ToString('HH:mm')

        $activeButtons = @(
            @{ Text = '25%'; WidthPercent = 25; Action = { Send-BossUpdate -BossName $bossName -Status '25%' -Percent 25 -DiscordMessageId $Spawn.discordMessageId; ShowToast -Title "$bossName (25%)" -Message "Spawned at $popTime$triggeredByText" -Type 'Warning' -Key $notifKey -TimeoutSeconds 0 -Progress 25 -IgnoreCancellation }.GetNewClosure() },
            @{ Text = 'Died'; WidthPercent = 25; Action = { Send-BossUpdate -BossName $bossName -Status 'Died' -Percent 0 -DiscordMessageId $Spawn.discordMessageId; CloseToast -Key $notifKey }.GetNewClosure() },
            @{ Text = 'False Alarm'; WidthPercent = 24; Action = { Send-BossUpdate -BossName $bossName -Status 'False Alarm' -Percent 0 -DiscordMessageId $Spawn.discordMessageId; CloseToast -Key $notifKey }.GetNewClosure() },
            @{ Text = 'Hide'; WidthPercent = 24; Action = { CloseToast -Key $notifKey; if (-not $global:DashboardConfig.State.WorldBossListener.HiddenNotifications.Contains($trackingKey)) { $global:DashboardConfig.State.WorldBossListener.HiddenNotifications.Add($trackingKey) | Out-Null } }.GetNewClosure() }
		)

        $finalStateButtons = @(
            @{ Text = 'Hide'; WidthPercent = 100; Action = { CloseToast -Key $notifKey; if (-not $global:DashboardConfig.State.WorldBossListener.HiddenNotifications.Contains($trackingKey)) { $global:DashboardConfig.State.WorldBossListener.HiddenNotifications.Add($trackingKey) | Out-Null } }.GetNewClosure() }
        )

        if ($hpStatus -eq 'False Alarm') {
            ShowInteractiveNotification `
                -Title "FALSE ALARM: $bossName" `
                -Message "Marked as False Alarm by $($Spawn.lastUpdateBy)." `
                -Type 'Info' `
                -Key $notifKey `
                -TimeoutSeconds 10 `
                -Progress 0 `
				-Buttons $finalStateButtons `
				-ImageUrl $bossImg `
                -IgnoreCancellation
            $State.HiddenNotifications.Remove($trackingKey) | Out-Null 
        }
        elseif ($hpStatus -eq 'Died') {
            ShowInteractiveNotification `
                -Title "$bossName Defeated!" `
                -Message "Boss marked as Died by $($Spawn.lastUpdateBy)." `
                -Type 'Info' `
                -Key $notifKey `
                -TimeoutSeconds 10 `
                -Progress 0 `
				-Buttons $finalStateButtons `
				-ImageUrl $bossImg `
                -IgnoreCancellation
            $State.HiddenNotifications.Remove($trackingKey) | Out-Null 
        }
        else {
            $notifyFilter = $true
            if ($global:DashboardConfig.Config['BossFilter'] -and $global:DashboardConfig.Config['BossFilter'].Contains($bossName)) {
                try { $notifyFilter = [bool]::Parse($global:DashboardConfig.Config['BossFilter'][$bossName]) } catch {}
            }

            if ($notifyFilter) {
                if ($hpStatus -eq 'Will Spawn' -or $hpStatus -eq 'Not Spawned') { return }

                ShowInteractiveNotification `
                    -Title "$bossName ($hpStatus)" `
                    -Message "Spawned at $popTime$triggeredByText" `
                    -Type 'Warning' `
                    -Key $notifKey `
                    -TimeoutSeconds 600 `
                    -Progress $hpPercent `
                    -IgnoreCancellation `
                    -Buttons $activeButtons `
                    -ImageUrl $bossImg `
            }
        }
    }
}

function Start-WorldBossListener
{
    $enabled = $false
    if ($global:DashboardConfig.UI.WorldBossListener) {
        $enabled = $global:DashboardConfig.UI.WorldBossListener.Checked
    } elseif ($global:DashboardConfig.Config['Options'] -and [int]$global:DashboardConfig.Config['Options']['WorldBossListener'] -eq 1) {
        $enabled = $true
    }

    if (-not $enabled) { Stop-WorldBossListener; return }
    if ($global:DashboardConfig.State.WorldBossListener.Timer) { return }

    if (-not $global:DashboardConfig.State.WorldBossListener.BossStates) {
        $global:DashboardConfig.State.WorldBossListener.BossStates = @{}
    }
    if (-not $global:DashboardConfig.State.WorldBossListener.HiddenNotifications) {
        $global:DashboardConfig.State.WorldBossListener.HiddenNotifications = [System.Collections.ArrayList]::new()
    }

    if (-not $global:DashboardConfig.State.WorldBossListener.RunspacePool) {
        if (Get-Command NewManagedRunspace -ErrorAction SilentlyContinue) {
            $global:DashboardConfig.State.WorldBossListener.RunspacePool = NewManagedRunspace -Name 'WorldBossListener' -MinRunspaces 1 -MaxRunspaces 1
        } else { return }
    }

    if ($global:DashboardConfig.Resources.BossData) {
        $bossData = $global:DashboardConfig.Resources.BossData
        $appPath = $global:DashboardConfig.Paths.Bosses
        [System.Threading.Tasks.Task]::Run([Action] {
            try {
                if (-not $global:DashboardConfig.Resources.ActiveDownloads) { $global:DashboardConfig.Resources.ActiveDownloads = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new()) }
                foreach ($key in $bossData.Keys) {
                    if ($global:DashboardConfig.State.IsStopping) { break }
                    $localPath = $null; $tempPath = $null
                    try {
                        $boss = $bossData[$key]
                        if ($boss.url) {
                            $fileName = $boss.url.Split('/')[-1]
                            $localPath = [System.IO.Path]::Combine($appPath, $fileName)
                            if (-not [System.IO.File]::Exists($localPath)) {
                                $tempPath = "$localPath.$([Guid]::NewGuid()).tmp"
                                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                                $wc = New-Object System.Net.WebClient;
                                $global:DashboardConfig.Resources.ActiveDownloads.Add($wc) | Out-Null
                                try { $wc.DownloadFile($boss.url, $tempPath); } finally { $wc.Dispose(); $global:DashboardConfig.Resources.ActiveDownloads.Remove($wc) }
                                if (Test-Path $localPath) { Remove-Item -LiteralPath $localPath -Force }
                                Move-Item -LiteralPath $tempPath -Destination $localPath -Force
                            }
                        }
                    } catch { } finally { if ($tempPath -and (Test-Path $tempPath)) { try { Remove-Item -LiteralPath $tempPath -Force } catch {} } }
                }
            } catch { }
        }) | Out-Null
    }

    $PollingIntervalSeconds = 30
    $TimerTickInterval = 1000

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $TimerTickInterval 
    
    $global:DashboardConfig.State.WorldBossListener.LastTimestamp = 0
    $global:DashboardConfig.State.WorldBossListener.IsFirstRun = $true
    $global:DashboardConfig.State.WorldBossListener.NextRunTime = [DateTime]::MinValue

    $timer.Add_Tick({
        $state = $global:DashboardConfig.State.WorldBossListener

        if ($global:DashboardConfig.UI.WorldBossListener -and -not $global:DashboardConfig.UI.WorldBossListener.Checked) {
            Stop-WorldBossListener
            return
        }

        if ($state.CurrentJob) {
            if ($state.CurrentJob.AsyncResult.IsCompleted) {
                try {
                    $ps = $state.CurrentJob.PowerShell
                    $response = $ps.EndInvoke($state.CurrentJob.AsyncResult)
                    $ps.Dispose()
                    $state.CurrentJob = $null

                   if ($response) { 
                        if ($state.IsFirstRun) {
                            $lines = @() 

                            $BuildBlock = {
                                param($ServerName, $Data)
                                $output = @()
                                
                                $IsBossEnabled = {
                                    param($bName)
                                    if ($global:DashboardConfig.Config['BossFilter'] -and $global:DashboardConfig.Config['BossFilter'].Contains($bName)) {
                                        try { return [bool]::Parse($global:DashboardConfig.Config['BossFilter'][$bName]) } catch {}
                                    }
                                    return $true
                                }

                                $activeSpawns = @($Data.activeSpawns) | Where-Object { & $IsBossEnabled $_.name }
                                $upcoming = @($Data.upcoming) | Where-Object { $_.time -notmatch '\(tomorrow\)' -and (& $IsBossEnabled $_.name) }
                                $randomRotation = @($Data.randomRotation) | Where-Object { & $IsBossEnabled $_.name }

                                $hasContent = ($upcoming.Count -gt 0) -or ($randomRotation.Count -gt 0) -or ($activeSpawns.Count -gt 0)
                                if ($hasContent) {
                                    
                                    $alive = @(); $dead = @()
                                    foreach ($b in $activeSpawns) {
                                        if ($b.hpStatus -eq 'Died') { $dead += $b } else { $alive += $b }
                                    }


                                    if ($upcoming.Count -gt 0) {
                                        $output += " [ UPCOMING ]"
											foreach ($b in $upcoming) {
												$localTime = [DateTimeOffset]::FromUnixTimeMilliseconds(
													[int64]$b.timeUtc
												).ToLocalTime().ToString('HH:mm')

												$output += "  > $localTime | $($b.name)"
											}
                                        $output += ""
                                    }
                                    if ($randomRotation.Count -gt 0) {
                                        $output += " [ RANDOM ROTATION ]"
                                        foreach ($b in $randomRotation) { $output += "  * $($b.name)" }
                                        $output += "" 
                                    }
                                }
                                return $output
                            }

                            $lines += $BuildBlock.Invoke("KHELDOR", $response.kheldor)
                            if ($lines.Count -gt 0) { $lines += "`n" }

                            if ($lines.Count -gt 0) {
                                $msg = $lines -join "`n"
                                $scheduleKey = "DailyBossSchedule".GetHashCode()
                                $buttons = @( @{ Text = 'Hide'; WidthPercent = 100; Action = { CloseToast -Key $scheduleKey }.GetNewClosure() } )

                                ShowInteractiveNotification `
                                    -Title 'Daily Boss Schedule' `
                                    -Message "$msg" `
                                    -Type 'Info' `
                                    -TimeoutSeconds 10 `
                                    -IgnoreCancellation `
                                    -Key $scheduleKey `
                                    -Buttons $buttons
                            }
                            
                            $state.IsFirstRun = $false 
                        }

                        $kheldorSpawns = @($response.kheldor.activeSpawns)


                        $allActiveAndRelevantSpawns = $kheldorSpawns
                        $knownIds = $allActiveAndRelevantSpawns | ForEach-Object { if ($_.id) { $_.id } else { $_.name } }
                        $localIds = $state.BossStates.Keys
                        
                        $recentlyFinishedIds = $localIds | Where-Object { $_ -notin $knownIds }

                        foreach ($rId in $recentlyFinishedIds) {
                            $state.BossStates[$rId].hpStatus = "Died"
                            $state.BossStates[$rId].hpPercent = 0
                            
                        }

                        foreach ($spawn in $allActiveAndRelevantSpawns) {
                            ProcessSpawnEvent -Spawn $spawn -LastTime ([int64]$state.LastTimestamp) -State $state
                        }
                        $state.LastTimestamp = $response.lastOverallUpdateTime

                    }
                }
                catch {
                    if ("$_" -match "Unauthorized") { ShowToast -Title "World Boss Listener" -Message "Unauthorized! Please update your Access Code." -Type "Error" -IgnoreCancellation }
                    if ($state.CurrentJob -and $state.CurrentJob.PowerShell) { $state.CurrentJob.PowerShell.Dispose() }
                    $state.CurrentJob = $null
                }
                
                $state.NextRunTime = [DateTime]::Now.AddSeconds($PollingIntervalSeconds)
            }
        }
        elseif ([DateTime]::Now -ge $state.NextRunTime) {
            $accessCode = $global:DashboardConfig.Config['Options']['AccessCode']
            if ([string]::IsNullOrWhiteSpace($accessCode)) { Stop-WorldBossListener; return }
            
            $isMinimal = -not $state.IsFirstRun

            $scriptBlock = {
                param($Url, $Code, $IsMinimal)
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                
                if (-not $global:RunspaceETag) { $global:RunspaceETag = $null }

                $minParam = if ($IsMinimal) { "&minimal=true" } else { "" }
                $runspacePollUrl = "$($Url)?code=$Code$minParam"
                
                try {
                    $wc = New-Object System.Net.WebClient
                    $wc.Headers.Add('User-Agent', 'EntropiaDashboardClient/1.0 (PowerShell)')
                    
                    if ($global:RunspaceETag -and $IsMinimal) {
                        $wc.Headers.Add('If-None-Match', $global:RunspaceETag)
                    }

                    try {
                        $runspaceRawResponse = $wc.DownloadString($runspacePollUrl)
                        
                        $newETag = $wc.ResponseHeaders['ETag']
                        if ($newETag) { $global:RunspaceETag = $newETag }

                        if ([string]::IsNullOrWhiteSpace($runspaceRawResponse)) { return $null }
                        [System.Object]$runspaceParsedObject = $runspaceRawResponse | ConvertFrom-Json -ErrorAction Stop
                        return $runspaceParsedObject
                    }
                    catch [System.Net.WebException] {
                        $resp = $_.Exception.Response
                        if ($resp -and [int]$resp.StatusCode -eq 304) {
                            return $null
                        }
                        throw $_
                    }
                } 
                catch { 
                    return $null 
                }
                finally {
                    if ($wc) { $wc.Dispose() }
                }
            }

            if ($state.RunspacePool -and (Get-Command InvokeInManagedRunspace -ErrorAction SilentlyContinue)) {
                $job = InvokeInManagedRunspace -RunspacePool $state.RunspacePool -ScriptBlock $scriptBlock -ArgumentList $script:ProxyUrl, $accessCode, $isMinimal -AsJob
                $state.CurrentJob = $job
            }
        }
    })

    $timer.Start()
    $global:DashboardConfig.State.WorldBossListener.Timer = $timer
    Write-Verbose 'WorldBoss Listener Started.'
}


function RefreshNoteGrid
{
	if ($global:DashboardConfig.UI.NoteGrid)
	{
		$grid = $global:DashboardConfig.UI.NoteGrid
		$grid.SuspendLayout()
		try
		{
			$grid.Rows.Clear()
			if ($global:DashboardConfig.Config['Notes'])
			{
				$keys = @($global:DashboardConfig.Config['Notes'].Keys)
				$now = [DateTime]::Now
				$notesModified = $false
				foreach ($key in $keys)
				{
					try
					{
						if (-not $global:DashboardConfig.Config['Notes'].Contains($key)) { continue }
						$note = $global:DashboardConfig.Config['Notes'][$key]
						if ($note -is [string] -and $note -eq 'System.Collections.Hashtable')
						{
							$global:DashboardConfig.Config['Notes'].Remove($key)
							$notesModified = $true
							continue
						}
						if ($note -is [string])
						{
							try
							{
								$converted = ConvertFrom-Json -InputObject $note -ErrorAction Stop
								if ($converted) { $note = $converted }
								if ($note -is [System.Collections.IDictionary]) { $note['StartupHandled'] = $false } else { $note | Add-Member -MemberType NoteProperty -Name 'StartupHandled' -Value $false -Force }
                        
								$global:DashboardConfig.Config['Notes'][$key] = $note
							}
							catch { Write-Verbose "Note JSON conversion failed for key '$key': $_" }
						}
						if ($note -is [string]) { continue }

						if ($note.DeleteWhenExpired -and $note.DueDate -and -not $note.AutoRenew -and $note.Notified)
						{
							try
							{
								$due = [DateTime]::Parse($note.DueDate)
								if ($now -ge $due)
								{
									$global:DashboardConfig.Config['Notes'].Remove($key)
									$notesModified = $true
									continue 
								}
							}
							catch {}
						}

						$content = $note.Content
						if ($note.DueDate) { $content = "Due: $($note.DueDate) - $($note.Content)" }
						$idx = $grid.Rows.Add($note.Title, $note.Type, $content)
						if ($idx -ge 0) { $grid.Rows[$idx].Tag = $key }
					}
					catch { Write-Verbose "Error processing note '$key': $_" }
				}
				if ($notesModified -and (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False))
				{
					WriteConfig
				}
			}
		}
		catch { Write-Verbose "RefreshNoteGrid Error: $_" }
		finally { $grid.ResumeLayout() }
	}
}

function Show-NoteDialog
{
	param($NoteData)
    
	$script:noteDialogResult = $null

	$form = New-Object System.Windows.Forms.Form
	$form.Text = if ($NoteData) { 'Edit Note' } else { 'Add Note' }
	$form.Size = New-Object System.Drawing.Size(400, 460)
	$form.StartPosition = 'CenterParent'
	$form.FormBorderStyle = 'None'
	$form.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
	$form.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    
	$topBar = New-Object System.Windows.Forms.Panel; $topBar.Height = 30; $topBar.Dock = 'Top'; $topBar.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $form.Controls.Add($topBar)
	$topBar.Add_MouseDown({ param($s,$e) if ($e.Button -eq 'Left') { if ('Custom.Native' -as [Type]) { [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($form.Handle, 0xA1, 0x2, 0) } } })
    
	$lblHeader = New-Object System.Windows.Forms.Label; $lblHeader.Text = $form.Text; $lblHeader.AutoSize = $true; $lblHeader.Location = New-Object System.Drawing.Point(10, 7); $lblHeader.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold); $topBar.Controls.Add($lblHeader)
	$lblHeader.Add_MouseDown({ param($s,$e) if ($e.Button -eq 'Left') { if ('Custom.Native' -as [Type]) { [Custom.Native]::ReleaseCapture(); [Custom.Native]::SendMessage($form.Handle, 0xA1, 0x2, 0) } } })
    
	$btnClose = New-Object System.Windows.Forms.Label; $btnClose.Text = 'X'; $btnClose.AutoSize = $true; $btnClose.Cursor = 'Hand'; $btnClose.Location = New-Object System.Drawing.Point(375, 7); $btnClose.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold); $btnClose.ForeColor = [System.Drawing.Color]::Gray; $topBar.Controls.Add($btnClose)
	$btnClose.Add_Click({ $form.DialogResult = 'Cancel'; $form.Close() })
	$btnClose.Add_MouseEnter({ $this.ForeColor = [System.Drawing.Color]::White }); $btnClose.Add_MouseLeave({ $this.ForeColor = [System.Drawing.Color]::Gray })
    
	$pnlContent = New-Object System.Windows.Forms.Panel; $pnlContent.Location = New-Object System.Drawing.Point(0, 30); $pnlContent.Size = New-Object System.Drawing.Size(400, 430); $form.Controls.Add($pnlContent)
    
	$AddLabel = { param($txt, $y) $l = New-Object System.Windows.Forms.Label; $l.Text = $txt; $l.Location = New-Object System.Drawing.Point(20, $y); $l.AutoSize = $true; $l.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160); $l.Font = New-Object System.Drawing.Font('Segoe UI', 9); $pnlContent.Controls.Add($l) }
    
	&$AddLabel 'Title' 15
	$txtTitle = New-Object System.Windows.Forms.TextBox; $txtTitle.Location = New-Object System.Drawing.Point(20, 35); $txtTitle.Size = New-Object System.Drawing.Size(360, 25); $txtTitle.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $txtTitle.ForeColor = [System.Drawing.Color]::White; $txtTitle.BorderStyle = 'FixedSingle'; if ($NoteData) { $txtTitle.Text = $NoteData.Title }; $pnlContent.Controls.Add($txtTitle)
    
	&$AddLabel 'Type' 70
	$cmbType = if ('Custom.DarkComboBox' -as [Type]) { New-Object Custom.DarkComboBox } else { New-Object System.Windows.Forms.ComboBox; $cmbType.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $cmbType.ForeColor = [System.Drawing.Color]::White }
	$cmbType.Location = New-Object System.Drawing.Point(20, 90); $cmbType.Size = New-Object System.Drawing.Size(150, 25); $cmbType.DropDownStyle = 'DropDownList'; $cmbType.Items.AddRange(@('Note', 'Timer', 'Reminder')); if ($NoteData) { $cmbType.SelectedItem = $NoteData.Type } else { $cmbType.SelectedIndex = 0 }; $pnlContent.Controls.Add($cmbType)
    
	$pnlDynamic = New-Object System.Windows.Forms.Panel; $pnlDynamic.Location = New-Object System.Drawing.Point(180, 70); $pnlDynamic.Size = New-Object System.Drawing.Size(215, 60); $pnlContent.Controls.Add($pnlDynamic)
    
	$lblTimerDuration = New-Object System.Windows.Forms.Label; $lblTimerDuration.Text = 'Duration:'; $lblTimerDuration.AutoSize = $true; $lblTimerDuration.Location = New-Object System.Drawing.Point(0, 0); $lblTimerDuration.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
	$txtTimerDuration = New-Object System.Windows.Forms.TextBox; $txtTimerDuration.Location = New-Object System.Drawing.Point(0, 20); $txtTimerDuration.Size = New-Object System.Drawing.Size(215, 25); $txtTimerDuration.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $txtTimerDuration.ForeColor = [System.Drawing.Color]::White; $txtTimerDuration.BorderStyle = 'FixedSingle'
	$lblTimerHint = New-Object System.Windows.Forms.Label; $lblTimerHint.Text = 'e.g. 10m, 1h 30m, 1d'; $lblTimerHint.AutoSize = $true; $lblTimerHint.Location = New-Object System.Drawing.Point(0, 46); $lblTimerHint.ForeColor = [System.Drawing.Color]::Gray; $lblTimerHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    
	$lblDate = New-Object System.Windows.Forms.Label; $lblDate.Text = 'Date'; $lblDate.AutoSize = $true; $lblDate.Location = New-Object System.Drawing.Point(0, 0); $lblDate.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
	$dtDate = New-Object System.Windows.Forms.DateTimePicker; $dtDate.Format = 'Short'; $dtDate.Location = New-Object System.Drawing.Point(0, 20); $dtDate.Size = New-Object System.Drawing.Size(100, 25); $dtDate.CalendarMonthBackground = [System.Drawing.Color]::FromArgb(45, 45, 48); $dtDate.CalendarForeColor = [System.Drawing.Color]::White
	$lblTime = New-Object System.Windows.Forms.Label; $lblTime.Text = 'Time'; $lblTime.AutoSize = $true; $lblTime.Location = New-Object System.Drawing.Point(110, 0); $lblTime.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
	$dtTime = New-Object System.Windows.Forms.DateTimePicker; $dtTime.Format = 'Time'; $dtTime.ShowUpDown = $true; $dtTime.Location = New-Object System.Drawing.Point(110, 20); $dtTime.Size = New-Object System.Drawing.Size(80, 25); $dtTime.CalendarMonthBackground = [System.Drawing.Color]::FromArgb(45, 45, 48)
    
	$lblRenewInterval = New-Object System.Windows.Forms.Label; $lblRenewInterval.Text = 'Interval (e.g. 10m, 2h, 1d):'; $lblRenewInterval.Location = New-Object System.Drawing.Point(20, 357); $lblRenewInterval.AutoSize = $true; $lblRenewInterval.Visible = $false; $pnlContent.Controls.Add($lblRenewInterval)
	$txtRenewInterval = New-Object System.Windows.Forms.TextBox; $txtRenewInterval.Location = New-Object System.Drawing.Point(200, 355); $txtRenewInterval.Size = New-Object System.Drawing.Size(180, 25); $txtRenewInterval.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $txtRenewInterval.ForeColor = [System.Drawing.Color]::White; $txtRenewInterval.BorderStyle = 'FixedSingle'; $txtRenewInterval.Visible = $false; if ($NoteData -and $NoteData.RenewInterval) { $txtRenewInterval.Text = $NoteData.RenewInterval } else { $txtRenewInterval.Text = '1d' }; $pnlContent.Controls.Add($txtRenewInterval)

	&$AddLabel 'Content' 130
	$txtContent = New-Object System.Windows.Forms.TextBox; $txtContent.Location = New-Object System.Drawing.Point(20, 150); $txtContent.Size = New-Object System.Drawing.Size(360, 100); $txtContent.Multiline = $true; $txtContent.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $txtContent.ForeColor = [System.Drawing.Color]::White; $txtContent.BorderStyle = 'FixedSingle'; if ($NoteData) { $txtContent.Text = $NoteData.Content }; $pnlContent.Controls.Add($txtContent)
    
	&$AddLabel 'Options' 260
	$chkShowOnStartup = New-Object System.Windows.Forms.CheckBox; $chkShowOnStartup.Text = 'Show on Startup'; $chkShowOnStartup.AutoSize = $true; $chkShowOnStartup.Location = New-Object System.Drawing.Point(20, 280); $chkShowOnStartup.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160); if ($NoteData -and $NoteData.ShowOnStartup) { $chkShowOnStartup.Checked = $true }; $pnlContent.Controls.Add($chkShowOnStartup)
	$chkDeleteExpired = New-Object System.Windows.Forms.CheckBox; $chkDeleteExpired.Text = 'Delete when expired (on restart)'; $chkDeleteExpired.AutoSize = $true; $chkDeleteExpired.Location = New-Object System.Drawing.Point(20, 305); $chkDeleteExpired.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160); if ($NoteData -and $NoteData.DeleteWhenExpired) { $chkDeleteExpired.Checked = $true }; $pnlContent.Controls.Add($chkDeleteExpired)
	$chkAutoRenew = New-Object System.Windows.Forms.CheckBox; $chkAutoRenew.Text = 'Renew automatically'; $chkAutoRenew.AutoSize = $true; $chkAutoRenew.Location = New-Object System.Drawing.Point(20, 330); $chkAutoRenew.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160); if ($NoteData -and $NoteData.AutoRenew) { $chkAutoRenew.Checked = $true }; $pnlContent.Controls.Add($chkAutoRenew)

	$UpdateDynamic = {
		$pnlDynamic.Controls.Clear()
		$chkDeleteExpired.Visible = $false
		$chkAutoRenew.Visible = $false
		$lblRenewInterval.Visible = $false
		$txtRenewInterval.Visible = $false
		if ($cmbType.SelectedItem -eq 'Timer')
		{
			$pnlDynamic.Controls.AddRange(@($lblTimerDuration, $txtTimerDuration, $lblTimerHint))
			if ($NoteData -and $NoteData.Type -eq 'Timer' -and $NoteData.DueDate)
			{
				try
				{ 
					$due = [DateTime]::Parse($NoteData.DueDate)
					$diff = $due - [DateTime]::Now
					if ($diff.TotalSeconds -gt 0)
					{
						$parts = @()
						if ($diff.Days -gt 0) { $parts += "$($diff.Days)d" }
						if ($diff.Hours -gt 0) { $parts += "$($diff.Hours)h" }
						if ($diff.Minutes -gt 0) { $parts += "$($diff.Minutes)m" }
						if ($diff.Seconds -gt 0) { $parts += "$($diff.Seconds)s" }
						$txtTimerDuration.Text = $parts -join ' '
					}
					else { $txtTimerDuration.Text = '10m' }
				}
				catch { $txtTimerDuration.Text = '10m' }
			}
			elseif ([string]::IsNullOrWhiteSpace($txtTimerDuration.Text)) { $txtTimerDuration.Text = '10m' }
			$chkDeleteExpired.Visible = $true
			$chkAutoRenew.Visible = $true
		}
		elseif ($cmbType.SelectedItem -eq 'Reminder')
		{
			$pnlDynamic.Controls.AddRange(@($lblDate, $dtDate, $lblTime, $dtTime ))
			if ($NoteData -and $NoteData.DueDate) { try { $d = [DateTime]::Parse($NoteData.DueDate); $dtDate.Value = $d; $dtTime.Value = $d } catch {} }
			$chkDeleteExpired.Visible = $true
			$chkAutoRenew.Visible = $true
			$lblRenewInterval.Visible = $chkAutoRenew.Checked
			$txtRenewInterval.Visible = $chkAutoRenew.Checked
		}
	}
	$cmbType.Add_SelectedIndexChanged($UpdateDynamic)
	$chkAutoRenew.Add_CheckedChanged($UpdateDynamic)

	&$UpdateDynamic

	$btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text = 'Save'; $btnSave.DialogResult = 'OK'; $btnSave.Location = New-Object System.Drawing.Point(200, 385); $btnSave.Size = New-Object System.Drawing.Size(80, 30); $btnSave.FlatStyle = 'Flat'; $btnSave.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215); $btnSave.ForeColor = [System.Drawing.Color]::White; $btnSave.FlatAppearance.BorderSize = 0; $pnlContent.Controls.Add($btnSave)
	$btnSave.Add_Click({ 
			$finalDueDate = $null
			if ($cmbType.SelectedItem -eq 'Timer')
			{
				$ts = ParseTimeSpanString -InputString $txtTimerDuration.Text
				if ($ts.TotalSeconds -eq 0) { $ts = [TimeSpan]::FromMinutes(10) }
				$finalDueDate = [DateTime]::Now.Add($ts).ToString('yyyy-MM-dd HH:mm:ss')
			}
			elseif ($cmbType.SelectedItem -eq 'Reminder')
			{
				$d = $dtDate.Value; $t = $dtTime.Value
				$finalDate = [DateTime]::new($d.Year, $d.Month, $d.Day, $t.Hour, $t.Minute, $t.Second)
				$finalDueDate = $finalDate.ToString('yyyy-MM-dd HH:mm:ss')
			}
			$script:noteDialogResult = @{
				Title             = $txtTitle.Text
				Type              = $cmbType.SelectedItem
				DueDate           = $finalDueDate
				ShowOnStartup     = $chkShowOnStartup.Checked
				DeleteWhenExpired = $chkDeleteExpired.Checked
				AutoRenew         = $chkAutoRenew.Checked
				RenewInterval     = $txtRenewInterval.Text
				Content           = $txtContent.Text
			}
			$form.DialogResult = 'OK'; $form.Close() 
		})
	$btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = 'Cancel'; $btnCancel.DialogResult = 'Cancel'; $btnCancel.Location = New-Object System.Drawing.Point(290, 385); $btnCancel.Size = New-Object System.Drawing.Size(80, 30); $btnCancel.FlatStyle = 'Flat'; $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60); $btnCancel.ForeColor = [System.Drawing.Color]::White; $btnCancel.FlatAppearance.BorderSize = 0; $pnlContent.Controls.Add($btnCancel)
	$btnCancel.Add_Click({ $form.DialogResult = 'Cancel'; $form.Close() })
    
	$form.AcceptButton = $btnSave; $form.CancelButton = $btnCancel
    
	$form.Owner = $global:DashboardConfig.UI.ExtraForm
	if ((Show-FormAsDialog -Form $form) -eq 'OK') { return $script:noteDialogResult }
	else { return $null }
}

function AddNote
{
	try
	{
		$data = Show-NoteDialog
		if ($data)
		{
			$id = [Guid]::NewGuid().ToString()
			$data.CreationDate = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
			$data.Notified = $false
			$global:DashboardConfig.Config['Notes'][$id] = $data
			if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
			RefreshNoteGrid
		}
	}
	catch { Write-Verbose "AddNote Error: $_" }
}

function EditNote
{
	$grid = $global:DashboardConfig.UI.NoteGrid
	if ($grid.SelectedRows.Count -gt 0)
	{
		$id = $grid.SelectedRows[0].Tag
		if ($global:DashboardConfig.Config['Notes'] -and $global:DashboardConfig.Config['Notes'][$id])
		{
			$originalNote = $global:DashboardConfig.Config['Notes'][$id]
			$data = Show-NoteDialog -NoteData $originalNote
			if ($data)
			{
				if ($originalNote.CreationDate -and $originalNote.DueDate -eq $data.DueDate)
				{
					$data.CreationDate = $originalNote.CreationDate
				}
				else
				{
					$data.CreationDate = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
				}
				$global:DashboardConfig.Config['Notes'][$id] = $data
				if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
				RefreshNoteGrid
			}
		}
	}
}

function RemoveNote
{
	$grid = $global:DashboardConfig.UI.NoteGrid
	if ($grid.SelectedRows.Count -gt 0)
	{
		$id = $grid.SelectedRows[0].Tag
		if ($global:DashboardConfig.Config['Notes'] -and $global:DashboardConfig.Config['Notes'][$id])
		{
			$notifKey = $id.GetHashCode()
			if ($notifKey -gt 0) { $notifKey = - $notifKey }
			CloseToast -Key $notifKey
			$global:DashboardConfig.Config['Notes'].Remove($id)
			if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
			RefreshNoteGrid
		}
	}
}

function Stop-NoteTimer
{
	if ($global:DashboardConfig.State.NoteTimer)
	{
		$global:DashboardConfig.State.NoteTimer.Stop()
		$global:DashboardConfig.State.NoteTimer.Dispose()
		$global:DashboardConfig.State.NoteTimer = $null
		Write-Verbose "Note Timer stopped."
	}
}

if (-not $global:DashboardConfig.State.NoteTimer)
{
	$noteTimer = New-Object System.Windows.Forms.Timer
	$noteTimer.Interval = 1000 
	$noteTimer.Add_Tick({
			$GetNewDueDate = {
				param($currentDueDate, $intervalString)
				$intervalTimeSpan = ParseTimeSpanString -InputString $intervalString
				if ($intervalTimeSpan.TotalSeconds -le 0)
				{
					Write-Verbose "NoteTimer: Invalid interval '$intervalString'. Defaulting to 1 day."
					$intervalTimeSpan = [TimeSpan]::FromDays(1)
				}
				return $currentDueDate.Add($intervalTimeSpan)
			}

			if ($global:DashboardConfig.Config['Notes'])
			{
				$now = [DateTime]::Now
				$keys = @($global:DashboardConfig.Config['Notes'].Keys)
				foreach ($key in $keys)
				{
					if (-not $global:DashboardConfig.Config['Notes'].Contains($key)) { continue }
					$note = $global:DashboardConfig.Config['Notes'][$key]
					if ($note -is [string]) { continue }

					if ($note.Type -eq 'Timer' -and $note.DueDate)
					{
						try
						{
							$due = [DateTime]::Parse($note.DueDate)
                        
							if ($now -ge $due)
							{
								if ($note.AutoRenew)
								{
									$created = [DateTime]::Parse($note.CreationDate)
									$originalDuration = $due - $created
									$newDueDate = [DateTime]::Now.Add($originalDuration)
									$note.DueDate = $newDueDate.ToString('yyyy-MM-dd HH:mm:ss')
									$note.CreationDate = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
									if ($note -is [System.Collections.IDictionary]) { $note['Notified'] = $false } else { $note | Add-Member -MemberType NoteProperty -Name 'Notified' -Value $false -Force }
									if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
									if (Get-Command RefreshNoteGrid -ErrorAction SilentlyContinue -Verbose:$False) { RefreshNoteGrid }
									continue
								}
								elseif (-not $note.Notified)
								{
									$notifKey = $key.GetHashCode()
									if ($notifKey -gt 0) { $notifKey = - $notifKey }

									CloseToast -Key $notifKey 

									$hideAction = [scriptblock]::Create("CloseToast -Key $notifKey")
									$editAction = [scriptblock]::Create("
                                    if (Get-Command EditNote -ErrorAction SilentlyContinue -Verbose:`$False) {
                                        `$grid = `$global:DashboardConfig.UI.NoteGrid
                                        if (`$grid) {
                                            `$rowToSelect = `$null
                                            foreach (`$r in `$grid.Rows) {
                                                if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') {
                                                    `$rowToSelect = `$r
                                                    break
                                                }
                                            }
                                            if (`$rowToSelect) {
                                                `$grid.ClearSelection()
                                                `$rowToSelect.Selected = `$true
                                                EditNote
                                            }
                                        }
                                    }
                                    CloseToast -Key $notifKey
                                ")
									$deleteAction = [scriptblock]::Create("
                                    if (Get-Command RemoveNote -ErrorAction SilentlyContinue -Verbose:`$False) {
                                        `$grid = `$global:DashboardConfig.UI.NoteGrid
                                        if (`$grid) {
                                            `$rowToSelect = `$null
                                            foreach (`$r in `$grid.Rows) {
                                                if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') {
                                                    `$rowToSelect = `$r
                                                    break
                                                }
                                            }
                                            if (`$rowToSelect) {
                                                `$grid.ClearSelection()
                                                `$rowToSelect.Selected = `$true
                                                RemoveNote
                                            }
                                        }
                                    }
                                    CloseToast -Key $notifKey
                                ")
									$btns = @( @{ Text = 'Hide'; Action = $hideAction }, @{ Text = 'Edit'; Action = $editAction }, @{ Text = 'Delete'; Action = $deleteAction } )
									if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue -Verbose:$False) { ShowInteractiveNotification -Title "Timer Finished: $($note.Title)" -Message $note.Content -Buttons $btns -Type 'Info' -Key $notifKey -TimeoutSeconds 0 -IgnoreCancellation } else { ShowToast -Title "Timer Finished: $($note.Title)" -Message $note.Content -Type 'Info' -TimeoutSeconds 0 -IgnoreCancellation }
									if ($note -is [System.Collections.IDictionary])
									{
										$note['Notified'] = $true
										$note['IsNotificationActive'] = $false
									}
									else
									{
										$note | Add-Member -MemberType NoteProperty -Name 'Notified' -Value $true -Force
										$note | Add-Member -MemberType NoteProperty -Name 'IsNotificationActive' -Value $false -Force
									}
								}
							} 
							else
							{
								if (($note -isnot [System.Collections.IDictionary]) -and (-not $note.PSObject.Properties['IsPaused'])) {
                                    $note | Add-Member -MemberType NoteProperty -Name 'IsPaused' -Value $false -Force
                                    $note | Add-Member -MemberType NoteProperty -Name 'RemainingOnPause' -Value $null -Force
                                    $note | Add-Member -MemberType NoteProperty -Name 'LastButtonState' -Value $null -Force
                                }

								$notifKey = $key.GetHashCode()
								if ($notifKey -gt 0) { $notifKey = - $notifKey }

								$pauseResumeAction = [scriptblock]::Create("
                                `$note = `$global:DashboardConfig.Config['Notes']['$key']
                                if (`$note) {
                                    if (`$note.IsPaused) {
                                        if (`$note.RemainingOnPause) {
                                            `$newDueDate = [DateTime]::Now.AddSeconds([double]`$note.RemainingOnPause)
                                            `$note.DueDate = `$newDueDate.ToString('yyyy-MM-dd HH:mm:ss')
                                            `$note.CreationDate = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
                                            `$note.IsPaused = `$false
                                            `$note.RemainingOnPause = `$null
                                        }
                                    } else {
                                        try {
                                            `$due = [DateTime]::Parse(`$note.DueDate)
                                            `$remaining = (`$due - [DateTime]::Now).TotalSeconds
                                            if (`$remaining -gt 0) {
                                                `$note.RemainingOnPause = `$remaining
                                                `$note.IsPaused = `$true
                                            }
                                        } catch {}
                                    }
                                    if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:`$False) { WriteConfig }
                                }
                            ")
								$editAction = [scriptblock]::Create("if (Get-Command EditNote -ErrorAction SilentlyContinue -Verbose:`$False) { `$grid = `$global:DashboardConfig.UI.NoteGrid; if (`$grid) { `$rowToSelect = `$null; foreach (`$r in `$grid.Rows) { if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') { `$rowToSelect = `$r; break } }; if (`$rowToSelect) { `$grid.ClearSelection(); `$rowToSelect.Selected = `$true; EditNote } }; CloseToast -Key $notifKey }")
								$deleteAction = [scriptblock]::Create("if (Get-Command RemoveNote -ErrorAction SilentlyContinue -Verbose:`$False) { `$grid = `$global:DashboardConfig.UI.NoteGrid; if (`$grid) { `$rowToSelect = `$null; foreach (`$r in `$grid.Rows) { if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') { `$rowToSelect = `$r; break } }; if (`$rowToSelect) { `$grid.ClearSelection(); `$rowToSelect.Selected = `$true; RemoveNote } }; CloseToast -Key $notifKey }")

								$currentState = if ($note.IsPaused) { 'Paused' } else { 'Running' }
                            
								if ($note.IsPaused)
								{
									$remainingOnPause = [timespan]::FromSeconds([double]$note.RemainingOnPause)
									$remainingString = if ($remainingOnPause.Days -gt 0) { $remainingOnPause.ToString("d'd 'h'h 'm'm 's's'") } elseif ($remainingOnPause.Hours -gt 0) { $remainingOnPause.ToString("h'h 'm'm 's's'") } elseif ($remainingOnPause.Minutes -gt 0) { $remainingOnPause.ToString("m'm 's's'") } else { $remainingOnPause.ToString("s's'") }
									$message = "PAUSED - Remaining: $remainingString`n$($note.Content)"
									$btns = @( @{ Text = 'Resume'; Action = $pauseResumeAction }, @{ Text = 'Edit'; Action = $editAction }, @{ Text = 'Delete'; Action = $deleteAction } )
									$titleText = "Timer Paused: $($note.Title)"
									$progress = -1
								}
								else
								{
									if (-not $note.CreationDate) { $note.CreationDate = [DateTime]::Now.AddSeconds(-1).ToString('yyyy-MM-dd HH:mm:ss') }
									$created = [DateTime]::Parse($note.CreationDate)
									$totalDuration = $due - $created
									$remaining = $due - $now
									$elapsed = $now - $created

									if ($totalDuration.TotalSeconds -gt 0) { $progress = [int]([Math]::Min(100, ($elapsed.TotalSeconds / $totalDuration.TotalSeconds) * 100)) }
									else { $progress = 100 }

									$parts = @()
									if ($remaining.Days -gt 0)
									{
										$parts = @("$($remaining.Days)d", "$($remaining.Hours)h", "$($remaining.Minutes)m", "$($remaining.Seconds)s")
									}
									elseif ($remaining.Hours -gt 0)
									{
										$parts = @("$($remaining.Hours)h", "$($remaining.Minutes)m", "$($remaining.Seconds)s")
									}
									elseif ($remaining.Minutes -gt 0)
									{
										$parts = @("$($remaining.Minutes)m", "$($remaining.Seconds)s")
									}
									else
									{
										$parts = @("$($remaining.Seconds)s")
									}
									$remainingString = $parts -join ' '

									$message = "Remaining: $remainingString`n$($note.Content)"
									$btns = @( @{ Text = 'Pause'; Action = $pauseResumeAction }, @{ Text = 'Edit'; Action = $editAction }, @{ Text = 'Delete'; Action = $deleteAction } )
									$titleText = "Timer: $($note.Title)"
								}

								$map = $global:DashboardConfig.State.LoginNotificationMap
								$notificationActive = ($map.ContainsKey($notifKey) -and -not $map[$notifKey].IsDisposed)

								if ($notificationActive -and $note.LastButtonState -eq $currentState)
								{
									ShowToast -Title $titleText -Message $message -Type 'Info' -Key $notifKey -TimeoutSeconds 0 -Progress $progress -IgnoreCancellation
								}
								else
								{
									ShowInteractiveNotification -Title $titleText -Message $message -Buttons $btns -Type 'Info' -Key $notifKey -TimeoutSeconds 0 -Progress $progress -IgnoreCancellation
									if ($note -is [System.Collections.IDictionary]) { $note['LastButtonState'] = $currentState } else { $note | Add-Member -MemberType NoteProperty -Name 'LastButtonState' -Value $currentState -Force }
								}

								if ($note -is [System.Collections.IDictionary])
								{
									$note['IsNotificationActive'] = $true
									$note['Notified'] = $false
								}
								else
								{
									$note | Add-Member -MemberType NoteProperty -Name 'IsNotificationActive' -Value $true -Force
									$note | Add-Member -MemberType NoteProperty -Name 'Notified' -Value $false -Force
								}
							}
						}
						catch { Write-Verbose "NoteTimer (Timer) Error for note '$($note.Title)': $_" }
					}
					elseif ($note.Type -eq 'Note')
					{
						if (($note.ShowOnStartup -and -not $note.StartupHandled) -or (-not $note.Notified))
						{
							$notifKey = $key.GetHashCode()
							if ($notifKey -gt 0) { $notifKey = - $notifKey }
                            
							$hideAction = [scriptblock]::Create("CloseToast -Key $notifKey")
							$deleteAction = [scriptblock]::Create("
                            if (Get-Command RemoveNote -ErrorAction SilentlyContinue -Verbose:`$False) {
                                `$grid = `$global:DashboardConfig.UI.NoteGrid
                                if (`$grid) {
                                    `$rowToSelect = `$null
                                    foreach (`$r in `$grid.Rows) {
                                        if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') {
                                            `$rowToSelect = `$r
                                            break
                                        }
                                    }
                                    if (`$rowToSelect) {
                                        `$grid.ClearSelection()
                                        `$rowToSelect.Selected = `$true
                                        RemoveNote
                                    }
                                }
                            }
                            CloseToast -Key $notifKey
                        ")
							$editAction = [scriptblock]::Create("
                            if (Get-Command EditNote -ErrorAction SilentlyContinue -Verbose:`$False) {
                                `$grid = `$global:DashboardConfig.UI.NoteGrid
                                if (`$grid) {
                                    `$rowToSelect = `$null
                                    foreach (`$r in `$grid.Rows) {
                                        if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') {
                                            `$rowToSelect = `$r
                                            break
                                        }
                                    }
                                    if (`$rowToSelect) {
                                        `$grid.ClearSelection()
                                        `$rowToSelect.Selected = `$true
                                        EditNote
                                    }
                                }
                            }
                            CloseToast -Key $notifKey
                        ")
							$btns = @( @{ Text = 'Hide'; Action = $hideAction }, @{ Text = 'Edit'; Action = $editAction }, @{ Text = 'Delete'; Action = $deleteAction } )
							if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue -Verbose:$False)
							{ 
								ShowInteractiveNotification -Title "Note: $($note.Title)" -Message $note.Content -Buttons $btns -Type 'Info' -Key $notifKey -TimeoutSeconds 0 -IgnoreCancellation
							}
							else
							{ 
								ShowToast -Title "Note: $($note.Title)" -Message $note.Content -Type 'Info' -TimeoutSeconds 0 -IgnoreCancellation
							}
							if ($note.ShowOnStartup)
							{
								if ($note -is [System.Collections.IDictionary]) { $note['StartupHandled'] = $true } else { $note | Add-Member -MemberType NoteProperty -Name 'StartupHandled' -Value $true -Force }
							}
							if ($note -is [System.Collections.IDictionary]) { $note['Notified'] = $true } else { $note | Add-Member -MemberType NoteProperty -Name 'Notified' -Value $true -Force }
						}
					}
					elseif ($note.Type -eq 'Reminder' -and $note.DueDate)
					{
						if ($note.ShowOnStartup -and -not $note.StartupHandled)
						{
							$notifKey = $key.GetHashCode()
							if ($notifKey -gt 0) { $notifKey = - $notifKey }
                            
							$hideAction = [scriptblock]::Create("CloseToast -Key $notifKey")
							$editAction = [scriptblock]::Create("
                            if (Get-Command EditNote -ErrorAction SilentlyContinue -Verbose:`$False) {
                                `$grid = `$global:DashboardConfig.UI.NoteGrid
                                if (`$grid) {
                                    `$rowToSelect = `$null
                                    foreach (`$r in `$grid.Rows) {
                                        if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') {
                                            `$rowToSelect = `$r
                                            break
                                        }
                                    }
                                    if (`$rowToSelect) {
                                        `$grid.ClearSelection()
                                        `$rowToSelect.Selected = `$true
                                        EditNote
                                    }
                                }
                            }
                            CloseToast -Key $notifKey
                        ")
							$deleteAction = [scriptblock]::Create("
                            if (Get-Command RemoveNote -ErrorAction SilentlyContinue -Verbose:`$False) {
                                `$grid = `$global:DashboardConfig.UI.NoteGrid
                                if (`$grid) {
                                    `$rowToSelect = `$null
                                    foreach (`$r in `$grid.Rows) {
                                        if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') {
                                            `$rowToSelect = `$r
                                            break
                                        }
                                    }
                                    if (`$rowToSelect) {
                                        `$grid.ClearSelection()
                                        `$rowToSelect.Selected = `$true
                                        RemoveNote
                                    }
                                }
                            }
                            CloseToast -Key $notifKey
                        ")
							$btns = @(
								@{ Text = 'Hide'; Action = $hideAction },
								@{ Text = 'Edit'; Action = $editAction },
								@{ Text = 'Delete'; Action = $deleteAction }
							)
							$message = "Due: $($note.DueDate)`n$($note.Content)"
							if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue -Verbose:$False) { ShowInteractiveNotification -Title "$($note.Type): $($note.Title)" -Message $message -Buttons $btns -Type 'Info' -Key $notifKey -TimeoutSeconds 0 -IgnoreCancellation } else { ShowToast -Title "$($note.Type): $($note.Title)" -Message $message -Type 'Info' -TimeoutSeconds 0 -IgnoreCancellation }
							if ($note -is [System.Collections.IDictionary]) { $note['StartupHandled'] = $true } else { $note | Add-Member -MemberType NoteProperty -Name 'StartupHandled' -Value $true -Force }
						}

						try
						{
							$due = [DateTime]::Parse($note.DueDate)
							if ($now -ge $due -and $note.AutoRenew -and $note.Notified)
							{
								$newDueDate = &$GetNewDueDate $due $note.RenewInterval
								$note.DueDate = $newDueDate.ToString('yyyy-MM-dd HH:mm:ss')
								if ($note -is [System.Collections.IDictionary]) { $note['Notified'] = $false } else { $note | Add-Member -MemberType NoteProperty -Name 'Notified' -Value $false -Force }
								if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
								if (Get-Command RefreshNoteGrid -ErrorAction SilentlyContinue -Verbose:$False) { RefreshNoteGrid }
								continue
							}

							if ($now -ge $due -and -not $note.Notified)
							{
								$notifKey = $key.GetHashCode()
								if ($notifKey -gt 0) { $notifKey = - $notifKey }
                            
                            
								$editAction = [scriptblock]::Create("
                                if (Get-Command EditNote -ErrorAction SilentlyContinue -Verbose:`$False) {
                                    `$grid = `$global:DashboardConfig.UI.NoteGrid
                                    if (`$grid) {
                                        `$rowToSelect = `$null
                                        foreach (`$r in `$grid.Rows) {
                                            if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') {
                                                `$rowToSelect = `$r
                                                break
                                            }
                                        }
                                        if (`$rowToSelect) {
                                            `$grid.ClearSelection()
                                            `$rowToSelect.Selected = `$true
                                            EditNote
                                        }
                                    }
                                }
                                CloseToast -Key $notifKey
                            ")
								$deleteAction = [scriptblock]::Create("
                                if (Get-Command RemoveNote -ErrorAction SilentlyContinue -Verbose:`$False) {
                                    `$grid = `$global:DashboardConfig.UI.NoteGrid
                                    if (`$grid) {
                                        `$rowToSelect = `$null
                                        foreach (`$r in `$grid.Rows) {
                                            if (`$r.Tag -and `$r.Tag.ToString() -eq '$key') {
                                                `$rowToSelect = `$r
                                                break
                                            }
                                        }
                                        if (`$rowToSelect) {
                                            `$grid.ClearSelection()
                                            `$rowToSelect.Selected = `$true
                                            RemoveNote
                                        }
                                    }
                                }
                                CloseToast -Key $notifKey
                            ")
								$btns = @(
									@{ Text = 'Hide'; Action = $hideAction },
									@{ Text = 'Edit'; Action = $editAction },
									@{ Text = 'Delete'; Action = $deleteAction }
								)
								$message = "Due: $($note.DueDate)`n$($note.Content)"
								try
								{
									if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue -Verbose:$False) { ShowInteractiveNotification -Title "$($note.Type): $($note.Title)" -Message $message -Buttons $btns -Type 'Info' -Key $notifKey -TimeoutSeconds 0 -IgnoreCancellation } else { ShowToast -Title "$($note.Type): $($note.Title)" -Message $message -Type 'Info' -TimeoutSeconds 0 -IgnoreCancellation }
								}
								catch
								{
									Write-Verbose "Reminder Notification Failed: $_"
								}
                            
								if ($note.AutoRenew)
								{
									$newDueDate = &$GetNewDueDate $due $note.RenewInterval
									$note.DueDate = $newDueDate.ToString('yyyy-MM-dd HH:mm:ss')
									if ($note -is [System.Collections.IDictionary]) { $note['Notified'] = $false } else { $note | Add-Member -MemberType NoteProperty -Name 'Notified' -Value $false -Force }
									if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
									if (Get-Command RefreshNoteGrid -ErrorAction SilentlyContinue -Verbose:$False) { RefreshNoteGrid }
								}
								else
								{
									if ($note -is [System.Collections.IDictionary]) { $note['Notified'] = $true } else { $note | Add-Member -MemberType NoteProperty -Name 'Notified' -Value $true -Force }
									if (Get-Command WriteConfig -ErrorAction SilentlyContinue -Verbose:$False) { WriteConfig }
								}
							}
						}
						catch { Write-Verbose "NoteTimer (Reminder) Error for note '$($note.Title)': $_" }
					}
				}
			}
		})
	$noteTimer.Start()
	$global:DashboardConfig.State.NoteTimer = $noteTimer
}
function UpdateUpgradeItemSelection
{
	try
	{
		$UI = $global:DashboardConfig.UI

		if ($null -eq $UI -or $null -eq $UI.UpgradeItemSelector -or $null -eq $UI.UpgradeTypeSelector ) { return }

		$data = if ($global:DashboardConfig.Resources) { $global:DashboardConfig.Resources.UpgradeData } else { $null }
		if ($null -eq $data -or $null -eq $UI.UpgradeItemSelector.SelectedItem) { return }

		$sel = $UI.UpgradeItemSelector.SelectedItem.ToString()
		$max = 0

		$categoryObject = if ($data -is [System.Collections.IDictionary]) { $data[$sel] } else { $data.$sel }

		if ($null -ne $categoryObject)
		{
			$firstPropName = $categoryObject.psobject.Properties.Name | Select-Object -First 1

			if ($firstPropName)
			{
				$firstPropValue = if ($categoryObject -is [System.Collections.IDictionary])
				{
					$categoryObject[$firstPropName]
				}
				else
				{
					$categoryObject.$firstPropName
				}

				if ($firstPropValue -is [System.Collections.ICollection])
				{
					$max = $firstPropValue.Count
				}
			}
		}

		$UI.UpgradeLvlStart.Items.Clear()
		$UI.UpgradeLvlEnd.Items.Clear()

		for ($i = 0; $i -lt $max; $i++)
		{
			$UI.UpgradeLvlStart.Items.Add($i) | Out-Null
			$UI.UpgradeLvlEnd.Items.Add($i + 1) | Out-Null
		}

		$UI.UpgradeLvlStart.SelectedIndex = if ($UI.UpgradeLvlStart.Items.Count) { 0 } else { -1 }
		$UI.UpgradeLvlEnd.SelectedIndex   = if ($UI.UpgradeLvlEnd.Items.Count)   { $max - 1 } else { -1 }

		$UI.UpgradeTypeSelector.Enabled = ($sel -eq "Weapon")

		if ($sel -eq "Weapon")
		{
			if ($UI.UpgradeTypeSelector.SelectedIndex -lt 0)
			{
				$UI.UpgradeTypeSelector.SelectedIndex = 0
			}
		}
		else
		{
			$UI.UpgradeTypeSelector.SelectedIndex = -1
		}

	}
	catch
	{
		Write-Verbose "Error in Update-UpgradeItemSelection: $_"
	}
}

function UpdateUpgradeTypeSelection
{
	try
	{
		$UI = $global:DashboardConfig.UI
		if ($null -eq $UI -or $null -eq $UI.UpgradeItemSelector -or $null -eq $UI.UpgradeTypeSelector) { return }
		if ($null -eq $UI.UpgradeItemSelector.SelectedItem -or $null -eq $UI.UpgradeTypeSelector.SelectedItem) { return }
		if ($UI.UpgradeItemSelector.SelectedItem.ToString() -ne "Weapon" -or $UI.UpgradeTypeSelector.SelectedItem.ToString() -ne "2H") { return }
		
	}
	catch
	{
		Write-Verbose "Error in Update-UpgradeTypeSelection: $_"
	}
}

function CalculateUpgradeTotalMaterials
{
	try
	{
		$UI = $global:DashboardConfig.UI

		if (
			$null -eq $UI -or
			$null -eq $UI.UpgradeResultsGrid -or
			$null -eq $UI.UpgradeTotalPerin -or
			$null -eq $UI.UpgradeTotalEcp
		) { return }

		$cart = if ($global:DashboardConfig.Resources) { $global:DashboardConfig.Resources.UpgradeCart } else { $null }
		$data = if ($global:DashboardConfig.Resources) { $global:DashboardConfig.Resources.UpgradeData } else { $null }

		$UI.UpgradeResultsGrid.Rows.Clear()

		if ($null -eq $cart -or $cart.Count -eq 0 -or $null -eq $data)
		{
			$UI.UpgradeTotalPerin.Text = "Total Perin: 0"
			$UI.UpgradeTotalEcp.Text = "Total ECP: 0"
			return
		}

		$materialTotals = [ordered]@{}
		$perinTotal = 0
		$ecpTotal = 0

		foreach ($item in $cart)
		{
			if ($null -eq $item -or -not $item.Category) { continue }

			$catData = if ($data -is [System.Collections.IDictionary])
			{
				$data[$item.Category]
			}
			else
			{
				$data.$($item.Category)
			}

			if ($null -eq $catData) { continue }

			foreach ($prop in $catData.psobject.Properties)
			{
				$mName = $prop.Name
				$costs = $prop.Value

				if ($null -eq $costs) { continue }

				$sum = 0

				for ($i = $item.CurrentLevel; $i -lt $item.TargetLevel; $i++)
				{
					if ($i -ge 0 -and $i -lt $costs.Count)
					{
						$sum += $costs[$i]
					}
				}

				$q = 1
				try {
					if ($item -and $item.PSObject -and $item.PSObject.Properties['Quantity']) {
						$q = [int]$item.Quantity
					}
				} catch {
					$q = 1
				}
				$total = $sum * $q * $item.Multiplier

				if (-not $materialTotals.Contains($mName))
				{
					$materialTotals[$mName] = 0
				}

				$materialTotals[$mName] += $total

				switch ($mName)
				{
					"Perin" { $perinTotal += $total }
					"ECP"   { $ecpTotal += $total }
				}
			}
		}

		$hideZero = $UI.UpgradeHideZero -and $UI.UpgradeHideZero.Checked

		foreach ($m in $materialTotals.Keys)
		{
			$amt = $materialTotals[$m]

			if ($hideZero -and $amt -eq 0) { continue }

			$UI.UpgradeResultsGrid.Rows.Add(
				$m,
				("{0:N0}" -f $amt)
			) | Out-Null
		}

		$UI.UpgradeTotalPerin.Text = "Total Perin: {0:N0}" -f $perinTotal
		$UI.UpgradeTotalEcp.Text = "Total ECP: {0:N0}" -f $ecpTotal
	}
	catch
	{
		Write-Verbose "Error in Calculate-UpgradeTotalMaterials: $_"
	}
}

function AddUpgradeItem
{
	try
	{
		$UI = $global:DashboardConfig.UI

		if (
			$null -eq $UI -or
			$null -eq $UI.UpgradeLvlStart -or
			$null -eq $UI.UpgradeLvlEnd -or
			$null -eq $UI.UpgradeItemSelector -or
			$null -eq $UI.UpgradeCartGrid
		) { return }


		if ($null -eq $UI.UpgradeLvlStart.SelectedItem -or $null -eq $UI.UpgradeLvlEnd.SelectedItem)
		{
			return
		}

		$cur = [int]$UI.UpgradeLvlStart.SelectedItem
		$tar = [int]$UI.UpgradeLvlEnd.SelectedItem

		if ($tar -le $cur)
		{
			return
		}

		if ($null -eq $UI.UpgradeItemSelector.SelectedItem) { return }

		$cat = $UI.UpgradeItemSelector.SelectedItem.ToString()

		$type = $null

		if (
			$cat -eq "Weapon" -and
			$null -ne $UI.UpgradeTypeSelector -and
			$null -ne $UI.UpgradeTypeSelector.SelectedItem
		)
		{
			$type = $UI.UpgradeTypeSelector.SelectedItem.ToString()
		}

		$mult = 1
		$dispName = $cat

		if ($cat -eq "Weapon" -and $type)
		{
			$dispName = "Weapon ($type)"

			if ($type -eq "2H")
			{
				$mult = 2
			}
		}

		$qty = 1

		$newItem = [PSCustomObject]@{
			Category     = $cat
			Display      = $dispName
			CurrentLevel = $cur
			TargetLevel  = $tar
			Multiplier   = $mult
			Quantity     = $qty
		}

		if ($null -eq $global:DashboardConfig.Resources)
		{
			$global:DashboardConfig | Add-Member -MemberType NoteProperty -Name Resources -Value @{} -Force
		}

		$cartRef = $global:DashboardConfig.Resources.UpgradeCart
		if ($null -eq $cartRef -or -not ($cartRef -is [System.Collections.IList])) {
			$newCart = [System.Collections.ArrayList]::new()
			if ($cartRef -ne $null) {
				try {
					if ($cartRef -is [System.Collections.IEnumerable]) {
						foreach ($el in $cartRef) { $newCart.Add($el) | Out-Null }
					} else {
						$newCart.Add($cartRef) | Out-Null
					}
				} catch { }
			}
			$global:DashboardConfig.Resources.UpgradeCart = $newCart
			$cartRef = $newCart
		}

		$global:DashboardConfig.Resources.UpgradeCart.Add($newItem) | Out-Null

		$UI.UpgradeCartGrid.Rows.Add(
			$newItem.Quantity,
			$dispName,
			"+$cur -> +$tar"
		) | Out-Null

		CalculateUpgradeTotalMaterials
	}
	catch
	{
		Write-Verbose "Error in Add-UpgradeItem: $_"
	}
}

function RemoveUpgradeItem
{
	try
	{
		$UI = $global:DashboardConfig.UI

		if ($null -eq $UI -or $null -eq $UI.UpgradeCartGrid) { return }
		if ($UI.UpgradeCartGrid.SelectedRows.Count -eq 0) { return }

		$selRow = $UI.UpgradeCartGrid.SelectedRows[0]

		try { $dispVal = $selRow.Cells[1].Value.ToString() } catch { $dispVal = $null }
		try { $levelsVal = $selRow.Cells[2].Value.ToString() } catch { $levelsVal = $null }

		$cart = if ($global:DashboardConfig.Resources) { $global:DashboardConfig.Resources.UpgradeCart } else { $null }

		function GetRebuiltCart
		{
			param($Collection, $ItemToRemove)

			$new = [System.Collections.ArrayList]::new()
			$removed = $false

			foreach ($el in $Collection)
			{
				if (-not $removed -and $el -eq $ItemToRemove)
				{
					$removed = $true
					continue
				}

				$new.Add($el) | Out-Null
			}

			$new
		}

		if ($null -ne $cart -and $null -ne $dispVal -and $null -ne $levelsVal)
		{
			$cur = $null
			$tar = $null

			if ($levelsVal -match '\+(\d+)\s*->\s*\+(\d+)')
			{
				$cur = [int]$matches[1]
				$tar = [int]$matches[2]
			}

			$toRemove = $cart | Where-Object {
				($null -eq $dispVal -or $_.Display -eq $dispVal) -and
				($null -eq $cur -or $_.CurrentLevel -eq $cur) -and
				($null -eq $tar -or $_.TargetLevel -eq $tar)
			} | Select-Object -First 1

			if ($null -ne $toRemove)
			{
				try
				{
					if ($cart -is [System.Collections.ArrayList]) {
						$cart.Remove($toRemove) | Out-Null
					} elseif ($cart -is [System.Management.Automation.PSCustomObject] -and $cart -eq $toRemove) {
						$global:DashboardConfig.Resources.UpgradeCart = [System.Collections.ArrayList]::new()
					} else {
						try {
							if ($cart -is [System.Collections.IList]) { $cart.Remove($toRemove) | Out-Null }
							else { $global:DashboardConfig.Resources.UpgradeCart = GetRebuiltCart $cart $toRemove }
						} catch {
							$global:DashboardConfig.Resources.UpgradeCart = GetRebuiltCart $cart $toRemove
						}
					}
				}
				catch
				{
					$global:DashboardConfig.Resources.UpgradeCart = [System.Collections.ArrayList]::new()
				}
			}
		}

		foreach ($r in $UI.UpgradeCartGrid.SelectedRows)
		{
			$UI.UpgradeCartGrid.Rows.Remove($r)
		}

		CalculateUpgradeTotalMaterials
	}
	catch
	{
		Write-Verbose "Error in Remove-UpgradeItem: $_"
	}
}

function ResetUpgradeCalculator
{
	try
	{
		$UI = $global:DashboardConfig.UI
		if ($null -eq $UI) { return }
		if ($null -eq $global:DashboardConfig.Resources) { $global:DashboardConfig | Add-Member -MemberType NoteProperty -Name Resources -Value @{} -Force }
		$global:DashboardConfig.Resources.UpgradeCart = [System.Collections.ArrayList]::new()
		if ($null -ne $UI.UpgradeCartGrid) { $UI.UpgradeCartGrid.Rows.Clear() }
		if ($null -ne $UI.UpgradeItemSelector -and $UI.UpgradeItemSelector.Items.Count -gt 0) {
			$UI.UpgradeItemSelector.SelectedIndex = 0
			try { Update-UpgradeItemSelection } catch { }
		}
		CalculateUpgradeTotalMaterials
	}
	catch
	{
		Write-Verbose "Error in Reset-UpgradeCalculator: $_"
	}
}

function HandleUpgradeCartRowsAdded
{
	$UI = $global:DashboardConfig.UI
	if ($null -eq $UI -or $null -eq $UI.UpgradeCartGrid -or $UI.UpgradeCartGrid.Rows.Count -eq 0) { return }
	$UI.UpgradeCartGrid.FirstDisplayedScrollingRowIndex = $UI.UpgradeCartGrid.Rows.Count - 1
}

Export-ModuleMember -Function *