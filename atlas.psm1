

function Initialize-AtlasCore
{
	if ($script:AtlasCoreLoaded) { return $true }
	if (([System.Management.Automation.PSTypeName]'Atlas.Canvas').Type) { $script:AtlasCoreLoaded = $true; return $true }
	try
	{
		$init = Get-Command -Name 'Initialize-DashboardAtlasTypes' -ErrorAction SilentlyContinue
		if (-not $init)
		{
			Import-Module (Join-Path $PSScriptRoot 'classes.psm1') -Force -DisableNameChecking -ErrorAction Stop
			$init = Get-Command -Name 'Initialize-DashboardAtlasTypes' -ErrorAction Stop
		}
		$null = & $init
		$script:AtlasCoreLoaded = $true
		return $true
	}
	catch
	{
		Write-Warning "Atlas compiled core did not load: $($_.Exception.Message)"
		return $false
	}
}
#endregion

#region PowerShell - path resolution and config

function Get-AtlasConfigValue
{
	param([Parameter(Mandatory = $true)][string]$Key, [string]$Default = '')
	try
	{
		if ($global:DashboardConfig -and $global:DashboardConfig.Config -and $global:DashboardConfig.Config['Atlas'])
		{
			$cfg = $global:DashboardConfig.Config['Atlas']
			if ($cfg -is [System.Collections.IDictionary] -and $cfg.Contains($Key))
			{
				$v = [string]$cfg[$Key]
				if ($v) { return $v }
			}
		}
	}
	catch { }
	return $Default
}

function Set-AtlasConfigValue
{
	param([Parameter(Mandatory = $true)][string]$Key, [AllowEmptyString()][string]$Value = '')
	try
	{
		if (-not $global:DashboardConfig) { return }
		if (-not $global:DashboardConfig.Config) { return }
		if (-not $global:DashboardConfig.Config['Atlas'])
		{
			$global:DashboardConfig.Config['Atlas'] = [ordered]@{}
		}
		$global:DashboardConfig.Config['Atlas'][$Key] = $Value
	}
	catch { }
}

function Get-AtlasCandidateRoots
{
	$list = New-Object System.Collections.Generic.List[string]
	$add = {
		param($p)
		if ($p -and (Test-Path -LiteralPath $p)) { if (-not $list.Contains($p)) { $list.Add($p) } }
	}
	if ($global:DashboardConfig -and $global:DashboardConfig.Paths -and $global:DashboardConfig.Paths.Source)
	{
		& $add $global:DashboardConfig.Paths.Source
	}
	return $list
}

function Get-AtlasHomeDataStamp
{
	param([string]$Home)
	$best = [long]0
	foreach ($rel in @('src\data.js', 'data.js'))
	{
		$p = Join-Path $Home $rel
		try
		{
			if (Test-Path -LiteralPath $p -PathType Leaf)
			{
				$t = [System.IO.File]::GetLastWriteTimeUtc($p).Ticks
				if ($t -gt $best) { $best = $t }
			}
		}
		catch { }
	}
	return $best
}

function Get-AtlasHomes
{
	$list = New-Object System.Collections.Generic.List[string]
	$add = {
		param($p)
		if ($p -and (Test-Path -LiteralPath $p)) { if (-not $list.Contains($p)) { $list.Add($p) } }
	}
	$paths = $null
	try { $paths = $global:DashboardConfig.Paths } catch { }
	if ($paths)
	{
		if ($paths.App) { & $add (Join-Path $paths.App 'atlas') }
		if ($paths.Modules) { & $add (Join-Path $paths.Modules 'atlas') }
		if ($paths.Source) { & $add (Join-Path $paths.Source 'atlas') }
	}
	if ($list.Count -gt 1)
	{
		$ranked = @()
		for ($i = 0; $i -lt $list.Count; $i++)
		{
			$ranked += [pscustomobject]@{ Path = $list[$i]; Stamp = (Get-AtlasHomeDataStamp -Home $list[$i]); Order = $i }
		}
		$ordered = @($ranked | Sort-Object -Property @{ Expression = { $_.Stamp }; Descending = $true }, @{ Expression = { $_.Order }; Descending = $false })
		$list.Clear()
		foreach ($r in $ordered) { $list.Add($r.Path) }
	}
	return $list
}

function Get-AtlasAssetRoot
{
	$cfg = Get-AtlasConfigValue 'AssetRoot'
	if ($cfg -and (Test-Path -LiteralPath $cfg))
	{
		if (Test-Path -LiteralPath (Join-Path $cfg 'maps')) { return $cfg }
	}
	foreach ($atlasHome in (Get-AtlasHomes))
	{
		$own = Join-Path $atlasHome 'icons'
		if (Test-Path -LiteralPath (Join-Path $own 'maps')) { return $own }
	}
	foreach ($r in (Get-AtlasCandidateRoots))
	{
		if (Test-Path -LiteralPath (Join-Path $r 'maps')) { return $r }
	}
	foreach ($atlasHome in (Get-AtlasHomes))
	{
		return (Join-Path $atlasHome 'icons')
	}
	return $null
}

function Get-AtlasDataPath
{
	$cfg = Get-AtlasConfigValue 'DataPath'
	if ($cfg -and (Test-Path -LiteralPath $cfg)) { return $cfg }

	foreach ($atlasHome in (Get-AtlasHomes))
	{
		$p = Join-Path $atlasHome 'src\data.js'
		if (Test-Path -LiteralPath $p) { return $p }
		$p = Join-Path $atlasHome 'data.js'
		if (Test-Path -LiteralPath $p) { return $p }
	}

	foreach ($r in (Get-AtlasCandidateRoots))
	{
		$p = Join-Path $r 'atlas\src\data.js'
		if (Test-Path -LiteralPath $p) { return $p }
		$p = Join-Path $r 'atlas\data.js'
		if (Test-Path -LiteralPath $p) { return $p }
		$p = Join-Path $r 'src\data.js'
		if (Test-Path -LiteralPath $p) { return $p }
	}
	return $null
}

function Get-AtlasData
{
	Initialize-AtlasCore | Out-Null
	if ($script:AtlasData) { return $script:AtlasData }
	$dp = Get-AtlasDataPath
	$ar = Get-AtlasAssetRoot
	if (-not $dp)
	{
		throw "Atlas data.js was not found. Looked beside the module and beside the dashboard source. Set Atlas\DataPath in config.ini to point at it."
	}
	if (-not $ar) { $ar = Split-Path -Parent $dp }
	$script:AtlasData = [Atlas.Data]::Load($dp, $ar)
	return $script:AtlasData
}

#endregion

#region PowerShell - atlas files and updates

$script:AtlasRemoteDefault = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/atlas'
$script:AtlasManifest = $null
$script:AtlasManifestTried = [datetime]::MinValue
$script:AtlasDownloadState = $null
$script:AtlasOpenPending = $false

function Get-AtlasRemoteBase
{
	$r = Get-AtlasConfigValue 'Remote' ''
	if ($r) { return $r.TrimEnd('/') }
	return $script:AtlasRemoteDefault
}

function Get-AtlasInstallRoot
{
	$cand = New-Object System.Collections.Generic.List[string]
	$paths = $null
	try { $paths = $global:DashboardConfig.Paths } catch { }
	if ($paths)
	{
		if ($paths.App) { $cand.Add((Join-Path $paths.App 'atlas')) }
		if ($paths.Modules) { $cand.Add((Join-Path $paths.Modules 'atlas')) }
		if ($paths.Source) { $cand.Add((Join-Path $paths.Source 'atlas')) }
	}
	foreach ($p in $cand) { if (Test-Path -LiteralPath $p) { return $p } }
	foreach ($p in $cand)
	{
		try { [void][System.IO.Directory]::CreateDirectory($p); return $p } catch { }
	}
	return $null
}

function Get-AtlasFileHash
{
	param([string]$Path)
	$sha = $null; $fs = $null
	try
	{
		$sha = [System.Security.Cryptography.SHA256]::Create()
		$fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
		$bytes = $sha.ComputeHash($fs)
		return ([System.BitConverter]::ToString($bytes).Replace('-', '').ToLowerInvariant())
	}
	catch { return '' }
	finally
	{
		if ($fs) { $fs.Dispose() }
		if ($sha) { $sha.Dispose() }
	}
}

function Get-AtlasBytesHash
{
	param([byte[]]$Bytes)
	$sha = $null
	try
	{
		$sha = [System.Security.Cryptography.SHA256]::Create()
		$hash = $sha.ComputeHash($Bytes)
		return ([System.BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant())
	}
	catch { return '' }
	finally
	{
		if ($sha) { $sha.Dispose() }
	}
}

function Format-AtlasSize
{
	param([long]$Bytes)
	if ($Bytes -ge 1GB) { return ('{0:0.0} GB' -f ($Bytes / 1GB)) }
	if ($Bytes -ge 1MB) { return ('{0:0.0} MB' -f ($Bytes / 1MB)) }
	if ($Bytes -ge 1KB) { return ('{0:0} KB' -f ($Bytes / 1KB)) }
	return ('{0} B' -f $Bytes)
}

function Enable-AtlasTls
{
	try
	{
		$cur = [System.Net.ServicePointManager]::SecurityProtocol
		[System.Net.ServicePointManager]::SecurityProtocol = $cur -bor [System.Net.SecurityProtocolType]::Tls12
	}
	catch { }
}

function Get-AtlasWebText
{
	param([string]$Url, [int]$Timeout = 20000)
	Enable-AtlasTls
	$resp = $null; $sr = $null
	$req = [System.Net.HttpWebRequest]::Create($Url)
	$req.Timeout = $Timeout
	$req.ReadWriteTimeout = $Timeout
	$req.UserAgent = 'EntropiaDashboard'
	try { $req.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate } catch { }
	try
	{
		$resp = $req.GetResponse()
		$sr = New-Object System.IO.StreamReader($resp.GetResponseStream(), [System.Text.Encoding]::UTF8)
		return $sr.ReadToEnd()
	}
	finally
	{
		if ($sr) { $sr.Dispose() }
		if ($resp) { $resp.Close() }
	}
}

function Invoke-AtlasDownloadFile
{
	param([string]$Url, [string]$Path, [string]$Hash = '')
	Enable-AtlasTls
	$dir = Split-Path -Parent $Path
	if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
	$part = $Path + '.part'
	$resp = $null; $src = $null; $dst = $null
	$req = [System.Net.HttpWebRequest]::Create($Url)
	$req.Timeout = 30000
	$req.ReadWriteTimeout = 120000
	$req.UserAgent = 'EntropiaDashboard'
	try
	{
		$resp = $req.GetResponse()
		$src = $resp.GetResponseStream()
		$dst = [System.IO.File]::Create($part)
		$src.CopyTo($dst)
		$dst.Dispose(); $dst = $null
		$src.Dispose(); $src = $null
		$resp.Close(); $resp = $null
		if ($Hash)
		{
			$got = Get-AtlasFileHash $part
			if (-not $got -or $got -ne $Hash) { throw 'The downloaded file did not match its published checksum.' }
		}
		[System.IO.File]::Copy($part, $Path, $true)
		Remove-Item -LiteralPath $part -Force
		return $true
	}
	finally
	{
		if ($dst) { $dst.Dispose() }
		if ($src) { $src.Dispose() }
		if ($resp) { $resp.Close() }
		if (Test-Path -LiteralPath $part) { try { Remove-Item -LiteralPath $part -Force } catch { } }
	}
}

function Get-AtlasManifest
{
	param([switch]$Force)
	$tried = $script:AtlasManifestTried
	$age = ((Get-Date) - $tried).TotalMinutes
	if (-not $Force -and $tried -gt [datetime]::MinValue -and $age -lt 10) { return $script:AtlasManifest }
	$script:AtlasManifestTried = Get-Date
	$script:AtlasManifest = $null
	$url = (Get-AtlasRemoteBase) + '/manifest.json'
	try
	{
		$txt = Get-AtlasWebText -Url $url -Timeout 10000
		if ($txt)
		{
			$obj = $txt | ConvertFrom-Json
			if ($obj -and $obj.data) { $script:AtlasManifest = $obj }
		}
	}
	catch { $script:AtlasManifest = $null }
	return $script:AtlasManifest
}

function Read-AtlasUpdateAnswer
{
	param([string]$Text, [string]$Caption = 'Entropia Map')
	if (Test-DashboardCommand 'Show-DarkMessageBox')
	{
		try
		{
			$r = Show-DarkMessageBox -Text $Text -Caption $Caption -Buttons 'YesNo' -Icon 'Information'
			return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
		}
		catch { }
	}
	try { return (([System.Windows.Forms.MessageBox]::Show($Text, $Caption, 'YesNo', 'Question')) -eq [System.Windows.Forms.DialogResult]::Yes) } catch { }
	return $false
}

function Show-AtlasNotice
{
	param([string]$Text, [string]$Caption = 'Entropia Map', [string]$Icon = 'Information')
	if (Test-DashboardCommand 'Show-DarkMessageBox')
	{
		try { Show-DarkMessageBox -Text $Text -Caption $Caption -Buttons 'Ok' -Icon $Icon | Out-Null; return } catch { }
	}
	try { [void][System.Windows.Forms.MessageBox]::Show($Text, $Caption, 'OK', $Icon) } catch { }
}

function New-AtlasDownloadClient
{
	param([int]$Threads = 12)
	try { Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue } catch { }
	try
	{
		$client = [System.Net.Http.HttpClient]::new()
		$client.Timeout = [TimeSpan]::FromSeconds(120)
		$client.DefaultRequestHeaders.UserAgent.ParseAdd('EntropiaDashboard')
		return $client
	}
	catch { return $null }
}

function New-AtlasDownloadState
{
	Enable-AtlasTls
	try { Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue } catch { }
	$pal = Get-AtlasPalette
	$form = New-Object System.Windows.Forms.Form
	$form.Text = 'Entropia Atlas - Downloading'
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
	$form.MinimizeBox = $false
	$form.MaximizeBox = $false
	$form.ShowInTaskbar = $false
	$form.ControlBox = $false
	$form.TopMost = $false
	$form.ClientSize = New-Object System.Drawing.Size(430, 118)
	$form.BackColor = $pal.Panel
	$form.ForeColor = $pal.Fg

	$lblStatus = New-Object System.Windows.Forms.Label
	$lblStatus.SetBounds(16, 14, 400, 18)
	$lblStatus.Text = 'Preparing...'
	$lblStatus.ForeColor = $pal.Fg
	$lblStatus.BackColor = [System.Drawing.Color]::Transparent
	$form.Controls.Add($lblStatus)

	$bar = New-Object System.Windows.Forms.ProgressBar
	$bar.SetBounds(16, 38, 398, 14)
	$bar.Minimum = 0
	$bar.Maximum = 1
	$bar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
	$form.Controls.Add($bar)

	$lblBytes = New-Object System.Windows.Forms.Label
	$lblBytes.SetBounds(16, 58, 260, 18)
	$lblBytes.Text = ''
	$lblBytes.ForeColor = $pal.Mut
	$lblBytes.BackColor = [System.Drawing.Color]::Transparent
	$form.Controls.Add($lblBytes)

	$btnCancel = New-AtlasButton -Text 'Cancel' -Width 80
	$btnCancel.SetBounds(334, 82, 80, 24)
	$form.Controls.Add($btnCancel)

	$threads = 12
	try
	{
		$rawThreads = Get-AtlasConfigValue 'DownloadThreads' ''
		$n = 0
		if ([int]::TryParse($rawThreads, [ref]$n) -and $n -gt 0) { $threads = [Math]::Min(32, $n) }
	}
	catch { }
	try
	{
		if ([System.Net.ServicePointManager]::DefaultConnectionLimit -lt $threads)
		{
			[System.Net.ServicePointManager]::DefaultConnectionLimit = $threads
		}
	}
	catch { }

	$state = @{
		Form = $form
		Status = $lblStatus
		Bar = $bar
		Bytes = $lblBytes
		CancelButton = $btnCancel
		Queue = (New-Object System.Collections.Generic.List[object])
		Job = $null
		Client = $null
		Threads = $threads
		CancelAll = $false
		Timer = $null
	}
	$state.Client = New-AtlasDownloadClient -Threads $threads

	$btnCancel.Add_Click({
			$st = $script:AtlasDownloadState
			if ($null -eq $st) { return }
			$st.CancelAll = $true
			try { $st.CancelButton.Enabled = $false } catch { }
		})

	$timer = New-Object System.Windows.Forms.Timer
	$timer.Interval = 40
	$timer.Add_Tick({ Update-AtlasDownload })
	$state.Timer = $timer
	$script:AtlasDownloadState = $state

	try
	{
		if ($global:DashboardConfig -and $global:DashboardConfig.UI -and $global:DashboardConfig.UI.MainForm -and -not $global:DashboardConfig.UI.MainForm.IsDisposed)
		{
			$form.Owner = $global:DashboardConfig.UI.MainForm
		}
	}
	catch { }

	$form.Show()
	try { $form.Activate() } catch { }
	$timer.Start()
	return $state
}

function Start-AtlasDownload
{
	param([string]$Title, [object[]]$Items, [scriptblock]$OnComplete, [string]$Kind = '')

	if ($null -eq $Items) { $Items = @() }
	if ($Items.Count -eq 0)
	{
		$result = @{ Failed = (New-Object System.Collections.Generic.List[string]); Cancelled = $false; Done = 0 }
		if ($OnComplete)
		{
			try { & $OnComplete $result }
			catch { Write-Verbose "ATLAS: empty download callback failed: $($_.Exception.Message)" }
		}
		return $true
	}

	$state = $script:AtlasDownloadState
	if ($null -eq $state -or $null -eq $state.Form -or $state.Form.IsDisposed)
	{
		try { $state = New-AtlasDownloadState }
		catch
		{
			Write-Verbose "ATLAS: the download window could not be opened: $($_.Exception.Message)"
			$failed = New-Object System.Collections.Generic.List[string]
			foreach ($it in $Items) { $failed.Add([string]$it.Path) }
			if ($OnComplete)
			{
				try { & $OnComplete @{ Failed = $failed; Cancelled = $false; Done = 0 } }
				catch { }
			}
			return $false
		}
	}

	$job = @{
		Title = $Title
		Kind = [string]$Kind
		Items = @($Items)
		OnComplete = $OnComplete
		Failed = (New-Object System.Collections.Generic.List[string])
		Pending = (New-Object System.Collections.Generic.List[object])
		Active = (New-Object System.Collections.ArrayList)
		Next = 0
		Done = 0
		BytesDone = [long]0
		TotalBytes = [long]0
		LastName = ''
		LastUi = [DateTime]::MinValue
		DirsMade = @{}
		Cancelled = $false
		Finished = $false
		Sequential = ($null -eq $state.Client)
	}
	foreach ($it in $job.Items)
	{
		$job.Pending.Add($it)
		$job.TotalBytes += [long]$it.Size
	}
	$state.Queue.Add($job)
	if ($null -eq $state.Job) { Start-AtlasDownloadJob }
	return $true
}

function Test-AtlasDownloadActive
{
	param([string]$Kind = '')
	$state = $script:AtlasDownloadState
	if ($null -eq $state) { return $false }
	if ($null -ne $state.Job)
	{
		if (-not $Kind -or [string]$state.Job.Kind -eq $Kind) { return $true }
	}
	try
	{
		foreach ($job in $state.Queue)
		{
			if (-not $Kind -or [string]$job.Kind -eq $Kind) { return $true }
		}
	}
	catch { }
	return $false
}

function Start-AtlasDownloadJob
{
	$state = $script:AtlasDownloadState
	if ($null -eq $state -or $null -ne $state.Job) { return }
	if ($state.Queue.Count -eq 0)
	{
		Stop-AtlasDownloadState
		return
	}

	$job = $state.Queue[0]
	$state.Queue.RemoveAt(0)
	$state.Job = $job
	$job.LastUi = [DateTime]::MinValue
	try
	{
		$state.Form.Text = [string]$job.Title
		$state.Status.Text = 'Preparing...'
		$state.Bar.Maximum = [Math]::Max(1, $job.Items.Count)
		$state.Bar.Value = 0
		$state.Bytes.Text = ''
		$state.CancelButton.Enabled = $true
	}
	catch { }
}

function Stop-AtlasDownloadState
{
	$state = $script:AtlasDownloadState
	if ($null -eq $state) { return }
	$script:AtlasDownloadState = $null
	try { if ($state.Timer) { $state.Timer.Stop() } } catch { }
	try { if ($state.Timer) { $state.Timer.Dispose() } } catch { }
	try { if ($state.Client) { $state.Client.Dispose() } } catch { }
	try
	{
		if ($state.Form -and -not $state.Form.IsDisposed)
		{
			$state.Form.Close()
			$state.Form.Dispose()
		}
	}
	catch { }
}

function Complete-AtlasDownloadJob
{
	param($Job)
	$state = $script:AtlasDownloadState
	if ($null -eq $state -or $null -eq $Job -or $Job.Finished) { return }

	$Job.Finished = $true
	if ($null -ne $state.Job -and [object]::ReferenceEquals($state.Job, $Job)) { $state.Job = $null }
	$result = @{ Failed = $Job.Failed; Cancelled = [bool]$Job.Cancelled; Done = [int]$Job.Done }
	$hasMore = $state.Queue.Count -gt 0
	if (-not $hasMore) { Stop-AtlasDownloadState }

	if ($Job.OnComplete)
	{
		try { & $Job.OnComplete $result }
		catch { Write-Verbose "ATLAS: download completion callback failed: $($_.Exception.Message)" }
	}

	$live = $script:AtlasDownloadState
	if ($null -ne $live -and $null -eq $live.Job -and $live.Queue.Count -gt 0)
	{
		Start-AtlasDownloadJob
	}
}

function Cancel-AtlasDownloads
{
	$state = $script:AtlasDownloadState
	if ($null -eq $state) { return }

	$jobs = New-Object System.Collections.Generic.List[object]
	if ($null -ne $state.Job) { $jobs.Add($state.Job) }
	foreach ($queued in $state.Queue) { $jobs.Add($queued) }
	$state.Job = $null
	$state.Queue.Clear()

	try { if ($state.Client) { $state.Client.CancelPendingRequests() } } catch { }
	Stop-AtlasDownloadState

	foreach ($job in $jobs)
	{
		if ($job.Finished) { continue }
		$job.Finished = $true
		$job.Cancelled = $true
		$result = @{ Failed = $job.Failed; Cancelled = $true; Done = [int]$job.Done }
		if ($job.OnComplete)
		{
			try { & $job.OnComplete $result }
			catch { Write-Verbose "ATLAS: download cancel callback failed: $($_.Exception.Message)" }
		}
	}
}

function Update-AtlasDownload
{
	$state = $script:AtlasDownloadState
	if ($null -eq $state) { return }
	if ($null -eq $state.Form -or $state.Form.IsDisposed)
	{
		Stop-AtlasDownloadState
		return
	}
	if ($state.CancelAll)
	{
		Cancel-AtlasDownloads
		return
	}

	$job = $state.Job
	if ($null -eq $job)
	{
		Start-AtlasDownloadJob
		return
	}

	if ($job.Next -ge $job.Pending.Count -and $job.Active.Count -eq 0)
	{
		Complete-AtlasDownloadJob -Job $job
		return
	}

	if ($job.Sequential)
	{
		if ($job.Next -lt $job.Pending.Count)
		{
			$it = $job.Pending[$job.Next]
			$job.Next++
			$job.LastName = Split-Path -Leaf ([string]$it.Path)
			try
			{
				Invoke-AtlasDownloadFile -Url ([string]$it.Url) -Path ([string]$it.Path) -Hash ([string]$it.Hash) | Out-Null
				$job.Done++
				$job.BytesDone += [long]$it.Size
			}
			catch { $job.Failed.Add([string]$it.Path) }
		}
	}
	else
	{
		if ($null -eq $state.Client)
		{
			$state.Client = New-AtlasDownloadClient -Threads $state.Threads
			if ($null -eq $state.Client) { $job.Sequential = $true; return }
		}

		while ($job.Active.Count -lt $state.Threads -and $job.Next -lt $job.Pending.Count)
		{
			$it = $job.Pending[$job.Next]
			$job.Next++
			$entry = @{ Item = $it; Name = (Split-Path -Leaf ([string]$it.Path)); Task = $null; Error = '' }
			try { $entry.Task = $state.Client.GetByteArrayAsync([string]$it.Url) }
			catch { $entry.Error = $_.Exception.Message }
			[void]$job.Active.Add($entry)
		}

		for ($i = $job.Active.Count - 1; $i -ge 0; $i--)
		{
			$entry = $job.Active[$i]
			$path = [string]$entry.Item.Path
			if ($entry.Error)
			{
				$job.Failed.Add($path)
				$job.Active.RemoveAt($i)
				continue
			}
			$task = $entry.Task
			if ($null -eq $task -or -not $task.IsCompleted) { continue }

			$job.Active.RemoveAt($i)
			try
			{
				if ($task.IsFaulted) { throw $task.Exception.GetBaseException() }
				if ($task.IsCanceled) { throw 'The download was cancelled.' }
				$bytes = [byte[]]$task.Result
				$want = [string]$entry.Item.Hash
				if ($want)
				{
					$got = Get-AtlasBytesHash -Bytes $bytes
					if (-not $got -or $got -ne $want) { throw 'The downloaded file did not match its published checksum.' }
				}
				$dir = [System.IO.Path]::GetDirectoryName($path)
				if ($dir -and -not $job.DirsMade.ContainsKey($dir))
				{
					[void][System.IO.Directory]::CreateDirectory($dir)
					$job.DirsMade[$dir] = $true
				}
				[System.IO.File]::WriteAllBytes($path, $bytes)
				$job.Done++
				$job.BytesDone += [long]$entry.Item.Size
				$job.LastName = [string]$entry.Name
			}
			catch
			{
				$job.Failed.Add($path)
			}
		}
	}

	if ($job.Next -ge $job.Pending.Count -and $job.Active.Count -eq 0)
	{
		Complete-AtlasDownloadJob -Job $job
		return
	}

	$now = [DateTime]::UtcNow
	if (($now - $job.LastUi).TotalMilliseconds -ge 80)
	{
		try
		{
			$state.Bar.Value = [Math]::Min([int]$state.Bar.Maximum, [int]$job.Done)
			$state.Bytes.Text = ('{0} of {1}' -f (Format-AtlasSize $job.BytesDone), (Format-AtlasSize $job.TotalBytes))
			if ($job.LastName)
			{
				$state.Status.Text = ('Downloading {0}   ({1} of {2})' -f $job.LastName, [Math]::Min($job.Done + 1, $job.Items.Count), $job.Items.Count)
			}
			else
			{
				$state.Status.Text = ('Downloading {0} files...' -f $job.Items.Count)
			}
		}
		catch { }
		$job.LastUi = $now
	}
}

function Get-AtlasIconsState
{
	param([string]$Root)
	$p = Join-Path $Root 'icons\.atlas-icons.json'
	if (-not (Test-Path -LiteralPath $p)) { return $null }
	try { return (Get-Content -LiteralPath $p -Raw | ConvertFrom-Json) } catch { return $null }
}

function Save-AtlasIconsState
{
	param([string]$Root, [string]$Version)
	try
	{
		$dir = Join-Path $Root 'icons'
		if (-not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
		$txt = ([ordered]@{ version = [string]$Version; checked = (Get-Date).ToString('s') } | ConvertTo-Json -Compress)
		Set-Content -LiteralPath (Join-Path $dir '.atlas-icons.json') -Value $txt -Encoding UTF8
	}
	catch { }
}

function Get-AtlasIconDiff
{
	param($Manifest, [string]$Root)
	$iconsRoot = Join-Path $Root 'icons'
	$prefix = $iconsRoot + [System.IO.Path]::DirectorySeparatorChar
	$out = @{
		Missing = 0
		Changed = 0
		Total = 0
		Bytes = [long]0
		Files = (New-Object System.Collections.Generic.List[object])
	}
	$sha = $null
	try { $sha = [System.Security.Cryptography.SHA256]::Create() } catch { $sha = $null }
	foreach ($f in @($Manifest.icons.files))
	{
		if (-not $f -or -not $f.p) { continue }
		$out.Total++
		$rel = [string]$f.p
		$full = $prefix + $rel.Replace('/', [System.IO.Path]::DirectorySeparatorChar.ToString())
		$size = [long]0
		if ($null -ne $f.s) { $size = [long]$f.s }
		$want = [string]$f.h
		$exists = $false
		try { $exists = [System.IO.File]::Exists($full) } catch { $exists = $false }
		if (-not $exists)
		{
			$out.Missing++
			$out.Bytes += $size
			$out.Files.Add(@{ Rel = $rel; Path = $full; Size = $size; Hash = $want })
			continue
		}
		if (-not $want) { continue }

		$got = ''
		$sizeMismatch = $false
		if ($null -ne $sha)
		{
			$fs = $null
			try
			{
				$fs = [System.IO.File]::Open($full, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
				if ($size -gt 0 -and $fs.Length -ne $size) { $sizeMismatch = $true }
				else
				{
					$bytes = $sha.ComputeHash($fs)
					$got = [System.BitConverter]::ToString($bytes).Replace('-', '').ToLowerInvariant()
				}
			}
			catch { $got = '' }
			finally { if ($fs) { $fs.Dispose() } }
		}
		else
		{
			if ($size -gt 0)
			{
				$actual = [long]-1
				try { $actual = ([System.IO.FileInfo]::new($full)).Length } catch { $actual = -1 }
				if ($actual -ge 0 -and $actual -ne $size) { $sizeMismatch = $true }
			}
			if (-not $sizeMismatch) { $got = Get-AtlasFileHash $full }
		}

		if ($sizeMismatch -or $got -ne $want)
		{
			$out.Changed++
			$out.Bytes += $size
			$out.Files.Add(@{ Rel = $rel; Path = $full; Size = $size; Hash = $want })
		}
	}
	if ($null -ne $sha) { try { $sha.Dispose() } catch { } }
	return $out
}

function Sync-AtlasAssets
{
	$root = Get-AtlasInstallRoot
	if (-not $root) { return $false }
	$manifest = Get-AtlasManifest
	if (-not $manifest) { return $false }
	$base = Get-AtlasRemoteBase
	$changed = $false

	$dataRel = 'src/data.js'
	if ($manifest.data.path) { $dataRel = [string]$manifest.data.path }
	$dataUrl = $base + '/' + $dataRel
	$dataPath = Get-AtlasDataPath
	if (-not $dataPath) { $dataPath = Join-Path $root ($dataRel -replace '/', '\') }
	$dataSize = [long]0
	if ($manifest.data.size) { $dataSize = [long]$manifest.data.size }
	$dataWant = [string]$manifest.data.sha256
	$dataDate = [string]$manifest.data.date
	$dataBusy = Test-AtlasDownloadActive -Kind 'AtlasData'

	if (-not (Test-Path -LiteralPath $dataPath))
	{
		$msg = "The Entropia Map data is not installed yet.`n`nDownload it now? ($(Format-AtlasSize $dataSize))"
		if (-not $dataBusy -and (Read-AtlasUpdateAnswer -Text $msg -Caption 'Entropia Map - Map data'))
		{
			Start-AtlasDownload -Title 'Entropia Map - map data' -Kind 'AtlasData' -Items @(@{ Url = $dataUrl; Path = $dataPath; Size = $dataSize; Hash = $dataWant }) -OnComplete {
				param($r)
				if ($r.Done -gt 0)
				{
					Reset-AtlasCaches
					if ($script:AtlasOpenPending)
					{
						$script:AtlasOpenPending = $false
						try { Show-Atlas -SkipAssetSync | Out-Null } catch { }
					}
				}
				elseif ($r.Failed.Count -gt 0)
				{
					$script:AtlasOpenPending = $false
					Show-AtlasNotice -Text 'The map data could not be downloaded. Check your connection and open the map again.' -Caption 'Entropia Map - Download failed' -Icon 'Error'
				}
			} | Out-Null
		}
	}
	elseif ($dataWant)
	{
		$got = Get-AtlasFileHash $dataPath
		$offer = [bool]($got -and $got -ne $dataWant)
		if ($offer)
		{
			# A manifest older than the data on disk describes an older build.
			# Do not offer to replace newer local data with it: the data that
			# ships beside this module can be newer than the published manifest.
			$genUtc = [datetime]::MinValue
			$genRaw = $manifest.generated
			try
			{
				if ($genRaw -is [datetime]) { $genUtc = ([datetime]$genRaw).ToUniversalTime() }
				elseif ($genRaw) { [void][datetime]::TryParse([string]$genRaw, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$genUtc) }
			}
			catch { }
			if ($genUtc -gt [datetime]::MinValue)
			{
				try { if ((Get-Item -LiteralPath $dataPath).LastWriteTimeUtc -gt $genUtc) { $offer = $false } } catch { }
			}
		}
		if ($offer)
		{
			$when = ''
			if ($dataDate) { $when = ' (updated ' + $dataDate + ')' }
			$msg = "A newer version of the map data is available$($when).`n`nDownload it now? ($(Format-AtlasSize $dataSize))"
			if (-not $dataBusy -and (Read-AtlasUpdateAnswer -Text $msg -Caption 'Entropia Map - Update'))
			{
				Start-AtlasDownload -Title 'Entropia Map - updating map data' -Kind 'AtlasData' -Items @(@{ Url = $dataUrl; Path = $dataPath; Size = $dataSize; Hash = $dataWant }) -OnComplete {
					param($r)
					if ($r.Done -gt 0)
					{
						Reset-AtlasCaches
						if ($script:AtlasOpenPending)
						{
							$script:AtlasOpenPending = $false
							try { Show-Atlas -SkipAssetSync | Out-Null } catch { }
						}
					}
					elseif ($r.Failed.Count -gt 0)
					{
						$script:AtlasOpenPending = $false
						Show-AtlasNotice -Text 'The new map data could not be downloaded. The map will keep using the copy it has.' -Caption 'Entropia Map - Download failed' -Icon 'Error'
					}
				} | Out-Null
			}
		}
	}

	if (Sync-AtlasIcons) { $changed = $true }

	return $changed
}

function Sync-AtlasIcons
{
	[CmdletBinding()]
	[OutputType([bool])]
	param([scriptblock]$OnComplete = $null)

	$root = Get-AtlasInstallRoot
	if (-not $root) { return $false }
	$manifest = Get-AtlasManifest
	if (-not $manifest -or -not $manifest.icons -or -not $manifest.icons.files) { return $false }

	$base = Get-AtlasRemoteBase
	$changed = $false
	$iconFiles = @($manifest.icons.files)
	if ($iconFiles.Count -eq 0) { return $false }

	$iconsRoot = Join-Path $root 'icons'
	$remoteVer = ''
	if ($manifest.icons.version) { $remoteVer = [string]$manifest.icons.version }
	$state = Get-AtlasIconsState -Root $root
	$haveVer = ''
	if ($state -and $state.version) { $haveVer = [string]$state.version }
	$scan = $true
	if ($haveVer -and $remoteVer -and $haveVer -eq $remoteVer -and (Test-Path -LiteralPath $iconsRoot)) { $scan = $false }
	if (-not $scan) { return $false }
	if (Test-AtlasDownloadActive -Kind 'AtlasIcons') { return $false }

	$diff = Get-AtlasIconDiff -Manifest $manifest -Root $root
	if ($diff.Files.Count -eq 0)
	{
		Save-AtlasIconsState -Root $root -Version $remoteVer
		return $false
	}

	if ($diff.Missing -eq $diff.Files.Count)
	{
		$what = 'The atlas icons are not installed.'
	}
	else
	{
		$bits = @()
		if ($diff.Missing -gt 0) { $bits += ('{0} missing' -f $diff.Missing) }
		if ($diff.Changed -gt 0) { $bits += ('{0} changed' -f $diff.Changed) }
		$what = 'Atlas icons: ' + ($bits -join ', ') + '.'
	}
	$count = $diff.Files.Count
	$word = if ($count -eq 1) { 'icon' } else { 'icons' }
	$msg = "$what`n`nDownload $count ${word}? ($(Format-AtlasSize $diff.Bytes))`n`nThe map and market work without them - pictures are then left out."
	if (Read-AtlasUpdateAnswer -Text $msg -Caption 'Entropia Atlas - Icons')
	{
		$items = New-Object System.Collections.Generic.List[object]
		foreach ($f in $diff.Files)
		{
			$items.Add(@{ Url = ($base + '/icons/' + $f.Rel); Path = $f.Path; Size = $f.Size; Hash = [string]$f.Hash })
		}
		$rootForDone = $root
		$versionForDone = $remoteVer
		$after = $OnComplete
		$done = {
			param($r)
			if ($r.Failed.Count -eq 0 -and -not $r.Cancelled)
			{
				Save-AtlasIconsState -Root $rootForDone -Version $versionForDone
				Reset-AtlasCaches
				if ($after) { try { & $after $r } catch { } }
			}
			elseif ($r.Failed.Count -gt 0)
			{
				$fw = if ($r.Failed.Count -eq 1) { 'icon' } else { 'icons' }
				Show-AtlasNotice -Text ('{0} {1} could not be downloaded. They will be offered again the next time the atlas is opened.' -f $r.Failed.Count, $fw) -Caption 'Entropia Atlas - Download failed' -Icon 'Error'
			}
		}.GetNewClosure()
		Start-AtlasDownload -Title 'Entropia Atlas - downloading icons' -Kind 'AtlasIcons' -Items $items.ToArray() -OnComplete $done | Out-Null
	}
	return $changed
}

function Reset-AtlasCaches
{
	$script:AtlasData = $null
	try { if ('Atlas.Cache' -as [Type]) { [Atlas.Cache]::Clear() } } catch { }
}

function Get-AtlasPalette
{
	$p = @{}
	$get = { param($k, $d) try { $c = Get-Theme $k -ErrorAction SilentlyContinue; if ($c) { return $c } } catch { } return $d }
	$p.Bg        = & $get 'MainBg' ([System.Drawing.Color]::FromArgb(17, 24, 35))
	$p.Panel     = & $get 'SecondaryBg' ([System.Drawing.Color]::FromArgb(17, 24, 35))
	$p.Panel2    = & $get 'TertiaryBg' ([System.Drawing.Color]::FromArgb(24, 33, 46))
	$p.Input     = & $get 'InputBg' ([System.Drawing.Color]::FromArgb(24, 33, 46))
	$p.Line      = & $get 'BorderColor' ([System.Drawing.Color]::FromArgb(38, 51, 69))
	$p.Fg        = & $get 'WhiteText' ([System.Drawing.Color]::FromArgb(231, 238, 247))
	$p.Mut       = & $get 'HeaderSubtitleText' ([System.Drawing.Color]::FromArgb(132, 150, 172))
	$p.Dim       = & $get 'SecondaryText' ([System.Drawing.Color]::FromArgb(93, 110, 131))
	$p.Acc       = & $get 'AccentBlue' ([System.Drawing.Color]::FromArgb(90, 176, 255))
	return $p
}

function New-AtlasCombo
{
	param([int]$Width = 0)
	if ('Custom.DarkComboBox' -as [Type]) { $c = New-Object Custom.DarkComboBox }
	else
	{
		$c = New-Object System.Windows.Forms.ComboBox
		$c.FlatStyle = 'Flat'
	}
	$c.DropDownStyle = 'DropDownList'
	$c.Font = Get-UIFont 'Segoe UI' 8.5
	if ($Width -gt 0) { $c.Width = $Width }
	if (Test-DashboardCommand 'ApplyControlInternal')
	{
		try { ApplyControlInternal $c (Get-Theme) } catch { }
	}
	if ($c.PSObject.Properties['HeaderColor']) { try { $c.HeaderColor = (Get-AtlasPalette).Dim } catch { } }
	return $c
}

function Get-AtlasNodeToolTip
{
	if ($script:AtlasNodeToolTip -and -not $script:AtlasNodeToolTip.IsDisposed) { return $script:AtlasNodeToolTip }
	$tip = New-Object System.Windows.Forms.ToolTip
	$tip.OwnerDraw = $true
	$tip.InitialDelay = 350
	$tip.ReshowDelay = 120
	$tip.AutoPopDelay = 9000
	$tip | Add-Member -MemberType NoteProperty -Name 'Hit' -Value $null -Force
	$tip | Add-Member -MemberType NoteProperty -Name 'Size' -Value (New-Object System.Drawing.Size(20, 20)) -Force
	$tip.Add_Popup({
			param($sender, $e)
			if ($sender.Hit) { $e.ToolTipSize = $sender.Size }
			else { $e.ToolTipSize = New-Object System.Drawing.Size(1, 1) }
		}.GetNewClosure())
	$tip.Add_Draw({
			param($sender, $e)
			DrawAtlasNodeTip -G $e.Graphics -Hit $sender.Hit -Bounds $e.Bounds
		}.GetNewClosure())
	$script:AtlasNodeToolTip = $tip
	return $tip
}

function Get-AtlasNodeTipLayout
{
	param($Hit)
	Initialize-AtlasCore | Out-Null
	$l = @{}
	$l.Name = $Hit.Label
	$l.Lvl = ''
	if ($Hit.Kind -eq 'mob' -and $Hit.Lvl) { $l.Lvl = [Atlas.State]::LvlText($Hit.Lvl) }
	$l.Alt = ''
	if ($Hit.Alt) { $l.Alt = ([string][char]0x00B7) + ' ' + $Hit.Alt }
	$l.Coords = ('{0:0}, {1:0}' -f $Hit.WX, $Hit.WZ)
	$l.Boss = [bool]$Hit.Boss
	return $l
}

function Get-AtlasNodeTipParchment
{
	param($Hit)
	if (-not $Hit -or $Hit.Kind -ne 'treasure' -or -not $Hit.Full) { return $null }
	$root = $script:AtlasData
	if (-not $root) { return $null }
	$path = Join-Path $root.AssetRoot ($Hit.Full -replace '/', '\')
	if (-not (Test-Path -LiteralPath $path)) { return $null }
	try
	{
		$fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
		$src = [System.Drawing.Image]::FromStream($fs)
		$bmp = New-Object System.Drawing.Bitmap($src)
		$src.Dispose(); $fs.Dispose()
		return $bmp
	}
	catch { return $null }
}

function Get-AtlasNodeTipSize
{
	param($Hit)
	$l = Get-AtlasNodeTipLayout -Hit $Hit
	$fName = Get-UIFont 'Segoe UI' 9 'Bold'
	$fSmall = Get-UIFont 'Segoe UI' 8
	$fItal = Get-UIFont 'Segoe UI' 8.5 'Italic'
	$fMono = Get-UIFont 'Consolas' 8
	try
	{
		$bm = New-Object System.Drawing.Bitmap(1, 1)
		$g = [System.Drawing.Graphics]::FromImage($bm)
		$w = 8 + 14 + 6
		$w += $g.MeasureString($l.Name, $fName).Width
		if ($l.Lvl) { $w += $g.MeasureString($l.Lvl, $fSmall).Width + 6 }
		if ($l.Alt) { $w += $g.MeasureString($l.Alt, $fItal).Width + 6 }
		if ($l.Boss) { $w += 34 }
		$w += 10 + $g.MeasureString($l.Coords, $fMono).Width + 8
		$g.Dispose(); $bm.Dispose()
		$h = 24
		$parch = Get-AtlasNodeTipParchment -Hit $Hit
		if ($parch)
		{
			$ph = [Math]::Min(96, $parch.Height)
			if ($ph -gt $h) { $h = $ph + 8 }
			$pw = [int]($parch.Width * ($ph / [double]$parch.Height))
			$w += $pw + 6
			$parch.Dispose()
		}
		return New-Object System.Drawing.Size([int]$w, [int]$h)
	}
	finally { }
}

function DrawAtlasNodeTip
{
	param([System.Drawing.Graphics]$G, $Hit, [System.Drawing.Rectangle]$Bounds)
	if (-not $Hit) { return }
	$pal = Get-AtlasPalette
	$l = Get-AtlasNodeTipLayout -Hit $Hit
	$fName = Get-UIFont 'Segoe UI' 9 'Bold'
	$fSmall = Get-UIFont 'Segoe UI' 8
	$fItal = Get-UIFont 'Segoe UI' 8.5 'Italic'
	$fMono = Get-UIFont 'Consolas' 8
	$parch = $null
	try
	{
		$G.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
		$G.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
		$bg = New-Object System.Drawing.SolidBrush($pal.Panel2)
		$G.FillRectangle($bg, $Bounds.X, $Bounds.Y, $Bounds.Width, $Bounds.Height)
		$bg.Dispose()
		$pen = New-Object System.Drawing.Pen($pal.Line)
		$G.DrawRectangle($pen, $Bounds.X, $Bounds.Y, $Bounds.Width - 1, $Bounds.Height - 1)
		$pen.Dispose()

		$x = $Bounds.X + 8
		$cy = $Bounds.Y + 12

		$parch = Get-AtlasNodeTipParchment -Hit $Hit
		if ($parch)
		{
			$ph = [Math]::Min(96, $parch.Height)
			$pw = [int]($parch.Width * ($ph / [double]$parch.Height))
			$G.DrawImage($parch, $x, ($Bounds.Y + 4), $pw, $ph)
			$x += $pw + 6
		}

		$col = $Hit.Color
		if (-not $col -or $col.IsEmpty) { $col = $pal.Acc }
		$mk = New-Object System.Drawing.SolidBrush($col)
		if ($Hit.Kind -eq 'species' -or $Hit.Kind -eq 'mob')
		{
			$G.FillEllipse($mk, $x, ($cy - 5), 10, 10)
		}
		else
		{
			$pts = New-Object 'System.Drawing.Point[]' 4
			$pts[0] = New-Object System.Drawing.Point(($x + 5), ($cy - 6))
			$pts[1] = New-Object System.Drawing.Point(($x + 11), $cy)
			$pts[2] = New-Object System.Drawing.Point(($x + 5), ($cy + 6))
			$pts[3] = New-Object System.Drawing.Point(($x - 1), $cy)
			$G.FillPolygon($mk, $pts)
		}
		$mk.Dispose()
		$x += 16

		$fb = New-Object System.Drawing.SolidBrush($pal.Fg)
		$mb = New-Object System.Drawing.SolidBrush($pal.Mut)
		$db = New-Object System.Drawing.SolidBrush($pal.Dim)
		$G.DrawString($l.Name, $fName, $fb, $x, ($cy - 8))
		$x += [int]$G.MeasureString($l.Name, $fName).Width
		if ($l.Lvl)
		{
			$G.DrawString($l.Lvl, $fSmall, $mb, ($x + 4), ($cy - 6))
			$x += [int]$G.MeasureString($l.Lvl, $fSmall).Width + 6
		}
		if ($l.Alt)
		{
			$G.DrawString($l.Alt, $fItal, $mb, ($x + 2), ($cy - 6))
			$x += [int]$G.MeasureString($l.Alt, $fItal).Width + 6
		}
		if ($l.Boss)
		{
			$tb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 110, 168))
			$G.DrawString('boss', $fSmall, $tb, ($x + 2), ($cy - 6))
			$tb.Dispose()
			$x += 36
		}
		$co = '{0:0}, {1:0}' -f $Hit.WX, $Hit.WZ
		$cw = [int]$G.MeasureString($co, $fMono).Width
		$G.DrawString($co, $fMono, $db, ($Bounds.X + $Bounds.Width - $cw - 8), ($cy - 7))
		$fb.Dispose(); $mb.Dispose(); $db.Dispose()
	}
	finally
	{
		if ($parch) { $parch.Dispose() }
	}
}

function Get-AtlasToolTip
{
	if ($script:AtlasToolTip -and -not $script:AtlasToolTip.IsDisposed) { return $script:AtlasToolTip }
	$tip = $null
	if (Test-DashboardCommand 'New-ThemedToolTip') { try { $tip = New-ThemedToolTip } catch { $tip = $null } }
	if (-not $tip) { $tip = New-Object System.Windows.Forms.ToolTip }
	$tip.AutoPopDelay = 9000
	$script:AtlasToolTip = $tip
	return $tip
}

function New-AtlasButton
{
	param(
		[string]$Text, [int]$Width = 0, [string]$Tip = '', [switch]$Primary,
		[System.Drawing.Color]$Accent = ([System.Drawing.Color]::Empty)
	)
	$pal = Get-AtlasPalette
	if ($Accent -eq [System.Drawing.Color]::Empty) { $Accent = $pal.Acc }
	$b = New-Object System.Windows.Forms.Button
	$b.Text = $Text
	$b.FlatStyle = 'Flat'
	$b.ForeColor = $pal.Mut
	$b.BackColor = $pal.Panel2
	$b.FlatAppearance.BorderColor = $pal.Line
	$b.FlatAppearance.BorderSize = 1
	$b.Cursor = [System.Windows.Forms.Cursors]::Hand
	$b.Font = Get-UIFont 'Segoe UI' 8.5
	$b.AutoSize = $false
	$b.Height = 24
	$b.Margin = New-Object System.Windows.Forms.Padding(0, 0, 4, 0)
	if ($Width -gt 0) { $b.Width = $Width }
	if ($Primary) { $b.ForeColor = $pal.Acc; $b.Font = Get-UIFont 'Segoe UI' 8.5 'Bold' }
	if ($PSBoundParameters.ContainsKey('Accent')) {
		$b | Add-Member -MemberType NoteProperty -Name 'AtlasAccent' -Value $Accent -Force
	}
	if ($Tip)
	{
		$tt = Get-AtlasToolTip
		if ($tt) { try { $tt.SetToolTip($b, $Tip) } catch { } }
	}
	if (Test-DashboardCommand 'Register-ThemedButtonFadeEffect') {
		try { Register-ThemedButtonFadeEffect -Button $b } catch { }
	}
	return $b
}

function Set-AtlasTip
{
	param([System.Windows.Forms.Control]$Control, [string]$Text)
	if (-not $Control -or -not $Text) { return }
	$tt = Get-AtlasToolTip
	if (-not $tt) { return }
	try { $tt.SetToolTip($Control, $Text) } catch { }
}

function Set-AtlasButtonState
{
	param([System.Windows.Forms.Button]$Button, [bool]$On, [System.Drawing.Color]$Accent)
	if (-not $Button) { return }
	$pal = Get-AtlasPalette
	if ($On) {
		$Button.ForeColor = $Accent
		$Button.BackColor = [System.Drawing.Color]::FromArgb(38, $Accent.R, $Accent.G, $Accent.B)
		$Button.FlatAppearance.BorderColor = $Accent
	}
	else {
		$Button.ForeColor = $pal.Mut
		$Button.BackColor = $pal.Panel2
		$Button.FlatAppearance.BorderColor = $pal.Line
	}
}

#endregion

#region PowerShell - window chrome

function Add-AtlasResizeGrips
{
	param([System.Windows.Forms.Form]$Form)
	$specs = @(
		@{ dock = 'Top';    size = 6; hit = 12; cur = 'SizeNS' }
		@{ dock = 'Bottom'; size = 6; hit = 15; cur = 'SizeNS' }
		@{ dock = 'Left';   size = 6; hit = 10; cur = 'SizeWE' }
		@{ dock = 'Right';  size = 6; hit = 11; cur = 'SizeWE' }
	)
	$cornerZone = 18
	$curNS = [System.Windows.Forms.Cursors]::SizeNS
	$curWE = [System.Windows.Forms.Cursors]::SizeWE
	$curNWSE = [System.Windows.Forms.Cursors]::SizeNWSE
	$curNESW = [System.Windows.Forms.Cursors]::SizeNESW
	$curMap = @{
		'SizeNS'  = $curNS
		'SizeWE'  = $curWE
		'SizeNWSE' = $curNWSE
		'SizeNESW' = $curNESW
	}
	$outp = New-Object System.Collections.Generic.List[System.Windows.Forms.Panel]
	foreach ($sp in $specs)
	{
		$hitCode = $sp.hit
		$g = New-Object System.Windows.Forms.Panel
		$g.Name = 'AtlasGrip' + $sp.dock
		$g.BackColor = $Form.BackColor
		$g.TabStop = $false
		$g.Dock = $sp.dock
		if ($sp.dock -eq 'Top' -or $sp.dock -eq 'Bottom') { $g.Height = $sp.size }
		else { $g.Width = $sp.size }
		$g.Cursor = $curMap[$sp.cur]
		$g.Add_MouseMove({
				param($s, $e)
				$nearTop = $e.Y -le $cornerZone
				$nearBottom = ($s.Height - $e.Y) -le $cornerZone
				$nearLeft = $e.X -le $cornerZone
				$nearRight = ($s.Width - $e.X) -le $cornerZone
				switch ($hitCode)
				{
					12 { if ($nearLeft) { $s.Cursor = $curNWSE } elseif ($nearRight) { $s.Cursor = $curNESW } else { $s.Cursor = $curNS } }
					15 { if ($nearLeft) { $s.Cursor = $curNESW } elseif ($nearRight) { $s.Cursor = $curNWSE } else { $s.Cursor = $curNS } }
					10 { if ($nearTop) { $s.Cursor = $curNWSE } elseif ($nearBottom) { $s.Cursor = $curNESW } else { $s.Cursor = $curWE } }
					11 { if ($nearTop) { $s.Cursor = $curNESW } elseif ($nearBottom) { $s.Cursor = $curNWSE } else { $s.Cursor = $curWE } }
				}
			}.GetNewClosure())
		$g.Add_MouseDown({
				param($s, $e)
				if ($e.Button -ne 'Left' -or -not ('Custom.Native' -as [Type])) { return }
				$hit = $hitCode
				$nearTop = $e.Y -le $cornerZone
				$nearBottom = ($s.Height - $e.Y) -le $cornerZone
				$nearLeft = $e.X -le $cornerZone
				$nearRight = ($s.Width - $e.X) -le $cornerZone
				switch ($hitCode)
				{
					12 { if ($nearLeft) { $hit = 13 } elseif ($nearRight) { $hit = 14 } }
					15 { if ($nearLeft) { $hit = 16 } elseif ($nearRight) { $hit = 17 } }
					10 { if ($nearTop) { $hit = 13 } elseif ($nearBottom) { $hit = 16 } }
					11 { if ($nearTop) { $hit = 14 } elseif ($nearBottom) { $hit = 17 } }
				}
				[Custom.Native]::ReleaseCapture()
				[Custom.Native]::SendMessage($s.FindForm().Handle, 0xA1, $hit, [IntPtr]::Zero)
			}.GetNewClosure())
		$outp.Add($g)
	}
	return $outp
}

function Add-AtlasDrag
{
	param([System.Windows.Forms.Control]$Control)
	$Control.Add_MouseDown({
			param($s, $e)
			if ($e.Button -eq 'Left' -and ('Custom.Native' -as [Type]))
			{
				[Custom.Native]::ReleaseCapture()
				[Custom.Native]::SendMessage($s.FindForm().Handle, 0xA1, 0x2, [IntPtr]::Zero)
			}
		}.GetNewClosure())
}

#endregion
#region PowerShell - the window

function Get-AtlasAccent
{
	param([System.Windows.Forms.Button]$Button, [string]$Default = '')
	$h = @{ 'Statues' = '#ffffff'; 'Ore' = '#ffb020'; 'Plant' = '#3fd16a'; 'Omega' = '#ffb64a'
		'Astral' = '#78dcff'; 'Treasure' = '#d4953a'; 'Daily Quest' = '#ff6ea8' }
	if ($Button)
	{
		try { if ($Button.PSObject.Properties['AtlasAccent'] -and $Button.AtlasAccent) { return $Button.AtlasAccent } } catch { }
	}
	if ($Default -and $h.ContainsKey($Default)) { return $h[$Default] }
	if (-not $Button) { return $Default }
	foreach ($k in $h.Keys) { if ([string]$Button.Text -eq $k -or [string]$Button.Text.EndsWith(' ' + $k)) { return $h[$k] } }
	return $Default
}

function Update-AtlasStatus
{
	param([hashtable]$Ctx)
	$S = $Ctx.S; $canvas = $Ctx.canvas
	$c = $S.StatusCounts()
	$Ctx.stMap.Text = if ($S.Meta) { $S.Meta.Title } else { '-' }
	$word = if ($c[0] -eq 1) { 'spot' } else { 'spots' }
	$Ctx.stSpots.Text = ('{0:N0} {1} · {2} species' -f $c[0], $word, $c[1])
	if ($S.ShowSp)
	{
		$Ctx.stSpawn.Text = '{0:N0} monsters · {1:N0} items · {2:N0} maps · {3:N0} statues · {4:N0} npcs' -f $c[2], $c[3], $c[4], $c[5], $c[6]
	}
	else { $Ctx.stSpawn.Text = 'spawns hidden' }
	$Ctx.stScale.Text = ('{0:0.00} px/m' -f ($canvas.Sc * $S.D.K * $canvas.Ppx))
}

function Sync-AtlasToggles
{
	param([hashtable]$Ctx)
	$S = $Ctx.S
	foreach ($b in $Ctx.guideButtons)
	{
		$acc = Get-AtlasAccent -Button $b
		Set-AtlasButtonState -Button $b -On ($b.Tag -eq $S.Query) -Accent $acc
	}
	foreach ($b in $Ctx.mineButtons)
	{
		Set-AtlasButtonState -Button $b -On ($b.Tag -eq $S.Query) -Accent (Get-AtlasAccent -Button $b -Default '#ffb020')
	}
	Set-AtlasButtonState -Button $Ctx.btnSpawns -On $S.ShowSp -Accent '#5ab0ff'
	Set-AtlasButtonState -Button $Ctx.btnZones -On $S.ShowPl -Accent '#5ab0ff'
}

function Sync-AtlasJump
{
	param([hashtable]$Ctx)
	Initialize-AtlasCore | Out-Null
	$S = $Ctx.S; $cmb = $Ctx.cmbJump
	$visible = $S.VisibleWorlds($S.MapGroup)
	$idx = -1
	for ($i = 0; $i -lt $visible.Count; $i++) { if ($visible[$i].Key -eq $S.MapKey) { $idx = $i; break } }
	$nextIdx = if ($idx -lt 0) { 0 } else { $idx + 1 }
	$Ctx.btnPrev.Enabled = $idx -gt 0
	$Ctx.btnNext.Enabled = ($visible.Count -gt 0 -and $nextIdx -lt $visible.Count)
	if ($idx -gt 0) { Set-AtlasTip -Control $Ctx.btnPrev -Text ('Previous world: ' + $visible[$idx - 1].Title) }
	else { Set-AtlasTip -Control $Ctx.btnPrev -Text 'First world in this group' }
	if ($nextIdx -lt $visible.Count) { Set-AtlasTip -Control $Ctx.btnNext -Text ('Next world: ' + $visible[$nextIdx].Title) }
	else { Set-AtlasTip -Control $Ctx.btnNext -Text 'Last world in this group' }

	$lines = New-Object System.Collections.Generic.List[object]
	$addHead = { param([string]$Text) $lines.Add(@{ t = $Text; h = $true }) }
	$addWorld = { param($wm, [string]$Suffix) $lines.Add(@{ t = ('   ' + $wm.Title + $Suffix); h = $false }) }

	if ($S.MapGroup -eq 'farm')
	{
		$parts = @(
			@{ lab = 'Farm Maps - Penya'; keys = @('wdheaven03', 'duflmas', 'wdoasisdesert', 'dreadcove') }
			@{ lab = 'Farm Maps - ECP'; keys = @('wddarkrartesia', 'amethyst', 'rifloatingcity') }
		)
		foreach ($p in $parts)
		{
			$found = @($p.keys | ForEach-Object { $S.Find($_) } | Where-Object { $_ })
			if ($found.Count -eq 0) { continue }
			& $addHead $p.lab
			foreach ($wm in $found) { & $addWorld $wm (' [' + [Atlas.State]::FarmLabel($wm.Key) + ']') }
		}
	}
	elseif ($S.MapGroup -in @('platin', 'royal'))
	{
		$pl = ($S.MapGroup -eq 'platin')
		& $addHead ([Atlas.State]::GRP_TITLE[$S.MapGroup])
		foreach ($wm in $visible)
		{
			$t = [Atlas.State]::TierText($S.D.W($wm.Key), $pl)
			if (-not $t) { $t = 'RL' }
			& $addWorld $wm (' <' + $t + '>')
		}
	}
	elseif ($S.MapGroup -eq 'dungeons')
	{
		foreach ($t in $S.DungeonTiers())
		{
			& $addHead ([Atlas.State]::GRP_TITLE['dungeons'] + ' - ' + $t.Key + ' (' + $t.Value.Count + ')')
			foreach ($wm in $t.Value) { & $addWorld $wm '' }
		}
	}
	else
	{
		$groups = @('main', 'dungeons', 'fields', 'event')
		if ($S.MapGroup -ne 'all') { $groups = @($S.MapGroup) }
		foreach ($gname in $groups)
		{
			$list = $S.VisibleWorlds($gname)
			if ($list.Count -eq 0) { continue }
			if ($gname -eq 'dungeons')
			{
				foreach ($t in $S.DungeonTiers())
				{
					& $addHead ([Atlas.State]::GRP_TITLE[$gname] + ' - ' + $t.Key + ' (' + $t.Value.Count + ')')
					foreach ($wm in $t.Value) { & $addWorld $wm '' }
				}
			}
			else
			{
				& $addHead ([Atlas.State]::GRP_TITLE[$gname])
				foreach ($wm in $list) { & $addWorld $wm '' }
			}
		}
	}

	$heads = New-Object System.Collections.Generic.List[int]
	for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i].h) { $heads.Add($i) } }

	$hasSup = [bool]$cmb.PSObject.Properties['SuppressChanges']
	if ($hasSup) { $cmb.SuppressChanges = $true }
	try
	{
		$cmb.Items.Clear()
		foreach ($ln in $lines) { [void]$cmb.Items.Add($ln.t) }
		if ($cmb.GetType().GetMethod('SetHeaderIndices')) { $cmb.SetHeaderIndices($heads.ToArray()) }

		$sel = -1
		for ($i = 0; $i -lt $cmb.Items.Count; $i++)
		{
			$txt = ([string]$cmb.Items[$i]).Trim()
			if ($txt -eq $S.Meta.Title) { $sel = $i; break }
		}
		$cmb.SelectedIndex = $sel
	}
	finally
	{
		if ($hasSup) { $cmb.SuppressChanges = $false }
	}
	if ($cmb.GetType().GetMethod('FitDropDownWidth')) { try { $cmb.FitDropDownWidth(360) } catch { } }
}

function Format-AtlasShopText
{
	param($Shop)
	if (-not $Shop) { return '' }
	$lines = New-Object System.Collections.Generic.List[string]
	if ($Shop.Trade)
	{
		$head = 'TRADE'
		if ($Shop.Trade.Currency -and $Shop.Trade.Currency.Name) {
			$head += '  ' + [string][char]0x00B7 + '  ' + $Shop.Trade.Currency.Name
		}
		$lines.Add($head)
		$tabs = @($Shop.Trade.Tabs)
		$items = @($Shop.Trade.Items)
		$count = [Math]::Max(1, $tabs.Count)
		for ($ti = 0; $ti -lt $count; $ti++)
		{
			if ($tabs.Count -gt 1) { $lines.Add('  [' + $tabs[$ti] + ']') }
			for ($i = 0; $i -lt $items.Count; $i++)
			{
				if ([int]$items[$i].Tab -ne $ti) { continue }
				$q = ''
				if ([int]$items[$i].Price -ne 1) { $q = '  x' + [int]$items[$i].Price }
				$lines.Add('  ' + $items[$i].Name + $q)
			}
		}
	}
	foreach ($ex in @($Shop.Exchange))
	{
		if (-not $ex) { continue }
		if ($lines.Count -gt 0) { $lines.Add('') }
		$lines.Add('EXCHANGE  ' + [string][char]0x00B7 + '  ' + $ex.Label)
		foreach ($set in @($ex.Sets))
		{
			$give = @($set.Cost | ForEach-Object {
				$s = [string]$_.Name
				if ([int]$_.Qty -gt 1) { $s += ' x' + [int]$_.Qty }
				$s
			}) -join ', '
			$get = @($set.Get | ForEach-Object {
				$s = [string]$_.Name
				if ([int]$_.Qty -gt 1) { $s += ' x' + [int]$_.Qty }
				$s
			}) -join ', '
			$lines.Add('  Give: ' + $give)
			$lines.Add('  Get:  ' + $get)
		}
	}
	if ($lines.Count -eq 0) { $lines.Add('No stock recorded.') }
	return ($lines -join "`r`n")
}

function Set-AtlasCardRegion
{
	param($Panel, [bool]$Round = $true)
	if (-not $Panel) { return }
	if (-not $Round)
	{
		$old = $Panel.Region
		$Panel.Region = $null
		if ($old) { $old.Dispose() }
		return
	}
	$w = [int]$Panel.Width; $h = [int]$Panel.Height
	if ($w -lt 4 -or $h -lt 4) { return }
	$r = [Math]::Min(11, [Math]::Min($w, $h) / 2)
	$d = $r * 2
	$p = New-Object System.Drawing.Drawing2D.GraphicsPath
	$p.AddArc(0, 0, $d, $d, 180, 90)
	$p.AddArc(($w - $d), 0, $d, $d, 270, 90)
	$p.AddArc(($w - $d), ($h - $d), $d, $d, 0, 90)
	$p.AddArc(0, ($h - $d), $d, $d, 90, 90)
	$p.CloseFigure()
	$old = $Panel.Region
	$Panel.Region = New-Object System.Drawing.Region($p)
	if ($old) { $old.Dispose() }
	$p.Dispose()
}

function Update-AtlasDetail
{
	param([hashtable]$Ctx, [Atlas.Hit]$Hit)
	Initialize-AtlasCore | Out-Null
	$detail = $Ctx.detail
	$S = $Ctx.S; $pal = $Ctx.pal
	$detail.Visible = $false
	$detail.Tag = $null
	$Ctx.detailDrag = $null
	@($detail.Controls) | ForEach-Object { $_.Dispose() }
	if (-not $Hit) { return }

	$isShop = ($Hit.Kind -eq 'npc' -and $Hit.Shop)
	$cardW = 252
	if ($isShop)
	{
		$cardW = 452
		$stageW = 0
		if ($Ctx.stage) { $stageW = [int]$Ctx.stage.ClientSize.Width }
		if ($stageW -gt 0) { $cardW = [Math]::Min(452, [Math]::Max(320, $stageW - 90)) }
	}
	$detail.Width = $cardW
	$detail.BackColor = if ($isShop) { $pal.Panel } else { $pal.Panel2 }
	$detail.BorderStyle = if ($isShop) { 'None' } else { 'FixedSingle' }
	$detail.AutoScroll = $false
	$detail.Cursor = [System.Windows.Forms.Cursors]::Default

	$barColor = [System.Drawing.Color]::FromArgb(90, 176, 255)
	$kindWord = 'node'
	switch ($Hit.Kind)
	{
		'species' { $kindWord = $Hit.Cat; if ($Hit.Color) { $barColor = $Hit.Color } }
		'treasure' { $kindWord = 'treasure dig spot'; $barColor = [System.Drawing.Color]::FromArgb(212, 149, 58) }
		'statue' { $kindWord = 'statue'; $barColor = [System.Drawing.Color]::FromArgb(77, 157, 224) }
		'npc' {
			$kindWord = if ($Hit.Shop) { ([string]$Hit.Shop.Kind).ToLowerInvariant() } else { 'npc' }
			$barColor = if ($Hit.Shop) { [System.Drawing.Color]::FromArgb(255, 200, 87) } else { [System.Drawing.Color]::FromArgb(63, 209, 106) }
		}
	}

	$bar = New-Object System.Windows.Forms.Panel
	$bar.BackColor = $barColor
	if ($isShop) { $bar.SetBounds(0, 0, $cardW, 3) } else { $bar.SetBounds(0, 0, 4, 24) }
	$detail.Controls.Add($bar)

	$nameX = 12; $nameY = 5; $nameW = $cardW - 46; $nameSize = 9.5
	$kindY = 24
	if ($isShop)
	{
		$lblDot = New-Object System.Windows.Forms.Label
		$lblDot.Text = [string][char]0x25C6
		$lblDot.ForeColor = $barColor
		$lblDot.Font = Get-UIFont 'Segoe UI' 7.5
		$lblDot.AutoSize = $false
		$lblDot.Size = New-Object System.Drawing.Size(14, 14)
		$lblDot.Location = New-Object System.Drawing.Point(12, 11)
		$lblDot.BackColor = [System.Drawing.Color]::Transparent
		$detail.Controls.Add($lblDot)
		$nameX = 27; $nameY = 8; $nameW = $cardW - 64; $nameSize = 10.5
		$kindY = 29
	}

	$lblName = New-Object System.Windows.Forms.Label
	$lblName.Text = $Hit.Label
	$lblName.ForeColor = $pal.Fg
	$lblName.Font = Get-UIFont 'Segoe UI' $nameSize 'Bold'
	$lblName.AutoSize = $false
	$lblName.Size = New-Object System.Drawing.Size($nameW, 20)
	$lblName.Location = New-Object System.Drawing.Point($nameX, $nameY)
	$lblName.Cursor = [System.Windows.Forms.Cursors]::SizeAll
	$detail.Controls.Add($lblName)

	$lblKind = New-Object System.Windows.Forms.Label
	$bits = @($kindWord)
	if ($Hit.Zone) { $bits += $Hit.Zone }
	if ($S.Meta) { $bits += $S.Meta.Title }
	$kindText = ($bits -join ' · ')
	if ($isShop) { $kindText = $kindText.ToUpperInvariant() }
	$lblKind.Text = $kindText
	$lblKind.ForeColor = $pal.Mut
	$lblKind.Font = Get-UIFont 'Segoe UI' 7.5
	$lblKind.AutoSize = $false
	$lblKind.Size = New-Object System.Drawing.Size(($(if ($isShop) { $cardW - 24 } else { $cardW - 46 })), 15)
	$lblKind.Location = New-Object System.Drawing.Point(12, $kindY)
	$lblKind.Cursor = [System.Windows.Forms.Cursors]::SizeAll
	$detail.Controls.Add($lblKind)

	$y = if ($isShop) { 49 } else { 44 }
	if ($Hit.Kind -eq 'treasure' -and $Hit.Full)
	{
		$imgPath = Join-Path $S.D.AssetRoot ($Hit.Full -replace '/', '\')
		if (Test-Path -LiteralPath $imgPath)
		{
			$pb = New-Object System.Windows.Forms.PictureBox
			try
			{
				$fs = New-Object System.IO.FileStream($imgPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
				$src = [System.Drawing.Image]::FromStream($fs)
				$pb.Image = New-Object System.Drawing.Bitmap($src)
				$src.Dispose()
				$fs.Dispose()
			}
			catch { }
			$pb.SizeMode = 'Zoom'
			$pb.Size = New-Object System.Drawing.Size(($cardW - 24), 124)
			$pb.Location = New-Object System.Drawing.Point(12, $y)
			$pb.BackColor = [System.Drawing.Color]::FromArgb(24, 20, 14)
			$detail.Controls.Add($pb)
			$y += 130
		}
	}

	if ($isShop)
	{
		$inv = [System.Globalization.CultureInfo]::InvariantCulture
		$dlNames = New-Object System.Windows.Forms.Label
		$dlNames.Text = "World x`r`nWorld z"
		$dlNames.ForeColor = $pal.Dim
		$dlNames.Font = Get-UIFont 'Segoe UI' 8
		$dlNames.AutoSize = $false
		$dlNames.Size = New-Object System.Drawing.Size(52, 34)
		$dlNames.Location = New-Object System.Drawing.Point(12, $y)
		$detail.Controls.Add($dlNames)

		$dlVals = New-Object System.Windows.Forms.Label
		$dlVals.Text = (([double]$Hit.WX).ToString('0.0', $inv) + "`r`n" + ([double]$Hit.WZ).ToString('0.0', $inv))
		$dlVals.ForeColor = $pal.Fg
		$dlVals.Font = Get-UIFont 'Consolas' 8
		$dlVals.AutoSize = $false
		$dlVals.Size = New-Object System.Drawing.Size(($cardW - 76), 34)
		$dlVals.Location = New-Object System.Drawing.Point(64, $y)
		$detail.Controls.Add($dlVals)
		$y += 50
	}
	else
	{
		$dl = New-Object System.Windows.Forms.Label
		$lines = @(('World x{0,8:0.0}' -f $Hit.WX), ('World z{0,8:0.0}' -f $Hit.WZ))
		if ($Hit.Kind -eq 'species') { $lines += ('Known at{0,7:N0} spots here' -f $S.PinsFor($Hit.Label).Count) }
		$dl.Text = ($lines -join "`r`n")
		$dl.ForeColor = $pal.Fg
		$dl.Font = Get-UIFont 'Consolas' 8
		$dl.AutoSize = $false
		$dl.Size = New-Object System.Drawing.Size(($cardW - 24), (16 + 14 * $lines.Count))
		$dl.Location = New-Object System.Drawing.Point(12, $y)
		$detail.Controls.Add($dl)
		$y += $dl.Height + 6
	}

	$btnX = New-Object System.Windows.Forms.Button
	$btnX.Text = [string][char]0x00D7
	$btnX.FlatStyle = 'Flat'
	$btnX.FlatAppearance.BorderSize = 0
	$btnX.BackColor = [System.Drawing.Color]::Transparent
	$btnX.ForeColor = $pal.Dim
	$btnX.Size = New-Object System.Drawing.Size(22, 20)
	$btnX.Location = New-Object System.Drawing.Point(($cardW - 30), $(if ($isShop) { 8 } else { 3 }))
	if ($isShop)
	{
		$btnX.Font = Get-UIFont 'Segoe UI' 12
		$btnX.Size = New-Object System.Drawing.Size(24, 24)
		$btnX.Location = New-Object System.Drawing.Point(($cardW - 34), 5)
	}
	$btnX.Add_Click({ Update-AtlasDetail $Ctx $null }.GetNewClosure())
	$detail.Controls.Add($btnX)

	$hitX = $Hit.X
	$hitY = $Hit.Y
	$btnCentre = New-AtlasButton -Text 'Centre' -Tip 'Centre the map on this spot'
	$btnW = if ($isShop) { $cardW - 24 } else { 104 }
	$btnH = if ($isShop) { 26 } else { 24 }
	$btnCentre.SetBounds(12, $y, $btnW, $btnH)
	$btnCentre.Add_Click({
			$Ctx.canvas.CentreOn($hitX, $hitY)
			Update-AtlasDetail $Ctx $null
		}.GetNewClosure())
	$detail.Controls.Add($btnCentre)

	if ($Hit.Kind -eq 'species')
	{
		$nm = $Hit.Label
		$btnOnly = New-AtlasButton -Text 'Show only this' -Tip 'Switch every other entry off'
		$btnOnly.SetBounds(120, $y, 120, 24)
		$btnOnly.Add_Click({
				$Ctx.S.OnlyThis($nm)
				Sync-Atlas $Ctx
			}.GetNewClosure())
		$detail.Controls.Add($btnOnly)
	}
	$y += $btnH + 8

	if ($Hit.Shop)
	{
		$viewW = $cardW - 24
		$view = New-Object Atlas.ShopView
		$view.Shop = $Hit.Shop
		$view.AssetRoot = $S.D.AssetRoot
		$view.SetPalette($pal.Panel, $pal.Panel2, [System.Drawing.Color]::FromArgb(31, 42, 57), $pal.Line, [System.Drawing.Color]::FromArgb(49, 65, 90), $pal.Fg, $pal.Mut, $pal.Dim, $pal.Acc, [System.Drawing.Color]::FromArgb(255, 200, 87))
		$view.SetBounds(12, $y, $viewW, 200)
		$viewH = [int]$view.ContentHeight
		$maxCard = 560
		$stageH = 0
		if ($Ctx.stage) { $stageH = [int]$Ctx.stage.ClientSize.Height }
		if ($stageH -gt 0) { $maxCard = [Math]::Min($maxCard, [Math]::Max(220, $stageH - 16)) }
		$room = [Math]::Max(80, $maxCard - $y - 12)
		if ($viewH -gt $room) { $viewH = $room }
		if ($viewH -lt 40) { $viewH = 40 }
		$view.Height = $viewH
		$view.Add_MouseEnter({
			$f = $view.FindForm()
			if ($f -and $f.ContainsFocus -and $view.CanFocus) { $view.Focus() }
		}.GetNewClosure())
		$view.Add_MouseLeave({
			try { if ($Ctx.canvas -and $Ctx.canvas.CanFocus) { $Ctx.canvas.Focus() } } catch { }
		}.GetNewClosure())
		$detail.Controls.Add($view)
		$y += $viewH
	}

	if ($isShop -and -not $detail.PSObject.Properties['AtlasRound'])
	{
		$detail | Add-Member -MemberType NoteProperty -Name 'AtlasRound' -Value $true
		$detail.Add_Paint({
			param($s, $e)
			if (-not $s.Tag -or -not $s.Tag.Shop) { return }
			$w = [int]$s.ClientSize.Width; $h = [int]$s.ClientSize.Height
			if ($w -lt 4 -or $h -lt 4) { return }
			$x0 = 0.5; $y0 = 0.5; $x1 = $w - 0.5; $y1 = $h - 0.5
			$r = [Math]::Min(11, [Math]::Min($w, $h) / 2); $d = $r * 2
			$p = New-Object System.Drawing.Drawing2D.GraphicsPath
			$p.AddArc($x0, $y0, $d, $d, 180, 90)
			$p.AddArc(($x1 - $d), $y0, $d, $d, 270, 90)
			$p.AddArc(($x1 - $d), ($y1 - $d), $d, $d, 0, 90)
			$p.AddArc($x0, ($y1 - $d), $d, $d, 90, 90)
			$p.CloseFigure()
			$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(49, 65, 90), 1)
			$e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
			$e.Graphics.DrawPath($pen, $p)
			$pen.Dispose(); $p.Dispose()
		}.GetNewClosure())
	}

	$detail.Tag = $Hit
	$detail.Visible = $true
	$sw = [int]$Ctx.stage.ClientSize.Width
	$sh = [int]$Ctx.stage.ClientSize.Height
	$detail.Height = $y + $(if ($isShop) { 10 } else { 32 })
	if ($sw -gt 0 -and $sh -gt 0)
	{
		$detail.Height = [Math]::Min($detail.Height, [Math]::Max(120, $sh - 16))
		$pos = $Ctx.detailPos
		if ($pos) {
			$px = [Math]::Min([Math]::Max(8, $pos.X), [Math]::Max(8, $sw - $detail.Width - 8))
			$py = [Math]::Min([Math]::Max(8, $pos.Y), [Math]::Max(8, $sh - $detail.Height - 8))
			$Ctx.detailPos = New-Object System.Drawing.Point([int]$px, [int]$py)
			$detail.Location = $Ctx.detailPos
		} else {
			$lx = [Math]::Max(8, ($sw - $detail.Width) - 60)
			$ly = [Math]::Max(8, ($sh - $detail.Height) - 14)
			$detail.Location = New-Object System.Drawing.Point([int]$lx, [int]$ly)
		}
	}

	Set-AtlasCardRegion -Panel $detail -Round $isShop
	$bar.Cursor = [System.Windows.Forms.Cursors]::SizeAll
	$startDrag = {
		param($s, $e)
		if ($e.Button -ne 'Left') { return }
		$p = [System.Windows.Forms.Cursor]::Position
		$Ctx.detailDrag = @{ X = $p.X; Y = $p.Y; L = $detail.Left; T = $detail.Top }
		$detail.Capture = $true
		$detail.Cursor = [System.Windows.Forms.Cursors]::SizeAll
	}.GetNewClosure()
	$moveDrag = {
		param($s, $e)
		$drag = $Ctx.detailDrag
		if (-not $drag) { return }
		$p = [System.Windows.Forms.Cursor]::Position
		$sw = [int]$Ctx.stage.ClientSize.Width
		$sh = [int]$Ctx.stage.ClientSize.Height
		$nx = [Math]::Min([Math]::Max(8, $drag.L + ($p.X - $drag.X)), [Math]::Max(8, $sw - $detail.Width - 8))
		$ny = [Math]::Min([Math]::Max(8, $drag.T + ($p.Y - $drag.Y)), [Math]::Max(8, $sh - $detail.Height - 8))
		$detail.Location = New-Object System.Drawing.Point([int]$nx, [int]$ny)
		$Ctx.detailPos = $detail.Location
	}.GetNewClosure()
	$endDrag = {
		if ($Ctx.detailDrag) {
			$Ctx.detailDrag = $null
			$detail.Capture = $false
			$detail.Cursor = [System.Windows.Forms.Cursors]::Default
		}
	}.GetNewClosure()
	foreach ($hdr in @($bar, $lblName, $lblKind))
	{
		$hdr.Cursor = [System.Windows.Forms.Cursors]::SizeAll
		$hdr.Add_MouseDown($startDrag)
	}
	$detail.Add_MouseDown($startDrag)
	$detail.Add_MouseMove($moveDrag)
	$detail.Add_MouseUp($endDrag)
}

function Sync-Atlas
{
	param([hashtable]$Ctx, [switch]$KeepDetail)
	Initialize-AtlasCore | Out-Null
	$S = $Ctx.S
	$S.BuildGroups()
	$Ctx.sidebar.SetState($S)
	$Ctx.canvas.RefreshFilters()
	Sync-AtlasToggles -Ctx $Ctx
	Update-AtlasStatus -Ctx $Ctx
	$Ctx.mining.Visible = ($S.MapKey -eq 'cardia' -or $S.MapKey -eq 'cave') -and [Atlas.State]::IsOreFilter($S.Query)
	$Ctx.sideTop.Height = if ($Ctx.mining.Visible) { 147 } else { 87 }
	if (-not $KeepDetail) { Update-AtlasDetail -Ctx $Ctx -Hit $Ctx.detail.Tag }
}

function GoAtlasWorld
{
	param([hashtable]$Ctx, [string]$Key, [switch]$KeepView)
	Initialize-AtlasCore | Out-Null
	$S = $Ctx.S
	if (-not $S.Find($Key)) { return }
	if (-not $KeepView)
	{
		$Ctx.canvas.SetWorld($Key, $false)
		Update-AtlasDetail -Ctx $Ctx -Hit $null
	}
	else { $Ctx.canvas.SetWorld($Key, $true) }

	$inGroup = $false
	foreach ($w in $S.VisibleWorlds($S.MapGroup)) { if ($w.Key -eq $Key) { $inGroup = $true; break } }
	if (-not $inGroup)
	{
		$follow = [Atlas.State]::GroupOf($Key)
		if ($follow -ne 'unsorted')
		{
			$S.MapGroup = $follow
			$want = [Atlas.State]::GRP_TITLE[$follow]
			for ($i = 0; $i -lt $Ctx.cmbGroup.Items.Count; $i++)
			{
				if ([string]$Ctx.cmbGroup.Items[$i] -eq $want) { $Ctx.cmbGroup.SelectedIndex = $i; break }
			}
		}
	}
	Sync-AtlasJump -Ctx $Ctx
	Sync-Atlas -Ctx $Ctx
}

function Set-AtlasQuery
{
	param([hashtable]$Ctx, [string]$Q, [switch]$NoBox)
	$Ctx.S.SetQuery($Q)
	if (-not $NoBox) { $Ctx.txtSearch.Text = $Ctx.S.Query }
	Sync-Atlas -Ctx $Ctx
}

function Start-AtlasLoadNotice
{
	param([hashtable]$Ctx)
	$lbl = $Ctx.lblLoading
	$cv = $Ctx.canvas
	$timer = New-Object System.Windows.Forms.Timer
	$timer.Interval = 250
	$Ctx.timers.Add($timer)
	$timer.Add_Tick({
			if ($cv.BackgroundStatus -ne '')
			{
				$lbl.Text = $cv.BackgroundStatus
				$lbl.Visible = $true
				$timer.Stop()
				return
			}
			if ($cv.BackgroundReady) { $lbl.Visible = $false; $timer.Stop() }
		}.GetNewClosure())
	$timer.Start()
}

function Show-Atlas
{
	param([switch]$SkipAssetSync)

	Initialize-AtlasCore | Out-Null
	if ($script:AtlasForm -and -not $script:AtlasForm.IsDisposed)
	{
		try
		{
			if (-not $script:AtlasForm.Visible) { $script:AtlasForm.Show() | Out-Null }
			$script:AtlasForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
			$script:AtlasForm.BringToFront()
			$script:AtlasForm.Activate()
			return $script:AtlasForm
		}
		catch { }
	}

	if (-not $SkipAssetSync)
	{
		try
		{
			if (Sync-AtlasAssets) { Reset-AtlasCaches }
		}
		catch { }
	}

	$data = $null
	try { $data = Get-AtlasData }
	catch
	{
		if (Test-AtlasDownloadActive -Kind 'AtlasData')
		{
			$script:AtlasOpenPending = $true
			Show-AtlasNotice -Text 'The map data is downloading. The map will open as soon as it is ready.' -Caption 'Entropia Map' -Icon 'Information'
			return $null
		}
		$msg = "The Entropia Map data could not be loaded.`r`n`r`n$($_.Exception.Message)"
		if (Test-DashboardCommand 'Show-DarkMessageBox') { Show-DarkMessageBox $msg 'Entropia Map' 'Ok' 'Error' | Out-Null }
		else { [System.Windows.Forms.MessageBox]::Show($msg, 'Entropia Map') | Out-Null }
		return $null
	}

	$pal = Get-AtlasPalette
	$S = New-Object Atlas.State
	$S.D = $data
	$S.MapKey = $data.Maps[0].Key
	$S.MapGroup = 'all'

	$cfgMap = Get-AtlasConfigValue 'Map' ''
	$cfgGrp = Get-AtlasConfigValue 'Group' 'all'
	$cfgQ = Get-AtlasConfigValue 'Query' ''
	if ($cfgMap -and $S.Find($cfgMap)) { $S.MapKey = $cfgMap }
	if ($cfgGrp -in [Atlas.State]::GROUPS) { $S.MapGroup = $cfgGrp }
	$S.SetQuery($cfgQ)
	$S.BuildGroups()

	$form = New-Object System.Windows.Forms.Form
	$form.Text = 'Entropia Map'
	$form.Size = New-Object System.Drawing.Size(1360, 800)
	$form.MinimumSize = New-Object System.Drawing.Size(960, 580)
	$form.StartPosition = 'CenterScreen'
	$form.BackColor = $pal.Bg
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
	$form.KeyPreview = $true
	try { if ($global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.Icon) { $form.Icon = $global:DashboardConfig.UI.MainForm.Icon } } catch { }

	$Ctx = @{
		S = $S; pal = $pal; form = $form
		canvas = $null; sidebar = $null
		top = $null; stage = $null; side = $null; detail = $null; sideTop = $null
		detailPos = $null; detailDrag = $null
		mining = $null; mineButtons = $null; guideButtons = $null
		cmbGroup = $null; cmbJump = $null
		btnPrev = $null; btnNext = $null
		btnSpawns = $null; btnZones = $null
		txtSearch = $null; lblLoading = $null
		searchHost = $null; btnClear = $null
		stMap = $null; stSpots = $null; stSpawn = $null; stScale = $null
		zoomPanel = $null; btnClose = $null; brand = $null
		btnAll = $null; btnNone = $null
		btnAllMaps = $null; btnNoneMaps = $null
		btnFold = $null; btnUnfold = $null
		zoomBtns = $null
		timers = (New-Object System.Collections.Generic.List[System.Windows.Forms.Timer])
	}

	$script:AtlasForm = $form
	if (-not $global:DashboardConfig.Resources.ContainsKey('AtlasForm')) { $global:DashboardConfig.Resources['AtlasForm'] = $null }
	$global:DashboardConfig.Resources.AtlasForm = $form

	$top = New-Object System.Windows.Forms.Panel
	$top.Name = 'AtlasTopBar'
	$top.Dock = 'Top'
	$top.Height = 42
	$top.BackColor = $pal.Panel
	$top.Padding = New-Object System.Windows.Forms.Padding(10, 0, 10, 0)
	Add-AtlasDrag -Control $top
	if (Test-DashboardCommand 'Register-ThemedTopBarPaint') { try { Register-ThemedTopBarPaint $top } catch { } }
	else
	{
		$top.Add_Paint({
				param($s, $e)
				$e.Graphics.FillRectangle((New-Object System.Drawing.SolidBrush((Get-AtlasPalette).Line)), 0, $s.Height - 1, $s.Width, 1)
			}.GetNewClosure())
	}

	$strip = New-Object System.Windows.Forms.FlowLayoutPanel
	$strip.Dock = 'Fill'
	$strip.FlowDirection = 'LeftToRight'
	$strip.WrapContents = $true
	$strip.AutoScroll = $false
	$strip.BackColor = [System.Drawing.Color]::Transparent
	$strip.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
	$strip.Margin = New-Object System.Windows.Forms.Padding(0)
	$top.Controls.Add($strip)
	Add-AtlasDrag -Control $strip

	$closeHost = New-Object System.Windows.Forms.Panel
	$closeHost.Dock = 'Right'
	$closeHost.Width = 42
	$closeHost.BackColor = [System.Drawing.Color]::Transparent
	$top.Controls.Add($closeHost)

	$btnClose = New-Object System.Windows.Forms.Button
	$btnClose.Text = [string][char]0x00D7
	$btnClose.FlatStyle = 'Flat'
	$btnClose.FlatAppearance.BorderSize = 0
	$btnClose.BackColor = [System.Drawing.Color]::Transparent
	$btnClose.ForeColor = $pal.Mut
	$btnClose.Font = Get-UIFont 'Segoe UI' 12
	$btnClose.Size = New-Object System.Drawing.Size(28, 26)
	$btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
	$btnClose.Location = New-Object System.Drawing.Point(10, 2)
	if (Test-DashboardCommand 'Register-ThemedButtonFadeEffect') { try { Register-ThemedButtonFadeEffect -Button $btnClose -Variant 'Danger' } catch { } }
	$closeHost.Controls.Add($btnClose)

	$syncHeader = {
		$w = [int]$strip.ClientSize.Width
		if ($w -le 0) { return }
		$need = 0
		try { $need = [int]$strip.GetPreferredSize((New-Object System.Drawing.Size($w, 0))).Height } catch { $need = 0 }
		$strip.PerformLayout()
		foreach ($c in $strip.Controls) { if ($c.Visible -and $c.Bottom -gt $need) { $need = $c.Bottom } }
		$need += 8
		if ($need -lt 42) { $need = 42 }
		if ($top.Height -ne $need)
		{
			$top.Height = $need
			$strip.Height = $need
			$top.PerformLayout()
		}
	}.GetNewClosure()
	$strip.Add_SizeChanged({ & $syncHeader }.GetNewClosure())
	$form.Add_Resize({ & $syncHeader }.GetNewClosure())

	$brand = New-Object System.Windows.Forms.Label
	$brand.Text = 'Entropia Map'
	$brand.ForeColor = $pal.Fg
	$brand.Font = Get-UIFont 'Segoe UI' 9.5 'Bold'
	$brand.AutoSize = $false
	$brand.Size = New-Object System.Drawing.Size(132, 26)
	$brand.TextAlign = 'MiddleLeft'
	$brand.BackColor = [System.Drawing.Color]::Transparent
	$brand.Margin = New-Object System.Windows.Forms.Padding(4, 0, 10, 0)
	Add-AtlasDrag -Control $brand
	$strip.Controls.Add($brand)

	$cmbGroup = New-AtlasCombo -Width 122
	$cmbGroup.Margin = New-Object System.Windows.Forms.Padding(0, 3, 4, 0)
	$sep = ([string][char]0x2500).PadRight(10, [char]0x2500)
	foreach ($t in @('All Worlds', 'Main Maps', 'Dungeons', 'Others', 'Event', $sep, 'Farm Maps', 'Platin Level Maps', 'Royal Level Maps'))
	{
		[void]$cmbGroup.Items.Add($t)
	}
	if ($cmbGroup.GetType().GetMethod('SetHeaderIndices')) { $cmbGroup.SetHeaderIndices(@(5)) }
	$groupSel = switch ($S.MapGroup)
	{
		'all' { 0 }
		'main' { 1 }
		'dungeons' { 2 }
		'fields' { 3 }
		'event' { 4 }
		'farm' { 6 }
		'platin' { 7 }
		'royal' { 8 }
		default { 0 }
	}
	$hasSup = [bool]$cmbGroup.PSObject.Properties['SuppressChanges']
	if ($hasSup) { $cmbGroup.SuppressChanges = $true }
	try { $cmbGroup.SelectedIndex = $groupSel }
	finally { if ($hasSup) { $cmbGroup.SuppressChanges = $false } }
	if ($cmbGroup.GetType().GetMethod('FitDropDownWidth')) { try { $cmbGroup.FitDropDownWidth(240) } catch { } }
	$strip.Controls.Add($cmbGroup)

	$btnPrev = New-AtlasButton -Text ([string][char]0x2039) -Width 26 -Tip 'Previous world'
	$btnPrev.Height = 22
	$btnPrev.Margin = New-Object System.Windows.Forms.Padding(0, 4, 2, 0)
	$strip.Controls.Add($btnPrev)

	$cmbJump = New-AtlasCombo -Width 168
	$cmbJump.Margin = New-Object System.Windows.Forms.Padding(0, 3, 2, 0)
	$strip.Controls.Add($cmbJump)

	$btnNext = New-AtlasButton -Text ([string][char]0x203A) -Width 26 -Tip 'Next world'
	$btnNext.Height = 22
	$btnNext.Margin = New-Object System.Windows.Forms.Padding(0, 4, 8, 0)
	$strip.Controls.Add($btnNext)

	$guideDefs = @(
		@{ t = 'Statues'; q = 'statue'; g = [string][char]0x265C; a = '#ffffff' }
		@{ t = 'Ore'; q = '(ore)'; g = [string][char]0x26CF; a = '#ffb020' }
		@{ t = 'Plant'; q = '(plant)'; g = [string][char]0x2740; a = '#3fd16a' }
		@{ t = 'Omega'; q = 'omega'; g = [string][char]0x03A9; a = '#ffb64a' }
		@{ t = 'Astral'; q = 'astral crystal|celestial map'; g = [string][char]0x2726; a = '#78dcff' }
		@{ t = 'Treasure'; q = 'treasure'; g = [string][char]0x25C8; a = '#d4953a' }
		@{ t = 'Daily Quest'; q = 'daily quest'; g = [string][char]0x2694; a = '#ff6ea8' }
	)
	$guideButtons = New-Object System.Collections.Generic.List[System.Windows.Forms.Button]
	foreach ($gd in $guideDefs)
	{
		$label = $gd.g + ' ' + $gd.t
		$b = New-AtlasButton -Text $label -Tip ("{0} guide: filter this world and list every other world's hits" -f $gd.t) -Accent $gd.a
		$b.Font = Get-UIFont 'Segoe UI Symbol' 8.5
		$q = $gd.q
		$b.Tag = $q
		$b.Width = [Math]::Max(54, ([System.Windows.Forms.TextRenderer]::MeasureText($label, $b.Font).Width + 16))
		$b.Height = 22
		$b.Margin = New-Object System.Windows.Forms.Padding(0, 4, 3, 0)
		$b.Add_Click({
				if ($Ctx.S.Query -eq $q) { Set-AtlasQuery -Ctx $Ctx -Q '' } else { Set-AtlasQuery -Ctx $Ctx -Q $q }
			}.GetNewClosure())
		$guideButtons.Add($b)
		$strip.Controls.Add($b)
	}

	$searchHost = New-Object System.Windows.Forms.Panel
	$searchHost.Size = New-Object System.Drawing.Size(184, 23)
	$searchHost.BorderStyle = 'FixedSingle'
	$searchHost.BackColor = $pal.Input
	$searchHost.Margin = New-Object System.Windows.Forms.Padding(6, 5, 6, 0)

	$txtSearch = New-Object System.Windows.Forms.TextBox
	$txtSearch.Font = Get-UIFont 'Segoe UI' 8.5
	$txtSearch.BackColor = $pal.Input
	$txtSearch.ForeColor = $pal.Fg
	$txtSearch.BorderStyle = 'None'
	$txtSearch.AutoSize = $false
	$txtSearch.SetBounds(4, 4, 156, 16)
	$txtSearch.Text = $S.Query
	$searchHost.Controls.Add($txtSearch)

	$btnClear = New-Object System.Windows.Forms.Button
	$btnClear.Text = [string][char]0x00D7
	$btnClear.FlatStyle = 'Flat'
	$btnClear.FlatAppearance.BorderSize = 0
	$btnClear.BackColor = $pal.Input
	$btnClear.ForeColor = $pal.Dim
	$btnClear.Font = Get-UIFont 'Segoe UI' 9
	$btnClear.SetBounds(162, 3, 18, 18)
	$btnClear.Cursor = [System.Windows.Forms.Cursors]::Hand
	$btnClear.TabStop = $false
	$btnClear.Visible = ($S.Query -ne '')
	$btnClear.Add_Click({
			Set-AtlasQuery -Ctx $Ctx -Q ''
			$btnClear.Visible = $false
		}.GetNewClosure())
	$searchHost.Controls.Add($btnClear)
	$strip.Controls.Add($searchHost)

	try { $null = $txtSearch.Handle; [Atlas.Cue]::Set($txtSearch, 'Search anything...') } catch { }

	$btnSpawns = New-AtlasButton -Text 'Spawns' -Width 64 -Tip 'Show or hide every spawn point on the map'
	$btnSpawns.Height = 22
	$btnSpawns.Margin = New-Object System.Windows.Forms.Padding(0, 4, 3, 0)
	$strip.Controls.Add($btnSpawns)

	$btnZones = New-AtlasButton -Text 'Zones' -Width 56 -Tip 'Show zone names'
	$btnZones.Height = 22
	$btnZones.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
	$strip.Controls.Add($btnZones)

	$main = New-Object System.Windows.Forms.Panel
	$main.Dock = 'Fill'
	$main.BackColor = $pal.Bg

	$stage = New-Object System.Windows.Forms.Panel
	$stage.Dock = 'Fill'
	$stage.BackColor = $pal.Bg

	$canvas = New-Object Atlas.Canvas
	$canvas.Dock = 'Fill'
	$canvas.S = $S
	$canvas.TabStop = $true
	$canvas.BackColor = [System.Drawing.Color]::FromArgb(10, 14, 19)
	$nodeTip = Get-AtlasNodeToolTip
	$nodeTip.SetToolTip($canvas, ' ')

	$lblLoading = New-Object System.Windows.Forms.Label
	$lblLoading.Text = 'Loading map...'
	$lblLoading.ForeColor = $pal.Mut
	$lblLoading.Font = Get-UIFont 'Segoe UI' 9
	$lblLoading.AutoSize = $true
	$lblLoading.BackColor = [System.Drawing.Color]::Transparent
	$lblLoading.Location = New-Object System.Drawing.Point(16, 14)
	$stage.Controls.Add($lblLoading)

	$zoomPanel = New-Object System.Windows.Forms.Panel
	$zoomPanel.Size = New-Object System.Drawing.Size(34, 114)
	$zoomPanel.BackColor = [System.Drawing.Color]::Transparent
	$stage.Controls.Add($zoomPanel)
	$zoomBtns = [ordered]@{}
	$zi = 0
	foreach ($zb in @(@{ k = 'in'; t = '+'; tip = 'Zoom in' }, @{ k = 'out'; t = '-'; tip = 'Zoom out' }, @{ k = 'fit'; t = [string][char]0x2262; tip = 'Fit the whole map' }))
	{
		$b = New-AtlasButton -Text $zb.t -Width 34 -Tip $zb.tip
		$b.Height = 34
		$b.Font = Get-UIFont 'Segoe UI' 12
		$b.Location = New-Object System.Drawing.Point(0, ($zi * 40))
		$zoomPanel.Controls.Add($b)
		$zoomBtns[$zb.k] = $b
		$zi++
	}

	$detail = New-Object System.Windows.Forms.Panel
	$detail.Size = New-Object System.Drawing.Size(252, 120)
	$detail.BackColor = $pal.Panel2
	$detail.BorderStyle = 'FixedSingle'
	$detail.Visible = $false
	$stage.Controls.Add($detail)

	$stage.Controls.Add($canvas)

	$side = New-Object System.Windows.Forms.Panel
	$side.Dock = 'Left'
	$side.Width = 274
	$side.BackColor = $pal.Panel

	$sidePad = 10
	$sideGap = 6
	$sideCol = [int](($side.Width - (2 * $sidePad) - $sideGap) / 2)
	$sideRight = $sidePad + (2 * $sideCol) + $sideGap

	$sidebar = New-Object Atlas.Sidebar
	$sidebar.Dock = 'Fill'
	$sidebar.S = $S
	$sidebar.BackColor = $pal.Panel
	$side.Controls.Add($sidebar)

	$sideTop = New-Object System.Windows.Forms.Panel
	$sideTop.Dock = 'Top'
	$sideTop.Height = 87
	$sideTop.BackColor = $pal.Panel
	$side.Controls.Add($sideTop)

	$sideHead = New-Object System.Windows.Forms.Panel
	$sideHead.SetBounds(0, 0, 274, 29)
	$sideHead.BackColor = $pal.Panel
	$sideTop.Controls.Add($sideHead)

	$lblFilters = New-Object System.Windows.Forms.Label
	$lblFilters.Text = 'FILTERS'
	$lblFilters.ForeColor = $pal.Mut
	$lblFilters.Font = Get-UIFont 'Segoe UI' 7.5 'Bold'
	$lblFilters.AutoSize = $true
	$lblFilters.Location = New-Object System.Drawing.Point(12, 8)
	$lblFilters.BackColor = [System.Drawing.Color]::Transparent
	$sideHead.Controls.Add($lblFilters)

	$btnAll = New-AtlasButton -Text 'All' -Width 42 -Tip 'Select everything on this world'
	$btnAll.SetBounds(($sideRight - 48 - $sideGap - 42), 3, 42, 23)
	$sideHead.Controls.Add($btnAll)
	$btnNone = New-AtlasButton -Text 'None' -Width 48 -Tip 'Deselect everything on this world'
	$btnNone.SetBounds(($sideRight - 48), 3, 48, 23)
	$sideHead.Controls.Add($btnNone)

	$bar1 = New-Object System.Windows.Forms.Panel
	$bar1.SetBounds(0, 29, 274, 29)
	$bar1.BackColor = $pal.Panel
	$sideTop.Controls.Add($bar1)
	$btnAllMaps = New-AtlasButton -Text 'Select all worlds' -Tip 'Select everything on every world'
	$btnAllMaps.SetBounds($sidePad, 3, $sideCol, 23)
	$bar1.Controls.Add($btnAllMaps)
	$btnNoneMaps = New-AtlasButton -Text 'Clear all worlds' -Tip 'Deselect everything on every world'
	$btnNoneMaps.SetBounds(($sidePad + $sideCol + $sideGap), 3, $sideCol, 23)
	$bar1.Controls.Add($btnNoneMaps)

	$bar2 = New-Object System.Windows.Forms.Panel
	$bar2.SetBounds(0, 58, 274, 29)
	$bar2.BackColor = $pal.Panel
	$sideTop.Controls.Add($bar2)
	$btnUnfold = New-AtlasButton -Text 'Expand all' -Tip 'Open every list, and every area inside it'
	$btnUnfold.SetBounds($sidePad, 3, $sideCol, 23)
	$bar2.Controls.Add($btnUnfold)
	$btnFold = New-AtlasButton -Text 'Collapse all' -Tip 'Collapse every list, and every area inside it'
	$btnFold.SetBounds(($sidePad + $sideCol + $sideGap), 3, $sideCol, 23)
	$bar2.Controls.Add($btnFold)

	$mining = New-Object System.Windows.Forms.Panel
	$mining.SetBounds(0, 87, 274, 60)
	$mining.BackColor = $pal.Panel
	$mining.Visible = $false
	$sideTop.Controls.Add($mining)
	$lblMining = New-Object System.Windows.Forms.Label
	$lblMining.Text = 'MINING SETS'
	$lblMining.ForeColor = $pal.Mut
	$lblMining.Font = Get-UIFont 'Segoe UI' 7.5 'Bold'
	$lblMining.AutoSize = $true
	$lblMining.Location = New-Object System.Drawing.Point(12, 6)
	$lblMining.BackColor = [System.Drawing.Color]::Transparent
	$mining.Controls.Add($lblMining)
	$mineButtons = New-Object System.Collections.Generic.List[System.Windows.Forms.Button]
	$mineCol = [int](($side.Width - (2 * $sidePad) - (3 * $sideGap)) / 4)
	for ($i = 0; $i -lt 7; $i++)
	{
		$b = New-AtlasButton -Text ('Set ' + ($i + 1)) -Tip ([Atlas.State]::MINING_LABEL[$i])
		$b.SetBounds(($sidePad + ($i % 4) * ($mineCol + $sideGap)), (23 + [int][Math]::Floor($i / 4) * 20), $mineCol, 18)
		$b.Font = Get-UIFont 'Segoe UI' 7.5
		$q = [Atlas.State]::MINING_SETS[$i]
		$b.Tag = $q
		$b.Add_Click({
				if ($Ctx.S.Query -eq $q) { Set-AtlasQuery -Ctx $Ctx -Q '' } else { Set-AtlasQuery -Ctx $Ctx -Q $q }
			}.GetNewClosure())
		$mineButtons.Add($b)
		$mining.Controls.Add($b)
	}

	$status = New-Object System.Windows.Forms.Panel
	$status.Dock = 'Bottom'
	$status.Height = 25
	$status.BackColor = $pal.Panel
	$status.Padding = New-Object System.Windows.Forms.Padding(0)

	$stStrip = New-Object System.Windows.Forms.FlowLayoutPanel
	$stStrip.Dock = 'Fill'
	$stStrip.WrapContents = $true
	$stStrip.AutoScroll = $false
	$stStrip.BackColor = [System.Drawing.Color]::Transparent
	$stStrip.Padding = New-Object System.Windows.Forms.Padding(12, 0, 12, 0)
	$status.Controls.Add($stStrip)

	function New-StatusLabel
	{
		param([string]$Text, [System.Drawing.Color]$Color, [System.Drawing.Font]$Font)
	Initialize-AtlasCore | Out-Null
		$l = New-Object System.Windows.Forms.Label
		$l.Text = $Text
		$l.ForeColor = $Color
		$l.AutoSize = $true
		if (-not $Font) { $Font = Get-UIFont 'Segoe UI' 8 }
		$l.Font = $Font
		$l.BackColor = [System.Drawing.Color]::Transparent
		$l.Margin = New-Object System.Windows.Forms.Padding(0, 5, 22, 0)
		$stStrip.Controls.Add($l)
		return $l
	}
	$stMap = New-StatusLabel -Text '-' -Color $pal.Fg -Font (Get-UIFont 'Segoe UI' 8.5 'Bold')
	$stSpots = New-StatusLabel -Text '' -Color $pal.Mut
	$stSpawn = New-StatusLabel -Text '' -Color $pal.Mut
	$stScale = New-StatusLabel -Text '' -Color $pal.Mut -Font (Get-UIFont 'Consolas' 8)

	$main.Controls.Add($stage)
	$main.Controls.Add($side)
	$main.Controls.Add($status)
	$form.Controls.Add($main)
	$form.Controls.Add($top)

	$Ctx.canvas = $canvas; $Ctx.sidebar = $sidebar
	$Ctx.top = $top; $Ctx.stage = $stage; $Ctx.side = $side; $Ctx.detail = $detail
	$Ctx.sideTop = $sideTop
	$Ctx.mining = $mining; $Ctx.mineButtons = $mineButtons; $Ctx.guideButtons = $guideButtons
	$Ctx.cmbGroup = $cmbGroup; $Ctx.cmbJump = $cmbJump
	$Ctx.btnPrev = $btnPrev; $Ctx.btnNext = $btnNext
	$Ctx.btnSpawns = $btnSpawns; $Ctx.btnZones = $btnZones
	$Ctx.txtSearch = $txtSearch; $Ctx.lblLoading = $lblLoading
	$Ctx.searchHost = $searchHost; $Ctx.btnClear = $btnClear
	$Ctx.stMap = $stMap; $Ctx.stSpots = $stSpots; $Ctx.stSpawn = $stSpawn; $Ctx.stScale = $stScale
	$Ctx.zoomPanel = $zoomPanel; $Ctx.btnClose = $btnClose; $Ctx.brand = $brand
	$Ctx.btnAll = $btnAll; $Ctx.btnNone = $btnNone
	$Ctx.btnAllMaps = $btnAllMaps; $Ctx.btnNoneMaps = $btnNoneMaps
	$Ctx.btnFold = $btnFold; $Ctx.btnUnfold = $btnUnfold
	$Ctx.zoomBtns = $zoomBtns

	$btnPrev.Add_Click({
			$visible = $Ctx.S.VisibleWorlds($Ctx.S.MapGroup)
			for ($i = 0; $i -lt $visible.Count; $i++)
			{
				if ($visible[$i].Key -eq $Ctx.S.MapKey)
				{
					if ($i -gt 0) { GoAtlasWorld -Ctx $Ctx -Key $visible[$i - 1].Key }
					break
				}
			}
		}.GetNewClosure())
	$btnNext.Add_Click({
			$visible = $Ctx.S.VisibleWorlds($Ctx.S.MapGroup)
			if ($visible.Count -eq 0) { return }
			$pos = -1
			for ($i = 0; $i -lt $visible.Count; $i++)
			{
				if ($visible[$i].Key -eq $Ctx.S.MapKey) { $pos = $i; break }
			}
			if ($pos -lt 0) { GoAtlasWorld -Ctx $Ctx -Key $visible[0].Key; return }
			if ($pos -lt $visible.Count - 1) { GoAtlasWorld -Ctx $Ctx -Key $visible[$pos + 1].Key }
		}.GetNewClosure())
	$cmbJump.Add_SelectedIndexChanged({
			$sel = $Ctx.cmbJump.SelectedIndex
			if ($sel -lt 0) { return }
			$txt = ([string]$Ctx.cmbJump.Items[$sel]).Trim()
			if (-not $txt -or $txt -eq $Ctx.S.MapKey) { return }
			foreach ($wm in $Ctx.S.D.Maps)
			{
				if ($txt -eq $wm.Title -or $txt.StartsWith($wm.Title + ' [') -or $txt.StartsWith($wm.Title + ' <'))
				{
					if ($wm.Key -ne $Ctx.S.MapKey) { GoAtlasWorld -Ctx $Ctx -Key $wm.Key }
					return
				}
			}
		}.GetNewClosure())
	$cmbGroup.Add_SelectedIndexChanged({
			switch ([string]$Ctx.cmbGroup.SelectedItem)
			{
				'All Worlds' { $Ctx.S.MapGroup = 'all' }
				'Main Maps' { $Ctx.S.MapGroup = 'main' }
				'Dungeons' { $Ctx.S.MapGroup = 'dungeons' }
				'Others' { $Ctx.S.MapGroup = 'fields' }
				'Event' { $Ctx.S.MapGroup = 'event' }
				'Farm Maps' { $Ctx.S.MapGroup = 'farm' }
				'Platin Level Maps' { $Ctx.S.MapGroup = 'platin' }
				'Royal Level Maps' { $Ctx.S.MapGroup = 'royal' }
				default { return }
			}
			Sync-AtlasJump -Ctx $Ctx
			Sync-Atlas -Ctx $Ctx
		}.GetNewClosure())

	$txtSearch.Add_KeyDown({
			param($s, $e)
			if ($e.KeyCode -eq 'Enter' -or $e.KeyCode -eq 'Escape')
			{
				$Ctx.S.SetQuery($txtSearch.Text)
				Sync-Atlas -Ctx $Ctx
				$e.SuppressKeyPress = $true
			}
		}.GetNewClosure())
	$txtSearch.Add_TextChanged({
			if ($txtSearch.Text -ne $Ctx.S.Query) { Set-AtlasQuery -Ctx $Ctx -Q $txtSearch.Text -NoBox }
			$btnClear.Visible = ($txtSearch.Text -ne '')
		}.GetNewClosure())

	$btnSpawns.Add_Click({
			$Ctx.S.ShowSp = -not $Ctx.S.ShowSp
			$Ctx.canvas.RefreshFilters()
			Sync-AtlasToggles -Ctx $Ctx
			Update-AtlasStatus -Ctx $Ctx
		}.GetNewClosure())
	$btnZones.Add_Click({
			$Ctx.S.ShowPl = -not $Ctx.S.ShowPl
			$Ctx.canvas.RefreshFilters()
			Sync-AtlasToggles -Ctx $Ctx
		}.GetNewClosure())

	$sidebar.Add_FilterChanged({
			$Ctx.canvas.RefreshFilters()
			Update-AtlasStatus -Ctx $Ctx
		}.GetNewClosure())
	$sidebar.Add_RowActivated({
			param($s, $e)
			$arg = [string]$e.Key
			if ($arg.Contains('|'))
			{
				$parts = $arg -split '\|', 2
				$mapKey = $parts[0]
				$label = $parts[1]
				GoAtlasWorld -Ctx $Ctx -Key $mapKey -KeepView
				$key = $null
				foreach ($g in $Ctx.S.Groups)
				{
					foreach ($r in $g.Rows)
					{
						if ($r.Label -eq $label -or $r.Key -eq $label) { $key = $r.Key; break }
					}
					if ($key) { break }
				}
				if (-not $key)
				{
					foreach ($pre in @('@s:', '@n:', '@t:', '@i:', '@c:'))
					{
						if ($label.StartsWith($pre)) { $key = $pre + $label.Substring(3); break }
					}
				}
				if ($key)
				{
					$Ctx.S.SetOn($key, $true)
					$pts = $Ctx.S.PinsFor($key)
					if ($pts.Count -eq 0) { $pts = [Atlas.State]::PinsForOn($mapKey, $key, $Ctx.S.D) }
					$Ctx.canvas.FitPoints($pts)
					Sync-Atlas -Ctx $Ctx
				}
			}
			else { $Ctx.canvas.FitPoints($Ctx.S.PinsFor($arg, $e.Zone)) }
		}.GetNewClosure())

	$sidebar.Add_ZoneActivated({
			param($s, $e)
			$pts = $Ctx.S.PinsForZone($e.Group, $e.Zone)
			if ($pts.Count -eq 0) { return }
			$Ctx.canvas.FitPoints($pts)
		}.GetNewClosure())
	$sidebar.Add_WorldActivated({
			param($s, $e)
			if (-not $e.Key) { return }
			GoAtlasWorld -Ctx $Ctx -Key $e.Key
		}.GetNewClosure())
	$refreshAfter = {
		$Ctx.sidebar.SetState($Ctx.S)
		$Ctx.canvas.RefreshFilters()
		Update-AtlasStatus -Ctx $Ctx
	}.GetNewClosure()
	$btnAll.Add_Click({ $Ctx.S.SetAllHere($true) | Out-Null; & $refreshAfter }.GetNewClosure())
	$btnNone.Add_Click({ $Ctx.S.SetAllHere($false) | Out-Null; & $refreshAfter }.GetNewClosure())
	$btnAllMaps.Add_Click({ $Ctx.S.SetAllGlobal($true) | Out-Null; & $refreshAfter }.GetNewClosure())
	$btnNoneMaps.Add_Click({ $Ctx.S.SetAllGlobal($false) | Out-Null; & $refreshAfter }.GetNewClosure())
	$btnFold.Add_Click({ $Ctx.sidebar.FoldAll($false) }.GetNewClosure())
	$btnUnfold.Add_Click({ $Ctx.sidebar.FoldAll($true) }.GetNewClosure())

	$canvas.Add_Picked({
			param($s, $e)
			Update-AtlasDetail -Ctx $Ctx -Hit $e.Hit
		}.GetNewClosure())
	$canvas.Add_HoverChanged({
			$hit = $Ctx.canvas.Hover
			$tip = Get-AtlasNodeToolTip
			if (-not $hit)
			{
				$tip.Hit = $null
				$tip.SetToolTip($Ctx.canvas, ' ')
				return
			}
			$tip.Hit = $hit
			$tip.Size = Get-AtlasNodeTipSize -Hit $hit
			$tip.SetToolTip($Ctx.canvas, ('{0} @ {1:0}, {2:0}' -f $hit.Label, $hit.WX, $hit.WZ))
		}.GetNewClosure())
	$canvas.Add_ViewChanged({
			$Ctx.stScale.Text = ('{0:0.00} px/m' -f ($Ctx.canvas.Sc * $Ctx.S.D.K * $Ctx.canvas.Ppx))
		}.GetNewClosure())
	$canvas.Add_WorldChange({
			param($s, $e)
			if ($e.Key -and $e.Key -ne $Ctx.S.MapKey) { GoAtlasWorld -Ctx $Ctx -Key $e.Key }
		}.GetNewClosure())
	$zoomBtns['in'].Add_Click({ $Ctx.canvas.ZoomAt($Ctx.canvas.Width / 2, $Ctx.canvas.Height / 2, 1.4) }.GetNewClosure())
	$zoomBtns['out'].Add_Click({ $Ctx.canvas.ZoomAt($Ctx.canvas.Width / 2, $Ctx.canvas.Height / 2, 1 / 1.4) }.GetNewClosure())
	$zoomBtns['fit'].Add_Click({ $Ctx.canvas.Fit() }.GetNewClosure())

	$form.Add_KeyDown({
			param($s, $e)
			if ($e.KeyCode -eq 'Escape' -and $detail.Visible) { Update-AtlasDetail -Ctx $Ctx -Hit $null; $e.Handled = $true }
		}.GetNewClosure())
	$btnClose.Add_Click({ $form.Close() }.GetNewClosure())

	$stage.Add_Resize({
			$w = [int]$stage.ClientSize.Width
			$h = [int]$stage.ClientSize.Height
			if ($w -le 0 -or $h -le 0) { return }
			$zoomPanel.Location = New-Object System.Drawing.Point((($w - 34) - 14), (($h - 114) - 14))
			if ($detail.Visible)
			{
				if ($Ctx.detailPos) {
					$px = [Math]::Min([Math]::Max(8, $Ctx.detailPos.X), [Math]::Max(8, $w - $detail.Width - 8))
					$py = [Math]::Min([Math]::Max(8, $Ctx.detailPos.Y), [Math]::Max(8, $h - $detail.Height - 8))
					$Ctx.detailPos = New-Object System.Drawing.Point([int]$px, [int]$py)
					$detail.Location = $Ctx.detailPos
				} else {
					$detail.Location = New-Object System.Drawing.Point(([Math]::Max(8, ($w - $detail.Width) - 60)), ([Math]::Max(8, ($h - $detail.Height) - 14)))
				}
			}
		}.GetNewClosure())

	$applyTheme = {
		$p = Get-AtlasPalette
		$Ctx.form.BackColor = $p.Bg
		$Ctx.stage.BackColor = $p.Bg
		$Ctx.side.BackColor = $p.Panel
		$Ctx.sideTop.BackColor = $p.Panel
		$Ctx.top.BackColor = $p.Panel
		$Ctx.detail.BackColor = $p.Panel2
		foreach ($c in $Ctx.form.Controls)
		{
			if ($c.Name -like 'AtlasGrip*') { $c.BackColor = $p.Bg }
		}
		$list = New-Object System.Collections.Generic.List[System.Windows.Forms.Button]
		foreach ($b in $Ctx.guideButtons) { $list.Add($b) }
		foreach ($b in $Ctx.mineButtons) { $list.Add($b) }
		foreach ($k in $Ctx.zoomBtns.Keys) { $list.Add($Ctx.zoomBtns[$k]) }
		foreach ($b in @($Ctx.btnAll, $Ctx.btnNone, $Ctx.btnAllMaps, $Ctx.btnNoneMaps, $Ctx.btnFold, $Ctx.btnUnfold, $Ctx.btnPrev, $Ctx.btnNext))
		{
			if ($b) { $list.Add($b) }
		}
		foreach ($b in $list)
		{
			$b.BackColor = $p.Panel2
			$b.FlatAppearance.BorderColor = $p.Line
		}
		foreach ($c in @($Ctx.cmbGroup, $Ctx.cmbJump))
		{
			if (-not $c) { continue }
			if (Test-DashboardCommand 'ApplyControlInternal')
			{
				try { ApplyControlInternal $c (Get-Theme); continue } catch { }
			}
			$c.BackColor = $p.Panel2; $c.ForeColor = $p.Mut
		}
		$Ctx.txtSearch.BackColor = $p.Input; $Ctx.txtSearch.ForeColor = $p.Fg
		$Ctx.searchHost.BackColor = $p.Input
		$Ctx.btnClear.BackColor = $p.Input; $Ctx.btnClear.ForeColor = $p.Dim
		$Ctx.brand.ForeColor = $p.Fg
		$Ctx.sidebar.SetPalette($p.Panel, $p.Panel2, $p.Line, $p.Fg, $p.Mut, $p.Dim, $p.Acc)
		$Ctx.stMap.ForeColor = $p.Fg
		foreach ($l in @($Ctx.stSpots, $Ctx.stSpawn, $Ctx.stScale, $Ctx.lblLoading)) { $l.ForeColor = $p.Mut }
		Sync-AtlasToggles -Ctx $Ctx
	}.GetNewClosure()

	$pos = $null
	try { $pos = $global:DashboardConfig.Config['WindowPosition'] } catch { }
	if ($pos)
	{
		try
		{
			$wx = [int]$pos['AtlasFormX']; $wy = [int]$pos['AtlasFormY']
			$ww = [int]$pos['AtlasFormWidth']; $wh = [int]$pos['AtlasFormHeight']
			if ($ww -ge 960 -and $wh -ge 580)
			{
				$rect = New-Object System.Drawing.Rectangle($wx, $wy, $ww, 20)
				if ([System.Windows.Forms.Screen]::AllScreens.WorkingArea.IntersectsWith($rect))
				{
					$form.StartPosition = 'Manual'
					$form.Location = New-Object System.Drawing.Point($wx, $wy)
					$form.Size = New-Object System.Drawing.Size($ww, $wh)
				}
			}
		}
		catch { }
	}
	if ($global:DashboardConfig.UI.MainForm -and $form.StartPosition -ne 'Manual')
	{
		try { $form.Owner = $global:DashboardConfig.UI.MainForm; $form.StartPosition = 'CenterParent' } catch { }
	}

	$form.Add_FormClosed({
			param($s, $e)
			try
			{
				if ($s.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal -and $global:DashboardConfig.Config['WindowPosition'])
				{
					$global:DashboardConfig.Config['WindowPosition']['AtlasFormX'] = $s.Location.X
					$global:DashboardConfig.Config['WindowPosition']['AtlasFormY'] = $s.Location.Y
					$global:DashboardConfig.Config['WindowPosition']['AtlasFormWidth'] = $s.Width
					$global:DashboardConfig.Config['WindowPosition']['AtlasFormHeight'] = $s.Height
				}
				Set-AtlasConfigValue 'Map' $S.MapKey
				Set-AtlasConfigValue 'Group' $S.MapGroup
				Set-AtlasConfigValue 'Query' $S.Query
				if (Test-DashboardCommand 'WriteConfig') { try { WriteConfig } catch { } }
			}
			catch { }
			if ($global:DashboardConfig.Resources.ContainsKey('AtlasForm')) { $global:DashboardConfig.Resources.Remove('AtlasForm') }
			$script:AtlasForm = $null
			foreach ($t in $Ctx.timers) { try { $t.Stop(); $t.Dispose() } catch { } }
			if ($script:AtlasToolTip -and -not $script:AtlasToolTip.IsDisposed)
			{
				try { $script:AtlasToolTip.Dispose() } catch { }
			}
			$script:AtlasToolTip = $null
			if ($script:AtlasNodeToolTip -and -not $script:AtlasNodeToolTip.IsDisposed)
			{
				try { $script:AtlasNodeToolTip.Dispose() } catch { }
			}
			$script:AtlasNodeToolTip = $null
			try { $canvas.Dispose() } catch { }
			try { $sidebar.Dispose() } catch { }
			try { $s.Dispose() } catch { }
			[GC]::Collect()
			[GC]::WaitForPendingFinalizers()
		}.GetNewClosure())

	$form.Add_Shown({
			& $applyTheme
			$stage.PerformLayout()
			$canvas.SetWorld($S.MapKey, $false)
			Sync-AtlasJump -Ctx $Ctx
			Sync-Atlas -Ctx $Ctx
			$canvas.InvalidateLayer()
			$canvas.Invalidate()
			Start-AtlasLoadNotice -Ctx $Ctx
		}.GetNewClosure())

	foreach ($g in (Add-AtlasResizeGrips -Form $form)) { $form.Controls.Add($g) }

	if (Test-DashboardCommand 'Register-OpenForm') { try { Register-OpenForm $form } catch { } }

	$form.Show() | Out-Null
	$canvas.Focus() | Out-Null
	return $form
}

function Close-Atlas
{
	if ($script:AtlasForm -and -not $script:AtlasForm.IsDisposed)
	{
		try { $script:AtlasForm.Close() | Out-Null } catch { }
	}
	$script:AtlasForm = $null
}

#endregion

Export-ModuleMember -Function *