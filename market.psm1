<# market.psm1 #>

#region Internal state

$script:MarketLocalDb = $null
$script:MarketLocalVersion = 0
$script:MarketDbPathOverride = ''
$script:MarketNamesCache = $null
$script:MarketCatalog = $null
$script:MarketCatalogSrcWriteUtc = [datetime]::MinValue
$script:MarketTicksCache = @{}
$script:MarketInitializing = $false
$script:MarketLastSuggestions = @()
$script:MarketLastQuery = ''
$script:MarketSuggestTimer = $null
$script:MarketIconCache = @{}
$script:MarketRecCache = @{}
$script:MarketIconMap = $null
$script:MarketIconMapPartial = $null
$script:MarketIconMapLoading = $false
$script:MarketIconMapPowerShell = $null
$script:MarketIconMapRunspace = $null
$script:MarketIconMapHandle = $null
$script:MarketIconMapTimer = $null
$script:MarketIconMapError = ''
$script:MarketDataRootCache = $null
$script:MarketIconBaseRootCache = @{}
$script:MarketVaultCacheUrl = ''
$script:MarketVaultCacheBytes = $null
$script:MarketVaultTempUrl = ''
$script:MarketItemImages = @{} 
$script:MarketHistCache = $null 
$script:MarketPickedName = ''
$script:MarketPickedTicks = [long]0
$script:MarketPickedPrice = 0.0
$script:MarketPickedFrom = ''
$script:MarketPickedSeller = ''
$script:MarketHistGridWired = $null
$script:MarketLastFilterText = $null

#endregion
#region Configuration helpers

function Get-MarketConfigValue
{
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [string]$Default = ''
    )
    try
    {
        $cfg = $global:DashboardConfig.Config
        if ($null -eq $cfg) { return $Default }
        if (-not $cfg.Contains('Market')) { return $Default }
        $sec = $cfg['Market']
        if ($null -eq $sec) { return $Default }
        if ($sec.Contains($Key))
        {
            $v = $sec[$Key]
            if ($null -ne $v) { return [string]$v }
        }
    }
    catch { }
    return $Default
}

function Set-MarketConfigValue
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [AllowEmptyString()]
        [string]$Value = ''
    )
    try
    {
        $cfg = $global:DashboardConfig.Config
        if ($null -eq $cfg) { return }
        if (-not $cfg.Contains('Market')) { $cfg['Market'] = [ordered]@{} }
        $sec = $cfg['Market']
        if ($null -eq $sec -or ($sec -isnot [System.Collections.IDictionary]))
        {
            $sec = [ordered]@{}
            $cfg['Market'] = $sec
        }
        $sec[$Key] = $Value
    }
    catch { Write-Verbose "MARKET: could not store config key '$Key': $($_.Exception.Message)" }
}

function Ensure-MarketConfig
{
    [CmdletBinding()]
    param()
    try
    {
        $cfg = $global:DashboardConfig.Config
        if ($null -eq $cfg) { return }
        if (-not $cfg.Contains('Market')) { $cfg['Market'] = [ordered]@{} }
        $sec = $cfg['Market']
        if ($null -eq $sec -or ($sec -isnot [System.Collections.IDictionary]))
        {
            $sec = [ordered]@{}
            $cfg['Market'] = $sec
        }
        if (-not $sec.Contains('VaultUrl')) { $sec['VaultUrl'] = '' }
        if (-not $sec.Contains('AccessCode')) { $sec['AccessCode'] = '' }
        if (-not $sec.Contains('TrendRange')) { $sec['TrendRange'] = '' }
    }
    catch { Write-Verbose "MARKET: could not seed config: $($_.Exception.Message)" }
}

#endregion
#region Local price journal

function Get-MarketLocalDbPath
{
    [CmdletBinding()]
    [OutputType([string])]
    param()
    if (-not [string]::IsNullOrWhiteSpace($script:MarketDbPathOverride)) { return $script:MarketDbPathOverride }
    $base = $env:APPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [string]$PWD.Path }
    return (Join-Path $base 'Entropia_Dashboard\market_local.json')
}

#region Item catalog

function Get-MarketDataRoot
{
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $base = $env:APPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [string]$PWD.Path }
    $key = "$base"
    $cache = $script:MarketDataRootCache
    if ($null -eq $cache) { $cache = @{}; $script:MarketDataRootCache = $cache }
    if ($cache.ContainsKey($key)) { return $cache[$key] }
    $root = Join-Path $base 'Entropia_Dashboard'
    $cache[$key] = $root
    return $root
}

function Get-MarketCatalogPath
{
    [CmdletBinding()]
    [OutputType([string])]
    param()
    return (Join-Path (Get-MarketDataRoot) 'market_catalog.json')
}

function Get-MarketCatalogShippedPath
{
    [CmdletBinding()]
    [OutputType([string])]
    param()
    try
    {
        $src = [string]$global:DashboardConfig.Paths.Source
        if (-not [string]::IsNullOrWhiteSpace($src))
        {
            $p = Join-Path $src 'market_catalog.json'
            if (Test-Path -LiteralPath $p -PathType Leaf) { return $p }
        }
    }
    catch { }
    return ''
}

function Get-MarketCatalogStamp
{
    [CmdletBinding()]
    param([string]$Path = '')

    $out = @{ Version = 0; Updated = '' }
    if ([string]::IsNullOrWhiteSpace($Path)) { return $out }
    try
    {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $out }
        $fs = [System.IO.File]::OpenRead($Path)
        try
        {
            $buf = [byte[]]::new(400)
            $n = $fs.Read($buf, 0, $buf.Length)
            $head = [System.Text.Encoding]::UTF8.GetString($buf, 0, $n)
        }
        finally { try { $fs.Dispose() } catch { } }
        $mv = [regex]::Match($head, '"version"\s*:\s*(\d+)')
        if ($mv.Success) { $out.Version = [int]$mv.Groups[1].Value }
        $mu = [regex]::Match($head, '"updated"\s*:\s*"([^"]*)"')
        if ($mu.Success) { $out.Updated = $mu.Groups[1].Value }
    }
    catch { }
    return $out
}

function Ensure-MarketCatalog
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $src = Get-MarketCatalogShippedPath
    if ([string]::IsNullOrWhiteSpace($src)) { return $false }
    $dst = Get-MarketCatalogPath
    try
    {
        if (-not (Test-Path -LiteralPath $dst -PathType Leaf))
        {
            $dir = Split-Path -Path $dst -Parent
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
            Copy-Item -LiteralPath $src -Destination $dst -Force
            return $true
        }
        if ([string]::Equals($src, $dst, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }

		try
        {
            $srcInfo = [System.IO.FileInfo]::new($src)
            $srcWrite = $srcInfo.LastWriteTimeUtc
            if ($script:MarketCatalogSrcWriteUtc -eq $srcWrite) { return $false }
            $script:MarketCatalogSrcWriteUtc = $srcWrite
        }
        catch { }

        $s = Get-MarketCatalogStamp -Path $src
        $d = Get-MarketCatalogStamp -Path $dst
        $newer = $false
        if ($s.Version -gt $d.Version) { $newer = $true }
        elseif ($s.Version -eq $d.Version -and -not [string]::IsNullOrWhiteSpace($s.Updated) -and
            [string]::CompareOrdinal($s.Updated, $d.Updated) -gt 0) { $newer = $true }
        if ($newer)
        {
            Copy-Item -LiteralPath $src -Destination $dst -Force
            return $true
        }
    }
    catch { Write-Verbose "MARKET: catalog copy failed: $($_.Exception.Message)" }
    return $false
}

function Get-MarketCatalogIndexPath
{
    [CmdletBinding()]
    [OutputType([string])]
    param()
    return (Join-Path (Get-MarketDataRoot) 'market_catalog.index')
}

function Read-MarketCatalogIndex
{
    [CmdletBinding()]
    param(
        [string]$Path = '',
        [int]$Version = 0,
        [string]$Updated = ''
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    try
    {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }

        if ('Custom.MarketCatalogIndex' -as [Type])
        {
            $iv = 0
            $iUpdated = $null
            if (-not [Custom.MarketCatalogIndex]::ReadHeader($Path, [ref]$iv, [ref]$iUpdated)) { return $null }
            if ($iv -ne $Version) { return $null }
            if (-not [string]::Equals($iUpdated, "$Updated", [System.StringComparison]::Ordinal)) { return $null }

            $cByName = $null
            $cNames = $null
            if (-not [Custom.MarketCatalogIndex]::Read($Path, [ref]$cByName, [ref]$cNames)) { return $null }
            return @{ Names = $cNames; ByName = $cByName }
        }

        $lines = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
        if ($lines.Count -lt 2) { return $null }
        $hdr = $lines[0].Split([char]9)
        if ($hdr.Length -lt 3 -or $hdr[0] -ne 'v2') { return $null }
        $iv = 0
        if (-not [int]::TryParse($hdr[1], [ref]$iv)) { return $null }
        if ($iv -ne $Version) { return $null }
        if (-not [string]::Equals($hdr[2], "$Updated", [System.StringComparison]::Ordinal)) { return $null }

        $byName = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $names = [System.Collections.Generic.List[string]]::new()
        for ($i = 1; $i -lt $lines.Count; $i++)
        {
            $ln = $lines[$i]
            if ([string]::IsNullOrWhiteSpace($ln)) { continue }
            $tab = $ln.IndexOf([char]9)
            if ($tab -lt 1) { continue }
            $n = $ln.Substring(0, $tab).Trim()
            if ($n -eq '') { continue }
            $rest = $ln.Substring($tab + 1)
            $conn = ''
            $desc = ''
            $tab2 = $rest.IndexOf([char]9)
            if ($tab2 -ge 0)
            {
                $conn = $rest.Substring(0, $tab2)
                $desc = $rest.Substring($tab2 + 1)
            }
            else
            {
                $conn = $rest
            }
            $nk = $n.ToLowerInvariant()
            if ($byName.ContainsKey($nk)) { continue }
            $byName[$nk] = [pscustomobject]@{ name = $n; connection = $conn; desc = $desc }
            [void]$names.Add($n)
        }
        if ($byName.Count -eq 0) { return $null }
        return @{ Names = $names.ToArray(); ByName = $byName }
    }
    catch { return $null }
}

function Write-MarketCatalogIndex
{
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$Path = '',
        [int]$Version = 0,
        [string]$Updated = '',
        $ByName = $null
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or $null -eq $ByName) { return $false }
    try
    {
        $sb = [System.Text.StringBuilder]::new(1048576)
        [void]$sb.Append('v2').Append([char]9).Append([string]$Version).Append([char]9).Append("$Updated").Append([char]10)
        foreach ($kv in $ByName.GetEnumerator())
        {
            $n = ''
            try { $n = ([string]$kv.Value.name).Trim() } catch { $n = '' }
            if ($n -eq '') { continue }
            $conn = ''
            try { $conn = [string]$kv.Value.connection } catch { $conn = '' }
            $desc = ''
            try { $desc = [string]$kv.Value.desc } catch { $desc = '' }
            $desc = ($desc -replace "[`t`r`n]", ' ')
            [void]$sb.Append($n).Append([char]9).Append($conn).Append([char]9).Append($desc).Append([char]10)
        }

        $tmpPath = "$Path.tmp"
        [System.IO.File]::WriteAllText($tmpPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
        if ([System.IO.File]::Exists($Path))
        {
            try { [System.IO.File]::Replace($tmpPath, $Path, $null) }
            catch { [System.IO.File]::Delete($Path); [System.IO.File]::Move($tmpPath, $Path) }
        }
        else
        {
            [System.IO.File]::Move($tmpPath, $Path)
        }
        return $true
    }
    catch { return $false }
}

function Get-MarketCatalog
{
    [CmdletBinding()]
    param([switch]$Force)

    $copied = $false
    try { $copied = [bool](Ensure-MarketCatalog) } catch { }
    if ($copied)
    {
        try { $script:MarketIconMap = $null } catch { }
        try { $script:MarketIconMapPartial = $null } catch { }
        try { $script:MarketNamesCache = $null } catch { }
        $script:MarketCatalog = $null
    }

    if (-not $Force -and $null -ne $script:MarketCatalog) { return $script:MarketCatalog }

    $cat = [pscustomobject]@{
        Version = 0
        Updated = ''
        ByName  = $null
        Names   = @()
    }

    $path = Get-MarketCatalogPath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf))
    {
        Write-Verbose "MARKET: no market_catalog.json at $path (item data unavailable)"
        $script:MarketCatalog = $cat
        return $cat
    }

    $stamp = Get-MarketCatalogStamp -Path $path
    $cat.Version = [int]$stamp.Version
    $cat.Updated = [string]$stamp.Updated

    $idx = Read-MarketCatalogIndex -Path (Get-MarketCatalogIndexPath) -Version $cat.Version -Updated $cat.Updated
    if ($null -ne $idx)
    {
        $cat.Names = $idx.Names
        $cat.ByName = $idx.ByName
    }

    if ($null -eq $cat.ByName)
    {
        try
        {
            $json = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            if ($null -ne $json -and $null -ne $json.items)
            {
                $byName = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $names = [System.Collections.Generic.List[string]]::new()
                foreach ($p in $json.items.PSObject.Properties)
                {
                    $rec = $p.Value
                    if ($null -eq $rec) { continue }
                    $n = ''
                    try { $n = [string]$rec.name } catch { $n = '' }
                    if ([string]::IsNullOrWhiteSpace($n)) { continue }
                    $conn = ''
                    try { $conn = [string]$rec.connection } catch { $conn = '' }
                    $desc = ''
                    try { $desc = [string]$rec.desc } catch { $desc = '' }
                    $n = $n.Trim()
                    $nk = $n.ToLowerInvariant()
                    if ($byName.ContainsKey($nk)) { continue }
                    $byName[$nk] = [pscustomobject]@{ name = $n; connection = $conn; desc = $desc }
                    [void]$names.Add($n)
                }
                $cat.ByName = $byName
                $cat.Names = $names.ToArray()
                try
                {
                    Write-MarketCatalogIndex -Path (Get-MarketCatalogIndexPath) -Version $cat.Version -Updated $cat.Updated -ByName $byName | Out-Null
                }
                catch { }
            }
        }
        catch { Write-Verbose "MARKET: catalog load failed: $($_.Exception.Message)" }

        if ($null -eq $cat.ByName)
        {
            $cat.ByName = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $cat.Names = @()
        }
    }

    $script:MarketCatalog = $cat
    return $cat
}

function Get-MarketIconBaseRoot
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Relative = '')

    $probe = ''
    $rel = "$Relative".Trim()
    if ($rel -ne '')
    {
        $norm = $rel.Replace('/', '\')
        $cut = $norm.IndexOf('\')
        if ($cut -ge 0) { $probe = $norm.Substring(0, $cut) }
        else { $probe = $norm }
        $probe = $probe.Trim()
    }

    $cache = $script:MarketIconBaseRootCache
    if ($null -eq $cache) { $cache = @{}; $script:MarketIconBaseRootCache = $cache }
    $key = "$probe".ToLowerInvariant()
    if ($cache.ContainsKey($key)) { return [string]$cache[$key] }

    $candidates = [System.Collections.Generic.List[string]]::new()
    try { [void]$candidates.Add((Get-MarketDataRoot)) } catch { }
    $src = ''
    try { $src = [string]$global:DashboardConfig.Paths.Source } catch { $src = '' }
    if (-not [string]::IsNullOrWhiteSpace($src)) { [void]$candidates.Add($src) }
    $up = ''
    try { $up = "$env:USERPROFILE" } catch { $up = '' }
    if (-not [string]::IsNullOrWhiteSpace($up))
    {
        [void]$candidates.Add((Join-Path $up 'Entropia_Dashboard\.main'))
        [void]$candidates.Add((Join-Path $up 'Entropia_Dashboard'))
    }

    $fallback = ''
    $found = ''
    foreach ($r in $candidates)
    {
        if ([string]::IsNullOrWhiteSpace($r)) { continue }
        if ($fallback -eq '') { $fallback = $r }
        if ($probe -eq '') { if ($found -eq '') { $found = $r }; continue }
        try
        {
            if (Test-Path -LiteralPath (Join-Path $r $probe) -PathType Container) { $found = $r; break }
        }
        catch { }
    }
    if ($found -eq '') { $found = $fallback }
    if ($found -eq '') { return '' }
    $cache[$key] = $found
    return $found
}

function Resolve-MarketIconPath
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Path = '')

    $p = "$Path".Trim()
    if ($p -eq '') { return '' }
    if ($p.IndexOf('%') -ge 0) { $p = [Environment]::ExpandEnvironmentVariables($p) }
    try
    {
        if (-not [System.IO.Path]::IsPathRooted($p))
        {
            $rel = $p.Replace('/', [string][System.IO.Path]::DirectorySeparatorChar)
            $base = ''
            try { $base = Get-MarketIconBaseRoot -Relative $p } catch { $base = '' }
            if ([string]::IsNullOrWhiteSpace($base)) { $base = Get-MarketDataRoot }
            $p = Join-Path $base $rel
        }
    }
    catch { }
    return $p
}

#region Remote catalog and icons

$script:MarketRemoteDefault = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/market'
$script:MarketManifest = $null
$script:MarketManifestTried = [datetime]::MinValue
$script:MarketAssetsChecked = $false
$script:MarketAssetSyncTimer = $null

function Get-MarketRemoteBase
{
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $r = Get-MarketConfigValue -Key 'Remote' -Default ''
    if (-not [string]::IsNullOrWhiteSpace($r)) { return $r.TrimEnd('/') }
    return $script:MarketRemoteDefault
}

function Format-MarketSize
{
    [CmdletBinding()]
    [OutputType([string])]
    param([long]$Bytes = 0)
    if ($Bytes -ge 1GB) { return ('{0:0.0} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:0.0} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:0} KB' -f ($Bytes / 1KB)) }
    return ('{0} B' -f $Bytes)
}

function Get-MarketWebText
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Url = '', [int]$Timeout = 20000)
    try
    {
        $cur = [System.Net.ServicePointManager]::SecurityProtocol
        [System.Net.ServicePointManager]::SecurityProtocol = $cur -bor [System.Net.SecurityProtocolType]::Tls12
    }
    catch { }
    $resp = $null; $sr = $null
    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Timeout = $Timeout
    $req.ReadWriteTimeout = $Timeout
    $req.UserAgent = 'EntropiaDashboard'
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

function Get-MarketFileHash
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Path = '')
    $sha = $null; $fs = $null
    try
    {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $bytes = $sha.ComputeHash($fs)
        return (-join ($bytes | ForEach-Object { $_.ToString('x2') }))
    }
    catch { return '' }
    finally
    {
        if ($fs) { $fs.Dispose() }
        if ($sha) { $sha.Dispose() }
    }
}

function Invoke-MarketDownloadFile
{
    [CmdletBinding()]
    [OutputType([bool])]
    param([string]$Url = '', [string]$Path = '', [string]$Hash = '')
    try
    {
        $cur = [System.Net.ServicePointManager]::SecurityProtocol
        [System.Net.ServicePointManager]::SecurityProtocol = $cur -bor [System.Net.SecurityProtocolType]::Tls12
    }
    catch { }
    $dir = Split-Path -Path $Path -Parent
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
            $got = Get-MarketFileHash -Path $part
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

function Get-MarketManifest
{
    [CmdletBinding()]
    param([switch]$Force)
    $tried = $script:MarketManifestTried
    $age = ((Get-Date) - $tried).TotalMinutes
    if (-not $Force -and $tried -gt [datetime]::MinValue -and $age -lt 10) { return $script:MarketManifest }
    $script:MarketManifestTried = Get-Date
    $script:MarketManifest = $null
    $url = (Get-MarketRemoteBase) + '/manifest.json'
    try
    {
        $txt = Get-MarketWebText -Url $url -Timeout 10000
        if ($txt)
        {
            $obj = $txt | ConvertFrom-Json
            if ($null -ne $obj -and ($obj.catalog -or $obj.icons)) { $script:MarketManifest = $obj }
        }
    }
    catch { $script:MarketManifest = $null }
    return $script:MarketManifest
}

function Read-MarketPrompt
{
    [CmdletBinding()]
    [OutputType([bool])]
    param([string]$Text = '', [string]$Caption = 'Entropia Market')
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

function Show-MarketNotice
{
    [CmdletBinding()]
    param([string]$Text = '', [string]$Caption = 'Entropia Market', [string]$Icon = 'Information')
    if (Test-DashboardCommand 'Show-DarkMessageBox')
    {
        try { Show-DarkMessageBox -Text $Text -Caption $Caption -Buttons 'Ok' -Icon $Icon | Out-Null; return } catch { }
    }
    try { [void][System.Windows.Forms.MessageBox]::Show($Text, $Caption, 'OK', $Icon) } catch { }
}

function Invoke-MarketDownload
{
    [CmdletBinding()]
    param([string]$Title = 'Entropia Market', [object[]]$Items = @())
    if ($Items.Count -eq 0) { return @{ Failed = (New-Object System.Collections.Generic.List[string]); Cancelled = $false; Done = 0 } }
    if (Test-DashboardCommand 'Invoke-AtlasDownload')
    {
        try { return (Invoke-AtlasDownload -Title $Title -Items $Items) } catch { }
    }
    $failed = New-Object System.Collections.Generic.List[string]
    $done = 0
    foreach ($it in $Items)
    {
        try
        {
            Invoke-MarketDownloadFile -Url ([string]$it.Url) -Path ([string]$it.Path) -Hash ([string]$it.Hash) | Out-Null
            $done++
        }
        catch { $failed.Add([string]$it.Path) }
    }
    return @{ Failed = $failed; Cancelled = $false; Done = $done }
}

function Get-MarketIconsState
{
    [CmdletBinding()]
    param([string]$Root = '')
    if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Get-MarketDataRoot }
    $p = Join-Path (Join-Path $Root 'icons') '.market-icons.json'
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    try { return (Get-Content -LiteralPath $p -Raw | ConvertFrom-Json) } catch { return $null }
}

function Save-MarketIconsState
{
    [CmdletBinding()]
    param([string]$Root = '', [string]$Version = '')
    try
    {
        if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Get-MarketDataRoot }
        $dir = Join-Path $Root 'icons'
        if (-not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
        $txt = ([ordered]@{ version = [string]$Version; checked = (Get-Date).ToString('s') } | ConvertTo-Json -Compress)
        Set-Content -LiteralPath (Join-Path $dir '.market-icons.json') -Value $txt -Encoding UTF8
    }
    catch { }
}

function Get-MarketIconDiff
{
    [CmdletBinding()]
    param($Manifest = $null)
    $root = Get-MarketDataRoot
    $iconsDir = Join-Path $root 'icons'
    $rootRel = 'icons'
    if ($Manifest -and $Manifest.icons -and $Manifest.icons.root) { $rootRel = [string]$Manifest.icons.root }

    $out = @{
        Missing = 0
        Changed = 0
        Total   = 0
        Bytes   = [long]0
        Files   = (New-Object System.Collections.Generic.List[object])
    }
    if ($null -eq $Manifest -or $null -eq $Manifest.icons) { return $out }

    foreach ($f in @($Manifest.icons.files))
    {
        if ($null -eq $f -or -not $f.p) { continue }
        $out.Total++
        $rel = [string]$f.p
        $want = [string]$f.h
        $size = [long]0
        if ($null -ne $f.s) { $size = [long]$f.s }

        $probe = $rootRel + '/' + $rel
        $have = ''
        try { $have = Resolve-MarketIconPath -Path $probe } catch { $have = '' }
        if ([string]::IsNullOrWhiteSpace($have) -or -not (Test-Path -LiteralPath $have -PathType Leaf))
        {
            $out.Missing++
            $out.Bytes += $size
            $out.Files.Add(@{ Rel = $rel; Path = (Join-Path $iconsDir ($rel -replace '/', '\')); Size = $size; Hash = $want })
            continue
        }
        if (-not $want) { continue }
        $got = Get-MarketFileHash -Path $have
        if ($got -and $got -ne $want)
        {
            $out.Changed++
            $out.Bytes += $size
            $out.Files.Add(@{ Rel = $rel; Path = (Join-Path $iconsDir ($rel -replace '/', '\')); Size = $size; Hash = $want })
        }
    }
    return $out
}

function Test-MarketStartupComplete
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    try
    {
        if ($global:DashboardConfig -and $global:DashboardConfig.State -and $global:DashboardConfig.State.Contains('StartupSplashClosed'))
        {
            return [bool]$global:DashboardConfig.State['StartupSplashClosed']
        }
    }
    catch { }
    try
    {
        if ($global:DashboardConfig -and $global:DashboardConfig.UI -and $global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.Visible)
        {
            return $true
        }
    }
    catch { }
    try { return [bool][System.Windows.Forms.Application]::MessageLoopRunning } catch { return $true }
}

function Invoke-MarketAssetSync
{
    [CmdletBinding()]
    param()
    if ($script:MarketAssetsChecked) { return }
    if (-not (Test-MarketStartupComplete))
    {
        Start-MarketAssetSync
        return
    }

    $script:MarketAssetsChecked = $true
    try
    {
        if (Sync-MarketAssets)
        {
            try { Stop-MarketIconMapLoad } catch { }
            $script:MarketCatalog = $null
            $script:MarketNamesCache = $null
            $script:MarketIconMap = $null
            $script:MarketIconMapPartial = $null
            try
            {
                if ($null -ne (Get-MarketUIControl -Name 'MarketSearchInput'))
                {
                    try { Start-MarketIconMapLoad } catch { }
                    try { Set-MarketIcons } catch { }
                    try { Update-MarketView } catch { }
                }
            }
            catch { }
        }
    }
    catch { Write-Verbose "MARKET: asset sync failed: $($_.Exception.Message)" }
}

function Start-MarketAssetSync
{
    [CmdletBinding()]
    param()
    if ($script:MarketAssetsChecked -or $null -ne $script:MarketAssetSyncTimer) { return }
    if (Test-MarketStartupComplete)
    {
        Invoke-MarketAssetSync
        return
    }

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $timer.Add_Tick({
            try
            {
                if (-not (Test-MarketStartupComplete)) { return }
                $t = $script:MarketAssetSyncTimer
                $script:MarketAssetSyncTimer = $null
                if ($null -ne $t)
                {
                    try { $t.Stop() } catch { }
                    try { $t.Dispose() } catch { }
                }
                try
                {
                    if ($global:DashboardConfig -and $global:DashboardConfig.Resources -and $global:DashboardConfig.Resources.Timers)
                    {
                        [void]$global:DashboardConfig.Resources.Timers.Remove('MarketAssetSync')
                    }
                }
                catch { }
                Invoke-MarketAssetSync
            }
            catch { Write-Verbose "MARKET: deferred asset sync failed: $($_.Exception.Message)" }
        })
    $script:MarketAssetSyncTimer = $timer
    try { $global:DashboardConfig.Resources.Timers['MarketAssetSync'] = $timer } catch { }
    $timer.Start()
}

function Sync-MarketAssets
{
    [CmdletBinding()]
    [OutputType([bool])]
    param([switch]$Force)

    if (-not (Test-MarketStartupComplete))
    {
        Start-MarketAssetSync
        return $false
    }

    $manifest = Get-MarketManifest -Force:$Force
    if ($null -eq $manifest) { return $false }
    $base = Get-MarketRemoteBase
    $changed = $false

    $catRel = 'market_catalog.json'
    if ($manifest.catalog -and $manifest.catalog.path) { $catRel = [string]$manifest.catalog.path }
    $catUrl = $base + '/' + $catRel
    $catPath = Get-MarketCatalogPath
    $catSize = [long]0
    if ($manifest.catalog -and $null -ne $manifest.catalog.size) { $catSize = [long]$manifest.catalog.size }
    $catWant = ''
    if ($manifest.catalog -and $manifest.catalog.sha256) { $catWant = [string]$manifest.catalog.sha256 }
    $catDate = ''
    if ($manifest.catalog -and $manifest.catalog.updated) { $catDate = [string]$manifest.catalog.updated }

    if ($manifest.catalog)
    {
        $missing = -not (Test-Path -LiteralPath $catPath -PathType Leaf)
        $stale = $false
        if (-not $missing -and $catWant)
        {
            $got = Get-MarketFileHash -Path $catPath
            if ($got -and $got -ne $catWant) { $stale = $true }
        }
        if ($missing -or $stale)
        {
            if ($missing)
            {
                $text = "The market item list is not installed yet.`n`nDownload it now? ($(Format-MarketSize $catSize))"
                $cap = 'Entropia Market - Item list'
                $title = 'Entropia Market - downloading the item list'
            }
            else
            {
                $when = ''
                if ($catDate) { $when = ' (updated ' + $catDate + ')' }
                $text = "A newer market item list is available$when.`n`nDownload it now? ($(Format-MarketSize $catSize))"
                $cap = 'Entropia Market - Update'
                $title = 'Entropia Market - updating the item list'
            }
            if (Read-MarketPrompt -Text $text -Caption $cap)
            {
                $r = Invoke-MarketDownload -Title $title -Items @(@{ Url = $catUrl; Path = $catPath; Size = $catSize; Hash = $catWant })
                if ($r.Done -gt 0)
                {
                    $changed = $true
                    $script:MarketCatalog = $null
                    $script:MarketNamesCache = $null
                    $script:MarketIconMap = $null
                    $script:MarketIconMapPartial = $null
                }
                elseif ($r.Failed.Count -gt 0)
                {
                    Show-MarketNotice -Text 'The market item list could not be downloaded. Check your connection and open the market again.' -Caption 'Entropia Market - Download failed' -Icon 'Error'
                }
            }
        }
    }

    $iconFiles = @()
    if ($manifest.icons -and $manifest.icons.files) { $iconFiles = @($manifest.icons.files) }
    if ($iconFiles.Count -gt 0)
    {
        $remoteVer = ''
        if ($manifest.icons.version) { $remoteVer = [string]$manifest.icons.version }
        $state = Get-MarketIconsState
        $haveVer = ''
        if ($state -and $state.version) { $haveVer = [string]$state.version }

        $scan = $true
        if ($haveVer -and $remoteVer -and $haveVer -eq $remoteVer) { $scan = $false }
        if ($scan)
        {
            $diff = Get-MarketIconDiff -Manifest $manifest
            if ($diff.Files.Count -eq 0)
            {
                Save-MarketIconsState -Version $remoteVer
            }
            else
            {
                if ($diff.Missing -eq $diff.Files.Count) { $what = 'The market icons are not installed.' }
                else
                {
                    $bits = @()
                    if ($diff.Missing -gt 0) { $bits += ('{0} missing' -f $diff.Missing) }
                    if ($diff.Changed -gt 0) { $bits += ('{0} changed' -f $diff.Changed) }
                    $what = 'Market icons: ' + ($bits -join ', ') + '.'
                }
                $count = $diff.Files.Count
                $word = if ($count -eq 1) { 'icon' } else { 'icons' }
                $text = "$what`n`nDownload $count ${word}? ($(Format-MarketSize $diff.Bytes))`n`nThe market works without them - rows are then listed without pictures."
                if (Read-MarketPrompt -Text $text -Caption 'Entropia Market - Icons')
                {
                    $rootRel = 'icons'
                    if ($manifest.icons.root) { $rootRel = [string]$manifest.icons.root }
                    $items = New-Object System.Collections.Generic.List[object]
                    foreach ($f in $diff.Files)
                    {
                        $items.Add(@{ Url = ($base + '/' + $rootRel + '/' + $f.Rel); Path = $f.Path; Size = $f.Size; Hash = [string]$f.Hash })
                    }
                    $r = Invoke-MarketDownload -Title 'Entropia Market - downloading icons' -Items $items.ToArray()
                    if ($r.Failed.Count -eq 0 -and -not $r.Cancelled)
                    {
                        Save-MarketIconsState -Version $remoteVer
                        $changed = $true
                    }
                    elseif ($r.Failed.Count -gt 0)
                    {
                        $fw = if ($r.Failed.Count -eq 1) { 'icon' } else { 'icons' }
                        Show-MarketNotice -Text ('{0} {1} could not be downloaded. They will be offered again the next time the market is opened.' -f $r.Failed.Count, $fw) -Caption 'Entropia Market - Download failed' -Icon 'Error'
                    }
                }
            }
        }
    }

    if ($changed)
    {
        try { $script:MarketIconBaseRootCache = @{} } catch { }
        try { $script:MarketIconCache.Clear() } catch { $script:MarketIconCache = @{} }
    }
    return $changed
}

#endregion

function ConvertTo-MarketNumber
{
    [CmdletBinding()]
    [OutputType([double])]
    param(
        $Value,
        [double]$Default = 0.0
    )
    if ($null -eq $Value) { return $Default }

    if ($Value -is [string])
    {
        $s = $Value.Trim()
        if ($s -eq '') { return $Default }
        $d = 0.0
        if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
        if ([double]::TryParse($s, [ref]$d)) { return $d }
        return $Default
    }

    try { return [double]$Value }
    catch { return $Default }
}

function ConvertTo-MarketPriceEntry
{
    [CmdletBinding()]
    param(
        $InputObject,
        [string]$DefaultSource = 'online'
    )
    if ($null -eq $InputObject) { return $null }

    $scalar = ($InputObject -is [double]) -or ($InputObject -is [int]) -or ($InputObject -is [decimal]) -or
        ($InputObject -is [long]) -or ($InputObject -is [float]) -or ($InputObject -is [int64]) -or
        ($InputObject -is [uint64]) -or ($InputObject -is [int16]) -or ($InputObject -is [uint16])
    if ($scalar)
    {
        $v = ConvertTo-MarketNumber -Value $InputObject
        if ($v -gt 0)
        {
            return [ordered]@{ price = $v; ts = ''; source = $DefaultSource; user = ''; currency = '' }
        }
        return $null
    }

    $props = $null
    try { $props = $InputObject.PSObject.Properties } catch { return $null }
    if ($null -eq $props) { return $null }

    $price = 0.0
    foreach ($k in @('price', 'value', 'amount', 'avg', 'average'))
    {
        $p = $props[$k]
        if ($null -eq $p -or $null -eq $p.Value) { continue }
        $v = ConvertTo-MarketNumber -Value $p.Value
        if ($v -gt 0) { $price = $v; break }
    }
    if ($price -le 0) { return $null }

    $ts = ''
    foreach ($k in @('ts', 'time', 'date', 'when'))
    {
        $p = $props[$k]
        if ($null -ne $p -and $p.Value) { $ts = [string]$p.Value; break }
    }
    $source = $DefaultSource
    $sp = $props['source']
    if ($null -ne $sp -and $sp.Value) { $source = [string]$sp.Value }
    $user = ''
    foreach ($k in @('user', 'seller', 'sellerName', 'seller_name', 'owner', 'player'))
    {
        $up = $props[$k]
        if ($null -ne $up -and $up.Value) { $user = [string]$up.Value; break }
    }

    $currency = ''
    foreach ($k in @('currency', 'cur'))
    {
        $cp = $props[$k]
        if ($null -ne $cp -and $cp.Value) { $currency = ([string]$cp.Value).Trim(); break }
    }

    return [ordered]@{ price = $price; ts = $ts; source = $source; user = $user; currency = $currency }
}

function Get-MarketLocalData
{
    [CmdletBinding()]
    param([switch]$Force)

    if (-not $Force -and $null -ne $script:MarketLocalDb) { return $script:MarketLocalDb }

    $db = [ordered]@{ version = 1; updated = ''; hidden = [System.Collections.ArrayList]::new(); items = [ordered]@{} }
    $path = Get-MarketLocalDbPath
    if (Test-Path -LiteralPath $path)
    {
        try
        {
            $raw = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
            if (-not [string]::IsNullOrWhiteSpace($raw))
            {
                $json = $raw | ConvertFrom-Json
                if ($null -ne $json)
                {
                    $jProps = $json.PSObject.Properties
                    $vp = $jProps['version']
                    if ($null -ne $vp -and $null -ne $vp.Value)
                    {
                        $v = 1
                        if ([int]::TryParse("$($vp.Value)", [ref]$v)) { $db.version = $v }
                    }
                    $up = $jProps['updated']
                    if ($null -ne $up -and $up.Value) { $db.updated = [string]$up.Value }

                    $itemsNode = $null
                    $ip = $jProps['items']
                    if ($null -ne $ip -and $null -ne $ip.Value) { $itemsNode = $ip.Value }
                    else { $itemsNode = $json }

                    if ($null -ne $itemsNode)
                    {
                        $iProps = $null
                        try { $iProps = $itemsNode.PSObject.Properties } catch { $iProps = $null }
                        if ($null -ne $iProps)
                        {
                            foreach ($p in $iProps)
                            {
                                if ($p.Name -in @('version', 'updated', 'items')) { continue }
                                if ($null -eq $p.Value) { continue }

                                $prices = [System.Collections.ArrayList]::new()
                                $node = $p.Value
                                $priceNode = $null
                                if ($node -is [System.Array]) { $priceNode = $node }
                                else
                                {
                                    $nProps = $null
                                    try { $nProps = $node.PSObject.Properties } catch { $nProps = $null }
                                    if ($null -ne $nProps)
                                    {
                                        $pp = $nProps['prices']
                                        if ($null -ne $pp -and $null -ne $pp.Value) { $priceNode = $pp.Value }
                                    }
                                }
                                if ($null -ne $priceNode)
                                {
                                    foreach ($e in @($priceNode))
                                    {
                                        if ($null -eq $e) { continue }

                                        $entry = $null
                                        try
                                        {
                                            $raw = $e.price
                                            $price = 0.0
                                            if ($raw -is [double]) { $price = $raw }
                                            elseif ($null -ne $raw)
                                            {
                                                $s = "$raw".Trim()
                                                if ($s -ne '')
                                                {
                                                    if (-not [double]::TryParse($s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$price))
                                                    {
                                                        if (-not [double]::TryParse($s, [ref]$price)) { $price = 0.0 }
                                                    }
                                                }
                                            }

                                            if ($price -gt 0)
                                            {
                                                $ts = ''
                                                $tp = $e.ts
                                                if ($null -ne $tp) { $ts = [string]$tp }
                                                $entry = [ordered]@{ price = $price; ts = $ts }
                                                $cu = ''
                                                try
                                                {
                                                    $cp = $e.currency
                                                    if ($null -ne $cp) { $cu = ([string]$cp).Trim() }
                                                }
                                                catch { $cu = '' }
                                                if ($cu -ne '') { $entry['currency'] = $cu }
                                            }
                                        }
                                        catch { $entry = $null }

                                        if ($null -eq $entry)
                                        {
                                            try
                                            {
                                                $fe = ConvertTo-MarketPriceEntry -InputObject $e -DefaultSource 'local'
                                                if ($null -ne $fe -and [double]$fe['price'] -gt 0)
                                                {
                                                    $entry = [ordered]@{ price = [double]$fe['price']; ts = [string]$fe['ts'] }
                                                    $cu = ''
                                                    try { $cu = ([string]$fe['currency']).Trim() } catch { $cu = '' }
                                                    if ($cu -ne '') { $entry['currency'] = $cu }
                                                }
                                            }
                                            catch { $entry = $null }
                                        }
                                        if ($null -ne $entry) { [void]$prices.Add($entry) }
                                    }
                                }
                                $db.items[$p.Name] = [ordered]@{ prices = $prices }
                            }
                        }

                        $hp = $jProps['hidden']
                        if ($null -ne $hp -and $null -ne $hp.Value)
                        {
                            foreach ($h in @($hp.Value))
                            {
                                $hk = ''
                                try { $hk = ([string]$h).Trim() } catch { $hk = '' }
                                if ($hk -ne '' -and -not $db.hidden.Contains($hk)) { [void]$db.hidden.Add($hk) }
                            }
                        }
                    }
                }
            }
        }
        catch { Write-Verbose "MARKET: local journal load failed: $($_.Exception.Message)" }
    }

    $script:MarketLocalDb = $db
    $script:MarketLocalVersion++
    return $db
}

function Save-MarketLocalData
{
    [CmdletBinding()]
    [OutputType([bool])]
    param($Data)
    if ($null -eq $Data) { $Data = $script:MarketLocalDb }
    if ($null -eq $Data) { return $false }

    $path = Get-MarketLocalDbPath
    $dir = Split-Path -Path $path -Parent
    try
    {
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
        $out = [ordered]@{ version = 1; updated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); hidden = @(); items = [ordered]@{} }

        try
        {
            $hidSrc = $null
            try { $hidSrc = $Data.hidden } catch { $hidSrc = $null }
            $hid = @()
            foreach ($h in @($hidSrc))
            {
                $hk = ''
                try { $hk = ([string]$h).Trim() } catch { $hk = '' }
                if ($hk -ne '') { $hid += $hk }
            }
            $out.hidden = $hid
        }
        catch { }

        foreach ($k in @($Data.items.Keys))
        {
            $item = $Data.items[$k]
            $prices = [System.Collections.Generic.List[object]]::new()
            if ($null -ne $item -and $null -ne $item.prices)
            {
                foreach ($e in $item.prices)
                {
                    if ($null -eq $e) { continue }
                    $row = [ordered]@{
                        price = [double]$e['price']
                        ts    = [string]$e['ts']
                    }
                    $cu = ''
                    try
                    {
                        $rawCur = $e['currency']
                        if ($null -ne $rawCur) { $cu = ([string]$rawCur).Trim() }
                    }
                    catch { $cu = '' }
                    if ($cu -ne '') { $row['currency'] = $cu }
                    [void]$prices.Add($row)
                }
            }
            $out.items[$k] = [ordered]@{ prices = $prices }
        }
        $json = ConvertTo-Json -InputObject $out -Depth 6
        [System.IO.File]::WriteAllText($path, $json, [System.Text.UTF8Encoding]::new($false))
        $script:MarketRecCache = @{}
        $script:MarketLocalVersion++
        return $true
    }
    catch
    {
        Write-Verbose "MARKET: local journal save failed: $($_.Exception.Message)"
        return $false
    }
}

function Find-MarketDictKey
{
    [CmdletBinding()]
    [OutputType([string])]
    param($Dict, [string]$Name)
    if ($null -eq $Dict) { return $null }
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $n = $Name.Trim()
    try
    {
        foreach ($k in $Dict.Keys)
        {
            if ([string]::Equals([string]$k, $n, [System.StringComparison]::OrdinalIgnoreCase)) { return [string]$k }
        }
    }
    catch { return $null }
    return $null
}

function ConvertTo-MarketStringList
{
    [CmdletBinding()]
    param($Value)

    $out = [System.Collections.Generic.List[string]]::new()
    if ($null -ne $Value)
    {
        if ($Value -is [string])
        {
            if ($Value -ne '') { $out.Add($Value) }
        }
        else
        {
            foreach ($item in @($Value))
            {
                if ($null -eq $item) { continue }
                if ($item -is [string])
                {
                    if ($item -ne '') { $out.Add($item) }
                }
                elseif ($item -is [System.Array])
                {
                    foreach ($sub in $item) { if ($sub -is [string] -and $sub -ne '') { $out.Add($sub) } }
                }
            }
        }
    }
    return ,($out.ToArray())
}

function Get-MarketItemNames
{
    [CmdletBinding()]
    param([switch]$IncludeOnline)

    $db = Get-MarketLocalData
    $cat = $null
    try { $cat = Get-MarketCatalog } catch { $cat = $null }

    $online = $null
    if ($IncludeOnline)
    {
        try
        {
            if ((Test-MarketAccess) -and (Test-DashboardCommand 'Get-MarketOnlineNames'))
            {
                $online = Get-MarketOnlineNames
            }
        }
        catch { $online = $null }
    }

    $c = $script:MarketNamesCache
    if ($null -ne $c -and
        $c.Version -eq $script:MarketLocalVersion -and
        [object]::ReferenceEquals($c.Catalog, $cat) -and
        [object]::ReferenceEquals($c.Online, $online) -and
        $c.Inc -eq [bool]$IncludeOnline)
    {
        return ,($c.Names)
    }

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $names = [System.Collections.Generic.List[string]]::new()
    if ($null -ne $db -and $null -ne $db.items)
    {
        foreach ($k in $db.items.Keys)
        {
            $s = [string]$k
            if ($seen.Add($s)) { $names.Add($s) }
        }
    }
    if ($null -ne $cat -and $null -ne $cat.Names)
    {
        foreach ($s in $cat.Names)
        {
            if ([string]::IsNullOrWhiteSpace($s)) { continue }
            if ($seen.Add($s)) { $names.Add($s) }
        }
    }
    if ($null -ne $online)
    {
        foreach ($s in @($online))
        {
            if ([string]::IsNullOrWhiteSpace($s)) { continue }
            if ($seen.Add($s)) { $names.Add($s) }
        }
    }

    $arr = $names.ToArray()
    $lower = [string[]]::new($arr.Length)
    for ($i = 0; $i -lt $arr.Length; $i++) { $lower[$i] = $arr[$i].ToLowerInvariant() }

    $script:MarketNamesCache = @{
        Version = $script:MarketLocalVersion
        Catalog = $cat
        Online  = $online
        Inc     = [bool]$IncludeOnline
        Names   = $arr
        Lower   = $lower
    }
    return ,($script:MarketNamesCache.Names)
}

function Get-MarketTicks
{
    [CmdletBinding()]
    [OutputType([long])]
    param([string]$Ts = '')

    if ([string]::IsNullOrWhiteSpace($Ts)) { return [long]0 }
    if ($script:MarketTicksCache.ContainsKey($Ts)) { return [long]$script:MarketTicksCache[$Ts] }
    $tk = [long]0
    try { $tk = ([datetime]::Parse($Ts, [System.Globalization.CultureInfo]::InvariantCulture)).Ticks }
    catch { $tk = [long]0 }
    if ($script:MarketTicksCache.Count -gt 20000) { $script:MarketTicksCache.Clear() }
    $script:MarketTicksCache[$Ts] = $tk
    return $tk
}

function Get-MarketItemEntries
{
    [CmdletBinding()]
    param(
        [string]$Name,
        [switch]$LocalOnly
    )

    $out = [System.Collections.ArrayList]::new()
    if (-not [string]::IsNullOrWhiteSpace($Name))
    {
        $db = Get-MarketLocalData
        $key = Find-MarketDictKey -Dict $db.items -Name $Name
        if ($null -ne $key)
        {
            $prices = $db.items[$key]['prices']
            if ($null -ne $prices)
            {
                foreach ($e in $prices) { if ($null -ne $e) { [void]$out.Add($e) } }
            }
        }
    }

    if (-not $LocalOnly -and (Test-MarketAccess) -and (Test-DashboardCommand 'Get-MarketOnlineEntries'))
    {
        $sync = $null
        try { $sync = ConvertTo-MarketEntryList -Value (Get-MarketOnlineEntries -Name $Name) } catch { $sync = $null }
        foreach ($e in @($sync)) { if ($null -ne $e) { [void]$out.Add($e) } }
    }

    $sorted = @()
    if ($out.Count -gt 0)
    {
        $pairs = [System.Collections.Generic.List[object]]::new()
        $rank = 0
        foreach ($e in $out)
        {
            $s = ''
            try { $s = [string]$e['ts'] } catch { $s = '' }
            $tk = Get-MarketTicks -Ts $s
            $pairs.Add([pscustomobject]@{ Key = ('{0:D19}|{1:D6}' -f $tk, $rank); E = $e })
            $rank++
        }
        $sorted = @($pairs | Sort-Object -Property Key | ForEach-Object { $_.E })
    }
    return ,$sorted
}

function ConvertTo-MarketEntryList
{
    [CmdletBinding()]
    param($Value)

    $out = [System.Collections.Generic.List[object]]::new()
    $stack = [System.Collections.Generic.Stack[object]]::new()
    if ($null -ne $Value) { $stack.Push($Value) }
    while ($stack.Count -gt 0)
    {
        $cur = $stack.Pop()
        if ($null -eq $cur) { continue }
        $leaf = $false
        if ($cur -is [string] -or $cur -is [System.Collections.IDictionary]) { $leaf = $true }
        elseif ($cur -isnot [System.Collections.IEnumerable]) { $leaf = $true }
        if (-not $leaf)
        {
            $items = @()
            try { $items = @($cur) } catch { $items = @() }
            if ($items.Count -eq 0) { continue }
            for ($i = $items.Count - 1; $i -ge 0; $i--) { $stack.Push($items[$i]) }
            continue
        }
        $out.Add($cur)
    }
    return ,($out.ToArray())
}

function Add-MarketPrice
{
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [double]$Price,
        [string]$Source = 'local',
        [string]$User = '',
        [string]$Currency = ''
    )

    $n = "$Name".Trim()
    if ($n -eq '') { return $false }
    if ($Price -le 0) { return $false }

    $db = Get-MarketLocalData
    $key = Find-MarketDictKey -Dict $db.items -Name $n
    if ($null -eq $key) { $key = $n }
    if ($null -eq $db.items[$key])
    {
        $db.items[$key] = [ordered]@{ prices = [System.Collections.ArrayList]::new() }
    }
    $prices = $db.items[$key]['prices']
    if ($null -eq $prices -or $prices -isnot [System.Collections.ArrayList])
    {
        $prices = [System.Collections.ArrayList]::new()
        $db.items[$key]['prices'] = $prices
    }

    $entry = [ordered]@{
        price = [double]$Price
        ts    = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
    $cu = "$Currency".Trim()
    if ($cu -ne '') { $entry['currency'] = $cu }
    [void]$prices.Add($entry)
    while ($prices.Count -gt 200) { $prices.RemoveAt(0) }

    [void](Save-MarketLocalData -Data $db)
    return $true
}

#endregion
#region Fuzzy autocomplete

function Get-MarketSuggestions
{
    [CmdletBinding()]
    param(
        [string]$Query,
        [int]$Limit = 50,
        [switch]$IncludeOnline
    )

    $q = "$Query".Trim()
    if ($q -eq '') { return ,@() }
    if ($Limit -lt 1) { $Limit = 8 }

    $names = $null
    try { $names = Get-MarketItemNames -IncludeOnline:$IncludeOnline }
    catch { $names = $null }
    if ($null -eq $names) { return ,@() }
    if ($names -isnot [array]) { $names = @($names) }
    if ($names.Count -eq 0) { return ,@() }

    $lower = $null
    $c = $script:MarketNamesCache
    if ($null -ne $c -and $null -ne $c.Lower -and [object]::ReferenceEquals($c.Names, $names)) { $lower = $c.Lower }

    $ql = $q.ToLowerInvariant()
    $qlen = $ql.Length
    $qch = $ql.ToCharArray()

    $subScores = [int[]]::new($names.Count)
    $subCount = 0
    for ($ni = 0; $ni -lt $names.Count; $ni++)
    {
        $n = $names[$ni]
        if ([string]::IsNullOrWhiteSpace($n)) { $subScores[$ni] = -1; continue }
        $tl = $null
        if ($null -ne $lower) { $tl = $lower[$ni] }
        if ($null -eq $tl) { $tl = "$n".ToLowerInvariant() }
        if ($tl.Length -eq 0) { $subScores[$ni] = -1; continue }

        if ($tl -ceq $ql)
        {
            $subScores[$ni] = 1000000
            $subCount++
            continue
        }

        $idx = $tl.IndexOf($ql, [System.StringComparison]::Ordinal)
        if ($idx -lt 0) { $subScores[$ni] = -1; continue }

        $s = 10000 - $idx
        if ($idx -eq 0) { $s += 5000 }
        elseif ($tl[$idx - 1] -eq ' ' -or $tl[$idx - 1] -eq '-') { $s += 2500 }
        $pen = ($tl.Length - $qlen) * 5
        if ($pen -gt 500) { $pen = 500 }
        $s += (500 - $pen)
        $subScores[$ni] = $s
        $subCount++
    }

    $hits = [System.Collections.Generic.List[object]]::new()

    for ($ni = 0; $ni -lt $names.Count; $ni++)
    {
        $n = $names[$ni]
        if ([string]::IsNullOrWhiteSpace($n)) { continue }

        $s = $subScores[$ni]
        if ($s -lt 0)
        {
            if ($subCount -ge $Limit) { continue }

            $tl = $null
            if ($null -ne $lower) { $tl = $lower[$ni] }
            if ($null -eq $tl) { $tl = "$n".ToLowerInvariant() }
            if ($tl.Length -eq 0) { continue }

            $s = 100
            $ti = 0
            $consec = 0
            foreach ($ch in $qch)
            {
                if ($ti -ge $tl.Length) { $s = -1; break }
                $found = $tl.IndexOf($ch, $ti)
                if ($found -lt 0) { $s = -1; break }
                if ($found -eq $ti -and $ti -gt 0) { $consec++ ; $s += (20 * $consec) }
                else { $consec = 0 }
                if ($found -eq 0) { $s += 30 }
                elseif ($tl[$found - 1] -eq ' ' -or $tl[$found - 1] -eq '-') { $s += 15 }
                $s += 10
                $ti = $found + 1
            }
            if ($s -ge 0)
            {
                $s -= ($tl.Length - $qlen)
                if ($s -lt 1) { $s = 1 }
            }
        }
        if ($s -lt 0) { continue }

        $len = $n.Length
        $full = ($hits.Count -ge $Limit)
        if ($full)
        {
            $w = $hits[$hits.Count - 1]
            if ($s -ne [int]$w.Score) { if ($s -lt [int]$w.Score) { continue } }
            else
            {
                if ($len -gt [int]$w.Len) { continue }
                if ($len -eq [int]$w.Len -and [string]::Compare($n, [string]$w.Name, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { continue }
            }
        }

        $p = $hits.Count
        if ($full) { $p = $hits.Count - 1 }
        while ($p -gt 0)
        {
            $b = $hits[$p - 1]
            $bs = [int]$b.Score
            $better = $false
            if ($s -ne $bs) { $better = ($s -gt $bs) }
            else
            {
                $bl = [int]$b.Len
                if ($len -ne $bl) { $better = ($len -lt $bl) }
                else { $better = ([string]::Compare($n, [string]$b.Name, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) }
            }
            if (-not $better) { break }
            $p--
        }
        if ($full) { $hits.RemoveAt($hits.Count - 1) }
        $hits.Insert($p, [pscustomobject]@{ Name = $n; Score = $s; Len = $len })
    }
    if ($hits.Count -eq 0) { return ,@() }

    $out = @()
    foreach ($h in $hits) { $out += [string]$h.Name }
    return ,$out
}

#endregion
#region Price trend

function Get-MarketTrend
{
    [CmdletBinding()]
    param(
        [double[]]$Prices = @(),
        [int]$Range = 25
    )

    $list = [System.Collections.Generic.List[double]]::new()
    foreach ($p in $Prices) { if ($p -gt 0) { [void]$list.Add([double]$p) } }
    $count = $list.Count

    $result = [pscustomobject]@{
        Direction = 'none'
        Arrow     = [string][char]0x2192
        DeltaPct  = [double]0
        OlderAvg  = $null
        RecentAvg = $null
        Count     = $count
        Used      = 0
    }
    if ($count -lt 2)
    {
        if ($count -eq 1) { $result.Direction = 'flat' }
        return $result
    }

    $range = $Range
    if ($range -lt 1) { $range = 1 }
    if ($range -gt 50) { $range = 50 }

    $used = [Math]::Min([Math]::Max($range, 2), $count)
    $result.Used = $used

    $window = [System.Collections.Generic.List[double]]::new()
    for ($i = ($count - $used); $i -lt $count; $i++) { [void]$window.Add($list[$i]) }

    $half = [int][Math]::Floor($used / 2.0)
    if ($half -lt 1) { $half = 1 }

    $older = [System.Collections.Generic.List[double]]::new()
    for ($i = 0; $i -lt $half; $i++) { [void]$older.Add($window[$i]) }
    $recent = [System.Collections.Generic.List[double]]::new()
    for ($i = $half; $i -lt $window.Count; $i++) { [void]$recent.Add($window[$i]) }
    if ($older.Count -lt 1 -or $recent.Count -lt 1) { return $result }

    $olderSum = 0.0
    foreach ($v in $older) { $olderSum += $v }
    $recentSum = 0.0
    foreach ($v in $recent) { $recentSum += $v }
    $olderAvg = [double]($olderSum / $older.Count)
    $recentAvg = [double]($recentSum / $recent.Count)
    $result.OlderAvg = $olderAvg
    $result.RecentAvg = $recentAvg
    if ($olderAvg -le 0) { return $result }

    $delta = ($recentAvg - $olderAvg) / $olderAvg
    $result.DeltaPct = [Math]::Round(($delta * 100.0), 2)

    if ($delta -gt 0.01)
    {
        $result.Direction = 'up'
        $result.Arrow = [string][char]0x2191
    }
    elseif ($delta -lt -0.01)
    {
        $result.Direction = 'down'
        $result.Arrow = [string][char]0x2193
    }
    else
    {
        $result.Direction = 'flat'
    }
    return $result
}

#endregion
#region Access

function ConvertTo-MarketBytes
{
    [CmdletBinding()]
    [OutputType([byte[]])]
    param([string]$Text)
    if ($null -eq $Text) { $Text = '' }
    return [System.Text.Encoding]::UTF8.GetBytes($Text)
}

function ConvertFrom-MarketBytes
{
    [CmdletBinding()]
    param([byte[]]$Bytes)
    if ($null -eq $Bytes) { return '' }
    $enc = New-Object System.Text.UTF8Encoding($false, $true)
    return $enc.GetString($Bytes)
}

function Get-MarketDerivedKeys
{
    [CmdletBinding()]
    param(
        [string]$Password,
        [byte[]]$Salt,
        [int]$Iterations
    )
    if ($Iterations -lt 10000) { $Iterations = 10000 }
    if ($Iterations -gt 10000000) { $Iterations = 10000000 }
    $kdf = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $Salt, $Iterations)
    try
    {
        $material = $kdf.GetBytes(64)
        $enc = [byte[]](New-Object byte[] 32)
        $mac = [byte[]](New-Object byte[] 32)
        [Array]::Copy($material, 0, $enc, 0, 32)
        [Array]::Copy($material, 32, $mac, 0, 32)
        return @{ Enc = $enc; Mac = $mac }
    }
    finally
    {
        try { $kdf.Reset() } catch { }
        try { $kdf.Dispose() } catch { }
    }
}

function Test-MarketTagEqual
{
    [CmdletBinding()]
    [OutputType([bool])]
    param([byte[]]$A, [byte[]]$B)
    if ($null -eq $A -or $null -eq $B) { return $false }
    if ($A.Length -ne $B.Length) { return $false }
    $diff = 0
    for ($i = 0; $i -lt $A.Length; $i++) { $diff = $diff -bor ($A[$i] -bxor $B[$i]) }
    return ($diff -eq 0)
}

function Protect-MarketPayload
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Password,
        [Parameter(Mandatory = $true)][string]$OutFile,
        [int]$Iterations = 200000
    )

    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try
    {
        $salt = [byte[]](New-Object byte[] 16)
        $iv = [byte[]](New-Object byte[] 16)
        $rng.GetBytes($salt)
        $rng.GetBytes($iv)
    }
    finally { try { $rng.Dispose() } catch { } }

    $keys = Get-MarketDerivedKeys -Password $Password -Salt $salt -Iterations $Iterations

    $aes = [System.Security.Cryptography.Aes]::Create()
    try
    {
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $aes.Key = $keys.Enc
        $aes.IV = $iv
        $plain = ConvertTo-MarketBytes -Text $Text
        $cipher = $aes.CreateEncryptor().TransformFinalBlock($plain, 0, $plain.Length)
    }
    finally { try { $aes.Dispose() } catch { } }

    $prefix = New-Object byte[] 41
    [System.Text.Encoding]::ASCII.GetBytes('EVMV') | ForEach-Object -Begin { $i = 0 } -Process { $prefix[$i] = $_; $i++ }
    $prefix[4] = 1
    $prefix[5] = [byte]($Iterations -band 0xFF)
    $prefix[6] = [byte](($Iterations -shr 8) -band 0xFF)
    $prefix[7] = [byte](($Iterations -shr 16) -band 0xFF)
    $prefix[8] = [byte](($Iterations -shr 24) -band 0xFF)
    [Array]::Copy($salt, 0, $prefix, 9, 16)
    [Array]::Copy($iv, 0, $prefix, 25, 16)

    $hmac = New-Object System.Security.Cryptography.HMACSHA256 -ArgumentList @(, $keys.Mac)
    try
    {
        $macInput = New-Object byte[] ($prefix.Length + $cipher.Length)
        [Array]::Copy($prefix, 0, $macInput, 0, $prefix.Length)
        [Array]::Copy($cipher, 0, $macInput, $prefix.Length, $cipher.Length)
        $tag = $hmac.ComputeHash($macInput)
    }
    finally { try { $hmac.Dispose() } catch { } }

    $out = New-Object byte[] (41 + 32 + $cipher.Length)
    [Array]::Copy($prefix, 0, $out, 0, 41)
    [Array]::Copy($tag, 0, $out, 41, 32)
    [Array]::Copy($cipher, 0, $out, 73, $cipher.Length)

    $dir = Split-Path -Path $OutFile -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    [System.IO.File]::WriteAllBytes($OutFile, $out)
    return $out.Length
}

function Unprotect-MarketPayload
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [Parameter(Mandatory = $true)][string]$Password
    )

    $keys = $null
    try
    {
        if ($null -eq $Bytes -or $Bytes.Length -lt 89) { return $null }

        $magic = [System.Text.Encoding]::ASCII.GetString($Bytes, 0, 4)
        if ($magic -ne 'EVMV') { return $null }
        if ($Bytes[4] -ne 1) { return $null }

        $iters = [int]$Bytes[5] -bor ([int]$Bytes[6] -shl 8) -bor ([int]$Bytes[7] -shl 16) -bor ([int]$Bytes[8] -shl 24)
        if ($iters -lt 10000 -or $iters -gt 10000000) { return $null }

        $salt = [byte[]](New-Object byte[] 16)
        $iv = [byte[]](New-Object byte[] 16)
        [Array]::Copy($Bytes, 9, $salt, 0, 16)
        [Array]::Copy($Bytes, 25, $iv, 0, 16)

        $tag = [byte[]](New-Object byte[] 32)
        [Array]::Copy($Bytes, 41, $tag, 0, 32)
        $cipher = [byte[]](New-Object byte[] ($Bytes.Length - 73))
        [Array]::Copy($Bytes, 73, $cipher, 0, $cipher.Length)

        $keys = Get-MarketDerivedKeys -Password $Password -Salt $salt -Iterations $iters

        $hmac = New-Object System.Security.Cryptography.HMACSHA256 -ArgumentList @(, $keys.Mac)
        try
        {
            $macInput = New-Object byte[] (41 + $cipher.Length)
            [Array]::Copy($Bytes, 0, $macInput, 0, 41)
            [Array]::Copy($cipher, 0, $macInput, 41, $cipher.Length)
            $calc = $hmac.ComputeHash($macInput)
        }
        finally { try { $hmac.Dispose() } catch { } }

        if (-not (Test-MarketTagEqual -A $tag -B $calc)) { return $null }

        $aes = [System.Security.Cryptography.Aes]::Create()
        try
        {
            $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
            $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
            $aes.Key = $keys.Enc
            $aes.IV = $iv
            $plain = $aes.CreateDecryptor().TransformFinalBlock($cipher, 0, $cipher.Length)
        }
        finally { try { $aes.Dispose() } catch { } }

        return (ConvertFrom-MarketBytes -Bytes $plain)
    }
    catch { return $null }
    finally
    {
        if ($null -ne $keys) { try { [Array]::Clear($keys.Enc, 0, $keys.Enc.Length) } catch { } try { [Array]::Clear($keys.Mac, 0, $keys.Mac.Length) } catch { } }
    }
}

function Get-MarketVaultUrl
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Code = '')

    $base = ''
    try { $base = Get-MarketSyncBaseUrl } catch { $base = '' }
    if ([string]::IsNullOrWhiteSpace($base)) { return '' }
    $url = "$base".TrimEnd('/') + '/vault'

    $cred = $null
    try { $cred = Get-MarketSyncCredentials -Code $Code } catch { $cred = $null }
    if ($null -eq $cred -or -not $cred.Ok) { return $url }

    $u = [uri]::EscapeDataString([string]$cred.User)
    $p = [uri]::EscapeDataString([string]$cred.Pass)
    return ("$url" + "?u=$u&p=$p")
}

function Test-MarketVaultRemoteAvailable
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    $u = Get-MarketVaultUrl
    if ([string]::IsNullOrWhiteSpace($u)) { return $false }
    return [bool]($u -match '^(?i)https://')
}

function Get-MarketVaultLocalPath
{
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $dir = Join-Path (Get-MarketDataRoot) 'vip'
    foreach ($n in @('market.vault', 'vault.dat'))
    {
        $p = Join-Path $dir $n
        try { if (Test-Path -LiteralPath $p -PathType Leaf) { return $p } } catch { }
    }
    return (Join-Path $dir 'market.vault')
}

function Test-MarketVaultLocalAvailable
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    $p = Get-MarketVaultLocalPath
    if ([string]::IsNullOrWhiteSpace($p)) { return $false }
    try { return [bool](Test-Path -LiteralPath $p -PathType Leaf) } catch { return $false }
}

function Get-MarketVaultTempPath
{
    [CmdletBinding()]
    [OutputType([string])]
    param()
    return (Join-Path (Get-MarketDataRoot) 'vip\market.vault.tmp')
}

function Remove-MarketVaultTempFile
{
    [CmdletBinding()]
    param()
    try
    {
        $p = Get-MarketVaultTempPath
        if (-not [string]::IsNullOrWhiteSpace($p) -and (Test-Path -LiteralPath $p -PathType Leaf))
        {
            Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
        }
    }
    catch { }
    $script:MarketVaultTempUrl = ''
}

function Test-MarketVaultAvailable
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    if (Test-MarketVaultRemoteAvailable) { return $true }
    if (Test-MarketVaultLocalAvailable) { return $true }
    return $false
}

function Get-MarketVaultBytes
{
    [CmdletBinding()]
    [OutputType([byte[]])]
    param([ValidateSet('Auto', 'Local', 'Remote')][string]$Source = 'Auto', [string]$Code = '')

    if ($Source -eq 'Auto')
    {
        if (Test-MarketVaultLocalAvailable) { $Source = 'Local' }
        else { $Source = 'Remote' }
    }

    if ($Source -eq 'Local')
    {
        if (-not (Test-MarketVaultLocalAvailable)) { throw 'No local update file is present.' }
        $p = Get-MarketVaultLocalPath
        try { return [System.IO.File]::ReadAllBytes($p) }
        catch { throw "The local update file could not be read: $($_.Exception.Message)" }
    }

    $url = Get-MarketVaultUrl -Code $Code
    if ([string]::IsNullOrWhiteSpace($url)) { throw 'No update source configured.' }
    if ($url -notmatch '^(?i)https://') { throw 'The update source must be a secure (https) address.' }

    $cred = $null
    try { $cred = Get-MarketSyncCredentials -Code $Code } catch { $cred = $null }
    if ($null -eq $cred -or -not $cred.Ok)
    {
        $m = 'Enter the access code first.'
        if ($null -ne $cred -and -not [string]::IsNullOrWhiteSpace("$($cred.Message)")) { $m = "$($cred.Message)" }
        throw $m
    }

    if ($null -ne $script:MarketVaultCacheBytes -and "$script:MarketVaultCacheUrl" -eq $url -and $script:MarketVaultCacheBytes.Length -gt 0)
    {
        return $script:MarketVaultCacheBytes
    }

    $tmp = Get-MarketVaultTempPath
    if ("$script:MarketVaultTempUrl" -eq $url)
    {
        try
        {
            if (Test-Path -LiteralPath $tmp -PathType Leaf)
            {
                $staged = [System.IO.File]::ReadAllBytes($tmp)
                if ($staged.Length -gt 0)
                {
                    $script:MarketVaultCacheUrl = $url
                    $script:MarketVaultCacheBytes = $staged
                    return $staged
                }
            }
        }
        catch { }
    }

    try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch { }
    $bytes = $null
    try
    {
        $wc = New-Object System.Net.WebClient
        try { $bytes = $wc.DownloadData($url) }
        finally { try { $wc.Dispose() } catch { } }
    }
    catch
    {
        $status = 0
        $ex = $_.Exception
        while ($null -ne $ex)
        {
            $resp = $null
            try { $resp = $ex.Response } catch { $resp = $null }
            if ($null -ne $resp) { try { $status = [int]$resp.StatusCode } catch { $status = 0 } }
            if ($status -ne 0) { break }
            $ex = $ex.InnerException
        }
        if ($status -eq 401 -or $status -eq 403) { throw 'The access code was refused by the update source.' }
        throw
    }
    if ($null -eq $bytes -or $bytes.Length -eq 0) { throw 'The update source answered with an empty file.' }

    $script:MarketVaultCacheUrl = $url
    $script:MarketVaultCacheBytes = $bytes

    try
    {
        $dir = Split-Path -Parent $tmp
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir -PathType Container))
        {
            [void](New-Item -Path $dir -ItemType Directory -Force)
        }
        [System.IO.File]::WriteAllBytes($tmp, $bytes)
        $script:MarketVaultTempUrl = $url
    }
    catch { $script:MarketVaultTempUrl = '' }

    return $bytes
}

#region Online sync file

function Get-MarketSyncBaseUrl
{
    [CmdletBinding()]
    [OutputType([string])]
    param()
    return 'https://entropiamarket.entropia-dashboard.workers.dev'
}

function Get-MarketSyncAccessCode
{
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $code = ''
    try
    {
        $box = Get-MarketUIControl -Name 'MarketCodeInput'
        if ($null -ne $box) { $code = ([string]$box.Text).Trim() }
    }
    catch { $code = '' }

    if ($code -eq '')
    {
        try { $code = ("$(Get-MarketConfigValue -Key 'AccessCode' -Default '')").Trim() } catch { $code = '' }
    }
    return $code
}

function Get-MarketSyncCredentials
{
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string]$Code = '')

    if ("$Code".Trim() -eq '')
    {
        try { $Code = Get-MarketSyncAccessCode } catch { $Code = '' }
    }
    $Code = "$Code".Trim()
    if ($Code -eq '')
    {
        return [pscustomobject]@{ Ok = $false; User = ''; Pass = ''; Message = 'Enter the access code first.' }
    }
    $dot = $Code.IndexOf([char]'.')
    if ($dot -lt 1 -or $dot -ge ($Code.Length - 1))
    {
        return [pscustomobject]@{ Ok = $false; User = ''; Pass = ''; Message = 'The access code must look like name.password for Sync.' }
    }
    return [pscustomobject]@{
        Ok      = $true
        User    = $Code.Substring(0, $dot)
        Pass    = $Code.Substring($dot + 1)
        Message = ''
    }
}

function Get-MarketSyncUrl
{
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $base = ''
    try { $base = Get-MarketSyncBaseUrl } catch { $base = '' }
    $base = "$base".TrimEnd('/')
    $feed = "$base/data"
    if ([string]::IsNullOrWhiteSpace($base)) { return '' }

    $cred = $null
    try { $cred = Get-MarketSyncCredentials } catch { $cred = $null }
    if ($null -eq $cred -or -not $cred.Ok) { return $feed }

    $u = [uri]::EscapeDataString([string]$cred.User)
    $p = [uri]::EscapeDataString([string]$cred.Pass)
    return ("$feed" + "?u=$u&p=$p")
}

function Update-MarketSyncFromUrl
{
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $out = [pscustomobject]@{ Ok = $false; Count = 0; Message = '' }

    $cred = $null
    try { $cred = Get-MarketSyncCredentials } catch { $cred = $null }
    if ($null -eq $cred -or -not $cred.Ok)
    {
        $m = 'Enter the access code first.'
        if ($null -ne $cred -and -not [string]::IsNullOrWhiteSpace("$($cred.Message)")) { $m = "$($cred.Message)" }
        $out.Message = $m
        return $out
    }

    $base = ''
    try { $base = Get-MarketSyncBaseUrl } catch { $base = '' }
    if ([string]::IsNullOrWhiteSpace($base))
    {
        $out.Message = 'No sync address configured.'
        return $out
    }
    if ($base -notmatch '^(?i)https://')
    {
        $out.Message = 'The sync address must be a secure (https) address.'
        return $out
    }

    $url = Get-MarketSyncUrl

    $raw = ''
    try
    {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $wc = New-Object System.Net.WebClient
        try
        {
            $wc.Encoding = [System.Text.Encoding]::UTF8
            $raw = $wc.DownloadString($url)
        }
        finally { try { $wc.Dispose() } catch { } }
    }
    catch
    {
        $status = 0
        $ex = $_.Exception
        while ($null -ne $ex)
        {
            $resp = $null
            try { $resp = $ex.Response } catch { $resp = $null }
            if ($null -ne $resp)
            {
                try { $status = [int]$resp.StatusCode } catch { $status = 0 }
            }
            if ($status -ne 0) { break }
            $ex = $ex.InnerException
        }

        if ($status -eq 401 -or $status -eq 403)
        {
            $out.Message = 'Sync refused the access code.'
        }
        else
        {
            $out.Message = "Sync download failed: $($_.Exception.Message)"
        }
        return $out
    }

    if ([string]::IsNullOrWhiteSpace($raw)) { $out.Message = 'The sync download was empty.'; return $out }

    $parsed = $null
    try { $parsed = $raw | ConvertFrom-Json } catch { $parsed = $null }
    if ($null -eq $parsed) { $out.Message = 'The downloaded file is not valid JSON.'; return $out }
    $offers = $null
    try { $offers = $parsed.offers } catch { $offers = $null }
    if ($null -eq $offers) { $out.Message = 'The downloaded file holds no offers.'; return $out }

    $path = Join-Path (Join-Path (Get-MarketDataRoot) 'market') 'market_sync.json'
    $tmp = "$path.new"
    $bak = "$path.bak"
    try
    {
        $dir = Split-Path -Parent $path
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir -PathType Container))
        {
            [void](New-Item -Path $dir -ItemType Directory -Force)
        }
        [System.IO.File]::WriteAllText($tmp, $raw, [System.Text.UTF8Encoding]::new($false))
        if (Test-Path -LiteralPath $path -PathType Leaf)
        {
            [System.IO.File]::Replace($tmp, $path, $bak)
            try { [System.IO.File]::Delete($bak) } catch { }
        }
        else { [System.IO.File]::Move($tmp, $path) }
    }
    catch
    {
        try { if (Test-Path -LiteralPath $tmp -PathType Leaf) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } } catch { }
        $out.Message = "Sync file could not be saved: $($_.Exception.Message)"
        return $out
    }

    $out.Ok = $true
    $out.Count = @($offers).Count

    try
    {
        $legacy = Join-Path (Get-MarketDataRoot) 'market_sync.json'
        if (Test-Path -LiteralPath $legacy -PathType Leaf) { Remove-Item -LiteralPath $legacy -Force -ErrorAction SilentlyContinue }
    }
    catch { }

    return $out
}

#endregion Online sync file

function Test-MarketAccess
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    $state = $global:DashboardConfig.State
    if ($null -eq $state) { return $false }
    try { return [bool]$state['MarketAccess'] } catch { return $false }
}

function Lock-MarketAccess
{
    [CmdletBinding()]
    param()

    try { Remove-MarketVaultTempFile } catch { }

    try
    {
        if ($null -ne (Get-Module -Name 'MarketOnline' -ErrorAction SilentlyContinue))
        {
            if (Test-DashboardCommand 'Close-MarketOnline') { try { [void](Close-MarketOnline) } catch { } }
            try { Remove-Module -Name 'MarketOnline' -Force -ErrorAction SilentlyContinue } catch { }
            try { Clear-DashboardCommandCache } catch { }
        }
    }
    catch { }

    $state = $global:DashboardConfig.State
    if ($null -eq $state) { return }
    $state['MarketAccess'] = $false
    try { $state['MarketOnlineActive'] = $false } catch { }
    try { $state['MarketOnlineView'] = '' } catch { }
    try { $state['MarketSyncData'] = $null } catch { }
    try { $state['MarketSyncError'] = '' } catch { }
    try { $state['MarketOnlineFetchedAt'] = $null } catch { }
    try { $state['MarketOnlineStatus'] = 'locked' } catch { }
    try { $state['MarketOnlineError'] = '' } catch { }
    try { $state['MarketLiveSource'] = '' } catch { }
    try { $state['MarketLiveCount'] = 0 } catch { }
}

function Unlock-MarketAccess
{
    [CmdletBinding()]
    param([string]$Code)

    $code = "$Code".Trim()
    if ($code -eq '')
    {
        return [pscustomobject]@{ Ok = $false; Message = 'Enter an access code first.' }
    }

    if (-not (Test-MarketVaultAvailable))
    {
        return [pscustomobject]@{ Ok = $false; Message = 'No update source configured.' }
    }

    $text = $null
    $wrongCode = $false
    $sourceError = ''
    foreach ($src in @('Local', 'Remote'))
    {
        $bytes = $null
        try { $bytes = Get-MarketVaultBytes -Source $src -Code $code }
        catch
        {
            $sourceError = "$($_.Exception.Message)"
            continue
        }

        $vaultPass = if ($code -match '\.') { $code.Split('.', 2)[1] } else { $code }

        $text = Unprotect-MarketPayload -Bytes $bytes -Password $vaultPass

        if ($null -eq $text -and $vaultPass -ne $code)
        {
            $text = Unprotect-MarketPayload -Bytes $bytes -Password $code
        }
        $bytes = $null
        if ($null -ne $text) { break }
        $wrongCode = $true
    }

    if ($null -eq $text)
    {
        if ($wrongCode) { return [pscustomobject]@{ Ok = $false; Message = 'The access code was not accepted.' } }
        if ("$sourceError" -match '(?i)access code') { return [pscustomobject]@{ Ok = $false; Message = "$sourceError" } }
        return [pscustomobject]@{ Ok = $false; Message = "Update source unreachable: $sourceError" }
    }

    $sb = $null
    try { $sb = [scriptblock]::Create($text) }
    catch { return [pscustomobject]@{ Ok = $false; Message = "The update could not be read: $($_.Exception.Message)" } }
    $text = $null

    try
    {
        if ($null -ne (Get-Module -Name 'MarketOnline' -ErrorAction SilentlyContinue))
        {
            if (Test-DashboardCommand 'Close-MarketOnline') { try { [void](Close-MarketOnline) } catch { } }
            Remove-Module -Name 'MarketOnline' -Force -ErrorAction SilentlyContinue
            try { Clear-DashboardCommandCache } catch { }
        }
    }
    catch { }

    $mod = $null
    try
    {
        $mod = New-Module -Name 'MarketOnline' -ScriptBlock $sb -ErrorAction Stop
        Import-Module -ModuleInfo $mod -Global -Force -DisableNameChecking -ErrorAction Stop
        try { Clear-DashboardCommandCache } catch { }
    }
    catch
    {
        $sb = $null
        return [pscustomobject]@{ Ok = $false; Message = "The update could not be loaded: $($_.Exception.Message)" }
    }
    $sb = $null

    if ($null -eq (Get-Module -Name 'MarketOnline' -ErrorAction SilentlyContinue))
    {
        return [pscustomobject]@{ Ok = $false; Message = 'The update could not be verified after import.' }
    }

    $state = $global:DashboardConfig.State
    $prevAccess = $false
    if ($null -ne $state) { try { $prevAccess = [bool]$state['MarketAccess'] } catch { $prevAccess = $false } }

    $initOk = $true
    $initMsg = ''
    if ($null -ne $state) { $state['MarketAccess'] = $true; $state['MarketOnlineActive'] = $true }
    try
    {
        if (Test-DashboardCommand 'Initialize-MarketOnline') { [void](Initialize-MarketOnline) }
        else { throw 'Initialize-MarketOnline not found in the update' }
    }
    catch
    {
        $initOk = $false
        $initMsg = "$($_.Exception.Message)"
    }

    if (-not $initOk)
    {
        try { if (Test-DashboardCommand 'Close-MarketOnline') { [void](Close-MarketOnline) } } catch { }
        try { Remove-Module -Name 'MarketOnline' -Force -ErrorAction SilentlyContinue } catch { }
        try { Clear-DashboardCommandCache } catch { }
        if ($null -ne $state) { $state['MarketOnlineActive'] = $false; $state['MarketAccess'] = $prevAccess }
        return [pscustomobject]@{ Ok = $false; Message = "The update could not be started: $initMsg" }
    }

    return [pscustomobject]@{ Ok = $true; Message = 'Access unlocked.' }
}

#endregion
#region Icons

function Get-MarketIcon
{
    [CmdletBinding()]
    [OutputType([System.Drawing.Bitmap])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Search', 'Plus', 'Minus', 'Lock', 'Unlock', 'Sync', 'Send', 'Arrow')]
        [string]$Kind = 'Search',
        [int]$Size = 14
    )

    if ($Size -lt 8) { $Size = 8 }
    if ($Size -gt 64) { $Size = 64 }

    $color = [System.Drawing.Color]::White
    try
    {
        if (Test-DashboardCommand 'Get-Theme')
        {
            $c = Get-Theme 'PrimaryText'
            if ($c -is [System.Drawing.Color]) { $color = $c }
        }
    }
    catch { }

    $key = "$Kind|$Size|$($color.ToArgb())"
    if ($script:MarketIconCache.ContainsKey($key)) { return $script:MarketIconCache[$key] }

    $bmp = [System.Drawing.Bitmap]::new($Size, $Size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $pen = $null
    try
    {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.Clear([System.Drawing.Color]::Transparent)

        $pen = [System.Drawing.Pen]::new($color, [float][Math]::Max(1.2, ($Size / 9.0)))
        $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
        $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
        $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round

        $s = [double]$Size
        switch ($Kind)
        {
            'Search'
            {
                $r = ($s * 0.30)
                $cx = ($s * 0.40)
                $cy = ($s * 0.40)
                $g.DrawEllipse($pen, [float]($cx - $r), [float]($cy - $r), [float](2 * $r), [float](2 * $r))
                $g.DrawLine($pen, [float]($cx + $r * 0.70), [float]($cy + $r * 0.70), [float]($s - 1.5), [float]($s - 1.5))
            }
            'Plus'
            {
                $g.DrawLine($pen, [float]($s / 2), [float]($s * 0.18), [float]($s / 2), [float]($s * 0.82))
                $g.DrawLine($pen, [float]($s * 0.18), [float]($s / 2), [float]($s * 0.82), [float]($s / 2))
            }
            'Minus'
            {
                $g.DrawLine($pen, [float]($s * 0.18), [float]($s / 2), [float]($s * 0.82), [float]($s / 2))
            }
            'Lock'
            {
                $bw = ($s * 0.56); $bx = ($s * 0.22); $by = ($s * 0.48); $bh = ($s * 0.36)
                $g.DrawRectangle($pen, [float]$bx, [float]$by, [float]$bw, [float]$bh)
                $g.DrawArc($pen, [float]($bx + $bw * 0.14), [float]($by - $bh * 0.55), [float]($bw * 0.72), [float]($bh * 1.10), 180, 180)
            }
            'Unlock'
            {
                $bw = ($s * 0.56); $bx = ($s * 0.22); $by = ($s * 0.48); $bh = ($s * 0.36)
                $g.DrawRectangle($pen, [float]$bx, [float]$by, [float]$bw, [float]$bh)
                $g.DrawArc($pen, [float]($bx + $bw * 0.44), [float]($by - $bh * 0.55), [float]($bw * 0.72), [float]($bh * 1.10), 180, 180)
            }
            'Sync'
            {
                $cx = ($s / 2); $cy = ($s / 2); $r = ($s * 0.32)
                $startA = -100.0; $sweep = 300.0
                $g.DrawArc($pen, [float]($cx - $r), [float]($cy - $r), [float](2 * $r), [float](2 * $r), [float]$startA, [float]$sweep)
                $endRad = (($startA + $sweep) * [Math]::PI / 180.0)
                $tipX = $cx + $r * [Math]::Cos($endRad)
                $tipY = $cy + $r * [Math]::Sin($endRad)
                $bx = -[Math]::Sin($endRad)
                $by = [Math]::Cos($endRad)
                $bl = ($s * 0.26)
                foreach ($a in @(30.0, -30.0))
                {
                    $rad = $a * [Math]::PI / 180.0
                    $rx = ($bx * [Math]::Cos($rad) - $by * [Math]::Sin($rad))
                    $ry = ($bx * [Math]::Sin($rad) + $by * [Math]::Cos($rad))
                    $g.DrawLine($pen, [float]$tipX, [float]$tipY, [float]($tipX + $rx * $bl), [float]($tipY + $ry * $bl))
                }
            }
            'Send'
            {
                $g.DrawLines($pen, [System.Drawing.PointF[]]@(
                        [System.Drawing.PointF]::new([float]($s * 0.14), [float]($s * 0.50)),
                        [System.Drawing.PointF]::new([float]($s * 0.86), [float]($s * 0.16)),
                        [System.Drawing.PointF]::new([float]($s * 0.86), [float]($s * 0.84)),
                        [System.Drawing.PointF]::new([float]($s * 0.14), [float]($s * 0.50))
                    ))
                $g.DrawLine($pen, [float]($s * 0.42), [float]($s * 0.35), [float]($s * 0.42), [float]($s * 0.65))
            }
            'Arrow'
            {
                $g.DrawLine($pen, [float]($s * 0.15), [float]($s * 0.50), [float]($s * 0.85), [float]($s * 0.50))
                $g.DrawLine($pen, [float]($s * 0.55), [float]($s * 0.22), [float]($s * 0.85), [float]($s * 0.50))
                $g.DrawLine($pen, [float]($s * 0.55), [float]($s * 0.78), [float]($s * 0.85), [float]($s * 0.50))
            }
        }
    }
    catch { Write-Verbose "MARKET: icon '$Kind' draw failed: $($_.Exception.Message)" }
    finally
    {
        if ($null -ne $pen) { try { $pen.Dispose() } catch { } }
        try { $g.Dispose() } catch { }
    }

    $script:MarketIconCache[$key] = $bmp
    return $bmp
}

function Build-MarketItemIconMap
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Map
    )

    if ($null -eq $Map) { return }

    $cat = $null
    try { $cat = Get-MarketCatalog } catch { $cat = $null }
    if ($null -eq $cat -or $null -eq $cat.ByName) { return }

    $sep = [string][System.IO.Path]::DirectorySeparatorChar
    $baseCache = @{}

    foreach ($kv in $cat.ByName.GetEnumerator())
    {
        $conn = ''
        try { $conn = [string]$kv.Value.connection } catch { $conn = '' }
        if ([string]::IsNullOrWhiteSpace($conn)) { continue }

        $p = "$conn".Trim()
        if ($p -eq '') { continue }
        try { if ($p.IndexOf('%') -ge 0) { $p = [Environment]::ExpandEnvironmentVariables($p) } } catch { }
        try
        {
            if (-not [System.IO.Path]::IsPathRooted($p))
            {
                $norm = $p.Replace('/', $sep)
                $segKey = $norm
                $cut = $norm.IndexOf($sep)
                if ($cut -ge 0) { $segKey = $norm.Substring(0, $cut) }
                $segKey = "$segKey".ToLowerInvariant()
                $base = $null
                if (-not $baseCache.TryGetValue($segKey, [ref]$base))
                {
                    try { $base = Get-MarketIconBaseRoot -Relative $p } catch { $base = '' }
                    if ([string]::IsNullOrWhiteSpace($base)) { try { $base = Get-MarketDataRoot } catch { $base = '' } }
                    $baseCache[$segKey] = $base
                }
                if ("$base" -ne '') { $p = "$base$sep$norm" }
            }
        }
        catch { }
        if ($p -eq '') { continue }
        if (-not $Map.ContainsKey($kv.Key)) { $Map[$kv.Key] = $p }
    }
}

function Get-MarketItemIconMap
{
    [CmdletBinding()]
    param()

    if ($null -ne $script:MarketIconMap) { return ,$script:MarketIconMap }

    if ($script:MarketIconMapLoading)
    {
        if ($null -eq $script:MarketIconMapPartial)
        {
            $script:MarketIconMapPartial = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        }
        return ,$script:MarketIconMapPartial
    }

    $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    try { Build-MarketItemIconMap -Map $map } catch { Write-Verbose "MARKET: icon index build failed: $($_.Exception.Message)" }

    $script:MarketIconMap = $map
    return ,$script:MarketIconMap
}

function Get-MarketIconMapStatus
{
    [CmdletBinding()]
    param()
    $entries = 0
    if ($null -ne $script:MarketIconMap) { $entries = $script:MarketIconMap.Count }
    return [pscustomobject]@{
        Ready   = ($null -ne $script:MarketIconMap)
        Loading = [bool]$script:MarketIconMapLoading
        Entries = $entries
        Error   = [string]$script:MarketIconMapError
    }
}

function Test-MarketIconMapReady
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    if ($null -ne $script:MarketIconMap) { return $true }
    return (-not $script:MarketIconMapLoading)
}

function Start-MarketIconMapLoad
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    if ($null -ne $script:MarketIconMap) { return $false }
    if ($script:MarketIconMapLoading) { return $false }

    $path = ''
    try { $path = [string]$MyInvocation.MyCommand.Module.Path } catch { $path = '' }
    if ([string]::IsNullOrWhiteSpace($path)) { try { $path = [string]$PSCommandPath } catch { $path = '' } }
    if ([string]::IsNullOrWhiteSpace($path)) { return $false }
    try { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $false } } catch { return $false }

    $ps = $null
    $rs = $null
    $handle = $null
    try
    {
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        $q = $path.Replace("'", "''")
        [void]$ps.AddScript("Import-Module -Name '$q' -Force -DisableNameChecking -ErrorAction Stop; `$m = Get-MarketItemIconMap; Write-Output -NoEnumerate `$m")
        $handle = $ps.BeginInvoke()
    }
    catch
    {
        Write-Verbose "MARKET: the icon index could not start in the background: $($_.Exception.Message)"
        try { if ($null -ne $ps) { $ps.Dispose() } } catch { }
        try { if ($null -ne $rs) { $rs.Dispose() } } catch { }
        return $false
    }

    $script:MarketIconMapPowerShell = $ps
    $script:MarketIconMapRunspace = $rs
    $script:MarketIconMapHandle = $handle
    $script:MarketIconMapLoading = $true

    if ($null -eq $script:MarketIconMapTimer)
    {
        $t = New-Object System.Windows.Forms.Timer
        $t.Interval = 250
        $t.Add_Tick({ Complete-MarketIconMapLoad })
        $script:MarketIconMapTimer = $t
    }
    try { $script:MarketIconMapTimer.Start() } catch { }
    return $true
}

function Complete-MarketIconMapLoad
{
    [CmdletBinding()]
    param()

    $handle = $script:MarketIconMapHandle
    $ps = $script:MarketIconMapPowerShell
    if ($null -eq $handle -or $null -eq $ps)
    {
        try { Stop-MarketIconMapLoad } catch { }
        return
    }
    try { if (-not $handle.IsCompleted) { return } } catch { try { Stop-MarketIconMapLoad } catch { }; return }

    $map = $null
    $script:MarketIconMapError = ''
    try
    {
        $res = $ps.EndInvoke($handle)
        if ($null -ne $res -and $res.Count -eq 1)
        {
            $map = $res[0]
            try { if ($null -ne $res[0].BaseObject) { $map = $res[0].BaseObject } } catch { }
        }
        elseif ($null -ne $res)
        {
            $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($it in $res)
            {
                $o = $it
                try { if ($null -ne $it.BaseObject) { $o = $it.BaseObject } } catch { }
                if ($o -is [System.Collections.DictionaryEntry]) { $map[[string]$o.Key] = [string]$o.Value }
                elseif ($o -is [System.Collections.Generic.KeyValuePair[string, string]]) { $map[[string]$o.Key] = [string]$o.Value }
            }
        }
        foreach ($e in $ps.Streams.Error) { $script:MarketIconMapError += "$e`n" }
    }
    catch { $script:MarketIconMapError = "$($_.Exception.Message)" }

    try { Stop-MarketIconMapLoad } catch { }

    if ($null -eq $map -or $map.Count -eq 0)
    {
        if ($script:MarketIconMapError -ne '') { Write-Verbose "MARKET: background icon index failed: $($script:MarketIconMapError)" }
        return
    }

    $script:MarketIconMap = $map
    try { Refresh-MarketIconConsumers } catch { }
}

function Refresh-MarketIconConsumers
{
    [CmdletBinding()]
    param()

    try
    {
        if (Test-DashboardCommand 'Clear-MarketOfferIconCache') { Clear-MarketOfferIconCache }
    }
    catch { }

    try
    {
        if (Test-DashboardCommand 'Refresh-MarketHistoryGrid') { Refresh-MarketHistoryGrid }
    }
    catch { }

    try
    {
        if (Test-DashboardCommand 'Set-MarketCurrencyComboIcons') { Set-MarketCurrencyComboIcons }
    }
    catch { }

    foreach ($n in @('MarketHistoryGrid', 'MarketLiveGrid'))
    {
        $g = $null
        try { $g = Get-MarketUIControl -Name $n } catch { $g = $null }
        if ($null -ne $g)
        {
            try { $g.Invalidate() } catch { }
            try { $g.Update() } catch { }
        }
    }
}

function Stop-MarketIconMapLoad
{
    [CmdletBinding()]
    param()

    $t = $script:MarketIconMapTimer
    $script:MarketIconMapTimer = $null
    if ($null -ne $t)
    {
        try { $t.Stop() } catch { }
        try { $t.Dispose() } catch { }
    }

    $ps = $script:MarketIconMapPowerShell
    $rs = $script:MarketIconMapRunspace
    $script:MarketIconMapPowerShell = $null
    $script:MarketIconMapRunspace = $null
    $script:MarketIconMapHandle = $null
    $script:MarketIconMapLoading = $false
    try { if ($null -ne $ps) { $ps.Dispose() } } catch { }
    try { if ($null -ne $rs) { $rs.Dispose() } } catch { }
}

function Get-MarketItemIconPath
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Name = '')

    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $map = $null
    try { $map = Get-MarketItemIconMap } catch { $map = $null }
    if ($null -eq $map) { return '' }
    $p = ''
    try { $null = $map.TryGetValue($Name.Trim(), [ref]$p) } catch { $p = '' }
    if ($null -eq $p) { $p = '' }
    if ("$p" -eq '') { return '' }
    $resolved = ''
    try { $resolved = Resolve-MarketIconPath -Path $p } catch { $resolved = $p }
    if ("$resolved" -ne '') { $p = "$resolved" }
    return [string]$p
}

function Get-MarketItemIconImage
{
    [CmdletBinding()]
    [OutputType([System.Drawing.Bitmap])]
    param([string]$Name = '')

    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $path = Get-MarketItemIconPath -Name $Name
    if ([string]::IsNullOrWhiteSpace($path)) { return $null }
    if ($script:MarketItemImages.ContainsKey($path)) { return $script:MarketItemImages[$path] }

    $bmp = $null
    $exists = $false
    try { $exists = (Test-Path -LiteralPath $path -PathType Leaf) } catch { $exists = $false }
    if ($exists)
    {
        try
        {
            $bytes = [System.IO.File]::ReadAllBytes($path)
            $ms = [System.IO.MemoryStream]::new($bytes)
            try
            {
                $img = [System.Drawing.Image]::FromStream($ms)
                try { $bmp = [System.Drawing.Bitmap]::new($img) }
                finally { try { $img.Dispose() } catch { } }
            }
            finally { try { $ms.Dispose() } catch { } }
        }
        catch { $bmp = $null; Write-Verbose "MARKET: icon load failed for '$Name': $($_.Exception.Message)" }
    }

    $script:MarketItemImages[$path] = $bmp
    return $bmp
}

function Draw-MarketSuggestionItem
{
    [CmdletBinding()]
    param(
        $Sender,
        $EventArgs
    )

    $e = $EventArgs
    if ($null -eq $e -or $null -eq $Sender) { return }
    $g = $null
    try { $g = $e.Graphics } catch { $g = $null }
    if ($null -eq $g) { return }
    if ($e.Index -lt 0) { return }

    $back = $null
    $fore = $null
    try { $back = $Sender.BackColor } catch { }
    try { $fore = $Sender.ForeColor } catch { }

    $sel = $false
    try { $sel = (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -ne 0) } catch { $sel = $false }
    if ($sel)
    {
        $c1 = $null
        try { $c1 = Get-Theme 'GridSelectedRow' } catch { $c1 = $null }
        if ($c1 -is [System.Drawing.Color]) { $back = $c1 } else { $back = [System.Drawing.SystemColors]::Highlight }
        $c2 = $null
        try { $c2 = Get-Theme 'GridSelectedText' } catch { $c2 = $null }
        if ($c2 -is [System.Drawing.Color]) { $fore = $c2 } else { $fore = [System.Drawing.SystemColors]::HighlightText }
    }
    if ($null -eq $back) { $back = [System.Drawing.SystemColors]::Window }
    if ($null -eq $fore) { $fore = [System.Drawing.SystemColors]::ControlText }

    $bBack = $null
    $bFore = $null
    try
    {
        $bBack = [System.Drawing.SolidBrush]::new($back)
        $bFore = [System.Drawing.SolidBrush]::new($fore)
        $r = $e.Bounds
        $g.FillRectangle($bBack, $r)

        $txt = ''
        try
        {
            if ($e.Index -ge 0 -and $e.Index -lt $Sender.Items.Count) { $txt = [string]$Sender.Items[$e.Index] }
        }
        catch { $txt = '' }

        $tx = $r.X + 4
        $img = $null
        if ($txt -ne '') { try { $img = Get-MarketItemIconImage -Name $txt } catch { $img = $null } }
        if ($null -ne $img)
        {
            $iy = $r.Y + [int](($r.Height - 16) / 2)
            if ($iy -lt $r.Y) { $iy = $r.Y }
            try { $g.DrawImage($img, $r.X + 4, $iy, 16, 16) } catch { }
            $tx = $r.X + 24
        }

        $font = $null
        try { $font = $Sender.Font } catch { $font = $null }
        if ($txt -ne '' -and $null -ne $font)
        {
            $w = $r.Right - $tx - 4
            if ($w -lt 8) { $w = 8 }
            $rect = New-Object System.Drawing.RectangleF ([float]$tx, [float]$r.Y, [float]$w, [float]$r.Height)
            $fit = $font
            try { $fit = Get-ThemedFitFont -Text $txt -MaxWidth ([int]$w) -Base $font } catch { $fit = $font }
            if ($null -eq $fit) { $fit = $font }
            try { $g.DrawString($txt, $fit, $bFore, $rect) }
            catch { try { $g.DrawString($txt, $font, $bFore, $rect) } catch { } }
        }
    }
    catch { Write-Verbose "MARKET: suggestion draw failed: $($_.Exception.Message)" }
    finally
    {
        if ($null -ne $bBack) { try { $bBack.Dispose() } catch { } }
        if ($null -ne $bFore) { try { $bFore.Dispose() } catch { } }
    }
}

#endregion
#region Small UI helpers

function Get-MarketUIControl
{
    [CmdletBinding()]
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $ui = $global:DashboardConfig.UI
    if ($null -eq $ui) { return $null }

    if ($ui -is [System.Collections.IDictionary])
    {
        try { if ($ui.Contains($Name)) { return $ui[$Name] } } catch { }
        return $null
    }
    try
    {
        $prop = $ui.PSObject.Properties[$Name]
        if ($null -ne $prop) { return $prop.Value }
    }
    catch { }
    return $null
}

function Set-MarketFeedback
{
    [CmdletBinding()]
    param(
        [string]$Text = '',
        [ValidateSet('Info', 'Success', 'Error')]
        [string]$Type = 'Info'
    )

    $lbl = Get-MarketUIControl -Name 'MarketFeedback'
    if ($null -eq $lbl) { return }
    try
    {
        Set-ThemedFitText -Control $lbl -Text $Text
        $key = 'SecondaryText'
        if ($Type -eq 'Success') { $key = 'StatusActive' }
        elseif ($Type -eq 'Error') { $key = 'AccentDanger' }
        if (Test-DashboardCommand 'Get-Theme')
        {
            $c = Get-Theme $key
            if ($c -is [System.Drawing.Color]) { $lbl.ForeColor = $c }
        }
    }
    catch { Write-Verbose "MARKET: feedback update failed: $($_.Exception.Message)" }
}

function Set-MarketCountTip
{
    [CmdletBinding()]
    param([string]$Text = '')

    $cnt = $null
    try { $cnt = Get-MarketUIControl -Name 'MarketHistoryCount' } catch { $cnt = $null }
    if ($null -eq $cnt) { return }

    $tip = $null
    try
    {
        $ui = $global:DashboardConfig.UI
        if ($null -ne $ui)
        {
            foreach ($n in @('ToolTipFtool', 'ToolTip', 'ToolTipMain'))
            {
                $t = $null
                try { $t = $ui.$n } catch { $t = $null }
                if ($null -ne $t) { $tip = $t; break }
            }
        }
    }
    catch { $tip = $null }
    if ($null -eq $tip) { return }

    try { $tip.SetToolTip($cnt, "$Text") } catch { }
}

function Limit-MarketText
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Text = '', [int]$Max = 40)
    $t = "$Text"
    if ($t.Length -le $Max) { return $t }
    $cut = $t.Substring(0, [Math]::Max(1, ($Max - 3)))
    $sp = $cut.LastIndexOf(' ')
    if ($sp -ge ([int](($Max - 3) / 2))) { $cut = $cut.Substring(0, $sp) }
    return ($cut.TrimEnd() + '...')
}

#endregion
#region View updaters

function Set-MarketIcons
{
    [CmdletBinding()]
    param()
    $map = @{
        MarketFindBtn = 'Search'
        MarketAddBtn  = 'Plus'
    }
    foreach ($k in $map.Keys)
    {
        $b = Get-MarketUIControl -Name $k
        if ($null -eq $b) { continue }
        try { $b.Image = Get-MarketIcon -Kind $map[$k] -Size 14 } catch { }
    }
}

function Update-MarketAccessUI
{
    [CmdletBinding()]
    param()
    $unlocked = Test-MarketAccess

    $payloadActive = $false
    try { if (Test-DashboardCommand 'Test-MarketOnlineActive') { $payloadActive = [bool](Test-MarketOnlineActive) } } catch { $payloadActive = $false }

    $status = Get-MarketUIControl -Name 'MarketStatus'
    if ($null -ne $status)
    {
        try
        {
            if ($unlocked) { $status.Text = 'online' } else { $status.Text = 'local' }
            if (Test-DashboardCommand 'Get-Theme')
            {
                $key = 'SecondaryText'
                if ($unlocked) { $key = 'StatusActive' }
                $c = Get-Theme $key
                if ($c -is [System.Drawing.Color]) { $status.ForeColor = $c }
            }
        }
        catch { }
    }

    $hint = Get-MarketUIControl -Name 'MarketScopeHint'
    if ($null -ne $hint)
    {
        try { if ($unlocked) { $hint.Text = 'live offers' } else { $hint.Text = 'journal only' } } catch { }
    }

    $combo = Get-MarketUIControl -Name 'MarketTrendRange'
    $unit = Get-MarketUIControl -Name 'MarketRangeUnit'
    if ($unlocked)
    {
        if ($null -ne $combo) { try { $combo.Visible = $true } catch { } }
        if ($null -ne $unit)
        {
            try { $unit.Text = 'sales' } catch { }
            if (-not $payloadActive) { try { $unit.Left = 354 } catch { } }
        }
    }
    else
    {
        if ($null -ne $combo) { try { $combo.Visible = $false } catch { } }
        if ($null -ne $unit) { try { $unit.Text = '5 sales'; $unit.Left = 310 } catch { } }
    }
    $sellerLbl = Get-MarketUIControl -Name 'MarketSeller'
    if ($null -ne $sellerLbl) { try { $sellerLbl.Visible = [bool]$unlocked } catch { } }

    $syncBtn = Get-MarketUIControl -Name 'MarketSyncBtn'
    if ($null -ne $syncBtn) { try { $syncBtn.Visible = [bool]$unlocked } catch { } }

    $btn = Get-MarketUIControl -Name 'MarketUnlockBtn'
    if ($null -ne $btn)
    {
        try
        {
            if ($unlocked) { $btn.Text = 'Lock' } else { $btn.Text = 'Unlock' }
            $kind = 'Lock'
            if ($unlocked) { $kind = 'Unlock' }
            $btn.Image = Get-MarketIcon -Kind $kind -Size 14
        }
        catch { }
    }
}

function Hide-MarketSuggestions
{
    [CmdletBinding()]
    param()
    $lst = Get-MarketUIControl -Name 'MarketSuggestions'
    if ($null -ne $lst) { try { $lst.Visible = $false } catch { } }
    Sync-MarketViewForCurrentName
}

function Hide-MarketSuggestionsIfOutside
{
    [CmdletBinding()]
    param()
    $lst = Get-MarketUIControl -Name 'MarketSuggestions'
    if ($null -eq $lst) { return }
    try { if (-not $lst.Visible) { return } } catch { return }
    try
    {
        $rect = $lst.RectangleToScreen($lst.ClientRectangle)
        $cur = [System.Windows.Forms.Cursor]::Position
        if (-not $rect.Contains($cur)) { Hide-MarketSuggestions }
    }
    catch { Hide-MarketSuggestions }
}

function Sync-MarketViewForCurrentName
{
    [CmdletBinding()]
    param()

    $name = ""
    try { $name = Get-MarketViewName }
    catch { try { $name = Get-MarketSelectedName } catch { $name = "" } }
    $name = [string]$name

    $term = ''
    try
    {
        $txt = Get-MarketUIControl -Name 'MarketSearchInput'
        if ($null -ne $txt) { $term = ([string]$txt.Text).Trim() }
    }
    catch { $term = '' }

    $last = $null
    try { $last = $script:MarketLastViewName } catch { $last = $null }
    $lastTerm = $null
    try { $lastTerm = $script:MarketLastFilterText } catch { $lastTerm = $null }

    $sameName = ($null -ne $last -and [string]::Equals([string]$last, $name, [System.StringComparison]::Ordinal))
    $sameTerm = ($null -ne $lastTerm -and [string]::Equals([string]$lastTerm, $term, [System.StringComparison]::Ordinal))
    if ($sameName -and $sameTerm) { return }

    try { Update-MarketView } catch { Write-Verbose "MARKET: view sync failed: $($_.Exception.Message)" }
}

function Update-MarketSuggestions
{
    [CmdletBinding()]
    param()

    $state = $global:DashboardConfig.State
    if ($null -ne $state -and $state['MarketTextLock']) { return }

    if ($script:MarketPickedName -ne '') { Clear-MarketPickedItem }

    if ($null -eq $script:MarketSuggestTimer)
    {
        $t = New-Object System.Windows.Forms.Timer
        $t.Interval = 160
        $t.Add_Tick({
                try { Invoke-MarketSuggestionsNow } catch { Write-Verbose "MARKET: $($_.Exception.Message)" }
            })
        $script:MarketSuggestTimer = $t
    }
    $script:MarketSuggestTimer.Stop()
    $script:MarketSuggestTimer.Start()
}

function Flush-MarketSuggestions
{
    $t = $script:MarketSuggestTimer
    if ($null -eq $t -or -not $t.Enabled) { return }
    $t.Stop()
    try { Invoke-MarketSuggestionsNow } catch { Write-Verbose "MARKET: $($_.Exception.Message)" }
}

function Invoke-MarketSuggestionsNow
{
    [CmdletBinding()]
    param()
    $t = $script:MarketSuggestTimer
    if ($null -ne $t) { $t.Stop() }

    $txt = Get-MarketUIControl -Name 'MarketSearchInput'
    $lst = Get-MarketUIControl -Name 'MarketSuggestions'
    if ($null -eq $txt -or $null -eq $lst) { return }

    $state = $global:DashboardConfig.State
    if ($null -ne $state -and $state['MarketTextLock']) { return }

    $q = "$($txt.Text)"
    if ([string]::IsNullOrWhiteSpace($q))
    {
        Hide-MarketSuggestions
        return
    }

    $items = @()
    try { $items = ConvertTo-MarketStringList -Value (Get-MarketSuggestions -Query $q -Limit 50 -IncludeOnline:(Test-MarketAccess)) }
    catch { $items = @() }

    $script:MarketLastSuggestions = $items
    $script:MarketLastQuery = $q.Trim()

    if ($items.Count -eq 0)
    {
        Hide-MarketSuggestions
        return
    }

    try
    {
        $lst.BeginUpdate()
        try
        {
            $lst.Items.Clear()
            foreach ($n in $items) { [void]$lst.Items.Add([string]$n) }
        }
        finally { $lst.EndUpdate() }

        $lst.SelectedIndex = -1
        $lst.Visible = $true
        $lst.BringToFront()
    }
    catch { Write-Verbose "MARKET: suggestion list update failed: $($_.Exception.Message)" }

    Sync-MarketViewForCurrentName
}

function Accept-MarketSelection
{
    [CmdletBinding()]
    param([int]$Index = -1)

    Flush-MarketSuggestions
    $txt = Get-MarketUIControl -Name 'MarketSearchInput'
    $lst = Get-MarketUIControl -Name 'MarketSuggestions'
    if ($null -eq $txt -or $null -eq $lst) { return }
    if ($lst.Items.Count -eq 0) { Hide-MarketSuggestions; return }

    $i = $Index
    if ($i -lt 0) { $i = $lst.SelectedIndex }
    if ($i -lt 0) { $i = 0 }
    if ($i -ge $lst.Items.Count) { $i = $lst.Items.Count - 1 }

    $name = [string]$lst.Items[$i]
    $state = $global:DashboardConfig.State
    if ($null -ne $state) { $state['MarketTextLock'] = $true }
    try
    {
        $txt.Text = $name
        try { $lst.SelectedIndex = $i } catch { }
    }
    finally
    {
        if ($null -ne $state) { $state['MarketTextLock'] = $false }
    }

    $script:MarketLastSuggestions = @()
    $script:MarketLastQuery = ''
    try { $lst.Visible = $false } catch { }
    try
    {
        $txt.SelectionStart = $txt.Text.Length
        $txt.SelectionLength = 0
        [void]$txt.Focus()
    }
    catch { }

    Clear-MarketPickedItem
    Update-MarketView
}

function Move-MarketSuggestion
{
    [CmdletBinding()]
    param([int]$Delta = 1)

    Flush-MarketSuggestions
    $lst = Get-MarketUIControl -Name 'MarketSuggestions'
    if ($null -eq $lst) { return }
    try { if (-not $lst.Visible -or $lst.Items.Count -eq 0) { return } } catch { return }

    $i = $lst.SelectedIndex + $Delta
    if ($i -lt 0) { $i = 0 }
    if ($i -ge $lst.Items.Count) { $i = $lst.Items.Count - 1 }
    try
    {
        $lst.SelectedIndex = $i
        $lst.TopIndex = $i
    }
    catch { }
}

function Select-MarketSuggestionAt
{
    [CmdletBinding()]
    param([int]$X = 0, [int]$Y = 0)

    Flush-MarketSuggestions
    $lst = Get-MarketUIControl -Name 'MarketSuggestions'
    if ($null -eq $lst) { return }
    $idx = -1
    try { $idx = $lst.IndexFromPoint($X, $Y) } catch { $idx = -1 }
    if ($idx -lt 0 -or $idx -ge $lst.Items.Count)
    {
        Hide-MarketSuggestions
        return
    }
    Accept-MarketSelection -Index $idx
}

function Get-MarketSelectedName
{
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $txt = Get-MarketUIControl -Name 'MarketSearchInput'
    $raw = ''
    if ($null -ne $txt) { try { $raw = [string]$txt.Text } catch { } }
    $raw = "$raw".Trim()
    if ($raw -eq '') { return '' }

    $lst = Get-MarketUIControl -Name 'MarketSuggestions'
    if ($null -ne $lst)
    {
        try
        {
            if ($lst.Visible -and $lst.SelectedIndex -ge 0 -and $lst.SelectedIndex -lt $lst.Items.Count)
            {
                return ([string]$lst.Items[$lst.SelectedIndex]).Trim()
            }
        }
        catch { }
    }

    try
    {
        $cat = Get-MarketCatalog
        if ($cat -and $cat.ByName)
        {
            $hit = $null
            if ($cat.ByName.TryGetValue($raw, [ref]$hit) -and $null -ne $hit)
            {
                $matched = [string]$hit.name
                if (-not [string]::IsNullOrWhiteSpace($matched)) { return $matched }
            }
        }
    }
    catch { }

    $names = @()
    try { $names = ConvertTo-MarketStringList -Value (Get-MarketItemNames -IncludeOnline:(Test-MarketAccess)) }
    catch { $names = @() }
    foreach ($n in $names)
    {
        if (-not [string]::IsNullOrWhiteSpace($n) -and [string]::Equals($n, $raw, [System.StringComparison]::OrdinalIgnoreCase))
        {
            return [string]$n
        }
    }

    if ($script:MarketLastSuggestions.Count -gt 0 -and
        [string]::Equals([string]$script:MarketLastQuery, $raw, [System.StringComparison]::OrdinalIgnoreCase))
    {
        return ([string]$script:MarketLastSuggestions[0]).Trim()
    }

    return $raw
}

function Get-MarketViewName
{
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $picked = ''
    try { $picked = "$script:MarketPickedName" } catch { $picked = '' }
    if ($null -eq $picked) { $picked = '' }
    $picked = $picked.Trim()
    if ($picked -ne '') { return $picked }

    $n = ''
    try { $n = Get-MarketSelectedName } catch { $n = '' }
    return ([string]$n)
}

function Clear-MarketPickedItem
{
    [CmdletBinding()]
    param()
    $script:MarketPickedName = ''
    $script:MarketPickedTicks = [long]0
    $script:MarketPickedPrice = 0.0
    $script:MarketPickedFrom = ''
    $script:MarketPickedSeller = ''
}

function Select-MarketItem
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [long]$Ticks = [long]0,
        [double]$Price = 0.0,
        [string]$Seller = '',
        [ValidateSet('history', 'live')]
        [string]$From = 'history'
    )

    $n = "$Name".Trim()
    if ($n -eq '') { return }

    $script:MarketPickedName = $n
    $script:MarketPickedTicks = $Ticks
    $script:MarketPickedPrice = $Price
    $script:MarketPickedFrom = $From
    $script:MarketPickedSeller = "$Seller".Trim()
    $script:MarketLastViewName = $n

    try { Update-MarketTrendDisplay } catch { Write-Verbose "MARKET: trend follow failed: $($_.Exception.Message)" }
    if (Test-DashboardCommand 'Update-MarketOnlineDetail')
    {
        try { Update-MarketOnlineDetail } catch { Write-Verbose "MARKET: info follow failed: $($_.Exception.Message)" }
    }
}

function Get-MarketCurrencyList
{
    [CmdletBinding()]
    [OutputType([string[]])]
    param()
    return @('Penya', 'Perin', 'Entropia Chip', 'Entropia Chip Perin', 'Red Diamond')
}

function Get-MarketCurrencyDisplayMap
{
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return @{
        'Entropia Chip Perin' = 'ECP'
    }
}

function Get-MarketCurrencyFullName
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Display = '')
    $d = "$Display".Trim()
    if ($d -eq '') { return '' }
    if ($d -eq 'ECP') { return 'Entropia Chip Perin' }
    return $d
}

function Get-MarketCurrencyIcon
{
    [CmdletBinding()]
    [OutputType([System.Drawing.Bitmap])]
    param([string]$Name = '')
    $n = "$Name".Trim()
    if ($n -eq '') { return $null }
    try { return Get-MarketItemIconImage -Name $n } catch { return $null }
}

function Set-MarketCurrencyComboIcons
{
    [CmdletBinding()]
    param()
    $combo = Get-MarketUIControl -Name 'MarketCurCombo'
    if ($null -eq $combo) { return }
    try
    {
        $imgs = $combo.ItemImages
        $names = $combo.DisplayNames
        if ($null -eq $imgs -or $null -eq $names) { return }
        $imgs.Clear()
        $names.Clear()

        $map = Get-MarketCurrencyDisplayMap
        for ($i = 0; $i -lt $combo.Items.Count; $i++)
        {
            $item = "$($combo.Items[$i])"
            $full = Get-MarketCurrencyFullName -Display $item
            if ($full -ne '')
            {
                $ic = $null
                try { $ic = Get-MarketCurrencyIcon -Name $full } catch { $ic = $null }
                if ($null -ne $ic) { $imgs[$i] = $ic }
            }
            if ($null -ne $map -and $map.ContainsKey($item)) { $names[$item] = "$($map[$item])" }
        }

        try { $combo.DropDownWidth = $combo.Width } catch { }
        try { $combo.Invalidate() } catch { }
    }
    catch { Write-Verbose "MARKET: currency icons failed: $($_.Exception.Message)" }
}

function Set-MarketAddCurrency
{
    [CmdletBinding()]
    param()
    if ($script:MarketInitializing) { return }

    $combo = Get-MarketUIControl -Name 'MarketCurCombo'
    if ($null -eq $combo) { return }

    $txt = ''
    try { $txt = "$($combo.Text)".Trim() } catch { return }
    $txt = Get-MarketCurrencyFullName -Display $txt
    if ($txt -eq 'Perin') { $txt = 'Perin' }

    $cur = Get-MarketConfigValue -Key 'AddCurrency' -Default 'Perin'
    if ($cur -eq $txt) { return }
    Set-MarketConfigValue -Key 'AddCurrency' -Value "$txt"
    if (Test-DashboardCommand 'WriteConfig')
    {
        try { WriteConfig } catch { }
    }
}

function Set-MarketTrendRange
{
    [CmdletBinding()]
    param()
    if ($script:MarketInitializing) { return }

    $combo = Get-MarketUIControl -Name 'MarketTrendRange'
    if ($null -eq $combo) { return }

    $txt = ''
    try { $txt = "$($combo.Text)" } catch { return }
    $n = 0
    if (-not [int]::TryParse($txt, [ref]$n)) { return }
    if ($n -lt 1 -or $n -gt 50) { return }

    $cur = Get-MarketConfigValue -Key 'TrendRange' -Default '25'
    if ($cur -eq "$n")
    {
        Update-MarketTrendDisplay
        return
    }
    Set-MarketConfigValue -Key 'TrendRange' -Value "$n"
    if (Test-DashboardCommand 'WriteConfig')
    {
        try { WriteConfig } catch { }
    }
    Update-MarketTrendDisplay
}

function Update-MarketTrendDisplay
{
    [CmdletBinding()]
    param([string]$Name = '')

    $arrow = Get-MarketUIControl -Name 'MarketTrendArrow'
    $text = Get-MarketUIControl -Name 'MarketTrendText'
    $detail = Get-MarketUIControl -Name 'MarketTrendDetail'
    if ($null -eq $arrow) { return }

    $inv = [System.Globalization.CultureInfo]::InvariantCulture

    $range = 25
    if (Test-MarketAccess)
    {
        $rawRange = Get-MarketConfigValue -Key 'TrendRange' -Default '25'
        if (-not [int]::TryParse($rawRange, [ref]$range)) { $range = 25 }
        if ($range -lt 1) { $range = 1 }
        if ($range -gt 50) { $range = 50 }
    }
    else
    {
        $range = 5
    }

    $name = "$Name"
    if ($name -eq '') { $name = Get-MarketViewName }
    $entries = @()
    if ($name -ne '')
    {
        try { $entries = ConvertTo-MarketEntryList -Value (Get-MarketItemEntries -Name $name) } catch { $entries = @() }
        if ($null -eq $entries) { $entries = @() }
    }

    $prices = @()
    foreach ($e in $entries)
    {
        if ($null -eq $e) { continue }
        $p = $null
        try { $p = $e['price'] } catch { $p = $null }
        if ($null -ne $p) { $prices += [double]$p }
    }

    $t = $null
    try { $t = Get-MarketTrend -Prices $prices -Range $range } catch { $t = $null }

    $arrowText = [string][char]0x2192
    $arrowColor = 'SecondaryText'
    $line = 'No price data yet'
    $detailLine = ''

    if ($null -eq $t)
    {
        $line = 'Trend unavailable'
    }
    elseif ($t.Count -eq 0)
    {
        if ($name -ne '') { $detailLine = "Record a price for '$name' to start a trend." }
        else { $detailLine = 'Search for an item to see its trend.' }
    }
    elseif ($t.Count -lt 2)
    {
        $arrowText = $t.Arrow
        $line = 'Not enough data yet'
        $detailLine = 'At least 2 recorded prices are needed for a trend.'
    }
    else
    {
        $arrowText = $t.Arrow
        $sign = ''
        if ($t.DeltaPct -gt 0) { $sign = '+' }
        elseif ($t.DeltaPct -lt 0) { $sign = '-' }
        $num = [Math]::Abs([double]$t.DeltaPct).ToString('0.00', $inv)
        $deltaStr = "$sign$num%"

        if ($t.Direction -eq 'up') { $arrowColor = 'StatusActive'; $line = "Rising $deltaStr" }
        elseif ($t.Direction -eq 'down') { $arrowColor = 'AccentDanger'; $line = "Falling $deltaStr" }
        else { $arrowColor = 'SecondaryText'; $line = "Steady $deltaStr" }

        $older = [double]$t.OlderAvg
        $recent = [double]$t.RecentAvg
        $detailLine = "Window $($t.Used) of $($t.Count) sales - older $([string]::Format($inv, '{0:0.##}', $older)) / recent $([string]::Format($inv, '{0:0.##}', $recent))"
    }

    try
    {
        $arrow.Text = $arrowText
        $text.Text = $line
        $detail.Text = $detailLine
        if (Test-DashboardCommand 'Get-Theme')
        {
            $c = Get-Theme $arrowColor
            if ($c -is [System.Drawing.Color]) { $arrow.ForeColor = $c }
        }
    }
    catch { Write-Verbose "MARKET: trend display update failed: $($_.Exception.Message)" }
}

function Fill-MarketHistoryRows
{
    [CmdletBinding()]
    param(
        $Items,
        $Table,
        [long]$Seq = 0,
        $Invariant = ([System.Globalization.CultureInfo]::InvariantCulture)
    )

    foreach ($k in $Items.Keys)
    {
        $prices = $Items[$k]['prices']
        if ($null -eq $prices) { continue }
        $name = [string]$k
        $key = $name.ToLowerInvariant()

        foreach ($e in $prices)
        {
            if ($null -eq $e) { continue }
            $ts = ''; $src = ''; $pr = 0.0; $us = ''; $cu = ''
            try { $ts = [string]$e['ts'] } catch { $ts = '' }
            try { $src = [string]$e['source'] } catch { $src = '' }
            try { $pr = [double]$e['price'] } catch { $pr = 0.0 }
            try { $us = [string]$e['user'] } catch { $us = '' }
            if ([string]::IsNullOrWhiteSpace($us)) { try { $us = [string]$e['seller'] } catch { $us = '' } }
            try
            {
                $rawCur = $e['currency']
                if ($null -ne $rawCur) { $cu = ([string]$rawCur).Trim() }
            }
            catch { $cu = '' }
            if ($us.Trim() -eq '') { $us = '-' } else { $us = $us.Trim() }
            if ($src -eq '') { $src = 'local' }

            $when = [System.DBNull]::Value
            $tk = [long]0
            if (-not [string]::IsNullOrWhiteSpace($ts))
            {
                $d = $null
                try { $d = [datetime]::ParseExact($ts, 'yyyy-MM-dd HH:mm:ss', $Invariant) } catch { $d = $null }
                if ($null -eq $d) { try { $d = [datetime]::Parse($ts, $Invariant) } catch { $d = $null } }
                if ($null -ne $d) { $when = $d; $tk = $d.Ticks }
            }

            $row = $Table.NewRow()
            $row['Icon'] = [System.DBNull]::Value
            $row['When'] = $when
            $row['Item'] = $name
            $row['ItemKey'] = $key
            $row['Price'] = $pr
            $row['Currency'] = $(if ($cu -ne '') { $cu } else { [System.DBNull]::Value })
            $row['Seller'] = $us
            $row['Rec'] = [System.DBNull]::Value
            $row['Source'] = $src
            $row['Ticks'] = $tk
            $row['Seq'] = $Seq
            $Table.Rows.Add($row)
            $Seq++
        }
    }
    return $Seq
}

function Get-MarketHistoryTable
{
    [CmdletBinding()]
    param()

    $unlocked = $false
    try { $unlocked = [bool](Test-MarketAccess) } catch { $unlocked = $false }

    $db = $null
    try { $db = Get-MarketLocalData } catch { $db = $null }

    $c = $script:MarketHistCache
    if ($null -ne $c -and
        $null -ne $c.View -and
        $c.Version -eq $script:MarketLocalVersion -and
        $c.Unlocked -eq $unlocked)
    {
        return ,$c.View
    }

    $dt = New-Object System.Data.DataTable 'MarketHistory'
    [void]$dt.Columns.Add('Icon', [object])
    [void]$dt.Columns.Add('When', [datetime])
    [void]$dt.Columns.Add('Item', [string])
    [void]$dt.Columns.Add('ItemKey', [string])
    [void]$dt.Columns.Add('Price', [double])
    [void]$dt.Columns.Add('Currency', [string])
    [void]$dt.Columns.Add('Seller', [string])
    [void]$dt.Columns.Add('Rec', [object])
    [void]$dt.Columns.Add('Source', [string])
    [void]$dt.Columns.Add('Ticks', [long])
    [void]$dt.Columns.Add('Seq', [long])

    $dt.BeginLoadData()
    $seq = [long]0
    if ($null -ne $db -and $null -ne $db.items) { $seq = Fill-MarketHistoryRows -Items $db.items -Table $dt -Seq $seq }
    $dt.EndLoadData()

    $dv = $dt.DefaultView
    $dv.Sort = 'Ticks DESC, Seq DESC'

    $script:MarketHistCache = @{
        Version  = $script:MarketLocalVersion
        Online   = $null
        Unlocked = $unlocked
        Table    = $dt
        View     = $dv
    }
    return ,$dv
}

function Format-MarketHistoryCell
{
    [CmdletBinding()]
    param(
        $Sender,
        $EventArgs
    )

    $e = $EventArgs
    if ($null -eq $e -or $null -eq $Sender) { return }
    if ($e.ColumnIndex -lt 0 -or $e.RowIndex -lt 0) { return }

    $col = $null
    try { $col = $Sender.Columns[$e.ColumnIndex] } catch { $col = $null }
    if ($null -eq $col) { return }
    $cn = [string]$col.Name
    if ($cn -ne 'MarketColIcon' -and $cn -ne 'MarketColRec' -and $cn -ne 'MarketColCur') { return }

    if ($cn -eq 'MarketColCur')
    {
        $cuName = ''
        try
        {
            $cv = $e.Value
            if ($null -ne $cv -and $cv -isnot [System.DBNull]) { $cuName = ([string]$cv).Trim() }
        }
        catch { $cuName = '' }
        if ($cuName -eq '') { $e.Value = [System.DBNull]::Value; return }
        $curImg = $null
        try { $curImg = Get-MarketItemIconImage -Name $cuName } catch { $curImg = $null }
        if ($null -ne $curImg) { $e.Value = $curImg } else { $e.Value = [System.DBNull]::Value }
        return
    }

    $itemName = ''
    try { $itemName = [string]$Sender.Rows[$e.RowIndex].Cells['MarketColItem'].Value } catch { $itemName = '' }
    if ([string]::IsNullOrWhiteSpace($itemName)) { return }

    if ($cn -eq 'MarketColIcon')
    {
        $img = $null
        try { $img = Get-MarketItemIconImage -Name $itemName } catch { $img = $null }
        if ($null -ne $img) { $e.Value = $img }
        return
    }

    $cur = $null
    try { $cur = $e.Value } catch { $cur = $null }
    if ($null -ne $cur -and $cur -isnot [System.DBNull]) { return }

    $unlocked = $false
    try { $unlocked = [bool](Test-MarketAccess) } catch { $unlocked = $false }
    if (-not $unlocked) { return }
    if (-not (Test-DashboardCommand 'Get-MarketItemRecommendation')) { return }

    $txt = ''
    try
    {
        $rr = Get-MarketItemRecommendation -Name $itemName
        if ($null -ne $rr -and $null -ne $rr.Value)
        {
            $txt = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.##}', $rr.Value)
        }
    }
    catch { $txt = '' }
    if ($txt -ne '') { $e.Value = $txt }
}

function ConvertTo-MarketLikePattern
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Text)

    $t = "$Text"
    if ($t -eq '') { return '' }
    $t = $t.Replace('[', '[[]')
    $t = $t.Replace('*', '[*]')
    $t = $t.Replace('?', '[?]')
    $t = $t.Replace('%', '[%]')
    return $t
}

function Refresh-MarketHistoryGrid
{
    [CmdletBinding()]
    param()

    $grid = Get-MarketUIControl -Name 'MarketHistoryGrid'
    if ($null -eq $grid) { return }
    try { if ($grid.IsDisposed) { return } } catch { return }

    $payload = $false
    try { if (Test-DashboardCommand 'Test-MarketOnlineActive') { $payload = [bool](Test-MarketOnlineActive) } } catch { $payload = $false }
    if ($payload)
    {
        $v = ''
        try { $v = [string]$global:DashboardConfig.State['MarketOnlineView'] } catch { $v = '' }
        if ($v -ne 'hist') { return }
    }

    $recCol = $null
    try { $recCol = $grid.Columns['MarketColRec'] } catch { $recCol = $null }
    if ($null -ne $recCol) { try { $recCol.Visible = (Test-MarketAccess) } catch { } }
    $sellerCol = $null
    try { $sellerCol = $grid.Columns['MarketColSeller'] } catch { $sellerCol = $null }
    if ($null -ne $sellerCol) { try { $sellerCol.Visible = (Test-MarketAccess) } catch { } }

    $term = ''
    try
    {
        $txt = Get-MarketUIControl -Name 'MarketSearchInput'
        if ($null -ne $txt) { $term = ([string]$txt.Text).Trim() }
    }
    catch { $term = '' }

    $view = $null
    try { $view = Get-MarketHistoryTable } catch { $view = $null; Write-Verbose "MARKET: history table failed: $($_.Exception.Message)" }
    if ($null -eq $view) { return }

    $filter = ''
    if (-not [string]::IsNullOrWhiteSpace($term))
    {
        $like = ConvertTo-MarketLikePattern -Text $term.ToLowerInvariant()
        $filter = "ItemKey LIKE '*" + $like + "*'"
    }
    try
    {
        if ([string]$view.RowFilter -ne $filter) { $view.RowFilter = $filter }
    }
    catch
    {
        Write-Verbose "MARKET: history filter failed: $($_.Exception.Message)"
        try { $view.RowFilter = '' } catch { }
    }
    $script:MarketLastFilterText = $term

    try
    {
        if (-not [object]::ReferenceEquals($grid.DataSource, $view)) { $grid.DataSource = $view }
        $cnt = Get-MarketUIControl -Name 'MarketHistoryCount'
        if ($null -ne $cnt)
        {
            try { $cnt.Text = ('{0:N0} entries' -f [int]$view.Count) }
            catch { $cnt.Text = "$($view.Count) entries" }
            try { Set-MarketCountTip -Text '' } catch { }
        }
        try { if (Get-Member -InputObject $grid -Name 'RefreshScroll' -MemberType Method) { $grid.RefreshScroll() } } catch { }
    }
    catch { Write-Verbose "MARKET: history grid refresh failed: $($_.Exception.Message)" }
}

function Update-MarketView
{
    [CmdletBinding()]
    param()
    Update-MarketTrendDisplay
    if (Test-DashboardCommand 'Update-MarketOnlineDetail') { try { Update-MarketOnlineDetail } catch { } }
    Refresh-MarketHistoryGrid
    if (Test-DashboardCommand 'Test-MarketOfferView')
    {
        $offerView = $false
        try { $offerView = [bool](Test-MarketOfferView) } catch { $offerView = $false }
        if ($offerView -and (Test-DashboardCommand 'Update-MarketOfferGrid'))
        {
            try { Update-MarketOfferGrid } catch { }
        }
    }
    $script:MarketLastViewName = ""
    try { $script:MarketLastViewName = [string](Get-MarketViewName) } catch { $script:MarketLastViewName = "" }
}

#endregion
#region Actions

function Get-MarketPriceFromUI
{
    [CmdletBinding()]
    param()

    $box = Get-MarketUIControl -Name 'MarketPriceInput'
    $raw = ''
    if ($null -ne $box) { try { $raw = [string]$box.Text } catch { } }
    $raw = "$raw".Trim()

    if ($raw -eq '')
    {
        Set-MarketFeedback -Text 'Enter a price first.' -Type 'Error'
        return $null
    }

    $price = ConvertTo-MarketNumber -Value $raw
    if ($price -le 0)
    {
        Set-MarketFeedback -Text ('Invalid price: ' + $raw) -Type 'Error'
        return $null
    }
    return [double]$price
}

function Add-MarketPriceFromUI
{
    [CmdletBinding()]
    param()

    $name = Get-MarketViewName
    if ($name -eq '')
    {
        Set-MarketFeedback -Text 'Pick an item name first.' -Type 'Error'
        return
    }
    $price = Get-MarketPriceFromUI
    if ($null -eq $price) { return }

    $cu = ''
    try
    {
        $cb = Get-MarketUIControl -Name 'MarketCurCombo'
        if ($null -ne $cb)
        {
            $t = ''
            try { $t = "$($cb.Text)".Trim() } catch { $t = '' }
            $t = Get-MarketCurrencyFullName -Display $t
            if ($t -ne '' -and $t -ne 'Perin') { $cu = $t }
        }
    }
    catch { $cu = '' }

    $ok = $false
    try { $ok = Add-MarketPrice -Name $name -Price $price -Source 'local' -Currency $cu }
    catch { $ok = $false }
    if (-not $ok)
    {
        Set-MarketFeedback -Text 'Could not save that price.' -Type 'Error'
        return
    }

    $box = Get-MarketUIControl -Name 'MarketPriceInput'
    if ($null -ne $box) { try { $box.Text = '' } catch { } }

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $priceStr = [string]::Format($inv, '{0:0.##}', $price)
    $msg = "Saved $priceStr"
    if ($cu -ne '') { $msg = "$msg $cu" }
    Set-MarketFeedback -Text $msg -Type 'Success'
    Update-MarketView
}

function Find-MarketLocalEntry
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [long]$Ticks = [long]0,
        [double]$Price = 0.0
    )

    $n = "$Name".Trim()
    if ($n -eq '') { return $null }

    $db = $null
    try { $db = Get-MarketLocalData } catch { $db = $null }
    if ($null -eq $db -or $null -eq $db.items) { return $null }

    $key = Find-MarketDictKey -Dict $db.items -Name $n
    if ($null -eq $key) { return $null }
    $prices = $null
    try { $prices = $db.items[$key]['prices'] } catch { $prices = $null }
    if ($null -eq $prices) { return $null }

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $idx = -1

    if ($Ticks -gt 0)
    {
        for ($i = 0; $i -lt @($prices).Count; $i++)
        {
            $e = $null
            try { $e = $prices[$i] } catch { $e = $null }
            if ($null -eq $e) { continue }
            $tk = [long]0
            try
            {
                $ts = [string]$e['ts']
                if (-not [string]::IsNullOrWhiteSpace($ts)) { $tk = ([datetime]::Parse($ts, $inv)).Ticks }
            }
            catch { $tk = [long]0 }
            if ($tk -ne $Ticks) { continue }
            if ($Price -gt 0)
            {
                $ep = 0.0
                try { $ep = [double]$e['price'] } catch { $ep = 0.0 }
                if ([Math]::Abs($ep - $Price) -gt 0.0001) { continue }
            }
            $idx = $i
            break
        }
    }

    if ($idx -lt 0 -and $Ticks -le 0)
    {
        $list = @($prices)
        for ($i = $list.Count - 1; $i -ge 0; $i--)
        {
            $e = $null
            try { $e = $list[$i] } catch { $e = $null }
            if ($null -eq $e) { continue }
            if ($Price -gt 0)
            {
                $ep = 0.0
                try { $ep = [double]$e['price'] } catch { $ep = 0.0 }
                if ([Math]::Abs($ep - $Price) -gt 0.0001) { continue }
            }
            $idx = $i
            break
        }
    }

    if ($idx -lt 0) { return $null }

    $p = 0.0
    try { $p = [double]$prices[$idx]['price'] } catch { $p = 0.0 }
    return @{ Key = $key; Index = $idx; Price = $p }
}

function Remove-MarketPriceFromUI
{
    [CmdletBinding()]
    param()

    $name = Get-MarketViewName
    if ($name -eq '')
    {
        Set-MarketFeedback -Text 'Pick an item name first.' -Type 'Error'
        return
    }

    $ticks = [long]0
    try { $ticks = [long]$script:MarketPickedTicks } catch { $ticks = [long]0 }
    $picked = 0.0
    try { $picked = [double]$script:MarketPickedPrice } catch { $picked = 0.0 }
    $from = ''
    try { $from = [string]$script:MarketPickedFrom } catch { $from = '' }

    if ($from -eq 'live')
    {
        $canHide = $false
        try { $canHide = [bool](Test-MarketAccess) } catch { $canHide = $false }
        if (-not $canHide)
        {
            Set-MarketFeedback -Text 'Live offer - not in your local list.' -Type 'Error'
            return
        }
        if (-not (Test-DashboardCommand 'Remove-MarketLiveRow'))
        {
            Set-MarketFeedback -Text 'Could not hide that row.' -Type 'Error'
            return
        }

        $seller = ''
        try { $seller = [string]$script:MarketPickedSeller } catch { $seller = '' }
        $ok = $false
        try { $ok = [bool](Remove-MarketLiveRow -Name $name -Seller $seller) } catch { $ok = $false }
        if (-not $ok)
        {
            Set-MarketFeedback -Text 'Could not hide that row.' -Type 'Error'
            return
        }

        Clear-MarketPickedItem
        Set-MarketFeedback -Text 'Row hidden from the offer list.' -Type 'Success'
        Update-MarketView
        return
    }
    if ($from -ne 'history')
    {
        Set-MarketFeedback -Text 'Click a row in the list first.' -Type 'Error'
        return
    }

    $hit = $null
    try { $hit = Find-MarketLocalEntry -Name $name -Ticks $ticks -Price $picked } catch { $hit = $null }
    if ($null -eq $hit)
    {
        if ($ticks -gt 0)
        {
            Set-MarketFeedback -Text 'That row is not in your local list.' -Type 'Error'
            return
        }
        Set-MarketFeedback -Text "No local price for '$name'." -Type 'Error'
        return
    }

    $db = $null
    try { $db = Get-MarketLocalData } catch { $db = $null }
    if ($null -eq $db -or $null -eq $db.items) { Set-MarketFeedback -Text 'Local list unavailable.' -Type 'Error'; return }

    $ok = $false
    try
    {
        $prices = $db.items[$hit.Key]['prices']
        $prices.RemoveAt([int]$hit.Index)
        $ok = [bool](Save-MarketLocalData -Data $db)
    }
    catch { $ok = $false }
    if (-not $ok)
    {
        Set-MarketFeedback -Text 'Could not remove that price.' -Type 'Error'
        return
    }

    $script:MarketPickedTicks = [long]0
    $script:MarketPickedPrice = 0.0
    $script:MarketPickedFrom = ''

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $priceStr = [string]::Format($inv, '{0:0.##}', $hit.Price)
    Set-MarketFeedback -Text "Removed $priceStr" -Type 'Success'
    Update-MarketView
}

function Invoke-MarketUnlockAction
{
    [CmdletBinding()]
    param()

    if (Test-MarketAccess)
    {
        Lock-MarketAccess
        Set-MarketFeedback -Text 'Access locked.' -Type 'Info'
        Update-MarketAccessUI
        Update-MarketView
        return
    }

    $box = Get-MarketUIControl -Name 'MarketCodeInput'
    $code = ''
    if ($null -ne $box) { try { $code = [string]$box.Text } catch { } }
    $code = "$code".Trim()

    if ($code -eq '')
    {
        $m = 'Enter an access code first.'
        if (Test-DashboardCommand 'Show-DarkMessageBox')
        {
            try { Show-DarkMessageBox -Text $m -Caption 'Market Access' -Buttons 'OK' -Icon 'Warning' | Out-Null } catch { }
        }
        Set-MarketFeedback -Text $m -Type 'Error'
        return
    }

    $r = $null
    try { $r = Unlock-MarketAccess -Code $code }
    catch { $r = [pscustomobject]@{ Ok = $false; Message = "$($_.Exception.Message)" } }

    if ($r.Ok)
    {
        $msg = "$($r.Message)"
        if (Test-DashboardCommand 'ShowToast')
        {
            try { ShowToast -Title 'Market' -Message $msg -Type 'Success' -IgnoreCancellation | Out-Null } catch { }
        }
        Set-MarketFeedback -Text 'Access unlocked.' -Type 'Success'
        Update-MarketAccessUI
        Update-MarketView
    }
    else
    {
        $m = "$($r.Message)"
        if ([string]::IsNullOrWhiteSpace($m)) { $m = 'Access code rejected.' }
        if (Test-DashboardCommand 'Show-DarkMessageBox')
        {
            try { Show-DarkMessageBox -Text $m -Caption 'Market Access' -Buttons 'OK' -Icon 'Warning' | Out-Null } catch { }
        }
        Set-MarketFeedback -Text $m -Type 'Error'
        Update-MarketAccessUI
    }
}

#endregion
#region Initialization

function Get-MarketOfferGridOnScreen
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $payload = $false
    try
    {
        if (Test-DashboardCommand 'Test-MarketOnlineActive') { $payload = [bool](Test-MarketOnlineActive) }
    }
    catch { $payload = $false }
    if (-not $payload) { return $false }

    $v = ''
    try { $v = [string]$global:DashboardConfig.State['MarketOnlineView'] } catch { $v = '' }
    return ($v -ne 'hist')
}

function Get-MarketSearchTerm
{
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $t = ''
    try
    {
        $txt = Get-MarketUIControl -Name 'MarketSearchInput'
        if ($null -ne $txt) { $t = ([string]$txt.Text).Trim() }
    }
    catch { $t = '' }
    return $t
}

function Get-MarketVisibleItemKeys
{
    [CmdletBinding()]
    param()

    $view = $null
    try
    {
        $grid = Get-MarketUIControl -Name 'MarketHistoryGrid'
        if ($null -ne $grid -and $grid.DataSource -is [System.Data.DataView]) { $view = $grid.DataSource }
    }
    catch { $view = $null }
    if ($null -eq $view)
    {
        try { $view = Get-MarketHistoryTable } catch { $view = $null }
    }
    if ($null -eq $view) { return [pscustomobject]@{ Filtered = $false; Keys = @(); Unknown = $true } }

    $filter = ''
    try { $filter = [string]$view.RowFilter } catch { $filter = '' }
    if ([string]::IsNullOrWhiteSpace($filter)) { return [pscustomobject]@{ Filtered = $false; Keys = @(); Unknown = $false } }

    $keys = @{}
    try
    {
        foreach ($rv in $view)
        {
            $ik = ''
            try { $ik = [string]$rv['ItemKey'] } catch { $ik = '' }
            if ($ik -ne '') { $keys[$ik] = $true }
        }
    }
    catch { return [pscustomobject]@{ Filtered = $true; Keys = @(); Unknown = $true } }

    return [pscustomobject]@{ Filtered = $true; Keys = @($keys.Keys); Unknown = $false }
}

function Remove-MarketPriceAllFromUI
{
    [CmdletBinding()]
    param()

    if (Get-MarketOfferGridOnScreen)
    {
        $canHide = $false
        try { $canHide = [bool](Test-MarketAccess) } catch { $canHide = $false }
        if (-not $canHide)
        {
            Set-MarketFeedback -Text 'Live offer - not in your local list.' -Type 'Error'
            return
        }

        $grid = $null
        try { $grid = Get-MarketUIControl -Name 'MarketLiveGrid' } catch { $grid = $null }
        if ($null -eq $grid) { Set-MarketFeedback -Text 'Offer grid not found.' -Type 'Error'; return }

        $want = [System.Collections.Generic.List[object]]::new()
        $rows = @()
        try { $rows = @(foreach ($r in $grid.Rows) { $r }) } catch { $rows = @() }
        foreach ($r in $rows)
        {
            try { if (-not $r.Visible) { continue } } catch { continue }
            $v = $null
            try { $v = $r.Cells['OfferColItem'].Value } catch { $v = $null }
            if ($null -eq $v) { continue }
            $seller = ''
            try { $offer = $r.Tag; if ($null -ne $offer) { $seller = "$($offer.seller)".Trim() } } catch { $seller = '' }
            $want.Add([pscustomobject]@{ Name = [string]$v; Seller = $seller })
        }
        if ($want.Count -eq 0)
        {
            Set-MarketFeedback -Text 'No offer rows to hide.' -Type 'Info'
            return
        }

        $hidden = 0
        if (Test-DashboardCommand 'Hide-MarketLiveRows')
        {
            $hidden = -1
            try { $hidden = [int](Hide-MarketLiveRows -Rows $want) } catch { $hidden = -1 }
        }
        else
        {
            if (-not (Test-DashboardCommand 'Remove-MarketLiveRow'))
            {
                Set-MarketFeedback -Text 'Could not hide those rows.' -Type 'Error'
                return
            }
            foreach ($w in $want)
            {
                $ok = $false
                try { $ok = [bool](Remove-MarketLiveRow -Name $w.Name -Seller $w.Seller) } catch { $ok = $false }
                if ($ok) { $hidden++ }
            }
        }
        if ($hidden -lt 0)
        {
            Set-MarketFeedback -Text 'Could not hide those rows.' -Type 'Error'
            return
        }

        Clear-MarketPickedItem
        Set-MarketFeedback -Text "Hidden $hidden offer row(s)." -Type 'Success'
        Update-MarketView
        return
    }

    $db = $null
    try { $db = Get-MarketLocalData } catch { $db = $null }
    if ($null -eq $db -or $null -eq $db.items) { Set-MarketFeedback -Text 'Local list unavailable.' -Type 'Error'; return }

    $items = $db.items
    if ($items.Count -eq 0)
    {
        Set-MarketFeedback -Text 'The local list is already empty.' -Type 'Info'
        return
    }

    $removed = 0
    $term = Get-MarketSearchTerm

    if ($term -eq '')
    {
        foreach ($k in @($items.Keys))
        {
            $prices = $null
            try { $prices = $items[$k]['prices'] } catch { $prices = $null }
            $removed += @($prices).Count
        }
        try { $items.Clear() } catch { }
    }
    else
    {
        $vis = Get-MarketVisibleItemKeys
        if ($vis.Unknown)
        {
            Set-MarketFeedback -Text 'Could not read the current search filter.' -Type 'Error'
            return
        }
        if (-not $vis.Filtered)
        {
            Set-MarketFeedback -Text 'Search filter not applied yet - try again.' -Type 'Error'
            return
        }

        $targets = [System.Collections.Generic.List[string]]::new()
        foreach ($ik in @($vis.Keys))
        {
            $key = $null
            try { $key = Find-MarketDictKey -Dict $items -Name $ik } catch { $key = $null }
            if ($null -eq $key) { continue }
            if (-not $targets.Contains($key)) { $targets.Add($key) }
        }
        if ($targets.Count -eq 0)
        {
            Set-MarketFeedback -Text 'Nothing to remove.' -Type 'Info'
            return
        }
        foreach ($key in $targets)
        {
            $prices = $null
            try { $prices = $items[$key]['prices'] } catch { $prices = $null }
            $removed += @($prices).Count
            try { [void]$items.Remove($key) } catch { }
        }
    }

    $ok = $false
    try { $ok = [bool](Save-MarketLocalData -Data $db) } catch { $ok = $false }
    if (-not $ok)
    {
        Set-MarketFeedback -Text 'Could not remove those prices.' -Type 'Error'
        return
    }

    Clear-MarketPickedItem
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    Set-MarketFeedback -Text ("Removed {0:N0} price(s)." -f [int]$removed) -Type 'Success'
    Update-MarketView
}

function Get-MarketHeaderControl
{
    [CmdletBinding()]
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }

    $c = $null
    try { $c = Get-MarketUIControl -Name $Name } catch { $c = $null }
    if ($null -ne $c)
    {
        try { if ($c.IsDisposed) { return $null } } catch { return $null }
        return $c
    }

    $tab = $null
    try { $tab = Get-MarketUIControl -Name 'MarketTab' } catch { $tab = $null }
    if ($null -eq $tab) { return $null }
    try
    {
        foreach ($ch in $tab.Controls)
        {
            if ($null -ne $ch -and [string]$ch.Name -eq $Name) { return $ch }
        }
    }
    catch { }
    return $null
}

function Update-MarketHeaderRow
{
    [CmdletBinding()]
    param()

    $right = 385
    $countW = 66
    $gap = 4
    $btnTop = 244
    $btnH = 24

    $lbl = Get-MarketHeaderControl -Name 'MarketHistoryLabel'
    if ($null -ne $lbl) { try { $lbl.Bounds = [System.Drawing.Rectangle]::new(15, 248, 52, 16) } catch { } }

    $cntLeft = $right - $countW
    $cnt = Get-MarketHeaderControl -Name 'MarketHistoryCount'
    if ($null -ne $cnt) { try { $cnt.Bounds = [System.Drawing.Rectangle]::new($cntLeft, 249, $countW, 16) } catch { } }

    $fb = Get-MarketHeaderControl -Name 'MarketFeedback'
    if ($null -ne $fb) { try { $fb.Bounds = [System.Drawing.Rectangle]::new(15, 222, 300, 16) } catch { } }

    $names = @('MarketRemoveBtn', 'MarketRemoveAllBtn', 'MarketOfferCopyAllBtn', 'MarketOfferCopySelBtn')
    $x = $cntLeft - $gap
    foreach ($n in $names)
    {
        $b = Get-MarketHeaderControl -Name $n
        if ($null -eq $b) { continue }

        $w = 60
        try
        {
            $m = [System.Windows.Forms.TextRenderer]::MeasureText([string]$b.Text, $b.Font, [System.Drawing.Size]::new(1000, 1000), [System.Windows.Forms.TextFormatFlags]::NoPadding)
            $w = [int]$m.Width + 10
            if ($w -lt 44) { $w = 44 }
        }
        catch { }

        $x -= $w
        try { $b.Bounds = [System.Drawing.Rectangle]::new($x, $btnTop, $w, $btnH) } catch { }
        $x -= $gap
    }
}

function Ensure-MarketRemoveBtn
{
    [CmdletBinding()]
    param()

    $tab = $null
    try { $tab = Get-MarketUIControl -Name 'MarketTab' } catch { $tab = $null }
    if ($null -eq $tab) { return }

    $btn = $null
    try { $btn = Get-MarketUIControl -Name 'MarketRemoveBtn' } catch { $btn = $null }
    if ($null -ne $btn)
    {
        try { if ($btn.IsDisposed) { $btn = $null } } catch { $btn = $null }
    }
    if ($null -eq $btn)
    {
        $p = @{
            type         = 'Button'
            width        = 70
            height       = 24
            top          = 244
            left         = 210
            id           = 'MarketRemoveBtn'
            text         = 'Remove'
            fs           = 'Flat'
            font         = (Get-UIFont 'Segoe UI' 8 'Bold')
            tooltip      = "Remove`nDelete ONE entry of your local list: the row you clicked in a grid,`nor the newest local price of the item shown. Every other price stays."
            noCustomPaint = $true
        }
        try { $btn = SetUIElement @p } catch { $btn = $null }
        if ($null -eq $btn) { return }

        try { [void]$tab.Controls.Add($btn) } catch { return }

        try
        {
            $ui = $global:DashboardConfig.UI
            if ($null -ne $ui) { $ui | Add-Member -MemberType NoteProperty -Name 'MarketRemoveBtn' -Value $btn -Force }
        }
        catch { }

        try
        {
            $btn.Add_Click({
                    try { Remove-MarketPriceFromUI }
                    catch { Set-MarketFeedback -Text 'Remove failed.' -Type 'Error' }
                })
        }
        catch { }
    }

    try { $btn.Image = $null } catch { }
    try { if (-not $btn.Visible) { $btn.Visible = $true } } catch { }

    $btnAll = $null
    try { $btnAll = Get-MarketUIControl -Name 'MarketRemoveAllBtn' } catch { $btnAll = $null }
    if ($null -ne $btnAll)
    {
        try { if ($btnAll.IsDisposed) { $btnAll = $null } } catch { $btnAll = $null }
    }
    if ($null -eq $btnAll)
    {
        $p = @{
            type         = 'Button'
            width        = 76
            height       = 24
            top          = 244
            left         = 185
            id           = 'MarketRemoveAllBtn'
            text         = 'Remove all'
            fs           = 'Flat'
            font         = (Get-UIFont 'Segoe UI' 8 'Bold')
            tooltip      = "Remove all`nDelete every price the grid shows right now.`nA search narrows it to the rows that match; without one the whole list goes.`nThe picked row stays untouched - that is what Remove is for."
            noCustomPaint = $true
        }
        try { $btnAll = SetUIElement @p } catch { $btnAll = $null }
        if ($null -ne $btnAll)
        {
            try { [void]$tab.Controls.Add($btnAll) } catch { }
            try
            {
                $ui = $global:DashboardConfig.UI
                if ($null -ne $ui) { $ui | Add-Member -MemberType NoteProperty -Name 'MarketRemoveAllBtn' -Value $btnAll -Force }
            }
            catch { }
            try
            {
                $btnAll.Add_Click({
                        try { Remove-MarketPriceAllFromUI }
                        catch { Set-MarketFeedback -Text 'Remove all failed.' -Type 'Error' }
                    })
            }
            catch { }
        }
    }

    try { if ($null -ne $btnAll -and -not $btnAll.Visible) { $btnAll.Visible = $true } } catch { }

    try { Update-MarketHeaderRow } catch { }
}

function Wire-MarketHistoryGrid
{
    [CmdletBinding()]
    param()

    $grid = $null
    try { $grid = Get-MarketUIControl -Name 'MarketHistoryGrid' } catch { $grid = $null }
    if ($null -eq $grid) { return }
    try { if ($grid.IsDisposed) { return } } catch { return }
    if ([object]::ReferenceEquals($script:MarketHistGridWired, $grid)) { return }

    try
    {
        $grid.Add_CellClick({
                param($src, $e)
                try
                {
                    if ($null -eq $e -or $e.RowIndex -lt 0) { return }
                    $g = $src
                    if ($null -eq $g) { return }
                    if ($e.RowIndex -ge $g.Rows.Count) { return }

                    $name = ''
                    try { $name = [string]$g.Rows[$e.RowIndex].Cells['MarketColItem'].Value } catch { $name = '' }
                    if ([string]::IsNullOrWhiteSpace($name)) { return }

                    $ticks = [long]0
                    $price = 0.0
                    try
                    {
                        $row = $g.Rows[$e.RowIndex].DataBoundItem
                        if ($row -is [System.Data.DataRowView])
                        {
                            $ticks = [long]$row['Ticks']
                            $price = [double]$row['Price']
                        }
                    }
                    catch { $ticks = [long]0; $price = 0.0 }

                    Select-MarketItem -Name $name -Ticks $ticks -Price $price
                }
                catch { Write-Verbose "MARKET: history pick failed: $($_.Exception.Message)" }
            })
        $script:MarketHistGridWired = $grid
    }
    catch { Write-Verbose "MARKET: history click wiring failed: $($_.Exception.Message)" }
}

#region Hover description tooltip (guest mode)

function Get-MarketItemDescription
{
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Name = '')

    $n = "$Name".Trim()
    if ($n -eq '') { return '' }
    $cat = $null
    try { $cat = Get-MarketCatalog } catch { $cat = $null }
    if ($null -eq $cat -or $null -eq $cat.ByName) { return '' }
    $hit = $null
    try { $null = $cat.ByName.TryGetValue($n.ToLowerInvariant(), [ref]$hit) } catch { $hit = $null }
    if ($null -eq $hit) { return '' }
    $d = ''
    try { $d = [string]$hit.desc } catch { $d = '' }
    return "$d".Trim()
}

function Initialize-MarketHistTipType
{
    [CmdletBinding()]
    param()
    if ('Entropia.HistHoverForm' -as [type]) { return }
    Add-Type -ReferencedAssemblies System.Windows.Forms -WarningAction SilentlyContinue $init = Get-Command -Name 'Initialize-DashboardMarketTypes' -ErrorAction SilentlyContinue
    if (-not $init)
    {
        Import-Module (Join-Path $PSScriptRoot 'classes.psm1') -Force -DisableNameChecking -ErrorAction Stop
        $init = Get-Command -Name 'Initialize-DashboardMarketTypes' -ErrorAction Stop
    }
    $null = & $init
}

function Get-MarketHistTipBodyFont
{
    [CmdletBinding()]
    param([double]$Size)
    $key = '{0:0.#}' -f $Size
    if (-not $script:MarketHistTipFonts.ContainsKey($key))
    {
        $script:MarketHistTipFonts[$key] = Get-UIFont 'Segoe UI' $Size
    }
    return $script:MarketHistTipFonts[$key]
}

function Get-MarketHistTipForm
{
    [CmdletBinding()]
    param()
    if ($null -ne $script:MarketHistTipForm) { return $script:MarketHistTipForm }

    Initialize-MarketHistTipType

    $bg = Get-Theme 'SecondaryBg'
    if ($bg -isnot [System.Drawing.Color]) { $bg = [System.Drawing.Color]::FromArgb(40, 40, 40) }
    $edge = Get-Theme 'BorderColor'
    if ($edge -isnot [System.Drawing.Color]) { $edge = [System.Drawing.Color]::FromArgb(100, 100, 100) }
    $fg = Get-Theme 'PrimaryText'
    if ($fg -isnot [System.Drawing.Color]) { $fg = [System.Drawing.Color]::FromArgb(240, 240, 240) }
    $fg2 = Get-Theme 'SecondaryText'
    if ($fg2 -isnot [System.Drawing.Color]) { $fg2 = [System.Drawing.Color]::FromArgb(230, 230, 230) }

    $f = New-Object Entropia.HistHoverForm
    $f.Name = 'MarketHistTip'
    $f.BackColor = $edge
    $f.Padding = New-Object System.Windows.Forms.Padding(1)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Name = 'HistTipPanel'
    $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel.BackColor = $bg
    $panel.AutoScroll = $false

    $title = New-Object System.Windows.Forms.Label
    $title.Name = 'HistTipTitle'
    $title.AutoSize = $true
    $title.Font = Get-UIFont 'Segoe UI' 9 'Bold'
    $title.ForeColor = $fg
    $title.BackColor = $bg
    $title.Location = New-Object System.Drawing.Point(9, 7)
    $title.MaximumSize = New-Object System.Drawing.Size(330, 0)

    $body = New-Object System.Windows.Forms.Label
    $body.Name = 'HistTipBody'
    $body.AutoSize = $true
    $body.Font = Get-MarketHistTipBodyFont 8
    $body.ForeColor = $fg2
    $body.BackColor = $bg
    $body.MaximumSize = New-Object System.Drawing.Size(330, 0)

    [void]$panel.Controls.Add($title)
    [void]$panel.Controls.Add($body)
    [void]$f.Controls.Add($panel)
    $f | Add-Member -NotePropertyName 'HistTipTitle' -NotePropertyValue $title
    $f | Add-Member -NotePropertyName 'HistTipBody' -NotePropertyValue $body
    $f | Add-Member -NotePropertyName 'HistTipPanel' -NotePropertyValue $panel

    $script:MarketHistTipForm = $f
    return $f
}

function Show-MarketHistTip
{
    [CmdletBinding()]
    param([int]$RowIndex)

    $grid = Get-MarketUIControl -Name 'MarketHistoryGrid'
    if ($null -eq $grid) { return }
    if ($RowIndex -lt 0 -or $RowIndex -ge $grid.Rows.Count) { Hide-MarketHistTip; return }

    $name = ''
    try
    {
        $ci = $grid.Columns['MarketColItem'].Index
        $name = [string]$grid.Rows[$RowIndex].Cells[$ci].Value
    }
    catch { $name = '' }
    if ([string]::IsNullOrWhiteSpace($name)) { Hide-MarketHistTip; return }

    $desc = Get-MarketItemDescription -Name $name
    if ($desc -eq '') { $desc = 'No description in the ingame item wiki for this item.' }

    $f = Get-MarketHistTipForm
    $title = $f.HistTipTitle
    $body = $f.HistTipBody

    $title.Text = $name
    $body.Text = $desc

    $pt = [System.Windows.Forms.Cursor]::Position
    $wa = [System.Windows.Forms.Screen]::FromPoint($pt).WorkingArea

    $availH = $wa.Height - 24
    $maxWrap = 520
    if ($maxWrap -gt ($wa.Width - 60)) { $maxWrap = $wa.Width - 60 }
    if ($maxWrap -lt 330) { $maxWrap = 330 }
    $wrap = 330
    $fpt = 8.0
    $w = 160
    $th = 0
    $bh = 0
    $h = 0
    while ($true)
    {
        $title.MaximumSize = New-Object System.Drawing.Size($wrap, 0)
        $body.MaximumSize = New-Object System.Drawing.Size($wrap, 0)
        $body.Font = Get-MarketHistTipBodyFont $fpt
        try { $title.PerformLayout(); $body.PerformLayout() } catch { }

        $w = [Math]::Max($title.PreferredSize.Width, $body.PreferredSize.Width)
        if ($w -lt 160) { $w = 160 }
        $th = $title.PreferredSize.Height
        $bh = $body.PreferredSize.Height
        $h = 7 + $th + 5 + $bh + 9
        if ($h -le $availH) { break }
        if ($wrap -lt $maxWrap)
        {
            $wrap = [Math]::Min($maxWrap, ($wrap + 65))
            continue
        }
        if ($fpt -gt 6.5) { $fpt = [Math]::Round(($fpt - 0.5), 1); continue }
        break
    }
    $title.Location = New-Object System.Drawing.Point(9, 7)
    $body.Location = New-Object System.Drawing.Point(9, (7 + $th + 5))
    $f.ClientSize = New-Object System.Drawing.Size(($w + 18), $h)

    $x = $pt.X + 18
    $y = $pt.Y + 22
    if (($x + $f.Width) -gt $wa.Right) { $x = $pt.X - 18 - $f.Width }
    if (($y + $f.Height) -gt $wa.Bottom) { $y = $pt.Y - 22 - $f.Height }
    if ($x -lt $wa.Left) { $x = $wa.Left }
    if ($y -lt $wa.Top) { $y = $wa.Top }
    $f.Location = New-Object System.Drawing.Point($x, $y)

    if (-not $f.Visible) { [void]$f.Show() }
    if ($null -ne $script:MarketHistTipTimer) { $script:MarketHistTipTimer.Stop() }
}

function Hide-MarketHistTip
{
    [CmdletBinding()]
    param()
    try
    {
        if ($null -ne $script:MarketHistTipForm -and $script:MarketHistTipForm.Visible) { $script:MarketHistTipForm.Hide() }
    }
    catch { }
}

function Start-MarketHistTipHide
{
    [CmdletBinding()]
    param()
    if ($null -eq $script:MarketHistTipTimer) { return }
    $script:MarketHistTipTimer.Stop()
    $script:MarketHistTipTimer.Start()
}

function Wire-MarketHistTip
{
    [CmdletBinding()]
    param()

    $grid = $null
    try { $grid = Get-MarketUIControl -Name 'MarketHistoryGrid' } catch { $grid = $null }
    if ($null -eq $grid) { return }
    try { if ($grid.IsDisposed) { return } } catch { return }
    if ($script:MarketHistTipWired -eq $grid) { return }

    $tt = New-Object System.Windows.Forms.Timer
    $tt.Interval = 140
    $tt.Add_Tick({
            try { if ($null -ne $script:MarketHistTipTimer) { $script:MarketHistTipTimer.Stop() } } catch { }
            Hide-MarketHistTip
        })
    $script:MarketHistTipTimer = $tt

    $grid.Add_CellMouseEnter({
            param($src, $e)
            try
            {
                if ($null -ne $e -and $e.RowIndex -ge 0) { Show-MarketHistTip -RowIndex $e.RowIndex }
                else { Hide-MarketHistTip }
            }
            catch { }
        })
    $grid.Add_CellMouseLeave({ param($src, $e) try { Start-MarketHistTipHide } catch { } })
    $grid.Add_MouseLeave({ param($src, $e) try { Start-MarketHistTipHide } catch { } })
    $grid.Add_Scroll({ param($src, $e) try { Hide-MarketHistTip } catch { } })

    $script:MarketHistTipWired = $grid
}

$script:MarketHistTipForm = $null
$script:MarketHistTipTimer = $null
$script:MarketHistTipFonts = @{}
$script:MarketHistTipWired = $null

#endregion

function Initialize-MarketTab
{
    [CmdletBinding()]
    param()

    $state = $global:DashboardConfig.State
    if ($null -eq $state) { return }
    if ($script:MarketInitializing) { return }

    $search = Get-MarketUIControl -Name 'MarketSearchInput'
    if ($null -eq $search) { return }

    $script:MarketInitializing = $true
    try
    {
        Ensure-MarketConfig

        try { Start-MarketIconMapLoad } catch { }

        $combo = Get-MarketUIControl -Name 'MarketTrendRange'
        if ($null -ne $combo)
        {
            try
            {
                if ($combo.Items.Count -eq 0)
                {
                    for ($i = 1; $i -le 50; $i++) { [void]$combo.Items.Add("$i") }
                }
                $range = 25
                $raw = Get-MarketConfigValue -Key 'TrendRange' -Default '25'
                if (-not [int]::TryParse($raw, [ref]$range)) { $range = 25 }
                if ($range -lt 1) { $range = 1 }
                if ($range -gt 50) { $range = 50 }
                $want = "$range"
                if ("$($combo.Text)" -ne $want)
                {
                    try { $combo.SelectedItem = $want } catch { }
                    if ("$($combo.Text)" -ne $want)
                    {
                        try { $combo.Text = $want } catch { }
                    }
                }
            }
            catch { Write-Verbose "MARKET: range combo init failed: $($_.Exception.Message)" }
        }

        $curCombo = Get-MarketUIControl -Name 'MarketCurCombo'
        if ($null -ne $curCombo)
        {
            try
            {
                if ($curCombo.Items.Count -eq 0)
                {
                    [void]$curCombo.Items.Add('Perin')
                    foreach ($c in Get-MarketCurrencyList) { if ($c -ne 'Perin') { [void]$curCombo.Items.Add($c) } }
                    Set-MarketCurrencyComboIcons
                }
                $want = Get-MarketConfigValue -Key 'AddCurrency' -Default 'Perin'
                if ($want -eq '') { $want = 'Perin' }
                $wantDisplay = $want
                $map = Get-MarketCurrencyDisplayMap
                if ($map.ContainsKey($want)) { $wantDisplay = $map[$want] }
                if ($curCombo.Items.Contains($wantDisplay)) { $curCombo.SelectedItem = $wantDisplay }
                elseif ($curCombo.Items.Contains($want)) { $curCombo.SelectedItem = $want }
                elseif ($curCombo.SelectedIndex -lt 0) { $curCombo.SelectedIndex = 0 }
            }
            catch { Write-Verbose "MARKET: currency combo init failed: $($_.Exception.Message)" }
        }

        $codeBox = Get-MarketUIControl -Name 'MarketCodeInput'
        if ($null -ne $codeBox)
        {
            try
            {
                if ([string]::IsNullOrWhiteSpace($codeBox.Text))
                {
                    $saved = Get-MarketConfigValue -Key 'AccessCode'
                    if (-not [string]::IsNullOrWhiteSpace($saved)) { $codeBox.Text = $saved }
                }
            }
            catch { }
        }

        Set-MarketIcons
        Ensure-MarketRemoveBtn
        Wire-MarketHistoryGrid
        Wire-MarketHistTip
        Update-MarketAccessUI
        Update-MarketView
    }
    finally
    {
        $script:MarketInitializing = $false
    }
}

function Open-MarketTab
{
    [CmdletBinding()]
    param()

    $tab = $null
    try { $tab = Get-MarketUIControl -Name 'MarketTab' } catch { $tab = $null }
    if ($null -eq $tab) { return }
    try { if (-not $tab.Visible) { return } } catch { return }

    try { Start-MarketAssetSync }
    catch { Write-Verbose "MARKET: asset sync failed: $($_.Exception.Message)" }
}

#endregion

try { Remove-MarketVaultTempFile } catch { }

try
{
    $syncRoot = Get-MarketDataRoot
    $legacySync = Join-Path $syncRoot 'market_sync.json'
    $movedSync = Join-Path (Join-Path $syncRoot 'market') 'market_sync.json'
    if ((Test-Path -LiteralPath $legacySync -PathType Leaf) -and -not (Test-Path -LiteralPath $movedSync -PathType Leaf))
    {
        $syncDir = Split-Path -Parent $movedSync
        if (-not [string]::IsNullOrWhiteSpace($syncDir) -and -not (Test-Path -LiteralPath $syncDir -PathType Container))
        {
            [void](New-Item -Path $syncDir -ItemType Directory -Force)
        }
        Move-Item -LiteralPath $legacySync -Destination $movedSync -Force
    }
}
catch { }

try { $ExecutionContext.SessionState.Module.OnRemove = { Stop-MarketIconMapLoad; Remove-MarketVaultTempFile } } catch { }

Export-ModuleMember -Function *