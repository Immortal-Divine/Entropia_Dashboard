<# runspace-helper.psm1 #>

function NewManagedRunspace
{
	[CmdletBinding()]
	param(
		[string]$Name = 'ManagedPool',
		[int]$MinRunspaces = 1,
		[int]$MaxRunspaces = 2,
		[hashtable]$SessionVariables = @{},
		[string[]]$Assemblies = @()
	)

	$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
	
	if ($null -ne $global:DashboardConfig)
	{
		$iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry('DashboardConfig', $global:DashboardConfig, 'Main Config', 'None')))
	}

	
	foreach ($entry in $SessionVariables.GetEnumerator())
	{
		$iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry($entry.Key, $entry.Value, '', 0)))
	}

	
	foreach ($asm in $Assemblies)
	{
		if (-not [string]::IsNullOrWhiteSpace($asm))
		{
			$iss.Assemblies.Add((New-Object System.Management.Automation.Runspaces.SessionStateAssemblyEntry($asm)))
		}
	}

	
	@('Custom.Native', 'Custom.Ftool', 'Custom.MouseHookManager') | ForEach-Object {
		$type = $_ -as [Type]
		if ($type -and -not [string]::IsNullOrEmpty($type.Assembly.Location))
		{
			$iss.Assemblies.Add((New-Object System.Management.Automation.Runspaces.SessionStateAssemblyEntry($type.Assembly.Location)))
		}
	}

	$pool = [RunspaceFactory]::CreateRunspacePool($MinRunspaces, $MaxRunspaces, $iss, $Host)
	$pool.ThreadOptions = 'ReuseThread'
	$pool.Open()
	return $pool
}

function InvokeInManagedRunspace
{
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)] $RunspacePool,
		[Parameter(Mandatory = $true)] [ScriptBlock]$ScriptBlock,
		[switch]$AsJob,
		[Alias('Args')]
		[object[]]$ArgumentList
	)

	$ps = [System.Management.Automation.PowerShell]::Create()
	$ps.RunspacePool = $RunspacePool
	$ps.AddScript($ScriptBlock.ToString()) | Out-Null
    
	if ($ArgumentList)
	{ 
		foreach ($arg in $ArgumentList) { $ps.AddArgument($arg) | Out-Null } 
	}

	if ($AsJob)
	{
		return @{ 
			PowerShell   = $ps 
			AsyncResult  = $ps.BeginInvoke()
			RunspacePool = $RunspacePool 
		}
	}
 else
	{
		try { return $ps.Invoke() } finally { $ps.Dispose() }
	}
}

function DisposeManagedRunspace
{
	[CmdletBinding()]
	param([hashtable]$JobResource)
	if ($null -eq $JobResource) { return }
    
	$ps = if ($JobResource.PowerShell) { $JobResource.PowerShell } else { $JobResource.PowerShellInstance }
	$rs = if ($JobResource.RunspacePool) { $JobResource.RunspacePool } else { $JobResource.Runspace }
	$ar = $JobResource.AsyncResult

	try
	{
		if ($ps -and $ps.InvocationStateInfo.State -eq 'Running')
		{
			$ps.Stop()
		}
		if ($ps -and $ar)
		{
			try { $ps.EndInvoke($ar) } catch {}
		}
		if ($ps) { $ps.Dispose() }
		if ($rs) { $rs.Dispose() }
	}
 catch { Write-Verbose "Runspace cleanup error: $($_.Exception.Message)" }
}

Export-ModuleMember -Function *