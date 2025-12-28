<# hotkeys.psm1 #>

if (-not $global:RegisteredHotkeys) { $global:RegisteredHotkeys = @{} }
if (-not $global:RegisteredHotkeyByString) { $global:RegisteredHotkeyByString = @{} }
if (-not $global:PausedRegisteredHotkeys) { $global:PausedRegisteredHotkeys = @{} }
if (-not $global:PausedIdMap) { $global:PausedIdMap = @{} }
if (-not $global:NextPausedFakeId) { $global:NextPausedFakeId = -1 }

function SetHotkey
{
	param(
		[Parameter(Mandatory = $true)]
		[string]$KeyCombinationString,
		[Parameter(Mandatory = $true)]
		$Action,
		[Parameter(Mandatory = $false)]
		[AllowEmptyString()]
		[string]$OwnerKey = $null,
		[Parameter(Mandatory = $false)]
		[int]$OldHotkeyId = $null
	)

	if ($Action -is [System.Action])
	{
		$actionDelegate = $Action
	}
	elseif ($Action -is [System.Delegate])
	{
		$actionDelegate = $Action -as [System.Action]
	}
	else
	{
		$actionDelegate = [System.Action]$Action
	}

	if (-not [string]::IsNullOrEmpty($OwnerKey))
	{
		$idsToClean = @()
		foreach ($kvp in $global:RegisteredHotkeys.GetEnumerator())
		{
			if ($kvp.Value.Owners -and $kvp.Value.Owners.ContainsKey($OwnerKey))
			{
				$idsToClean += $kvp.Key
			}
		}

		foreach ($id in $idsToClean)
		{
			Write-Verbose "HOTKEYS: Found existing registration for owner '$OwnerKey' on ID $id. Removing..."
           
			$ownerAction = $global:RegisteredHotkeys[$id].Owners[$OwnerKey]
           
			$unregisteredSpecific = $false
			try
			{
				if ($ownerAction)
				{
					$unregisterMethod = [Custom.HotkeyManager].GetMethod('UnregisterAction')
					if ($unregisterMethod)
					{
						[Custom.HotkeyManager]::UnregisterAction($id, $ownerAction.ActionDelegate)
						$unregisteredSpecific = $true
					}
				}
			}
			catch { Write-Verbose "HOTKEYS: Specific unregister failed: $unregisteredSpecific $_" }

			if ($global:RegisteredHotkeys.ContainsKey($id))
			{
				$global:RegisteredHotkeys[$id].Owners.Remove($OwnerKey)
               
				if ($global:RegisteredHotkeys[$id].Owners.Count -eq 0)
				{
					$unregisteredSuccessfully = $false
					try
					{
						[Custom.HotkeyManager]::Unregister($id)
						$unregisteredSuccessfully = $true
					}
					catch { Write-Warning "HOTKEYS: Failed to unregister OS hotkey ID $id for '$OwnerKey'. Error: $_" }

					if ($unregisteredSuccessfully)
					{
						$ks = $global:RegisteredHotkeys[$id].KeyString
						if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) }
						$global:RegisteredHotkeys.Remove($id)
					}
				}
			}
		}
	}

	if ([string]::IsNullOrEmpty($KeyCombinationString))
	{
		return $null
	}

	if ($global:PausedRegisteredHotkeys)
	{
		$pausedKeys = @($global:PausedRegisteredHotkeys.Keys)
		foreach ($k in $pausedKeys)
		{
			if (-not $global:PausedRegisteredHotkeys.ContainsKey($k)) { continue }
			$meta = $global:PausedRegisteredHotkeys[$k]
			if ($meta -and $meta.Owners -and $meta.Owners.ContainsKey($OwnerKey))
			{
				Write-Verbose "HOTKEYS: Found and removed owner '$OwnerKey' from paused hotkey (fakeId $k)."
				$meta.Owners.Remove($OwnerKey)
				if ($meta.Owners.Count -eq 0)
				{
					$global:PausedRegisteredHotkeys.Remove($k)
					$ks = $meta.KeyString
					if ($ks -and $global:RegisteredHotkeyByString.ContainsKey($ks) -and $global:RegisteredHotkeyByString[$ks] -eq $k)
					{
						$global:RegisteredHotkeyByString.Remove($ks)
					}
				}
			}
		}
	}
   
	if ($OldHotkeyId -ne $null -and $OldHotkeyId -ne 0)
	{
		if ($global:RegisteredHotkeys.ContainsKey($OldHotkeyId))
		{
			try { [Custom.HotkeyManager]::Unregister($OldHotkeyId) } catch {}
			$global:RegisteredHotkeys.Remove($OldHotkeyId)
		}
	}

	$parsedResult = ParseKeyString $KeyCombinationString
	$modsFromParse = @()
	if ($parsedResult.Modifiers) { if ($parsedResult.Modifiers -is [System.Array]) { $modsFromParse = $parsedResult.Modifiers } else { $modsFromParse = @($parsedResult.Modifiers) } }
	[System.Collections.ArrayList]$parsedModifierKeys = @(); if ($modsFromParse.Count -gt 0) { $parsedModifierKeys.AddRange($modsFromParse) }
	[string]$parsedPrimaryKey = if ($parsedResult.Primary) { [string]$parsedResult.Primary } else { $null }

	if ([string]::IsNullOrEmpty($parsedPrimaryKey) -and $parsedModifierKeys.Count -gt 0) { throw 'A primary key is required.' }
	if ([string]::IsNullOrEmpty($parsedPrimaryKey) -and $parsedModifierKeys.Count -eq 0) { throw 'No key provided.' }

	[uint32]$mod = 0
	if ($parsedModifierKeys)
	{
		foreach ($mKey in $parsedModifierKeys)
		{
			switch ($mKey) { 'Alt' { $mod += 1 } 'Ctrl' { $mod += 2 } 'Shift' { $mod += 4 } 'Win' { $mod += 8 } }
		}
	}

	$vkNameFromNormalized = if ($parsedResult.Normalized) { ($parsedResult.Normalized -split '\+')[-1] } else { $null }
	$virtualKeyName = if (-not [string]::IsNullOrEmpty($vkNameFromNormalized)) { $vkNameFromNormalized } else { $parsedPrimaryKey }
	$virtualKey = (GetVirtualKeyMappings)[([string]$virtualKeyName).ToUpper()]
	if (-not $virtualKey) { throw "Invalid primary key: $virtualKeyName" }

	$canonicalOrder = @('Ctrl','Alt','Shift','Win')
	$canonicalModifiers = @(); foreach ($m in $canonicalOrder) { if ($parsedModifierKeys -contains $m) { $canonicalModifiers += $m } }
	$verboseKeyString = if ($canonicalModifiers.Count -gt 0) { "$(($canonicalModifiers -join '+'))+$parsedPrimaryKey" } else { $parsedPrimaryKey }
	$normalizedKeyString = $verboseKeyString.ToUpper()

	$actionType = [Custom.HotkeyManager+HotkeyActionType]::Normal
	if ($OwnerKey -and $OwnerKey -like 'global_toggle_*')
	{
		$actionType = [Custom.HotkeyManager+HotkeyActionType]::GlobalToggle
	}

	$hotkeyActionEntry = New-Object Custom.HotkeyManager+HotkeyActionEntry
	$hotkeyActionEntry.ActionDelegate = $actionDelegate
	$hotkeyActionEntry.Type = $actionType

	$ownerToggleOn = $true
	if ($OwnerKey)
	{
		$instanceId = $OwnerKey
		if ($OwnerKey -like 'ext_*')
		{
			$parts = $OwnerKey -split '_'
			if ($parts.Count -ge 2) { $instanceId = $parts[1] }
		}
		try
		{
			if ($global:DashboardConfig.Resources.InstanceHotkeysPaused -and $global:DashboardConfig.Resources.InstanceHotkeysPaused.Contains($instanceId))
			{
				$ownerToggleOn = -not [bool]$global:DashboardConfig.Resources.InstanceHotkeysPaused[$instanceId]
			}
			else
			{
				if ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
				{
					$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
					if ($form -and $form.Tag -and $form.Tag.BtnHotkeyToggle)
					{
						$ownerToggleOn = [bool]$form.Tag.BtnHotkeyToggle.Checked
					}
				}
			}
		}
		catch { }
	}

	if (-not $ownerToggleOn)
	{
		$existingPausedId = $null
		if ($global:PausedRegisteredHotkeys)
		{
			foreach ($k in $global:PausedRegisteredHotkeys.Keys)
			{
				try
				{
					if ($global:PausedRegisteredHotkeys[$k].KeyString -and $global:PausedRegisteredHotkeys[$k].KeyString -eq $normalizedKeyString)
					{
						$existingPausedId = $k
						break
					}
				}
				catch {}
			}
		}
       
		if ($existingPausedId)
		{
			$metaToUpdate = $global:PausedRegisteredHotkeys[$existingPausedId]
			if (-not $metaToUpdate.Owners) { $metaToUpdate.Owners = @{} }
           
			$updatedHotkeyActionEntry = New-Object Custom.HotkeyManager+HotkeyActionEntry
			$updatedHotkeyActionEntry.ActionDelegate = $actionDelegate
			$updatedHotkeyActionEntry.Type = $actionType

			$metaToUpdate.Owners[$OwnerKey] = $updatedHotkeyActionEntry
           
			Write-Verbose "HOTKEYS: Registered hotkey (paused, updated) for owner '$OwnerKey' on key '$normalizedKeyString'."
			return $existingPausedId
		}
		else
		{
			if (-not $global:NextPausedFakeId) { $global:NextPausedFakeId = -1 }
			$fakeId = $global:NextPausedFakeId
			$global:NextPausedFakeId = $global:NextPausedFakeId - 1
			$meta = @{ Modifier = $mod; Key = $virtualKey; KeyString = $normalizedKeyString; Owners = @{}; Action = $actionDelegate; ActionType = $actionType }
			$meta.Owners[$OwnerKey] = $hotkeyActionEntry
			if (-not $global:PausedRegisteredHotkeys) { $global:PausedRegisteredHotkeys = @{} }
			$global:PausedRegisteredHotkeys[$fakeId] = $meta
			$global:RegisteredHotkeyByString[$normalizedKeyString] = $fakeId
			$id = $fakeId
			Write-Verbose "HOTKEYS: Stored hotkey (paused) for owner $OwnerKey on key $normalizedKeyString."
			return $id
		}
	}

	if ($global:RegisteredHotkeyByString.ContainsKey($normalizedKeyString))
	{
		$existingId = $global:RegisteredHotkeyByString[$normalizedKeyString]
		if ($OwnerKey -and $global:RegisteredHotkeys.ContainsKey($existingId) -and $global:RegisteredHotkeys[$existingId].Owners.ContainsKey($OwnerKey))
		{
			$staleAction = $global:RegisteredHotkeys[$existingId].Owners[$OwnerKey]
			try { [Custom.HotkeyManager]::UnregisterAction($existingId, $staleAction.ActionDelegate) } catch {}
			$global:RegisteredHotkeys[$existingId].Owners.Remove($OwnerKey)
		}
		try
		{
			$id = [Custom.HotkeyManager]::Register($mod, $virtualKey, $hotkeyActionEntry)
		}
		catch
		{
			Write-Warning "HOTKEYS: Failed to register hotkey $($KeyCombinationString): $_"
			$id = $null
			throw 
		}
	}
	else
	{
		try
		{
			$id = [Custom.HotkeyManager]::Register($mod, $virtualKey, $hotkeyActionEntry)
		}
		catch
		{
			Write-Warning "HOTKEYS: Failed to register hotkey $($KeyCombinationString): $_"
			$id = $null
			throw 
		}
	}

	if ($global:RegisteredHotkeys.ContainsKey($id))
	{
		$existingKeyString = $global:RegisteredHotkeys[$id].KeyString
		if ($existingKeyString -and $existingKeyString -ne $normalizedKeyString)
		{
			$errorMsg = "HOTKEYS: CRITICAL hotkey conflict. ID $id."
			Write-Warning $errorMsg
			throw $errorMsg
		}
	}

	if (-not $global:RegisteredHotkeys.ContainsKey($id))
	{
		$global:RegisteredHotkeys[$id] = @{Modifier = $mod; Key = $virtualKey; KeyString = $normalizedKeyString; Owners = @{}; Action = $actionDelegate; ActionType = $actionType}
	}
	else
	{
		if (-not $global:RegisteredHotkeys[$id].KeyString) { $global:RegisteredHotkeys[$id].KeyString = $normalizedKeyString }
	}
   
	if ($OwnerKey)
	{
		if (-not $global:RegisteredHotkeys[$id].Owners) { $global:RegisteredHotkeys[$id].Owners = @{} }
		$global:RegisteredHotkeys[$id].Owners[$OwnerKey] = $hotkeyActionEntry
	}
	$global:RegisteredHotkeyByString[$normalizedKeyString] = $id

	try
	{
		if ($OwnerKey -and $global:PausedRegisteredHotkeys)
		{
			$pausedKeys = @($global:PausedRegisteredHotkeys.Keys)
			foreach ($k in $pausedKeys)
			{
				if (-not $global:PausedRegisteredHotkeys.ContainsKey($k)) { continue }
				$meta = $global:PausedRegisteredHotkeys[$k]
				if ($meta -and $meta.KeyString -and $meta.KeyString -eq $normalizedKeyString)
				{
					if ($meta.Owners -and $meta.Owners.ContainsKey($OwnerKey))
					{
						$meta.Owners.Remove($OwnerKey)
						if ($meta.Owners.Count -eq 0)
						{
							$global:PausedRegisteredHotkeys.Remove($k)
						}
						break 
					}
				}
			}
		}
	}
	catch {}
	RefreshHotkeysList
	return $id
}

function ResumeAllHotkeys
{
	if (-not $global:PausedRegisteredHotkeys -or $global:PausedRegisteredHotkeys.Count -eq 0) { return }
	if (-not $global:PausedIdMap) { $global:PausedIdMap = @{} }

	$pausedKeys = @($global:PausedRegisteredHotkeys.Keys)
	foreach ($oldId in $pausedKeys)
	{
		try
		{
			$meta = $global:PausedRegisteredHotkeys[$oldId]
			if (-not $meta) { continue }

			$mod = $meta.Modifier
			$vk = $meta.Key
			$ks = $meta.KeyString
			$owners = if ($meta.Owners) { @($meta.Owners.Keys) } else { @() }

			if (-not $global:RegisteredHotkeys) { $global:RegisteredHotkeys = @{} }

			foreach ($ok in $owners)
			{
				$shouldResume = $true
				try
				{
					$instId = $ok
					if ($ok -like 'ext_*') { $parts = $ok -split '_'; if ($parts.Count -ge 2) { $instId = $parts[1] } }
					if ($global:DashboardConfig.Resources.InstanceHotkeysPaused -and $global:DashboardConfig.Resources.InstanceHotkeysPaused.Contains($instId))
					{
						if ([bool]$global:DashboardConfig.Resources.InstanceHotkeysPaused[$instId]) { $shouldResume = $false }
					}
				}
				catch { }

				if (-not $shouldResume) { continue }

				$ownerAct = $meta.Owners[$ok]
				if (-not $ownerAct) { continue }
               
				$actionToRun = $ownerAct.ActionDelegate
				$delegate = [System.Action]({ InvokeHotkeyAction -Action $actionToRun }.GetNewClosure())

				try
				{
					$hotkeyActionEntry = New-Object Custom.HotkeyManager+HotkeyActionEntry
					$hotkeyActionEntry.ActionDelegate = $delegate
					$hotkeyActionEntry.Type = if ($ownerAct.Type) { $ownerAct.Type } else { [Custom.HotkeyManager+HotkeyActionType]::Normal }

					$registeredId = $null
					try
					{
						$registeredId = [Custom.HotkeyManager]::Register($mod, $vk, $hotkeyActionEntry)
					}
					catch
					{
						Write-Warning ("HOTKEYS: ResumeAllHotkeys failed owner '{0}': {1}" -f $ok, $_.Exception.Message)
					}

					if (-not $registeredId) { continue }
                   
					$global:PausedIdMap[$oldId] = $registeredId
					if (-not $global:RegisteredHotkeys.ContainsKey($registeredId))
					{
						$global:RegisteredHotkeys[$registeredId] = @{ Modifier = $mod; Key = $vk; KeyString = $ks; Owners = @{}; Action = $null; ActionType = $hotkeyActionEntry.Type }
					}
					$global:RegisteredHotkeys[$registeredId].Owners[$ok] = $hotkeyActionEntry
					if ($ks) { $global:RegisteredHotkeyByString[$ks] = $registeredId }
					$global:RegisteredHotkeys[$registeredId].ActionType = $hotkeyActionEntry.Type 

					$meta.Owners.Remove($ok)
				}
				catch {}
			}

			if (-not $meta.Owners -or $meta.Owners.Count -eq 0)
			{
				$global:PausedRegisteredHotkeys.Remove($oldId)
			}
			else
			{
				$global:PausedRegisteredHotkeys[$oldId] = $meta
			}
		}
		catch {}
	}
}

function ResumePausedKeys
{
	param([Parameter(Mandatory = $true)][array]$Keys)
	if (-not $Keys -or $Keys.Count -eq 0) { return }
	if (-not $global:PausedIdMap) { $global:PausedIdMap = @{} }
	if (-not $global:RegisteredHotkeys) { $global:RegisteredHotkeys = @{} }

	foreach ($oldId in $Keys)
	{
		if (-not $global:PausedRegisteredHotkeys.Contains($oldId)) { continue }
		try
		{
			$meta = $global:PausedRegisteredHotkeys[$oldId]
			if (-not $meta) { continue }

			$mod = $meta.Modifier; $vk = $meta.Key; $ks = $meta.KeyString
			$owners = if ($meta.Owners) { @($meta.Owners.Keys) } else { @() }

			foreach ($ok in $owners)
			{
				$shouldResume = $true
				try
				{
					$instId = $ok
					if ($ok -like 'ext_*') { $parts = $ok -split '_'; if ($parts.Count -ge 2) { $instId = $parts[1] } }
					if ($global:DashboardConfig.Resources.InstanceHotkeysPaused -and $global:DashboardConfig.Resources.InstanceHotkeysPaused.Contains($instId))
					{
						if ([bool]$global:DashboardConfig.Resources.InstanceHotkeysPaused[$instId]) { $shouldResume = $false }
					}
				}
				catch {}

				if (-not $shouldResume) { continue }

				$ownerAct = $meta.Owners[$ok]
				if (-not $ownerAct) { continue }
               
				$actionToRun = $ownerAct.ActionDelegate
				$delegate = [System.Action]({ InvokeHotkeyAction -Action $actionToRun }.GetNewClosure())
                               
				try
				{
					$hotkeyActionEntry = New-Object Custom.HotkeyManager+HotkeyActionEntry
					$hotkeyActionEntry.ActionDelegate = $delegate
					$hotkeyActionEntry.Type = if ($meta.ActionType) { $meta.ActionType } else { [Custom.HotkeyManager+HotkeyActionType]::Normal }
               
					$registeredId = [Custom.HotkeyManager]::Register($mod, $vk, $hotkeyActionEntry)
					$global:PausedIdMap[$oldId] = $registeredId
					if (-not $global:RegisteredHotkeys.ContainsKey($registeredId))
					{
						$global:RegisteredHotkeys[$registeredId] = @{ Modifier = $mod; Key = $vk; KeyString = $ks; Owners = @{}; Action = $null; ActionType = $hotkeyActionEntry.Type }
					}
					$global:RegisteredHotkeys[$registeredId].Owners[$ok] = $hotkeyActionEntry
					$global:RegisteredHotkeys[$registeredId].ActionType = $hotkeyActionEntry.Type
					if ($ks) { $global:RegisteredHotkeyByString[$ks] = $registeredId }
                   
					$meta.Owners.Remove($ok)
				}
				catch {}
			}

			if (-not $meta.Owners -or $meta.Owners.Count -eq 0)
			{
				$global:PausedRegisteredHotkeys.Remove($oldId)
			}
			else
			{
				$global:PausedRegisteredHotkeys[$oldId] = $meta
			}
		}
		catch {}
	}
}

function ResumeHotkeysForOwner
{
	param([Parameter(Mandatory = $true)][string]$OwnerKey)
	if (-not $OwnerKey) { return }
	if (-not $global:PausedRegisteredHotkeys -or $global:PausedRegisteredHotkeys.Count -eq 0) { return }
	if (-not $global:PausedIdMap) { $global:PausedIdMap = @{} }
   
	$idsToResume = @()
	foreach ($kvp in $global:PausedRegisteredHotkeys.GetEnumerator())
	{
		$oldId = $kvp.Key
		$meta = $kvp.Value
		if ($meta.Owners -and $meta.Owners.ContainsKey($OwnerKey))
		{
			$idsToResume += $oldId
		}
	}

	foreach ($oldId in $idsToResume)
	{
		try
		{
			$meta = $global:PausedRegisteredHotkeys[$oldId]
			$mod = $meta.Modifier; $vk = $meta.Key; $ks = $meta.KeyString; $owners = $meta.Owners

			foreach ($ok in $owners.Keys)
			{
				$ownerAct = $owners[$ok]
				if (-not $ownerAct) { continue }

				if ($ownerAct -is [Custom.HotkeyManager+HotkeyActionEntry])
				{
					$actionToRun = $ownerAct.ActionDelegate
				}
				else
				{
					$actionToRun = $ownerAct
				}
				$delegate = [System.Action]({ InvokeHotkeyAction -Action $actionToRun }.GetNewClosure())
               
				try
				{
					$hotkeyActionEntry = New-Object Custom.HotkeyManager+HotkeyActionEntry
					$hotkeyActionEntry.ActionDelegate = $delegate
					$hotkeyActionEntry.Type = if ($meta.ActionType) { $meta.ActionType } else { [Custom.HotkeyManager+HotkeyActionType]::Normal }

					$registeredId = [Custom.HotkeyManager]::Register($mod, $vk, $hotkeyActionEntry)
					$global:PausedIdMap[$oldId] = $registeredId
					if (-not $global:RegisteredHotkeys.ContainsKey($registeredId))
					{
						$global:RegisteredHotkeys[$registeredId] = @{Modifier = $mod; Key = $vk; KeyString = $ks; Owners = @{}; Action = $null; ActionType = $hotkeyActionEntry.Type }
					}
					$global:RegisteredHotkeys[$registeredId].Owners[$ok] = $hotkeyActionEntry
					$global:RegisteredHotkeys[$registeredId].ActionType = $hotkeyActionEntry.Type
					if ($ks) { $global:RegisteredHotkeyByString[$ks] = $registeredId }
				}
				catch {}
			}
			$global:PausedRegisteredHotkeys.Remove($oldId)
		}
		catch {}
	}
}

function RemoveAllHotkeys
{
	[Custom.HotkeyManager]::UnregisterAll()
	$global:RegisteredHotkeys.Clear()
	$global:RegisteredHotkeyByString.Clear()
	Write-Verbose 'HOTKEYS: All hotkeys unregistered.'
}

function TestHotkeyConflict
{
	param(
		[Parameter(Mandatory = $true)]
		[string]$KeyCombinationString,
		[Parameter(Mandatory = $true)]
		[Custom.HotkeyManager+HotkeyActionType]$NewHotkeyType,
		[Parameter(Mandatory = $false)]
		[string]$OwnerKeyToExclude
	)

	$normalizedKeyString = (NormalizeKeyString $KeyCombinationString).ToUpper()
	if (-not $normalizedKeyString) { return $false }

	foreach ($id in $global:RegisteredHotkeys.Keys)
	{
		$meta = $global:RegisteredHotkeys[$id]
		if ($meta -and $meta.KeyString -eq $normalizedKeyString)
		{
			$existingHotkeyType = if ($meta.ActionType) { $meta.ActionType } else { [Custom.HotkeyManager+HotkeyActionType]::Normal }

			$isConflict = $false
			if ($NewHotkeyType -ne $existingHotkeyType)
			{
				if (-not ($NewHotkeyType -eq [Custom.HotkeyManager+HotkeyActionType]::Normal -and $existingHotkeyType -eq [Custom.HotkeyManager+HotkeyActionType]::GlobalToggle))
				{
					$isConflict = $true
				}
			}
			elseif ($NewHotkeyType -eq [Custom.HotkeyManager+HotkeyActionType]::GlobalToggle)
			{
				$isConflict = $true
			}
           
			if ($isConflict)
			{
				if ($OwnerKeyToExclude)
				{
					if ($meta.Owners.ContainsKey($OwnerKeyToExclude))
					{
						continue 
					}
				}
				return $true
			}
		}
	}

	if ($global:PausedRegisteredHotkeys)
	{
		foreach ($id in $global:PausedRegisteredHotkeys.Keys)
		{
			$meta = $global:PausedRegisteredHotkeys[$id]
			if ($meta -and $meta.KeyString -eq $normalizedKeyString)
			{
				$existingHotkeyType = if ($meta.ActionType) { $meta.ActionType } else { [Custom.HotkeyManager+HotkeyActionType]::Normal }

				$isConflict = $false
				if (!($NewHotkeyType -eq [Custom.HotkeyManager+HotkeyActionType]::Normal -and $existingHotkeyType -eq [Custom.HotkeyManager+HotkeyActionType]::GlobalToggle))
				{
					$isConflict = $true
				}

				if ($isConflict)
				{
					if ($OwnerKeyToExclude)
					{
						if ($meta.Owners.ContainsKey($OwnerKeyToExclude))
						{
							continue
						}
					}
					return $true
				}
			}
		}
	}

	return $false
}

function InvokeHotkeyAction
{
	param([Parameter(Mandatory = $true)]$Action)
	try
	{
		if ($Action -is [ScriptBlock])
		{ 
			& $Action 
		}
		elseif ($Action -is [System.Delegate] -or $Action -is [System.Action])
		{ 
			$Action.Invoke() 
		}
		elseif ($Action -is [string])
		{
			$cmd = Get-Command $Action -ErrorAction SilentlyContinue
			if ($cmd) { & $cmd.Name } else { Invoke-Expression $Action }
		}
	}
	catch
	{
		Write-Verbose ('HOTKEYS-DEBUG: InvokeHotkeyAction: Exception invoking action: {0}' -f $_.Exception.Message)
	}
}

function PauseAllHotkeys
{
	[Custom.HotkeyManager]::AreHotkeysGloballyPaused = $true
	if (-not $global:RegisteredHotkeys -or $global:RegisteredHotkeys.Count -eq 0) { return }
   
	$global:PausedIdMap = @{}
	if (-not $global:PausedRegisteredHotkeys) { $global:PausedRegisteredHotkeys = @{} }

	$idsToProcess = @($global:RegisteredHotkeys.Keys)

	foreach ($id in $idsToProcess)
	{
		try
		{
			$meta = $global:RegisteredHotkeys[$id]
			if (-not $meta -or -not $meta.Owners) { continue }

			$ownersToRemove = @()
			$ownerEntriesToPause = @{}
			$shouldFullyUnregisterOSHotkey = $true
           
			$ownerKeys = @($meta.Owners.Keys) 
			foreach ($ownerKey in $ownerKeys)
			{
				$ownerEntry = $meta.Owners[$ownerKey]
				if ($null -eq $ownerEntry -or $null -eq $ownerEntry.ActionDelegate) { continue }

				try
				{
					[Custom.HotkeyManager]::UnregisterAction($id, $ownerEntry.ActionDelegate)
					$ownersToRemove += $ownerKey 
					$ownerEntriesToPause[$ownerKey] = $ownerEntry
				}
				catch
				{
					$shouldFullyUnregisterOSHotkey = $false
				}
			}

			foreach ($ownerKey in $ownersToRemove)
			{
				$meta.Owners.Remove($ownerKey)
			}

			$pausedMetaForId = @{
				Modifier   = $meta.Modifier
				Key        = $meta.Key
				KeyString  = $meta.KeyString
				Owners     = @{}
				Action     = $meta.Action
				ActionType = $meta.ActionType
			}
			foreach ($ownerKey in $ownerEntriesToPause.Keys)
			{
				$pausedMetaForId.Owners[$ownerKey] = $ownerEntriesToPause[$ownerKey]
			}

			if ($pausedMetaForId.Owners.Count -gt 0)
			{
				$global:PausedRegisteredHotkeys[$id] = $pausedMetaForId
				if ($meta.KeyString) { $global:RegisteredHotkeyByString.Remove($meta.KeyString) }
				$global:RegisteredHotkeyByString[$meta.KeyString] = $id
			}

			if ($shouldFullyUnregisterOSHotkey -and $meta.Owners.Count -eq 0)
			{
				try { [Custom.HotkeyManager]::Unregister($id) } catch {}
				$global:RegisteredHotkeys.Remove($id)
			}
			else
			{
				$global:RegisteredHotkeys[$id] = $meta
			}

		}
		catch {}
	}
}

function PauseHotkeysForOwner
{
	param([Parameter(Mandatory = $true)][string]$OwnerKey)
	if (-not $OwnerKey) { return }
	if (-not $global:RegisteredHotkeys -or $global:RegisteredHotkeys.Count -eq 0) { return }
	if (-not $global:PausedRegisteredHotkeys) { $global:PausedRegisteredHotkeys = @{} }
	if (-not $global:PausedIdMap) { $global:PausedIdMap = @{} }

	$idsToProcess = @()
	foreach ($kvp in $global:RegisteredHotkeys.GetEnumerator())
	{
		$id = $kvp.Key
		$meta = $kvp.Value
		if ($meta.Owners -and $meta.Owners.ContainsKey($OwnerKey))
		{
			$idsToProcess += $id
		}
	}

	foreach ($id in $idsToProcess)
	{
		try
		{
			$meta = $global:RegisteredHotkeys[$id]
			$ks = $null
			try { $ks = $meta.KeyString } catch {}

			$ownerEntry = $null
			try { $ownerEntry = $meta.Owners[$OwnerKey] } catch {}

			if ($null -eq $ownerEntry -or $null -eq $ownerEntry.ActionDelegate)
			{
				continue
			}

			if (-not $global:PausedRegisteredHotkeys) { $global:PausedRegisteredHotkeys = @{} }
			if (-not $global:NextPausedFakeId) { $global:NextPausedFakeId = -1 }

			$unregisterActionMethod = $null
			try { $unregisterActionMethod = [Custom.HotkeyManager].GetMethod('UnregisterAction') } catch {}
			if ($unregisterActionMethod)
			{
				try
				{
					[Custom.HotkeyManager]::UnregisterAction($id, $ownerEntry.ActionDelegate)

					if ($global:RegisteredHotkeys.ContainsKey($id) -and $global:RegisteredHotkeys[$id].Owners.ContainsKey($OwnerKey))
					{
						$global:RegisteredHotkeys[$id].Owners.Remove($OwnerKey)
					}

					$fakeId = $global:NextPausedFakeId; $global:NextPausedFakeId = $global:NextPausedFakeId - 1
					$pausedMeta = @{ Modifier = $meta.Modifier; Key = $meta.Key; KeyString = $ks; Owners = @{}; Action = $ownerEntry.ActionDelegate; ActionType = $ownerEntry.Type } 
					$pausedMeta.Owners[$OwnerKey] = $ownerEntry
					$global:PausedRegisteredHotkeys[$fakeId] = $pausedMeta
					if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) }
					$global:RegisteredHotkeyByString[$ks] = $fakeId
					continue
				}
				catch {}
			}

			$global:PausedRegisteredHotkeys[$id] = $meta
			try { [Custom.HotkeyManager]::Unregister($id) } catch { }
			try { if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) } } catch {}
			try { $global:RegisteredHotkeys.Remove($id) } catch {}
		}
		catch {}
	}
}

function UnregisterHotkeyInstance
{
	param(
		[Parameter(Mandatory = $true)][int]$Id,
		[Parameter(Mandatory = $false)][string]$OwnerKey
	)
	if (-not $Id) { return }
	try
	{
		$didUnregister = $false
		$translateTarget = $null
		if ($global:PausedIdMap -and $global:PausedIdMap.ContainsKey($Id))
		{
			$translateTarget = $global:PausedIdMap[$Id]
		}
		if ($translateTarget) { $IdToUse = $translateTarget } else { $IdToUse = $Id }
       
		if (-not $global:RegisteredHotkeys.ContainsKey($IdToUse) -and $global:PausedRegisteredHotkeys -and $global:PausedRegisteredHotkeys.ContainsKey($IdToUse))
		{
			$pausedMeta = $global:PausedRegisteredHotkeys[$IdToUse]
			if ($OwnerKey -and $pausedMeta.Owners -and $pausedMeta.Owners.ContainsKey($OwnerKey))
			{
				$pausedMeta.Owners.Remove($OwnerKey)
				if ($pausedMeta.Owners.Count -eq 0) { $global:PausedRegisteredHotkeys.Remove($IdToUse); }
				$didUnregister = $true
			}
			else
			{
				$global:PausedRegisteredHotkeys.Remove($IdToUse)
				$didUnregister = $true
			}
		}
		if ($OwnerKey -and $global:RegisteredHotkeys.ContainsKey($Id) -and $global:RegisteredHotkeys[$Id].Owners.ContainsKey($OwnerKey))
		{
			$ownerEntry = $global:RegisteredHotkeys[$Id].Owners[$OwnerKey]
			$unregisterMethod = $null
			try { $unregisterMethod = [Custom.HotkeyManager].GetMethod('UnregisterAction') } catch {}
			if ($unregisterMethod)
			{
				try
				{
					[Custom.HotkeyManager]::UnregisterAction($Id, $ownerEntry.ActionDelegate) 
					$didUnregister = $true
				}
				catch
				{
					Write-Warning 'HOTKEYS: UnregisterAction failed. Fallback.'
				}
			} 
			if ($global:RegisteredHotkeys.ContainsKey($Id) -and $global:RegisteredHotkeys[$Id].Owners.ContainsKey($OwnerKey))
			{
				$global:RegisteredHotkeys[$Id].Owners.Remove($OwnerKey)
			}
			if ($global:RegisteredHotkeys.ContainsKey($Id) -and $global:RegisteredHotkeys[$Id].Owners.Count -eq 0)
			{
				$ks = $global:RegisteredHotkeys[$Id].KeyString
				if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) }
				$global:RegisteredHotkeys.Remove($Id)
			}
		}
		if (-not $didUnregister)
		{
			if ($global:RegisteredHotkeys.ContainsKey($Id))
			{
				$unregisteredSuccessfully = $false
				try
				{
					[Custom.HotkeyManager]::Unregister($Id)
					$unregisteredSuccessfully = $true
				}
				catch {}
               
				if ($unregisteredSuccessfully)
				{
					if ($global:RegisteredHotkeys.ContainsKey($Id))
					{
						$ks = $global:RegisteredHotkeys[$Id].KeyString
						if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) }
						$global:RegisteredHotkeys.Remove($Id)
					}
				}
			}
		}
	}
	catch
	{
		Write-Warning "HOTKEYS: Failed to unregister hotkey ID $Id. Error: $_"
	}
}

function GetVirtualKeyMappings
{
	return @{
		'LEFT_MOUSE' = 0x01
		'RIGHT_MOUSE' = 0x02
		'CANCEL' = 0x03 
		'MIDDLE_MOUSE' = 0x04 
		'MOUSE_X1' = 0x05 
		'MOUSE_X2' = 0x06 

		'F1' = 0x70; 'F2' = 0x71; 'F3' = 0x72; 'F4' = 0x73; 'F5' = 0x74; 'F6' = 0x75
		'F7' = 0x76; 'F8' = 0x77; 'F9' = 0x78; 'F10' = 0x79; 'F11' = 0x7A; 'F12' = 0x7B
		'F13' = 0x7C; 'F14' = 0x7D; 'F15' = 0x7E; 'F16' = 0x7F; 'F17' = 0x80; 'F18' = 0x81
		'F19' = 0x82; 'F20' = 0x83; 'F21' = 0x84; 'F22' = 0x85; 'F23' = 0x86; 'F24' = 0x87

		'0' = 0x30; '1' = 0x31; '2' = 0x32; '3' = 0x33; '4' = 0x34
		'5' = 0x35; '6' = 0x36; '7' = 0x37; '8' = 0x38; '9' = 0x39

		'A' = 0x41; 'B' = 0x42; 'C' = 0x43; 'D' = 0x44; 'E' = 0x45; 'F' = 0x46; 'G' = 0x47; 'H' = 0x48
		'I' = 0x49; 'J' = 0x4A; 'K' = 0x4B; 'L' = 0x4C; 'M' = 0x4D; 'N' = 0x4E; 'O' = 0x4F; 'P' = 0x50
		'Q' = 0x51; 'R' = 0x52; 'S' = 0x53; 'T' = 0x54; 'U' = 0x55; 'V' = 0x56; 'W' = 0x57; 'X' = 0x58
		'Y' = 0x59; 'Z' = 0x5A

       
		'SPACE' = 0x20; 'ENTER' = 0x0D; 'TAB' = 0x09; 'ESCAPE' = 0x1B 
		'SHIFT' = 0x10; 'CONTROL' = 0x11; 'ALT' = 0x12
		'UP_ARROW' = 0x26; 'DOWN_ARROW' = 0x28; 'LEFT_ARROW' = 0x25; 'RIGHT_ARROW' = 0x27
		'HOME' = 0x24; 'END' = 0x23; 'PAGE_UP' = 0x21; 'PAGE_DOWN' = 0x22
		'INSERT' = 0x2D; 'DELETE' = 0x2E; 'BACKSPACE' = 0x08

		'NAV_VIEW' = 0x88; 'NAV_MENU' = 0x89
		'NAV_UP' = 0x8A; 'NAV_DOWN' = 0x8B
		'NAV_LEFT' = 0x8C; 'NAV_RIGHT' = 0x8D
		'NAV_ACCEPT' = 0x8E; 'NAV_CANCEL' = 0x8F

		'CAPS_LOCK' = 0x14; 'NUM_LOCK' = 0x90; 'SCROLL_LOCK' = 0x91
		'PRINT_SCREEN' = 0x2C; 'PAUSE_BREAK' = 0x13
		'LEFT_WINDOWS' = 0x5B; 'RIGHT_WINDOWS' = 0x5C; 'APPLICATION' = 0x5D
		'LEFT_SHIFT' = 0xA0; 'RIGHT_SHIFT' = 0xA1
		'LEFT_CONTROL' = 0xA2; 'RIGHT_CONTROL' = 0xA3
		'LEFT_ALT' = 0xA4; 'RIGHT_ALT' = 0xA5; 'SLEEP' = 0x5F

		'NUMPAD_0' = 0x60; 'NUMPAD_1' = 0x61; 'NUMPAD_2' = 0x62; 'NUMPAD_3' = 0x63; 'NUMPAD_4' = 0x64
		'NUMPAD_5' = 0x65; 'NUMPAD_6' = 0x66; 'NUMPAD_7' = 0x67; 'NUMPAD_8' = 0x68; 'NUMPAD_9' = 0x69
		'NUMPAD_MULTIPLY' = 0x6A; 'NUMPAD_ADD' = 0x6B; 'NUMPAD_SEPARATOR' = 0x6C
		'NUMPAD_SUBTRACT' = 0x6D; 'NUMPAD_DECIMAL' = 0x6E; 'NUMPAD_DIVIDE' = 0x6F

		'SEMICOLON' = 0xBA 
		'EQUALS' = 0xBB 
		'COMMA' = 0xBC 
		'MINUS' = 0xBD 
		'PERIOD' = 0xBE 
		'FORWARD_SLASH' = 0xBF 
		'BACKTICK' = 0xC0 
		'LEFT_BRACKET' = 0xDB 
		'BACKSLASH' = 0xDC 
		'RIGHT_BRACKET' = 0xDD 
		'APOSTROPHE' = 0xDE 

		'<' = 0xE2 
		'OEM_8' = 0xDF 
		'AX' = 0xE1 
		'PACKET' = 0xE7 

		'BROWSER_BACK' = 0xA6; 'BROWSER_FORWARD' = 0xA7; 'BROWSER_REFRESH' = 0xA8; 'BROWSER_STOP' = 0xA9
		'BROWSER_SEARCH' = 0xAA; 'BROWSER_FAVORITES' = 0xAB; 'BROWSER_HOME' = 0xAC; 'VOLUME_MUTE' = 0xAD
		'VOLUME_DOWN' = 0xAE; 'VOLUME_UP' = 0xAF; 'MEDIA_NEXT_TRACK' = 0xB0; 'MEDIA_PREVIOUS_TRACK' = 0xB1
		'MEDIA_STOP' = 0xB2; 'MEDIA_PLAY_PAUSE' = 0xB3; 'LAUNCH_MAIL' = 0xB4; 'LAUNCH_MEDIA_PLAYER' = 0xB5
		'LAUNCH_MY_COMPUTER' = 0xB6; 'LAUNCH_CALCULATOR' = 0xB7

		'GAMEPAD_A' = 0xC3; 'GAMEPAD_B' = 0xC4
		'GAMEPAD_X' = 0xC5; 'GAMEPAD_Y' = 0xC6
		'GAMEPAD_RIGHT_BUMPER' = 0xC7; 'GAMEPAD_LEFT_BUMPER' = 0xC8
		'GAMEPAD_LEFT_TRIGGER' = 0xC9; 'GAMEPAD_RIGHT_TRIGGER' = 0xCA
		'GAMEPAD_DPAD_UP' = 0xCB; 'GAMEPAD_DPAD_DOWN' = 0xCC
		'GAMEPAD_DPAD_LEFT' = 0xCD; 'GAMEPAD_DPAD_RIGHT' = 0xCE
		'GAMEPAD_MENU' = 0xCF; 'GAMEPAD_VIEW' = 0xD0
		'GAMEPAD_LEFT_THUMB_BUTTON' = 0xD1; 'GAMEPAD_RIGHT_THUMB_BUTTON' = 0xD2
		'GAMEPAD_LEFT_THUMB_UP' = 0xD3; 'GAMEPAD_LEFT_THUMB_DOWN' = 0xD4
		'GAMEPAD_LEFT_THUMB_RIGHT' = 0xD5; 'GAMEPAD_LEFT_THUMB_LEFT' = 0xD6
		'GAMEPAD_RIGHT_THUMB_UP' = 0xD7; 'GAMEPAD_RIGHT_THUMB_DOWN' = 0xD8
		'GAMEPAD_RIGHT_THUMB_RIGHT' = 0xD9; 'GAMEPAD_RIGHT_THUMB_LEFT' = 0xDA

		'IME_KANA_HANGUL' = 0x15; 'IME_JUNJA' = 0x17; 'IME_FINAL' = 0x18; 'IME_HANJA_KANJI' = 0x19
		'IME_CONVERT' = 0x1C; 'IME_NONCONVERT' = 0x1D; 'IME_ACCEPT' = 0x1E; 'IME_MODE_CHANGE' = 0x1F; 'IME_PROCESS' = 0xE5

		'OEM_FJ_JISHO' = 0x92; 'OEM_FJ_MASSHOU' = 0x93; 'OEM_FJ_TOUROKU' = 0x94; 'OEM_FJ_LOYA' = 0x95; 'OEM_FJ_ROYA' = 0x96
       
		'SELECT' = 0x29; 'PRINT' = 0x2A; 'EXECUTE' = 0x2B; 'HELP' = 0x2F; 'CLEAR' = 0x0C
		'ATTN' = 0xF6; 'CRSEL' = 0xF7; 'EXSEL' = 0xF8; 'ERASE_EOF' = 0xF9; 'PLAY' = 0xFA; 'ZOOM' = 0xFB
		'PA1' = 0xFD; 'OEM_CLEAR' = 0xFE
	}
}

function NormalizeKeyString
{
	param(
		[Parameter(Mandatory = $true)][string]$KeyCombinationString
	)

	if ([string]::IsNullOrEmpty($KeyCombinationString)) { return $null }
	$parts = @($KeyCombinationString -split '[\+\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
   
	$parsedModifierKeys = @()
	$parsedPrimaryKey = $null
	foreach ($part in $parts)
	{
		switch ($part.ToUpper())
		{
			'ALT' { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'MENU' { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'LALT' { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'RALT' { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'LEFT_ALT' { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'RIGHT_ALT' { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'CTRL' { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'CONTROL' { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'LCTRL' { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'RCTRL' { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'LEFT_CONTROL' { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'RIGHT_CONTROL' { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'SHIFT' { if ($parsedModifierKeys -notcontains 'Shift') { $parsedModifierKeys += 'Shift' } }
			'LSHIFT' { if ($parsedModifierKeys -notcontains 'Shift') { $parsedModifierKeys += 'Shift' } }
			'RSHIFT' { if ($parsedModifierKeys -notcontains 'Shift') { $parsedModifierKeys += 'Shift' } }
			'WIN' { if ($parsedModifierKeys -notcontains 'Win') { $parsedModifierKeys += 'Win' } }
			'LWIN' { if ($parsedModifierKeys -notcontains 'Win') { $parsedModifierKeys += 'Win' } }
			'RWIN' { if ($parsedModifierKeys -notcontains 'Win') { $parsedModifierKeys += 'Win' } }
			'WINDOWS' { if ($parsedModifierKeys -notcontains 'Win') { $parsedModifierKeys += 'Win' } }
			default
			{ 
				if ([string]::IsNullOrEmpty($parsedPrimaryKey))
				{ 
					$parsedPrimaryKey = $part 
				}
				else
				{
					$parsedPrimaryKey = "$parsedPrimaryKey $part"
				}
			}
		}
	}

	$canonicalOrder = @('Ctrl','Alt','Shift','Win')
	$canonicalModifiers = @()
	foreach ($m in $canonicalOrder) { if ($parsedModifierKeys -contains $m) { $canonicalModifiers += $m } }

	if ($canonicalModifiers.Count -gt 0)
	{
		if ([string]::IsNullOrEmpty($parsedPrimaryKey)) { return ([string]($canonicalModifiers -join '+')).ToUpper() }
		return ([string]("$(($canonicalModifiers -join '+'))+$parsedPrimaryKey")).ToUpper()
	}
	else
	{
		return ([string]$parsedPrimaryKey).ToUpper()
	}
}

function ParseKeyString
{
	param(
		[Parameter(Mandatory = $true)][string]$KeyCombinationString
	)

	$normalized = NormalizeKeyString $KeyCombinationString
	if (-not $normalized) { return @{Modifiers = @(); Primary = $null; Normalized = $null} }

	$parts = @($normalized -split '\s*\+\s*' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
   
	if ($parts.Count -eq 0) { return @{Modifiers = @(); Primary = $null; Normalized = $normalized} }

	if ($parts.Count -eq 1)
	{
		return @{Modifiers = @(); Primary = ([string]$parts[0]).ToUpper(); Normalized = $normalized}
	}

	$primary = ([string]$parts[-1]).ToUpper()
	$mods = $parts[0..($parts.Count - 2)] | ForEach-Object {
		switch ($_.ToUpper())
		{
			'CTRL' { 'Ctrl' }
			'CONTROL' { 'Ctrl' }
			'LCTRL' { 'Ctrl' }
			'RCTRL' { 'Ctrl' }
			'LEFT_CONTROL' { 'Ctrl' }
			'RIGHT_CONTROL' { 'Ctrl' }
			'ALT' { 'Alt' }
			'MENU' { 'Alt' }
			'LALT' { 'Alt' }
			'RALT' { 'Alt' }
			'LEFT_ALT' { 'Alt' }
			'RIGHT_ALT' { 'Alt' }
			'SHIFT' { 'Shift' }
			'LSHIFT' { 'Shift' }
			'RSHIFT' { 'Shift' }
			'WIN' { 'Win' }
			'LWIN' { 'Win' }
			'RWIN' { 'Win' }
			'WINDOWS' { 'Win' }
			default { $_ }
		}
	}
	return @{Modifiers = $mods; Primary = $primary; Normalized = $normalized}
}

function GetKeyCombinationString
{
	param([string[]]$modifiers, [string]$primaryKey)
	$parts = @()
	if ($modifiers) { $parts += $modifiers }
	if ($primaryKey) { $parts += $primaryKey }
	if ($parts.Count -eq 0) { return 'none' }
	return ($parts -join ' ')
}

function IsModifierKeyCode
{
	param([System.Windows.Forms.Keys]$keyCode)
	return (
		$keyCode -eq [System.Windows.Forms.Keys]::ControlKey -or  
		$keyCode -eq [System.Windows.Forms.Keys]::LControlKey -or
		$keyCode -eq [System.Windows.Forms.Keys]::RControlKey -or
		$keyCode -eq [System.Windows.Forms.Keys]::ShiftKey -or    
		$keyCode -eq [System.Windows.Forms.Keys]::LShiftKey -or
		$keyCode -eq [System.Windows.Forms.Keys]::RShiftKey -or
		$keyCode -eq [System.Windows.Forms.Keys]::Menu -or        
		$keyCode -eq [System.Windows.Forms.Keys]::LMenu -or       
		$keyCode -eq [System.Windows.Forms.Keys]::RMenu -or       
		$keyCode -eq [System.Windows.Forms.Keys]::LWin -or
		$keyCode -eq [System.Windows.Forms.Keys]::RWin
	)
}

function Show-KeyCaptureDialog
{
    param(
        [string]$currentKey = '',
        [System.Windows.Forms.IWin32Window]$Owner = $null,
        [System.Windows.Forms.Form]$OwnerForm
    )
    try
    {
        PauseAllHotkeys
        $script:KeyCapture_PausedSnapshot = @()
        if ($global:PausedRegisteredHotkeys) { $script:KeyCapture_PausedSnapshot = @($global:PausedRegisteredHotkeys.Keys) }
    }
    catch
    {
        Write-Verbose ('HOTKEYS: Pause-AllHotkeys failed: {0}' -f $_.Exception.Message)
    }

    $script:capturedModifierKeys = @()
    $script:capturedPrimaryKey = $null

   
    if (-not [string]::IsNullOrEmpty($currentKey) -and $currentKey -ne 'Hotkey' -and $currentKey -ne 'none')
    {
        $parts = $currentKey.Split(' ')
        foreach ($part in $parts)
        {
            switch ($part.ToUpper())
            {
                'ALT' { if ($script:capturedModifierKeys -notcontains 'Alt') { $script:capturedModifierKeys += 'Alt' } }
                'CTRL' { if ($script:capturedModifierKeys -notcontains 'Ctrl') { $script:capturedModifierKeys += 'Ctrl' } }
                'SHIFT' { if ($script:capturedModifierKeys -notcontains 'Shift') { $script:capturedModifierKeys += 'Shift' } }
                'WIN' { if ($script:capturedModifierKeys -notcontains 'Win') { $script:capturedModifierKeys += 'Win' } }
                default { if ([string]::IsNullOrEmpty($script:capturedPrimaryKey)) { $script:capturedPrimaryKey = $part }; break } 
            }
        }
    }

    $captureForm = New-Object System.Windows.Forms.Form
    $captureForm.Text = 'Capture Input'
    $captureForm.Size = New-Object System.Drawing.Size(340, 180)
    $captureForm.StartPosition = 'CenterScreen'
    $captureForm.FormBorderStyle = 'FixedDialog'
    $captureForm.MaximizeBox = $false
    $captureForm.MinimizeBox = $false
    $captureForm.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $captureForm.ForeColor = [System.Drawing.Color]::White
    $captureForm.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = 'Press any key combination or click a mouse button inside this window.'
    $label.Size = New-Object System.Drawing.Size(300, 60)
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.TextAlign = 'MiddleCenter'
    $label.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $label.ForeColor = [System.Drawing.Color]::White
    $label.Enabled = $false

    $label.Add_Paint({
            param($s, $e) 

       
            $graphics = $e.Graphics
            $text = $this.Text

       
            $font = $this.Font
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
            $rect = $this.ClientRectangle

       
            $format = New-Object System.Drawing.StringFormat
            $format.Alignment = [System.Drawing.StringAlignment]::Center
            $format.LineAlignment = [System.Drawing.StringAlignment]::Center

            $rectF = New-Object System.Drawing.RectangleF($rect.X, $rect.Y, $rect.Width, $rect.Height)

       
            $graphics.DrawString($text, $font, $brush, $rectF, $format)

       
            $brush.Dispose()

        })
    $captureForm.Controls.Add($label)

   
    $resultPanel = New-Object System.Windows.Forms.Panel
    $resultPanel.Location = New-Object System.Drawing.Point(10, 80)
    $resultPanel.Size = New-Object System.Drawing.Size(300, 40)
    $resultPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $captureForm.Controls.Add($resultPanel)

    $resultLabel = New-Object System.Windows.Forms.Label
    $resultLabel.Text = (GetKeyCombinationString $script:capturedModifierKeys $script:capturedPrimaryKey)
    $resultLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $resultLabel.TextAlign = 'MiddleCenter'
    $resultLabel.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $resultLabel.ForeColor = [System.Drawing.Color]::White
    $resultLabel.Enabled = $false

    $resultLabel.Add_Paint({

            param($s, $e) 

       
            $graphics = $e.Graphics
            $text = $this.Text

       
            $font = $this.Font
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
            $rect = $this.ClientRectangle

       
            $format = New-Object System.Drawing.StringFormat
            $format.Alignment = [System.Drawing.StringAlignment]::Center
            $format.LineAlignment = [System.Drawing.StringAlignment]::Center

            $rectF = New-Object System.Drawing.RectangleF($rect.X, $rect.Y, $rect.Width, $rect.Height)

       
            $graphics.DrawString($text, $font, $brush, $rectF, $format)

       
            $brush.Dispose()
        })
    $resultPanel.Controls.Add($resultLabel)

   
    $captureForm.Add_MouseDown({
            param($s, $e)
       
            $formBounds = $s.Bounds
            $cursorPos = [System.Windows.Forms.Cursor]::Position
            if (-not $formBounds.Contains($cursorPos))
            {
                return
            }

       
            [System.Collections.ArrayList]$currentModifiers = @()
            if ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control) { $currentModifiers.Add('Ctrl') }
            if ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Alt) { $currentModifiers.Add('Alt') }
            if ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Shift) { $currentModifiers.Add('Shift') }
       

       
            $btnName = $null
            switch ($e.Button)
            {
                'Left' { $btnName = 'LEFT_MOUSE' }
                'Right' { $btnName = 'RIGHT_MOUSE' }
                'Middle' { $btnName = 'MIDDLE_MOUSE' }
                'XButton1' { $btnName = 'MOUSE_X1' }
                'XButton2' { $btnName = 'MOUSE_X2' }
            }

            if ($btnName)
            {
                $script:capturedPrimaryKey = $btnName
                $script:capturedModifierKeys = $currentModifiers
           
                $resultLabel.Text = "Captured: $(GetKeyCombinationString $script:capturedModifierKeys $script:capturedPrimaryKey)"
                $resultLabel.ForeColor = [System.Drawing.Color]::Green
                $captureForm.DialogResult = 'OK'
                $captureForm.Close()
            }
        })
   
   
    $captureForm.Add_KeyDown({
            param($form, $e)

            [System.Collections.ArrayList]$currentModifiersTemp = @()
            if ($e.Control) { $currentModifiersTemp.Add('Ctrl') }
            if ($e.Alt) { $currentModifiersTemp.Add('Alt') }
            if ($e.Shift) { $currentModifiersTemp.Add('Shift') }
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::LWin -or $e.KeyCode -eq [System.Windows.Forms.Keys]::RWin)
            {
                if ($currentModifiersTemp -notcontains 'Win') { $currentModifiersTemp.Add('Win') }
            }

            $keyMappings = GetVirtualKeyMappings
            [string]$actualPressedKeyName = $null
            foreach ($kvp in $keyMappings.GetEnumerator())
            {
                if ($kvp.Value -eq $e.KeyValue)
                {
                    $actualPressedKeyName = $kvp.Key
                    break
                }
            }
            if (-not $actualPressedKeyName)
            {
                $actualPressedKeyName = $e.KeyCode.ToString()
            }
            $pressedKeyName = $actualPressedKeyName

            $isPhysicalModifier = IsModifierKeyCode $e.KeyCode
            $knownModifierNames = @(
                'CONTROLKEY', 'MENU', 'SHIFTKEY', 'LWIN', 'RWIN', 
                'CONTROL', 'ALT', 'SHIFT', 'WIN',                
                'LEFT_CONTROL', 'RIGHT_CONTROL', 'LEFT_ALT', 'RIGHT_ALT', 'LEFT_SHIFT', 'RIGHT_SHIFT', 
                'LBUTTON', 'RBUTTON', 'MBUTTON', 'XBUTTON1', 'XBUTTON2' 
            )
            $isNamedModifier = $knownModifierNames -contains $pressedKeyName.ToUpper()


            if (-not $isPhysicalModifier -and -not $isNamedModifier)
            {
                $script:capturedPrimaryKey = $pressedKeyName
                $script:capturedModifierKeys = $currentModifiersTemp
           
                $resultLabel.Text = "Captured: $(GetKeyCombinationString $script:capturedModifierKeys $script:capturedPrimaryKey)"
                $resultLabel.ForeColor = [System.Drawing.Color]::Green
                $captureForm.DialogResult = 'OK'
                $captureForm.Close()
            }
            else
            {
                $resultLabel.Text = (GetKeyCombinationString $currentModifiersTemp $null) 
                $resultLabel.ForeColor = [System.Drawing.Color]::White
            }
        })
   
    $captureForm.Add_KeyUp({
            param($form, $e)
            if ([string]::IsNullOrEmpty($script:capturedPrimaryKey))
            {
                [System.Collections.ArrayList]$currentModifiersOnUp = @()
                if ($e.Control) { $currentModifiersOnUp.Add('Ctrl') }
                if ($e.Alt) { $currentModifiersOnUp.Add('Alt') }
                if ($e.Shift) { $currentModifiersOnUp.Add('Shift') }
           
                $resultLabel.Text = (GetKeyCombinationString $currentModifiersOnUp $null)
                $resultLabel.ForeColor = [System.Drawing.Color]::Yellow
            }
        })

    $captureForm.KeyPreview = $true
    $captureForm.TopMost = $true
    
    # --- CHANGE START ---
    # Configure ownership but rely on Show-FormAsDialog for the display loop
    if ($Owner)
    {
        $captureForm.Owner = $Owner
        $captureForm.StartPosition = 'CenterScreen'
    }
    elseif ($OwnerForm)
    {
        $captureForm.Owner = $OwnerForm
        $captureForm.StartPosition = 'CenterScreen'
    }

    $result = Show-FormAsDialog -Form $captureForm

    try
    {
        if ($result -eq 'OK' -and -not [string]::IsNullOrEmpty($script:capturedPrimaryKey)) 
        {
            return (GetKeyCombinationString $script:capturedModifierKeys $script:capturedPrimaryKey)
        }
        else
        {
            return $currentKey  
        }
    }
     finally
    {
        try
        {
            [Custom.HotkeyManager]::AreHotkeysGloballyPaused = $false
            if ($script:KeyCapture_PausedSnapshot -and $script:KeyCapture_PausedSnapshot.Count -gt 0)
            {
                ResumePausedKeys -Keys $script:KeyCapture_PausedSnapshot
            }
            else
            {
                ResumeAllHotkeys
            }
        }
        catch
        {
            Write-Verbose ('HOTKEYS: ResumeAllHotkeys (or ResumePausedKeys) failed: {0}' -f $_.Exception.Message)
        }
        finally
        {
            Remove-Variable -Name KeyCapture_PausedSnapshot -Scope Script -ErrorAction SilentlyContinue
            $script:capturedPrimaryKey = $null
            $script:capturedModifierKeys = @()
        }
    }
}

Export-ModuleMember -Function *
