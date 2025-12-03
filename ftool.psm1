<# ftool.psm1 
	.SYNOPSIS
		Interface for Ftool.

	.DESCRIPTION
		This module creates and manages the complete ftool interface for Entropia Dashboard:
		- Builds the main application window and all dialog forms
		- Creates interactive controls (buttons, panels, grids, text boxes)
		- Handles window positioning
		- Handles hotkeys
		- Implements settings management
		- Activates Ftool

	.NOTES
		Author: Immortal / Divine
		Version: 1.2
		Requires: PowerShell 5.1, .NET Framework 4.5+, classes.psm1, ini.psm1, datagrid.psm1, ftool.dll
#>

#region Hotkey Management

# Add C# code for a hidden message window to handle hotkey messages
if (-not ([System.Management.Automation.PSTypeName]'HotkeyManager').Type) {
    Add-Type -TypeDefinition @"
	using System;
	using System.Collections.Generic;
	using System.Runtime.InteropServices;
	using System.Windows.Forms;

	public class HotkeyManager : IDisposable
	{
		public const int WM_HOTKEY = 0x0312;
		private static int _nextId = 1;
		private static readonly MessageWindow _window = new MessageWindow();
		
		private static DateTime _lastHotkeyTime = DateTime.MinValue; 
		
		private static Dictionary<int, List<HotkeyActionEntry>> _hotkeyActions = new Dictionary<int, List<HotkeyActionEntry>>();
		private static Dictionary<Tuple<System.UInt32, System.UInt32>, int> _hotkeyIds = new Dictionary<Tuple<System.UInt32, System.UInt32>, int>();
		private static readonly object _hotkeyActionsLock = new object();
		private static bool _isProcessingHotkey = false;

		public enum HotkeyActionType
		{
			Normal,
			GlobalToggle
		}

		public struct HotkeyActionEntry
		{
			public Action ActionDelegate;
			public HotkeyActionType Type;
		}
        
        public static bool AreHotkeysGloballyPaused { get; set; }

		[DllImport("user32.dll", SetLastError=true)]
		private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

		[DllImport("user32.dll", SetLastError=true)]
		private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

		public static int Register(System.UInt32 modifiers, System.UInt32 virtualKey, HotkeyActionEntry actionEntry)
		{
			lock (_hotkeyActionsLock)
			{
				var hotkeyTuple = Tuple.Create(modifiers, virtualKey);
				if (_hotkeyIds.ContainsKey(hotkeyTuple))
				{
					int existingId = _hotkeyIds[hotkeyTuple];
					if (!_hotkeyActions.ContainsKey(existingId)) _hotkeyActions[existingId] = new List<HotkeyActionEntry>();
					_hotkeyActions[existingId].Add(actionEntry);
					return existingId;
				}

				int id = _nextId++;
				if (!RegisterHotKey(_window.Handle, id, modifiers, virtualKey))
				{
					throw new Exception("Failed to register hotkey. Error: " + Marshal.GetLastWin32Error());
				}
				_hotkeyActions[id] = new List<HotkeyActionEntry>() { actionEntry };
				_hotkeyIds[hotkeyTuple] = id;
				return id;
			}
		}

		public static void Unregister(int id)
		{
			lock (_hotkeyActionsLock)
			{
				if (!_hotkeyActions.ContainsKey(id)) return;
				UnregisterHotKey(_window.Handle, id);
				_hotkeyActions.Remove(id);

				Tuple<uint, uint> keyToRemove = null;
				foreach(var pair in _hotkeyIds)
				{
					if(pair.Value == id)
					{
						keyToRemove = pair.Key;
						break;
					}
				}
				if(keyToRemove != null)
				{
					_hotkeyIds.Remove(keyToRemove);
				}
			}
		}

		public static void UnregisterAction(int id, Action action)
		{
			lock (_hotkeyActionsLock)
			{
				if (!_hotkeyActions.ContainsKey(id)) return;
				var list = _hotkeyActions[id];
				list.RemoveAll(entry => entry.ActionDelegate == action);
				if (list.Count == 0)
				{
					UnregisterHotKey(_window.Handle, id);
					_hotkeyActions.Remove(id);
					Tuple<uint, uint> keyToRemove = null;
					foreach(var pair in _hotkeyIds)
					{
						if(pair.Value == id)
						{
							keyToRemove = pair.Key;
							break;
						}
					}
					if(keyToRemove != null)
					{
						_hotkeyIds.Remove(keyToRemove);
					}
				}
			}
		}

		public static void UnregisterAll()
		{
			lock (_hotkeyActionsLock)
			{
				foreach (var id in new List<int>(_hotkeyActions.Keys))
				{
					UnregisterHotKey(_window.Handle, id);
				}
				_hotkeyActions.Clear();
				_hotkeyIds.Clear();
			}
		}

		public void Dispose()
		{
			UnregisterAll();
			_window.Dispose();
		}

		private class MessageWindow : Form
		{
			protected override CreateParams CreateParams
			{
				get
				{
					var cp = base.CreateParams;
					cp.Parent = (IntPtr)(-3); // HWND_MESSAGE
					return cp;
				}
			}

			protected override void WndProc(ref Message m)
			{
				if (m.Msg == WM_HOTKEY)
				{
					if ((DateTime.Now - HotkeyManager._lastHotkeyTime).TotalMilliseconds < 300) 
					{
						return;
					}

					if (_isProcessingHotkey) {
						return;
					}

					try
					{
						_isProcessingHotkey = true;
						HotkeyManager._lastHotkeyTime = DateTime.Now;

						int id = m.WParam.ToInt32();
						if (_hotkeyActions.ContainsKey(id) && _hotkeyActions[id] != null)
						{
							List<HotkeyActionEntry> actionsToInvokeCopy;
							lock (_hotkeyActionsLock)
							{
								if (!_hotkeyActions.ContainsKey(id)) return;
								actionsToInvokeCopy = new List<HotkeyActionEntry>(_hotkeyActions[id]);
							}
							
							List<HotkeyActionEntry> globalToggleActions = new List<HotkeyActionEntry>();
							List<HotkeyActionEntry> normalActions = new List<HotkeyActionEntry>();

							foreach (var entry in actionsToInvokeCopy)
							{
								if (entry.Type == HotkeyActionType.GlobalToggle)
									globalToggleActions.Add(entry);
								else
									normalActions.Add(entry);
							}

							if (globalToggleActions.Count > 0)
							{
								try 
								{
									if (globalToggleActions[0].ActionDelegate != null) 
									{
										globalToggleActions[0].ActionDelegate.Invoke();
									}
								} 
								catch { }
								return; 
							}
							
							if (HotkeyManager.AreHotkeysGloballyPaused) 
							{
								return; 
							}

							foreach (var entry in normalActions)
							{
								try { if (entry.ActionDelegate != null) { entry.ActionDelegate.Invoke(); } } catch {}
							}
						}
					}
					finally
					{
						_isProcessingHotkey = false;
					}
				}
				base.WndProc(ref m);
			}
		}
	}
"@ -ReferencedAssemblies "System.Windows.Forms", "System.Drawing"
}

# Do not overwrite existing hotkey maps if they already exist from a previous instance
if (-not $global:RegisteredHotkeys) { $global:RegisteredHotkeys = @{} }
if (-not $global:RegisteredHotkeyByString) { $global:RegisteredHotkeyByString = @{} }

function Set-Hotkey
{
    param(
        [Parameter(Mandatory=$true)]
        [string]$KeyCombinationString,
        [Parameter(Mandatory=$true)]
        $Action,
		[Parameter(Mandatory=$false)]
		[string]$OwnerKey = $null,
        [Parameter(Mandatory=$false)]
        [int]$OldHotkeyId = $null
    )

    # 1. Normalize Action to a specific Delegate instance.
    if ($Action -is [System.Action]) {
        $actionDelegate = $Action
    } elseif ($Action -is [System.Delegate]) {
        $actionDelegate = $Action -as [System.Action]
    } else {
        $actionDelegate = [System.Action]$Action
    }

    # 2. Unregistration
    if (-not [string]::IsNullOrEmpty($OwnerKey)) {
        $idsToClean = @()
        foreach ($kvp in $global:RegisteredHotkeys.GetEnumerator()) {
            if ($kvp.Value.Owners -and $kvp.Value.Owners.ContainsKey($OwnerKey)) {
                $idsToClean += $kvp.Key
            }
        }

        foreach ($id in $idsToClean) {
            Write-Verbose "FTOOL: Found existing registration for owner '$OwnerKey' on ID $id. Removing..."
            
            $ownerAction = $global:RegisteredHotkeys[$id].Owners[$OwnerKey]
            
            $unregisteredSpecific = $false
            try {
                if ($ownerAction) {
                    $unregisterMethod = [HotkeyManager].GetMethod('UnregisterAction')
                    if ($unregisterMethod) {
                        [HotkeyManager]::UnregisterAction($id, $ownerAction.ActionDelegate)
                        $unregisteredSpecific = $true
                    }
                }
            } catch { Write-Verbose "FTOOL: Specific unregister failed: $unregisteredSpecific $_" }

            if ($global:RegisteredHotkeys.ContainsKey($id)) {
                $global:RegisteredHotkeys[$id].Owners.Remove($OwnerKey)
                
                if ($global:RegisteredHotkeys[$id].Owners.Count -eq 0) {
                    $unregisteredSuccessfully = $false
                    try {
                        [HotkeyManager]::Unregister($id)
                        $unregisteredSuccessfully = $true
                    } catch { Write-Warning "FTOOL: Failed to unregister OS hotkey ID $id for '$OwnerKey'. Error: $_" }

                    if ($unregisteredSuccessfully) {
                        $ks = $global:RegisteredHotkeys[$id].KeyString
                        if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) }
                        $global:RegisteredHotkeys.Remove($id)
                    }
                }
            }
        }
    }

    if ($global:PausedRegisteredHotkeys) {
        $pausedKeys = @($global:PausedRegisteredHotkeys.Keys)
        foreach ($k in $pausedKeys) {
            if (-not $global:PausedRegisteredHotkeys.ContainsKey($k)) { continue }
            $meta = $global:PausedRegisteredHotkeys[$k]
            if ($meta -and $meta.Owners -and $meta.Owners.ContainsKey($OwnerKey)) {
                Write-Verbose "FTOOL: Found and removed owner '$OwnerKey' from paused hotkey (fakeId $k)."
                $meta.Owners.Remove($OwnerKey)
                if ($meta.Owners.Count -eq 0) {
                    $global:PausedRegisteredHotkeys.Remove($k)
                    $ks = $meta.KeyString
                    if ($ks -and $global:RegisteredHotkeyByString.ContainsKey($ks) -and $global:RegisteredHotkeyByString[$ks] -eq $k) {
                        $global:RegisteredHotkeyByString.Remove($ks)
                    }
                }
            }
        }
    }
    
    if ($OldHotkeyId -ne $null -and $OldHotkeyId -ne 0) {
        if ($global:RegisteredHotkeys.ContainsKey($OldHotkeyId)) {
             try { [HotkeyManager]::Unregister($OldHotkeyId) } catch {}
             $global:RegisteredHotkeys.Remove($OldHotkeyId)
        }
    }

	# 3. Parse New Hotkey String
	$parsedResult = ParseKeyString $KeyCombinationString
	$modsFromParse = @()
	if ($parsedResult.Modifiers) { if ($parsedResult.Modifiers -is [System.Array]) { $modsFromParse = $parsedResult.Modifiers } else { $modsFromParse = @($parsedResult.Modifiers) } }
	[System.Collections.ArrayList]$parsedModifierKeys = @(); if ($modsFromParse.Count -gt 0) { $parsedModifierKeys.AddRange($modsFromParse) }
	[string]$parsedPrimaryKey = if ($parsedResult.Primary) { [string]$parsedResult.Primary } else { $null }

    if ([string]::IsNullOrEmpty($parsedPrimaryKey) -and $parsedModifierKeys.Count -gt 0) { throw "A primary key is required." }
    if ([string]::IsNullOrEmpty($parsedPrimaryKey) -and $parsedModifierKeys.Count -eq 0) { throw "No key provided." }

	[uint32]$mod = 0
    if($parsedModifierKeys){
        foreach ($mKey in $parsedModifierKeys) {
            switch ($mKey) { 'Alt' { $mod+=1 } 'Ctrl' { $mod+=2 } 'Shift' { $mod+=4 } 'Win' { $mod+=8 } }
        }
    }

	$vkNameFromNormalized = if ($parsedResult.Normalized) { ($parsedResult.Normalized -split '\+')[-1] } else { $null }
	$virtualKeyName = if (-not [string]::IsNullOrEmpty($vkNameFromNormalized)) { $vkNameFromNormalized } else { $parsedPrimaryKey }
	$virtualKey = (Get-VirtualKeyMappings)[([string]$virtualKeyName).ToUpper()]
	if (-not $virtualKey) { throw "Invalid primary key: $virtualKeyName" }

	$canonicalOrder = @('Ctrl','Alt','Shift','Win')
	$canonicalModifiers = @(); foreach ($m in $canonicalOrder) { if ($parsedModifierKeys -contains $m) { $canonicalModifiers += $m } }
	$verboseKeyString = if ($canonicalModifiers.Count -gt 0) { "$(($canonicalModifiers -join ' + '))+$parsedPrimaryKey" } else { $parsedPrimaryKey }
	$normalizedKeyString = $verboseKeyString.ToUpper()

	# 4. Register New Hotkey
    $actionType = [HotkeyManager+HotkeyActionType]::Normal
    if ($OwnerKey -and $OwnerKey -like "global_toggle_*") {
        $actionType = [HotkeyManager+HotkeyActionType]::GlobalToggle
    }

    $hotkeyActionEntry = New-Object HotkeyManager+HotkeyActionEntry
    $hotkeyActionEntry.ActionDelegate = $actionDelegate
    $hotkeyActionEntry.Type = $actionType

	$ownerToggleOn = $true
	if ($OwnerKey) {
		$instanceId = $OwnerKey
		if ($OwnerKey -like 'ext_*') {
			$parts = $OwnerKey -split '_'
			if ($parts.Count -ge 2) { $instanceId = $parts[1] }
		}
		try {
			if ($global:DashboardConfig.Resources.InstanceHotkeysPaused -and $global:DashboardConfig.Resources.InstanceHotkeysPaused.Contains($instanceId)) {
				$ownerToggleOn = -not [bool]$global:DashboardConfig.Resources.InstanceHotkeysPaused[$instanceId]
			} else {
				if ($global:DashboardConfig.Resources.FtoolForms -and $global:DashboardConfig.Resources.FtoolForms.Contains($instanceId)) {
					$form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
					if ($form -and $form.Tag -and $form.Tag.BtnHotkeyToggle) {
						$ownerToggleOn = [bool]$form.Tag.BtnHotkeyToggle.Checked
					}
				}
			}
		} catch { }
	}

	if (-not $ownerToggleOn) {
		$existingPausedId = $null
		if ($global:PausedRegisteredHotkeys) {
			foreach ($k in $global:PausedRegisteredHotkeys.Keys) {
				try {
					if ($global:PausedRegisteredHotkeys[$k].KeyString -and $global:PausedRegisteredHotkeys[$k].KeyString -eq $normalizedKeyString) {
						$existingPausedId = $k
						break
					}
				} catch {}
			}
		}
		
		if ($existingPausedId) {
			$metaToUpdate = $global:PausedRegisteredHotkeys[$existingPausedId]
			if (-not $metaToUpdate.Owners) { $metaToUpdate.Owners = @{} }
			
            $updatedHotkeyActionEntry = New-Object HotkeyManager+HotkeyActionEntry
			$updatedHotkeyActionEntry.ActionDelegate = $actionDelegate
			$updatedHotkeyActionEntry.Type = $actionType

			$metaToUpdate.Owners[$OwnerKey] = $updatedHotkeyActionEntry
			
			Write-Verbose "FTOOL: Registered hotkey (paused, updated) for owner '$OwnerKey' on key '$normalizedKeyString'."
			return $existingPausedId
		} else {
			if (-not $global:NextPausedFakeId) { $global:NextPausedFakeId = -1 }
			$fakeId = $global:NextPausedFakeId
			$global:NextPausedFakeId = $global:NextPausedFakeId - 1
			$meta = @{ Modifier = $mod; Key = $virtualKey; KeyString = $normalizedKeyString; Owners = @{}; Action = $actionDelegate; ActionType = $actionType }
			$meta.Owners[$OwnerKey] = $hotkeyActionEntry
			if (-not $global:PausedRegisteredHotkeys) { $global:PausedRegisteredHotkeys = @{} }
			$global:PausedRegisteredHotkeys[$fakeId] = $meta
			$global:RegisteredHotkeyByString[$normalizedKeyString] = $fakeId
			$id = $fakeId
			Write-Verbose "FTOOL: Stored hotkey (paused) for owner $OwnerKey on key $normalizedKeyString."
			return $id
		}
	}

	if ($global:RegisteredHotkeyByString.ContainsKey($normalizedKeyString)) {
		$existingId = $global:RegisteredHotkeyByString[$normalizedKeyString]
		if ($OwnerKey -and $global:RegisteredHotkeys.ContainsKey($existingId) -and $global:RegisteredHotkeys[$existingId].Owners.ContainsKey($OwnerKey)) {
			$staleAction = $global:RegisteredHotkeys[$existingId].Owners[$OwnerKey]
			try { [HotkeyManager]::UnregisterAction($existingId, $staleAction.ActionDelegate) } catch {}
			$global:RegisteredHotkeys[$existingId].Owners.Remove($OwnerKey)
		}
		try {
			$id = [HotkeyManager]::Register($mod, $virtualKey, $hotkeyActionEntry)
		} catch {
			Write-Warning "FTOOL: Failed to register hotkey $($KeyCombinationString): $_"
			$id = $null
			throw 
		}
	} else {
		try {
			$id = [HotkeyManager]::Register($mod, $virtualKey, $hotkeyActionEntry)
		} catch {
			Write-Warning "FTOOL: Failed to register hotkey $($KeyCombinationString): $_"
			$id = $null
			throw 
		}
	}

	# 5. Update Maps - Handle re-sync if C# returns existing ID
	if ($global:RegisteredHotkeys.ContainsKey($id)) {
		$existingKeyString = $global:RegisteredHotkeys[$id].KeyString
		if ($existingKeyString -and $existingKeyString -ne $normalizedKeyString) {
			$errorMsg = "FTOOL: CRITICAL hotkey conflict. ID $id."
			Write-Warning $errorMsg
			throw $errorMsg
		}
	}

	if (-not $global:RegisteredHotkeys.ContainsKey($id)) {
		$global:RegisteredHotkeys[$id] = @{Modifier=$mod; Key=$virtualKey; KeyString=$normalizedKeyString; Owners = @{}; Action = $actionDelegate; ActionType = $actionType}
	} else {
		if (-not $global:RegisteredHotkeys[$id].KeyString) { $global:RegisteredHotkeys[$id].KeyString = $normalizedKeyString }
	}
	
	if ($OwnerKey) {
		if (-not $global:RegisteredHotkeys[$id].Owners) { $global:RegisteredHotkeys[$id].Owners = @{} }
		$global:RegisteredHotkeys[$id].Owners[$OwnerKey] = $hotkeyActionEntry
	}
	$global:RegisteredHotkeyByString[$normalizedKeyString] = $id

	try {
		if ($OwnerKey -and $global:PausedRegisteredHotkeys) {
			$pausedKeys = @($global:PausedRegisteredHotkeys.Keys)
			foreach ($k in $pausedKeys) {
				if (-not $global:PausedRegisteredHotkeys.ContainsKey($k)) { continue }
				$meta = $global:PausedRegisteredHotkeys[$k]
				if ($meta -and $meta.KeyString -and $meta.KeyString -eq $normalizedKeyString) {
					if ($meta.Owners -and $meta.Owners.ContainsKey($OwnerKey)) {
						$meta.Owners.Remove($OwnerKey)
						if ($meta.Owners.Count -eq 0) {
							$global:PausedRegisteredHotkeys.Remove($k)
						}
						break 
					}
				}
			}
		}
	} catch {}

	return $id
}

function Resume-AllHotkeys {
	if (-not $global:PausedRegisteredHotkeys -or $global:PausedRegisteredHotkeys.Count -eq 0) { return }
	if (-not $global:PausedIdMap) { $global:PausedIdMap = @{} }

	$pausedKeys = @($global:PausedRegisteredHotkeys.Keys)
	foreach ($oldId in $pausedKeys) {
		try {
			$meta  = $global:PausedRegisteredHotkeys[$oldId]
			if (-not $meta) { continue }

			$mod    = $meta.Modifier
			$vk     = $meta.Key
			$ks     = $meta.KeyString
			$owners = if ($meta.Owners) { @($meta.Owners.Keys) } else { @() }

			if (-not $global:RegisteredHotkeys) { $global:RegisteredHotkeys = @{} }

			foreach ($ok in $owners) {
				$shouldResume = $true
				try {
					$instId = $ok
					if ($ok -like 'ext_*') { $parts = $ok -split '_'; if ($parts.Count -ge 2) { $instId = $parts[1] } }
					if ($global:DashboardConfig.Resources.InstanceHotkeysPaused -and $global:DashboardConfig.Resources.InstanceHotkeysPaused.Contains($instId)) {
						if ([bool]$global:DashboardConfig.Resources.InstanceHotkeysPaused[$instId]) { $shouldResume = $false }
					}
				} catch { }

				if (-not $shouldResume) { continue }

				$ownerAct = $meta.Owners[$ok]
				if (-not $ownerAct) { continue }
				
                # FIX: Explicit variable capture for closure
                $actionToRun = $ownerAct.ActionDelegate
				$delegate = [System.Action]({ Invoke-FtoolAction -Action $actionToRun }.GetNewClosure())

				try {
					$hotkeyActionEntry = New-Object HotkeyManager+HotkeyActionEntry
					$hotkeyActionEntry.ActionDelegate = $delegate
					$hotkeyActionEntry.Type = if ($ownerAct.Type) { $ownerAct.Type } else { [HotkeyManager+HotkeyActionType]::Normal }

					$registeredId = $null
					try {
						$registeredId = [HotkeyManager]::Register($mod, $vk, $hotkeyActionEntry)
					} catch {
						Write-Warning ("FTOOL: Resume-AllHotkeys failed owner '{0}': {1}" -f $ok, $_.Exception.Message)
					}

					if (-not $registeredId) { continue }
					
					$global:PausedIdMap[$oldId] = $registeredId
					if (-not $global:RegisteredHotkeys.ContainsKey($registeredId)) {
						$global:RegisteredHotkeys[$registeredId] = @{ Modifier = $mod; Key = $vk; KeyString = $ks; Owners = @{}; Action = $null; ActionType = $hotkeyActionEntry.Type }
					}
					$global:RegisteredHotkeys[$registeredId].Owners[$ok] = $hotkeyActionEntry
					if ($ks) { $global:RegisteredHotkeyByString[$ks] = $registeredId }
					$global:RegisteredHotkeys[$registeredId].ActionType = $hotkeyActionEntry.Type 

					$meta.Owners.Remove($ok)
				} catch {}
			}

			if (-not $meta.Owners -or $meta.Owners.Count -eq 0) {
				$global:PausedRegisteredHotkeys.Remove($oldId)
			} else {
				$global:PausedRegisteredHotkeys[$oldId] = $meta
			}
		} catch {}
	}
}

function Resume-PausedKeys {
	param([Parameter(Mandatory=$true)][array]$Keys)
	if (-not $Keys -or $Keys.Count -eq 0) { return }
	if (-not $global:PausedIdMap) { $global:PausedIdMap = @{} }
	if (-not $global:RegisteredHotkeys) { $global:RegisteredHotkeys = @{} }

	foreach ($oldId in $Keys) {
		if (-not $global:PausedRegisteredHotkeys.Contains($oldId)) { continue }
		try {
			$meta = $global:PausedRegisteredHotkeys[$oldId]
			if (-not $meta) { continue }

			$mod = $meta.Modifier; $vk = $meta.Key; $ks = $meta.KeyString
			$owners = if ($meta.Owners) { @($meta.Owners.Keys) } else { @() }

			foreach ($ok in $owners) {
				$shouldResume = $true
				try {
					$instId = $ok
					if ($ok -like 'ext_*') { $parts = $ok -split '_'; if ($parts.Count -ge 2) { $instId = $parts[1] } }
					if ($global:DashboardConfig.Resources.InstanceHotkeysPaused -and $global:DashboardConfig.Resources.InstanceHotkeysPaused.Contains($instId)) {
						if ([bool]$global:DashboardConfig.Resources.InstanceHotkeysPaused[$instId]) { $shouldResume = $false }
					}
				} catch {}

				if (-not $shouldResume) { continue }

				$ownerAct = $meta.Owners[$ok]
				if (-not $ownerAct) { continue }
				
                $actionToRun = $ownerAct.ActionDelegate
				$delegate = [System.Action]({ Invoke-FtoolAction -Action $actionToRun }.GetNewClosure())
								
				try {
					$hotkeyActionEntry = New-Object HotkeyManager+HotkeyActionEntry
					$hotkeyActionEntry.ActionDelegate = $delegate
					$hotkeyActionEntry.Type = if ($meta.ActionType) { $meta.ActionType } else { [HotkeyManager+HotkeyActionType]::Normal }
				
					$registeredId = [HotkeyManager]::Register($mod, $vk, $hotkeyActionEntry)
					$global:PausedIdMap[$oldId] = $registeredId
					if (-not $global:RegisteredHotkeys.ContainsKey($registeredId)) {
						$global:RegisteredHotkeys[$registeredId] = @{ Modifier = $mod; Key = $vk; KeyString = $ks; Owners = @{}; Action = $null; ActionType = $hotkeyActionEntry.Type }
					}
					$global:RegisteredHotkeys[$registeredId].Owners[$ok] = $hotkeyActionEntry
					$global:RegisteredHotkeys[$registeredId].ActionType = $hotkeyActionEntry.Type
					if ($ks) { $global:RegisteredHotkeyByString[$ks] = $registeredId }
					
					$meta.Owners.Remove($ok)
				} catch {}
			}

			if (-not $meta.Owners -or $meta.Owners.Count -eq 0) {
				$global:PausedRegisteredHotkeys.Remove($oldId)
			} else {
				$global:PausedRegisteredHotkeys[$oldId] = $meta
			}
		} catch {}
	}
}

function Resume-HotkeysForOwner {
	param([Parameter(Mandatory=$true)][string]$OwnerKey)
	if (-not $OwnerKey) { return }
	if (-not $global:PausedRegisteredHotkeys -or $global:PausedRegisteredHotkeys.Count -eq 0) { return }
	if (-not $global:PausedIdMap) { $global:PausedIdMap = @{} }
	
	$idsToResume = @()
	foreach ($kvp in $global:PausedRegisteredHotkeys.GetEnumerator()) {
		$oldId = $kvp.Key
		$meta = $kvp.Value
		if ($meta.Owners -and $meta.Owners.ContainsKey($OwnerKey)) {
			$idsToResume += $oldId
		}
	}

	foreach ($oldId in $idsToResume) {
		try {
			$meta = $global:PausedRegisteredHotkeys[$oldId]
			$mod = $meta.Modifier; $vk = $meta.Key; $ks = $meta.KeyString; $owners = $meta.Owners

			foreach ($ok in $owners.Keys) {
				$ownerAct = $owners[$ok]
				if (-not $ownerAct) { continue }

                if ($ownerAct -is [HotkeyManager+HotkeyActionEntry]) {
                    $actionToRun = $ownerAct.ActionDelegate
                } else {
                    $actionToRun = $ownerAct
                }
				$delegate = [System.Action]({ Invoke-FtoolAction -Action $actionToRun }.GetNewClosure())
				
				try {
					$hotkeyActionEntry = New-Object HotkeyManager+HotkeyActionEntry
					$hotkeyActionEntry.ActionDelegate = $delegate
					$hotkeyActionEntry.Type = if ($meta.ActionType) { $meta.ActionType } else { [HotkeyManager+HotkeyActionType]::Normal }

					$registeredId = [HotkeyManager]::Register($mod, $vk, $hotkeyActionEntry)
					$global:PausedIdMap[$oldId] = $registeredId
					if (-not $global:RegisteredHotkeys.ContainsKey($registeredId)) {
						$global:RegisteredHotkeys[$registeredId] = @{Modifier=$mod; Key=$vk; KeyString=$ks; Owners = @{}; Action = $null; ActionType = $hotkeyActionEntry.Type }
					}
					$global:RegisteredHotkeys[$registeredId].Owners[$ok] = $hotkeyActionEntry
                    $global:RegisteredHotkeys[$registeredId].ActionType = $hotkeyActionEntry.Type
					if ($ks) { $global:RegisteredHotkeyByString[$ks] = $registeredId }
				} catch {}
			}
			$global:PausedRegisteredHotkeys.Remove($oldId)
		} catch {}
	}
}

function Remove-AllHotkeys {
    [HotkeyManager]::UnregisterAll()
    $global:RegisteredHotkeys.Clear()
	$global:RegisteredHotkeyByString.Clear()
    Write-Verbose "FTOOL: All hotkeys unregistered."
}

function Test-HotkeyConflict {
    param(
        [Parameter(Mandatory=$true)]
        [string]$KeyCombinationString,
        [Parameter(Mandatory=$true)]
        [HotkeyManager+HotkeyActionType]$NewHotkeyType,
        [Parameter(Mandatory=$false)]
        [string]$OwnerKeyToExclude
    )

    $normalizedKeyString = (NormalizeKeyString $KeyCombinationString).ToUpper()
    if (-not $normalizedKeyString) { return $false }

    foreach ($id in $global:RegisteredHotkeys.Keys) {
        $meta = $global:RegisteredHotkeys[$id]
        if ($meta -and $meta.KeyString -eq $normalizedKeyString) {
            $existingHotkeyType = if ($meta.ActionType) { $meta.ActionType } else { [HotkeyManager+HotkeyActionType]::Normal }

            $isConflict = $false
            if (!($NewHotkeyType -eq [HotkeyManager+HotkeyActionType]::Normal -and $existingHotkeyType -eq [HotkeyManager+HotkeyActionType]::GlobalToggle)) {
                $isConflict = $true
            }
            
            if ($isConflict) {
                if ($OwnerKeyToExclude) {
                    if ($meta.Owners.ContainsKey($OwnerKeyToExclude)) {
                        continue 
                    }
                }
                return $true
            }
        }
    }

    if ($global:PausedRegisteredHotkeys) {
        foreach ($id in $global:PausedRegisteredHotkeys.Keys) {
            $meta = $global:PausedRegisteredHotkeys[$id]
            if ($meta -and $meta.KeyString -eq $normalizedKeyString) {
                $existingHotkeyType = if ($meta.ActionType) { $meta.ActionType } else { [HotkeyManager+HotkeyActionType]::Normal }

                $isConflict = $false
                if (!($NewHotkeyType -eq [HotkeyManager+HotkeyActionType]::Normal -and $existingHotkeyType -eq [HotkeyManager+HotkeyActionType]::GlobalToggle)) {
                    $isConflict = $true
                }

                if ($isConflict) {
                    if ($OwnerKeyToExclude) {
                        if ($meta.Owners.ContainsKey($OwnerKeyToExclude)) {
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


function Invoke-FtoolAction {
	param([Parameter(Mandatory=$true)]$Action)
	try {
		if ($Action -is [ScriptBlock]) { 
			& $Action 
		} elseif ($Action -is [System.Delegate] -or $Action -is [System.Action]) { 
			$Action.Invoke() 
		} elseif ($Action -is [string]) {
			$cmd = Get-Command $Action -ErrorAction SilentlyContinue
			if ($cmd) { & $cmd.Name } else { Invoke-Expression $Action }
		}
	} catch {
		Write-Verbose ("FTOOL-DEBUG: Invoke-FtoolAction: Exception invoking action: {0}" -f $_.Exception.Message)
	}
}

function PauseAllHotkeys {
	[HotkeyManager]::AreHotkeysGloballyPaused = $true
	if (-not $global:RegisteredHotkeys -or $global:RegisteredHotkeys.Count -eq 0) { return }
	
	$global:PausedIdMap = @{}
	if (-not $global:PausedRegisteredHotkeys) { $global:PausedRegisteredHotkeys = @{} }

	$idsToProcess = @($global:RegisteredHotkeys.Keys)

	foreach ($id in $idsToProcess) {
		try {
			$meta = $global:RegisteredHotkeys[$id]
			if (-not $meta -or -not $meta.Owners) { continue }

            $ownersToRemove = @()
            $ownerEntriesToPause = @{}
            $shouldFullyUnregisterOSHotkey = $true
            
            $ownerKeys = @($meta.Owners.Keys) 
            foreach ($ownerKey in $ownerKeys) {
                $ownerEntry = $meta.Owners[$ownerKey]
                if ($null -eq $ownerEntry -or $null -eq $ownerEntry.ActionDelegate) { continue }

                try {
                    [HotkeyManager]::UnregisterAction($id, $ownerEntry.ActionDelegate)
                    $ownersToRemove += $ownerKey 
                    $ownerEntriesToPause[$ownerKey] = $ownerEntry
                } catch {
                    $shouldFullyUnregisterOSHotkey = $false
                }
            }

            foreach ($ownerKey in $ownersToRemove) {
                $meta.Owners.Remove($ownerKey)
            }

            $pausedMetaForId = @{
                Modifier = $meta.Modifier;
                Key = $meta.Key;
                KeyString = $meta.KeyString;
                Owners = @{}
                Action = $meta.Action;
                ActionType = $meta.ActionType;
            }
            foreach ($ownerKey in $ownerEntriesToPause.Keys) {
                $pausedMetaForId.Owners[$ownerKey] = $ownerEntriesToPause[$ownerKey]
            }

            if ($pausedMetaForId.Owners.Count -gt 0) {
                $global:PausedRegisteredHotkeys[$id] = $pausedMetaForId
                if ($meta.KeyString) { $global:RegisteredHotkeyByString.Remove($meta.KeyString) }
                $global:RegisteredHotkeyByString[$meta.KeyString] = $id
            }

            if ($shouldFullyUnregisterOSHotkey -and $meta.Owners.Count -eq 0) {
                try { [HotkeyManager]::Unregister($id) } catch {}
                $global:RegisteredHotkeys.Remove($id)
            } else {
                $global:RegisteredHotkeys[$id] = $meta
            }

		} catch {}
	}
}

function PauseHotkeysForOwner {
	param([Parameter(Mandatory=$true)][string]$OwnerKey)
	if (-not $OwnerKey) { return }
	if (-not $global:RegisteredHotkeys -or $global:RegisteredHotkeys.Count -eq 0) { return }
	if (-not $global:PausedRegisteredHotkeys) { $global:PausedRegisteredHotkeys = @{} }
	if (-not $global:PausedIdMap) { $global:PausedIdMap = @{} }

	$idsToProcess = @()
	foreach ($kvp in $global:RegisteredHotkeys.GetEnumerator()) {
		$id = $kvp.Key
		$meta = $kvp.Value
		if ($meta.Owners -and $meta.Owners.ContainsKey($OwnerKey)) {
			$idsToProcess += $id
		}
	}

	foreach ($id in $idsToProcess) {
		try {
			$meta = $global:RegisteredHotkeys[$id]
			$ks = $null
			try { $ks = $meta.KeyString } catch {}

			$ownerEntry = $null
			try { $ownerEntry = $meta.Owners[$OwnerKey] } catch {}

			if ($null -eq $ownerEntry -or $null -eq $ownerEntry.ActionDelegate) {
				continue
			}

			if (-not $global:PausedRegisteredHotkeys) { $global:PausedRegisteredHotkeys = @{} }
			if (-not $global:NextPausedFakeId) { $global:NextPausedFakeId = -1 }

			$unregisterActionMethod = $null
			try { $unregisterActionMethod = [HotkeyManager].GetMethod('UnregisterAction') } catch {}
			if ($unregisterActionMethod) {
					try {
						[HotkeyManager]::UnregisterAction($id, $ownerEntry.ActionDelegate)

						if ($global:RegisteredHotkeys.ContainsKey($id) -and $global:RegisteredHotkeys[$id].Owners.ContainsKey($OwnerKey)) {
							$global:RegisteredHotkeys[$id].Owners.Remove($OwnerKey)
						}

						$fakeId = $global:NextPausedFakeId; $global:NextPausedFakeId = $global:NextPausedFakeId - 1
						$pausedMeta = @{ Modifier = $meta.Modifier; Key = $meta.Key; KeyString = $ks; Owners = @{}; Action = $ownerEntry.ActionDelegate; ActionType = $ownerEntry.Type } # Store ActionDelegate and Type
						$pausedMeta.Owners[$OwnerKey] = $ownerEntry
						$global:PausedRegisteredHotkeys[$fakeId] = $pausedMeta
						if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) }
						$global:RegisteredHotkeyByString[$ks] = $fakeId
						continue
					} catch {}
			}

			$global:PausedRegisteredHotkeys[$id] = $meta
		    try { [HotkeyManager]::Unregister($id) } catch { }
		    try { if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) } } catch {}
		    try { $global:RegisteredHotkeys.Remove($id) } catch {}
		} catch {}
	}}

function Unregister-HotkeyInstance {
	param(
		[Parameter(Mandatory=$true)][int]$Id,
		[Parameter(Mandatory=$false)][string]$OwnerKey
	)
	if (-not $Id) { return }
	try {
		$didUnregister = $false
		$translateTarget = $null
		if ($global:PausedIdMap -and $global:PausedIdMap.ContainsKey($Id)) {
			$translateTarget = $global:PausedIdMap[$Id]
		}
		if ($translateTarget) { $IdToUse = $translateTarget } else { $IdToUse = $Id }
		
		if (-not $global:RegisteredHotkeys.ContainsKey($IdToUse) -and $global:PausedRegisteredHotkeys -and $global:PausedRegisteredHotkeys.ContainsKey($IdToUse)) {
			$pausedMeta = $global:PausedRegisteredHotkeys[$IdToUse]
			if ($OwnerKey -and $pausedMeta.Owners -and $pausedMeta.Owners.ContainsKey($OwnerKey)) {
				$pausedMeta.Owners.Remove($OwnerKey)
				if ($pausedMeta.Owners.Count -eq 0) { $global:PausedRegisteredHotkeys.Remove($IdToUse); }
				$didUnregister = $true
			} else {
				$global:PausedRegisteredHotkeys.Remove($IdToUse)
				$didUnregister = $true
			}
		}
        if ($OwnerKey -and $global:RegisteredHotkeys.ContainsKey($Id) -and $global:RegisteredHotkeys[$Id].Owners.ContainsKey($OwnerKey)) {
            $ownerEntry = $global:RegisteredHotkeys[$Id].Owners[$OwnerKey]
            $unregisterMethod = $null
            try { $unregisterMethod = [HotkeyManager].GetMethod('UnregisterAction') } catch {}
            if ($unregisterMethod) {
                try {
                    [HotkeyManager]::UnregisterAction($Id, $ownerEntry.ActionDelegate) 
                    $didUnregister = $true
                } catch {
                    Write-Warning "FTOOL: UnregisterAction failed. Fallback."
                }
            } 
            if ($global:RegisteredHotkeys.ContainsKey($Id) -and $global:RegisteredHotkeys[$Id].Owners.ContainsKey($OwnerKey)) {
                $global:RegisteredHotkeys[$Id].Owners.Remove($OwnerKey)
            }
            if ($global:RegisteredHotkeys.ContainsKey($Id) -and $global:RegisteredHotkeys[$Id].Owners.Count -eq 0) {
				$ks = $global:RegisteredHotkeys[$Id].KeyString
				if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) }
				$global:RegisteredHotkeys.Remove($Id)
			}
		}
		if (-not $didUnregister) {
			if ($global:RegisteredHotkeys.ContainsKey($Id)) {
                $unregisteredSuccessfully = $false
				try {
					[HotkeyManager]::Unregister($Id)
                    $unregisteredSuccessfully = $true
				} catch {}
				
                if ($unregisteredSuccessfully) {
                    if ($global:RegisteredHotkeys.ContainsKey($Id)) {
                        $ks = $global:RegisteredHotkeys[$Id].KeyString
                        if ($ks) { $global:RegisteredHotkeyByString.Remove($ks) }
                        $global:RegisteredHotkeys.Remove($Id)
                    }
                }
			}
		}
	} catch {
		Write-Warning "FTOOL: Failed to unregister hotkey ID $Id. Error: $_"
	}
}


function ToggleSpecificFtoolInstance {
    param(
        [Parameter(Mandatory=$true)]
        [string]$InstanceId,
        [Parameter(Mandatory=$false)]
        [string]$ExtKey
    )

    if (-not $global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId)) {
        return
    }

    $form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
    if ($form -and -not $form.IsDisposed) {
        $formData = $form.Tag

        if (-not [string]::IsNullOrEmpty($ExtKey)) {
            if ($global:DashboardConfig.Resources.ExtensionData.Contains($ExtKey)) {
                $extData = $global:DashboardConfig.Resources.ExtensionData[$ExtKey]
                if ($extData.RunningSpammer) {
                    $extData.BtnStop.PerformClick()
                } else {
                    $extData.BtnStart.PerformClick()
                }
            }
        } else {
            if ($formData.RunningSpammer) {
                $formData.BtnStop.PerformClick()
            } else {
                $formData.BtnStart.PerformClick()
            }
        }
    }
}


function ToggleInstanceHotkeys
{
    param(
        [Parameter(Mandatory=$true)]
        [string]$InstanceId,
        [Parameter(Mandatory=$true)]
        [bool]$ToggleState
    )
    
    if (-not $global:DashboardConfig.Resources.InstanceHotkeysPaused) { $global:DashboardConfig.Resources.InstanceHotkeysPaused = @{} }
    
    $global:DashboardConfig.Resources.InstanceHotkeysPaused[$InstanceId] = (-not $ToggleState)

    try {
        if ($ToggleState) {
            try { Resume-HotkeysForOwner -OwnerKey $InstanceId } catch {}
            if ($global:DashboardConfig.Resources.ExtensionData) {
                $extKeys = $global:DashboardConfig.Resources.ExtensionData.Keys | Where-Object { $_ -like "ext_${InstanceId}_*" }
                foreach ($extKey in $extKeys) {
                    try { Resume-HotkeysForOwner -OwnerKey $extKey } catch {}
                }
            }
        } else {
            try { PauseHotkeysForOwner -OwnerKey $InstanceId } catch {}
            if ($global:DashboardConfig.Resources.ExtensionData) {
                $extKeys = $global:DashboardConfig.Resources.ExtensionData.Keys | Where-Object { $_ -like "ext_${InstanceId}_*" }
                foreach ($extKey in $extKeys) {
                    try { PauseHotkeysForOwner -OwnerKey $extKey } catch {}
                }
            }
        }
    } catch {}

	if ($global:DashboardConfig.Resources.FtoolForms.Contains($InstanceId)) {
		$form = $global:DashboardConfig.Resources.FtoolForms[$InstanceId]
		if ($form -and -not $form.IsDisposed) {
			UpdateSettings -formData $form.Tag -forceWrite
		}
	}
}


#endregion

#region Helper Functions

function Get-VirtualKeyMappings
{
	return @{
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

		'SPACE' = 0x20; 'ENTER' = 0x0D; 'TAB' = 0x09; 'ESCAPE' = 0x1B; 'SHIFT' = 0x10; 'CONTROL' = 0x11; 'ALT' = 0x12
		'UP_ARROW' = 0x26; 'DOWN_ARROW' = 0x28; 'LEFT_ARROW' = 0x25; 'RIGHT_ARROW' = 0x27; 'HOME' = 0x24; 'END' = 0x23
		'PAGE_UP' = 0x21; 'PAGE_DOWN' = 0x22; 'INSERT' = 0x2D; 'DELETE' = 0x2E; 'BACKSPACE' = 0x08

		'CAPS_LOCK' = 0x14; 'NUM_LOCK' = 0x90; 'SCROLL_LOCK' = 0x91; 'PRINT_SCREEN' = 0x2C; 'PAUSE_BREAK' = 0x13
		'LEFT_WINDOWS' = 0x5B; 'RIGHT_WINDOWS' = 0x5C; 'APPLICATION' = 0x5D; 'LEFT_SHIFT' = 0xA0; 'RIGHT_SHIFT' = 0xA1
		'LEFT_CONTROL' = 0xA2; 'RIGHT_CONTROL' = 0xA3; 'LEFT_ALT' = 0xA4; 'RIGHT_ALT' = 0xA5; 'SLEEP' = 0x5F

		'NUMPAD_0' = 0x60; 'NUMPAD_1' = 0x61; 'NUMPAD_2' = 0x62; 'NUMPAD_3' = 0x63; 'NUMPAD_4' = 0x64
		'NUMPAD_5' = 0x65; 'NUMPAD_6' = 0x66; 'NUMPAD_7' = 0x67; 'NUMPAD_8' = 0x68; 'NUMPAD_9' = 0x69
		'NUMPAD_MULTIPLY' = 0x6A; 'NUMPAD_ADD' = 0x6B; 'NUMPAD_SEPARATOR' = 0x6C; 'NUMPAD_SUBTRACT' = 0x6D; 'NUMPAD_DECIMAL' = 0x6E; 'NUMPAD_DIVIDE' = 0x6F

		'SEMICOLON' = 0xBA; 'EQUALS' = 0xBB; 'COMMA' = 0xBC; 'MINUS' = 0xBD; 'PERIOD' = 0xBE
		'FORWARD_SLASH' = 0xBF; 'BACKTICK' = 0xC0; 'LEFT_BRACKET' = 0xDB; 'BACKSLASH' = 0xDC; 'RIGHT_BRACKET' = 0xDD
		'APOSTROPHE' = 0xDE

		'BROWSER_BACK' = 0xA6; 'BROWSER_FORWARD' = 0xA7; 'BROWSER_REFRESH' = 0xA8; 'BROWSER_STOP' = 0xA9
		'BROWSER_SEARCH' = 0xAA; 'BROWSER_FAVORITES' = 0xAB; 'BROWSER_HOME' = 0xAC; 'VOLUME_MUTE' = 0xAD
		'VOLUME_DOWN' = 0xAE; 'VOLUME_UP' = 0xAF; 'MEDIA_NEXT_TRACK' = 0xB0; 'MEDIA_PREVIOUS_TRACK' = 0xB1
		'MEDIA_STOP' = 0xB2; 'MEDIA_PLAY_PAUSE' = 0xB3; 'LAUNCH_MAIL' = 0xB4; 'LAUNCH_MEDIA_PLAYER' = 0xB5
		'LAUNCH_MY_COMPUTER' = 0xB6; 'LAUNCH_CALCULATOR' = 0xB7

		'IME_KANA_HANGUL' = 0x15; 'IME_JUNJA' = 0x17; 'IME_FINAL' = 0x18; 'IME_HANJA_KANJI' = 0x19
		'IME_CONVERT' = 0x1C; 'IME_NONCONVERT' = 0x1D; 'IME_ACCEPT' = 0x1E; 'IME_MODE_CHANGE' = 0x1F; 'IME_PROCESS' = 0xE5
		'SELECT' = 0x29; 'PRINT' = 0x2A; 'EXECUTE' = 0x2B; 'HELP' = 0x2F; 'CLEAR' = 0x0C
		'ATTN' = 0xF6; 'CRSEL' = 0xF7; 'EXSEL' = 0xF8; 'ERASE_EOF' = 0xF9; 'PLAY' = 0xFA; 'ZOOM' = 0xFB
		'PA1' = 0xFD; 'OEM_CLEAR' = 0xFE
	}
}

function NormalizeKeyString {
	param(
		[Parameter(Mandatory=$true)][string]$KeyCombinationString
	)

	if ([string]::IsNullOrEmpty($KeyCombinationString)) { return $null }
	$parts = @($KeyCombinationString -split '\s*\+\s*' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
	
	$parsedModifierKeys = @()
	$parsedPrimaryKey = $null
	foreach ($part in $parts) {
		switch ($part.ToUpper()) {
			'ALT'   { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'MENU'  { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'LALT'  { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'RALT'  { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'LEFT_ALT'  { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'RIGHT_ALT' { if ($parsedModifierKeys -notcontains 'Alt') { $parsedModifierKeys += 'Alt' } }
			'CTRL'  { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'CONTROL'{ if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'LCTRL' { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'RCTRL' { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'LEFT_CONTROL' { if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'RIGHT_CONTROL'{ if ($parsedModifierKeys -notcontains 'Ctrl') { $parsedModifierKeys += 'Ctrl' } }
			'SHIFT' { if ($parsedModifierKeys -notcontains 'Shift') { $parsedModifierKeys += 'Shift' } }
			'LSHIFT'{ if ($parsedModifierKeys -notcontains 'Shift') { $parsedModifierKeys += 'Shift' } }
			'RSHIFT'{ if ($parsedModifierKeys -notcontains 'Shift') { $parsedModifierKeys += 'Shift' } }
			'WIN'   { if ($parsedModifierKeys -notcontains 'Win') { $parsedModifierKeys += 'Win' } }
			'LWIN'  { if ($parsedModifierKeys -notcontains 'Win') { $parsedModifierKeys += 'Win' } }
			'RWIN'  { if ($parsedModifierKeys -notcontains 'Win') { $parsedModifierKeys += 'Win' } }
			'WINDOWS' { if ($parsedModifierKeys -notcontains 'Win') { $parsedModifierKeys += 'Win' } }
			default { 
                if ([string]::IsNullOrEmpty($parsedPrimaryKey)) { 
                    $parsedPrimaryKey = $part 
                } else {
                    $parsedPrimaryKey = "$parsedPrimaryKey $part"
                }
            }
		}
	}

	$canonicalOrder = @('Ctrl','Alt','Shift','Win')
	$canonicalModifiers = @()
	foreach ($m in $canonicalOrder) { if ($parsedModifierKeys -contains $m) { $canonicalModifiers += $m } }

	if ($canonicalModifiers.Count -gt 0) {
		if ([string]::IsNullOrEmpty($parsedPrimaryKey)) { return ([string]($canonicalModifiers -join ' + ')).ToUpper() }
		return ([string]("$(($canonicalModifiers -join ' + '))+$parsedPrimaryKey")).ToUpper()
	} else {
		return ([string]$parsedPrimaryKey).ToUpper()
	}
}

function ParseKeyString {
	param(
		[Parameter(Mandatory=$true)][string]$KeyCombinationString
	)

	$normalized = NormalizeKeyString $KeyCombinationString
	if (-not $normalized) { return @{Modifiers=@(); Primary=$null; Normalized=$null} }

	$parts = @($normalized -split '\s*\+\s*' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
	
	if ($parts.Count -eq 0) { return @{Modifiers=@(); Primary=$null; Normalized=$normalized} }

	if ($parts.Count -eq 1) {
		return @{Modifiers=@(); Primary=([string]$parts[0]).ToUpper(); Normalized=$normalized}
	}

	$primary = ([string]$parts[-1]).ToUpper()
	$mods = $parts[0..($parts.Count - 2)] | ForEach-Object {
		switch ($_.ToUpper()) {
			'CTRL' { 'Ctrl' }
			'CONTROL' { 'Ctrl' }
			'LCTRL' { 'Ctrl' }
			'RCTRL' { 'Ctrl' }
			'LEFT_CONTROL' { 'Ctrl' }
			'RIGHT_CONTROL' { 'Ctrl' }
			'ALT'  { 'Alt' }
			'MENU' { 'Alt' }
			'LALT'  { 'Alt' }
			'RALT'  { 'Alt' }
			'LEFT_ALT' { 'Alt' }
			'RIGHT_ALT' { 'Alt' }
			'SHIFT'{ 'Shift' }
			'LSHIFT'{ 'Shift' }
			'RSHIFT'{ 'Shift' }
			'WIN'  { 'Win' }
			'LWIN' { 'Win' }
			'RWIN' { 'Win' }
			'WINDOWS' { 'Win' }
			default { $_ }
		}
	}
	return @{Modifiers=$mods; Primary=$primary; Normalized=$normalized}
}

function Get-KeyCombinationString {
    param([string[]]$modifiers, [string]$primaryKey)
    $parts = @()
    if ($modifiers) { $parts += $modifiers }
    if ($primaryKey) { $parts += $primaryKey }
    if ($parts.Count -eq 0) { return "none" }
    return ($parts -join ' + ')
}

function IsModifierKeyCode {
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
		[string]$currentKey = '' 
	)
	try {
		PauseAllHotkeys
		$script:KeyCapture_PausedSnapshot = @()
		if ($global:PausedRegisteredHotkeys) { $script:KeyCapture_PausedSnapshot = @($global:PausedRegisteredHotkeys.Keys) }
	} catch {
		Write-Verbose ("FTOOL: Pause-AllHotkeys failed: {0}" -f $_.Exception.Message)
	}

    $script:capturedModifierKeys = @()
    $script:capturedPrimaryKey = $null

    if (-not [string]::IsNullOrEmpty($currentKey) -and $currentKey -ne "Hotkey" -and $currentKey -ne "none") {
        $parts = $currentKey.Split(' + ')
        foreach ($part in $parts) {
            switch ($part.ToUpper()) {
                'ALT'   { if ($script:capturedModifierKeys -notcontains 'Alt') { $script:capturedModifierKeys += 'Alt' } }
                'CTRL'  { if ($script:capturedModifierKeys -notcontains 'Ctrl') { $script:capturedModifierKeys += 'Ctrl' } }
                'SHIFT' { if ($script:capturedModifierKeys -notcontains 'Shift') { $script:capturedModifierKeys += 'Shift' } }
                'WIN'   { if ($script:capturedModifierKeys -notcontains 'Win') { $script:capturedModifierKeys += 'Win' } }
                default { if ([string]::IsNullOrEmpty($script:capturedPrimaryKey)) { $script:capturedPrimaryKey = $part }; break } 
            }
        }
    }

	$captureForm = New-Object System.Windows.Forms.Form
	$captureForm.Text = "Press a Key Combination" 
	$captureForm.Size = New-Object System.Drawing.Size(300, 150)
	$captureForm.StartPosition = 'CenterParent'
	$captureForm.FormBorderStyle = 'FixedDialog'
	$captureForm.MaximizeBox = $false
	$captureForm.MinimizeBox = $false
	$captureForm.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
	$captureForm.ForeColor = [System.Drawing.Color]::White
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Press any key combination.`n`nPress ESC to cancel." 
	$label.Size = New-Object System.Drawing.Size(280, 60)
	$label.Location = New-Object System.Drawing.Point(10, 10)
	$label.TextAlign = 'MiddleCenter'
	$label.Font = New-Object System.Drawing.Font('Segoe UI', 10)
	$captureForm.Controls.Add($label)
	
	$resultLabel = New-Object System.Windows.Forms.Label
	$resultLabel.Text = (Get-KeyCombinationString $script:capturedModifierKeys $script:capturedPrimaryKey) 
	$resultLabel.Size = New-Object System.Drawing.Size(280, 25)
	$resultLabel.Location = New-Object System.Drawing.Point(10, 75)
	$resultLabel.TextAlign = 'MiddleCenter'
	$resultLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
	$resultLabel.ForeColor = [System.Drawing.Color]::Yellow
	$captureForm.Controls.Add($resultLabel)
	
	$captureForm.Add_KeyDown({
        param($form, $e)

        if ($e.KeyCode -eq 'Escape') {
            $script:capturedPrimaryKey = $null
            $script:capturedModifierKeys = @()
            $captureForm.DialogResult = 'Cancel'
            $captureForm.Close()
            return
        }

        [System.Collections.ArrayList]$currentModifiersTemp = @()
        if ($e.Control) { $currentModifiersTemp.Add('Ctrl') }
        if ($e.Alt) { $currentModifiersTemp.Add('Alt') }
        if ($e.Shift) { $currentModifiersTemp.Add('Shift') }
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::LWin -or $e.KeyCode -eq [System.Windows.Forms.Keys]::RWin) {
            if ($currentModifiersTemp -notcontains 'Win') { $currentModifiersTemp.Add('Win') }
        }

        $keyMappings = Get-VirtualKeyMappings
        [string]$actualPressedKeyName = $null
        foreach ($kvp in $keyMappings.GetEnumerator()) {
            if ($kvp.Value -eq $e.KeyValue) {
                $actualPressedKeyName = $kvp.Key
                break
            }
        }
        if (-not $actualPressedKeyName) {
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


        if (-not $isPhysicalModifier -and -not $isNamedModifier) {
            $script:capturedPrimaryKey = $pressedKeyName
            $script:capturedModifierKeys = $currentModifiersTemp
            
            $resultLabel.Text = "Captured: $(Get-KeyCombinationString $script:capturedModifierKeys $script:capturedPrimaryKey)"
            $resultLabel.ForeColor = [System.Drawing.Color]::Green
            $captureForm.DialogResult = 'OK'
            $captureForm.Close()
        } else {
            $resultLabel.Text = (Get-KeyCombinationString $currentModifiersTemp $null) 
            $resultLabel.ForeColor = [System.Drawing.Color]::White
        }
    })
    
    $captureForm.Add_KeyUp({
        param($form, $e)
        if ([string]::IsNullOrEmpty($script:capturedPrimaryKey)) {
            [System.Collections.ArrayList]$currentModifiersOnUp = @()
            if ($e.Control) { $currentModifiersOnUp.Add('Ctrl') }
            if ($e.Alt) { $currentModifiersOnUp.Add('Alt') }
            if ($e.Shift) { $currentModifiersOnUp.Add('Shift') }
            
            $resultLabel.Text = (Get-KeyCombinationString $currentModifiersOnUp $null)
            $resultLabel.ForeColor = [System.Drawing.Color]::Yellow
        }
    })


	$captureForm.KeyPreview = $true
	$captureForm.TopMost = $true
	
	$result = $captureForm.ShowDialog()
	try {
		if ($result -eq 'OK' -and -not [string]::IsNullOrEmpty($script:capturedPrimaryKey)) 
		{
			return (Get-KeyCombinationString $script:capturedModifierKeys $script:capturedPrimaryKey)
		}
		else
		{
			return $currentKey  
		}
	} finally {
		try {
			[HotkeyManager]::AreHotkeysGloballyPaused = $false
			if ($script:KeyCapture_PausedSnapshot -and $script:KeyCapture_PausedSnapshot.Count -gt 0) {
				Resume-PausedKeys -Keys $script:KeyCapture_PausedSnapshot
			} else {
				Resume-AllHotkeys
			}
		} catch {
			Write-Verbose ("FTOOL: Resume-AllHotkeys (or Resume-PausedKeys) failed: {0}" -f $_.Exception.Message)
		} finally {
			Remove-Variable -Name KeyCapture_PausedSnapshot -Scope Script -ErrorAction SilentlyContinue
            $script:capturedPrimaryKey = $null
            $script:capturedModifierKeys = @()
		}
	}
}

function LoadFtoolSettings
{
	param($formData)
	
	if (-not $global:DashboardConfig.Config)
	{ 
		$global:DashboardConfig.Config = [ordered]@{} 
	}
	if (-not $global:DashboardConfig.Config.Contains('Ftool'))
	{ 
		$global:DashboardConfig.Config['Ftool'] = [ordered]@{} 
	}
	
	$profilePrefix = FindOrCreateProfile $formData.WindowTitle
	$formData.ProfilePrefix = $profilePrefix

	if ($profilePrefix)
	{
        $globalHotkeyName = "GlobalHotkey_$profilePrefix"
        if ($global:DashboardConfig.Config['Ftool'].Contains($globalHotkeyName))
        {
            $formData.GlobalHotkey = $global:DashboardConfig.Config['Ftool'][$globalHotkeyName]
        }
        
        $hotkeyName = "Hotkey_$profilePrefix"
        if ($global:DashboardConfig.Config['Ftool'].Contains($hotkeyName))
        {
            $formData.Hotkey = $global:DashboardConfig.Config['Ftool'][$hotkeyName]
            $formData.BtnHotKey.Text = $formData.Hotkey 
        }
        else
        {
            $formData.Hotkey = $null 
            $formData.BtnHotKey.Text = "Hotkey" 
        }

		$hotkeysEnabledName = "hotkeys_enabled_$profilePrefix"
		if ($global:DashboardConfig.Config['Ftool'].Contains($hotkeysEnabledName)) {
			$value = $global:DashboardConfig.Config['Ftool'][$hotkeysEnabledName]
			try { $formData.BtnHotkeyToggle.Checked = [bool]::Parse($value) } catch { $formData.BtnHotkeyToggle.Checked = $true }
		} else {
			$formData.BtnHotkeyToggle.Checked = $true
		}

		$keyName = "key1_$profilePrefix"
		$intervalName = "inpt1_$profilePrefix"
		$nameName = "name1_$profilePrefix"
		$positionXName = "pos1X_$profilePrefix"
		$positionYName = "pos1Y_$profilePrefix"
	
		if ($global:DashboardConfig.Config['Ftool'].Contains($keyName))
		{
			$formData.BtnKeySelect.Text = $global:DashboardConfig.Config['Ftool'][$keyName]
		}
		else
		{
			$formData.BtnKeySelect.Text = 'none'
		}
		
		if ($global:DashboardConfig.Config['Ftool'].Contains($intervalName))
		{
			$intervalValue = [int]$global:DashboardConfig.Config['Ftool'][$intervalName]
			if ($intervalValue -lt 10)
			{
				$intervalValue = 100
			}
			$formData.Interval.Text = $intervalValue.ToString()
		}
		else
		{
			$formData.Interval.Text = '100'
		}

		if ($global:DashboardConfig.Config['Ftool'].Contains($nameName))
		{
			$nameValue = [string]$global:DashboardConfig.Config['Ftool'][$nameName]
			$formData.Name.Text = $nameValue.ToString()
		}
		else
		{
			$formData.Name.Text = 'Main'
		}

		if ($global:DashboardConfig.Config['Ftool'].Contains($positionXName))
		{
			$positionValue = [int]$global:DashboardConfig.Config['Ftool'][$positionXName]
			$formData.PositionSliderX.Value = $positionValue
		}
		else
		{
			$formData.PositionSliderX.Value = 0
		}

		if ($global:DashboardConfig.Config['Ftool'].Contains($positionYName))
		{
			$positionValue = [int]$global:DashboardConfig.Config['Ftool'][$positionYName]
			$formData.PositionSliderY.Value = $positionValue
		}
		else
		{
			$formData.PositionSliderY.Value = 100
		}
	}
	else
	{
        $formData.Hotkey = $null
        $formData.BtnHotKey.Text = "Hotkey" 
		$formData.BtnKeySelect.Text = 'none'
		$formData.Interval.Text = '1000'
		$formData.Name.Text = 'Main'
	}
}

function FindOrCreateProfile
{
	param($windowTitle)
	
	$profilePrefix = $null
	$profileFound = $false
	
	if ($windowTitle)
	{
		foreach ($key in $global:DashboardConfig.Config['Ftool'].Keys)
		{
			if ($key -like 'profile_*' -and $global:DashboardConfig.Config['Ftool'][$key] -eq $windowTitle)
			{
				$profilePrefix = $key -replace 'profile_', ''
				$profileFound = $true
				break
			}
		}
		
		if (-not $profileFound)
		{
			$maxProfileNum = 0
			foreach ($key in $global:DashboardConfig.Config['Ftool'].Keys)
			{
				if ($key -match 'profile_(\d+)')
				{
					$profileNum = [int]$matches[1]
					if ($profileNum -gt $maxProfileNum)
					{
						$maxProfileNum = $profileNum
					}
				}
			}
			
			$profilePrefix = ($maxProfileNum + 1).ToString()
			$profileKey = "profile_$profilePrefix"
			$global:DashboardConfig.Config['Ftool'][$profileKey] = $windowTitle
		}
	}
	return $profilePrefix
}

function InitializeExtensionTracking
{
	param($instanceId)
	
	$instanceKey = "instance_$instanceId"
	
	if (-not $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
	{
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey] = @{
			NextExtNum       = 2
			ActiveExtensions = @()
			RemovedExtNums   = @()
		}
	}
	
	$validActiveExtensions = @()
	foreach ($key in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions)
	{
		if ($global:DashboardConfig.Resources.ExtensionData.Contains($key))
		{
			$validActiveExtensions += $key
		}
	}
	$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions = $validActiveExtensions
	
	if ($global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
	{
		$validRemovedNums = @()
		foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
		{
			$intNum = 0
			if ([int]::TryParse($num.ToString(), [ref]$intNum))
			{
				# Only add if not already in the list
				if ($validRemovedNums -notcontains $intNum)
				{
					$validRemovedNums += $intNum
				}
			}
		}
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums = $validRemovedNums
	}
}

function GetNextExtensionNumber
{
	param($instanceId)
	
	$instanceKey = "instance_$instanceId"
	
	if ($global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums.Count -gt 0)
	{
		$sortedNums = @()
		foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
		{
			$intNum = 0
			if ([int]::TryParse($num.ToString(), [ref]$intNum))
			{
				$sortedNums += $intNum
			}
		}
		
		$sortedNums = $sortedNums | Sort-Object
		
		if ($sortedNums.Count -gt 0)
		{
			$extNum = $sortedNums[0]
			
			$newRemovedNums = @()
			foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
			{
				$intNum = 0
				if ([int]::TryParse($num.ToString(), [ref]$intNum) -and $intNum -ne $extNum)
				{
					$newRemovedNums += $intNum
				}
			}
			$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums = $newRemovedNums
			
			Write-Verbose "FTOOL: Reusing extension number $extNum for $instanceId" -ForegroundColor DarkGray
		}
		else
		{
			$extNum = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum
			$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum++
			Write-Verbose "FTOOL: Using new extension number $extNum for $instanceId (no valid reusable numbers)" -ForegroundColor Yellow
		}
	}
	else
	{
		$extNum = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].NextExtNum++
		Write-Verbose "FTOOL: Using new extension number $extNum for $instanceId" -ForegroundColor DarkGray
	}
	
	return $extNum
}

function FindExtensionKeyByControl
{
	param($control, $controlType)
	
	if (-not $control)
	{
		return $null 
	}
	if (-not $controlType -or -not ($controlType -is [string]))
	{
		return $null 
	}
	
	$controlId = $control.GetHashCode()
	$form = $control.FindForm()
	
	if ($form -and $form.Tag -and $form.Tag.ControlToExtensionMap.Contains($controlId))
	{
		$extKey = $form.Tag.ControlToExtensionMap[$controlId]
		
		if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
		{
			return $extKey
		}
		else
		{
			$form.Tag.ControlToExtensionMap.Remove($controlId)
		}
	}
	
	foreach ($key in $global:DashboardConfig.Resources.ExtensionData.Keys)
	{
		$extData = $global:DashboardConfig.Resources.ExtensionData[$key]
		if ($extData -and $extData.$controlType -eq $control)
		{
			if ($form -and $form.Tag)
			{
				$form.Tag.ControlToExtensionMap[$controlId] = $key
			}
			return $key
		}
	}
	
	return $null
}

function LoadExtensionSettings
{
	param($extData, $profilePrefix)
	
	$extNum = $extData.ExtNum
	
	# Load settings from config if they exist
	$keyName = "key${extNum}_$profilePrefix"
	$intervalName = "inpt${extNum}_$profilePrefix"
	$nameName = "name${extNum}_$profilePrefix"
    $extHotkeyName = "ExtHotkey_${extNum}_$profilePrefix" 
	
    if ($global:DashboardConfig.Config['Ftool'].Contains($extHotkeyName))
    {
        $extData.Hotkey = $global:DashboardConfig.Config['Ftool'][$extHotkeyName]
        $extData.BtnHotKey.Text = $extData.Hotkey
    }
    else
    {
        $extData.Hotkey = $null
        $extData.BtnHotKey.Text = "Hotkey" 
    }

	if ($global:DashboardConfig.Config['Ftool'].Contains($keyName))
	{
		$keyValue = $global:DashboardConfig.Config['Ftool'][$keyName]
		$extData.BtnKeySelect.Text = $keyValue
	}
	else
	{
		$extData.BtnKeySelect.Text = 'none'
	}
	
	if ($global:DashboardConfig.Config['Ftool'].Contains($intervalName))
	{
		$intervalValue = [int]$global:DashboardConfig.Config['Ftool'][$intervalName]
		if ($intervalValue -lt 10)
		{
			$intervalValue = 100
		}
		$extData.Interval.Text = $intervalValue.ToString()
	}
	else
	{
		$extData.Interval.Text = '1000'
	}

	if ($global:DashboardConfig.Config['Ftool'].Contains($nameName))
	{
		$nameValue = [string]$global:DashboardConfig.Config['Ftool'][$nameName]

		$extData.Name.Text = $nameValue.ToString()
	}
	else
	{
		$extData.Name.Text = 'Name'
	}
}

function UpdateSettings
{
	param($formData, $extData = $null, [switch]$forceWrite)
	
	$profilePrefix = FindOrCreateProfile $formData.WindowTitle
	
	if ($profilePrefix)
	{
		if ($extData)
		{
			$extNum = $extData.ExtNum
			$keyName = "key${extNum}_$profilePrefix"
			$intervalName = "inpt${extNum}_$profilePrefix"
			$nameName = "name${extNum}_$profilePrefix"
            $extHotkeyName = "ExtHotkey_${extNum}_$profilePrefix" 
            
            if ($extData.Hotkey) { 
                $global:DashboardConfig.Config['Ftool'][$extHotkeyName] = $extData.Hotkey
            } else {
                if ($global:DashboardConfig.Config['Ftool'].Contains($extHotkeyName)) {
                    $global:DashboardConfig.Config['Ftool'].Remove($extHotkeyName)
                }
            }
			
			if ($extData.BtnKeySelect -and $extData.BtnKeySelect.Text)
			{
				$global:DashboardConfig.Config['Ftool'][$keyName] = $extData.BtnKeySelect.Text
			}
			
			if ($extData.Interval)
			{
				$global:DashboardConfig.Config['Ftool'][$intervalName] = $extData.Interval.Text
			}

			if ($extData.Name)
			{
				$global:DashboardConfig.Config['Ftool'][$nameName] = $extData.Name.Text
			}
		}
		else
		{
            $global:DashboardConfig.Config['Ftool']["GlobalHotkey_$profilePrefix"] = $formData.GlobalHotkey
            $global:DashboardConfig.Config['Ftool']["Hotkey_$profilePrefix"] = $formData.Hotkey
			$global:DashboardConfig.Config['Ftool']["hotkeys_enabled_$profilePrefix"] = $formData.BtnHotkeyToggle.Checked
			$global:DashboardConfig.Config['Ftool']["key1_$profilePrefix"] = $formData.BtnKeySelect.Text
			$global:DashboardConfig.Config['Ftool']["inpt1_$profilePrefix"] = $formData.Interval.Text
			$global:DashboardConfig.Config['Ftool']["name1_$profilePrefix"] = $formData.Name.Text
			$global:DashboardConfig.Config['Ftool']["pos1X_$profilePrefix"] = $formData.PositionSliderX.Value
			$global:DashboardConfig.Config['Ftool']["pos1Y_$profilePrefix"] = $formData.PositionSliderY.Value
		}
		
		if ($forceWrite) {
			Write-Config
			if ($global:DashboardConfig.Resources.Timers.Contains('ConfigWriteTimer')) {
				$global:DashboardConfig.Resources.Timers['ConfigWriteTimer'].Stop()
			}
		}
		else
		{
            if (-not $global:DashboardConfig.Resources.Timers.Contains('ConfigWriteTimer'))
            {
                $ConfigWriteTimer = New-Object System.Windows.Forms.Timer
                $ConfigWriteTimer.Interval = 1000
                $ConfigWriteTimer.Add_Tick({
                        param($timerSender, $timerArgs)
                        try
                        {
                            if ($timerSender)
                            {
                                $timerSender.Stop()
                            }
                            Write-Config
                        }
                        catch
                        {
                            Write-Verbose ("FTOOL: Error in config write timer: {0}" -f $_.Exception.Message) -ForegroundColor Red
                        }
                    })
                $global:DashboardConfig.Resources.Timers['ConfigWriteTimer'] = $ConfigWriteTimer
            }
            else
            {
                $ConfigWriteTimer = $global:DashboardConfig.Resources.Timers['ConfigWriteTimer']
            }
            
            $ConfigWriteTimer.Stop()
            $ConfigWriteTimer.Start()
        }
	}
}

function CreatePositionTimer
{
	param($formData)
	
	$positionTimer = New-Object System.Windows.Forms.Timer
	$positionTimer.Interval = 10
	$positionTimer.Tag = @{
		WindowHandle = $formData.SelectedWindow
		FtoolForm    = $formData.Form
		InstanceId   = $formData.InstanceId
		FormData     = $formData
	}
	
	$positionTimer.Add_Tick({
			param($s, $e)
			try
			{
				if (-not $s -or -not $s.Tag)
				{
					return
				}
				$timerData = $s.Tag
				if (-not $timerData -or $timerData['WindowHandle'] -eq [IntPtr]::Zero)
				{
					return
				}
			
				if (-not $timerData['FtoolForm'] -or $timerData['FtoolForm'].IsDisposed)
				{
					return
				}
		
				$rect = New-Object Custom.Native+RECT
				if ([Custom.Native]::GetWindowRect($timerData['WindowHandle'], [ref]$rect))
				{
					try
					{
						$sliderValueX = $timerData.FormData.PositionSliderX.Value
						$maxLeft = $rect.Right - $timerData['FtoolForm'].Width - 8
						$targetLeft = $rect.Left + 8 + (($maxLeft - ($rect.Left + 8)) * $sliderValueX / 100)
						
						$sliderValueY = 100 - $timerData.FormData.PositionSliderY.Value
						$maxTop = $rect.Bottom - $timerData['FtoolForm'].Height - 8
						$targetTop = $rect.Top + 8 + (($maxTop - ($rect.Top + 8)) * $sliderValueY / 100)

						$currentLeft = $timerData['FtoolForm'].Left
						$newLeft = $currentLeft + ($targetLeft - $currentLeft) * 0.2 
						$timerData['FtoolForm'].Left = [int]$newLeft

						$currentTop = $timerData['FtoolForm'].Top
						$newTop = $currentTop + ($targetTop - $currentTop) * 0.2 
						$timerData['FtoolForm'].Top = [int]$newTop
					
						$foregroundWindow = [Custom.Native]::GetForegroundWindow()
					
						if ($foregroundWindow -eq $timerData['WindowHandle'])
						{
							if (-not $timerData['FtoolForm'].TopMost)
							{
								$timerData['FtoolForm'].TopMost = $true
								$timerData['FtoolForm'].BringToFront()
							}
						}
						elseif ($foregroundWindow -ne $timerData['FtoolForm'].Handle)
						{
							if ($timerData['FtoolForm'].TopMost)
							{
								$timerData['FtoolForm'].TopMost = $false
							}
						}
					}
					catch
					{
						Write-Verbose ("FTOOL: Position timer error: {0} for {1}" -f $_.Exception.Message, $($timerData['InstanceId'])) -ForegroundColor Red
					}
				}
			}
			catch
			{
				Write-Verbose ("FTOOL: Position timer critical error: {0}" -f $_.Exception.Message) -ForegroundColor Red
			}
		})
	
	$positionTimer.Start()
	$global:DashboardConfig.Resources.Timers["ftoolPosition_$($formData.InstanceId)"] = $positionTimer
}

function RepositionExtensions
{
	param($form, $instanceId)
	
	if (-not $form -or -not $instanceId)
	{
		return 
	}
	
	$instanceKey = "instance_$instanceId"
	
	if (-not $global:DashboardConfig.Resources.ExtensionTracking -or 
		-not $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
	{
		return
	}
	
	$form.SuspendLayout()
	
	try
	{
		$baseHeight = 130
		
		$activeExtensions = @()
		if ($global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions)
		{
			$activeExtensions = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions
		}
		
		$newHeight = 130  
		$position = 0
		
		if ($activeExtensions.Count -gt 0)
		{
			$sortedExtensions = @()
			foreach ($extKey in $activeExtensions)
			{
				if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
				{
					$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
					$sortedExtensions += [PSCustomObject]@{
						Key    = $extKey
						ExtNum = [int]$extData.ExtNum 
					}
				}
			}
			$sortedExtensions = $sortedExtensions | Sort-Object ExtNum
			
			foreach ($extObj in $sortedExtensions)
			{
				$extKey = $extObj.Key
				if ($global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
				{
					$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
					if ($extData -and $extData.Panel -and -not $extData.Panel.IsDisposed)
					{
						$newTop = $baseHeight + ($position * 70)
						
						$extData.Panel.Top = $newTop
						$extData.Position = $position
						
						$position++
					}
				}
			}
			
			if ($position -gt 0)
			{
				$newHeight = 130 + ($position * 70)
			}
		}
		
		if (-not $form.Tag.IsCollapsed)
		{
			$form.Height = $newHeight
			$form.Tag.OriginalHeight = $newHeight
			$form.Tag.PositionSliderY.Height = $newHeight - 20
		}
	}
	catch
	{
		Write-Verbose ("FTOOL: Error in RepositionExtensions: {0}" -f $_.Exception.Message) -ForegroundColor Red
	}
	finally
	{
		$form.ResumeLayout()
	}
}

function CreateSpammerTimer
{
	param($windowHandle, $keyValue, $instanceId, $interval, $extNum = $null, $extKey = $null)
	
	$spamTimer = New-Object System.Windows.Forms.Timer
	$spamTimer.Interval = $interval
	
	$timerTag = @{
		WindowHandle = $windowHandle
		Key          = $keyValue
		InstanceId   = $instanceId
	}
	
	if ($null -ne $extNum)
	{
		$timerTag['ExtNum'] = $extNum
		$timerTag['ExtKey'] = $extKey
	}
	
	$spamTimer.Tag = $timerTag
	
	$spamTimer.Add_Tick({
			param($s, $evt)
			try
			{
				if (-not $s -or -not $s.Tag)
				{
					return
				}
				$timerData = $s.Tag
				if (-not $timerData -or $timerData['WindowHandle'] -eq [IntPtr]::Zero)
				{
					return
				}
				
				$keyMappings = Get-VirtualKeyMappings
				$virtualKeyCode = $keyMappings[$timerData['Key']]
				
				if ($virtualKeyCode)
				{
					[Custom.Ftool]::fnPostMessage($timerData['WindowHandle'], 256, $virtualKeyCode, 0)
				}
				else
				{
					Write-Verbose "FTOOL: Unknown key '$($timerData['Key'])' for $($timerData['InstanceId'])" -ForegroundColor Yellow
				}
			}
			catch
			{
				Write-Verbose ("FTOOL: Spammer timer error: {0}" -f $_.Exception.Message) -ForegroundColor Red
			}
		})
	
	$spamTimer.Start()
	return $spamTimer
}

function ToggleButtonState
{
	param($startBtn, $stopBtn, $isStarting)
	
	if ($isStarting)
	{
		$startBtn.Enabled = $false
		$startBtn.Visible = $false
		$stopBtn.Enabled = $true
		$stopBtn.Visible = $true
	}
	else
	{
		$startBtn.Enabled = $true
		$startBtn.Visible = $true
		$stopBtn.Enabled = $false
		$stopBtn.Visible = $false
	}
}

function CheckRateLimit
{
	param($controlId, $minInterval = 100)
	
	$currentTime = Get-Date
	if ($global:DashboardConfig.Resources.LastEventTimes.Contains($controlId) -and 
	($currentTime - $global:DashboardConfig.Resources.LastEventTimes[$controlId]).TotalMilliseconds -lt $minInterval)
	{
		return $false
	}
	
	$global:DashboardConfig.Resources.LastEventTimes[$controlId] = $currentTime
	return $true
}

function AddFormCleanupHandler
{
	param($form)
	
	if (-not $form.Tag.ExtensionCleanupAdded)
	{
		$form.Add_FormClosing({
				param($src, $e)
			
				try
				{
					$instanceId = $src.Tag.InstanceId
					if ($instanceId)
					{
						CleanupInstanceResources $instanceId
					}
				}
				catch
				{
					Write-Verbose ("FTOOL: Error during form cleanup: {0}" -f $_.Exception.Message) -ForegroundColor Red
				}
			})
		
		$form.Tag | Add-Member -NotePropertyName ExtensionCleanupAdded -NotePropertyValue $true -Force
	}
}

function CleanupInstanceResources
{
	param($instanceId)
	
	$instanceKey = "instance_$instanceId"
	
	$keysToRemove = @()
	
	$activeExtensions = @()
	if ($global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey) -and 
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions)
	{
		$activeExtensions = @() + $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions
	}
	
	foreach ($key in $activeExtensions)
	{
		if ($global:DashboardConfig.Resources.ExtensionData.Contains($key))
		{
			$extData = $global:DashboardConfig.Resources.ExtensionData[$key]
			if ($extData.RunningSpammer)
			{
				$extData.RunningSpammer.Stop()
				$extData.RunningSpammer.Dispose()
			}
			$keysToRemove += $key
		}
	}
	
	foreach ($key in $keysToRemove)
	{
		$global:DashboardConfig.Resources.ExtensionData.Remove($key)
	}
	
	$timerKeysToRemove = @()
	foreach ($key in $global:DashboardConfig.Resources.Timers.Keys)
	{
		if ($key -like "ExtSpammer_${instanceId}_*" -or $key -eq "ftoolPosition_$instanceId")
		{
			$timer = $global:DashboardConfig.Resources.Timers[$key]
			if ($timer)
			{
				$timer.Stop()
				$timer.Dispose()
			}
			$timerKeysToRemove += $key
		}
	}
	
	foreach ($key in $timerKeysToRemove)
	{
		$global:DashboardConfig.Resources.Timers.Remove($key)
	}
	
    if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId)) {
        $form = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
        if ($form -and $form.Tag -and $form.Tag.HotkeyId) {
			try {
				Unregister-HotkeyInstance -Id $form.Tag.HotkeyId -OwnerKey $form.Tag.InstanceId
			} catch {
				Write-Warning "FTOOL: Failed to unregister hotkey with ID $($form.Tag.HotkeyId) for instance $instanceId. Error: $_"
			}
        }
    }

	if ($global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
	{
		$global:DashboardConfig.Resources.ExtensionTracking.Remove($instanceKey)
	}
	
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
}

function Stop-FtoolForm
{
	param($Form)
	
	if (-not $Form -or $Form.IsDisposed)
	{
		return
	}
	
	try
	{
		$instanceId = $Form.Tag.InstanceId
		if ($instanceId)
		{
			CleanupInstanceResources $instanceId
			
			if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
			{
				$global:DashboardConfig.Resources.FtoolForms.Remove($instanceId)
			}
		}
		
		$Form.Close()
		$Form.Dispose()
	}
	catch
	{
		Write-Verbose ("FTOOL: Error stopping Ftool form: {0}" -f $_.Exception.Message) -ForegroundColor Red
	}
}

function RemoveExtension
{
	param($form, $extKey)
	
	if (-not $form -or -not $extKey)
	{
		return $false 
	}
	
	if (-not $global:DashboardConfig.Resources.ExtensionData -or 
		-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
	{
		return $false
	}
	
	$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
	if (-not $extData)
	{
		return $false 
	}
	
	$extNum = [int]$extData.ExtNum 
	$instanceId = $extData.InstanceId
	$instanceKey = "instance_$instanceId"
	
	try
	{
		if ($extData.RunningSpammer)
		{
			$extData.RunningSpammer.Stop()
			$extData.RunningSpammer.Dispose()
			$extData.RunningSpammer = $null
			
			$timerKey = "ExtSpammer_${instanceId}_$extNum"
			if ($global:DashboardConfig.Resources.Timers.Contains($timerKey))
			{
				$global:DashboardConfig.Resources.Timers.Remove($timerKey)
			}
		}
		
		if ($extData.Panel -and -not $extData.Panel.IsDisposed)
		{
			$form.Controls.Remove($extData.Panel)
			$extData.Panel.Dispose()
		}
		
		if (-not $global:DashboardConfig.Resources.ExtensionTracking -or 
			-not $global:DashboardConfig.Resources.ExtensionTracking.Contains($instanceKey))
		{
			InitializeExtensionTracking $instanceId
		}
		
		$intExtNum = [int]$extNum 
		
		$found = $false
		foreach ($num in $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums)
		{
			$intNum = 0
			if ([int]::TryParse($num.ToString(), [ref]$intNum) -and $intNum -eq $intExtNum)
			{
				$found = $true
				break
			}
		}
		
		if (-not $found)
		{
			$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].RemovedExtNums += $intExtNum
		}
		
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions = 
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions | 
		Where-Object { $_ -ne $extKey }
		
		$global:DashboardConfig.Resources.ExtensionData.Remove($extKey)

        if ($extData.HotkeyId) {
			try {
				Unregister-HotkeyInstance -Id $extData.HotkeyId -OwnerKey $extKey
			} catch {
				Write-Warning "FTOOL: Failed to unregister hotkey ID $($extData.HotkeyId) for extension $($extKey). Error: $_"
			}
        }
		
		RepositionExtensions $form $instanceId
		
		return $true
	}
	catch
	{
		Write-Verbose ("FTOOL: Error in RemoveExtension: {0}" -f $_.Exception.Message) -ForegroundColor Red
		return $false
	}
}

#endregion

#region Core Functions

function FtoolSelectedRow
{
	param($row)
	
	if (-not $row -or -not $row.Cells -or $row.Cells.Count -lt 3)
	{
		Write-Verbose 'FTOOL: Invalid row data, skipping' -ForegroundColor Yellow
		return
	}
	
	$instanceId = $row.Cells[2].Value.ToString()
	if (-not $row.Tag -or -not $row.Tag.MainWindowHandle)
	{
		Write-Verbose "FTOOL: Missing window handle for instance $instanceId, skipping" -ForegroundColor Yellow
		return
	}
	
	$windowHandle = $row.Tag.MainWindowHandle
	if (-not $instanceId)
	{
		Write-Verbose 'FTOOL: Missing instance ID, skipping' -ForegroundColor Yellow
		return
	}
	
	if ($global:DashboardConfig.Resources.FtoolForms.Contains($instanceId))
	{
		$existingForm = $global:DashboardConfig.Resources.FtoolForms[$instanceId]
		if (-not $existingForm.IsDisposed)
		{
			$existingForm.BringToFront()
			return
		}
		else
		{
			$global:DashboardConfig.Resources.FtoolForms.Remove($instanceId)
		}
	}
	
		$targetWindowRect = New-Object Custom.Native+RECT
		[Custom.Native]::GetWindowRect($windowHandle, [ref]$targetWindowRect)
	
	$windowTitle = if ($row.Tag -and $row.Tag.MainWindowTitle)
	{
		$row.Tag.MainWindowTitle
	}
	elseif ($row.Cells[1].Value)
	{
		$row.Cells[1].Value.ToString()
	}
	else
	{
		"Window_$instanceId"
	}
	
	$ftoolForm = CreateFtoolForm $instanceId $targetWindowRect $windowTitle $row
	
	$global:DashboardConfig.Resources.FtoolForms[$instanceId] = $ftoolForm
	
    $ftoolForm.Show()
	$ftoolForm.BringToFront()
}

function CreateFtoolForm
{
	param($instanceId, $targetWindowRect, $windowTitle, $row)
	
	$ftoolForm = Set-UIElement -type 'Form' -visible $true -width 250 -height 170 -top ($targetWindowRect.Top + 30) -left ($targetWindowRect.Left + 10) -bg @(30, 30, 30) -fg @(255, 255, 255) -text "FTool - $instanceId" -startPosition 'Manual' -formBorderStyle 0 -opacity 1
	if ($global:DashboardConfig.Paths.Icon -and (Test-Path $global:DashboardConfig.Paths.Icon))
	{
		try
		{
			$icon = New-Object System.Drawing.Icon($global:DashboardConfig.Paths.Icon)
			$ftoolForm.Icon = $icon
		}
		catch
		{
			Write-Verbose "FTOOL: Failed to load icon from $($global:DashboardConfig.Paths.Icon): $_" -ForegroundColor Red
		}
	}
	if (-not $ftoolForm)
	{
		Write-Verbose "FTOOL: Failed to create form for $instanceId, skipping" -ForegroundColor Red
		return $null
	}
	
	$headerPanel = Set-UIElement -type 'Panel' -visible $true -width 250 -height 20 -top 0 -left 0 -bg @(40, 40, 40)
	$ftoolForm.Controls.Add($headerPanel)

	$btnInstanceHotkeyToggle = Set-UIElement -type 'Label' -visible $true -width 15 -height 14 -top 2 -left 118 -bg @(40, 40, 40, 0) -fg @(255, 255, 255) -text ([char]0x2328) -font (New-Object System.Drawing.Font('Segoe UI', 10))
	$headerPanel.Controls.Add($btnInstanceHotkeyToggle)

	$labelWinTitle = Set-UIElement -type 'Label' -visible $true -width 120 -height 20 -top 5 -left 5 -bg @(40, 40, 40, 0) -fg @(255, 255, 255) -text $row.Cells[1].Value -font (New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Regular))
	$headerPanel.Controls.Add($labelWinTitle)
	
	$btnHotkeyToggle = Set-UIElement -type 'Toggle' -visible $true -width 30 -height 14 -top 3 -left 135 -bg @(40, 80, 80) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8)) -checked $true
	$headerPanel.Controls.Add($btnHotkeyToggle)

	$btnAdd = Set-UIElement -type 'Button' -visible $true -width 14 -height 14 -top 3 -left 170 -bg @(40, 80, 80) -fg @(255, 255, 255) -text ([char]0x2795) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11))
	$headerPanel.Controls.Add($btnAdd)
	
	$btnShowHide = Set-UIElement -type 'Button' -visible $true -width 14 -height 14 -top 3 -left 185 -bg @(60, 60, 100) -fg @(255, 255, 255) -text ([char]0x25B2) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11))
	$headerPanel.Controls.Add($btnShowHide)
	
	$btnClose = Set-UIElement -type 'Button' -visible $true -width 14 -height 14 -top 3 -left 210 -bg @(200, 45, 45) -fg @(255, 255, 255) -text ([char]0x166D) -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 11))
	$headerPanel.Controls.Add($btnClose)
	
	$panelSettings = Set-UIElement -type 'Panel' -visible $true -width 190 -height 60 -top 60 -left 40 -bg @(50, 50, 50)
	$ftoolForm.Controls.Add($panelSettings)
	
	$btnKeySelect = Set-UIElement -type 'Button' -visible $true -width 55 -height 25 -top 4 -left 3 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8))
	$panelSettings.Controls.Add($btnKeySelect)
	
	$interval = Set-UIElement -type 'TextBox' -visible $true -width 47 -height 15 -top 5 -left 59 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '1000' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$panelSettings.Controls.Add($interval)
	
	$name = Set-UIElement -type 'TextBox' -visible $true -width 37 -height 17 -top 5 -left 108 -bg @(40, 40, 40) -fg @(255, 255, 255) -text 'Main' -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular))
	$panelSettings.Controls.Add($name)
	
	$btnStart = Set-UIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 10 -bg @(0, 120, 215) -fg @(255, 255, 255) -text 'Start' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$panelSettings.Controls.Add($btnStart)
	
	$btnStop = Set-UIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 67 -bg @(200, 50, 50) -fg @(255, 255, 255) -text 'Stop' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$btnStop.Enabled = $false
	$btnStop.Visible = $false
	$panelSettings.Controls.Add($btnStop)

	$btnHotKey = Set-UIElement -type 'Button' -visible $true -width 40 -height 30 -top 1 -left 146 -bg @(200, 50, 50) -fg @(255, 255, 255) -text 'Hotkey' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 6))
	$panelSettings.Controls.Add($btnHotKey)
	
	$positionSliderY = New-Object System.Windows.Forms.TrackBar
	$positionSliderY.Orientation = 'Vertical'
	$positionSliderY.Minimum = 0
	$positionSliderY.Maximum = 100
	$positionSliderY.TickFrequency = 300
	$positionSliderY.Value = 0
	$positionSliderY.Size = New-Object System.Drawing.Size(1, 110)
	$positionSliderY.Location = New-Object System.Drawing.Point(5, 20)
	$ftoolForm.Controls.Add($positionSliderY)
		
	$positionSliderX = New-Object System.Windows.Forms.TrackBar
	$positionSliderX.Minimum = 0
	$positionSliderX.Maximum = 100
	$positionSliderX.TickFrequency = 300
	$positionSliderX.Value = 0
	$positionSliderX.Size = New-Object System.Drawing.Size(190, 1)
	$positionSliderX.Location = New-Object System.Drawing.Point(45, 25)
	$ftoolForm.Controls.Add($positionSliderX)

	$formData = [PSCustomObject]@{
		InstanceId            = $instanceId
		SelectedWindow        = $row.Tag.MainWindowHandle
		BtnKeySelect          = $btnKeySelect
		Interval              = $interval
		Name                  = $name
		BtnStart              = $btnStart
		BtnStop               = $btnStop
		BtnHotKey             = $btnHotKey
		BtnInstanceHotkeyToggle = $btnInstanceHotkeyToggle
		BtnHotkeyToggle       = $btnHotkeyToggle
		BtnAdd                = $btnAdd
		BtnClose              = $btnClose
		BtnShowHide           = $btnShowHide
		PositionSliderX       = $positionSliderX
		PositionSliderY       = $positionSliderY
		Form                  = $ftoolForm
		RunningSpammer        = $null
		WindowTitle           = $windowTitle
		Process               = $row.Tag
		OriginalLeft          = $targetWindowRect.Left
		IsCollapsed           = $false
		LastExtensionAdded    = 0
		ExtensionCount        = 0
		ControlToExtensionMap = @{}
		OriginalHeight        = 130
		HotkeyId              = $null 
		GlobalHotkeyId        = $null 
		ProfilePrefix         = $null 
		Hotkey                = $null 
		GlobalHotkey          = $null 
	}
	
	$ftoolForm.Tag = $formData
	
	LoadFtoolSettings $formData
	ToggleInstanceHotkeys -InstanceId $formData.InstanceId -ToggleState $formData.BtnHotkeyToggle.Checked
	
	if (-not [string]::IsNullOrEmpty($formData.Hotkey)) {
		try {
			$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($formData.InstanceId)' -ExtKey `$null")
			$newHotkeyId = Set-Hotkey -KeyCombinationString $formData.Hotkey -Action $scriptBlock -OwnerKey $formData.InstanceId
			$formData.HotkeyId = $newHotkeyId
			$hotkeyIdDisplay = if ($formData.HotkeyId) { $formData.HotkeyId } else { 'None' }
			Write-Verbose "FTOOL: Registered hotkey $($formData.Hotkey) (ID: $hotkeyIdDisplay) for main instance $($formData.InstanceId) on load."
		} catch {
			Write-Warning "FTOOL: Failed to register hotkey $($formData.Hotkey) for main instance $($formData.InstanceId) on load. Error: $_"
			$formData.HotkeyId = $null 
		}
	}

	if (-not [string]::IsNullOrEmpty($formData.GlobalHotkey)) {
		try {
			$ownerKey = "global_toggle_$($formData.InstanceId)"
			$script = @"
Write-Verbose "FTOOL: Global-toggle hotkey triggered for instance '$($formData.InstanceId)'"
if (`$global:DashboardConfig.Resources.FtoolForms.Contains('$($formData.InstanceId)')) {
	`$f = `$global:DashboardConfig.Resources.FtoolForms['$($formData.InstanceId)']
	if (`$f -and -not `$f.IsDisposed -and `$f.Tag) {
		`$toggle = `$f.Tag.BtnHotkeyToggle
		if (`$toggle) {
			try {
                if (`$toggle.InvokeRequired) {
                    `$action = [System.Action]{ `$toggle.Checked = -not `$toggle.Checked }
                    `$toggle.Invoke(`$action)
                } else {
                    `$toggle.Checked = -not `$toggle.Checked
                }
                ToggleInstanceHotkeys -InstanceId '$($formData.InstanceId)' -ToggleState `$toggle.Checked
                Write-Verbose "FTOOL: Global toggle action completed for instance '$($formData.InstanceId)'."
            } catch {
                Write-Warning "FTOOL: Error during global toggle action for instance '$($formData.InstanceId)'. Error: `$_"
            }
		}
	}
}
"@
			$scriptBlock = [scriptblock]::Create($script)
			$formData.GlobalHotkeyId = Set-Hotkey -KeyCombinationString $formData.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey
			Write-Verbose "FTOOL: Registered global-toggle hotkey $($formData.GlobalHotkey) (ID: $($formData.GlobalHotkeyId)) for instance $($formData.InstanceId) on load."
		} catch {
			Write-Warning "FTOOL: Failed to register global-toggle hotkey $($formData.GlobalHotkey) for instance $($formData.InstanceId) on load. Error: $_"
			$formData.GlobalHotkeyId = $null
		}
	}

	CreatePositionTimer $formData
	AddFtoolEventHandlers $formData
	
	return $ftoolForm
}

function AddFtoolEventHandlers
{
	param($formData)
	
	$formData.Interval.Add_TextChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			$intervalValue = 0
			if (-not [int]::TryParse($this.Text, [ref]$intervalValue) -or $intervalValue -lt 10)
			{
				$this.Text = '100'
			}
		
			UpdateSettings $data -forceWrite
		})

	$formData.Name.Add_TextChanged({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$data = $form.Tag
	
		UpdateSettings $data -forceWrite
	})

	$formData.PositionSliderX.Add_ValueChanged({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return
		}
		$data = $form.Tag
		UpdateSettings $data -forceWrite
	})

	$formData.PositionSliderY.Add_ValueChanged({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return
		}
		$data = $form.Tag
		UpdateSettings $data -forceWrite
	})
	
	$formData.BtnKeySelect.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$data = $form.Tag
		
		$currentKey = $data.BtnKeySelect.Text
		$newKey = Show-KeyCaptureDialog $currentKey
		
		if ($newKey -and $newKey -ne $currentKey)
		{
			if ($data.RunningSpammer)
			{
				$data.RunningSpammer.Stop()
				$data.RunningSpammer.Dispose()
				$data.RunningSpammer = $null
			}
			
			$data.BtnKeySelect.Text = $newKey
			UpdateSettings $data -forceWrite
			
			ToggleButtonState $data.BtnStart $data.BtnStop $false
		}
	})
	
	$formData.BtnStart.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			$keyValue = $data.BtnKeySelect.Text
			if (-not $keyValue -or $keyValue.Trim() -eq '')
			{
			[System.Windows.Forms.MessageBox]::Show('Please select a key', 'Error')
			return
			}
		
			$intervalNum = 0
			if (-not [int]::TryParse($data.Interval.Text, [ref]$intervalNum) -or $intervalNum -lt 10)
			{
				$intervalNum = 100
				$data.Interval.Text = '100'
			}
		
			ToggleButtonState $data.BtnStart $data.BtnStop $true
		
			$spamTimer = CreateSpammerTimer $data.SelectedWindow $keyValue $data.InstanceId $intervalNum
		
			$data.RunningSpammer = $spamTimer
			$global:DashboardConfig.Resources.Timers["ExtSpammer_$($data.InstanceId)_1"] = $spamTimer
		})
	
	$formData.BtnStop.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			if ($data.RunningSpammer)
			{
				$data.RunningSpammer.Stop()
				$data.RunningSpammer.Dispose()
				$data.RunningSpammer = $null
			}
		
			ToggleButtonState $data.BtnStart $data.BtnStop $false
		
			if ($global:DashboardConfig.Resources.Timers.Contains("ExtSpammer_$($data.InstanceId)_1"))
			{
				$global:DashboardConfig.Resources.Timers.Remove("ExtSpammer_$($data.InstanceId)_1")
			}
		})

	$formData.BtnHotKey.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$data = $form.Tag

		$currentHotkeyText = $data.BtnHotKey.Text
		if ($currentHotkeyText -eq "Hotkey") { $currentHotkeyText = $null } 

		$oldHotkeyIdToUnregister = $data.HotkeyId

		$newHotkey = Show-KeyCaptureDialog $currentHotkeyText
		
		        if ($newHotkey -and $newHotkey -ne $currentHotkeyText) { 
		
		            $data.Hotkey = $newHotkey 		
			try {
				$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($data.InstanceId)' -ExtKey `$null")
				$data.HotkeyId = Set-Hotkey -KeyCombinationString $data.Hotkey -Action $scriptBlock -OwnerKey $data.InstanceId -OldHotkeyId $oldHotkeyIdToUnregister
				$data.BtnHotKey.Text = $newHotkey 
				Write-Verbose "FTOOL: Registered hotkey $($data.Hotkey) (ID: $($data.HotkeyId)) for main instance $($data.InstanceId)."
			} catch {
				Write-Warning "FTOOL: Failed to register hotkey $($data.Hotkey) for main instance $($data.InstanceId). Error: $_"
                $data.HotkeyId = $null 
                $data.Hotkey = $currentHotkeyText
                $data.BtnHotKey.Text = $currentHotkeyText -or "Hotkey"
			}
			
			UpdateSettings $data -forceWrite 
		} elseif (-not $newHotkey -and $oldHotkeyIdToUnregister) { 
			try {
				Unregister-HotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $data.InstanceId
				Write-Verbose "FTOOL: Unregistered hotkey (ID: $($oldHotkeyIdToUnregister)) for main instance $($data.InstanceId) due to user clear."
			} catch {
				Write-Warning "FTOOL: Failed to unregister hotkey (ID: $($oldHotkeyIdToUnregister)) for main instance $($data.InstanceId) on clear. Error: $_"
			}
            $data.HotkeyId = $null
            $data.Hotkey = $null
            $data.BtnHotKey.Text = "Hotkey"
            UpdateSettings $data -forceWrite
        }
	})

	$formData.BtnHotkeyToggle.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag) { return }
		$data = $form.Tag
		$toggleOn = $this.Checked
        ToggleInstanceHotkeys -InstanceId $data.InstanceId -ToggleState $toggleOn
	})

			$formData.BtnInstanceHotkeyToggle.Add_Click({
				$form = $this.FindForm()
				if (-not $form -or -not $form.Tag) { return }
				$data = $form.Tag
		
				$currentHotkeyText = $data.GlobalHotkey
		
				$oldHotkeyIdToUnregister = $data.GlobalHotkeyId
				$ownerKey = "global_toggle_$($data.InstanceId)"
		
				$newHotkey = Show-KeyCaptureDialog $currentHotkeyText
		
				if ($newHotkey -and $newHotkey -ne $currentHotkeyText) {
					if (Test-HotkeyConflict -KeyCombinationString $newHotkey -NewHotkeyType ([HotkeyManager+HotkeyActionType]::GlobalToggle) -OwnerKeyToExclude $ownerKey) {
						[System.Windows.Forms.MessageBox]::Show("This hotkey combination ('$newHotkey') is already assigned to another function or instance. Please choose a different key.", "Hotkey Conflict", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
						$data.GlobalHotkey = $currentHotkeyText
						return 
					}
		
					$data.GlobalHotkey = $newHotkey
					try {
						$script = @"
		Write-Verbose "FTOOL: Global-toggle hotkey triggered for instance '$($data.InstanceId)'"
		if (`$global:DashboardConfig.Resources.FtoolForms.Contains('$($data.InstanceId)')) {
			`$f = `$global:DashboardConfig.Resources.FtoolForms['$($data.InstanceId)']
			if (`$f -and -not `$f.IsDisposed -and `$f.Tag) {
				`$toggle = `$f.Tag.BtnHotkeyToggle
				if (`$toggle) {
					try {
						if (`$toggle.InvokeRequired) {
							`$action = [System.Action]{ `$toggle.Checked = -not `$toggle.Checked }
							`$toggle.Invoke(`$action)
						} else {
							`$toggle.Checked = -not `$toggle.Checked
						}
						ToggleInstanceHotkeys -InstanceId '$($data.InstanceId)' -ToggleState `$toggle.Checked
						Write-Verbose "FTOOL: Global toggle action completed for instance '$($data.InstanceId)'."
					} catch {
						Write-Warning "FTOOL: Error during global toggle action for instance '$($data.InstanceId)'. Error: `$_"
					}
				}
			}
		}
"@
						$scriptBlock = [scriptblock]::Create($script)
						$data.GlobalHotkeyId = Set-Hotkey -KeyCombinationString $data.GlobalHotkey -Action $scriptBlock -OwnerKey $ownerKey -OldHotkeyId $oldHotkeyIdToUnregister
						Write-Verbose "FTOOL: Registered global-toggle hotkey $($data.GlobalHotkey) (ID: $($data.GlobalHotkeyId)) for instance $($data.InstanceId)."
					} catch {
						Write-Warning "FTOOL: Failed to register global-toggle hotkey $($data.GlobalHotkey) for instance $($data.InstanceId). Error: $_"
						$data.GlobalHotkeyId = $null
						$data.GlobalHotkey = $currentHotkeyText
					}
		
					UpdateSettings $data -forceWrite
				} elseif (-not $newHotkey -and $oldHotkeyIdToUnregister) {
					try {
						Unregister-HotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $ownerKey
						Write-Verbose "FTOOL: Unregistered global-toggle hotkey (ID: $($oldHotkeyIdToUnregister)) for instance $($data.InstanceId) due to user clear."
					} catch {
						Write-Warning "FTOOL: Failed to unregister global-toggle hotkey (ID: $($oldHotkeyIdToUnregister)) for instance $($data.InstanceId) on clear. Error: $_"
					}
					$data.GlobalHotkeyId = $null
					$data.GlobalHotkey = $null
					UpdateSettings $data -forceWrite
				}
			})

	
	$formData.BtnShowHide.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			if (-not (CheckRateLimit $this.GetHashCode() 200))
			{
				return 
			}
		
			if ($data.IsCollapsed)
			{
				$form.Height = $data.OriginalHeight
				$data.IsCollapsed = $false
				$this.Text = [char]0x25B2  
			}
			else
			{
				$data.OriginalHeight = $form.Height
				$form.Height = 25
				$data.IsCollapsed = $true
				$this.Text = [char]0x25BC  
			}
		})
	
	$formData.BtnAdd.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag -or $form.Tag.IsCollapsed)
			{
				return 
			}
			$data = $form.Tag
		
			if (-not (CheckRateLimit $this.GetHashCode() 200))
			{
				return 
			}
		
			if ($data.ExtensionCount -ge 8)
			{
				[System.Windows.Forms.MessageBox]::Show('Maximum number of extensions reached.', 'Warning')
				return
			}
		
			$data.ExtensionCount++
		
			InitializeExtensionTracking $data.InstanceId
		
			$extNum = GetNextExtensionNumber $data.InstanceId
		
			$extData = CreateExtensionPanel $form $currentHeight $extNum $data.InstanceId $data.SelectedWindow
            $extKeyForScriptBlock = "ext_$($data.InstanceId)_$extNum"
		
			$profilePrefix = FindOrCreateProfile $data.WindowTitle
		
			LoadExtensionSettings $extData $profilePrefix
            
            if (-not [string]::IsNullOrEmpty($extData.Hotkey)) {
				try {
					$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($extData.InstanceId)' -ExtKey '$($extKeyForScriptBlock)'")
					$newHotkeyId = Set-Hotkey -KeyCombinationString $extData.Hotkey -Action $scriptBlock -OwnerKey $extKeyForScriptBlock
					$extData.HotkeyId = $newHotkeyId
					$extHotKeyIdDisplay = if ($extData.HotkeyId) { $extData.HotkeyId } else { 'None' }
					Write-Verbose "FTOOL: Registered hotkey $($extData.Hotkey) (ID: $extHotKeyIdDisplay) for extension $($extKeyForScriptBlock) on load."
				} catch {
                    Write-Warning "FTOOL: Failed to register hotkey $($extData.Hotkey) for extension $($extKeyForScriptBlock) on load. Error: $_"
                    $extData.HotkeyId = $null 
                }
            }
		
			AddExtensionEventHandlers $extData $data
		
			AddFormCleanupHandler $form
		
			RepositionExtensions $form $extData.InstanceId
				
			$form.Refresh()
		})
	
	$formData.BtnClose.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			CleanupInstanceResources $data.InstanceId
		
			if ($global:DashboardConfig.Resources.FtoolForms.Contains($data.InstanceId))
			{
				$global:DashboardConfig.Resources.FtoolForms.Remove($data.InstanceId)
			}
		
			$form.Close()
			$form.Dispose()
		})
	
	$formData.Form.Add_FormClosed({
			param($src, $e)
		
			$instanceId = $src.Tag.InstanceId
			if ($instanceId)
			{
				CleanupInstanceResources $instanceId
			}
		})
}

function CreateExtensionPanel
{
	param($form, $currentHeight, $extNum, $instanceId, $windowHandle)
	
	$panelExt = Set-UIElement -type 'Panel' -visible $true -width 190 -height 60 -top 0 -left 40 -bg @(50, 50, 50)
	$form.Controls.Add($panelExt)
	$panelExt.BringToFront()
	
	$btnKeySelectExt = Set-UIElement -type 'Button' -visible $true -width 55 -height 25 -top 4 -left 3 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 8))
	$panelExt.Controls.Add($btnKeySelectExt)
	
	$intervalExt = Set-UIElement -type 'TextBox' -visible $true -width 47 -height 15 -top 5 -left 59 -bg @(40, 40, 40) -fg @(255, 255, 255) -text '1000' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$panelExt.Controls.Add($intervalExt)
	
	$btnStartExt = Set-UIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 10 -bg @(0, 120, 215) -fg @(255, 255, 255) -text 'Start' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$panelExt.Controls.Add($btnStartExt)
	
	$btnStopExt = Set-UIElement -type 'Button' -visible $true -width 45 -height 20 -top 35 -left 67 -bg @(200, 50, 50) -fg @(255, 255, 255) -text 'Stop' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 9))
	$btnStopExt.Enabled = $false
	$btnStopExt.Visible = $false
	$panelExt.Controls.Add($btnStopExt)

	$btnHotKeyExt = Set-UIElement -type 'Button' -visible $true -width 40 -height 30 -top 1 -left 146 -bg @(200, 50, 50) -fg @(255, 255, 255) -text 'Hotkey' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 6))
	$panelExt.Controls.Add($btnHotKeyExt)
	
	$nameExt = Set-UIElement -type 'TextBox' -visible $true -width 37 -height 17 -top 5 -left 108 -bg @(40, 40, 40) -fg @(255, 255, 255) -text "Name" -font (New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular))
	$panelExt.Controls.Add($nameExt)
	
	$btnRemoveExt = Set-UIElement -type 'Button' -visible $true -width 40 -height 20 -top 35 -left 120 -bg @(150, 50, 50) -fg @(255, 255, 255) -text 'Close' -fs 'Flat' -font (New-Object System.Drawing.Font('Segoe UI', 7))
	$panelExt.Controls.Add($btnRemoveExt)
	
	$extKey = "ext_${instanceId}_$extNum"
	$instanceKey = "instance_$instanceId"
	
	$extData = [PSCustomObject]@{
		Panel          = $panelExt
		BtnKeySelect   = $btnKeySelectExt
		Interval       = $intervalExt
		BtnStart       = $btnStartExt
		BtnStop        = $btnStopExt
		BtnHotKey      = $btnHotKeyExt
		BtnRemove      = $btnRemoveExt
		Name           = $nameExt
		ExtNum         = $extNum
		InstanceId     = $instanceId
		WindowHandle   = $windowHandle
		RunningSpammer = $null
		Position       = $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions.Count
        Hotkey         = $null 
        HotkeyId       = $null 
	}
	
	$global:DashboardConfig.Resources.ExtensionData[$extKey] = $extData
	
	if (-not $global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions.Contains($extKey))
	{
		$global:DashboardConfig.Resources.ExtensionTracking[$instanceKey].ActiveExtensions += $extKey
	}
	
	return $extData
}

function AddExtensionEventHandlers
{
	param($extData, $formData)
	
	$extData.Interval.Add_TextChanged({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			if (-not (CheckRateLimit $this.GetHashCode()))
			{
				return 
			}
		
			$extKey = FindExtensionKeyByControl $this 'Interval'
			if (-not $extKey)
			{
				return 
			}
		
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
		
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
		
			$intervalValue = 0
			if (-not [int]::TryParse($this.Text, [ref]$intervalValue) -or $intervalValue -lt 10)
			{
				$this.Text = '100'
				$intervalValue = 100
			}
		
			UpdateSettings $data $extData -forceWrite
		})

	$extData.Name.Add_TextChanged({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$data = $form.Tag
	
		if (-not (CheckRateLimit $this.GetHashCode()))
		{
			return 
		}
	
		$extKey = FindExtensionKeyByControl $this 'Name'
		if (-not $extKey)
		{
			return 
		}
	
		if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
		{
			return 
		}
	
		$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
		if (-not $extData)
		{
			return 
		}
	
		UpdateSettings $data $extData -forceWrite
	})
	
	$extData.BtnKeySelect.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$data = $form.Tag
		
		$extKey = FindExtensionKeyByControl $this 'BtnKeySelect'
		if (-not $extKey)
		{
			return 
		}
		
		if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
		{
			return 
		}
		
		$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
		if (-not $extData)
		{
			return 
		}
		
		$currentKey = $extData.BtnKeySelect.Text
		$newKey = Show-KeyCaptureDialog $currentKey
		
		if ($newKey -and $newKey -ne $currentKey)
		{
			if ($extData.RunningSpammer)
			{
				$extData.RunningSpammer.Stop()
				$extData.RunningSpammer.Dispose()
				$extData.RunningSpammer = $null
			}
			
			$extData.BtnKeySelect.Text = $newKey
			UpdateSettings $data $extData -forceWrite
			
			ToggleButtonState $extData.BtnStart $extData.BtnStop $false
		}
	})
	
	$extData.BtnStart.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			if (-not (CheckRateLimit $this.GetHashCode() 500))
			{
				return 
			}
		
			$extKey = FindExtensionKeyByControl $this 'BtnStart'
			if (-not $extKey)
			{
				return 
			}
		
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
		
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
		
			$extNum = $extData.ExtNum
		
			$keyValue = $extData.BtnKeySelect.Text
			 if (-not $keyValue -or $keyValue.Trim() -eq '')
			{
			 [System.Windows.Forms.MessageBox]::Show('Please select a key', 'Error')
			 return
			}
		
			$intervalNum = 0
		
			if (-not $extData.Interval -or -not $extData.Interval.Text)
			{
				$intervalNum = 100
			}
			else
			{
				if (-not [int]::TryParse($extData.Interval.Text, [ref]$intervalNum) -or $intervalNum -lt 10)
				{
					$intervalNum = 100
					if ($extData.Interval)
					{
						$extData.Interval.Text = '100'
					}
				}
			}
		
			ToggleButtonState $extData.BtnStart $extData.BtnStop $true
		
			$timerKey = "ExtSpammer_$($data.InstanceId)_$extNum"
			if ($global:DashboardConfig.Resources.Timers.Contains($timerKey))
			{
				$existingTimer = $global:DashboardConfig.Resources.Timers[$timerKey]
				if ($existingTimer)
				{
					$existingTimer.Stop()
					$existingTimer.Dispose()
					$global:DashboardConfig.Resources.Timers.Remove($timerKey)
				}
			}
		
			$spamTimer = CreateSpammerTimer $data.SelectedWindow $keyValue $data.InstanceId $intervalNum $extNum $extKey
		
			$extData.RunningSpammer = $spamTimer
			$global:DashboardConfig.Resources.Timers[$timerKey] = $spamTimer
		})
	
	$extData.BtnStop.Add_Click({
			$form = $this.FindForm()
			if (-not $form -or -not $form.Tag)
			{
				return 
			}
			$data = $form.Tag
		
			if (-not (CheckRateLimit $this.GetHashCode() 500))
			{
				return 
			}
		
			$extKey = FindExtensionKeyByControl $this 'BtnStop'
			if (-not $extKey)
			{
				return 
			}
		
			if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
			{
				return 
			}
		
			$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
			if (-not $extData)
			{
				return 
			}
		
			$extNum = $extData.ExtNum
		
			if ($extData.RunningSpammer)
			{
				$extData.RunningSpammer.Stop()
				$extData.RunningSpammer.Dispose()
				$extData.RunningSpammer = $null
			}
		
			ToggleButtonState $extData.BtnStart $extData.BtnStop $false
		
			$timerKey = "ExtSpammer_$($data.InstanceId)_$extNum"
			if ($global:DashboardConfig.Resources.Timers.Contains($timerKey))
			{
				$global:DashboardConfig.Resources.Timers.Remove($timerKey)
			}
		})

	$extData.BtnHotKey.Add_Click({
		$form = $this.FindForm()
		if (-not $form -or -not $form.Tag)
		{
			return 
		}
		$formData = $form.Tag 

		$extKey = FindExtensionKeyByControl $this 'BtnHotKey'
		if (-not $extKey)
		{
			return 
		}
		
		if (-not $global:DashboardConfig.Resources.ExtensionData.Contains($extKey))
		{
			return 
		}
		
		$extData = $global:DashboardConfig.Resources.ExtensionData[$extKey]
		if (-not $extData)
		{
			return 
		}
		
		$currentHotkeyText = $extData.BtnHotKey.Text
		if ($currentHotkeyText -eq "Hotkey") { $currentHotkeyText = $null } 

        $oldHotkeyIdToUnregister = $extData.HotkeyId

		$newHotkey = Show-KeyCaptureDialog $currentHotkeyText
		
		if ($newHotkey -and $newHotkey -ne $currentHotkeyText) { 

			$extData.Hotkey = $newHotkey
			
			try {
				$scriptBlock = [scriptblock]::Create("ToggleSpecificFtoolInstance -InstanceId '$($extData.InstanceId)' -ExtKey '$($extKey)'")
				$extData.HotkeyId = Set-Hotkey -KeyCombinationString $extData.Hotkey -Action $scriptBlock -OwnerKey $extKey -OldHotkeyId $oldHotkeyIdToUnregister
				$extData.BtnHotKey.Text = $newHotkey 
				$extHotKeyIdDisplay = if ($extData.HotkeyId) { $extData.HotkeyId } else { 'None' }
				Write-Verbose "FTOOL: Registered hotkey $($extData.Hotkey) (ID: $extHotKeyIdDisplay) for extension $($extKey)."
			} catch {
				Write-Warning "FTOOL: Failed to register hotkey $($extData.Hotkey) for extension $($extKey). Error: $_"
                $extData.HotkeyId = $null 
                $extData.Hotkey = $currentHotkeyText
                $extData.BtnHotKey.Text = $currentHotkeyText -or "Hotkey"
			}
			
			UpdateSettings $formData $extData -forceWrite
		} elseif (-not $newHotkey -and $oldHotkeyIdToUnregister) { 
			try {
				Unregister-HotkeyInstance -Id $oldHotkeyIdToUnregister -OwnerKey $extKey
				Write-Verbose "FTOOL: Unregistered hotkey (ID: $($oldHotkeyIdToUnregister)) for extension $($extKey) due to user clear."
			} catch {
				Write-Warning "FTOOL: Failed to unregister hotkey (ID: $($oldHotkeyIdToUnregister)) for extension $($extKey) on clear. Error: $_"
			}
            $extData.HotkeyId = $null
            $extData.Hotkey = $null
            $extData.BtnHotKey.Text = "Hotkey"
            UpdateSettings $formData $extData -forceWrite
        }
	})
	
	$extData.BtnRemove.Add_Click({
			try
			{
				$form = $this.FindForm()
				if (-not $form -or -not $form.Tag)
				{
					return 
				}
			
				if (-not (CheckRateLimit $this.GetHashCode() 1000))
				{
					return 
				}
			
				$extKey = FindExtensionKeyByControl $this 'BtnRemove'
				if (-not $extKey)
				{
					return 
				}
			
				$this.Enabled = $false
			
				RemoveExtension $form $extKey
			
				$form.Tag.ExtensionCount--
			}
			catch
			{
				Write-Verbose ("FTOOL: Error in Remove button handler: {0}" -f $_.Exception.Message) -ForegroundColor Red
			}
		})
}

#endregion

#region Module Exports

Export-ModuleMember -Function FtoolSelectedRow, Set-Hotkey, Resume-AllHotkeys, Resume-PausedKeys, Resume-HotkeysForOwner, Remove-AllHotkeys, Test-HotkeyConflict, Invoke-FtoolAction, PauseAllHotkeys, PauseHotkeysForOwner, Unregister-HotkeyInstance, ToggleSpecificFtoolInstance, ToggleInstanceHotkeys, Get-VirtualKeyMappings, NormalizeKeyString, ParseKeyString, Get-KeyCombinationString, IsModifierKeyCode, Show-KeyCaptureDialog, LoadFtoolSettings, FindOrCreateProfile, InitializeExtensionTracking, GetNextExtensionNumber, FindExtensionKeyByControl, LoadExtensionSettings, UpdateSettings, CreatePositionTimer, RepositionExtensions, CreateSpammerTimer, ToggleButtonState, CheckRateLimit, AddFormCleanupHandler, CleanupInstanceResources, Stop-FtoolForm, RemoveExtension, CreateFtoolForm, AddFtoolEventHandlers, CreateExtensionPanel, AddExtensionEventHandlers

#endregion