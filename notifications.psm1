<# notifications.psm1 #>

#region Globals
if (-not $global:DashboardConfig) { $global:DashboardConfig = @{} }
if (-not $global:DashboardConfig.State) { $global:DashboardConfig.State = @{} }

if (-not $global:LoginNotificationStack) { $global:LoginNotificationStack = [System.Collections.ArrayList]::new() }
if (-not $global:DashboardConfig.State.LoginNotificationMap) { $global:DashboardConfig.State.LoginNotificationMap = @{} }
if (-not $global:DashboardConfig.State.ContainsKey('NotificationHoverActive')) { $global:DashboardConfig.State.NotificationHoverActive = $false }


#endregion


function UpdateNotificationPositions
{
	$global:LoginNotificationStack = [System.Collections.ArrayList]@($global:LoginNotificationStack | Where-Object { $_ -and -not $_.IsDisposed })

	$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
	$baseX = $screen.Right - 320
    
	
	$currentY = $screen.Bottom - 2 
 
	for ($i = 0; $i -lt $global:LoginNotificationStack.Count; $i++)
 {
		$form = $global:LoginNotificationStack[$i]
		if ($form -and -not $form.IsDisposed -and $form.Visible)
		{
			
			$currentY -= $form.Height
			$form.Location = New-Object System.Drawing.Point($baseX, $currentY)
            
			
			$currentY -= 10
		}
	}
}

function CloseToast
{
	param([int]$Key)
    
	$action = {
		if ($global:DashboardConfig.State.LoginNotificationMap.ContainsKey($Key))
		{
			try
			{
				$f = $global:DashboardConfig.State.LoginNotificationMap[$Key]
				if ($f -and -not $f.IsDisposed)
				{
					$f.Close()
					if ($global:LoginNotificationStack.Contains($f)) { $global:LoginNotificationStack.Remove($f) }
				}
			}
			catch {}
			$global:DashboardConfig.State.LoginNotificationMap.Remove($Key)
			UpdateNotificationPositions
		}
	}

	if ($global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.InvokeRequired)
	{
		$global:DashboardConfig.UI.MainForm.BeginInvoke([Action]$action)
	}
 else
	{
		& $action
	}
}


function ShowToast
{
	param(
		[string]$Title,
		[string]$Message,
		[string]$Type = 'Info',
		[int]$Key = 0,
		[int]$TimeoutSeconds = 5,
		[int]$Progress = -1,
		[int]$Height = 0
	)

	
	if ($Key -ne 0 -and $global:DashboardConfig.State.LoginNotificationMap.ContainsKey($Key))
	{
		$existingForm = $global:DashboardConfig.State.LoginNotificationMap[$Key]
		if ($existingForm -and -not $existingForm.IsDisposed)
		{
			
			$pnl = $null
			foreach ($c in $existingForm.Controls)
			{
				if ($c -is [System.Windows.Forms.Panel] -and ($c.Dock -eq 'Fill' -or $c.BorderStyle -eq 'FixedSingle')) { $pnl = $c; break }
			}
			if ($pnl)
			{
				$msgLabel = $null
				foreach ($ctrl in $pnl.Controls)
				{
					if ($ctrl -is [System.Windows.Forms.Label])
					{
						
						$isTitle = ($ctrl.Tag -eq 'Title' -or $ctrl.Font.Bold)
                        
						if ($isTitle)
						{
							$ctrl.Tag = 'Title'
							$ctrl.Text = $Title
						}
						else
						{
							$ctrl.Tag = 'MessageContent'
							$ctrl.Text = $Message
							$msgLabel = $ctrl
						}
					}
					if ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Dock -eq 'Left')
					{
						
						if ($ctrl.Controls.Count -gt 0)
						{
							$strip = $ctrl.Controls[0]
							
							if ($Progress -ge 0)
							{
								$strip.Height = [int]($ctrl.Height * ($Progress / 100))
							}
							else
							{
								
								$strip.Height = $ctrl.Height
							}
							
							$strip.BackColor = if ($Type -eq 'Warning') { [System.Drawing.Color]::Orange } elseif ($Type -eq 'Info') { [System.Drawing.Color]::CornflowerBlue } else { [System.Drawing.Color]::IndianRed }
							$strip.Invalidate(); $strip.Update()
						}
					}
				}

				$existingForm.Refresh()

				
				if ($msgLabel)
				{
					$requiredContentHeight = [Math]::Max(60, $msgLabel.Bottom + 15)
					
					$diff = $requiredContentHeight - $pnl.Height
					if ($diff -ne 0)
					{
						$existingForm.Height += $diff
					}
				}

				
				if ($TimeoutSeconds -gt 0 -and $existingForm.Tag.Timer)
				{
					$t = $existingForm.Tag.Timer
					$tTag = $t.Tag
					$tTag.TotalMs = $TimeoutSeconds * 1000
					$tTag.RemainingMs = $tTag.TotalMs
					
					if ($tTag.Strip.Parent) { $tTag.Strip.Height = $tTag.Strip.Parent.Height }
				}
			}
			if ($global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.InvokeRequired)
			{
				$global:DashboardConfig.UI.MainForm.BeginInvoke([Action] { UpdateNotificationPositions })
			}
			else
			{
				UpdateNotificationPositions
			}
			return
		}
		else
		{
			
			$global:DashboardConfig.State.LoginNotificationMap.Remove($Key)
		}
	}



	$form = New-Object System.Windows.Forms.Form
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	$form.Location = New-Object System.Drawing.Point(-32000, -32000)
	$form.TopMost = $true
	$form.ShowInTaskbar = $false
	$form.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)

	if ($Key -ne 0) { $global:DashboardConfig.State.LoginNotificationMap[$Key] = $form }
	$global:LoginNotificationStack.Add($form)
	$form.Show()
	$form.Add_FormClosed({ param($s,$e) if ($global:LoginNotificationStack.Contains($s)) { $global:LoginNotificationStack.Remove($s); $map = $global:DashboardConfig.State.LoginNotificationMap.GetEnumerator() | Where-Object { $_.Value -eq $s } | Select-Object -First 1; if ($map) { $global:DashboardConfig.State.LoginNotificationMap.Remove($map.Key) } } 
			if ($global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.InvokeRequired)
			{
				$global:DashboardConfig.UI.MainForm.BeginInvoke([Action] { UpdateNotificationPositions })
			}
			else
			{
				UpdateNotificationPositions
			}
		})

	$pnl = New-Object System.Windows.Forms.Panel; $pnl.Dock = 'Fill'; $pnl.BorderStyle = 'FixedSingle'; $form.Controls.Add($pnl)

	$stripContainer = New-Object System.Windows.Forms.Panel
	$stripContainer.Width = 5; $stripContainer.Dock = 'Left'; $stripContainer.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
	$pnl.Controls.Add($stripContainer)

	$strip = New-Object System.Windows.Forms.Panel
	$strip.Width = 5; $strip.Dock = 'Bottom'
    
	$strip.BackColor = if ($Type -eq 'Warning') { [System.Drawing.Color]::Orange } elseif ($Type -eq 'Info') { [System.Drawing.Color]::CornflowerBlue } else { [System.Drawing.Color]::IndianRed }
	$stripContainer.Controls.Add($strip)
	$lblTitle = New-Object System.Windows.Forms.Label
	$lblTitle.Text = $Title; $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold); $lblTitle.ForeColor = [System.Drawing.Color]::White
	$lblTitle.Location = New-Object System.Drawing.Point(15, 10); $lblTitle.AutoSize = $true
	$lblTitle.Tag = 'Title'
	$pnl.Controls.Add($lblTitle)

	$lblMsg = New-Object System.Windows.Forms.Label
	$lblMsg.Text = $Message; $lblMsg.Font = New-Object System.Drawing.Font('Segoe UI', 9); $lblMsg.ForeColor = [System.Drawing.Color]::LightGray
	$lblMsg.Location = New-Object System.Drawing.Point(15, 30)
	$lblMsg.MaximumSize = New-Object System.Drawing.Size(290, 0) 
	$lblMsg.AutoSize = $true
	$lblMsg.Tag = 'MessageContent' 
	$pnl.Controls.Add($lblMsg)

	
	$minHeight = 60
	$calcHeight = $lblMsg.Bottom + 15
	if ($Height -eq 0) { $Height = [Math]::Max($minHeight, $calcHeight) }
    
	$form.Size = New-Object System.Drawing.Size(320, $Height)
	$stripContainer.Height = $Height 
	$strip.Height = $Height

	$form.Tag = @{
		TitleLabel    = $lblTitle
		MsgLabel      = $lblMsg
		ProgressStrip = $strip
	}

	if ($TimeoutSeconds -gt 0)
	{
		$timer = New-Object System.Windows.Forms.Timer
		$timer.Interval = 100

		$timer.Tag = @{
			Form          = $form
			TotalMs       = ($TimeoutSeconds * 1000)
			RemainingMs   = ($TimeoutSeconds * 1000)
			Strip         = $strip 
			InitialHeight = $Height
		}
	}
 else
	{
		
		if ($Progress -ge 0)
		{
			$strip.Height = [int]($stripContainer.Height * ($Progress / 100))
		}
		else
		{
			$strip.Height = $stripContainer.Height
		}
	}

	if ($global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.InvokeRequired)
	{
		$global:DashboardConfig.UI.MainForm.BeginInvoke([Action] { UpdateNotificationPositions })
	}
 else
	{
		UpdateNotificationPositions
	}

	if ($TimeoutSeconds -gt 0)
	{
		$form.Tag.Timer = $timer
		$timer.Add_Tick({ 
				param($s,$e) 
				$tag = $s.Tag
				if ($tag.Form.IsDisposed) { $s.Stop(); $s.Dispose(); return }
            
				
				$mousePos = [System.Windows.Forms.Cursor]::Position
				$anyHover = $false
				$stackCopy = [object[]]@($global:LoginNotificationStack)
				foreach ($f in $stackCopy)
				{
					if ($f -and -not $f.IsDisposed -and $f.Visible -and $f.Bounds.Contains($mousePos))
					{
						$anyHover = $true
						break
					}
				}
				$global:DashboardConfig.State.NotificationHoverActive = $anyHover

				
				if ($global:DashboardConfig.State.LoginActive -or $global:DashboardConfig.State.NotificationHoverActive) { return }

				$tag.RemainingMs -= $s.Interval
				if ($tag.RemainingMs -le 0) { $s.Stop(); if ($tag.Form -and -not $tag.Form.IsDisposed) { $tag.Form.Close() }; $s.Dispose(); return } 
            
				$currentContainerHeight = if ($tag.Strip.Parent) { $tag.Strip.Parent.Height } else { $tag.Form.Height }
				$pct = $tag.RemainingMs / $tag.TotalMs
				if ($pct -lt 0) { $pct = 0 } if ($pct -gt 1) { $pct = 1 } 
				$newHeight = [int]($currentContainerHeight * $pct)
				$tag.Strip.Height = $newHeight 
			})

		$timer.Start()
	}
	$form.Show()
}

function ShowInteractiveNotification
{
	param(
		[string]$Title, [string]$Message, [array]$Buttons, [string]$Type = 'Info', [int]$Key = 0, [int]$TimeoutSeconds = 15
	)
	
	
	ShowToast -Title $Title -Message $Message -Type $Type -Key $Key -TimeoutSeconds $TimeoutSeconds -Height 0

	
	$form = $null
	if ($Key -ne 0 -and $global:DashboardConfig.State.LoginNotificationMap.ContainsKey($Key))
	{
		$form = $global:DashboardConfig.State.LoginNotificationMap[$Key]
	}
 else
	{
		if ($global:LoginNotificationStack.Count -gt 0) { $form = $global:LoginNotificationStack[$global:LoginNotificationStack.Count - 1] }
	}

	if (-not $form -or $form.IsDisposed) { return }

	try
	{
		
		$existingFlows = @($form.Controls | Where-Object { $_.Tag -eq 'InteractiveButtons' -or $_ -is [System.Windows.Forms.FlowLayoutPanel] })
		foreach ($flow in $existingFlows)
		{
			$form.Height -= $flow.Height
			$form.Controls.Remove($flow)
			$flow.Dispose()
		}

		
		$btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
		$btnPanel.Tag = 'InteractiveButtons'
		$btnPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
		$btnPanel.AutoSize = $true
		$btnPanel.Padding = [System.Windows.Forms.Padding]::new(6)
		$btnPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight

		
		$nonFullWidthButtons = $Buttons | Where-Object { -not ($_.ContainsKey('FullWidth') -and $_['FullWidth']) }
		$nonFullWidthCount = $nonFullWidthButtons.Count
		$totalWidth = 300 
		$margin = 3
		$totalButtonWidth = $totalWidth
		if ($nonFullWidthCount -gt 1) { $totalButtonWidth = $totalWidth - ($margin * ($nonFullWidthCount - 1)) }
		$baseWidth = if ($nonFullWidthCount -gt 0) { [math]::Floor($totalButtonWidth / $nonFullWidthCount) } else { 0 }
		$remainder = if ($nonFullWidthCount -gt 0) { $totalButtonWidth % $nonFullWidthCount } else { 0 }

		
		Write-Verbose "ShowInteractiveNotification: Attaching $($Buttons.Count) button(s)"
		$nonFullWidthIndex = 0
		foreach ($b in $Buttons)
		{
			try
			{
				$text = if ($b.ContainsKey('Text')) { $b['Text'] } else { $b.Text }
				$action = if ($b.ContainsKey('Action')) { $b['Action'] } else { $b.Action }
				try { Write-Verbose "Button '$text' action type: $($action.GetType().FullName)" } catch { Write-Verbose "Button '$text' action type: <unknown>" }
				try { Write-Verbose "Button '$text' action value: $([string]::Copy(($action.ToString())))" } catch { }
				$btn = New-Object System.Windows.Forms.Button
				$btn.Text = $text
				$btn.AutoSize = $false
				$btn.Height = 30
				$btn.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
				$btn.ForeColor = [System.Drawing.Color]::White
				$btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
				$btn.FlatAppearance.BorderSize = 0
				$btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
				$btn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
				
				try { $btn.Tag = $action } catch {}

				if ($b.ContainsKey('FullWidth') -and $b['FullWidth'])
				{
					$btn.Width = ($totalButtonWidth + 6)
					$btn.Margin = [System.Windows.Forms.Padding]::new(0, $margin, 0, 0)
				}
				else
				{
					$btn.Width = $baseWidth
					if ($nonFullWidthIndex -lt $remainder)
					{
						$btn.Width++
					}
					if ($nonFullWidthIndex -lt ($nonFullWidthCount - 1))
					{
						$btn.Margin = [System.Windows.Forms.Padding]::new(0, 0, $margin, 0)
					}
					else
					{
						$btn.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 0)
					}
					$nonFullWidthIndex++
				}

				
				$null = $localAction; $localAction = $action

				
				$btn.Add_Click({ param($s,$e) try
						{
							$act = $null
							try { $act = $s.Tag } catch {}
							$qExists = ($null -ne $global:DashboardConfig.State -and $null -ne $global:DashboardConfig.State.NotificationActionQueue)
							Write-Verbose "Button Click: Text='$($s.Text)' QExists=$qExists ActionType=$($act.GetType().FullName)"

							if (-not $global:DashboardConfig.State.NotificationActionQueue) { $global:DashboardConfig.State.NotificationActionQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new()) }

							
							if ($act -is [System.Collections.IDictionary])
							{
								$actionName = if ($act.ContainsKey('Action')) { $act['Action'] } else { $null }
								$actionPid = if ($act.ContainsKey('Pid')) { [int]$act['Pid'] } else { 0 }
								$closeNow = $true
								if ($act.ContainsKey('CloseOnAction')) { $closeNow = [bool]$act['CloseOnAction'] }

								if ($actionName) { $global:DashboardConfig.State.NotificationActionQueue.Enqueue(@{ Action = $actionName; Pid = $actionPid }) }
								if ($closeNow -and $actionPid -ne 0) { try { CloseToast -Key $actionPid } catch {} }
							}
							else
							{
								if ($act -is [scriptblock]) { & $act } else { & ([scriptblock]::Create($act.ToString())) }
							}
						}
						catch { Write-Verbose "Button click handler error: $($_.Exception.Message)"; Write-Verbose ($_.Exception | Out-String) } })

				$btnPanel.Controls.Add($btn)
			}
			catch {}
		}

		
		$form.Controls.Add($btnPanel)
		
		$form.Height += $btnPanel.Height
        
		
		if ($form.Tag.Timer)
		{
			$tTag = $form.Tag.Timer.Tag
			if ($tTag.RemainingMs -eq $tTag.TotalMs)
			{
				if ($tTag.Strip.Parent) { $tTag.Strip.Height = $tTag.Strip.Parent.Height }
			}
		}

		
		if ($global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.InvokeRequired)
		{
			$global:DashboardConfig.UI.MainForm.BeginInvoke([Action] { UpdateNotificationPositions })
		}
		else
		{
			UpdateNotificationPositions
		}
	}
 catch {}
}

function ShowReconnectInteractiveNotification
{
	param(
		[string]$Title,
		[string]$Message,
		[hashtable]$Buttons,
		[string]$Type = 'Warning',
		[int]$RelatedPid = 0,
		[int]$TimeoutSeconds = 15
	)

	
	$btnArray = @()
	$delayBtn = $null

	if ($Buttons)
	{
		foreach ($key in $Buttons.Keys)
		{
			$cmd = $Buttons[$key]
			$escapedCmd = [string]::Copy($cmd)
			$pidVal = $RelatedPid
			$closeOnActionSetting = $true

			
			
			$btnItem = @{ Text = $key; Action = @{ Action = $escapedCmd; Pid = $pidVal; CloseOnAction = $closeOnActionSetting } }
                
			if ($cmd -eq 'Delay')
			{
				$btnItem['FullWidth'] = $true
				$delayBtn = $btnItem
			}
			else
			{
				$btnArray += $btnItem
			}
		}
	}

	if ($delayBtn)
	{
		$btnArray += $delayBtn
	}

	ShowInteractiveNotification -Title $Title -Message $Message -Buttons $btnArray -Type $Type -Key $RelatedPid -TimeoutSeconds $TimeoutSeconds
}

Export-ModuleMember -Function *
