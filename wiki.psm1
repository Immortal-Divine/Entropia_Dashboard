<# wiki.psm1 #>

function Show-Wiki
{
	[CmdletBinding()]
	param()


	$LocalWikiPath = Join-Path $PSScriptRoot 'wiki.json'
	
	$LocalWikiData = $null
	$OnlineWikiData = $null
	$SourceUsed = 'Connecting...'
	$IsOffline = $false
	$Disposables = New-Object System.Collections.Generic.List[object]
	$WikiJsonURL = 'https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/main/wiki.json'


	if (Test-Path $LocalWikiPath)
	{
		try
		{
			$RawContent = Get-Content $LocalWikiPath -Raw -Encoding UTF8
			$LocalWikiData = $RawContent | ConvertFrom-Json

			if ($LocalWikiData -is [string]) {
				$LocalWikiData = $LocalWikiData | ConvertFrom-Json
			}
		}
		catch {}
	}

	if ($LocalWikiData) {
		$SourceUsed = 'Offline (Local File)'
		$WikiData = $LocalWikiData
	} else {
		$SourceUsed = 'Offline'
		$IsOffline = $true
		$WikiData = [ordered]@{}
	}

	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	if (-not ([System.Threading.SynchronizationContext]::Current)) {
		[System.Windows.Forms.WindowsFormsSynchronizationContext]::SetSynchronizationContext((New-Object System.Windows.Forms.WindowsFormsSynchronizationContext))
	}
	$syncContext = [System.Threading.SynchronizationContext]::Current

	$Theme = @{
		Background   = [System.Drawing.ColorTranslator]::FromHtml('#0f1219')
		PanelColor   = [System.Drawing.ColorTranslator]::FromHtml('#232838')
		InputBack    = [System.Drawing.ColorTranslator]::FromHtml('#161a26')
		AccentColor  = [System.Drawing.ColorTranslator]::FromHtml('#ff2e4c')
		TextColor    = [System.Drawing.ColorTranslator]::FromHtml('#ffffff')
		SubTextColor = [System.Drawing.ColorTranslator]::FromHtml('#a0a5b0')
		SuccessColor = [System.Drawing.ColorTranslator]::FromHtml('#28a745')
	}


	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Entropia Wiki - $SourceUsed"
	$form.Size = New-Object System.Drawing.Size(1340, 750)
	$form.StartPosition = 'CenterScreen'
	$form.BackColor = $Theme.Background
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

	if ($global:DashboardConfig.UI.MainForm -and $global:DashboardConfig.UI.MainForm.Icon)
	{
		$form.Icon = $global:DashboardConfig.UI.MainForm.Icon
	}

    # Register for global cleanup
    if (-not $global:DashboardConfig.Resources.ContainsKey('WikiForm')) {
        $global:DashboardConfig.Resources['WikiForm'] = $null
    }
    $global:DashboardConfig.Resources.WikiForm = $form

	$form.Tag = $WikiData
    

	$NavHistory = New-Object System.Collections.Stack



	$headerPanel = New-Object System.Windows.Forms.Panel
	$headerPanel.Dock = 'Top'
	$headerPanel.Height = 40
	$headerPanel.BackColor = $Theme.PanelColor
	$headerPanel.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
	$form.Controls.Add($headerPanel) | Out-Null


	$statusPanel = New-Object System.Windows.Forms.Panel
	$statusPanel.Dock = 'Bottom'
	$statusPanel.Height = 25
	$statusPanel.BackColor = $Theme.PanelColor
	$form.Controls.Add($statusPanel) | Out-Null
    
	$lblStatus = New-Object System.Windows.Forms.Label
	$lblStatus.Text = "Source: $SourceUsed"
	$lblStatus.ForeColor = if ($IsOffline) { $Theme.SubTextColor } else { $Theme.SuccessColor }
	$lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8)
	$lblStatus.AutoSize = $true
	$lblStatus.Location = New-Object System.Drawing.Point(10, 5)
	$statusPanel.Controls.Add($lblStatus) | Out-Null


    $splitContainer = New-Object System.Windows.Forms.SplitContainer
    $splitContainer.Dock = "Fill"
    $splitContainer.Orientation = "Vertical"
    $splitContainer.SplitterWidth = 2
    $splitContainer.FixedPanel = 'Panel1'
    $splitContainer.BackColor = $Theme.Background
    $form.Controls.Add($splitContainer) | Out-Null
    $splitContainer.SplitterDistance = 300

	$splitContainer.BringToFront()


	$titleLabel = New-Object System.Windows.Forms.Label
	$titleLabel.Text = 'ENTROPIA KNOWLEDGE BASE'
	$titleLabel.ForeColor = $Theme.AccentColor
	$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
	$titleLabel.AutoSize = $true
	$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
	$headerPanel.Controls.Add($titleLabel) | Out-Null

	$closeBtn = New-Object System.Windows.Forms.Label
	$closeBtn.Text = 'X'
	$closeBtn.ForeColor = $Theme.SubTextColor
	$closeBtn.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
	$closeBtn.Cursor = 'Hand'
	$closeBtn.AutoSize = $true
	$closeBtn.Anchor = 'Top, Right'
	$closeBtn.Location = New-Object System.Drawing.Point(($form.Width - 35), 8) 
	$headerPanel.Controls.Add($closeBtn) | Out-Null


	$navPanel = New-Object System.Windows.Forms.Panel
	$navPanel.Dock = 'Fill'
	$navPanel.Padding = New-Object System.Windows.Forms.Padding(0)
	$splitContainer.Panel1.Controls.Add($navPanel) | Out-Null

	$searchContainer = New-Object System.Windows.Forms.Panel
	$searchContainer.Dock = 'Top'
	$searchContainer.Height = 50
	$searchContainer.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
	$navPanel.Controls.Add($searchContainer) | Out-Null

	$searchBox = New-Object System.Windows.Forms.TextBox
	$searchBox.Dock = 'Fill' 
	$searchBox.BackColor = $Theme.InputBack
	$searchBox.ForeColor = $Theme.SubTextColor
	$searchBox.BorderStyle = 'FixedSingle'
	$searchBox.Font = New-Object System.Drawing.Font('Segoe UI', 10)
	$searchBox.Text = 'Search...'
	$searchContainer.Controls.Add($searchBox) | Out-Null

	$treeView = New-Object Custom.DarkTreeView
	$treeView.Dock = 'Fill'
	$treeView.BackColor = $Theme.InputBack
	$treeView.ForeColor = $Theme.TextColor
	$treeView.BorderStyle = 'None'
	$treeView.Font = New-Object System.Drawing.Font('Segoe UI', 10)
	$treeView.FullRowSelect = $true
	$treeView.ShowLines = $false
	$treeView.ShowPlusMinus = $true
	$treeView.Indent = 20
	$treeView.ItemHeight = 28
	$navPanel.Controls.Add($treeView) | Out-Null
	if ([Custom.Native] -as [Type]) { [Custom.Native]::UseImmersiveDarkMode($treeView.Handle) }
	$treeView.BringToFront()


    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Dock = "Fill"
    $contentPanel.Padding = New-Object System.Windows.Forms.Padding(25, 20, 25, 20)
    $splitContainer.Panel2.Controls.Add($contentPanel) | Out-Null


    $navToolbar = New-Object System.Windows.Forms.Panel
    $navToolbar.Dock = "Top"
    $navToolbar.Height = 40
    $contentPanel.Controls.Add($navToolbar) | Out-Null
    $navToolbar.BringToFront()

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [char]0x25C0+"BACK"
    $btnBack.Size = New-Object System.Drawing.Size(80, 30)
    $btnBack.Location = New-Object System.Drawing.Point(0, 5)
    $btnBack.FlatStyle = "Flat"
    $btnBack.FlatAppearance.BorderSize = 0
    $btnBack.BackColor = $Theme.SubTextColor
    $btnBack.ForeColor = $Theme.TextColor
    $btnBack.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnBack.Cursor = "Hand"
    $btnBack.Enabled = $false
    $navToolbar.Controls.Add($btnBack) | Out-Null


    $pnlEditFields = New-Object System.Windows.Forms.Panel
    $pnlEditFields.Location = New-Object System.Drawing.Point(25, 130)
    $pnlEditFields.Size = New-Object System.Drawing.Size(700, 110)
    $pnlEditFields.Visible = $false
    $contentPanel.Controls.Add($pnlEditFields) | Out-Null

    $lblEditTitle = New-Object System.Windows.Forms.Label
    $lblEditTitle.Text = "Title:"
    $lblEditTitle.Location = New-Object System.Drawing.Point(0, 3)
    $lblEditTitle.AutoSize = $true
    $lblEditTitle.ForeColor = $Theme.SubTextColor
    $pnlEditFields.Controls.Add($lblEditTitle)

    $txtEditTitle = New-Object System.Windows.Forms.TextBox
    $txtEditTitle.Location = New-Object System.Drawing.Point(50, 0)
    $txtEditTitle.Width = 640
    $txtEditTitle.Anchor = "Top, Left, Right"
    $txtEditTitle.BackColor = $Theme.InputBack
    $txtEditTitle.ForeColor = $Theme.TextColor
    $txtEditTitle.BorderStyle = "FixedSingle"
    $pnlEditFields.Controls.Add($txtEditTitle)
    if ([Custom.Native] -as [Type]) { [Custom.Native]::UseImmersiveDarkMode($txtEditTitle.Handle) }

    $lblEditUrl = New-Object System.Windows.Forms.Label
    $lblEditUrl.Text = "URL:"
    $lblEditUrl.Location = New-Object System.Drawing.Point(0, 33)
    $lblEditUrl.AutoSize = $true
    $lblEditUrl.ForeColor = $Theme.SubTextColor
    $pnlEditFields.Controls.Add($lblEditUrl)

    $txtEditUrl = New-Object System.Windows.Forms.TextBox
    $txtEditUrl.Location = New-Object System.Drawing.Point(50, 30)
    $txtEditUrl.Width = 640
    $txtEditUrl.Anchor = "Top, Left, Right"
    $txtEditUrl.BackColor = $Theme.InputBack
    $txtEditUrl.ForeColor = $Theme.TextColor
    $txtEditUrl.BorderStyle = "FixedSingle"
    $pnlEditFields.Controls.Add($txtEditUrl)
    if ([Custom.Native] -as [Type]) { [Custom.Native]::UseImmersiveDarkMode($txtEditUrl.Handle) }

    $lblEditNote = New-Object System.Windows.Forms.Label
    $lblEditNote.Text = "Note:"
    $lblEditNote.Location = New-Object System.Drawing.Point(0, 63)
    $lblEditNote.AutoSize = $true
    $lblEditNote.ForeColor = $Theme.SubTextColor
    $pnlEditFields.Controls.Add($lblEditNote)

    $txtEditNote = New-Object System.Windows.Forms.TextBox
    $txtEditNote.Location = New-Object System.Drawing.Point(50, 60)
    $txtEditNote.Width = 640
    $txtEditNote.Anchor = "Top, Left, Right"
    $txtEditNote.BackColor = $Theme.InputBack
    $txtEditNote.ForeColor = $Theme.TextColor
    $txtEditNote.BorderStyle = "FixedSingle"
    $pnlEditFields.Controls.Add($txtEditNote)
    if ([Custom.Native] -as [Type]) { [Custom.Native]::UseImmersiveDarkMode($txtEditNote.Handle) }

    $txtDescription = New-Object System.Windows.Forms.RichTextBox
    $txtDescription.Location = New-Object System.Drawing.Point(25, 180)
    $txtDescription.Size = New-Object System.Drawing.Size(1000, 480)
    $txtDescription.Anchor = "Top, Bottom, Left, Right"
    $txtDescription.BackColor = $Theme.Background
    $txtDescription.ForeColor = $Theme.TextColor
    $txtDescription.BorderStyle = "None"
    $txtDescription.ReadOnly = $true
    $txtDescription.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $txtDescription.Text = "Select an item from the menu on the left to view details.`n`nUse the Search box to quickly find guides, items, or systems."
    $contentPanel.Controls.Add($txtDescription) | Out-Null
    if ([Custom.Native] -as [Type]) { [Custom.Native]::UseImmersiveDarkMode($txtDescription.Handle) }

    $webDescription = New-Object System.Windows.Forms.WebBrowser
    $webDescription.Location = New-Object System.Drawing.Point(25, 180)
    $webDescription.Size = New-Object System.Drawing.Size(1000, 480)
    $webDescription.Anchor = "Top, Bottom, Left, Right"
    $webDescription.ScriptErrorsSuppressed = $true
    $webDescription.Visible = $false
    $contentPanel.Controls.Add($webDescription) | Out-Null


	$pnlFormatting = New-Object System.Windows.Forms.FlowLayoutPanel
	$pnlFormatting.Location = New-Object System.Drawing.Point(25, 245)
	$pnlFormatting.Size = New-Object System.Drawing.Size(700, 35)
	$pnlFormatting.Visible = $false
	$pnlFormatting.AutoSize = $true
	$contentPanel.Controls.Add($pnlFormatting) | Out-Null

	$AddFormatBtn = {
		param($text, $tag, $tooltip, $targetCtrl)
		$btn = New-Object System.Windows.Forms.Button
		$btn.Text = $text
		$btn.Tag = $tag
		$btn.Size = New-Object System.Drawing.Size(40, 30)
		$btn.FlatStyle = 'Flat'
		$btn.FlatAppearance.BorderSize = 1
		$btn.FlatAppearance.BorderColor = $Theme.PanelColor
		$btn.BackColor = $Theme.InputBack
		$btn.ForeColor = $Theme.TextColor
		$btn.Cursor = 'Hand'
		$btn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
		if ($tooltip) {
			$tt = New-Object System.Windows.Forms.ToolTip
			$tt.SetToolTip($btn, $tooltip)
		}
		$btn.Add_Click({
			$t = $targetCtrl
            if ($null -eq $t) { return }
			$sel = $t.SelectedText
			$mode = $this.Tag
			
			if ($mode -eq 'Table') {
				$t.SelectedText = "`n| Header 1 | Header 2 |`n|---|---|`n| Cell 1 | Cell 2 |`n"
			} elseif ($mode -eq 'Link') {
				$clip = try { [System.Windows.Forms.Clipboard]::GetText() } catch { '' }
				if ([string]::IsNullOrEmpty($sel)) { $sel = "text" }
				$t.SelectedText = "[$sel]($clip)"
			} elseif ($mode -eq 'Img') {
				$clip = try { [System.Windows.Forms.Clipboard]::GetText() } catch { '' }
				if ([string]::IsNullOrEmpty($sel)) { $sel = "Image" }
				$t.SelectedText = "[$sel]$clip"
			} elseif ($mode -eq 'Wiki') {
				$clip = try { [System.Windows.Forms.Clipboard]::GetText() } catch { '' }
				if ([string]::IsNullOrEmpty($sel)) { $sel = "Node" }
				$t.SelectedText = "[$sel](wiki:$sel)"
			} else {
				$p = $mode[0]; $s = $mode[1]
				if ([string]::IsNullOrEmpty($sel)) { $sel = "text" }
				$t.SelectedText = "$p$sel$s"
			}
			$t.Focus()
		}.GetNewClosure())
		$pnlFormatting.Controls.Add($btn)
	}

	& $AddFormatBtn "B" @('**','**') "Bold" $txtDescription
	& $AddFormatBtn "I" @('*','*') "Italic" $txtDescription
	& $AddFormatBtn "S" @('~~','~~') "Strikethrough" $txtDescription
	& $AddFormatBtn "</>" @('`','`') "Code" $txtDescription
	& $AddFormatBtn "Link" "Link" "Link" $txtDescription
	& $AddFormatBtn "Img" "Img" "Image" $txtDescription
	& $AddFormatBtn "Wiki" "Wiki" "Node" $txtDescription
	& $AddFormatBtn "Table" "Table" "Insert Table" $txtDescription


    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "SAVE"
    $btnSave.Size = New-Object System.Drawing.Size(80, 30)
    $btnSave.Anchor = 'Top, Right'
    $btnSave.Location = New-Object System.Drawing.Point(($navToolbar.Width - 90), 5)
    $btnSave.FlatStyle = "Flat"
    $btnSave.FlatAppearance.BorderSize = 0
    $btnSave.BackColor = $Theme.SuccessColor
    $btnSave.ForeColor = "White"
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnSave.Cursor = "Hand"
    $btnSave.Visible = $false
    $navToolbar.Controls.Add($btnSave) | Out-Null

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "ADD"
    $btnAdd.Size = New-Object System.Drawing.Size(80, 30)
    $btnAdd.Anchor = 'Top, Right'
    $btnAdd.Location = New-Object System.Drawing.Point(($navToolbar.Width - 270), 5)
    $btnAdd.FlatStyle = "Flat"
    $btnAdd.FlatAppearance.BorderSize = 0
    $btnAdd.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078d7")
    $btnAdd.ForeColor = "White"
    $btnAdd.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnAdd.Cursor = "Hand"
    $btnAdd.Visible = $false
    $navToolbar.Controls.Add($btnAdd) | Out-Null

    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "DELETE"
    $btnDelete.Size = New-Object System.Drawing.Size(80, 30)
    $btnDelete.Anchor = 'Top, Right'
    $btnDelete.Location = New-Object System.Drawing.Point(($navToolbar.Width - 180), 5)
    $btnDelete.FlatStyle = "Flat"
    $btnDelete.FlatAppearance.BorderSize = 0
    $btnDelete.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#d9534f")
    $btnDelete.ForeColor = "White"
    $btnDelete.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnDelete.Cursor = "Hand"
    $btnDelete.Visible = $false
    $navToolbar.Controls.Add($btnDelete) | Out-Null

    $chkEdit = New-Object System.Windows.Forms.CheckBox
    $chkEdit.Name = "chkEdit"
    $chkEdit.Text = "EDIT"
    $chkEdit.Appearance = 'Button'
    $chkEdit.Size = New-Object System.Drawing.Size(80, 30)
    $chkEdit.Anchor = 'Top, Right'
    $chkEdit.Location = New-Object System.Drawing.Point(($navToolbar.Width - 360), 5)
    $chkEdit.FlatStyle = "Flat"
    $chkEdit.FlatAppearance.BorderSize = 1
    $chkEdit.FlatAppearance.BorderColor = $Theme.SubTextColor
    $chkEdit.BackColor = $Theme.Background
    $chkEdit.ForeColor = $Theme.SubTextColor
    $chkEdit.TextAlign = 'MiddleCenter'
    $chkEdit.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $chkEdit.Cursor = "Hand"
    $navToolbar.Controls.Add($chkEdit) | Out-Null

    $chkViewMode = New-Object System.Windows.Forms.CheckBox
    $chkViewMode.Name = "chkViewMode"
    $chkViewMode.Text = "VIEW ONLINE"
    $chkViewMode.Appearance = 'Button'
    $chkViewMode.Size = New-Object System.Drawing.Size(110, 30)
    $chkViewMode.Anchor = 'Top, Right'
    $chkViewMode.Location = New-Object System.Drawing.Point(($navToolbar.Width - 480), 5)
    $chkViewMode.FlatStyle = "Flat"
    $chkViewMode.FlatAppearance.BorderSize = 1
    $chkViewMode.FlatAppearance.BorderColor = $Theme.SubTextColor
    $chkViewMode.BackColor = $Theme.Background
    $chkViewMode.ForeColor = $Theme.SubTextColor
    $chkViewMode.TextAlign = 'MiddleCenter'
    $chkViewMode.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $chkViewMode.Cursor = "Hand"
    $navToolbar.Controls.Add($chkViewMode) | Out-Null

    $lblContentTitle = New-Object System.Windows.Forms.Label
    $lblContentTitle.Location = New-Object System.Drawing.Point(25, 80)
    $lblContentTitle.Size = New-Object System.Drawing.Size(700, 45)
    $lblContentTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $lblContentTitle.ForeColor = $Theme.TextColor
    $lblContentTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblContentTitle.Text = "Welcome"
    $contentPanel.Controls.Add($lblContentTitle) | Out-Null

    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.Location = New-Object System.Drawing.Point(25, 130)
    $btnPanel.Size = New-Object System.Drawing.Size(650, 45)
    $btnPanel.AutoSize = $true
    $contentPanel.Controls.Add($btnPanel) | Out-Null

    $btnOpenLink = New-Object System.Windows.Forms.Button
    $btnOpenLink.Text = "OPEN LINK"
    $btnOpenLink.Size = New-Object System.Drawing.Size(140, 35)
    $btnOpenLink.FlatStyle = "Flat"
    $btnOpenLink.FlatAppearance.BorderSize = 0
    $btnOpenLink.BackColor = $Theme.AccentColor
    $btnOpenLink.ForeColor = "White"
    $btnOpenLink.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOpenLink.Cursor = "Hand"
    $btnOpenLink.Visible = $false
    $btnPanel.Controls.Add($btnOpenLink) | Out-Null

    $btnCopyLink = New-Object System.Windows.Forms.Button
    $btnCopyLink.Text = "COPY LINK"
    $btnCopyLink.Size = New-Object System.Drawing.Size(120, 35)
    $btnCopyLink.FlatStyle = "Flat"
    $btnCopyLink.FlatAppearance.BorderSize = 1
    $btnCopyLink.FlatAppearance.BorderColor = $Theme.SubTextColor
    $btnCopyLink.BackColor = $Theme.Background
    $btnCopyLink.ForeColor = $Theme.SubTextColor
    $btnCopyLink.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnCopyLink.Cursor = "Hand"
    $btnCopyLink.Visible = $false
    $btnCopyLink.Margin = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
    $btnPanel.Controls.Add($btnCopyLink) | Out-Null



	$closeBtn.Add_Click({ try { $form.Close() } catch {} }.GetNewClosure())

	$dragAction = {
		param($src, $e)
		if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left)
		{
			try
			{
				if ([Custom.Native] -as [Type])
				{
					[Custom.Native]::ReleaseCapture()
					[Custom.Native]::SendMessage($form.Handle, 0xA1, 0x2, 0)
				}
			}
			catch {}
		}
	}.GetNewClosure()

	$form.Add_MouseDown($dragAction)
	$headerPanel.Add_MouseDown($dragAction)
	$titleLabel.Add_MouseDown($dragAction)


	$getKeys = { 
		param($obj) 
		if ($obj -is [System.Collections.IDictionary]) { return $obj.Keys }
		if ($obj.PSObject) { return $obj.PSObject.Properties.Name }
		return @()
	}


	$AddWikiNode = {
		param($Self, $parentNode, $nodeName, $nodeData, $Filter, $isFiltering, $ParentData, $Key)
        

		$currentNode = New-Object System.Windows.Forms.TreeNode($nodeName)
		$currentNode.Tag = $nodeData
		$currentNode.Name = $nodeName
        

        $currentNode | Add-Member -MemberType NoteProperty -Name "WikiMeta" -Value @{ Parent = $ParentData; Key = $Key } -Force
        
		$hasMatches = $false
		$childFound = $false


		$isComplex = ($nodeData -is [System.Management.Automation.PSCustomObject] -or $nodeData -is [System.Collections.IDictionary])
		$isArray   = ($nodeData -is [System.Array] -or $nodeData -is [System.Collections.ArrayList]) -and $nodeData -isnot [string]

		if ($isComplex)
		{

			$childKeys = & $getKeys $nodeData

			foreach ($childName in $childKeys)
			{


				if ($childName -in 'meta','description','url','URL') { continue }

				$childContent = if ($nodeData -is [System.Collections.IDictionary]) { $nodeData[$childName] } else { $nodeData.$childName }
                

				$subNodeHasMatches = & $Self $Self $currentNode $childName $childContent $Filter $isFiltering $nodeData $childName
				if ($subNodeHasMatches) { 
					$childFound = $true
					$hasMatches = $true 
				}
			}
		}
		elseif ($isArray)
		{

            if ($nodeData -is [System.Array]) {
                $newList = [System.Collections.ArrayList]::new($nodeData)
                if ($ParentData -is [System.Collections.IDictionary]) { $ParentData[$Key] = $newList }
                elseif ($ParentData -is [System.Collections.IList]) { $ParentData[$Key] = $newList }
                elseif ($ParentData) { $ParentData | Add-Member -MemberType NoteProperty -Name $Key -Value $newList -Force }
                $nodeData = $newList
                $currentNode.Tag = $newList
            }


            $idx = 0
			foreach ($item in $nodeData)
			{
				if ($item -is [string]) {
					$isLeafMatch = if ($isFiltering) { $item -match $Filter } else { $true }
					if ($isLeafMatch) {
						$leafNode = New-Object System.Windows.Forms.TreeNode($item)
						$leafNode.Name = $item
						$leafNode.Tag = $item 
                        $leafNode | Add-Member -MemberType NoteProperty -Name "WikiMeta" -Value @{ Parent = $nodeData; Key = $idx } -Force
						$currentNode.Nodes.Add($leafNode) | Out-Null
						$childFound = $true
						$hasMatches = $true
					}
				}
				else {
					$subNodeHasMatches = & $Self $Self $currentNode "Item" $item $Filter $isFiltering $nodeData $idx
					if ($subNodeHasMatches) { 
						$childFound = $true
						$hasMatches = $true 
					}
				}
                $idx++
			}
		}
		else
		{


			$isLeafMatch = if ($isFiltering) { $nodeName -match $Filter -or ($nodeData -match $Filter) } else { $true }
			
			if ($isLeafMatch) {
				$currentNode.Text = "$nodeName"
				$hasMatches = $true
			}
		}


		$selfMatches = $false
		if ($isFiltering) {
			if ($nodeName -match $Filter) {
				$selfMatches = $true
			} elseif ($isComplex) {

				$props = & $getKeys $nodeData
				foreach ($p in $props) {
					$v = if ($nodeData -is [System.Collections.IDictionary]) { $nodeData[$p] } else { $nodeData.$p }
					if ($v -is [string] -and $v -match $Filter) {
						$selfMatches = $true; break
					}
				}
			}
		} else {
			$selfMatches = $true
		}


		if ($childFound -and ($hasMatches -or $selfMatches)) {
			$parentNode.Nodes.Add($currentNode) | Out-Null
			if ($isFiltering -or ($nodeName -in 'Home','Equipment Progression')) { $currentNode.Expand() }
			return $true
		}
		elseif (-not $childFound -and ($hasMatches -or $selfMatches)) { 

			$parentNode.Nodes.Add($currentNode) | Out-Null
			return $true
		}
		return $hasMatches -or $selfMatches 
	}.GetNewClosure()


	$PopulateTree = {
		param($Filter, $Data, $Tree)
        
		try
		{
			$Tree.BeginUpdate()
			$Tree.Nodes.Clear()

			$isFiltering = -not [string]::IsNullOrWhiteSpace($Filter) -and $Filter -ne 'Search...'
			$parentKeys = & $getKeys $Data

			foreach ($catName in $parentKeys)
			{
				if ($catName -eq 'meta') { continue }
				
				$catData = if ($Data -is [System.Collections.IDictionary]) { $Data[$catName] } else { $Data.$catName }
                

				& $AddWikiNode $AddWikiNode $Tree $catName $catData $Filter $isFiltering $Data $catName
			}
		}
		catch { Write-Verbose "Wiki Search Error: $_" } 
		finally
		{
			$Tree.EndUpdate()
		}
	}.GetNewClosure()

    $webDescription.Add_Navigating({
        param($s, $e)
        $url = $e.Url.ToString()
        if ($url -eq 'about:blank') { return }
        $e.Cancel = $true
        
        if ($url.StartsWith('http://wiki/') -or $url.StartsWith('about:wiki:') -or $url.StartsWith('wiki:')) {
            $target = $null
            if ($url.StartsWith('http://wiki/')) { $target = $url.Substring(12) }
            elseif ($url.StartsWith('about:wiki:')) { $target = $url.Substring(11) }
            else { $target = $url.Substring(5) }

            if ($null -ne $target) {
                try {
                    $target = [System.Uri]::UnescapeDataString($target)
                    $target = $target.Trim()
                    $nodes = $treeView.Nodes.Find($target, $true)
                    if ($nodes.Count -gt 0) {
                        $treeView.SelectedNode = $nodes[0]
                        $treeView.Focus()
                    }
                } catch {}
            }
        } else {
            try { [System.Diagnostics.Process]::Start($url) | Out-Null } catch {}
        }
    }.GetNewClosure())

    $ConvertMarkdownToHtml = {
        param($text)
        if ([string]::IsNullOrWhiteSpace($text)) { return "" }


        $text = $text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'


        $lines = $text -split "`n"
        $htmlLines = New-Object System.Collections.Generic.List[string]
        
        $inTable = $false
        $inList = $false
        $inOrderedList = $false

        foreach ($line in $lines) {
            $trimLine = $line.Trim()
            

            if ($trimLine -match '^\|(.+)\|$') {
                if (-not $inTable) {
                    $htmlLines.Add("<table>")
                    $inTable = $true
                    $isHeader = $true
                } else {
                    if ($trimLine -match '^\|[\s\-:|]+\|$') { continue }
                    $isHeader = $false
                }
                
                $cells = $trimLine.Trim('|').Split('|')
                $rowHtml = "<tr>"
                foreach ($cell in $cells) {
                    $tag = if ($isHeader) { "th" } else { "td" }
                    $rowHtml += "<$tag>$($cell.Trim())</$tag>"
                }
                $rowHtml += "</tr>"
                $htmlLines.Add($rowHtml)
                continue
            } elseif ($inTable) {
                $htmlLines.Add("</table>")
                $inTable = $false
            }


            if ($trimLine -match '^[\-\*]\s+(.+)$') {
                if ($inOrderedList) { $htmlLines.Add("</ol>"); $inOrderedList = $false }
                if (-not $inList) { $htmlLines.Add("<ul>"); $inList = $true }
                $htmlLines.Add("<li>$($matches[1])</li>")
                continue
            } elseif ($inList) {
                $htmlLines.Add("</ul>")
                $inList = $false
            }
            

            if ($trimLine -match '^\d+\.\s+(.+)$') {
                if ($inList) { $htmlLines.Add("</ul>"); $inList = $false }
                if (-not $inOrderedList) { $htmlLines.Add("<ol>"); $inOrderedList = $true }
                $htmlLines.Add("<li>$($matches[1])</li>")
                continue
            } elseif ($inOrderedList) {
                $htmlLines.Add("</ol>")
                $inOrderedList = $false
            }


            if ($trimLine -match '^(#{1,6})\s+(.+)$') {
                $level = $matches[1].Length
                $htmlLines.Add("<h$level>$($matches[2])</h$level>")
                continue
            }


            if ($trimLine -match '^(&gt;|>)[\s]+(.+)$') {
                $htmlLines.Add("<blockquote>$($matches[2])</blockquote>")
                continue
            }


            if ($trimLine -match '^(\*{3,}|-{3,}|_{3,})$') {
                $htmlLines.Add("<hr />")
                continue
            }

            if ($trimLine.Length -gt 0) {
                $htmlLines.Add("$line<br>")
            } else {
                $htmlLines.Add("<br>")
            }
        }
        
        if ($inTable) { $htmlLines.Add("</table>") }
        if ($inList) { $htmlLines.Add("</ul>") }
        if ($inOrderedList) { $htmlLines.Add("</ol>") }

        $finalHtml = $htmlLines -join "`n"


        



        $codePlaceholderPrefix = "##WIKICODEBLOCK" 
        $codeMap = @{}
        $codeId = 0
        

        $match = [regex]::Matches($finalHtml, '`([^`]+)`')
        

        for ($i = $match.Count - 1; $i -ge 0; $i--) {
            $m = $match[$i]
            $key = "${codePlaceholderPrefix}${codeId}##"
            $codeMap[$key] = $m.Groups[1].Value
            

            $pre = $finalHtml.Substring(0, $m.Index)
            $post = $finalHtml.Substring($m.Index + $m.Length)
            $finalHtml = $pre + $key + $post
            
            $codeId++
        }



        $finalHtml = $finalHtml -replace '\*\*(.+?)\*\*', '<b>$1</b>'
        

        $finalHtml = $finalHtml -replace '\*(.+?)\*', '<i>$1</i>'
        

        $finalHtml = $finalHtml -replace '~~(.+?)~~', '<strike>$1</strike>'


        $finalHtml = $finalHtml -replace '\[\[(.+?)\]\]', '<a href="http://wiki/$1">$1</a>'
        $finalHtml = $finalHtml -replace '\[([^\]]+)\]\(wiki:([^)]+)\)', '<a href="http://wiki/$2">$1</a>'
        $finalHtml = $finalHtml -replace '\[([^\]]+)\]\(([^)]+)\)', '<a href="$2">$1</a>'


        $finalHtml = $finalHtml -replace '\[imgs\]([^<\s]+)', '<img src="$1" style="width: 25%;" />'
        $finalHtml = $finalHtml -replace '\[imgm\]([^<\s]+)', '<img src="$1" style="width: 50%;" />'
        $finalHtml = $finalHtml -replace '\[imgi\]([^<\s]+)', '<img src="$1" style="width: 32px;" />'
        $finalHtml = $finalHtml -replace '\[img\]([^<\s]+)', '<img src="$1" />'


        foreach ($key in $codeMap.Keys) {
            $content = $codeMap[$key]
            $finalHtml = $finalHtml.Replace($key, "<code>$content</code>")
        }

        return $finalHtml
    }.GetNewClosure()


    $RefreshContent = {
        $node = $treeView.SelectedNode
        if (-not $node) { return }
        

        if ($node.Text.Contains(':')) {
             $lblContentTitle.Text = $node.Text.Split(':')[0]
        } else {
             $lblContentTitle.Text = $node.Text
        }

        $data = $node.Tag
        

        $descVal = ""
        $urlVal = ""
        $noteVal = ""
        
        if ($data -is [System.Collections.IDictionary] -or $data -is [System.Management.Automation.PSCustomObject]) {
            if ($data.description) { $descVal = $data.description }
            if ($data.note) { $noteVal = $data.note }
            if ($data.url) { $urlVal = $data.url } elseif ($data.URL) { $urlVal = $data.URL }
        } elseif ($data -is [string]) {
            $descVal = $data
        }

        if ($chkEdit.Checked) {

            $txtDescription.Text = $descVal
            $txtDescription.Visible = $true
            $webDescription.Visible = $false
            $txtEditTitle.Text = $node.Name
            $txtEditUrl.Text = $urlVal
            $txtEditNote.Text = $noteVal
        } else {

            $displayText = $descVal
            
            if ($noteVal) { $displayText += "`n`nNote: $noteVal" }
            if ($urlVal) { $displayText += "`n`nLink: $urlVal" }
            
            $txtDescription.Text = $displayText
            $txtDescription.Visible = $false
            $webDescription.Visible = $true

            $htmlContent = & $ConvertMarkdownToHtml $displayText
            
            $fullHtml = @"
<html>
<head>
<style>
body { background-color: #0f1219; color: #ffffff; font-family: 'Segoe UI', sans-serif; font-size: 11pt; margin: 0; padding: 0; border: none; overflow: auto; scrollbar-base-color: #3a3f4b; scrollbar-track-color: #161a26; scrollbar-arrow-color: #ffffff; }
img { max-width: 100%; height: auto; display: block; margin: 10px 0; border-radius: 4px; }
a { color: #ff2e4c; text-decoration: none; }
table { border-collapse: collapse; width: 100%; margin: 10px 0; }
th, td { border: 1px solid #444; padding: 6px; text-align: left; }
th { background-color: #232838; color: #ff2e4c; }
tr:nth-child(even) { background-color: #161a26; }
code { background-color: #161a26; padding: 2px 4px; border-radius: 3px; font-family: Consolas, monospace; color: #e6db74; }
blockquote { border-left: 3px solid #ff2e4c; margin: 10px 0; padding-left: 10px; color: #a0a5b0; }
hr { border: 0; border-top: 1px solid #444; margin: 15px 0; }
</style>
</head>
<body>
$htmlContent
</body>
</html>
"@
            $webDescription.DocumentText = $fullHtml
            

            $hasLink = -not [string]::IsNullOrEmpty($urlVal)
            if (-not $hasLink -and $descVal -match '^https?://') { $hasLink = $true; $urlVal = $descVal }
            
            $btnOpenLink.Visible = $hasLink
            $btnCopyLink.Visible = $hasLink
            if ($hasLink) {
                $btnOpenLink.Tag = $urlVal
                $btnCopyLink.Tag = $urlVal
            }
            $btnPanel.Visible = $true 
        }
    }.GetNewClosure()


    $chkViewMode.Add_CheckedChanged({
        param($s, $e)
        try {

            $restorePath = New-Object System.Collections.Generic.List[string]
            $curr = $treeView.SelectedNode
            while ($curr) {

                $val = if (-not [string]::IsNullOrEmpty($curr.Name)) { $curr.Name } else { $curr.Text }
                $restorePath.Insert(0, $val)
                $curr = $curr.Parent
            }

            $isOnlineView = $s.Checked
            if ($isOnlineView) {
                $s.Text = "VIEW LOCAL"
                $s.BackColor = $Theme.AccentColor
                $s.ForeColor = "White"

                $form.Tag = $OnlineWikiData
                & $PopulateTree $null $OnlineWikiData $treeView
                
                $lblStatus.Text = "Source: Online (GitHub) [Read-Only]"
                $chkEdit.Enabled = $false
                $chkEdit.Checked = $false
            } else {
                $s.Text = "VIEW ONLINE"
                $s.BackColor = $Theme.Background
                $s.ForeColor = $Theme.SubTextColor

                $form.Tag = $LocalWikiData
                & $PopulateTree $null $LocalWikiData $treeView

                $lblStatus.Text = "Source: Offline (Local File)"
                $chkEdit.Enabled = ($null -ne $LocalWikiData)
            }
            

            $NavHistory.Clear()
            $btnBack.Enabled = $false


            $restored = $false
            if ($restorePath.Count -gt 0) {
                $nodes = $treeView.Nodes
                $target = $null
                
                foreach ($part in $restorePath) {
                    $found = $false
                    

                    $match = $nodes.Find($part, $false)
                    if ($match.Count -gt 0) {
                        $target = $match[0]; $nodes = $target.Nodes; $found = $true
                    }
                    

                    if (-not $found) {
                        foreach ($n in $nodes) {
                            if ($n.Text -eq $part) {
                                $target = $n; $nodes = $n.Nodes; $found = $true; break
                            }
                        }
                    }
                    if (-not $found) { $target = $null; break }
                }
                
                if ($target) { $treeView.SelectedNode = $target; $target.EnsureVisible(); $restored = $true }
            }

            if (-not $restored) {
                $lblContentTitle.Text = "Select an item"
                $txtDescription.Text = ""
                $btnOpenLink.Visible = $false
                $btnCopyLink.Visible = $false
            }
        } catch { Write-Verbose "ViewMode Error: $_" }
    }.GetNewClosure())

    $chkEdit.Add_CheckedChanged({
        $isEdit = $chkEdit.Checked
        $txtDescription.ReadOnly = -not $isEdit
        $btnSave.Visible = $isEdit
        $btnDelete.Visible = $isEdit
        $btnAdd.Visible = $isEdit
        
        if ($isEdit) {
            $chkEdit.BackColor = $Theme.AccentColor
            $chkEdit.ForeColor = "White"
            

            $btnPanel.Visible = $false
            $pnlEditFields.Visible = $true
            $pnlFormatting.Visible = $true
            $txtDescription.Top = 285
            $txtDescription.Height = ($contentPanel.Height - 295)
            $webDescription.Top = 285
            $webDescription.Height = ($contentPanel.Height - 295)
            if ([Custom.Native] -as [Type]) { [Custom.Native]::UseImmersiveDarkMode($txtDescription.Handle) }
        } else {
            $chkEdit.BackColor = $Theme.Background
            $chkEdit.ForeColor = $Theme.SubTextColor
            

            $pnlEditFields.Visible = $false
            $pnlFormatting.Visible = $false
            $txtDescription.Top = 180
            $txtDescription.Height = ($contentPanel.Height - 195)
            $webDescription.Top = 180
            $webDescription.Height = ($contentPanel.Height - 195)
        }
        
        & $RefreshContent
    }.GetNewClosure())

	$searchBox.Add_TextChanged({
			try
			{
				$txt = $this.Text
				if ($txt -eq 'Search...') { return }
				$data = $form.Tag
				& $PopulateTree $txt $data $treeView
			}
			catch {}
	}.GetNewClosure())

	$searchBox.Add_GotFocus({ 
			try { if ($this.Text -eq 'Search...') { $this.Text = ''; $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#ffffff') } } catch {}
	}.GetNewClosure())
    
	$searchBox.Add_LostFocus({ 
			try { if ($this.Text -eq '') { $this.Text = 'Search...'; $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#a0a5b0') } } catch {}
	}.GetNewClosure())

    $treeView.Add_NodeMouseClick({
        param($src, $e)
        if ($e.Button -eq 'Left') {
            $hit = $src.HitTest($e.Location)
            if ($hit.Location -ne [System.Windows.Forms.TreeViewHitTestLocations]::PlusMinus) {
                if (-not $e.Node.IsExpanded) { $e.Node.Expand() }
                

                & $RefreshContent
            }
        }
    }.GetNewClosure())

    $treeView.Add_AfterSelect({
        param($src, $e)
        try {
            $node = $e.Node
            $null = $data; $data = $node.Tag
            $null = $nodeName; $nodeName = $node.Name

            if ($NavHistory.Count -eq 0 -or $NavHistory.Peek() -ne $node) {
                 $NavHistory.Push($node)
                 if ($NavHistory.Count -gt 1) { $btnBack.Enabled = $true; $btnBack.ForeColor = $Theme.TextColor }
            }

            & $RefreshContent
        } catch {}
    }.GetNewClosure())

    $btnBack.Add_Click({
        if ($NavHistory.Count -gt 1) {
            $NavHistory.Pop() | Out-Null
            $prevNode = $NavHistory.Peek()
            $treeView.SelectedNode = $prevNode
            $treeView.Focus()
        }
        if ($NavHistory.Count -le 1) {
            $btnBack.Enabled = $false
            $btnBack.ForeColor = $Theme.SubTextColor
        }
    }.GetNewClosure())

    $btnBack.Add_MouseEnter({ if ($this.Enabled) { $this.ForeColor = $Theme.AccentColor } }.GetNewClosure())
    $btnBack.Add_MouseLeave({ if ($this.Enabled) { $this.ForeColor = $Theme.TextColor } }.GetNewClosure())

    $btnOpenLink.Add_Click({ try { if ($this.Tag) { [System.Diagnostics.Process]::Start($this.Tag) | Out-Null } } catch {} }.GetNewClosure())

    $btnCopyLink.Add_Click({
        try {
            if ($this.Tag) {
                [System.Windows.Forms.Clipboard]::SetText($this.Tag)
                $oldText = $this.Text
                $this.Text = "COPIED!"
                $this.ForeColor = $Theme.SuccessColor
                
                $t = New-Object System.Windows.Forms.Timer
                $Disposables.Add($t)
                $t.Interval = 1500
                $t.Add_Tick({
                    $this.Stop()
                    $btnCopyLink.Text = $oldText
                    $btnCopyLink.ForeColor = $Theme.SubTextColor
                    $Disposables.Remove($this)
                    $this.Dispose()
                })
                $t.Start()
            }
        } catch {}
    }.GetNewClosure())

    $btnSave.Add_Click({
        try {
            $node = $treeView.SelectedNode
            if (-not $node) { return }
            
            $newText = $txtDescription.Text
            $newTitle = $txtEditTitle.Text
            $newUrl = $txtEditUrl.Text
            $newNote = $txtEditNote.Text

            $meta = $node.WikiMeta 
            if ($null -eq $meta) { return }
            
            $parent = $meta.Parent
            $key = $meta.Key
            

            if (-not [string]::IsNullOrWhiteSpace($newTitle) -and $newTitle -ne $key) {

                $exists = $false
                if ($parent -is [System.Collections.IDictionary]) { $exists = $parent.Contains($newTitle) }
                elseif ($parent.PSObject) { $exists = $null -ne $parent.PSObject.Properties[$newTitle] }
                
                if ($exists) {
                    if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Rename Error" -Message "A key with the name '$newTitle' already exists." -Type "Warning" }
                    else { [System.Windows.Forms.MessageBox]::Show("A key with the name '$newTitle' already exists.", "Rename Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null }
                    return
                }


                if ($parent -is [System.Collections.IDictionary]) {
                    $keys = [System.Collections.ArrayList]@($parent.Keys)
                    $temp = [ordered]@{}
                    foreach ($k in $keys) {
                        if ($k -eq $key) { $temp[$newTitle] = $parent[$key] }
                        else { $temp[$k] = $parent[$k] }
                    }
                    $parent.Clear()
                    foreach ($k in $temp.Keys) { $parent[$k] = $temp[$k] }
                }
                else {

                    $props = $parent.PSObject.Properties | Select-Object Name, Value
                    $temp = [ordered]@{}
                    foreach ($p in $props) {
                        if ($p.Name -eq $key) { $temp[$newTitle] = $p.Value }
                        else { $temp[$p.Name] = $p.Value }
                    }
                    $props | ForEach-Object { $parent.PSObject.Properties.Remove($_.Name) }
                    foreach ($k in $temp.Keys) {
                        $parent | Add-Member -MemberType NoteProperty -Name $k -Value $temp[$k] -Force
                    }
                }
                
                $key = $newTitle
                $meta.Key = $newTitle
                $node.Name = $newTitle
                $node.Text = if ($node.Tag -is [string]) { "$newTitle`: $($node.Tag)" } else { $newTitle }
            }

            $isSimple = ($node.Tag -is [string])
            $needsComplex = (-not [string]::IsNullOrWhiteSpace($newUrl) -or -not [string]::IsNullOrWhiteSpace($newNote))
            

            if ($isSimple -and -not $needsComplex) {

                $val = $newText
                if ($parent -is [System.Collections.IList]) { $parent[$key] = $val }
                elseif ($parent -is [System.Collections.IDictionary]) { $parent[$key] = $val }
                else { $parent.$key = $val }
                $node.Tag = $val
                if ($node.Text -match "^$([regex]::Escape($key)):") { $node.Text = "$key`: $val" }
            }
            else {

                $obj = $node.Tag
                if ($isSimple) {
                    $obj = [ordered]@{ description = $newText }
                    if ($parent -is [System.Collections.IList]) { $parent[$key] = $obj }
                    elseif ($parent -is [System.Collections.IDictionary]) { $parent[$key] = $obj }
                    else { $parent.$key = $obj }
                    $node.Tag = $obj
                }
                

                if ($obj -is [System.Collections.IDictionary]) {
                    $obj['description'] = $newText
                    if (-not [string]::IsNullOrWhiteSpace($newUrl)) { $obj['url'] = $newUrl } elseif ($obj.Contains('url')) { $obj.Remove('url') }
                    if (-not [string]::IsNullOrWhiteSpace($newNote)) { $obj['note'] = $newNote } elseif ($obj.Contains('note')) { $obj.Remove('note') }
                }
                else {
                    $obj | Add-Member -MemberType NoteProperty -Name "description" -Value $newText -Force
                    
                    if (-not [string]::IsNullOrWhiteSpace($newUrl)) { $obj | Add-Member -MemberType NoteProperty -Name "url" -Value $newUrl -Force }
                    elseif ($obj.PSObject.Properties['url']) { $obj.PSObject.Properties.Remove('url') }

                    if (-not [string]::IsNullOrWhiteSpace($newNote)) { $obj | Add-Member -MemberType NoteProperty -Name "note" -Value $newNote -Force }
                    elseif ($obj.PSObject.Properties['note']) { $obj.PSObject.Properties.Remove('note') }
                }
            }
            

            $currentForm = $this.FindForm()
            $json = $currentForm.Tag | ConvertTo-Json -Depth 20
            $json | Set-Content -Path $LocalWikiPath -Encoding UTF8 -Force
            
            if (Get-Command ShowToast -ErrorAction SilentlyContinue) {
                ShowToast -Title "Wiki Saved" -Message "Changes saved successfully." -Type "Info" -TimeoutSeconds 2
            }
            else { [System.Windows.Forms.MessageBox]::Show("Wiki saved successfully!") | Out-Null }
            
        } catch { }
    }.GetNewClosure())

    $btnDelete.Add_Click({
        try {
            $node = $treeView.SelectedNode
            if (-not $node) { return }
            
            $meta = $node.WikiMeta
            if ($null -eq $meta) { return }
            
            $siblingCount = if ($node.Parent) { $node.Parent.Nodes.Count } else { $treeView.Nodes.Count }
            
            if ($siblingCount -le 1) {
                if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Action Blocked" -Message "Cannot delete the last item. Please add another item first." -Type "Warning" }
                return
            }

            $nKey = "WikiDelete".GetHashCode()


            $targetPath = $LocalWikiPath
            $targetForm = $form

            $performDelete = {
                try {
                    if (Get-Command CloseToast -ErrorAction SilentlyContinue) { CloseToast -Key $nKey }

                    $parent = $meta.Parent
                    $key = $meta.Key
                    
                    if ($parent -is [System.Collections.IDictionary]) { $parent.Remove($key) }
                    elseif ($parent -is [System.Collections.IList]) { $parent.RemoveAt([int]$key) }
                    else { $parent.PSObject.Properties.Remove([string]$key) }
                    
                    $node.Remove()
                    
                    $json = $targetForm.Tag | ConvertTo-Json -Depth 20
                    $json | Set-Content -Path $targetPath -Encoding UTF8 -Force
                    
                    if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Deleted" -Message "Item deleted successfully." -Type "Info" }
                    else { [System.Windows.Forms.MessageBox]::Show("Item deleted successfully.") | Out-Null }
                } catch {}
            }.GetNewClosure()

            if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue) {
                $btns = @( @{ Text = "Yes"; Action = $performDelete }, @{ Text = "No"; Action = { CloseToast -Key $nKey }.GetNewClosure() } )
                ShowInteractiveNotification -Title "Confirm Delete" -Message "Are you sure you want to delete '$($node.Text)'?" -Buttons $btns -Type "Warning" -Key $nKey -TimeoutSeconds 10
            } else {
                if ([System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete '$($node.Text)'?", "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -eq 'Yes') { & $performDelete }
            }
        } catch { 
            if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Error" -Message "Error: $_" -Type "Error" }
            else { }
        }
    }.GetNewClosure())

    $btnAdd.Add_Click({
        try {
            $cms = New-Object System.Windows.Forms.ContextMenuStrip
            if ('Custom.DarkRenderer' -as [Type]) { $cms.Renderer = New-Object Custom.DarkRenderer }
            
            $tv = $treeView
            $node = $tv.SelectedNode
            $mainForm = $tv.FindForm()
            

            $askName = {
                return "New"
            }


            $itemSibling = $cms.Items.Add("Add Sibling")
            $itemSibling.Add_Click({
                try {
                    $name = $askName
                    if ([string]::IsNullOrWhiteSpace($name)) { return }
                    
                    $null = $parentData; $parentData = $null
                    $collection = $null
                    
                    if ($node) {
                        $meta = $node.WikiMeta
                        if ($meta -and $meta.Parent) { $collection = $meta.Parent }
                    } else {
                        if ($mainForm) { $collection = $mainForm.Tag }
                    }
                    
                    if ($null -eq $collection) { return }
                    
                    $exists = $false
                    if ($collection -is [System.Collections.IDictionary]) { $exists = $collection.Contains($name) }
                    else { $exists = $null -ne $collection.PSObject.Properties[$name] }

                    if ($exists) { 
                        if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Add Error" -Message "Item '$name' already exists." -Type "Warning" }
                        else { [System.Windows.Forms.MessageBox]::Show("Item '$name' already exists.") | Out-Null }
                        return 
                    }
                    
                    $newItem = "New Description"
                    if ($collection -is [System.Collections.IDictionary]) { $collection[$name] = $newItem }
                    else { $collection | Add-Member -MemberType NoteProperty -Name $name -Value $newItem -Force }
                    
                    $pNode = if ($node) { $node.Parent } else { $null }
                    $newNode = New-Object System.Windows.Forms.TreeNode($name)
                    $newNode.Name = $name
                    $newNode.Text = "$name"
                    $newNode.Tag = $newItem
                    $newNode | Add-Member -MemberType NoteProperty -Name "WikiMeta" -Value @{ Parent = $collection; Key = $name } -Force
                    
                    if ($pNode) { $pNode.Nodes.Add($newNode) | Out-Null } else { $tv.Nodes.Add($newNode) | Out-Null }
                    $tv.SelectedNode = $newNode
                    $btnSave.PerformClick()
                } catch {}
            }.GetNewClosure())


            $itemChild = $cms.Items.Add("Add Child")
            if (-not $node) { $itemChild.Enabled = $false }
            $itemChild.Add_Click({
                try {
                    $name = &$askName
                    if ([string]::IsNullOrWhiteSpace($name)) { return }
                    
                    $meta = $node.WikiMeta
                    $parentOfNode = $meta.Parent
                    $keyOfNode = $meta.Key
                    $currentData = $node.Tag
                    
                    $container = $currentData
                    

                    if ($currentData -is [string]) {
                        $container = [ordered]@{ description = $currentData }
                        if ($parentOfNode -is [System.Collections.IDictionary]) { $parentOfNode[$keyOfNode] = $container }
                        else { $parentOfNode.$keyOfNode = $container }
                        
                        $node.Tag = $container
                        $node.Text = $keyOfNode 
                        $node.WikiMeta.Parent = $parentOfNode
                    }
                    
                    $exists = $false
                    if ($container -is [System.Collections.IDictionary]) { $exists = $container.Contains($name) }
                    else { $exists = $null -ne $container.PSObject.Properties[$name] }

                    if ($exists) { 
                        if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Add Error" -Message "Item '$name' already exists." -Type "Warning" }
                        else { [System.Windows.Forms.MessageBox]::Show("Item '$name' already exists.") | Out-Null }
                        return 
                    }
                    
                    $newItem = "New Description"
                    if ($container -is [System.Collections.IDictionary]) { $container[$name] = $newItem }
                    else { $container | Add-Member -MemberType NoteProperty -Name $name -Value $newItem -Force }
                    
                    $newNode = New-Object System.Windows.Forms.TreeNode($name)
                    $newNode.Name = $name
                    $newNode.Text = "$name"
                    $newNode.Tag = $newItem
                    $newNode | Add-Member -MemberType NoteProperty -Name "WikiMeta" -Value @{ Parent = $container; Key = $name } -Force
                    
                    $node.Nodes.Add($newNode) | Out-Null
                    $node.Expand()
                    $tv.SelectedNode = $newNode
                    $btnSave.PerformClick()
                } catch { }
            }.GetNewClosure())
            

            $itemRoot = $cms.Items.Add("Add Root Node")
            $itemRoot.Add_Click({
                try {
                    $name = &$askName
                    if ([string]::IsNullOrWhiteSpace($name)) { return }
                    
                    $collection = if ($mainForm) { $mainForm.Tag } else { $form.Tag }
                    if ($null -eq $collection) {
                        $collection = [ordered]@{}
                        if ($mainForm) { $mainForm.Tag = $collection }
                    }
                    
                    $exists = $false
                    if ($collection -is [System.Collections.IDictionary]) { $exists = $collection.Contains($name) }
                    else { $exists = $null -ne $collection.PSObject.Properties[$name] }

                    if ($exists) { 
                        if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Add Error" -Message "Item '$name' already exists." -Type "Warning" }
                        else { [System.Windows.Forms.MessageBox]::Show("Item '$name' already exists.") | Out-Null }
                        return 
                    }
                    
                    $newItem = "New Description"
                    if ($collection -is [System.Collections.IDictionary]) { $collection[$name] = $newItem }
                    else { $collection | Add-Member -MemberType NoteProperty -Name $name -Value $newItem -Force }
                    
                    $newNode = New-Object System.Windows.Forms.TreeNode($name)
                    $newNode.Name = $name
                    $newNode.Text = "$name"
                    $newNode.Tag = $newItem
                    $newNode | Add-Member -MemberType NoteProperty -Name "WikiMeta" -Value @{ Parent = $collection; Key = $name } -Force
                    
                    $tv.Nodes.Add($newNode) | Out-Null
                    $tv.SelectedNode = $newNode
                    $btnSave.PerformClick()
                } catch {}
            }.GetNewClosure())

            $cms.Items.Add("-") | Out-Null


            $itemMoveUp = $cms.Items.Add("Move Up")
            if (-not $node -or -not $node.PrevNode) { $itemMoveUp.Enabled = $false }
            $itemMoveUp.Add_Click({
                try {
                    $prevNode = $node.PrevNode
                    $meta = $node.WikiMeta
                    $parent = $meta.Parent
                    if ($null -eq $parent) { return }
                    $key = $meta.Key
                    

                    if ($parent -is [System.Collections.IList]) {
                        $idx = [int]$key
                        $prevIdx = $idx - 1
                        $temp = $parent[$idx]; $parent[$idx] = $parent[$prevIdx]; $parent[$prevIdx] = $temp
                        $node.WikiMeta.Key = $prevIdx; $prevNode.WikiMeta.Key = $idx
                    } else {
                        $prevKey = $prevNode.WikiMeta.Key
                        $keys = if ($parent -is [System.Collections.IDictionary]) { [System.Collections.ArrayList]@($parent.Keys) } else { [System.Collections.ArrayList]@($parent.PSObject.Properties.Name) }
                        $idx = $keys.IndexOf($key); $prevIdx = $keys.IndexOf($prevKey)
                        $keys[$idx] = $prevKey; $keys[$prevIdx] = $key
                        
                        if ($parent -is [System.Collections.IDictionary]) {
                            $tempData = [ordered]@{}; foreach ($k in $keys) { $tempData[$k] = $parent[$k] }; $parent.Clear(); foreach ($k in $tempData.Keys) { $parent[$k] = $tempData[$k] }
                        } else {
                            $tempData = [ordered]@{}; foreach ($k in $keys) { $tempData[$k] = $parent.$k }; $keys | ForEach-Object { $parent.PSObject.Properties.Remove($_) }; foreach ($k in $tempData.Keys) { $parent | Add-Member -MemberType NoteProperty -Name $k -Value $tempData[$k] -Force }
                        }
                    }
                    

                    $pNode = $node.Parent; $nodesColl = if ($pNode) { $pNode.Nodes } else { $tv.Nodes }
                    $index = $node.Index; $nodesColl.RemoveAt($index); $nodesColl.Insert($index - 1, $node); $tv.SelectedNode = $node
                    $btnSave.PerformClick()
                } catch { if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Move Error" -Message "Error moving up: $_" -Type "Error" } }
            }.GetNewClosure())


            $itemMoveDown = $cms.Items.Add("Move Down")
            if (-not $node -or -not $node.NextNode) { $itemMoveDown.Enabled = $false }
            $itemMoveDown.Add_Click({
                try {
                    $nextNode = $node.NextNode
                    $meta = $node.WikiMeta
                    $parent = $meta.Parent
                    if ($null -eq $parent) { return }
                    $key = $meta.Key
                    

                    if ($parent -is [System.Collections.IList]) {
                        $idx = [int]$key
                        $nextIdx = $idx + 1
                        $temp = $parent[$idx]; $parent[$idx] = $parent[$nextIdx]; $parent[$nextIdx] = $temp
                        $node.WikiMeta.Key = $nextIdx; $nextNode.WikiMeta.Key = $idx
                    } else {
                        $nextKey = $nextNode.WikiMeta.Key
                        $keys = if ($parent -is [System.Collections.IDictionary]) { [System.Collections.ArrayList]@($parent.Keys) } else { [System.Collections.ArrayList]@($parent.PSObject.Properties.Name) }
                        $idx = $keys.IndexOf($key); $nextIdx = $keys.IndexOf($nextKey)
                        $keys[$idx] = $nextKey; $keys[$nextIdx] = $key
                        
                        if ($parent -is [System.Collections.IDictionary]) {
                            $tempData = [ordered]@{}; foreach ($k in $keys) { $tempData[$k] = $parent[$k] }; $parent.Clear(); foreach ($k in $tempData.Keys) { $parent[$k] = $tempData[$k] }
                        } else {
                            $tempData = [ordered]@{}; foreach ($k in $keys) { $tempData[$k] = $parent.$k }; $keys | ForEach-Object { $parent.PSObject.Properties.Remove($_) }; foreach ($k in $tempData.Keys) { $parent | Add-Member -MemberType NoteProperty -Name $k -Value $tempData[$k] -Force }
                        }
                    }
                    

                    $pNode = $node.Parent; $nodesColl = if ($pNode) { $pNode.Nodes } else { $tv.Nodes }
                    $index = $node.Index; $nodesColl.RemoveAt($index); $nodesColl.Insert($index + 1, $node); $tv.SelectedNode = $node
                    $btnSave.PerformClick()
                } catch { if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Move Error" -Message "Error moving down: $_" -Type "Error" } }
            }.GetNewClosure())

            $cms.Items.Add("-") | Out-Null

            $cms.Show($btnAdd, 0, $btnAdd.Height)
        } catch { 
            if (Get-Command ShowToast -ErrorAction SilentlyContinue) { ShowToast -Title "Error" -Message "Error showing menu: $_" -Type "Error" }
            else { [System.Windows.Forms.MessageBox]::Show("Error showing menu: $_") | Out-Null }
        }
    }.GetNewClosure())


	& $PopulateTree $null $WikiData $treeView


	$chkEdit.Enabled = ($null -ne $LocalWikiData)
	$chkViewMode.Enabled = ($null -ne $OnlineWikiData -and $null -ne $LocalWikiData)
	if (-not $chkViewMode.Enabled) {
		$chkViewMode.Text = "N/A"
	}


	if ('Custom.DarkMessageBox' -as [Type])
	{
		if ($global:DashboardConfig.UI.MainForm -and $form.StartPosition -ne 'Manual')
		{
			$form.Owner = $global:DashboardConfig.UI.MainForm
			$form.StartPosition = 'CenterParent'
		}
	}

	$SetupResize = {
		param($f)
		$gripSize = 6
		
		$pBot = New-Object System.Windows.Forms.Panel
		$pBot.Dock = 'Bottom'; $pBot.Height = $gripSize; $pBot.Cursor = 'SizeNS'; $pBot.BackColor = [System.Drawing.Color]::Transparent
		$pBot.Add_MouseDown({ param($s, $e) if ($e.Button -eq 'Left') {
			[Custom.Native]::ReleaseCapture(); $frm = $s.FindForm()
			if ($frm) {
				$pt = $s.PointToClient([System.Windows.Forms.Cursor]::Position); $mode = 15
				if ($pt.X -lt 15) { $mode = 16 } elseif ($pt.X -gt $s.Width - 15) { $mode = 17 }
				[Custom.Native]::SendMessage($frm.Handle, 0xA1, $mode, 0)
			}
		} }.GetNewClosure())
		$pBot.Add_MouseMove({ param($s, $e)
			if ($e.X -lt 15) { $s.Cursor = 'SizeNESW' } elseif ($e.X -gt $s.Width - 15) { $s.Cursor = 'SizeNWSE' } else { $s.Cursor = 'SizeNS' }
		}.GetNewClosure())
		$f.Controls.Add($pBot); $pBot.SendToBack()

		$pTop = New-Object System.Windows.Forms.Panel
		$pTop.Dock = 'Top'; $pTop.Height = $gripSize; $pTop.Cursor = 'SizeNS'; $pTop.BackColor = [System.Drawing.Color]::Transparent
		$pTop.Add_MouseDown({ param($s, $e) if ($e.Button -eq 'Left') {
			[Custom.Native]::ReleaseCapture(); $frm = $s.FindForm()
			if ($frm) {
				$pt = $s.PointToClient([System.Windows.Forms.Cursor]::Position); $mode = 12
				if ($pt.X -lt 15) { $mode = 13 } elseif ($pt.X -gt $s.Width - 15) { $mode = 14 }
				[Custom.Native]::SendMessage($frm.Handle, 0xA1, $mode, 0)
			}
		} }.GetNewClosure())
		$pTop.Add_MouseMove({ param($s, $e)
			if ($e.X -lt 15) { $s.Cursor = 'SizeNWSE' } elseif ($e.X -gt $s.Width - 15) { $s.Cursor = 'SizeNESW' } else { $s.Cursor = 'SizeNS' }
		}.GetNewClosure())
		$f.Controls.Add($pTop); $pTop.SendToBack()

		$pLeft = New-Object System.Windows.Forms.Panel
		$pLeft.Dock = 'Left'; $pLeft.Width = $gripSize; $pLeft.Cursor = 'SizeWE'; $pLeft.BackColor = [System.Drawing.Color]::Transparent
		$pLeft.Add_MouseDown({ param($s, $e) if ($e.Button -eq 'Left') {
			[Custom.Native]::ReleaseCapture(); $frm = $s.FindForm()
			if ($frm) { [Custom.Native]::SendMessage($frm.Handle, 0xA1, 10, 0) }
		} }.GetNewClosure())
		$f.Controls.Add($pLeft); $pLeft.SendToBack()

		$pRight = New-Object System.Windows.Forms.Panel
		$pRight.Dock = 'Right'; $pRight.Width = $gripSize; $pRight.Cursor = 'SizeWE'; $pRight.BackColor = [System.Drawing.Color]::Transparent
		$pRight.Add_MouseDown({ param($s, $e) if ($e.Button -eq 'Left') {
			[Custom.Native]::ReleaseCapture(); $frm = $s.FindForm()
			if ($frm) { [Custom.Native]::SendMessage($frm.Handle, 0xA1, 11, 0) }
		} }.GetNewClosure())
		$f.Controls.Add($pRight); $pRight.SendToBack()
	}
	& $SetupResize $form


	if ($global:DashboardConfig.Config['WindowPosition'])
	{
		$wp = $global:DashboardConfig.Config['WindowPosition']
		if ($wp.Contains('WikiFormWidth') -and $wp.Contains('WikiFormHeight'))
		{
			$form.Width = [int]$wp.WikiFormWidth
			$form.Height = [int]$wp.WikiFormHeight
		}
		if ($wp.Contains('WikiFormX') -and $wp.Contains('WikiFormY'))
		{
			$wx = [int]$wp.WikiFormX; $wy = [int]$wp.WikiFormY
			$wRect = [System.Drawing.Rectangle]::new($wx, $wy, $form.Width, 20)
			if ([System.Windows.Forms.Screen]::AllScreens.WorkingArea.IntersectsWith($wRect) -contains $true)
			{
				$form.StartPosition = 'Manual'
				$form.Location = [System.Drawing.Point]::new($wx, $wy)
			}
		}
	}


	$txtDescription.Height = [Math]::Max(100, ($contentPanel.Height - 190))
	$webDescription.Height = [Math]::Max(100, ($contentPanel.Height - 190))

	$form.Add_FormClosed({
			if ($this.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal)
			{
				if (-not $global:DashboardConfig.Config.Contains('WindowPosition'))
				{ 
					$global:DashboardConfig.Config['WindowPosition'] = [ordered]@{} 
				}
				$global:DashboardConfig.Config['WindowPosition']['WikiFormX'] = $this.Location.X
				$global:DashboardConfig.Config['WindowPosition']['WikiFormY'] = $this.Location.Y
				$global:DashboardConfig.Config['WindowPosition']['WikiFormWidth'] = $this.Width
				$global:DashboardConfig.Config['WindowPosition']['WikiFormHeight'] = $this.Height
            
				if (Get-Command WriteConfig -ErrorAction SilentlyContinue) { WriteConfig }
			}

			try {
				if ($webDescription) {
					if (-not $webDescription.IsDisposed) {
						# Explicitly clear the document content and disable navigation
						# This is a common and effective workaround for WebBrowser opening external windows on dispose
						$webDescription.DocumentText = "" 
						$webDescription.AllowNavigation = $false
						$webDescription.Stop()
					}
					$webDescription.Dispose()
				}
			} catch {}

			try {
				foreach ($d in $Disposables.ToArray()) { if ($d) { $d.Dispose() } }
			} catch {}

            if ($global:DashboardConfig.Resources.ContainsKey('WikiForm')) { $global:DashboardConfig.Resources.Remove('WikiForm') }
            if ($global:DashboardConfig.Resources.ContainsKey('WikiResources')) { $global:DashboardConfig.Resources.Remove('WikiResources') }

			$this.Dispose() 
			[System.GC]::Collect()
			[System.GC]::WaitForPendingFinalizers()
		}.GetNewClosure())


	$form.Add_Shown({
		$capturedForm = $this
		$capturedStatus = $lblStatus
		$capturedViewMode = $chkViewMode
		$capturedTheme = $Theme
		$capturedLocalData = $LocalWikiData
		$capturedLocalPath = $LocalWikiPath
		$capturedTree = $treeView
		$capturedPopulate = $PopulateTree

		$scriptBlock = {
			param($url)
			try {
				$content = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 5 -ErrorAction Stop
				if ($content -is [string]) { return $content | ConvertFrom-Json }
				return $content
			} catch { return $null }
		}

		$runspace = [runspacefactory]::CreateRunspace()
		$runspace.Open()
		$ps = [powershell]::Create().AddScript($scriptBlock).AddArgument($WikiJsonURL)
		$ps.Runspace = $runspace
		$Disposables.Add($ps)
		$Disposables.Add($runspace)

        $global:DashboardConfig.Resources['WikiResources'] = @{
            PowerShell = $ps
            Runspace   = $runspace
        }

		$asyncResult = $ps.BeginInvoke()

		$pollTimer = New-Object System.Windows.Forms.Timer
		$pollTimer.Interval = 100
		$pollTimer.Add_Tick({
			if ($asyncResult.IsCompleted) {
				$pollTimer.Stop()
				
				try {
					$OnlineWikiData = $ps.EndInvoke($asyncResult)
				} catch { 
                    $pollTimer.Dispose()
                    return 
                }

				if ($capturedForm.IsDisposed) { 
                    $pollTimer.Dispose()
                    return 
                }
				
				if ($OnlineWikiData) {
					$capturedForm.Text = "Entropia Wiki - Online (GitHub)"
					$capturedStatus.Text = "Source: Online (GitHub)"
					$capturedStatus.ForeColor = $capturedTheme.SuccessColor
					$capturedViewMode.Enabled = ($null -ne $capturedLocalData)
					if ($capturedViewMode.Enabled) {
						$capturedViewMode.Text = "VIEW ONLINE"
					}

					if (-not $capturedLocalData) {
						$capturedLocalData = $OnlineWikiData
						$capturedForm.Tag = $OnlineWikiData
						& $capturedPopulate $null $OnlineWikiData $capturedTree
					}

					if (-not (Test-Path $capturedLocalPath)) {
						try {
							($OnlineWikiData | ConvertTo-Json -Depth 20) | Set-Content -Path $capturedLocalPath -Encoding UTF8 -Force
						} catch { Write-Verbose "Failed to create local wiki.json: $_" }
					}
				}
                $pollTimer.Dispose()
			}
		}.GetNewClosure())
        
        $Disposables.Add($pollTimer)
        $pollTimer.Start()
	}.GetNewClosure())

	$form.Show() | Out-Null

}

Export-ModuleMember -Function *