function Show-Wiki
{
	[CmdletBinding()]
	param()

	$LocalWikiPath = Join-Path $PSScriptRoot 'wiki.json'
	
	$LocalWikiData = $null
	$null = $OnlineWikiData; $OnlineWikiData = $null
	$WikiState = @{
		LocalData = $null
		OnlineData = $null
		# Stores unsaved edits: Key = NodeFullPath, Value = @{ Text = "..."; Scroll = 0 }
		UnsavedBuffer = @{} 
		# Flat list of node names for autocomplete
		NodeList = @()
	}
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
	if (-not ([System.Management.Automation.PSTypeName]'WikiScrollNative').Type) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class WikiScrollNative {
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, int wMsg, int wParam, ref POINT lParam);
    
    [DllImport("user32.dll")]
    public static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);

    public const int EM_SETSCROLLPOS = 0x04DE;
    public const int EM_GETSCROLLPOS = 0x04DD;
    public const int EM_GETFIRSTVISIBLELINE = 0x00CE;
    public const int EM_GETLINECOUNT = 0x00BA;
    public const int EM_LINEFROMCHAR = 0x00C9;
    public const int EM_LINESCROLL = 0x00B6;
    public const int SB_VERT = 1;
}
"@
    }

	$Theme = @{
		Background   = [System.Drawing.ColorTranslator]::FromHtml('#0f1219')
		PanelColor   = [System.Drawing.ColorTranslator]::FromHtml('#232838')
		InputBack    = [System.Drawing.ColorTranslator]::FromHtml('#161a26')
		AccentColor  = [System.Drawing.ColorTranslator]::FromHtml('#ff2e4c')
		TextColor    = [System.Drawing.ColorTranslator]::FromHtml('#ffffff')
		SubTextColor = [System.Drawing.ColorTranslator]::FromHtml('#a0a5b0')
		SuccessColor = [System.Drawing.ColorTranslator]::FromHtml('#28a745')
        ListSelect   = [System.Drawing.ColorTranslator]::FromHtml('#3a3f4b')
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

    # Autocomplete ListBox
    $lstAutocomplete = New-Object System.Windows.Forms.ListBox
    $lstAutocomplete.Visible = $false
    $lstAutocomplete.Size = New-Object System.Drawing.Size(200, 150)
    $lstAutocomplete.BackColor = $Theme.PanelColor
    $lstAutocomplete.ForeColor = $Theme.TextColor
    $lstAutocomplete.BorderStyle = 'FixedSingle'
    $lstAutocomplete.Font = New-Object System.Drawing.Font("Consolas", 10)
    $contentPanel.Controls.Add($lstAutocomplete)
    $lstAutocomplete.BringToFront()

    $editPreviewSplitter = New-Object System.Windows.Forms.SplitContainer
    $editPreviewSplitter.Orientation = 'Horizontal'
    $editPreviewSplitter.Location = New-Object System.Drawing.Point(25, 180)
    $editPreviewSplitter.Size = New-Object System.Drawing.Size(1000, 480)
    $editPreviewSplitter.Anchor = 'Top, Bottom, Left, Right'
    $editPreviewSplitter.BackColor = $Theme.PanelColor
    $editPreviewSplitter.SplitterWidth = 8
    $editPreviewSplitter.Panel1Collapsed = $true
    $contentPanel.Controls.Add($editPreviewSplitter) | Out-Null

    $txtDescription = New-Object System.Windows.Forms.RichTextBox
    $txtDescription.Dock = 'Fill'
    $txtDescription.BackColor = $Theme.Background
    $txtDescription.ForeColor = $Theme.TextColor
    $txtDescription.BorderStyle = "None"
    $txtDescription.ReadOnly = $true
    $txtDescription.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $txtDescription.Text = "Select an item from the menu on the left to view details.`n`nUse the Search box to quickly find guides, items, or systems."
    $txtDescription.HideSelection = $false
    $editPreviewSplitter.Panel1.Controls.Add($txtDescription) | Out-Null
    if ([Custom.Native] -as [Type]) { [Custom.Native]::UseImmersiveDarkMode($txtDescription.Handle) }

    $webDescription = New-Object System.Windows.Forms.WebBrowser
    $webDescription.Dock = 'Fill'
    $webDescription.ScriptErrorsSuppressed = $true
    $webDescription.Visible = $true
    $webDescription.DocumentText = @"
<html><head><style>body{background-color:#0f1219;color:#ffffff;font-family:'Segoe UI',sans-serif;font-size:11pt;margin:0;padding:0;}</style></head><body>
<div style='padding:20px;'>Select an item from the menu on the left to view details.<br><br>Use the Search box to quickly find guides, items, or systems.</div>
</body></html>
"@
    $editPreviewSplitter.Panel2.Controls.Add($webDescription) | Out-Null

	$scrollSyncTimer = New-Object System.Windows.Forms.Timer
    $scrollSyncTimer.Interval = 300
    $Disposables.Add($scrollSyncTimer)
    
    $SyncState = @{
        IsScrolling = $false
        LastTxtScroll = -1
    }

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

    $btnSyncNode = New-Object System.Windows.Forms.Button
    $btnSyncNode.Text = "SYNC NODE"
    $btnSyncNode.Size = New-Object System.Drawing.Size(80, 30)
    $btnSyncNode.Anchor = 'Top, Right'
    $btnSyncNode.Location = New-Object System.Drawing.Point(($navToolbar.Width - 180), 5)
    $btnSyncNode.FlatStyle = "Flat"
    $btnSyncNode.FlatAppearance.BorderSize = 0
    $btnSyncNode.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#5bc0de")
    $btnSyncNode.ForeColor = "White"
    $btnSyncNode.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnSyncNode.Cursor = "Hand"
    $btnSyncNode.Visible = $false
    $ttSync = New-Object System.Windows.Forms.ToolTip
    $ttSync.SetToolTip($btnSyncNode, "Update only the selected item from online source")
    $navToolbar.Controls.Add($btnSyncNode) | Out-Null

    $btnSyncAll = New-Object System.Windows.Forms.Button
    $btnSyncAll.Text = "SYNC ALL"
    $btnSyncAll.Size = New-Object System.Drawing.Size(80, 30)
    $btnSyncAll.Anchor = 'Top, Right'
    $btnSyncAll.Location = New-Object System.Drawing.Point(($navToolbar.Width - 270), 5)
    $btnSyncAll.FlatStyle = "Flat"
    $btnSyncAll.FlatAppearance.BorderSize = 0
    $btnSyncAll.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#17a2b8") 
    $btnSyncAll.ForeColor = "White"
    $btnSyncAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnSyncAll.Cursor = "Hand"
    $btnSyncAll.Visible = $false
    $ttSync.SetToolTip($btnSyncAll, "Overwrite ALL local data with online data")
    $navToolbar.Controls.Add($btnSyncAll) | Out-Null

    $btnSmartSync = New-Object System.Windows.Forms.Button
    $btnSmartSync.Text = "SMART SYNC"
    $btnSmartSync.Size = New-Object System.Drawing.Size(90, 30)
    $btnSmartSync.Anchor = 'Top, Right'
    $btnSmartSync.Location = New-Object System.Drawing.Point(($navToolbar.Width - 370), 5)
    $btnSmartSync.FlatStyle = "Flat"
    $btnSmartSync.FlatAppearance.BorderSize = 0
    $btnSmartSync.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#20c997")
    $btnSmartSync.ForeColor = "White"
    $btnSmartSync.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnSmartSync.Cursor = "Hand"
    $btnSmartSync.Visible = $false
    $ttSync.SetToolTip($btnSmartSync, "Merge online data into local wiki (Adds missing nodes, keeps your changes)")
    $navToolbar.Controls.Add($btnSmartSync) | Out-Null

    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "DELETE"
    $btnDelete.Size = New-Object System.Drawing.Size(80, 30)
    $btnDelete.Anchor = 'Top, Right'
    $btnDelete.Location = New-Object System.Drawing.Point(($navToolbar.Width - 460), 5)
    $btnDelete.FlatStyle = "Flat"
    $btnDelete.FlatAppearance.BorderSize = 0
    $btnDelete.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#d9534f")
    $btnDelete.ForeColor = "White"
    $btnDelete.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnDelete.Cursor = "Hand"
    $btnDelete.Visible = $false
    $navToolbar.Controls.Add($btnDelete) | Out-Null

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "ADD"
    $btnAdd.Size = New-Object System.Drawing.Size(80, 30)
    $btnAdd.Anchor = 'Top, Right'
    $btnAdd.Location = New-Object System.Drawing.Point(($navToolbar.Width - 550), 5)
    $btnAdd.FlatStyle = "Flat"
    $btnAdd.FlatAppearance.BorderSize = 0
    $btnAdd.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078d7")
    $btnAdd.ForeColor = "White"
    $btnAdd.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnAdd.Cursor = "Hand"
    $btnAdd.Visible = $false
    $navToolbar.Controls.Add($btnAdd) | Out-Null

    $chkEdit = New-Object System.Windows.Forms.CheckBox
    $chkEdit.Name = "chkEdit"
    $chkEdit.Text = "EDIT"
    $chkEdit.Appearance = 'Button'
    $chkEdit.Size = New-Object System.Drawing.Size(80, 30)
    $chkEdit.Anchor = 'Top, Right'
    $chkEdit.Location = New-Object System.Drawing.Point(($navToolbar.Width - 640), 5)
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
    $chkViewMode.Location = New-Object System.Drawing.Point(($navToolbar.Width - 760), 5)
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

    # Helper to flatten node list for autocomplete
    $PopulateNodeList = {
        param($Data)
        $list = New-Object System.Collections.Generic.List[string]
        
        $recurse = {
            param($obj)
            $keys = & $getKeys $obj
            foreach ($k in $keys) {
                if ($k -in 'meta','description','url','URL') { continue }
                $list.Add($k)
                $val = if ($obj -is [System.Collections.IDictionary]) { $obj[$k] } else { $obj.$k }
                if ($val -is [System.Collections.IDictionary] -or $val -is [System.Management.Automation.PSCustomObject]) {
                    & $recurse $val
                }
            }
        }
        & $recurse $Data
        return $list
    }.GetNewClosure()

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

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            $trimLine = $line.Trim()
            $lineId = "line_$i"
            
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
                $rowHtml = "<tr id=""$lineId"">"
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
                $htmlLines.Add("<li id=""$lineId"">$($matches[1])</li>")
                continue
            } elseif ($inList) {
                $htmlLines.Add("</ul>")
                $inList = $false
            }
            
            if ($trimLine -match '^\d+\.\s+(.+)$') {
                if ($inList) { $htmlLines.Add("</ul>"); $inList = $false }
                if (-not $inOrderedList) { $htmlLines.Add("<ol>"); $inOrderedList = $true }
                $htmlLines.Add("<li id=""$lineId"">$($matches[1])</li>")
                continue
            } elseif ($inOrderedList) {
                $htmlLines.Add("</ol>")
                $inOrderedList = $false
            }

            if ($trimLine -match '^(#{1,6})\s+(.+)$') {
                $level = $matches[1].Length
                $htmlLines.Add("<h$level id=""$lineId"">$($matches[2])</h$level>")
                continue
            }

            if ($trimLine -match '^(&gt;|>)[\s]+(.+)$') {
                $htmlLines.Add("<blockquote id=""$lineId"">$($matches[2])</blockquote>")
                continue
            }

            if ($trimLine -match '^(\*{3,}|-{3,}|_{3,})$') {
                $htmlLines.Add("<hr id=""$lineId"" />")
                continue
            }

            if ($trimLine.Length -gt 0) {
                $htmlLines.Add("<a id=""$lineId""></a>$line<br>")
            } else {
                $htmlLines.Add("<a id=""$lineId""></a><br>")
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

        # Handle Video formats (MP4, WEBM) inside [img] tags
        $finalHtml = $finalHtml -replace '\[img\]\s*(\S+\.(?:mp4|webm))', '<video src="$1" autoplay loop muted playsinline style="max-width: 100%; border-radius: 4px;"></video>'

        # Handle standard Images
        $finalHtml = $finalHtml -replace '\[img\]\s*([^\s<]+)', '<img src="$1" />'

        $finalHtml = $finalHtml -replace '\[imgs\]([^<\s]+)', '<img src="$1" style="width: 25%;" />'
        $finalHtml = $finalHtml -replace '\[imgm\]([^<\s]+)', '<img src="$1" style="width: 50%;" />'
        $finalHtml = $finalHtml -replace '\[imgi\]([^<\s]+)', '<img src="$1" style="width: 32px;" />'

        foreach ($key in $codeMap.Keys) {
            $content = $codeMap[$key]
            $finalHtml = $finalHtml.Replace($key, "<code>$content</code>")
        }

        return $finalHtml
    }.GetNewClosure()

    $UpdateLineOffsets = {
        if ($txtDescription.IsDisposed) { return }
        $txt = $txtDescription.Text
        $lines = $txt -split "`n"
        $offsets = New-Object int[] ($lines.Count + 1)
        $runningTotal = 0
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $offsets[$i] = $runningTotal
            $runningTotal += $lines[$i].Length + 1
        }
        $offsets[$lines.Count] = $runningTotal
        $SyncState.LineCharOffsets = $offsets
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
        
        # Check for unsaved changes first
        if ($chkEdit.Checked -and $WikiState.UnsavedBuffer.ContainsKey($node.FullPath)) {
            $buffer = $WikiState.UnsavedBuffer[$node.FullPath]
            $descVal = $buffer.Text
            # We can retrieve other props from original data if needed, but text is priority
            if ($data -is [System.Collections.IDictionary] -or $data -is [System.Management.Automation.PSCustomObject]) {
                if ($data.url) { $urlVal = $data.url } elseif ($data.URL) { $urlVal = $data.URL }
                if ($data.note) { $noteVal = $data.note }
            }
        }
        else {
            if ($data -is [System.Collections.IDictionary] -or $data -is [System.Management.Automation.PSCustomObject]) {
                if ($data.description) { $descVal = $data.description }
                if ($data.note) { $noteVal = $data.note }
                if ($data.url) { $urlVal = $data.url } elseif ($data.URL) { $urlVal = $data.URL }
            } elseif ($data -is [string]) {
                $descVal = $data
            }
        }

        if ($chkEdit.Checked) {
            $txtDescription.Text = $descVal
            $txtDescription.Visible = $true
            $webDescription.Visible = $true # Visible in edit mode for live preview
            $txtEditTitle.Text = $node.Name
            $txtEditUrl.Text = $urlVal
            $txtEditNote.Text = $noteVal
            
            # Restore scroll position if available
            if ($WikiState.UnsavedBuffer.ContainsKey($node.FullPath)) {
                $scrollPos = $WikiState.UnsavedBuffer[$node.FullPath].Scroll
                if ($scrollPos -gt 0) {
                    $txtDescription.SelectionStart = $scrollPos
                    $txtDescription.ScrollToCaret()
                }
            }
            
            # Trigger initial preview render
            $htmlContent = & $ConvertMarkdownToHtml $descVal
            & $UpdateLineOffsets
            $fullHtml = @"
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<style>
body { background-color: #0f1219; color: #ffffff; font-family: 'Segoe UI', sans-serif; font-size: 11pt; margin: 0; padding: 0; border: none; scrollbar-base-color: #3a3f4b; scrollbar-track-color: #161a26; scrollbar-arrow-color: #ffffff; }
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

        } else {
            $displayText = $descVal
            
            if ($noteVal) { $displayText += "`n`nNote: $noteVal" }
            if ($urlVal) { $displayText += "`n`nLink: $urlVal" }
            
            $txtDescription.Text = $displayText
            $txtDescription.Visible = $false
            $webDescription.Visible = $true

            # Force sync reset when switching to view mode
            $SyncState.LastTxtScroll = -1

            $htmlContent = & $ConvertMarkdownToHtml $displayText
            
            # Added meta tag for IE Edge mode to support GIFs
            $fullHtml = @"
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<style>
body { background-color: #0f1219; color: #ffffff; font-family: 'Segoe UI', sans-serif; font-size: 11pt; margin: 0; padding: 0; border: none; scrollbar-base-color: #3a3f4b; scrollbar-track-color: #161a26; scrollbar-arrow-color: #ffffff; }
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

    $UpdateSyncButtonState = {
        $isLocalView = -not $chkViewMode.Checked
        $online = $WikiState['OnlineData']
        $hasOnlineData = $null -ne $online
        $hasNode = $null -ne $treeView.SelectedNode

        $btnSyncNode.Visible = $isLocalView -and $hasOnlineData
        $btnSyncNode.Enabled = $btnSyncNode.Visible -and $hasNode
        $btnSyncAll.Visible = $isLocalView -and $hasOnlineData
        $btnSmartSync.Visible = $isLocalView -and $hasOnlineData
    }.GetNewClosure()

    # --- Autocomplete Logic ---
    $txtDescription.Add_KeyUp({
        param($s, $e)
        
        # Hide if not relevant keys
        if ($e.KeyCode -in 'Up','Down','Enter','Tab','Escape') { return }

        $caret = $txtDescription.SelectionStart
        if ($caret -eq 0) { $lstAutocomplete.Visible = $false; return }

        # Get text up to caret
        $text = $txtDescription.Text.Substring(0, $caret)
        
        # Regex to find [[... or wiki:... at the end
        if ($text -match '(?<trigger>\[\[|wiki:)(?<query>[^\]\)]*)$') {
            $trigger = $matches['trigger']
            $query = $matches['query']
            
            # Filter nodes
            $matchesList = $WikiState.NodeList | Where-Object { $_ -like "*$query*" } | Select-Object -First 10
            
            if ($matchesList) {
                $lstAutocomplete.Items.Clear()
                foreach ($m in $matchesList) { $lstAutocomplete.Items.Add($m) | Out-Null }
                
                # Position ListBox
                $pt = $txtDescription.GetPositionFromCharIndex($caret)
                $pt.Y += 20 # Move below line
                
                # Adjust if out of bounds (basic)
                if ($pt.X + $lstAutocomplete.Width -gt $txtDescription.Width) { $pt.X = $txtDescription.Width - $lstAutocomplete.Width }
                
                $lstAutocomplete.Location = $pt
                $lstAutocomplete.Visible = $true
                $lstAutocomplete.Tag = @{ Trigger = $trigger; StartIndex = ($caret - $query.Length) }
                if ($lstAutocomplete.Items.Count -gt 0) { $lstAutocomplete.SelectedIndex = 0 }
            } else {
                $lstAutocomplete.Visible = $false
            }
        } else {
            $lstAutocomplete.Visible = $false
        }
    }.GetNewClosure())

    $txtDescription.Add_KeyDown({
        param($s, $e)
        if ($lstAutocomplete.Visible) {
            if ($e.KeyCode -eq 'Down') {
                $e.Handled = $true
                if ($lstAutocomplete.SelectedIndex -lt $lstAutocomplete.Items.Count - 1) {
                    $lstAutocomplete.SelectedIndex++
                }
            }
            elseif ($e.KeyCode -eq 'Up') {
                $e.Handled = $true
                if ($lstAutocomplete.SelectedIndex -gt 0) {
                    $lstAutocomplete.SelectedIndex--
                }
            }
            elseif ($e.KeyCode -in 'Tab','Enter') {
                $e.Handled = $true
                $e.SuppressKeyPress = $true # Prevent newline/tab in text
                
                $selected = $lstAutocomplete.SelectedItem
                if ($selected) {
                    $info = $lstAutocomplete.Tag
                    $start = $info.StartIndex
                    $len = $txtDescription.SelectionStart - $start
                    
                    $txtDescription.Select($start, $len)
                    $txtDescription.SelectedText = $selected
                    
                    # Close brackets if [[
                    if ($info.Trigger -eq '[[') {
                        $txtDescription.SelectedText = "]]"
                    } else {
                        $txtDescription.SelectedText = ")" # Close paren for wiki: link
                    }
                    
                    $lstAutocomplete.Visible = $false
                }
            }
            elseif ($e.KeyCode -eq 'Escape') {
                $lstAutocomplete.Visible = $false
                $e.Handled = $true
            }
        }
    }.GetNewClosure())
    
    $lstAutocomplete.Add_MouseClick({
        $selected = $lstAutocomplete.SelectedItem
        if ($selected) {
             $info = $lstAutocomplete.Tag
             $start = $info.StartIndex
             $len = $txtDescription.SelectionStart - $start
             
             $txtDescription.Select($start, $len)
             $txtDescription.SelectedText = $selected
             
             if ($info.Trigger -eq '[[') {
                 $txtDescription.SelectedText = "]]"
             } else {
                 $txtDescription.SelectedText = ")"
             }
             $lstAutocomplete.Visible = $false
             $txtDescription.Focus()
        }
    }.GetNewClosure())

    # --- Live Preview Logic ---
    
    $generalScrollTimer = New-Object System.Windows.Forms.Timer
    $generalScrollTimer.Interval = 50 # Faster tick for smoother sync
    $Disposables.Add($generalScrollTimer)

    # Helper: robustly find the element that controls scrolling
    $GetWebScrollable = {
        if (-not $webDescription.Document) { return $null }
        
        $html = $webDescription.Document.GetElementsByTagName("HTML")[0]
        $body = $webDescription.Document.Body
        
        # If one is scrolled, use it
        if ($html -and $html.ScrollTop -gt 0) { return $html }
        if ($body -and $body.ScrollTop -gt 0) { return $body }
        
        # Default to HTML in standards mode (which we use), fallback to Body
        if ($html) { return $html }
        return $body
    }.GetNewClosure()
    
    # 1. Editor -> Web Sync (Timer based)
    $generalScrollTimer.Add_Tick({
        if ($SyncState.IsScrolling) { return }
        if (-not $chkEdit.Checked) { return }
        if (-not $webDescription.Visible) { return }
        
        try {
            $firstLine = [WikiScrollNative]::SendMessage($txtDescription.Handle, [WikiScrollNative]::EM_GETFIRSTVISIBLELINE, 0, 0)
            
            if ($firstLine -ne $SyncState.LastTxtScroll) {
                $SyncState.LastTxtScroll = $firstLine
                
                # Lock to prevent the Web->Editor event from firing back
                $SyncState.IsScrolling = $true 
                
                if ($webDescription.Document) {
                    $targetEl = $null
                    # Search backwards from current line to find an anchor
                    for ($i = $firstLine; $i -ge 0; $i--) {
                        $el = $webDescription.Document.GetElementById("line_$i")
                        if ($el) {
                            $targetEl = $el
                            break
                        }
                    }
                    
                    if ($targetEl) {
                        $targetEl.ScrollIntoView($true)
                    }
                }
            }
        }
        catch {}
        finally {
             $SyncState.IsScrolling = $false
        }
    }.GetNewClosure())
    $generalScrollTimer.Start()

    # 2. Web -> Editor Sync (Event based)
    $AttachWebScroll = {
        if ($webDescription.Document -and $webDescription.Document.Window) {
            $webDescription.Document.Window.AttachEventHandler("onscroll", {
                if ($SyncState.IsScrolling) { return }
                if (-not $chkEdit.Checked) { return }
                
                try {
                    $SyncState.IsScrolling = $true
                    
                    # Use GetElementFromPoint to find which line is at the top
                    $pt = New-Object System.Drawing.Point(30, 20) # Offset slightly to hit content
                    $el = $webDescription.Document.GetElementFromPoint($pt)
                    
                    $foundLine = -1
                    while ($el) {
                        if ($el.Id -and $el.Id.StartsWith("line_")) {
                            $foundLine = [int]($el.Id.Substring(5))
                            break
                        }
                        $el = $el.Parent
                    }
                    
                    if ($foundLine -ge 0) {
                        $targetPhysicalLine = $foundLine
                        if ($SyncState.LineCharOffsets -and $foundLine -lt $SyncState.LineCharOffsets.Count) {
                            $charIdx = $SyncState.LineCharOffsets[$foundLine]
                            $targetPhysicalLine = [WikiScrollNative]::SendMessage($txtDescription.Handle, [WikiScrollNative]::EM_LINEFROMCHAR, $charIdx, 0)
                        }

                        $currentLine = [WikiScrollNative]::SendMessage($txtDescription.Handle, [WikiScrollNative]::EM_GETFIRSTVISIBLELINE, 0, 0)
                        $linesToScroll = $targetPhysicalLine - $currentLine

                        if ($linesToScroll -ne 0) {
                            [WikiScrollNative]::SendMessage($txtDescription.Handle, [WikiScrollNative]::EM_LINESCROLL, 0, $linesToScroll) | Out-Null
                        }
                        
                        # Update tracker so timer doesn't reverse it
                        $SyncState.LastTxtScroll = [WikiScrollNative]::SendMessage($txtDescription.Handle, [WikiScrollNative]::EM_GETFIRSTVISIBLELINE, 0, 0)
                    }
                }
                catch {}
                finally {
                    $SyncState.IsScrolling = $false
                }
            })
        }
    }.GetNewClosure()

    $webDescription.Add_DocumentCompleted({
        param($s, $e)
        & $AttachWebScroll
    }.GetNewClosure())

    # 3. Content Update (Debounced typing)
    $scrollSyncTimer.Add_Tick({
        $scrollSyncTimer.Stop()
        if ($chkEdit.Checked) {
            
            $htmlContent = & $ConvertMarkdownToHtml $txtDescription.Text
            & $UpdateLineOffsets
            
            # Pause sync while we mess with the DOM
            $SyncState.IsScrolling = $true 

            try {
                if ($webDescription.Document -and $webDescription.Document.Body) {
                    
                    # A. Capture Current Ratio (not pixels, in case height changes)
                    $scrollEl = & $GetWebScrollable
                    $prevRatio = 0.0
                    $hasScroll = $false

                    if ($scrollEl) {
                        $oldMax = $scrollEl.ScrollRectangle.Height - $scrollEl.ClientRectangle.Height
                        if ($oldMax -gt 0) {
                             $prevRatio = $scrollEl.ScrollTop / $oldMax
                             $hasScroll = $true
                        }
                    }

                    # B. Update InnerHtml (No Page Reload)
                    $webDescription.Document.Body.InnerHtml = $htmlContent
                    
                    # C. Restore Position by Ratio
                    if ($hasScroll) {
                        # Re-fetch element as DOM might have shifted
                        $scrollEl = & $GetWebScrollable
                        if ($scrollEl) {
                            $newMax = $scrollEl.ScrollRectangle.Height - $scrollEl.ClientRectangle.Height
                            if ($newMax -gt 0) {
                                $scrollEl.ScrollTop = [int]($prevRatio * $newMax)
                            }
                        }
                    }
                } 
                else {
                    # Initial Load - Add DOCTYPE for Standards Mode
                    $fullHtml = @"
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<style>
body { background-color: #0f1219; color: #ffffff; font-family: 'Segoe UI', sans-serif; font-size: 11pt; margin: 0; padding: 0; border: none; scrollbar-base-color: #3a3f4b; scrollbar-track-color: #161a26; scrollbar-arrow-color: #ffffff; }
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
                }
            } 
            catch {}
            finally {
                # Allow sync to resume after a short delay to let rendering finish
                $SyncState.IsScrolling = $false
            }
        }
    }.GetNewClosure())

    $txtDescription.Add_TextChanged({
        if ($chkEdit.Checked) {
            $scrollSyncTimer.Stop()
            $scrollSyncTimer.Start()
        }
    }.GetNewClosure())

    # --- End Autocomplete ---

    $btnSyncNode.Add_Click({
        try {
            $node = $treeView.SelectedNode
            if (-not $node) { return }

            $path = New-Object System.Collections.Generic.List[object]
            $curr = $node
            while ($curr) {
                if ($curr.WikiMeta) {
                    $path.Insert(0, $curr.WikiMeta.Key)
                } else {
                    if (Get-Command ShowToast -Ea SilentlyContinue) { ShowToast "Error" "Node is missing metadata." "Error" }
                    return
                }
                $curr = $curr.Parent
            }

            $onlineNodeData = $WikiState.OnlineData
            $found = $true
            foreach ($key in $path) {
                if ($null -eq $onlineNodeData) { $found = $false; break }

                $keyExists = $false
                if ($onlineNodeData -is [System.Collections.IDictionary]) {
                    $keyExists = $onlineNodeData.Contains($key)
                } elseif ($onlineNodeData -is [System.Collections.IList]) {
                    $idx = -1
                    if ([int]::TryParse($key, [ref]$idx)) {
                        $keyExists = $idx -ge 0 -and $idx -lt $onlineNodeData.Count
                    }
                } elseif ($onlineNodeData.PSObject) {
                    $keyExists = $null -ne $onlineNodeData.PSObject.Properties[$key]
                }

                if ($keyExists) {
                    if ($onlineNodeData -is [System.Collections.IList]) {
                        $onlineNodeData = $onlineNodeData[[int]$key]
                    } else {
                        $onlineNodeData = $onlineNodeData.$key
                    }
                } else {
                    $found = $false
                    break
                }
            }

            if (-not $found -or $null -eq $onlineNodeData) {
                if (Get-Command ShowToast -Ea SilentlyContinue) { ShowToast "Sync Error" "Could not find corresponding node in online data." "Warning" }
                return
            }

            $meta = $node.WikiMeta
            $parent = $meta.Parent
            $key = $meta.Key

            $onlineNodeDataClone = $onlineNodeData | ConvertTo-Json -Depth 20 | ConvertFrom-Json

            if ($parent -is [System.Collections.IDictionary]) {
                $parent[$key] = $onlineNodeDataClone
            } elseif ($parent -is [System.Collections.IList]) {
                $parent[[int]$key] = $onlineNodeDataClone
            } else {
                $parent | Add-Member -MemberType NoteProperty -Name $key -Value $onlineNodeDataClone -Force
            }
            
            $node.Tag = $onlineNodeDataClone
            $node.Nodes.Clear()

            $isComplex = ($onlineNodeDataClone -is [System.Management.Automation.PSCustomObject] -or $onlineNodeDataClone -is [System.Collections.IDictionary])
            $isArray   = ($onlineNodeDataClone -is [System.Array] -or $onlineNodeDataClone -is [System.Collections.ArrayList]) -and $onlineNodeDataClone -isnot [string]
            
            if ($isComplex) {
                $childKeys = & $getKeys $onlineNodeDataClone
                foreach ($childName in $childKeys) {
                    if ($childName -in 'meta','description','url','URL') { continue }
                    $childContent = if ($onlineNodeDataClone -is [System.Collections.IDictionary]) { $onlineNodeDataClone[$childName] } else { $onlineNodeDataClone.$childName }
                    & $AddWikiNode $AddWikiNode $node $childName $childContent $null $false $onlineNodeDataClone $childName
                }
            } elseif ($isArray) {
                $idx = 0
                foreach ($item in $onlineNodeDataClone) {
                    $name = if ($item -is [string]) { $item } else { "Item" }
                    & $AddWikiNode $AddWikiNode $node $name $item $null $false $onlineNodeDataClone $idx
                    $idx++
                }
            }

            $btnSave.PerformClick()
            & $RefreshContent
            
            if (Get-Command ShowToast -Ea SilentlyContinue) { ShowToast "Synced" "Node synced with online version." "Info" }
        } catch {
            if (Get-Command ShowToast -Ea SilentlyContinue) { ShowToast "Sync Error" "An error occurred during sync: $_" "Error" }
        }
    }.GetNewClosure())

	$btnSyncAll.Add_Click({
        try {
            $online = $WikiState['OnlineData']
            if ($null -eq $online) { return }

            $confirmKey = "WikiSyncAll".GetHashCode()
            
            $targetPath = $LocalWikiPath
            $targetForm = $form
            $targetState = $WikiState
            $targetTree = $treeView
            $targetPopulate = $PopulateTree
            $targetViewMode = $chkViewMode

            $performSyncAll = {
                try {
                    if (Get-Command CloseToast -ErrorAction SilentlyContinue) { CloseToast -Key $confirmKey }

                    $onlineJson = $targetState['OnlineData'] | ConvertTo-Json -Depth 20
                    $newData = $onlineJson | ConvertFrom-Json

                    $targetState['LocalData'] = $newData
                    $onlineJson | Set-Content -Path $targetPath -Encoding UTF8 -Force

                    if (-not $targetViewMode.Checked) {
                        $targetForm.Tag = $newData
                        & $targetPopulate $null $newData $targetTree
                    }
                    
                    if (Get-Command ShowToast -Ea SilentlyContinue) { 
                        ShowToast -Title "Sync Complete" -Message "Local wiki completely overwritten with online data." -Type "Success" 
                    } else { 
                        [System.Windows.Forms.MessageBox]::Show("Local wiki synced with online version!") | Out-Null 
                    }
                } catch {
                     if (Get-Command ShowToast -Ea SilentlyContinue) { ShowToast "Sync Error" "Error: $_" "Error" }
                     else { [System.Windows.Forms.MessageBox]::Show("Error: $_") | Out-Null }
                }
            }.GetNewClosure()

            $msg = "Are you sure you want to overwrite your ENTIRE local wiki with the online version?`n`nAll local custom edits will be lost."
            
            if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue) {
                $btns = @( @{ Text = "Yes, Overwrite"; Action = $performSyncAll }, @{ Text = "Cancel"; Action = { CloseToast -Key $confirmKey }.GetNewClosure() } )
                ShowInteractiveNotification -Title "Confirm Full Sync" -Message $msg -Buttons $btns -Type "Warning" -Key $confirmKey -TimeoutSeconds 15
            } else {
                if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Full Sync", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -eq 'Yes') { 
                    & $performSyncAll 
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error initiating sync: $_") | Out-Null
        }
    }.GetNewClosure())

    $btnSmartSync.Add_Click({
        try {
            $online = $WikiState['OnlineData']
            if ($null -eq $online) { return }
            
            $local = $WikiState['LocalData']
            if ($null -eq $local) { return }

            $confirmKey = "WikiSmartSync".GetHashCode()
            
            $targetPath = $LocalWikiPath
            $targetForm = $form
            $targetState = $WikiState
            $targetTree = $treeView
            $targetPopulate = $PopulateTree
            $targetViewMode = $chkViewMode
            $targetGetKeys = $getKeys

            $performSmartSync = {
                try {
                    if (Get-Command CloseToast -ErrorAction SilentlyContinue) { CloseToast -Key $confirmKey }

                    $MergeLogic = {
                        param($L, $O)
                        if ($null -eq $O) { return }
                        
                        $oKeys = & $targetGetKeys $O
                        foreach ($k in $oKeys) {
                            if ($k -in 'meta','description','url','URL','note') { continue }
                            
                            $lHas = $false
                            $lVal = $null
                            
                            if ($L -is [System.Collections.IDictionary]) {
                                if ($L.Contains($k)) { $lHas = $true; $lVal = $L[$k] }
                            } elseif ($L.PSObject) {
                                if ($L.PSObject.Properties[$k]) { $lHas = $true; $lVal = $L.$k }
                            }
                            
                            $oVal = if ($O -is [System.Collections.IDictionary]) { $O[$k] } else { $O.$k }
                            
                            if (-not $lHas) {
                                $clone = $oVal | ConvertTo-Json -Depth 20 -Compress | ConvertFrom-Json
                                if ($L -is [System.Collections.IDictionary]) { $L[$k] = $clone }
                                else { $L | Add-Member -MemberType NoteProperty -Name $k -Value $clone -Force }
                            } else {
                                $lIsCont = ($lVal -is [System.Collections.IDictionary] -or $lVal -is [System.Management.Automation.PSCustomObject])
                                $oIsCont = ($oVal -is [System.Collections.IDictionary] -or $oVal -is [System.Management.Automation.PSCustomObject])
                                if ($lIsCont -and $oIsCont) {
                                    & $MergeLogic $lVal $oVal
                                }
                            }
                        }
                    }

                    & $MergeLogic $local $online
                    
                    $json = $local | ConvertTo-Json -Depth 20
                    $json | Set-Content -Path $targetPath -Encoding UTF8 -Force
                    
                    if (-not $targetViewMode.Checked) {
                        $targetForm.Tag = $local
                        & $targetPopulate $null $local $targetTree
                    }

                    if (Get-Command ShowToast -Ea SilentlyContinue) { 
                        ShowToast -Title "Smart Sync Complete" -Message "Missing nodes added from online wiki." -Type "Success" 
                    } else { 
                        [System.Windows.Forms.MessageBox]::Show("Smart Sync Complete!") | Out-Null 
                    }

                } catch {
                     if (Get-Command ShowToast -Ea SilentlyContinue) { ShowToast "Sync Error" "Error: $_" "Error" }
                     else { [System.Windows.Forms.MessageBox]::Show("Error: $_") | Out-Null }
                }
            }.GetNewClosure()

            $msg = "Smart Sync will add any missing pages/categories from the online wiki to your local wiki.`n`nYour existing pages will NOT be overwritten.`n`nProceed?"
            
            if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue) {
                $btns = @( @{ Text = "Start Sync"; Action = $performSmartSync }, @{ Text = "Cancel"; Action = { CloseToast -Key $confirmKey }.GetNewClosure() } )
                ShowInteractiveNotification -Title "Confirm Smart Sync" -Message $msg -Buttons $btns -Type "Info" -Key $confirmKey -TimeoutSeconds 15
            } else {
                if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Smart Sync", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information) -eq 'Yes') { 
                    & $performSmartSync 
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error initiating sync: $_") | Out-Null
        }
    }.GetNewClosure())

    $btnSmartSync.Add_Click({
        try {
            $online = $WikiState['OnlineData']
            if ($null -eq $online) { return }
            
            $local = $WikiState['LocalData']
            if ($null -eq $local) { return }

            $confirmKey = "WikiSmartSync".GetHashCode()
            
            $targetPath = $LocalWikiPath
            $targetForm = $form
            $targetState = $WikiState
            $targetTree = $treeView
            $targetPopulate = $PopulateTree
            $targetViewMode = $chkViewMode
            $targetGetKeys = $getKeys

            $performSmartSync = {
                try {
                    if (Get-Command CloseToast -ErrorAction SilentlyContinue) { CloseToast -Key $confirmKey }

                    $MergeLogic = {
                        param($L, $O)
                        if ($null -eq $O) { return }
                        
                        $oKeys = & $targetGetKeys $O
                        foreach ($k in $oKeys) {
                            if ($k -in 'meta','description','url','URL','note') { continue }
                            
                            $lHas = $false
                            $lVal = $null
                            
                            if ($L -is [System.Collections.IDictionary]) {
                                if ($L.Contains($k)) { $lHas = $true; $lVal = $L[$k] }
                            } elseif ($L.PSObject) {
                                if ($L.PSObject.Properties[$k]) { $lHas = $true; $lVal = $L.$k }
                            }
                            
                            $oVal = if ($O -is [System.Collections.IDictionary]) { $O[$k] } else { $O.$k }
                            
                            if (-not $lHas) {
                                $clone = $oVal | ConvertTo-Json -Depth 20 -Compress | ConvertFrom-Json
                                if ($L -is [System.Collections.IDictionary]) { $L[$k] = $clone }
                                else { $L | Add-Member -MemberType NoteProperty -Name $k -Value $clone -Force }
                            } else {
                                $lIsCont = ($lVal -is [System.Collections.IDictionary] -or $lVal -is [System.Management.Automation.PSCustomObject])
                                $oIsCont = ($oVal -is [System.Collections.IDictionary] -or $oVal -is [System.Management.Automation.PSCustomObject])
                                
                                $contentDiff = $false
                                if (-not $lIsCont -and -not $oIsCont) {
                                    if ($lVal -ne $oVal) { $contentDiff = $true }
                                } elseif ($lIsCont -and $oIsCont) {
                                    $lDesc = if ($lVal -is [System.Collections.IDictionary]) { $lVal['description'] } else { $lVal.description }
                                    $oDesc = if ($oVal -is [System.Collections.IDictionary]) { $oVal['description'] } else { $oVal.description }
                                    if ($lDesc -ne $oDesc) { $contentDiff = $true }
                                } else {
                                    $contentDiff = $true
                                }

                                if ($contentDiff) {
                                    $newKey = "$k (Online)"
                                    $clone = $oVal | ConvertTo-Json -Depth 20 -Compress | ConvertFrom-Json
                                    if ($L -is [System.Collections.IDictionary]) { $L[$newKey] = $clone }
                                    else { $L | Add-Member -MemberType NoteProperty -Name $newKey -Value $clone -Force }
                                } elseif ($lIsCont -and $oIsCont) {
                                    & $MergeLogic $lVal $oVal
                                }
                            }
                        }
                    }

                    & $MergeLogic $local $online
                    
                    $json = $local | ConvertTo-Json -Depth 20
                    $json | Set-Content -Path $targetPath -Encoding UTF8 -Force
                    
                    if (-not $targetViewMode.Checked) {
                        $targetForm.Tag = $local
                        & $targetPopulate $null $local $targetTree
                    }

                    if (Get-Command ShowToast -Ea SilentlyContinue) { 
                        ShowToast -Title "Smart Sync Complete" -Message "Missing nodes added. Conflicts saved as '(Online)'." -Type "Success" 
                    } else { 
                        [System.Windows.Forms.MessageBox]::Show("Smart Sync Complete!") | Out-Null 
                    }

                } catch {
                     if (Get-Command ShowToast -Ea SilentlyContinue) { ShowToast "Sync Error" "Error: $_" "Error" }
                     else { [System.Windows.Forms.MessageBox]::Show("Error: $_") | Out-Null }
                }
            }.GetNewClosure()

            $msg = "Smart Sync will merge online data.`n`n1. Missing nodes will be added.`n2. Nodes with different content will be added as 'Node (Online)'.`n`nYour existing data will NOT be overwritten.`n`nProceed?"
            
            if (Get-Command ShowInteractiveNotification -ErrorAction SilentlyContinue) {
                $btns = @( @{ Text = "Start Sync"; Action = $performSmartSync }, @{ Text = "Cancel"; Action = { CloseToast -Key $confirmKey }.GetNewClosure() } )
                ShowInteractiveNotification -Title "Confirm Smart Sync" -Message $msg -Buttons $btns -Type "Info" -Key $confirmKey -TimeoutSeconds 15
            } else {
                if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Smart Sync", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information) -eq 'Yes') { 
                    & $performSmartSync 
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error initiating sync: $_") | Out-Null
        }
    }.GetNewClosure())

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

                $data = $WikiState.OnlineData
                $form.Tag = $data
                & $PopulateTree $null $data $treeView
                
                $lblStatus.Text = "Source: Online (GitHub) [Read-Only]"
                $chkEdit.Enabled = $false
                $chkEdit.Checked = $false
            } else {
                $s.Text = "VIEW ONLINE"
                $s.BackColor = $Theme.Background
                $s.ForeColor = $Theme.SubTextColor

                $data = $WikiState.LocalData
                $form.Tag = $data
                & $PopulateTree $null $data $treeView

                $lblStatus.Text = "Source: Offline (Local File)"
                $chkEdit.Enabled = ($null -ne $data)
            }
            
            $NavHistory.Clear()
            $btnBack.Enabled = $false

            & $UpdateSyncButtonState

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
        
        & $UpdateSyncButtonState
        
        if ($isEdit) {
            $chkEdit.BackColor = $Theme.AccentColor
            $chkEdit.ForeColor = "White"
            
            $btnPanel.Visible = $false
            $pnlEditFields.Visible = $true
            $pnlFormatting.Visible = $true
            
            # Split screen logic
            $editPreviewSplitter.Top = 285
            $editPreviewSplitter.Height = ($contentPanel.Height - 295)
            $editPreviewSplitter.Panel1Collapsed = $false
            $editPreviewSplitter.SplitterDistance = [int]($editPreviewSplitter.Height / 2)
            
            if ([Custom.Native] -as [Type]) { [Custom.Native]::UseImmersiveDarkMode($txtDescription.Handle) }
            
            # Trigger initial render for preview
            $txtDescription.Text = $txtDescription.Text # Triggers TextChanged
            
        } else {
            $chkEdit.BackColor = $Theme.Background
            $chkEdit.ForeColor = $Theme.SubTextColor
            
            # Clear unsaved buffer on cancel/exit edit
            $WikiState.UnsavedBuffer.Clear()

            $pnlEditFields.Visible = $false
            $pnlFormatting.Visible = $false
            
            $editPreviewSplitter.Top = 180
            $editPreviewSplitter.Height = ($contentPanel.Height - 195)
            $editPreviewSplitter.Panel1Collapsed = $true
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

    # --- Persistence Logic: BeforeSelect ---
    $treeView.Add_BeforeSelect({
        param($src, $e)
        if ($chkEdit.Checked -and $treeView.SelectedNode) {
            # Save current text to buffer before switching
            $path = $treeView.SelectedNode.FullPath
            $WikiState.UnsavedBuffer[$path] = @{
                Text = $txtDescription.Text
                Scroll = $txtDescription.SelectionStart
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
            & $UpdateSyncButtonState
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
            
            # Clear unsaved buffer for this node
            if ($WikiState.UnsavedBuffer.ContainsKey($node.FullPath)) {
                $WikiState.UnsavedBuffer.Remove($node.FullPath)
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

	$WikiState.LocalData = $LocalWikiData
	& $PopulateTree $null $WikiState.LocalData $treeView
    
    # Pre-populate autocomplete list
    $WikiState.NodeList = & $PopulateNodeList $WikiState.LocalData

	$chkEdit.Enabled = ($null -ne $WikiState.LocalData)
	$chkViewMode.Enabled = ($null -ne $WikiState.OnlineData -and $null -ne $WikiState.LocalData)
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
		$capturedState = $WikiState
		$capturedLocalPath = $LocalWikiPath
		$capturedTree = $treeView
		$capturedPopulate = $PopulateTree
        $capturedSyncUpdate = $UpdateSyncButtonState
        $capturedPopulateNodeList = $PopulateNodeList

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
		$pollTimer.Interval = 1000
		$pollTimer.Add_Tick({
			if ($asyncResult.IsCompleted) {
				$pollTimer.Stop()
				
				try {
					$resultCol = $ps.EndInvoke($asyncResult)
					if ($resultCol -and $resultCol.Count -gt 0) {
						$capturedState.OnlineData = $resultCol[0]
					}
				} catch { 
                    $pollTimer.Dispose()
                    return 
                }

				if ($capturedForm.IsDisposed) { 
                    $pollTimer.Dispose()
                    return 
                }
				
				if ($capturedState.OnlineData) {
					$capturedForm.Text = "Entropia Wiki - Online (GitHub)"
					$capturedStatus.Text = "Source: Online (GitHub)"
					$capturedStatus.ForeColor = $capturedTheme.SuccessColor
					
                    if (-not $capturedState.LocalData) {
                        $capturedState.LocalData = $capturedState.OnlineData
                        $capturedForm.Tag = $capturedState.OnlineData
						& $capturedPopulate $null $capturedState.OnlineData $capturedTree
                        
                        # Populate autocomplete list
                        $capturedState.NodeList = & $capturedPopulateNodeList $capturedState.LocalData
                    }

					$capturedViewMode.Enabled = ($null -ne $capturedState.LocalData -and $null -ne $capturedState.OnlineData)
					if ($capturedViewMode.Enabled) {
						$capturedViewMode.Text = "VIEW ONLINE"
					}
                    
                    & $capturedSyncUpdate

					if (-not (Test-Path $capturedLocalPath)) {
						try {
							($capturedState.OnlineData | ConvertTo-Json -Depth 20) | Set-Content -Path $capturedLocalPath -Encoding UTF8 -Force
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