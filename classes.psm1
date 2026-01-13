<# classes.psm1 #>

#region Global Configuration
$script:ReferencedAssemblies = @(
	'System.Windows.Forms',
	'System.Drawing',
	'System.Management.Automation'
)

$script:CompilationOptions = @{
	Language             = 'CSharp'
	ReferencedAssemblies = $script:ReferencedAssemblies
	WarningAction        = 'SilentlyContinue' 
}
if (-not ($global:DashboardConfig -and $global:DashboardConfig.Paths -and $global:DashboardConfig.Paths.FtoolDLL -and $global:DashboardConfig.Paths.FtoolDLL.Trim()))
{
	$script:FToolDll = (Join-Path $env:APPDATA 'Entropia_Dashboard\modules\ftool.dll')
}
else
{
	$script:FToolDll = $global:DashboardConfig.Paths.FtoolDLL
}
#endregion 

#region C#Class Definitions
$classes = @"
using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Reflection;

namespace Custom
{
    // --- FIXED DARK TREEVIEW ---
	public class DarkTreeView : TreeView
	{
		// API to force Dark Scrollbars (Windows 10 1809+)
		[DllImport("dwmapi.dll")]
		private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

		// API to set Explorer style (cleaner look)
		[DllImport("uxtheme.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
		private static extern int SetWindowTheme(IntPtr hwnd, string pszSubAppName, string pszSubIdList);

		private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;

		public DarkTreeView()
		{
			this.DrawMode = TreeViewDrawMode.OwnerDrawText; // Essential for background color control
			this.DoubleBuffered = true;                     // Prevents flickering
			this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
			
			// Default Dark Styles
			this.BackColor = Color.FromArgb(22, 26, 38); 
			this.ForeColor = Color.White;
			this.BorderStyle = BorderStyle.None;
			this.HideSelection = false; // Keep highlight when clicking buttons
            this.FullRowSelect = true;
            this.ShowLines = false;
            this.ShowPlusMinus = true;
		}

		protected override void OnHandleCreated(EventArgs e)
		{
			base.OnHandleCreated(e);
			
			// 1. Apply Explorer Style (Adds fade animations and cleaner chevrons)
			SetWindowTheme(this.Handle, "Explorer", null);

			// 2. Apply Native Dark Mode to Scrollbars
			int darkMode = 1;
			if (DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkMode, 4) != 0)
            {
                 DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref darkMode, 4);
            }
		}

        // Fix for black text glitch on selection loss
        protected override void OnDrawNode(DrawTreeNodeEventArgs e)
        {
            // If the script defines a DrawNode event, this allows it to run.
            // If not, we provide a safe default draw to prevent "invisible text".
            if (e.Node == null) return;

            Font font = e.Node.NodeFont ?? e.Node.TreeView.Font;
            Color fore = e.Node.TreeView.ForeColor;

            // Handle Selection Highlight
            if ((e.State & TreeNodeStates.Selected) != 0)
            {
                // Dark Blue/Gray Highlight
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(45, 50, 65)))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds);
                }
                fore = Color.White;
            }
            else
            {
                 using (SolidBrush brush = new SolidBrush(this.BackColor))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds);
                }
            }

            // Draw Text
            TextRenderer.DrawText(e.Graphics, e.Node.Text, font, e.Bounds, fore, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
        }
	}
    // ---------------------------

	public class DarkMessageBox : Form
	{
		#region DWM API for Dark Title Bar
		[DllImport("dwmapi.dll")]
		private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

		private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
		private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

		private void UseImmersiveDarkMode(IntPtr handle)
		{
			int darkMode = 1;
			if (DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkMode, sizeof(int)) != 0)
			{
				DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref darkMode, sizeof(int));
			}
		}
		#endregion

		private readonly Color _backColor = Color.FromArgb(28, 28, 28);
		private readonly Color _secondaryColor = Color.FromArgb(45, 45, 48);
		private readonly Color _accentColor = Color.FromArgb(0, 120, 215); 
		private readonly Color _textColor = Color.FromArgb(230, 230, 230);
		private bool _isSuccess = false;

		public DarkMessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, bool isSuccess = false)
		{
			this.SuspendLayout();

			this._isSuccess = isSuccess;
			
			// Form Setup
			this.Text = caption;
			this.BackColor = _backColor;
			this.ForeColor = _textColor;
			this.FormBorderStyle = FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.ShowInTaskbar = false;
			this.StartPosition = FormStartPosition.Manual;
            this.Top = (Screen.PrimaryScreen.Bounds.Height - this.Height)/2;
            this.Left = (Screen.PrimaryScreen.Bounds.Width - this.Width)/2;
			this.Font = new Font("Segoe UI", 9.5f);
			this.AutoSize = true;
			this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
			this.MinimumSize = new Size(350, 150);

			// Layout Container
			TableLayoutPanel mainLayout = new TableLayoutPanel
			{
				Dock = DockStyle.Fill,
				AutoSize = true,
				AutoSizeMode = AutoSizeMode.GrowAndShrink,
				ColumnCount = 2,
				RowCount = 2,
				Padding = new Padding(25)
			};
			mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
			mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
			this.Controls.Add(mainLayout);

			// Icon Handling (Modern Drawn Icons)
			if (icon != MessageBoxIcon.None || _isSuccess)
			{
				PictureBox iconBox = new PictureBox {
					Size = new Size(32, 32),
					Margin = new Padding(0, 0, 20, 0),
					Image = GetModernIcon(icon, _isSuccess) // Pass the flag here
				};
				mainLayout.Controls.Add(iconBox, 0, 0);
			}

			// Message Label
			Label messageLabel = new Label
			{
				Text = text,
				AutoSize = true,
				MaximumSize = new Size(450, 800),
				ForeColor = _textColor,
				TextAlign = ContentAlignment.MiddleLeft,
				Dock = DockStyle.Fill,
				Padding = new Padding(0, 5, 0, 0)
			};
			mainLayout.Controls.Add(messageLabel, 1, 0);

			// Button Panel
			FlowLayoutPanel buttonPanel = new FlowLayoutPanel
			{
				FlowDirection = FlowDirection.RightToLeft,
				AutoSize = true,
				Dock = DockStyle.Fill,
				Margin = new Padding(0, 25, 0, 0)
			};
			mainLayout.Controls.Add(buttonPanel, 0, 1);
			mainLayout.SetColumnSpan(buttonPanel, 2);

			AddButtons(buttons, buttonPanel);

			this.ResumeLayout(true);
			this.HandleCreated += (s, e) => UseImmersiveDarkMode(this.Handle);
		}

		private void AddButtons(MessageBoxButtons buttons, FlowLayoutPanel panel)
		{
			switch (buttons)
			{
				case MessageBoxButtons.OK: 
					AddButton(panel, "OK", DialogResult.OK, true); 
					break;
				case MessageBoxButtons.OKCancel: 
					AddButton(panel, "Cancel", DialogResult.Cancel); 
					AddButton(panel, "OK", DialogResult.OK, true); 
					break;
				case MessageBoxButtons.YesNo: 
					AddButton(panel, "No", DialogResult.No); 
					AddButton(panel, "Yes", DialogResult.Yes, true); 
					break;
			}
		}

		private Bitmap GetModernIcon(MessageBoxIcon icon, bool isSuccess = false)
		{
			Bitmap bmp = new Bitmap(32, 32);
			using (Graphics g = Graphics.FromImage(bmp))
			{
				g.SmoothingMode = SmoothingMode.AntiAlias;
				g.PixelOffsetMode = PixelOffsetMode.HighQuality;

				Color iconColor = Color.FromArgb(0, 140, 255); 
				if (isSuccess) 
					iconColor = Color.FromArgb(40, 200, 100); 
				else if (icon == MessageBoxIcon.Error || icon == MessageBoxIcon.Hand || icon == MessageBoxIcon.Stop)
					iconColor = Color.FromArgb(255, 60, 60); 
				else if (icon == MessageBoxIcon.Exclamation || icon == MessageBoxIcon.Warning)
					iconColor = Color.FromArgb(255, 180, 0); 

				using (Pen ringPen = new Pen(iconColor, 2))
				{
					g.DrawEllipse(ringPen, 2, 2, 28, 28);
				}

				using (Pen p = new Pen(iconColor, 2.5f))
				{
					p.StartCap = LineCap.Round;
					p.EndCap = LineCap.Round;
					p.LineJoin = LineJoin.Round;

					float centerX = 16f;

					if (isSuccess)
					{
						g.DrawLine(p, 10, 16, 14, 21);
						g.DrawLine(p, 14, 21, 22, 11);
					}
					else if (icon == MessageBoxIcon.Question)
					{
						GraphicsPath qPath = new GraphicsPath();
						qPath.AddArc(11, 9, 10, 8, 180, 180); 
						qPath.AddLine(21, 13, 16, 15);
						qPath.AddLine(16, 15, 16, 19);
						g.DrawPath(p, qPath);
						g.FillEllipse(new SolidBrush(iconColor), centerX - 1.5f, 22, 3, 3);
					}
					else if (icon == MessageBoxIcon.Exclamation || icon == MessageBoxIcon.Warning)
					{
						g.DrawLine(p, centerX, 9, centerX, 19);
						g.FillEllipse(new SolidBrush(iconColor), centerX - 1.5f, 22, 3, 3);
					}
					else if (icon == MessageBoxIcon.Error || icon == MessageBoxIcon.Hand || icon == MessageBoxIcon.Stop)
					{
						g.DrawLine(p, 11, 11, 21, 21);
						g.DrawLine(p, 21, 11, 11, 21);
					}
					else 
					{
						g.FillEllipse(new SolidBrush(iconColor), centerX - 1.5f, 8, 3, 3);
						g.DrawLine(p, centerX, 13, centerX, 23);
					}
				}
			}
			return bmp;
		}

		private void AddButton(FlowLayoutPanel panel, string text, DialogResult result, bool isPrimary = false)
		{
			Color btnBack = isPrimary ? Color.FromArgb(0, 90, 158) : Color.FromArgb(55, 55, 58);
			Color btnHover = isPrimary ? Color.FromArgb(0, 120, 215) : Color.FromArgb(70, 70, 75);

			Button btn = new Button
			{
				Text = text,
				DialogResult = result,
				Size = new Size(100, 32),
				FlatStyle = FlatStyle.Flat,
				Margin = new Padding(10, 0, 0, 0),
				Cursor = Cursors.Hand,
				BackColor = btnBack,
				ForeColor = Color.White,
				Font = new Font("Segoe UI", 9f, isPrimary ? FontStyle.Bold : FontStyle.Regular)
			};
			
			btn.FlatAppearance.BorderSize = 0;
			btn.FlatAppearance.MouseOverBackColor = btnHover;

			btn.Click += (s, e) => { 
				this.DialogResult = result; 
				this.Close(); 
			};

			panel.Controls.Add(btn);
			if (isPrimary) this.AcceptButton = btn;
			if (result == DialogResult.Cancel || result == DialogResult.No) this.CancelButton = btn;
		}

		public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, string type)
		{
			bool successFlag = (type.ToLower() == "success");
			using (var form = new DarkMessageBox(text, caption, buttons, icon, successFlag))
			{
				return form.ShowDialog();
			}
		}

		public static DialogResult Show(string text, string caption = "Notification", MessageBoxButtons buttons = MessageBoxButtons.OK, MessageBoxIcon icon = MessageBoxIcon.Information)
		{
    		using (var form = new DarkMessageBox(text, caption, buttons, icon)) 
			{
				return form.ShowDialog();
			}
		}

	}

	public class DarkInputBox : Form
    {
		#region DWM API for Dark Title Bar
		[DllImport("dwmapi.dll")]
		private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

		private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
		private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

		private void UseImmersiveDarkMode(IntPtr handle)
		{
			int darkMode = 1;
			if (DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkMode, sizeof(int)) != 0)
			{
				DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref darkMode, sizeof(int));
			}
		}
		#endregion

        public string ResultText { get; private set; }
        private TextBox _textBox;
		
		private readonly Color _backColor = Color.FromArgb(28, 28, 28);
		private readonly Color _secondaryColor = Color.FromArgb(45, 45, 48);
		private readonly Color _textColor = Color.FromArgb(230, 230, 230);

        public DarkInputBox(string title, string prompt, string defaultText)
        {
			this.SuspendLayout();

            this.Text = title;
            this.BackColor = _backColor;
            this.ForeColor = _textColor;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(350, 160);
            this.TopMost = true;
			this.Font = new Font("Segoe UI", 9.5f);

			// Layout Container
			TableLayoutPanel mainLayout = new TableLayoutPanel
			{
				Dock = DockStyle.Fill,
				ColumnCount = 1,
				RowCount = 3,
				Padding = new Padding(20)
			};
			mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Label
			mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // TextBox
			mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // Buttons
			this.Controls.Add(mainLayout);

            Label lbl = new Label();
            lbl.Text = prompt;
			lbl.AutoSize = true;
			lbl.MaximumSize = new Size(310, 0);
            lbl.ForeColor = _textColor;
			lbl.Margin = new Padding(0, 0, 0, 10);
            mainLayout.Controls.Add(lbl, 0, 0);

            _textBox = new TextBox();
            _textBox.Text = defaultText;
			_textBox.Dock = DockStyle.Top;
            _textBox.BackColor = _secondaryColor;
            _textBox.ForeColor = _textColor;
            _textBox.BorderStyle = BorderStyle.FixedSingle;
			_textBox.KeyDown += (s, e) => { 
				if (e.KeyCode == Keys.Enter) { 
					this.ResultText = _textBox.Text; 
					this.DialogResult = DialogResult.OK; 
					this.Close(); 
					e.Handled = true; 
					e.SuppressKeyPress = true; 
				} 
			};
            mainLayout.Controls.Add(_textBox, 0, 1);

			// Button Panel
			FlowLayoutPanel buttonPanel = new FlowLayoutPanel
			{
				FlowDirection = FlowDirection.RightToLeft,
				AutoSize = true,
				Dock = DockStyle.Bottom,
				Margin = new Padding(0, 20, 0, 0)
			};
			mainLayout.Controls.Add(buttonPanel, 0, 2);

			AddButton(buttonPanel, "Cancel", DialogResult.Cancel);
			AddButton(buttonPanel, "OK", DialogResult.OK, true);

			this.ResumeLayout(true);
			this.HandleCreated += (s, e) => UseImmersiveDarkMode(this.Handle);
			this.Shown += (s, e) => { _textBox.Focus(); _textBox.SelectAll(); };
        }

		private void AddButton(FlowLayoutPanel panel, string text, DialogResult result, bool isPrimary = false)
		{
			Color btnBack = isPrimary ? Color.FromArgb(0, 90, 158) : Color.FromArgb(55, 55, 58);
			Color btnHover = isPrimary ? Color.FromArgb(0, 120, 215) : Color.FromArgb(70, 70, 75);

			Button btn = new Button
			{
				Text = text,
				DialogResult = result,
				Size = new Size(90, 30),
				FlatStyle = FlatStyle.Flat,
				Margin = new Padding(10, 0, 0, 0),
				Cursor = Cursors.Hand,
				BackColor = btnBack,
				ForeColor = Color.White,
				Font = new Font("Segoe UI", 9f, isPrimary ? FontStyle.Bold : FontStyle.Regular)
			};
			
			btn.FlatAppearance.BorderSize = 0;
			btn.FlatAppearance.MouseOverBackColor = btnHover;

			btn.Click += (s, e) => { 
				if (result == DialogResult.OK) this.ResultText = _textBox.Text;
				this.DialogResult = result; 
				this.Close(); 
			};

			panel.Controls.Add(btn);
			if (isPrimary) this.AcceptButton = btn;
			if (result == DialogResult.Cancel) this.CancelButton = btn;
		}
    }

	public static class MouseHookManager 
	{
		public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
		private static HookProc hookDelegate;
		public static IntPtr HookId = IntPtr.Zero;

		public static void Start(HookProc proc) 
		{
			hookDelegate = proc;
			HookId = SetWindowsHookEx(14, proc, GetModuleHandle(null), 0);
		}

		public static void Stop() 
		{
			if (HookId != IntPtr.Zero)
   			{
				UnhookWindowsHookEx(HookId);
				HookId = IntPtr.Zero;
			}
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool UnhookWindowsHookEx(IntPtr hhk);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
		
		[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr GetModuleHandle(string lpModuleName);
	}

	public static class Native
	{

		[DllImport("dwmapi.dll")]
		private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

		private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
		private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

		public static void UseImmersiveDarkMode(IntPtr handle)
		{
			int darkMode = 1;
			if (DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkMode, 4) != 0)
			{
				DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref darkMode, 4);
			}
		}

		[DllImport("iphlpapi.dll", SetLastError = true)]
		private static extern uint GetExtendedTcpTable(IntPtr pTcpTable, ref int dwOutBufLen, bool sort, int ipVersion, int tblClass, int reserved);

		[DllImport("iphlpapi.dll")]
		private static extern int SetTcpEntry(ref MIB_TCPROW pTcpRow);

		[StructLayout(LayoutKind.Sequential)]
		private struct MIB_TCPROW
		{
			public uint dwState;
			public uint dwLocalAddr;
			public int dwLocalPort;
			public uint dwRemoteAddr;
			public int dwRemotePort;
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct MIB_TCPROW_OWNER_PID
		{
			public uint dwState;
			public uint dwLocalAddr;
			public int dwLocalPort;
			public uint dwRemoteAddr;
			public int dwRemotePort;
			public int dwOwningPid;
		}

		public static int CloseTcpConnectionsForPid(int pid)
		{
			int closedCount = 0;
			int buffSize = 0;
			GetExtendedTcpTable(IntPtr.Zero,ref buffSize, false, 2, 5, 0);
			IntPtr tcpTablePtr = Marshal.AllocHGlobal(buffSize);
			try
			{
				if (GetExtendedTcpTable(tcpTablePtr,ref buffSize, false, 2, 5, 0) == 0)
				{
					int entryCount = Marshal.ReadInt32(tcpTablePtr);
					IntPtr rowPtr = (IntPtr)((long)tcpTablePtr + 4);
					int rowSize = Marshal.SizeOf(typeof(MIB_TCPROW_OWNER_PID));

					for (int i = 0; i < entryCount; i++)
					{
						MIB_TCPROW_OWNER_PID row = (MIB_TCPROW_OWNER_PID)Marshal.PtrToStructure(rowPtr,typeof(MIB_TCPROW_OWNER_PID));
						if (row.dwOwningPid == pid && row.dwState == 5)
						{
							MIB_TCPROW setRow = new MIB_TCPROW();
							setRow.dwState = 12;
							setRow.dwLocalAddr = row.dwLocalAddr;
							setRow.dwLocalPort = row.dwLocalPort;
							setRow.dwRemoteAddr = row.dwRemoteAddr;
							setRow.dwRemotePort = row.dwRemotePort;
							
							if (SetTcpEntry(ref setRow) == 0) closedCount++;
						}
						rowPtr = (IntPtr)((long)rowPtr + rowSize);
					}
				}
			}
			catch {}
			finally
			{
				Marshal.FreeHGlobal(tcpTablePtr);
			}
			return closedCount;
		}

		[DllImport("user32.dll", EntryPoint = "SetWindowPos", SetLastError = true)]
		public static extern bool PositionWindow(IntPtr windowHandle, IntPtr insertAfterHandle, int X, int Y, int width, int height, uint flags);

		[DllImport("user32.dll", EntryPoint = "GetWindowRect")]
		public static extern bool GetWindowRect(IntPtr windowHandle, out RECT rect);

		[DllImport("user32.dll", EntryPoint = "ShowWindow")]
		public static extern bool ShowWindow(IntPtr windowHandle, int nCmdShow);

		[DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
		public static extern bool SetForegroundWindow(IntPtr windowHandle);

		[DllImport("user32.dll", EntryPoint = "GetForegroundWindow")]
		public static extern IntPtr GetForegroundWindow();

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsWindow(IntPtr hWnd);

		[DllImport("user32.dll", SetLastError = true)]
		public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

		[DllImport("user32.dll", EntryPoint = "IsIconic")]
		public static extern bool IsWindowMinimized(IntPtr windowHandle);

		[DllImport("user32.dll", EntryPoint = "SendMessageTimeout", SetLastError = true)]
		public static extern IntPtr GetWindowResponse(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

		[DllImport("psapi.dll", EntryPoint = "EmptyWorkingSet")]
		public static extern bool EmptyWorkingSet(IntPtr hProcess);

		[DllImport("user32.dll", EntryPoint = "MsgWaitForMultipleObjects", SetLastError = true)]
		public static extern uint AsyncExecution(uint nCount, IntPtr[] pHandles, bool bWaitAll, uint dwMilliseconds, uint dwWakeMask);

		[DllImport("user32.dll", EntryPoint = "PeekMessage", SetLastError = true)]
		public static extern bool PeekMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

		[DllImport("user32.dll", EntryPoint = "TranslateMessage" )]
		public static extern bool TranslateMessage(ref MSG lpMsg);

		[DllImport("user32.dll", EntryPoint = "DispatchMessage" )]
		public static extern IntPtr DispatchMessage(ref MSG lpMsg);

		[DllImport("user32.dll", EntryPoint = "ReleaseCapture" )]
		public static extern bool ReleaseCapture();

		[DllImport("user32.dll", EntryPoint = "SendMessage" )]
		public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

		[DllImport("user32.dll")]
		public static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);
		
		[DllImport("user32.dll", SetLastError = true)]
		public static extern bool SetCursorPos(int x, int y);
		
		[DllImport("user32.dll")]
		public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
		
		[DllImport("user32.dll", SetLastError = true)]
		public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);

		[DllImport("user32.dll")]
		public static extern bool ClipCursor(ref RECT lpRect);

		[DllImport("user32.dll")]
		public static extern bool ClipCursor(IntPtr lpRect);

		[DllImport("user32.dll")]
		public static extern bool GetCursorPos(out Point lpPoint);

		[DllImport("user32.dll", SetLastError = true)]
		public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

		[DllImport("user32.dll", EntryPoint = "GetActiveWindow")]
		public static extern IntPtr GetActiveWindowHandle();

		[DllImport("user32.dll", EntryPoint = "IsWindowVisible")]
		public static extern bool IsWindowActive(IntPtr windowHandle);

		[DllImport("user32.dll", EntryPoint = "SetWindowText", SetLastError = true, CharSet = CharSet.Auto)]
		public static extern bool SetWindowTitle(IntPtr windowHandle, string windowTitle);

		[DllImport("user32.dll", EntryPoint = "GetWindowLong")]
		private static extern IntPtr GetWindowLong32(IntPtr hWnd, int nIndex);

		[DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
		private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

		public static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex)
		{
			return IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd,nIndex) : GetWindowLong32(hWnd,nIndex);
		}

		[DllImport("user32.dll", EntryPoint = "SetWindowLong")]
		private static extern IntPtr SetWindowLong32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

		[DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
		private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

		public static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
		{
			return IntPtr.Size == 8 ? SetWindowLongPtr64(hWnd,nIndex, dwNewLong) : SetWindowLong32(hWnd,nIndex, dwNewLong);
		}
		
		[DllImport("user32.dll")]
		public static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

		public static IntPtr GetWindowHandle(object windowIdentifier)
		{
			if (windowIdentifier is IntPtr) return (IntPtr)windowIdentifier;
			if (windowIdentifier is int) return FindWindowByProcessId((int)windowIdentifier);
			return IntPtr.Zero;
		}

		public static IntPtr FindWindowByProcessId(int processId)
		{
			try { return Process.GetProcessById(processId).MainWindowHandle; }
			catch { return IntPtr.Zero; }
		}

		public static bool BringToFront(object windowIdentifier)
		{
			IntPtr windowHandle = GetWindowHandle(windowIdentifier);
			if (windowHandle == IntPtr.Zero) return false;
			if (Native.IsWindowMinimized(windowHandle)) Native.ShowWindow(windowHandle,Native.SW_RESTORE);
			return Native.SetForegroundWindow(windowHandle);
		}

		public static bool SendToBack(object windowIdentifier)
		{
			IntPtr windowHandle = GetWindowHandle(windowIdentifier);
			if (windowHandle == IntPtr.Zero) return false;
			return Native.ShowWindow(windowHandle,Native.SW_MINIMIZE);
		}

		public static bool IsMinimized(object windowIdentifier)
		{
			IntPtr windowHandle = GetWindowHandle(windowIdentifier);
			if (windowHandle == IntPtr.Zero) return false;
			return Native.IsWindowMinimized(windowHandle);
		}

		public static bool Responsive(IntPtr hWnd, uint timeout = 100)
		{
			try
			{
				if (hWnd == IntPtr.Zero) return false;
				IntPtr result;
				IntPtr sendResult = Native.GetWindowResponse(hWnd,Native.WM_NULL, IntPtr.Zero, IntPtr.Zero, Native.SMTO_ABORTIFHUNG, timeout, out result);
				return sendResult != IntPtr.Zero;
			}
			catch { return false; }
		}

		public static Task<bool> ResponsiveAsync(IntPtr hWnd, uint timeout = 100)
		{
			var tcs = new TaskCompletionSource<bool>();
			ThreadPool.QueueUserWorkItem(state => {
					try { tcs.SetResult(Responsive(hWnd,timeout)); }
					catch (Exception ex) { tcs.SetException(ex); }
				});
			return tcs.Task;
		}

		[StructLayout(LayoutKind.Sequential)]
		public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

		[StructLayout(LayoutKind.Sequential)]
		public struct MSG { public IntPtr hwnd; public uint message; public UIntPtr wParam; public IntPtr lParam; public uint time; public System.Drawing.Point pt; }

		public static readonly IntPtr TopWindowHandle = new IntPtr(0);
		public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
		public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
		
		public const uint WM_NULL = 0x0000;
		public const uint SMTO_ABORTIFHUNG = 0x0002;
		public const int WM_SETREDRAW = 0x000B; 
		
		public const uint GW_HWNDNEXT = 2;
		public const uint GW_HWNDPREV = 3;

		
		public const int SW_HIDE = 0;
		public const int SW_SHOW = 5;
		public const int SW_MAXIMIZE = 3;
		public const int SW_MINIMIZE = 6;
		public const int SW_SHOWMINNOACTIVE = 7; 
		public const int SW_SHOWNA = 8;
		public const int SW_RESTORE = 9;

		public const uint QS_ALLINPUT = 0x04FF;
		public const uint PM_REMOVE = 0x0001;
		
		public const uint SWP_NOSIZE = 0x0001;      
		public const uint SWP_NOMOVE = 0x0002;      
		public const uint SWP_NOZORDER = 0x0004;    
		public const uint SWP_NOACTIVATE = 0x0010;
		public const uint SWP_FRAMECHANGED = 0x0020;
		public const int SWP_SHOWWINDOW = 0x0040;   
		
		public const int ERROR_TIMEOUT = 1460;
		public const uint WAIT_TIMEOUT = 258;
		
		[DllImport("kernel32.dll", SetLastError = true)]
		public static extern IntPtr OpenProcess(
			uint dwDesiredAccess,
			bool bInheritHandle,
			int dwProcessId
		);

		[DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		public static extern bool QueryFullProcessImageName(
			IntPtr hProcess,
			uint dwFlags,
			StringBuilder lpExeName,
			ref uint lpdwSize
		);

		[DllImport("kernel32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool CloseHandle(IntPtr hObject);

		
		public const uint PROCESS_QUERY_INFORMATION = 0x0400;
		public const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;

		public static string GetProcessPathById(int processId)
		{
			IntPtr hProcess = IntPtr.Zero;
			try
			{
				
				hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION,false, processId);
				if (hProcess == IntPtr.Zero)
				{
					
					hProcess = OpenProcess(PROCESS_QUERY_INFORMATION,false, processId);
				}

				if (hProcess == IntPtr.Zero)
				{
					
					return null;
				}

				uint capacity = 260; 
				StringBuilder sb = new StringBuilder((int)capacity);
				
				
				if (!QueryFullProcessImageName(hProcess,0, sb, ref capacity))
				{
					
					if (Marshal.GetLastWin32Error() == 122) 
					{
						sb.Capacity = (int)capacity; 
						if (!QueryFullProcessImageName(hProcess,0, sb, ref capacity))
						{
							return null; 
						}
					}
					else
					{
						return null; 
					}
				}
				return sb.ToString();
			}
			catch
			{
				
				return null;
			}
			finally
			{
				
				if (hProcess != IntPtr.Zero)
				{
					CloseHandle(hProcess);
				}
			}
		}
		
		public const int GWL_EXSTYLE = -20;
		public const uint WS_EX_TOPMOST = 0x00000008;

		public const byte VK_MENU = 0x12; 
		public const uint KEYEVENTF_EXTENDEDKEY = 0x0001; 
		public const uint KEYEVENTF_KEYUP = 0x0002; 
		
		
		public const uint RDW_INVALIDATE = 0x0001;
		public const uint RDW_ALLCHILDREN = 0x0080;
		public const uint RDW_UPDATENOW = 0x0100;
		public const uint RDW_FRAME = 0x0400;

		[Flags]
		public enum WindowPositionOptions : uint 
		{ 
			NoZOrderChange = SWP_NOZORDER,
			DoNotActivate = SWP_NOACTIVATE,
			MakeVisible = SWP_SHOWWINDOW
		}
	}

	public static class TaskbarTool
	{
		
		[ComImport]
		[Guid("56FDF344-FD6D-11d0-958D-006097C9A090")]
		[ClassInterface(ClassInterfaceType.None)]
		private class TaskbarList { }

		
		[ComImport]
		[Guid("602D4995-B13A-429b-A66E-1935E44F4317")]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		private interface ITaskbarList2
		{
			void HrInit();
			void AddTab(IntPtr hwnd);
			void DeleteTab(IntPtr hwnd);
			void ActivateTab(IntPtr hwnd);
			void SetActiveAlt(IntPtr hwnd);
			void MarkFullscreenWindow(IntPtr hwnd, bool fFullscreen);
		}

		public static bool SetTaskbarState(IntPtr hwnd, bool visible)
		{
			bool result = false;
			
			if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
			{
				Thread t = new Thread(() => { result = SafeSetTaskbarState(hwnd,visible); });
				t.SetApartmentState(ApartmentState.STA);
				t.Start();
				t.Join();
			}
			else
			{
				result = SafeSetTaskbarState(hwnd,visible);
			}
			return result;
		}

		private static bool SafeSetTaskbarState(IntPtr hwnd, bool visible)
		{
			ITaskbarList2 taskbar = null;
			try
			{
				object obj = new TaskbarList();
				taskbar = (ITaskbarList2)obj;
				taskbar.HrInit();

				if (visible)
				taskbar.AddTab(hwnd);
				else
				taskbar.DeleteTab(hwnd);
					
				return true; 
			}
			catch
			{
				return false; 
			}
			finally
			{
				if (taskbar != null && Marshal.IsComObject(taskbar))
				{
					Marshal.ReleaseComObject(taskbar);
				}
			}
		}
	}

	public static class Ftool
	{
		[DllImport(@"$script:FToolDll", EntryPoint = "fnPostMessage", CallingConvention = CallingConvention.StdCall)]
			public static extern void fnPostMessage(IntPtr hWnd, int msg, int wParam, int lParam);
		}

		public class IniFile
		{
			private string filePath;
			public IniFile(string filePath) { this.filePath = filePath; }

			public OrderedDictionary ReadIniFile()
			{
				OrderedDictionary config = new OrderedDictionary();
				if (!File.Exists(filePath)) return config;
				try
				{
					string currentSection = null;
					string[] lines = File.ReadAllLines(filePath);
					foreach (string line in lines)
					{
						string trimmedLine = line.Trim();
						if (string.IsNullOrWhiteSpace(trimmedLine) || trimmedLine.StartsWith(";") || trimmedLine.StartsWith("#")) continue;
						if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
						{
							currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2).Trim();
							if (!config.Contains(currentSection)) config[currentSection] = new OrderedDictionary();
							continue;
						}
						int equalsPos = trimmedLine.IndexOf("=");
						if (equalsPos > 0 && currentSection != null)
						{
							string key = trimmedLine.Substring(0, equalsPos).Trim();
							string value = trimmedLine.Substring(equalsPos + 1).Trim();
							if (value.StartsWith("\"") && value.EndsWith("\"")) value = value.Substring(1, value.Length - 2);
							((OrderedDictionary)config[currentSection])[key] = value;
						}
					}
					return config;
				}
				catch { return new OrderedDictionary(); }
			}

			public void WriteIniFile(OrderedDictionary config)
			{
				try
				{
					string directory = Path.GetDirectoryName(filePath);
					if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) Directory.CreateDirectory(directory);
					List<string> lines = new List<string>();
					bool firstSection = true;
					foreach (DictionaryEntry sectionEntry in config)
					{
						string section = sectionEntry.Key.ToString();
						OrderedDictionary sectionData = (OrderedDictionary)sectionEntry.Value;
						if (!firstSection) lines.Add("");
						lines.Add("[" + section + "]");
						firstSection = false;
						foreach (DictionaryEntry kvpEntry in sectionData)
						{
							string key = kvpEntry.Key.ToString();
							string value = kvpEntry.Value.ToString();
							if (value.Contains(" ") && !(value.StartsWith("\"") && value.EndsWith("\""))) value = "\"" + value + "\"";
							lines.Add(key + " = " + value);
						}
					}
					File.WriteAllLines(filePath,lines);
				}
				catch { }
			}
		}

		public class DarkComboBox : ComboBox
		{
			private Color BackgroundColor = Color.FromArgb(40, 40, 40);
			private Color TextColor = Color.FromArgb(240, 240, 240);
			private Color BorderColor = Color.FromArgb(65, 65, 70);
			private Color ArrowColor = Color.FromArgb(240, 240, 240);
			private Color SelectedItemBackColor = Color.FromArgb(0, 120, 200);
			private Color SelectedItemForeColor = Color.FromArgb(240, 240, 240);
			private Color DropDownBackColor = Color.FromArgb(40, 40, 40);

			public DarkComboBox() : base()
			{
				this.SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw,true);
				this.DrawMode = DrawMode.OwnerDrawFixed;
				this.DropDownStyle = ComboBoxStyle.DropDownList;
				this.FlatStyle = FlatStyle.Flat;
			}

			protected override void OnPaint(PaintEventArgs e)
			{
				Graphics g = e.Graphics;
				Rectangle bounds = this.ClientRectangle;
				using (SolidBrush backBrush = new SolidBrush(this.Enabled ? BackgroundColor : SystemColors.ControlDark))
				g.FillRectangle(backBrush,bounds);

				if (this.DropDownStyle == ComboBoxStyle.DropDownList)
				{
					string text = this.Text;
					if (!string.IsNullOrEmpty(text))
					{
						Rectangle textRect = new Rectangle(bounds.Left - 8, bounds.Top, bounds.Width, bounds.Height);
						TextRenderer.DrawText(g,text, this.Font, textRect, this.Enabled ? TextColor : SystemColors.GrayText, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
					}
				}
				using (Pen borderPen = new Pen(this.Enabled ? BorderColor : SystemColors.ControlDarkDark))
				g.DrawRectangle(borderPen,bounds.Left, bounds.Top, bounds.Width - 2, bounds.Height);

				Rectangle buttonRect = new Rectangle(bounds.Width,bounds.Top, bounds.Left, bounds.Height);
				Point center = new Point(buttonRect.Left + buttonRect.Width / 2, buttonRect.Top + buttonRect.Height / 2);
				Point[] arrowPoints = new Point[] { new Point(center.X - 5, center.Y - 2), new Point(center.X + 5, center.Y - 2), new Point(center.X,center.Y + 3) };
				using (SolidBrush arrowBrush = new SolidBrush(this.Enabled ? ArrowColor : SystemColors.GrayText))
				g.FillPolygon(arrowBrush,arrowPoints);
			}

			protected override void OnDrawItem(DrawItemEventArgs e)
			{
				if (e.Index < 0 || e.Index >= this.Items.Count) return;
				Graphics g = e.Graphics;
				Rectangle bounds = e.Bounds;
				Color currentBackColor = ((e.State & DrawItemState.Selected) == DrawItemState.Selected) ? SelectedItemBackColor : DropDownBackColor;
				Color currentForeColor = ((e.State & DrawItemState.Selected) == DrawItemState.Selected) ? SelectedItemForeColor : TextColor;

				using (SolidBrush backBrush = new SolidBrush(currentBackColor))
				g.FillRectangle(backBrush,bounds.Left, bounds.Top, bounds.Width, bounds.Height);

				string itemText = this.Items[e.Index].ToString();
				Rectangle textBounds = new Rectangle(bounds.Left,bounds.Top, bounds.Width, bounds.Height);
				TextRenderer.DrawText(g,itemText, this.Font, textBounds, currentForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
			}
		}

		public class DarkTabControl : TabControl
		{
			public DarkTabControl()
			{
				this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw,true);
				this.DrawMode = TabDrawMode.OwnerDrawFixed;
				this.SizeMode = TabSizeMode.Fixed;
				this.ItemSize = new Size(197, 30);
			}

			protected override void OnPaint(PaintEventArgs e)
			{
				using (SolidBrush b = new SolidBrush(Color.FromArgb(30, 30, 30)))
				{
					e.Graphics.FillRectangle(b,this.ClientRectangle);
				}
			
				for (int i = 0; i < this.TabCount; i++)
				{
					Rectangle tabRect = this.GetTabRect(i);
					bool isSelected = (this.SelectedIndex == i);
				
					Color bg = isSelected ? Color.FromArgb(50, 50, 50) : Color.FromArgb(30, 30, 30);
				
					using (SolidBrush b = new SolidBrush(bg))
					{
						e.Graphics.FillRectangle(b,tabRect);
					}
				
					string text = this.TabPages[i].Text;
					TextRenderer.DrawText(e.Graphics,text, this.Font, tabRect, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
				
					if (isSelected)
					{
						using (Pen p = new Pen(Color.FromArgb(0, 120, 215), 2))
						{
							e.Graphics.DrawLine(p,tabRect.Left, tabRect.Bottom - 1, tabRect.Right, tabRect.Bottom - 1);
						}
					}
				}
			}
		
			protected override void OnPaintBackground(PaintEventArgs pevent) {}
		}

		public class Toggle : CheckBox
		{
			public Toggle()
			{
				this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint,true);
				this.Padding = new Padding(1);
				this.Appearance = Appearance.Button;
				this.FlatStyle = FlatStyle.Flat;
				this.FlatAppearance.BorderSize = 0;
			}
			protected override void OnPaint(PaintEventArgs e)
			{
				base.OnPaint(e);
				e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
				Color backColor = this.Parent != null ? this.Parent.BackColor : SystemColors.Control;
				e.Graphics.Clear(backColor);
				Color onColor = Color.FromArgb(35, 175, 75);
				Color offColor = Color.FromArgb(210, 45, 45);

				using (var path = new System.Drawing.Drawing2D.GraphicsPath())
				{
					var d = this.Padding.All;
					var r = this.Height - 2 * d;
					path.AddArc(d,d, r, r, 90, 180);
					path.AddArc(this.Width - r - d, d, r, r, -90, 180);
					path.CloseFigure();
					using(var brush = new SolidBrush(this.Checked ? onColor : offColor)) e.Graphics.FillPath(brush,path);

					r = this.Height - 6;
					var rect = this.Checked ? new Rectangle(this.Width - r - 3, 3, r, r) : new Rectangle(3, 3, r, r);
					e.Graphics.FillEllipse(Brushes.White,rect);
				
					Color trackColor = this.Checked ? onColor : offColor;
					double luminance = (0.2126 * trackColor.R) + (0.7152 * trackColor.G) + (0.0722 * trackColor.B);
					Color textColor = (luminance < 140) ? Color.White : Color.Black;

					using (var shadowBrush = new SolidBrush(Color.FromArgb(140, 0, 0, 0)))
					using (var textBrush = new SolidBrush(textColor))
					{
						var sf = new System.Drawing.StringFormat();
						sf.Alignment = System.Drawing.StringAlignment.Center;
						sf.LineAlignment = System.Drawing.StringAlignment.Center;
						var boundsF = new RectangleF(0, 0, this.Width, this.Height);
						var shadowBounds = new RectangleF(boundsF.X + 1, boundsF.Y + 1, boundsF.Width, boundsF.Height);
						e.Graphics.DrawString(this.Text,this.Font, shadowBrush, shadowBounds, sf);
						e.Graphics.DrawString(this.Text,this.Font, textBrush, boundsF, sf);
					}
				}
			}
		}

		public class TextProgressBar : ProgressBar
		{
			private string _customText;
			public string CustomText 
			{ 
				get { return _customText; }
				set 
				{
					_customText = value;
					this.Invalidate(); 
				}
			}
		
			public TextProgressBar()
			{
				CustomText = "";
				this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer,true);
			}

			protected override void OnPaint(PaintEventArgs e)
			{
				Rectangle rect = ClientRectangle;
				Graphics g = e.Graphics;

				ProgressBarRenderer.DrawHorizontalBar(g,rect);
			
				if (this.Value > 0)
				{
					Rectangle clip = new Rectangle(rect.X,rect.Y, (int)Math.Round(((float)this.Value / this.Maximum) * rect.Width), rect.Height);
				
				
					using (SolidBrush b = new SolidBrush(Color.FromArgb(0, 120, 215)))
					{
						g.FillRectangle(b,clip);
					}
				}

				using (Font f = new Font("Segoe UI", 8))
				{
					string text = string.IsNullOrEmpty(CustomText) ? "" : CustomText;
					SizeF len = g.MeasureString(text,f);
					Point location = new Point((int)((rect.Width / 2) - (len.Width / 2)), (int)((rect.Height / 2) - (len.Height / 2)));
				
				
					g.DrawString(text,f, Brushes.Black, location.X + 1, location.Y + 1);
				
					g.DrawString(text,f, Brushes.White, location);
				}
			}
		}

		public class FtoolFormWindow : Form 
		{
			private const int WS_EX_TOOLWINDOW = 0x00000080;
			private const int WS_EX_APPWINDOW = 0x00040000;
			private const int WS_CAPTION = 0x00C00000;
			private const int WS_SIZEBOX = 0x00040000;

			public FtoolFormWindow() 
			{
				this.FormBorderStyle = FormBorderStyle.None;
				this.ShowInTaskbar = false;
			}

			protected override CreateParams CreateParams 
			{
				get {
					CreateParams cp = base.CreateParams;

					cp.ExStyle |= WS_EX_TOOLWINDOW;
					cp.ExStyle &= ~WS_EX_APPWINDOW;
					cp.Style &= ~WS_CAPTION;
					cp.Style &= ~WS_SIZEBOX;

					return cp;
				}
			}
	}

	public class DarkScrollBar : Control
	{
		private int _value = 0;
		private int _maximum = 100;
		private int _largeChange = 10;
		private bool _isHovered = false;
		private bool _isDragging = false;
		private int _dragStartY;
		private int _dragStartValue;

		public event EventHandler Scroll;

		public int Value
		{
			get { return _value; }
			set 
			{ 
				int newVal = Math.Max(0, Math.Min(value, _maximum - (_largeChange > 0 ? _largeChange : 0)));
				if (_value != newVal)
				{
					_value = newVal;
					Invalidate();
					if (Scroll != null) Scroll(this, EventArgs.Empty);
				}
			}
		}

		public int Maximum { get { return _maximum; } set { _maximum = value; Invalidate(); } }
		public int LargeChange { get { return _largeChange; } set { _largeChange = value; Invalidate(); } }

		public DarkScrollBar()
		{
			this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
			this.Width = 12;
			this.BackColor = Color.FromArgb(30, 30, 30);
		}

		protected override void OnMouseEnter(EventArgs e) { _isHovered = true; Invalidate(); base.OnMouseEnter(e); }
		protected override void OnMouseLeave(EventArgs e) { _isHovered = false; Invalidate(); base.OnMouseLeave(e); }
		protected override void OnMouseDown(MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Left)
			{
				int trackHeight = this.Height;
				int thumbHeight = Math.Max(20, (int)((float)_largeChange / Math.Max(1, _maximum) * trackHeight));
				int scrollableHeight = trackHeight - thumbHeight;
				int thumbY = 0;
				if (scrollableHeight > 0 && (_maximum - _largeChange) > 0) thumbY = (int)((float)_value / (_maximum - _largeChange) * scrollableHeight);
				
				Rectangle thumbRect = new Rectangle(2, thumbY, Width - 4, thumbHeight);
				if (thumbRect.Contains(e.Location)) { _isDragging = true; _dragStartY = e.Y; _dragStartValue = _value; }
				else { if (e.Y < thumbY) Value -= _largeChange; else Value += _largeChange; }
			}
			base.OnMouseDown(e);
		}
		protected override void OnMouseMove(MouseEventArgs e)
		{
			if (_isDragging)
			{
				int trackHeight = this.Height;
				int thumbHeight = Math.Max(20, (int)((float)_largeChange / Math.Max(1, _maximum) * trackHeight));
				int scrollableHeight = trackHeight - thumbHeight;
				if (scrollableHeight > 0)
				{
					int deltaY = e.Y - _dragStartY;
					float deltaVal = (float)deltaY / scrollableHeight * (_maximum - _largeChange);
					this.Value = (int)(_dragStartValue + deltaVal);
				}
			}
			base.OnMouseMove(e);
		}
		protected override void OnMouseUp(MouseEventArgs e) { _isDragging = false; base.OnMouseUp(e); }

		protected override void OnPaint(PaintEventArgs e)
		{
			e.Graphics.Clear(this.BackColor);
			if (_maximum <= 0 || _largeChange >= _maximum) return;

			int trackHeight = this.Height;
			int thumbHeight = Math.Max(20, (int)((float)_largeChange / Math.Max(1, _maximum) * trackHeight));
			int scrollableHeight = trackHeight - thumbHeight;
			int thumbY = 0;
			if (scrollableHeight > 0 && (_maximum - _largeChange) > 0) thumbY = (int)((float)_value / (_maximum - _largeChange) * scrollableHeight);
			
			Color thumbColor = _isDragging ? Color.FromArgb(120, 120, 140) : (_isHovered ? Color.FromArgb(90, 90, 100) : Color.FromArgb(60, 60, 60));
			using (System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath())
			{
				int r = 4; Rectangle rect = new Rectangle(2, thumbY, Width - 4, thumbHeight);
				path.AddArc(rect.X, rect.Y, r, r, 180, 90); path.AddArc(rect.Right - r, rect.Y, r, r, 270, 90);
				path.AddArc(rect.Right - r, rect.Bottom - r, r, r, 0, 90); path.AddArc(rect.X, rect.Bottom - r, r, r, 90, 90);
				path.CloseFigure();
				using (SolidBrush b = new SolidBrush(thumbColor)) { e.Graphics.FillPath(b, path); }
				using (Pen p = new Pen(Color.FromArgb(100, 255, 255, 255), 1)) { e.Graphics.DrawLine(p, rect.X + 2, rect.Y + 4, rect.X + 2, rect.Bottom - 4); }
			}
		}
	}

	public class DarkDataGridView : DataGridView
	{
		private DarkScrollBar _vScroll;
		public DarkDataGridView()
		{
			this.ScrollBars = ScrollBars.None;
			_vScroll = new DarkScrollBar();
			_vScroll.Dock = DockStyle.Right;
			_vScroll.Visible = false;
			_vScroll.Scroll += (s, e) => { if (this.RowCount > 0) { try { this.FirstDisplayedScrollingRowIndex = _vScroll.Value; } catch {} } };
			this.Controls.Add(_vScroll);
			this.RowsAdded += (s,e) => UpdateScroll();
			this.RowsRemoved += (s,e) => UpdateScroll();
			this.Resize += (s,e) => UpdateScroll();
			this.MouseWheel += (s,e) => { if (_vScroll.Visible) _vScroll.Value -= (e.Delta / 120) * 3; };
		}
		private void UpdateScroll()
		{
			int visible = this.DisplayedRowCount(false);
			int total = this.RowCount;
			if (total > visible) {
				_vScroll.Visible = true;
				_vScroll.Maximum = total;
				_vScroll.LargeChange = visible;
				try { _vScroll.Value = this.FirstDisplayedScrollingRowIndex; } catch {}
			} else {
				_vScroll.Visible = false;
			}
		}
	}

	public class DarkColorTable : ProfessionalColorTable
	{
		public override Color MenuItemSelected { get { return Color.FromArgb(60, 60, 60); } }
		public override Color MenuItemBorder { get { return Color.FromArgb(60, 60, 60); } }
		public override Color MenuBorder { get { return Color.FromArgb(40, 40, 40); } }
		public override Color MenuItemSelectedGradientBegin { get { return Color.FromArgb(60, 60, 60); } }
		public override Color MenuItemSelectedGradientEnd { get { return Color.FromArgb(60, 60, 60); } }
		public override Color MenuItemPressedGradientBegin { get { return Color.FromArgb(60, 60, 60); } }
		public override Color MenuItemPressedGradientEnd { get { return Color.FromArgb(60, 60, 60); } }
		public override Color ImageMarginGradientBegin { get { return Color.FromArgb(30, 30, 30); } }
		public override Color ImageMarginGradientMiddle { get { return Color.FromArgb(30, 30, 30); } }
		public override Color ImageMarginGradientEnd { get { return Color.FromArgb(30, 30, 30); } }
		public override Color ToolStripDropDownBackground { get { return Color.FromArgb(30, 30, 30); } }
		public override Color SeparatorDark { get { return Color.FromArgb(60, 60, 60); } }
		public override Color SeparatorLight { get { return Color.FromArgb(60, 60, 60); } }
	}

	public class DarkRenderer : ToolStripProfessionalRenderer
	{
		public DarkRenderer() : base(new DarkColorTable()) { this.RoundedEdges = false; }

		protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
		{
			e.TextColor = Color.FromArgb(240, 240, 240);
			base.OnRenderItemText(e);
		}

		protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
		{
			e.ArrowColor = Color.FromArgb(240, 240, 240);
			base.OnRenderArrow(e);
		}
	}

	public class HotkeyManager : IDisposable
	{
		public const int WM_HOTKEY = 0x0312;
		private static int _nextId = 1;
		private static readonly MessageWindow _window = new MessageWindow();
		
		private static DateTime _lastHotkeyTime = DateTime.MinValue; 
		
		private static Dictionary<int, List<HotkeyActionEntry>> _hotkeyActions = new Dictionary<int, List<HotkeyActionEntry>>();
		private static Dictionary<Tuple<uint, uint>, int> _hotkeyIds = new Dictionary<Tuple<uint, uint>, int>();
		private static readonly object _hotkeyActionsLock = new object();
		private static bool _isProcessingHotkey = false;

		private const int WH_MOUSE_LL = 14;
		private const int WM_LBUTTONDOWN = 0x0201;
		private const int WM_RBUTTONDOWN = 0x0204;
		private const int WM_MBUTTONDOWN = 0x0207;
		private const int WM_XBUTTONDOWN = 0x020B;

		private const uint MOD_ALT = 0x0001;
		private const uint MOD_CONTROL = 0x0002;
		private const uint MOD_SHIFT = 0x0004;
		private const uint MOD_WIN = 0x0008;

		private static LowLevelMouseProc _mouseProcDelegate;
		private static IntPtr _mouseHookID = IntPtr.Zero;
		private static int _mouseHotkeysCount = 0;

		[StructLayout(LayoutKind.Sequential)]
		private struct MSLLHOOKSTRUCT
		{
			public System.Drawing.Point pt;
			public uint mouseData;
			public uint flags;
			public uint time;
			public IntPtr dwExtraInfo;
		}

		private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool UnhookWindowsHookEx(IntPtr hhk);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr GetModuleHandle(string lpModuleName);
		
		[DllImport("user32.dll")]
		static extern short GetAsyncKeyState(int vKey);

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

		private static bool IsMouseKey(uint virtualKey)
		{
			return virtualKey == 0x01 || virtualKey == 0x02 || virtualKey == 0x04 || virtualKey == 0x05 || virtualKey == 0x06;
		}

		private static void EnsureMouseHook()
		{
			if (_mouseHookID == IntPtr.Zero)
			{
				_mouseProcDelegate = MouseHookCallback;
				using (Process curProcess = Process.GetCurrentProcess())
				using (ProcessModule curModule = curProcess.MainModule)
				{
					_mouseHookID = SetWindowsHookEx(WH_MOUSE_LL, _mouseProcDelegate, GetModuleHandle(curModule.ModuleName), 0);
				}
			}
		}

		private static void ReleaseMouseHook()
		{
			if (_mouseHookID != IntPtr.Zero)
			{
				UnhookWindowsHookEx(_mouseHookID);
				_mouseHookID = IntPtr.Zero;
			}
		}

		private static IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
		{
			if (nCode >= 0 && !_isProcessingHotkey)
			{
				MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
				
				if ((hookStruct.flags & 0x01) == 0) 
				{
					uint vkCode = 0;
					if ((int)wParam == WM_LBUTTONDOWN) vkCode = 0x01;
					else if ((int)wParam == WM_RBUTTONDOWN) vkCode = 0x02;
					else if ((int)wParam == WM_MBUTTONDOWN) vkCode = 0x04;
					else if ((int)wParam == WM_XBUTTONDOWN)
					{
						uint xButton = hookStruct.mouseData >> 16;
						if (xButton == 1) vkCode = 0x05; 
						else if (xButton == 2) vkCode = 0x06; 
					}

					if (vkCode != 0)
					{
						
						uint modifiers = 0;
						if ((GetAsyncKeyState(0x12) & 0x8000) != 0) modifiers |= MOD_ALT;
						if ((GetAsyncKeyState(0x11) & 0x8000) != 0) modifiers |= MOD_CONTROL;
						if ((GetAsyncKeyState(0x10) & 0x8000) != 0) modifiers |= MOD_SHIFT;
						if (((GetAsyncKeyState(0x5B) & 0x8000) != 0) || ((GetAsyncKeyState(0x5C) & 0x8000) != 0)) modifiers |= MOD_WIN;

						var hotkeyTuple = Tuple.Create(modifiers, vkCode);

						lock (_hotkeyActionsLock)
						{
							if (_hotkeyIds.ContainsKey(hotkeyTuple))
							{
								_isProcessingHotkey = true;
								try
								{
									_lastHotkeyTime = DateTime.Now;
									int id = _hotkeyIds[hotkeyTuple];
									if (_hotkeyActions.ContainsKey(id) && _hotkeyActions[id] != null)
									{
										var actionsToInvokeCopy = new List<HotkeyActionEntry>(_hotkeyActions[id]);
										
										var globalToggleActions = new List<HotkeyActionEntry>();
										var normalActions = new List<HotkeyActionEntry>();

										foreach (var entry in actionsToInvokeCopy)
										{
											if (entry.Type == HotkeyActionType.GlobalToggle) globalToggleActions.Add(entry);
											else normalActions.Add(entry);
										}

										if (globalToggleActions.Count > 0)
										{
											if (globalToggleActions[0].ActionDelegate != null) globalToggleActions[0].ActionDelegate.Invoke();
										}
										else if (!AreHotkeysGloballyPaused)
										{
											foreach (var entry in normalActions)
											{
												if (entry.ActionDelegate != null) entry.ActionDelegate.Invoke();
											}
										}
									}
								}
								finally
								{
									_isProcessingHotkey = false;
								}
							}
						}
					}
				}
			}
			return CallNextHookEx(_mouseHookID, nCode, wParam, lParam);
		}

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
				if(IsMouseKey(virtualKey))
				{
					EnsureMouseHook();
					_mouseHotkeysCount++;
				}
				else
				{
					if (!RegisterHotKey(_window.Handle, id, modifiers, virtualKey))
					{
						throw new Exception("Failed to register hotkey. Error: " + Marshal.GetLastWin32Error());
					}
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
					if(IsMouseKey(keyToRemove.Item2))
					{
						_mouseHotkeysCount--;
						if(_mouseHotkeysCount <= 0)
						{
							ReleaseMouseHook();
							_mouseHotkeysCount = 0;
						}
					}
					else
					{
						UnregisterHotKey(_window.Handle, id);
					}
					_hotkeyIds.Remove(keyToRemove);
				}
				_hotkeyActions.Remove(id);
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
					Unregister(id);
				}
			}
		}

		public static void UnregisterAll()
		{
			lock (_hotkeyActionsLock)
			{
				foreach (var pair in _hotkeyIds)
				{
					if (!IsMouseKey(pair.Key.Item2))
					{
						UnregisterHotKey(_window.Handle, pair.Value);
					}
				}
				ReleaseMouseHook();
				_mouseHotkeysCount = 0;
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
					cp.Parent = (IntPtr)(-3); 
					return cp;
				}
			}

			protected override void WndProc(ref Message m)
			{
				if (m.Msg == WM_HOTKEY)
				{

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

	public class FolderBrowser
	{
		public string SelectedPath { get; set; }
		public string Description { get; set; }
		
		public DialogResult ShowDialog(IWin32Window owner)
		{
			var ofd = new OpenFileDialog();
			ofd.FileName = "Select Folder";
			ofd.Filter = "Folders|";
			ofd.CheckFileExists = false;
			ofd.CheckPathExists = true;
			ofd.Title = Description;
			if (!string.IsNullOrEmpty(SelectedPath)) ofd.InitialDirectory = SelectedPath;

			var type = ofd.GetType();
			var createVistaDialog = type.GetMethod("CreateVistaDialog", BindingFlags.Instance | BindingFlags.NonPublic);
			
			if (createVistaDialog != null)
			{
				var dialog = createVistaDialog.Invoke(ofd, null);
				var dialogType = dialog.GetType();
				var setOptions = dialogType.GetMethod("SetOptions");
				var getOptions = dialogType.GetMethod("GetOptions");
				
				if (setOptions != null && getOptions != null)
				{
					uint options = (uint)getOptions.Invoke(dialog, null);
					setOptions.Invoke(dialog, new object[] { options | 0x20 });
					
					var show = dialogType.GetMethod("Show");
					int result = (int)show.Invoke(dialog, new object[] { owner.Handle });
					
					if (result == 0)
					{
						var getResult = dialogType.GetMethod("GetResult");
						var item = getResult.Invoke(dialog, null);
						var getDisplayName = item.GetType().GetMethod("GetDisplayName");
						SelectedPath = (string)getDisplayName.Invoke(item, new object[] { 0x80058000 });
						return DialogResult.OK;
					}
					return DialogResult.Cancel;
				}
			}
			
			using (var fbd = new FolderBrowserDialog())
			{
				fbd.Description = Description;
				fbd.SelectedPath = SelectedPath;
				var res = fbd.ShowDialog(owner);
				if (res == DialogResult.OK) SelectedPath = fbd.SelectedPath;
				return res;
			}
		}
	}

	public class ColorWriter
	{
					public static void WriteColored(string message, string color)
					{
						ConsoleColor originalColor = Console.ForegroundColor;
						switch(color.ToLower()) {
							case "darkgray": Console.ForegroundColor = ConsoleColor.DarkGray; break;
							case "yellow": Console.ForegroundColor = ConsoleColor.Yellow; break;
							case "red": Console.ForegroundColor = ConsoleColor.Red; break;
							case "cyan": Console.ForegroundColor = ConsoleColor.Cyan; break;
							case "green": Console.ForegroundColor = ConsoleColor.Green; break;
							default: Console.ForegroundColor = ConsoleColor.DarkGray; break;
						}
						Console.WriteLine(message);
						Console.ForegroundColor = originalColor;
					}
	}

	public class SafeWindowCore 
	{
			[DllImport("user32.dll", SetLastError = true)]
			public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam, StringBuilder lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);

			[DllImport("user32.dll")]
			[return: MarshalAs(UnmanagedType.Bool)]
			private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

			[DllImport("user32.dll", SetLastError = true)]
			private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

			[DllImport("user32.dll")]
			[return: MarshalAs(UnmanagedType.Bool)]
			private static extern bool IsWindowVisible(IntPtr hWnd);

			[DllImport("user32.dll", EntryPoint = "GetWindowLong")]
			private static extern IntPtr GetWindowLongPtr32(IntPtr hWnd, int nIndex);

			[DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
			private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

			[DllImport("user32.dll", SetLastError = true)]
			private static extern bool GetWindowInfo(IntPtr hWnd, ref WINDOWINFO pwi);

			[DllImport("user32.dll")]
			private static extern IntPtr GetForegroundWindow();

			private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

			[StructLayout(LayoutKind.Sequential)]
			public struct WINDOWINFO 
			{
				public uint cbSize;
				public RECT rcWindow;
				public RECT rcClient;
				public uint dwStyle;
				public uint dwExStyle;
				public uint dwWindowStatus;
				public uint cxWindowBorders;
				public uint cyWindowBorders;
				public ushort atomWindowType;
				public ushort wCreatorVersion;
			}
			[StructLayout(LayoutKind.Sequential)]
			public struct RECT {
				public int Left, Top, Right, Bottom;
			}

			public static string GetText(IntPtr hWnd) {
				if (hWnd == IntPtr.Zero) return null;
				const uint WM_GETTEXT = 0x000D;
				const uint SMTO_ABORTIFHUNG = 0x0002;
				UIntPtr result;
				StringBuilder sb = new StringBuilder(512);
				IntPtr ret = SendMessageTimeout(hWnd, WM_GETTEXT, (UIntPtr)512, sb, SMTO_ABORTIFHUNG, 50, out result);
				if (ret == IntPtr.Zero) return null;
				return sb.ToString();
			}

			public static IntPtr FindBestWindow(int pid) {
				IntPtr bestHandle = IntPtr.Zero;
				EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
					uint windowPid;
					GetWindowThreadProcessId(hWnd, out windowPid);
					if (windowPid == pid) {
						if (IsWindowVisible(hWnd)) {
							bestHandle = hWnd;
							return false;
						}
					}
					return true;
				}, IntPtr.Zero);
				return bestHandle;
			}

			public static bool IsMinimized(IntPtr hWnd) {
				const int GWL_STYLE = -16;
				const long WS_MINIMIZE = 0x20000000;
				IntPtr ptrVal;
				if (IntPtr.Size == 8) ptrVal = GetWindowLongPtr64(hWnd, GWL_STYLE);
				else ptrVal = GetWindowLongPtr32(hWnd, GWL_STYLE);
				long style = ptrVal.ToInt64();
				return (style & WS_MINIMIZE) == WS_MINIMIZE;
			}

			public static bool IsWindowFlashing(IntPtr hWnd) {
				if (hWnd == IntPtr.Zero) return false;
				if (GetForegroundWindow() == hWnd) return false;
				WINDOWINFO pwi = new WINDOWINFO();
				pwi.cbSize = (uint)Marshal.SizeOf(pwi);
				if (GetWindowInfo(hWnd, ref pwi)) {
					return (pwi.dwWindowStatus & 0x0001) != 0;
				}
				return false;
			}
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct Win32Point {
		public int X;
		public int Y;
	}

	public class Win32MouseUtils 
	{
		[DllImport("user32.dll")]
		public static extern bool GetCursorPos(out Win32Point lpPoint);

		[DllImport("user32.dll")]
		public static extern bool ScreenToClient(IntPtr hWnd, ref Win32Point lpPoint);

		[DllImport("user32.dll")]
		public static extern short GetAsyncKeyState(int vKey);

		public static IntPtr MakeLParam(int x, int y) {
			return (IntPtr)((y << 16) | (x & 0xFFFF));
		}

		public static IntPtr GetCurrentInputStateWParam(int targetButtonFlag, bool isDown, int virtualKeyToExclude) {
			int wParam = 0;

			const int VK_LBUTTON  = 0x01;
			const int VK_RBUTTON  = 0x02;
			const int VK_MBUTTON  = 0x04;
			const int VK_XBUTTON1 = 0x05;
			const int VK_XBUTTON2 = 0x06;
			const int VK_SHIFT    = 0x10;
			const int VK_CONTROL  = 0x11;

			const int MK_LBUTTON  = 0x0001;
			const int MK_RBUTTON  = 0x0002;
			const int MK_SHIFT    = 0x0004;
			const int MK_CONTROL  = 0x0008;
			const int MK_MBUTTON  = 0x0010;
			const int MK_XBUTTON1 = 0x0020;
			const int MK_XBUTTON2 = 0x0040;

			if (virtualKeyToExclude != VK_LBUTTON && (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0) wParam |= MK_LBUTTON;
			if (virtualKeyToExclude != VK_RBUTTON && (GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0) wParam |= MK_RBUTTON;
			if (virtualKeyToExclude != VK_MBUTTON && (GetAsyncKeyState(VK_MBUTTON) & 0x8000) != 0) wParam |= MK_MBUTTON;
			if (virtualKeyToExclude != VK_XBUTTON1 && (GetAsyncKeyState(VK_XBUTTON1) & 0x8000) != 0) wParam |= MK_XBUTTON1;
			if (virtualKeyToExclude != VK_XBUTTON2 && (GetAsyncKeyState(VK_XBUTTON2) & 0x8000) != 0) wParam |= MK_XBUTTON2;
			
			if ((GetAsyncKeyState(VK_SHIFT)   & 0x8000) != 0) wParam |= MK_SHIFT;
			if ((GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0) wParam |= MK_CONTROL;

			if (isDown) wParam |= targetButtonFlag;
			else wParam &= ~targetButtonFlag;

			return (IntPtr)wParam;
		}
	}
	
}
"@ 
#endregion

#region Module Initialization
function InitializeClassesModule
{
	[CmdletBinding()]
	param()
	try
	{
		
		if ('Custom.Native' -as [Type])
		{
			Write-Verbose 'System integration classes already initialized.'
			return $true
		}
		Write-Verbose 'Initializing system integration classes...'
		$addTypeArgs = @{
			TypeDefinition       = $classes
			Language             = $script:CompilationOptions.Language
			ReferencedAssemblies = $script:CompilationOptions.ReferencedAssemblies
			WarningAction        = $script:CompilationOptions.WarningAction
			ErrorAction          = 'Stop'
		}
		Add-Type @addTypeArgs
		Write-Verbose 'System integration classes initialized successfully.'
		return $true
	}
	catch
	{
		Write-Error "Failed to initialize/compile system integration classes. Error: $($_.Exception.Message)"
		return $false
	}
}
InitializeClassesModule
Import-Module NetTCPIP 4>$null
#endregion 

#region Module Exports
Export-ModuleMember -Function *
#endregion