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
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Reflection;

namespace Custom
{
	
	public static class ColorWheelHelper 
	{
		public static void RenderWheel(Bitmap bmp, double lightness) 
		{
			int width = bmp.Width;
			int height = bmp.Height;
			double center = width / 2.0;
			double maxRadius = center - 1.0;

			int[] pixels = new int[width * height];

			for (int y = 0; y < height; y++) 
			{
				int rowOffset = y * width;
				double dy = y - center;

				for (int x = 0; x < width; x++) 
				{
					double dx = x - center;
					double dist = Math.Sqrt(dx * dx + dy * dy);

					if (dist <= maxRadius) 
					{
						double angle = Math.Atan2(dy, dx) * 180.0 / Math.PI;
						if (angle < 0) angle += 360.0;
						double sat = dist / maxRadius;

						double c = (1.0 - Math.Abs(2.0 * lightness - 1.0)) * sat;
						double hp = angle / 60.0;
						double xVal = c * (1.0 - Math.Abs((hp % 2.0) - 1.0));
						double m = lightness - c / 2.0;

						double r = 0, g = 0, b = 0;
						if      (angle < 60)  { r = c;    g = xVal; }
						else if (angle < 120) { r = xVal; g = c;    }
						else if (angle < 180) { g = c;    b = xVal; }
						else if (angle < 240) { g = xVal; b = c;    }
						else if (angle < 300) { r = xVal; b = c;    }
						else                  { r = c;    b = xVal; }

						int ir = (int)Math.Min(255, (r + m) * 255);
						int ig = (int)Math.Min(255, (g + m) * 255);
						int ib = (int)Math.Min(255, (b + m) * 255);

						pixels[rowOffset + x] = (255 << 24) | (ir << 16) | (ig << 8) | ib;
					} 
					else 
					{
						pixels[rowOffset + x] = (255 << 24) | (30 << 16) | (33 << 8) | 40;
					}
				}
			}

			BitmapData data = bmp.LockBits(new Rectangle(0, 0, width, height),
				ImageLockMode.WriteOnly, PixelFormat.Format32bppRgb);
			Marshal.Copy(pixels, 0, data.Scan0, pixels.Length);
			bmp.UnlockBits(data);
		}
	}

	public class NoFocusForm : Form
	{
		protected override bool ShowWithoutActivation
		{
			get { return true; }
		}

		private const int WM_MOUSEACTIVATE = 0x0021;
		private const int MA_NOACTIVATE = 3;
		private const int WS_EX_NOACTIVATE = 0x08000000;
		private const int WS_EX_TOOLWINDOW = 0x00000080;

		protected override CreateParams CreateParams
		{
			get
			{
				var cp = base.CreateParams;
				cp.ExStyle |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
				return cp;
			}
		}

		protected override void WndProc(ref Message m)
		{
			if (m.Msg == WM_MOUSEACTIVATE)
			{
				m.Result = (IntPtr)MA_NOACTIVATE;
				return;
			}
			base.WndProc(ref m);
		}
	}
	
	public class DarkTreeView : TreeView
	{
		[DllImport("dwmapi.dll")]
		private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

		[DllImport("uxtheme.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
		private static extern int SetWindowTheme(IntPtr hwnd, string pszSubAppName, string pszSubIdList);

		private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;

		private Color _selectionColor = Color.FromArgb(45, 50, 65);
		public Color SelectionColor { get { return _selectionColor; } set { _selectionColor = value; } }

		public DarkTreeView()
		{
			this.SelectionColor = Color.FromArgb(45, 50, 65);
			this.DrawMode = TreeViewDrawMode.OwnerDrawText;
			this.DoubleBuffered = true;
			this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
			
			this.BackColor = Color.FromArgb(22, 26, 38); 
			this.ForeColor = Color.White;
			this.BorderStyle = BorderStyle.None;
			this.HideSelection = false;
            this.FullRowSelect = true;
            this.ShowLines = false;
            this.ShowPlusMinus = true;
		}

		protected override void OnHandleCreated(EventArgs e)
		{
			base.OnHandleCreated(e);
			
			SetWindowTheme(this.Handle, "Explorer", null);

			int darkMode = 1;
			if (DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkMode, 4) != 0)
            {
                 DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref darkMode, 4);
            }
		}

        protected override void OnDrawNode(DrawTreeNodeEventArgs e)
        {
            if (e.Node == null) return;

            Font font = e.Node.NodeFont ?? e.Node.TreeView.Font;
            Color fore = e.Node.TreeView.ForeColor;

            if ((e.State & TreeNodeStates.Selected) != 0)
            {
                using (SolidBrush brush = new SolidBrush(SelectionColor))
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

            TextRenderer.DrawText(e.Graphics, e.Node.Text, font, e.Bounds, fore, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
        }
	}

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

		private readonly Color _backColor = Color.FromArgb(13, 17, 23);
		private readonly Color _secondaryColor = Color.FromArgb(22, 27, 34);
		private readonly Color _accentColor = Color.FromArgb(31, 111, 235); 
		private readonly Color _textColor = Color.FromArgb(230, 237, 243);
		private bool _isSuccess = false;

		public DarkMessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, bool isSuccess = false)
		{
			this.SuspendLayout();

			this._isSuccess = isSuccess;
			
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

			if (icon != MessageBoxIcon.None || _isSuccess)
			{
				PictureBox iconBox = new PictureBox {
					Size = new Size(32, 32),
					Margin = new Padding(0, 0, 20, 0),
					Image = GetModernIcon(icon, _isSuccess)
				};
				mainLayout.Controls.Add(iconBox, 0, 0);
			}

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

		private static readonly Dictionary<string, Bitmap> _modernIconCache = new Dictionary<string, Bitmap>();
		private static readonly object _modernIconLock = new object();

		private Bitmap GetModernIcon(MessageBoxIcon icon, bool isSuccess = false)
		{
			string cacheKey = icon.ToString() + (isSuccess ? "|success" : "|plain");
			lock (_modernIconLock)
			{
				Bitmap cached;
				if (_modernIconCache.TryGetValue(cacheKey, out cached)) { return cached; }
				Bitmap built = BuildModernIcon(icon, isSuccess);
				_modernIconCache[cacheKey] = built;
				return built;
			}
		}

		private Bitmap BuildModernIcon(MessageBoxIcon icon, bool isSuccess = false)
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

				using (SolidBrush fillBrush = new SolidBrush(iconColor))
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
						using (GraphicsPath qPath = new GraphicsPath())
						{
							qPath.AddArc(11, 9, 10, 8, 180, 180); 
							qPath.AddLine(21, 13, 16, 15);
							qPath.AddLine(16, 15, 16, 19);
							g.DrawPath(p, qPath);
						}
						g.FillEllipse(fillBrush, centerX - 1.5f, 22, 3, 3);
					}
					else if (icon == MessageBoxIcon.Exclamation || icon == MessageBoxIcon.Warning)
					{
						g.DrawLine(p, centerX, 9, centerX, 19);
						g.FillEllipse(fillBrush, centerX - 1.5f, 22, 3, 3);
					}
					else if (icon == MessageBoxIcon.Error || icon == MessageBoxIcon.Hand || icon == MessageBoxIcon.Stop)
					{
						g.DrawLine(p, 11, 11, 21, 21);
						g.DrawLine(p, 21, 11, 11, 21);
					}
					else 
					{
						g.FillEllipse(fillBrush, centerX - 1.5f, 8, 3, 3);
						g.DrawLine(p, centerX, 13, centerX, 23);
					}
				}
			}
			return bmp;
		}

		private void AddButton(FlowLayoutPanel panel, string text, DialogResult result, bool isPrimary = false)
		{
			Color btnBack = isPrimary ? Color.FromArgb(31, 111, 235) : Color.FromArgb(33, 38, 45);
			Color btnHover = isPrimary ? Color.FromArgb(58, 123, 213) : Color.FromArgb(48, 54, 61);

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

			TableLayoutPanel mainLayout = new TableLayoutPanel
			{
				Dock = DockStyle.Fill,
				ColumnCount = 1,
				RowCount = 3,
				Padding = new Padding(20)
			};
			mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
			mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
			mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
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

	public class ProcessData : IComparable<ProcessData>
	{
		public int Id;
		public string Name;
		public IntPtr MainWindowHandle;
		public DateTime StartTime;
		public string Profile;
		public string Title;
		public bool IsResponsive;
		public bool IsMinimized;

		public int CompareTo(ProcessData other)
		{
			if (other == null) return 1;
			int res = string.Compare(this.Profile, other.Profile, StringComparison.OrdinalIgnoreCase);
			if (res != 0) return res;
			return this.StartTime.CompareTo(other.StartTime);
		}
	}

	public static class ProcessHelper
	{
		private static Dictionary<int, IntPtr> _handleCache = new Dictionary<int, IntPtr>();
		private static Dictionary<int, string> _pathCache = new Dictionary<int, string>();
		private static Dictionary<int, string> _profileCache = new Dictionary<int, string>();
		private static Dictionary<int, string> _titleCache = new Dictionary<int, string>();
		private static Dictionary<int, DateTime> _startTimeCache = new Dictionary<int, DateTime>();

		public static List<ProcessData> GetProcessDataList(string processName, IDictionary profileMap)
		{
			var procs = Process.GetProcessesByName(processName);
			var list = new List<ProcessData>();
			var newHandleCache = new Dictionary<int, IntPtr>();
			var newPathCache = new Dictionary<int, string>();
			var newProfileCache = new Dictionary<int, string>();
			var newTitleCache = new Dictionary<int, string>();
			var newStartTimeCache = new Dictionary<int, DateTime>();
			
			var locker = new object();

			Parallel.ForEach(procs, p => {
				try
				{
					if (p.HasExited) return;
					int pid = p.Id;
					DateTime startTime = p.StartTime;
					IntPtr hWnd = p.MainWindowHandle;

					bool isSameProcess = false;
					lock (_startTimeCache) {
						if (_startTimeCache.ContainsKey(pid) && _startTimeCache[pid] == startTime) isSameProcess = true;
					}

					if (hWnd == IntPtr.Zero) {
						if (isSameProcess) {
							lock (_handleCache) {
								if (_handleCache.ContainsKey(pid) && Native.IsWindow(_handleCache[pid])) hWnd = _handleCache[pid];
							}
						}
						if (hWnd == IntPtr.Zero) hWnd = SafeWindowCore.FindBestWindow(pid);
					}

					string profile = "";
					if (isSameProcess) {
						lock (_profileCache) { if (_profileCache.ContainsKey(pid)) profile = _profileCache[pid]; }
					}
					
					if (string.IsNullOrEmpty(profile)) {
						string exePath = Native.GetProcessPathById(pid);
						if (!string.IsNullOrEmpty(exePath)) {
							string dir = Path.GetDirectoryName(exePath).TrimEnd('\\', '/') + "\\";
							foreach (DictionaryEntry entry in profileMap) {
								string pPath = ((string)entry.Value).TrimEnd('\\', '/') + "\\";
								if (dir.StartsWith(pPath, StringComparison.OrdinalIgnoreCase)) {
									profile = (string)entry.Key;
									break;
								}
							}
						}
					}

					string title = p.ProcessName;
					if (isSameProcess) {
						lock (_titleCache) { if (_titleCache.ContainsKey(pid)) title = _titleCache[pid]; }
					}

					bool isResponsive = true;
					bool isMinimized = false;

					if (hWnd != IntPtr.Zero) {
						isResponsive = Native.Responsive(hWnd, 50);
						isMinimized = SafeWindowCore.IsMinimized(hWnd);
						
						string fetchedTitle = SafeWindowCore.GetText(hWnd);
						if (!string.IsNullOrEmpty(fetchedTitle)) title = fetchedTitle;
					}

					var data = new ProcessData {
						Id = pid,
						Name = p.ProcessName,
						MainWindowHandle = hWnd,
						StartTime = startTime,
						Profile = profile,
						Title = title,
						IsResponsive = isResponsive,
						IsMinimized = isMinimized
					};

					lock (locker) {
						newHandleCache[pid] = hWnd;
						newPathCache[pid] = profile;
						newProfileCache[pid] = profile;
						newTitleCache[pid] = title;
						newStartTimeCache[pid] = startTime;
						list.Add(data);
					}
				}
				catch {}
				finally { p.Dispose(); }
			});

			_handleCache = newHandleCache;
			_profileCache = newProfileCache;
			_titleCache = newTitleCache;
			_startTimeCache = newStartTimeCache;

			list.Sort();
			return list;
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
		[DllImport("user32.dll", SetLastError = true)]
		public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

		[DllImport("user32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsIconic(IntPtr hWnd);

		public static void SendClick(IntPtr hWnd, int x, int y)
		{
			Point pt = new Point(x, y);
			ClientToScreen(hWnd, ref pt);
			SetCursorPos(pt.X, pt.Y);
			System.Threading.Thread.Sleep(10);

			IntPtr lParam = (IntPtr)(((uint)(y & 0xFFFF) << 16) | (uint)(x & 0xFFFF));
			IntPtr lParamScreen = (IntPtr)(((uint)(pt.Y & 0xFFFF) << 16) | (uint)(pt.X & 0xFFFF));

			const uint WM_MOUSEMOVE     = 0x0200;
			const uint WM_LBUTTONDOWN   = 0x0201;
			const uint WM_LBUTTONUP     = 0x0202;
			const uint WM_MOUSEACTIVATE = 0x0021;
			const uint WM_SETFOCUS      = 0x0007;
			const int  HTCLIENT         = 1;

			SendMessage(hWnd, (int)WM_MOUSEACTIVATE, (int)hWnd, (int)(((uint)WM_LBUTTONDOWN << 16) | HTCLIENT));
			SendMessage(hWnd, (int)WM_SETFOCUS, 0, 0);

			SendMessage(hWnd, (int)WM_MOUSEMOVE, 0, (int)lParam);
			SendMessage(hWnd, (int)WM_LBUTTONDOWN, 1, (int)lParam);
			System.Threading.Thread.Sleep(15);
			SendMessage(hWnd, (int)WM_LBUTTONUP, 0, (int)lParam);
		}


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

		public static Dictionary<int, int> GetTcpConnectionCounts()
		{
			var pidCounts = new Dictionary<int, int>();
			var pidIpMap = new Dictionary<int, Dictionary<string, int>>();
			
			int[] families = new int[] { 2, 23 };

			foreach (int family in families)
			{
				int buffSize = 0;
				uint ret = GetExtendedTcpTable(IntPtr.Zero, ref buffSize, false, family, 5, 0);
				
				for (int retry = 0; retry < 5; retry++)
				{
					if (buffSize == 0) buffSize = 32768;
					IntPtr tcpTablePtr = Marshal.AllocHGlobal(buffSize);
					try
					{
						ret = GetExtendedTcpTable(tcpTablePtr, ref buffSize, false, family, 5, 0);
						if (ret == 0)
						{
							int entryCount = Marshal.ReadInt32(tcpTablePtr);
							IntPtr rowPtr = (IntPtr)((long)tcpTablePtr + 4);

							if (family == 2)
							{
								for (int i = 0; i < entryCount; i++)
								{
									int state = Marshal.ReadInt32(rowPtr, 0);
									if (state == 5)
									{
										int pid = Marshal.ReadInt32(rowPtr, 20);
										uint remoteAddr = (uint)Marshal.ReadInt32(rowPtr, 12);
										string ip = remoteAddr.ToString();

										if (!pidIpMap.ContainsKey(pid)) pidIpMap[pid] = new Dictionary<string, int>();
										if (!pidIpMap[pid].ContainsKey(ip)) pidIpMap[pid][ip] = 0;
										pidIpMap[pid][ip]++;
									}
									rowPtr = (IntPtr)((long)rowPtr + 24);
								}
							}
							else
							{
								for (int i = 0; i < entryCount; i++)
								{
									int state = Marshal.ReadInt32(rowPtr, 48);
									if (state == 5)
									{
										int pid = Marshal.ReadInt32(rowPtr, 52);
										byte[] remoteIpBytes = new byte[16];
										Marshal.Copy((IntPtr)((long)rowPtr + 24), remoteIpBytes, 0, 16);
										string ip = string.Join(":", remoteIpBytes.Select(b => b.ToString("x2")));

										if (!pidIpMap.ContainsKey(pid)) pidIpMap[pid] = new Dictionary<string, int>();
										if (!pidIpMap[pid].ContainsKey(ip)) pidIpMap[pid][ip] = 0;
										pidIpMap[pid][ip]++;
									}
									rowPtr = (IntPtr)((long)rowPtr + 56); 
								}
							}
							break;
						}
						else if (ret == 122)
						{
							continue;
						}
						else
						{
							break;
						}
					}
					catch { break; }
					finally { Marshal.FreeHGlobal(tcpTablePtr); }
				}
			}

			foreach (var pidEntry in pidIpMap)
			{
				int totalValidConnections = 0;
				foreach (var ipEntry in pidEntry.Value)
				{
					if (ipEntry.Value >= 2)
					{
						totalValidConnections += ipEntry.Value;
					}
				}
				if (totalValidConnections > 0)
				{
					pidCounts[pidEntry.Key] = totalValidConnections;
				}
			}

			return pidCounts;
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

		public static bool EmptyWorkingSet(int processId)
		{
			IntPtr hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_SET_QUOTA, false, processId);
			if (hProcess == IntPtr.Zero) hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_SET_QUOTA, false, processId);
			if (hProcess == IntPtr.Zero) return false;
			try { return EmptyWorkingSet(hProcess); }
			finally { CloseHandle(hProcess); }
		}

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

		[DllImport("user32.dll", EntryPoint = "IsWindowActive")]
		public static extern bool IsWindowActive(IntPtr windowHandle);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsWindowVisible(IntPtr hWnd);

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
		public const uint PROCESS_SET_QUOTA = 0x0100;

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
			public Color BackgroundColor   { get { return _backgroundColor;   } set { _backgroundColor   = value; Invalidate(); } }
			public Color TextColor         { get { return _textColor;         } set { _textColor         = value; Invalidate(); } }
			public Color BorderColor       { get { return _borderColor;       } set { _borderColor       = value; Invalidate(); } }
			public Color ArrowColor        { get { return _arrowColor;        } set { _arrowColor        = value; Invalidate(); } }
			public Color SelectedItemBackColor { get { return _selectedItemBackColor; } set { _selectedItemBackColor = value; Invalidate(); } }
			public Color SelectedItemForeColor { get { return _selectedItemForeColor; } set { _selectedItemForeColor = value; Invalidate(); } }
			public Color DropDownBackColor { get { return _dropDownBackColor; } set { _dropDownBackColor = value; Invalidate(); } }
			public Color HeaderColor { get { return _headerColor; } set { _headerColor = value; Invalidate(); } }

			private System.Collections.Generic.Dictionary<int, Image> _itemImages = new System.Collections.Generic.Dictionary<int, Image>();
			public System.Collections.Generic.Dictionary<int, Image> ItemImages
			{
				get { return _itemImages; }
				set { _itemImages = value; Invalidate(); }
			}

			private System.Collections.Generic.Dictionary<string, string> _displayNames = new System.Collections.Generic.Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
			public System.Collections.Generic.Dictionary<string, string> DisplayNames
			{
				get { return _displayNames; }
				set { _displayNames = value; Invalidate(); }
			}

			private System.Collections.Generic.HashSet<int> _headerIndices = new System.Collections.Generic.HashSet<int>();
			public System.Collections.Generic.HashSet<int> HeaderIndices
			{
				get { return _headerIndices; }
				set
				{
					_headerIndices = value ?? new System.Collections.Generic.HashSet<int>();
					Invalidate();
				}
			}

			public void SetHeaderIndices(int[] indices)
			{
				_headerIndices = new System.Collections.Generic.HashSet<int>();
				if (indices != null) foreach (int i in indices) _headerIndices.Add(i);
				Invalidate();
			}

			public bool SuppressChanges { get; set; }

			private Font _headerFont;
			private Font HeaderFont
			{
				get
				{
					if (_headerFont == null)
						_headerFont = new Font(this.Font.FontFamily, Math.Max(6.5f, this.Font.Size - 1f), FontStyle.Bold);
					return _headerFont;
				}
			}

			protected override void OnFontChanged(EventArgs e)
			{
				if (_headerFont != null) { _headerFont.Dispose(); _headerFont = null; }
				base.OnFontChanged(e);
			}

			private int _lastValidIndex = -1;
			private bool _suppressHeaderRevert;

			protected override void OnSelectedIndexChanged(EventArgs e)
			{
				int i = this.SelectedIndex;
				bool header = i >= 0 && _headerIndices != null && _headerIndices.Contains(i);
				if (i < 0) _lastValidIndex = -1;
				else if (!header) _lastValidIndex = i;
				if (SuppressChanges) return;
				if (header)
				{
					if (_suppressHeaderRevert) return;
					int back = _lastValidIndex;
					if (back < -1 || back >= this.Items.Count) back = -1;
					_suppressHeaderRevert = true;
					try { this.SelectedIndex = back; }
					finally { _suppressHeaderRevert = false; }
					return;
				}
				base.OnSelectedIndexChanged(e);
			}

			public void FitDropDownWidth(int maxWidth)
			{
				int w = this.Width;
				for (int i = 0; i < this.Items.Count; i++)
				{
					bool head = _headerIndices != null && _headerIndices.Contains(i);
					System.Drawing.Size sz = TextRenderer.MeasureText(GetItemDisplayText(this.Items[i]), head ? HeaderFont : this.Font);
					if (sz.Width + 16 > w) w = sz.Width + 16;
				}
				if (maxWidth > 0 && w > maxWidth) w = maxWidth;
				this.DropDownWidth = w;
			}

			private string GetItemDisplayText(object item)
			{
				string t = item == null ? "" : item.ToString();
				if (_displayNames != null && _displayNames.ContainsKey(t)) return _displayNames[t];
				return t;
			}

			private Image GetItemImage(int index)
			{
				Image img;
				if (_itemImages != null && _itemImages.TryGetValue(index, out img)) return img;
				return null;
			}

			private Color _backgroundColor      = Color.FromArgb(22, 27, 34);   
			private Color _textColor            = Color.FromArgb(230, 237, 243); 
			private Color _borderColor          = Color.FromArgb(48, 54, 61);   
			private Color _arrowColor           = Color.FromArgb(230, 237, 243); 
			private Color _selectedItemBackColor= Color.FromArgb(31, 111, 235);  
			private Color _selectedItemForeColor= Color.FromArgb(255, 255, 255); 
			private Color _dropDownBackColor    = Color.FromArgb(22, 27, 34);   
			private Color _headerColor          = Color.FromArgb(139, 148, 158);

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
					string text = GetItemDisplayText(this.Text);
					if (!string.IsNullOrEmpty(text))
					{
						Image img = GetItemImage(this.SelectedIndex);
						int textX = bounds.Left;
						if (img != null)
						{
							int iy = bounds.Top + (bounds.Height - 16) / 2;
							g.DrawImage(img, bounds.Left + 2, iy, 16, 16);
							textX = bounds.Left + 20;
						}
						Rectangle textRect = new Rectangle(textX, bounds.Top, bounds.Width - (textX - bounds.Left), bounds.Height);
						TextRenderer.DrawText(g,text, this.Font, textRect, this.Enabled ? TextColor : SystemColors.GrayText, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
					}
				}
				using (Pen borderPen = new Pen(this.Enabled ? BorderColor : SystemColors.ControlDarkDark))
				g.DrawRectangle(borderPen,bounds.Left, bounds.Top, bounds.Width - 2, bounds.Height);

				Rectangle buttonRect = new Rectangle(bounds.Right - 18, bounds.Top, 18, bounds.Height);
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
				bool header = _headerIndices != null && _headerIndices.Contains(e.Index);
				bool chosen = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
				Color currentBackColor = (chosen && !header) ? SelectedItemBackColor : DropDownBackColor;
				Color currentForeColor = (chosen && !header) ? SelectedItemForeColor : (header ? HeaderColor : TextColor);

				using (SolidBrush backBrush = new SolidBrush(currentBackColor))
				g.FillRectangle(backBrush,bounds.Left, bounds.Top, bounds.Width, bounds.Height);

				string itemText = GetItemDisplayText(this.Items[e.Index]);

				if (header)
				{
					bool ruleOnly = itemText.Trim().Trim('\u2500', '\u2014', '-').Length == 0;
					using (Pen p = new Pen(Color.FromArgb(90, BorderColor)))
					{
						int y = ruleOnly ? bounds.Top + bounds.Height / 2 : bounds.Top;
						g.DrawLine(p, bounds.Left + 4, y, bounds.Right - 4, y);
					}
					if (!ruleOnly)
					{
						Rectangle hb = new Rectangle(bounds.Left + 6, bounds.Top + 1, bounds.Width - 10, bounds.Height - 2);
						TextRenderer.DrawText(g, itemText.ToUpperInvariant(), HeaderFont, hb, currentForeColor,
							TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
					}
					return;
				}

				Image img = GetItemImage(e.Index);
				int textX = bounds.Left;
				if (img != null)
				{
					int iy = bounds.Top + (bounds.Height - 16) / 2;
					g.DrawImage(img, bounds.Left + 2, iy, 16, 16);
					textX = bounds.Left + 20;
				}
				Rectangle textBounds = new Rectangle(textX, bounds.Top, bounds.Width - (textX - bounds.Left), bounds.Height);
				TextRenderer.DrawText(g,itemText, this.Font, textBounds, currentForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
			}
		}

		public class DarkTabControl : TabControl
		{
			private Color _tabBgColor = Color.FromArgb(13, 17, 23);
			private Color _tabActiveBgColor = Color.FromArgb(22, 27, 34);
			private Color _tabInactiveBgColor = Color.FromArgb(13, 17, 23);
			private Color _tabActiveTextColor = Color.FromArgb(230, 237, 243);
			private Color _tabInactiveTextColor = Color.FromArgb(139, 148, 158);
			private Color _tabUnderlineColor = Color.FromArgb(31, 111, 235);
			private Color _tabInactiveUnderlineColor = Color.FromArgb(48, 54, 61);
			private bool _fillWidth = false;
			private bool _allowFormDrag = false;

			[DllImport("user32.dll")]
			private static extern bool ReleaseCapture();

			[DllImport("user32.dll")]
			private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

			public Color TabBgColor { get { return _tabBgColor; } set { _tabBgColor = value; Invalidate(); } }
			public Color TabActiveBgColor { get { return _tabActiveBgColor; } set { _tabActiveBgColor = value; Invalidate(); } }
			public Color TabInactiveBgColor { get { return _tabInactiveBgColor; } set { _tabInactiveBgColor = value; Invalidate(); } }
			public Color TabActiveTextColor { get { return _tabActiveTextColor; } set { _tabActiveTextColor = value; Invalidate(); } }
			public Color TabInactiveTextColor { get { return _tabInactiveTextColor; } set { _tabInactiveTextColor = value; Invalidate(); } }
			public Color TabUnderlineColor { get { return _tabUnderlineColor; } set { _tabUnderlineColor = value; Invalidate(); } }
			public Color TabInactiveUnderlineColor { get { return _tabInactiveUnderlineColor; } set { _tabInactiveUnderlineColor = value; Invalidate(); } }
			public bool FillWidth { get { return _fillWidth; } set { _fillWidth = value; this.Multiline = !value; this.SizeMode = value ? TabSizeMode.Fixed : TabSizeMode.Normal; Invalidate(); } }
			public bool AllowFormDrag { get { return _allowFormDrag; } set { _allowFormDrag = value; } }

			public DarkTabControl()
			{
				this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
				this.DrawMode = TabDrawMode.OwnerDrawFixed;
				this.Multiline = true;
				this.SizeMode = TabSizeMode.Normal;
			}

			public Rectangle GetDisplayTabRect(int index)
			{
				if (FillWidth && this.TabCount > 0)
				{
					int tabHeight = this.GetTabRect(0).Height;
					int tabWidth = this.ClientSize.Width / this.TabCount;
					int width = tabWidth;
					if (index == this.TabCount - 1)
					{
						width += this.ClientSize.Width % this.TabCount;
					}
					return new Rectangle(index * tabWidth, 0, width, tabHeight);
				}
				return this.GetTabRect(index);
			}

			protected override void OnSelectedIndexChanged(EventArgs e)
			{
				base.OnSelectedIndexChanged(e);
				this.Invalidate();
			}

			protected override void OnPaint(PaintEventArgs e)
			{
				Graphics g = e.Graphics;
				g.SmoothingMode = SmoothingMode.AntiAlias;

				using (SolidBrush b = new SolidBrush(TabBgColor))
				{
					g.FillRectangle(b, this.ClientRectangle);
				}

				for (int i = 0; i < this.TabCount; i++)
				{
					Rectangle tabRect = GetDisplayTabRect(i);

					bool isSelected = (this.SelectedIndex == i);
					Color bg = isSelected ? TabActiveBgColor : TabInactiveBgColor;

					using (SolidBrush bgb = new SolidBrush(bg))
					{
						g.FillRectangle(bgb, tabRect);
					}

					string text = this.TabPages[i].Text;
					Color textColor = isSelected ? TabActiveTextColor : TabInactiveTextColor;
					TextRenderer.DrawText(g, text, this.Font, tabRect, textColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

					if (isSelected)
					{
						using (Pen p = new Pen(TabUnderlineColor, 2))
						{
							g.DrawLine(p, tabRect.Left, tabRect.Bottom - 1, tabRect.Right, tabRect.Bottom - 1);
						}
					}
					else
					{
						using (Pen p = new Pen(TabInactiveUnderlineColor, 1))
						{
							g.DrawLine(p, tabRect.Left, tabRect.Bottom - 1, tabRect.Right, tabRect.Bottom - 1);
						}
					}
				}
			}

			protected override void WndProc(ref Message m)
			{
				if (m.Msg == 0x0201 && FillWidth && this.TabCount > 0)
				{
					int x = m.LParam.ToInt32() & 0xFFFF;
					int y = (m.LParam.ToInt32() >> 16) & 0xFFFF;
					Point pt = new Point(x, y);

					Rectangle firstRect = GetDisplayTabRect(0);
					if (y >= 0 && y < firstRect.Bottom)
					{
						for (int i = 0; i < this.TabCount; i++)
						{
							if (GetDisplayTabRect(i).Contains(pt))
							{
								if (this.SelectedIndex != i)
								{
									this.SelectedIndex = i;
								}
								if (_allowFormDrag)
								{
									Form owner = this.FindForm();
									if (owner != null)
									{
										ReleaseCapture();
										SendMessage(owner.Handle, 0x00A1, (IntPtr)0x0002, IntPtr.Zero);
									}
								}
								return;
							}
						}
						return;
					}
				}
				base.WndProc(ref m);
			}

			protected override void OnPaintBackground(PaintEventArgs pevent) {}
		}

		public class TabStripDragHelper : NativeWindow
		{
			private const int WM_LBUTTONDOWN = 0x0201;
			private const int WM_NCLBUTTONDOWN = 0x00A1;
			private const int HTCAPTION = 0x0002;
			private IntPtr _formHandle;
			private DarkTabControl _tabControl;

			[DllImport("user32.dll")]
			private static extern bool ReleaseCapture();

			[DllImport("user32.dll")]
			private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

			public TabStripDragHelper(IntPtr formHandle, DarkTabControl tabControl)
			{
				_formHandle = formHandle;
				_tabControl = tabControl;
				this.AssignHandle(tabControl.Handle);
			}

			protected override void WndProc(ref Message m)
			{
				if (m.Msg == WM_LBUTTONDOWN)
				{
					int x = m.LParam.ToInt32() & 0xFFFF;
					int y = (m.LParam.ToInt32() >> 16) & 0xFFFF;
					Point pt = new Point(x, y);

					bool onTabHeader = false;
					for (int i = 0; i < _tabControl.TabCount; i++)
					{
						Rectangle tabRect = _tabControl.GetDisplayTabRect(i);
						if (tabRect.Contains(pt))
						{
							onTabHeader = true;
							break;
						}
					}

					if (!onTabHeader)
					{
						int tabStripHeight = _tabControl.ItemSize.Height;
						if (tabStripHeight <= 0 && _tabControl.TabCount > 0)
						{
							tabStripHeight = _tabControl.GetTabRect(0).Bottom + 2;
						}
						if (y >= 0 && y < tabStripHeight)
						{
							ReleaseCapture();
							SendMessage(_formHandle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
							return;
						}
					}
				}
				base.WndProc(ref m);
			}
		}

		public class Toggle : CheckBox
		{
			private static readonly StringFormat CenterFormat = new StringFormat
			{
				Alignment = StringAlignment.Center,
				LineAlignment = StringAlignment.Center
			};

			private Color _onColor = Color.FromArgb(35, 175, 75);
			private Color _offColor = Color.FromArgb(210, 45, 45);
			public Color OnColor { get { return _onColor; } set { _onColor = value; } }
			public Color OffColor { get { return _offColor; } set { _offColor = value; } }

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
				Color onColor = OnColor;
				Color offColor = OffColor;

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
						var sf = CenterFormat;
						var boundsF = new RectangleF(0, 0, this.Width, this.Height);
						var shadowBounds = new RectangleF(boundsF.X + 1, boundsF.Y + 1, boundsF.Width, boundsF.Height);
						e.Graphics.DrawString(this.Text,this.Font, shadowBrush, shadowBounds, sf);
						e.Graphics.DrawString(this.Text,this.Font, textBrush, boundsF, sf);
					}
				}
			}
		}

		public class KeySafeButton : Button {
			protected override void OnKeyDown(KeyEventArgs e) { e.Handled = true; }
			protected override void OnKeyPress(KeyPressEventArgs e) { e.Handled = true; }
			protected override void OnKeyUp(KeyEventArgs e) { e.Handled = true; }
		}

		public class KeySafeCheckBox : CheckBox {
			protected override void OnKeyDown(KeyEventArgs e) { e.Handled = true; }
			protected override void OnKeyPress(KeyPressEventArgs e) { e.Handled = true; }
			protected override void OnKeyUp(KeyEventArgs e) { e.Handled = true; }
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
		
			private Color _fillColor = Color.FromArgb(0, 120, 215);
			public Color FillColor { get { return _fillColor; } set { _fillColor = value; } }

			private static readonly Font LabelFont = new Font("Segoe UI", 8);

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
				
					using (SolidBrush b = new SolidBrush(FillColor))
					{
						g.FillRectangle(b,clip);
					}
				}

				{
					Font f = LabelFont;
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

		private Color _thumbNormalColor = Color.FromArgb(33, 38, 45);
		private Color _thumbHoverColor  = Color.FromArgb(48, 54, 61);
		private Color _thumbActiveColor = Color.FromArgb(88, 166, 255);
		public Color ThumbNormalColor { get { return _thumbNormalColor; } set { _thumbNormalColor = value; } }
		public Color ThumbHoverColor  { get { return _thumbHoverColor; }  set { _thumbHoverColor = value; } }
		public Color ThumbActiveColor { get { return _thumbActiveColor; } set { _thumbActiveColor = value; } }

		public DarkScrollBar()
		{
			this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
			this.Width = 12;
			this.BackColor = Color.FromArgb(13, 17, 23);
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
			
			Color thumbColor = _isDragging ? ThumbActiveColor : (_isHovered ? ThumbHoverColor : ThumbNormalColor);
			using (System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath())
			{
				int r = 4; Rectangle rect = new Rectangle(2, thumbY, Width - 4, thumbHeight);
				path.AddArc(rect.X, rect.Y, r, r, 180, 90); path.AddArc(rect.Right - r, rect.Y, r, r, 270, 90);
				path.AddArc(rect.Right - r, rect.Bottom - r, r, r, 0, 90); path.AddArc(rect.X, rect.Bottom - r, r, r, 90, 90);
				path.CloseFigure();
				using (SolidBrush b = new SolidBrush(thumbColor)) { e.Graphics.FillPath(b, path); }
			}
		}
	}

	public class DarkDataGridView : DataGridView
	{
		private DarkScrollBar _vScroll;
		private Color _checkBoxCheckedColor = Color.FromArgb(31, 111, 235);
		private Color _checkBoxUncheckedColor = Color.FromArgb(22, 27, 34);
		private Color _checkBoxBorderColor = Color.FromArgb(48, 54, 61);
		private Color _checkBoxCheckColor = Color.White;
		public Color CheckBoxCheckedColor { get { return _checkBoxCheckedColor; } set { _checkBoxCheckedColor = value; Invalidate(); } }
		public Color CheckBoxUncheckedColor { get { return _checkBoxUncheckedColor; } set { _checkBoxUncheckedColor = value; Invalidate(); } }
		public Color CheckBoxBorderColor { get { return _checkBoxBorderColor; } set { _checkBoxBorderColor = value; Invalidate(); } }
		public Color CheckBoxCheckColor { get { return _checkBoxCheckColor; } set { _checkBoxCheckColor = value; Invalidate(); } }
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
			this.DataBindingComplete += (s,e) => UpdateScroll();
			this.MouseWheel += (s,e) => { if (_vScroll.Visible) _vScroll.Value -= (e.Delta / 120) * 3; };
		}

		public void RefreshScroll()
		{
			UpdateScroll();
		}

		protected override void OnCellPainting(DataGridViewCellPaintingEventArgs e)
		{
			if (e.RowIndex >= 0 &&
				e.ColumnIndex >= 0 &&
				this.Columns[e.ColumnIndex] is DataGridViewCheckBoxColumn)
			{
				e.Paint(
					e.CellBounds,
					DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground
				);

				bool isChecked = false;

				if (e.Value != null && e.Value != DBNull.Value)
				{
					if (e.Value is bool)
					{
						isChecked = (bool)e.Value;
					}
					else
					{
						bool parsed;

						if (bool.TryParse(e.Value.ToString(), out parsed))
						{
							isChecked = parsed;
						}
					}
				}

				int size = 14;

				int x = e.CellBounds.X +
						(e.CellBounds.Width - size) / 2;

				int y = e.CellBounds.Y +
						(e.CellBounds.Height - size) / 2;

				Rectangle boxRect = new Rectangle(
					x,
					y,
					size,
					size
				);

				Color boxBg = isChecked ? CheckBoxCheckedColor : CheckBoxUncheckedColor;

				Color border = CheckBoxBorderColor;

				using (SolidBrush b = new SolidBrush(boxBg))
				{
					e.Graphics.FillRectangle(b, boxRect);
				}

				using (Pen p = new Pen(border))
				{
					e.Graphics.DrawRectangle(p, boxRect);
				}

				if (isChecked)
				{
					using (Pen checkPen = new Pen(CheckBoxCheckColor, 2))
					{
						e.Graphics.DrawLines(
							checkPen,
							new Point[]
							{
								new Point(x + 3, y + 7),
								new Point(x + 6, y + 10),
								new Point(x + 11, y + 4)
							}
						);
					}
				}

				e.Handled = true;
				return;
			}

			base.OnCellPainting(e);
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
		public override Color MenuItemSelected { get { return Color.FromArgb(31, 111, 235); } }
		public override Color MenuItemBorder { get { return Color.FromArgb(31, 111, 235); } }
		public override Color MenuBorder { get { return Color.FromArgb(48, 54, 61); } }
		public override Color MenuItemSelectedGradientBegin { get { return Color.FromArgb(31, 111, 235); } }
		public override Color MenuItemSelectedGradientEnd { get { return Color.FromArgb(31, 111, 235); } }
		public override Color MenuItemPressedGradientBegin { get { return Color.FromArgb(31, 111, 235); } }
		public override Color MenuItemPressedGradientEnd { get { return Color.FromArgb(31, 111, 235); } }
		public override Color ImageMarginGradientBegin { get { return Color.FromArgb(22, 27, 34); } }
		public override Color ImageMarginGradientMiddle { get { return Color.FromArgb(22, 27, 34); } }
		public override Color ImageMarginGradientEnd { get { return Color.FromArgb(22, 27, 34); } }
		public override Color ToolStripDropDownBackground { get { return Color.FromArgb(22, 27, 34); } }
		public override Color SeparatorDark { get { return Color.FromArgb(48, 54, 61); } }
		public override Color SeparatorLight { get { return Color.FromArgb(48, 54, 61); } }
	}

	public class DarkRenderer : ToolStripProfessionalRenderer
	{
		public DarkRenderer() : base(new DarkColorTable()) { this.RoundedEdges = false; }

		protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
		{
			e.TextColor = Color.FromArgb(230, 237, 243);
			base.OnRenderItemText(e);
		}

		protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
		{
			e.ArrowColor = Color.FromArgb(230, 237, 243);
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

	public sealed class MarketCatalogEntry
	{
		public string Name { get; set; }
		public string Connection { get; set; }
		public string Desc { get; set; }
	}

	public static class MarketCatalogIndex
	{
		public static bool Read(string path, out Dictionary<string, MarketCatalogEntry> byName, out string[] names)
		{
			byName = null;
			names = null;
			if (string.IsNullOrEmpty(path)) return false;
			try
			{
				string[] lines = File.ReadAllLines(path, Encoding.UTF8);
				if (lines.Length < 2) return false;

				var map = new Dictionary<string, MarketCatalogEntry>(StringComparer.OrdinalIgnoreCase);
				var list = new List<string>(lines.Length - 1);

				for (int i = 1; i < lines.Length; i++)
				{
					string ln = lines[i];
					if (ln.Length == 0) continue;

					int tab = ln.IndexOf('\t');
					if (tab < 1) continue;

					string n = ln.Substring(0, tab).Trim();
					if (n.Length == 0) continue;

					string rest = ln.Substring(tab + 1);
					string conn = "";
					string desc = "";
					int tab2 = rest.IndexOf('\t');
					if (tab2 >= 0)
					{
						conn = rest.Substring(0, tab2);
						desc = rest.Substring(tab2 + 1);
					}
					else
					{
						conn = rest;
					}

					if (map.ContainsKey(n)) continue;
					map.Add(n, new MarketCatalogEntry { Name = n, Connection = conn, Desc = desc });
					list.Add(n);
				}

				if (map.Count == 0) return false;
				byName = map;
				names = list.ToArray();
				return true;
			}
			catch
			{
				byName = null;
				names = null;
				return false;
			}
		}

		public static bool ReadHeader(string path, out int version, out string updated)
		{
			version = 0;
			updated = null;
			if (string.IsNullOrEmpty(path)) return false;
			try
			{
				string[] lines = File.ReadAllLines(path, Encoding.UTF8);
				if (lines.Length < 1) return false;
				string[] hdr = lines[0].Split('\t');
				if (hdr.Length < 3) return false;
				int v;
				if (!int.TryParse(hdr[1], out v)) return false;
				version = v;
				updated = hdr[2];
				return true;
			}
			catch { return false; }
		}
	}

    public static class LogClaimLedger
    {
        public sealed class Claim
        {
            public string Marker;
            public int Remaining;
            public bool Cancelled;
            public readonly ManualResetEventSlim Done = new ManualResetEventSlim(false);
        }

        private sealed class LogState
        {
            public int ProcessedLines;
            public readonly Dictionary<string, Queue<Claim>> Pending = new Dictionary<string, Queue<Claim>>(StringComparer.Ordinal);
        }

        private static readonly object Sync = new object();
        private static readonly Dictionary<string, LogState> Logs = new Dictionary<string, LogState>(StringComparer.OrdinalIgnoreCase);

        private static LogState GetState(string path)
        {
            LogState state;
            if (!Logs.TryGetValue(path, out state))
            {
                state = new LogState();
                var existing = ReadAllText(path);
                state.ProcessedLines = existing == null ? 0 : CountCompleteLines(existing);
                Logs[path] = state;
            }
            return state;
        }

        private static string ReadAllText(string path)
        {
            try
            {
                if (!File.Exists(path)) return string.Empty;
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                {
                    return reader.ReadToEnd();
                }
            }
            catch
            {
                return null;
            }
        }

        private static int CountCompleteLines(string text)
        {
            int count = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\n') count++;
            }
            return count;
        }

        public static Claim Register(string path, string marker, int required)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(marker)) return null;
            lock (Sync)
            {
                var state = GetState(path);
                Queue<Claim> queue;
                if (!state.Pending.TryGetValue(marker, out queue))
                {
                    queue = new Queue<Claim>();
                    state.Pending[marker] = queue;
                }
                var claim = new Claim { Marker = marker, Remaining = Math.Max(1, required) };
                queue.Enqueue(claim);
                return claim;
            }
        }

        public static bool IsDone(Claim claim)
        {
            return claim != null && claim.Done.IsSet;
        }

        public static void Cancel(Claim claim)
        {
            if (claim == null) return;
            lock (Sync)
            {
                claim.Cancelled = true;
                claim.Done.Set();
                foreach (var entry in Logs)
                {
                    Queue<Claim> queue;
                    if (!entry.Value.Pending.TryGetValue(claim.Marker, out queue) || queue.Count == 0) continue;
                    if (!queue.Contains(claim)) continue;
                    var rebuilt = new Queue<Claim>(queue.Count);
                    while (queue.Count > 0)
                    {
                        var item = queue.Dequeue();
                        if (!ReferenceEquals(item, claim)) rebuilt.Enqueue(item);
                    }
                    entry.Value.Pending[claim.Marker] = rebuilt;
                }
            }
        }

        public static void Poll(string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            lock (Sync)
            {
                var state = GetState(path);
                var text = ReadAllText(path);
                if (text == null) return;
                int complete = CountCompleteLines(text);
                if (complete < state.ProcessedLines) state.ProcessedLines = 0;
                if (complete <= state.ProcessedLines) return;

                var lines = text.Split('\n');
                for (int i = state.ProcessedLines; i < complete && i < lines.Length; i++)
                {
                    var line = lines[i];
                    foreach (var marker in new List<string>(state.Pending.Keys))
                    {
                        if (line.IndexOf(marker, StringComparison.Ordinal) >= 0)
                        {
                            Dispatch(state, marker);
                        }
                    }
                }
                state.ProcessedLines = complete;
            }
        }

        private static void Dispatch(LogState state, string marker)
        {
            Queue<Claim> queue;
            if (!state.Pending.TryGetValue(marker, out queue)) return;
            while (queue.Count > 0)
            {
                var claim = queue.Peek();
                if (claim.Cancelled || claim.Done.IsSet)
                {
                    queue.Dequeue();
                    continue;
                }
                claim.Remaining--;
                if (claim.Remaining <= 0)
                {
                    queue.Dequeue();
                    claim.Done.Set();
                }
                return;
            }
        }
    }
}
"@ 
#endregion

#region Atlas and Market Core (moved from atlas.psm1 and market.psm1)
$script:AtlasCoreSource = @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace Atlas
{
    #region JSON helpers

    internal static class J
    {
        internal static Dictionary<string, object> M(object o)
        {
            Dictionary<string, object> d = o as Dictionary<string, object>;
            return d;
        }

        internal static object[] A(object o)
        {
            return o as object[];
        }

        internal static string S(object o)
        {
            if (o == null) return null;
            string s = o as string;
            return s;
        }

        internal static object Val(Dictionary<string, object> d, string key)
        {
            if (d == null) return null;
            object v;
            return d.TryGetValue(key, out v) ? v : null;
        }

        internal static double D(object o, double def)
        {
            if (o == null) return def;
            if (o is double) return (double)o;
            if (o is int) return (int)o;
            if (o is long) return (long)o;
            if (o is decimal) return (double)(decimal)o;
            if (o is string)
            {
                double r;
                if (double.TryParse((string)o, NumberStyles.Float, CultureInfo.InvariantCulture, out r)) return r;
            }
            return def;
        }

        internal static Dictionary<string, object> Sub(Dictionary<string, object> root, string key)
        {
            if (root == null) return null;
            object v;
            if (!root.TryGetValue(key, out v)) return null;
            return v as Dictionary<string, object>;
        }

        internal static string[] StrArray(Dictionary<string, object> tbl, string map)
        {
            if (tbl == null) return new string[0];
            object v;
            if (!tbl.TryGetValue(map, out v)) return new string[0];
            object[] a = A(v);
            if (a == null) return new string[0];
            string[] r = new string[a.Length];
            for (int i = 0; i < a.Length; i++) r[i] = S(a[i]);
            return r;
        }
    }

    #endregion

    #region Model

    public sealed class WorldMeta
    {
        public string Key, Title, Img;
        public int W, H;
    }

    public sealed class Species
    {
        public string Name, Map, Kind, Zone, ColorHex, Icon;
        public Color Color;
        public double N;
        public int GlobalIndex;
    }

    public struct Pt
    {
        public double X, Y;
        public int Ref;
        public string Zone;
        public Pt(double x, double y, int r) { X = x; Y = y; Ref = r; Zone = null; }
        public Pt(double x, double y, int r, string z) { X = x; Y = y; Ref = r; Zone = z; }
    }

    public sealed class NpcStack
    {
        public string Name, Icon, Sym;
        public int Id, Qty;
    }

    public sealed class NpcTradeItem
    {
        public string Name, Icon, Sym;
        public int Id, Price, Tab;
    }

    public sealed class NpcTrade
    {
        public string[] Tabs = new string[0];
        public NpcStack Currency;
        public List<NpcTradeItem> Items = new List<NpcTradeItem>();
    }

    public sealed class NpcOffer
    {
        public List<NpcStack> Cost = new List<NpcStack>();
        public List<NpcStack> Get = new List<NpcStack>();
    }

    public sealed class NpcExchange
    {
        public string Menu, Label;
        public List<NpcOffer> Sets = new List<NpcOffer>();
    }

    public sealed class NpcShop
    {
        public string Script, Kind, Tag, Text;
        public NpcTrade Trade;
        public List<NpcExchange> Exchange = new List<NpcExchange>();
    }

    /* The trader card from the web map, painted here instead of assembled out
     * of labels: a shop window is a panel, trade stock is a grid of pictures
     * and prices, and an exchange set is one give -> get row.  A single NPC
     * can keep eighty-four offers, so the view scrolls itself and the detail
     * card stays on screen. */
    public sealed class ShopView : Control
    {
        private NpcShop shop;
        private string assetRoot;
        private string[] fallbackRoots;
        private int scroll;
        private int contentH;
        private bool scrollDrag;
        private int dragY;
        private int dragScroll;

        private const int HeadH = 25;
        private const int TabH = 16;
        private const int ItemH = 30;
        private const int OfferH = 29;
        private const int Gap = 3;
        private const int PanelGap = 8;
        private const int ScrollW = 6;
        private const int IconSize = 22;

        private Color CShop = Color.FromArgb(14, 20, 28);
        private Color CItem = Color.FromArgb(24, 33, 45);
        private Color CPanel2 = Color.FromArgb(24, 33, 46);
        private Color CPanel3 = Color.FromArgb(31, 42, 57);
        private Color CLine = Color.FromArgb(38, 51, 69);
        private Color CLine2 = Color.FromArgb(49, 65, 90);
        private Color CFg = Color.FromArgb(231, 238, 247);
        private Color CMut = Color.FromArgb(132, 150, 172);
        private Color CDim = Color.FromArgb(93, 110, 131);
        private Color CAcc = Color.FromArgb(90, 176, 255);
        private Color CGold = Color.FromArgb(255, 200, 87);
        private Color CScroll = Color.FromArgb(112, 130, 153);
        private Color CScrollTrack = Color.FromArgb(20, 27, 37);

        private Font fHead;
        private Font fBody;
        private Font fSmall;
        private Font fMono;
        private Font fItalic;

        private static Dictionary<string, Bitmap> iconCache = new Dictionary<string, Bitmap>();
        private static readonly object iconGate = new object();

        public NpcShop Shop
        {
            get { return shop; }
            set { shop = value; scroll = 0; Measure(); Invalidate(); }
        }

        public string AssetRoot
        {
            get { return assetRoot; }
            set { assetRoot = value; Invalidate(); }
        }

        /* Extra folders searched when a shelf picture is not under the
         * atlas root. The atlas now carries every picture itself, so this
         * is only a safety net for custom asset layouts. */
        public string[] FallbackRoots
        {
            get { return fallbackRoots; }
            set { fallbackRoots = value; Invalidate(); }
        }

        public int ContentHeight { get { return contentH; } }

        public void ScrollBy(int delta)
        {
            ScrollTo(scroll + delta);
        }

        public ShopView()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.Selectable, true);
            DoubleBuffered = true;
            TabStop = true;
            BackColor = Color.FromArgb(17, 24, 35);
            fHead = new Font("Segoe UI", 7.5f, FontStyle.Bold);
            fBody = new Font("Segoe UI", 7.5f, FontStyle.Regular);
            fSmall = new Font("Segoe UI", 6.75f, FontStyle.Bold);
            fMono = new Font("Consolas", 7.5f, FontStyle.Bold);
            fItalic = new Font("Segoe UI", 7.5f, FontStyle.Italic);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (fHead != null) fHead.Dispose();
                if (fBody != null) fBody.Dispose();
                if (fSmall != null) fSmall.Dispose();
                if (fMono != null) fMono.Dispose();
                if (fItalic != null) fItalic.Dispose();
            }
            base.Dispose(disposing);
        }

        private static Color Blend(Color under, Color over, double t)
        {
            return Color.FromArgb(
                (int)Math.Round(under.R + (over.R - under.R) * t),
                (int)Math.Round(under.G + (over.G - under.G) * t),
                (int)Math.Round(under.B + (over.B - under.B) * t));
        }

        public void SetPalette(Color card, Color panel2, Color panel3, Color line, Color line2,
                               Color fg, Color mut, Color dim, Color acc, Color gold)
        {
            BackColor = card;
            CShop = Blend(card, Color.FromArgb(10, 14, 19), 0.45);
            CItem = Blend(CShop, panel3, 0.60);
            CPanel2 = panel2;
            CPanel3 = panel3;
            CLine = line;
            CLine2 = line2;
            CFg = fg;
            CMut = mut;
            CDim = dim;
            CAcc = acc;
            CGold = gold;
            Invalidate();
        }

        private int PanelWidth
        {
            get
            {
                int w = Width - ScrollW;
                if (w < 80) w = Width;
                return w;
            }
        }

        private static GraphicsPath Round(Rectangle r, int radius)
        {
            GraphicsPath p = new GraphicsPath();
            if (r.Width <= 0 || r.Height <= 0) return p;
            int d = Math.Min(radius * 2, Math.Min(r.Width, r.Height));
            if (d <= 1) { p.AddRectangle(r); return p; }
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        private static GraphicsPath TopRound(Rectangle r, int radius)
        {
            GraphicsPath p = new GraphicsPath();
            if (r.Width <= 0 || r.Height <= 0) return p;
            int d = Math.Min(radius * 2, Math.Min(r.Width, r.Height));
            if (d <= 1) { p.AddRectangle(r); return p; }
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddLine(r.Right, r.Y + d / 2, r.Right, r.Bottom);
            p.AddLine(r.Right, r.Bottom, r.X, r.Bottom);
            p.AddLine(r.X, r.Bottom, r.X, r.Y + d / 2);
            p.CloseFigure();
            return p;
        }

        private int Columns(int innerW)
        {
            int cols = (innerW + Gap) / (132 + Gap);
            return Math.Max(1, cols);
        }

        private int GridHeight(int count, int cols)
        {
            if (count <= 0) return 0;
            int rows = (count + cols - 1) / cols;
            return 4 + rows * ItemH + (rows - 1) * Gap + 6;
        }

        private string[] TabsOf(NpcTrade tr)
        {
            if (tr.Tabs != null && tr.Tabs.Length > 0) return tr.Tabs;
            return new string[] { "Shop" };
        }

        private int CountTab(NpcTrade tr, int tab)
        {
            int n = 0;
            for (int i = 0; i < tr.Items.Count; i++)
            {
                int t = tr.Items[i].Tab;
                if (t < 0) t = 0;
                if (t == tab) n++;
            }
            return n;
        }

        private int TradeHeight(int w)
        {
            int innerW = w - 12;
            int cols = Columns(innerW);
            NpcTrade tr = shop.Trade;
            string[] tabs = TabsOf(tr);
            int h = HeadH;
            if (tabs.Length > 1)
            {
                for (int i = 0; i < tabs.Length; i++)
                {
                    int n = CountTab(tr, i);
                    if (n == 0) continue;
                    h += TabH + GridHeight(n, cols);
                }
            }
            else
            {
                h += GridHeight(tr.Items.Count, cols);
            }
            return h;
        }

        private int ExchangeHeight(NpcExchange ex)
        {
            int h = HeadH;
            for (int i = 0; i < ex.Sets.Count; i++)
            {
                NpcOffer set = ex.Sets[i];
                int n = Math.Max(set.Cost.Count, set.Get.Count);
                if (n < 1) n = 1;
                h += 10 + n * OfferH + (n - 1) * Gap;
            }
            return h;
        }

        private void Measure()
        {
            contentH = 0;
            if (shop == null) return;
            int w = Width - ScrollW;
            if (w < 120) w = 400;
            int y = 0;
            if (shop.Trade != null && shop.Trade.Items.Count > 0)
            {
                y += TradeHeight(w) + PanelGap;
            }
            if (shop.Exchange != null)
            {
                for (int i = 0; i < shop.Exchange.Count; i++)
                {
                    if (shop.Exchange[i].Sets.Count == 0) continue;
                    y += ExchangeHeight(shop.Exchange[i]) + PanelGap;
                }
            }
            if (y > 0) y -= PanelGap;
            contentH = y;
            int max = Math.Max(0, contentH - Height);
            if (scroll > max) scroll = max;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Measure();
            Invalidate();
        }

        private void Frame(Graphics g, Rectangle r)
        {
            using (GraphicsPath p = Round(r, 8))
            using (SolidBrush b = new SolidBrush(CShop))
            {
                g.FillPath(b, p);
            }
            Rectangle head = new Rectangle(r.X, r.Y, r.Width, HeadH);
            using (GraphicsPath hp = TopRound(head, 8))
            using (SolidBrush b = new SolidBrush(CPanel2))
            {
                g.FillPath(b, hp);
            }
            using (Pen pen = new Pen(CLine))
            {
                g.DrawLine(pen, r.X + 1, r.Y + HeadH - 1, r.Right - 1, r.Y + HeadH - 1);
            }
            using (GraphicsPath p = Round(r, 8))
            using (Pen pen = new Pen(CLine))
            {
                g.DrawPath(pen, p);
            }
        }

        private void DrawHeader(Graphics g, Rectangle panel, string left, string right, string icon)
        {
            int maxW = Math.Max(0, panel.Width / 2 - 4);
            int tw = 0;
            int contentW = 0;
            int startX = panel.Right - 8;
            if (!string.IsNullOrEmpty(right))
            {
                tw = TextRenderer.MeasureText(right, fBody).Width + 2;
                if (tw > maxW) tw = maxW;
                contentW = tw + (string.IsNullOrEmpty(icon) ? 0 : 22);
                startX = panel.Right - 8 - contentW;
            }
            int lw = panel.Width - 16 - (contentW > 0 ? contentW + 8 : 0);
            Rectangle lr = new Rectangle(panel.X + 8, panel.Y, Math.Max(0, lw), HeadH);
            TextRenderer.DrawText(g, left ?? "", fHead, lr, CAcc,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);

            if (string.IsNullOrEmpty(right)) return;
            if (!string.IsNullOrEmpty(icon))
            {
                DrawIcon(g, icon, startX, panel.Y + (HeadH - 18) / 2, 18);
            }
            Rectangle tr = new Rectangle(startX + (string.IsNullOrEmpty(icon) ? 0 : 22),
                                         panel.Y, Math.Max(0, tw), HeadH);
            TextRenderer.DrawText(g, right, fBody, tr, CMut,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
        }

        private void DrawTab(Graphics g, Rectangle panel, string tab, int y)
        {
            Rectangle r = new Rectangle(panel.X + 8, y, Math.Max(0, panel.Width - 16), TabH);
            TextRenderer.DrawText(g, (tab ?? "").ToUpperInvariant(), fSmall, r, CDim,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
        }

        private void DrawItemCell(Graphics g, Rectangle r, string name, string icon, int value)
        {
            using (GraphicsPath p = Round(r, 5))
            using (SolidBrush b = new SolidBrush(CItem))
            {
                g.FillPath(b, p);
            }
            int ix = r.X + 4;
            int iy = r.Y + (r.Height - IconSize) / 2;
            DrawIcon(g, icon, ix, iy, IconSize);

            string q = value > 1 ? value.ToString() : null;
            int qw = 0;
            if (q != null) qw = TextRenderer.MeasureText(q, fMono).Width + 4;
            int nx = ix + IconSize + 5;
            int nw = Math.Max(0, r.Right - 4 - qw - nx);
            TextRenderer.DrawText(g, name ?? "", fBody, new Rectangle(nx, r.Y, nw, r.Height), CFg,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
            if (q != null)
            {
                Rectangle qr = new Rectangle(r.Right - 4 - qw, r.Y, qw, r.Height);
                TextRenderer.DrawText(g, q, fMono, qr, CGold,
                    TextFormatFlags.Right | TextFormatFlags.VerticalCenter |
                    TextFormatFlags.NoPrefix);
            }
        }

        private void DrawTrade(Graphics g, int y, int w, int top, int bottom)
        {
            int h = TradeHeight(w);
            Rectangle panel = new Rectangle(0, y, w, h);
            Frame(g, panel);
            NpcTrade tr = shop.Trade;
            DrawHeader(g, panel, "TRADE", tr.Currency == null ? null : tr.Currency.Name,
                       tr.Currency == null ? null : tr.Currency.Icon);

            int innerW = panel.Width - 12;
            int cols = Columns(innerW);
            int cellW = (innerW - (cols - 1) * Gap) / cols;
            int x0 = panel.X + 6;
            int cy = panel.Y + HeadH;
            string[] tabs = TabsOf(tr);
            if (tabs.Length > 1)
            {
                for (int ti = 0; ti < tabs.Length; ti++)
                {
                    int n = CountTab(tr, ti);
                    if (n == 0) continue;
                    DrawTab(g, panel, tabs[ti], cy);
                    cy += TabH;
                    int gy = cy + 4;
                    int drawn = 0;
                    for (int i = 0; i < tr.Items.Count; i++)
                    {
                        NpcTradeItem it = tr.Items[i];
                        int t = it.Tab < 0 ? 0 : it.Tab;
                        if (t != ti) continue;
                        Rectangle cell = new Rectangle(x0 + (drawn % cols) * (cellW + Gap),
                                                       gy + (drawn / cols) * (ItemH + Gap), cellW, ItemH);
                        if (cell.Bottom >= top && cell.Y <= bottom)
                            DrawItemCell(g, cell, it.Name, it.Icon, it.Price);
                        drawn++;
                    }
                    cy += GridHeight(n, cols);
                }
            }
            else
            {
                int gy = cy + 4;
                for (int i = 0; i < tr.Items.Count; i++)
                {
                    NpcTradeItem it = tr.Items[i];
                    Rectangle cell = new Rectangle(x0 + (i % cols) * (cellW + Gap),
                                                   gy + (i / cols) * (ItemH + Gap), cellW, ItemH);
                    if (cell.Bottom >= top && cell.Y <= bottom)
                        DrawItemCell(g, cell, it.Name, it.Icon, it.Price);
                }
            }
        }

        private void DrawStacks(Graphics g, List<NpcStack> stacks, int x, int y, int w, int blockH,
                                int top, int bottom)
        {
            if (stacks == null || stacks.Count == 0)
            {
                if (y + blockH >= top && y <= bottom)
                {
                    TextRenderer.DrawText(g, "nothing", fItalic, new Rectangle(x, y, w, blockH), CDim,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                        TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
                }
                return;
            }
            for (int i = 0; i < stacks.Count; i++)
            {
                Rectangle cell = new Rectangle(x, y + i * (OfferH + Gap), w, OfferH);
                if (cell.Bottom < top || cell.Y > bottom) continue;
                NpcStack s = stacks[i];
                DrawItemCell(g, cell, s.Name, s.Icon, s.Qty);
            }
        }

        private void DrawExchange(Graphics g, int y, int w, NpcExchange ex, int top, int bottom)
        {
            int h = ExchangeHeight(ex);
            Rectangle panel = new Rectangle(0, y, w, h);
            Frame(g, panel);
            int n = ex.Sets.Count;
            DrawHeader(g, panel, (ex.Label == null ? "Exchange" : ex.Label).ToUpperInvariant(),
                       n.ToString() + (n == 1 ? " offer" : " offers"), null);

            int cy = panel.Y + HeadH;
            for (int i = 0; i < ex.Sets.Count; i++)
            {
                NpcOffer set = ex.Sets[i];
                int rows = Math.Max(set.Cost.Count, set.Get.Count);
                if (rows < 1) rows = 1;
                int blockH = rows * OfferH + (rows - 1) * Gap;
                int rowH = blockH + 10;
                if (cy + rowH >= top && cy <= bottom)
                {
                    if (i > 0)
                    {
                        using (Pen pen = new Pen(CLine))
                        {
                            g.DrawLine(pen, panel.X + 1, cy, panel.Right - 1, cy);
                        }
                    }
                    int contentW = panel.Width - 12;
                    int colW = Math.Max(60, (contentW - 24) / 2);
                    int lx = panel.X + 6;
                    int rx = lx + colW + 16 + 8;
                    int ty = cy + 5;
                    DrawStacks(g, set.Cost, lx, ty, colW, blockH, top, bottom);
                    DrawStacks(g, set.Get, rx, ty, colW, blockH, top, bottom);
                    Rectangle ar = new Rectangle(lx + colW, cy, 16, rowH);
                    TextRenderer.DrawText(g, "\u2192", fBody, ar, CDim,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                        TextFormatFlags.NoPrefix);
                }
                cy += rowH;
            }
        }

        private void DrawIcon(Graphics g, string rel, int x, int y, int size)
        {
            Bitmap bmp = LoadIcon(rel);
            if (bmp != null)
            {
                g.DrawImage(bmp, new Rectangle(x, y, size, size));
                return;
            }
            using (GraphicsPath p = Round(new Rectangle(x, y, size, size), 3))
            using (SolidBrush b = new SolidBrush(CPanel3))
            {
                g.FillPath(b, p);
            }
        }

        private Bitmap LoadIcon(string rel)
        {
            if (string.IsNullOrEmpty(rel) || string.IsNullOrEmpty(assetRoot)) return null;
            string key = rel;
            lock (iconGate)
            {
                Bitmap hit;
                if (iconCache.TryGetValue(key, out hit)) return hit;
            }
            Bitmap made = null;
            try
            {
                string path = Path.IsPathRooted(rel) ? rel :
                    Path.Combine(assetRoot, rel.Replace('/', Path.DirectorySeparatorChar));
                if (!File.Exists(path) && fallbackRoots != null)
                {
                    string leaf = Path.GetFileName(rel);
                    for (int i = 0; i < fallbackRoots.Length; i++)
                    {
                        string root = fallbackRoots[i];
                        if (string.IsNullOrEmpty(root)) continue;
                        string alt = Path.Combine(root, leaf);
                        if (File.Exists(alt)) { path = alt; break; }
                    }
                }
                if (File.Exists(path))
                {
                    using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (Image src = Image.FromStream(fs))
                    {
                        made = new Bitmap(22, 22, PixelFormat.Format32bppArgb);
                        using (Graphics g = Graphics.FromImage(made))
                        {
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            g.CompositingQuality = CompositingQuality.HighQuality;
                            g.DrawImage(src, new Rectangle(0, 0, 22, 22));
                        }
                    }
                }
            }
            catch { made = null; }
            lock (iconGate)
            {
                iconCache[key] = made;
            }
            return made;
        }

        private Rectangle ThumbRect()
        {
            int track = Math.Max(1, Height - 4);
            int thumb = Math.Max(24, (int)((long)Height * track / Math.Max(1, contentH)));
            if (thumb > track) thumb = track;
            int max = Math.Max(1, contentH - Height);
            int travel = Math.Max(0, track - thumb);
            int top = 2 + (int)((long)travel * scroll / max);
            return new Rectangle(Width - 5, top, 4, thumb);
        }

        private void ScrollTo(int value)
        {
            int max = Math.Max(0, contentH - Height);
            int next = Math.Max(0, Math.Min(value, max));
            if (next == scroll) return;
            scroll = next;
            Invalidate();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            ScrollTo(scroll - (e.Delta / 120) * 60);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            if (e.Button != MouseButtons.Left || contentH <= Height) return;
            Rectangle thumb = ThumbRect();
            if (e.X >= Width - 10)
            {
                if (thumb.Contains(e.Location))
                {
                    scrollDrag = true;
                    dragY = e.Y;
                    dragScroll = scroll;
                    Capture = true;
                }
                else
                {
                    ScrollTo(scroll + (e.Y < thumb.Y ? -Height : Height));
                }
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!scrollDrag) return;
            int max = contentH - Height;
            int track = Math.Max(1, Height - 4);
            int thumb = ThumbRect().Height;
            int travel = Math.Max(1, track - thumb);
            ScrollTo(dragScroll + (int)((long)(e.Y - dragY) * max / travel));
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (scrollDrag)
            {
                scrollDrag = false;
                Capture = false;
            }
            base.OnMouseUp(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Measure();
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            e.Graphics.Clear(BackColor);
            e.Graphics.SetClip(ClientRectangle);
            if (shop == null) return;

            int y = -scroll;
            int w = PanelWidth;
            int top = 0;
            int bottom = Height;
            if (shop.Trade != null && shop.Trade.Items.Count > 0)
            {
                DrawTrade(e.Graphics, y, w, top, bottom);
                y += TradeHeight(w) + PanelGap;
            }
            if (shop.Exchange != null)
            {
                for (int i = 0; i < shop.Exchange.Count; i++)
                {
                    NpcExchange ex = shop.Exchange[i];
                    if (ex.Sets.Count == 0) continue;
                    if (y + ExchangeHeight(ex) >= top && y <= bottom)
                        DrawExchange(e.Graphics, y, w, ex, top, bottom);
                    y += ExchangeHeight(ex) + PanelGap;
                }
            }
            if (contentH > Height)
            {
                using (SolidBrush track = new SolidBrush(CScrollTrack))
                {
                    e.Graphics.FillRectangle(track, new Rectangle(Width - 5, 2, 4, Math.Max(0, Height - 4)));
                }
                using (GraphicsPath p = Round(ThumbRect(), 2))
                using (SolidBrush b = new SolidBrush(CScroll))
                {
                    e.Graphics.FillPath(b, p);
                }
            }
        }
    }

    public sealed class Region
    {
        public string Title;
        public double X, Y;
        public bool Big;
    }

    public sealed class World
    {
        public WorldMeta Meta;

        public List<Species> Species = new List<Species>();
        public List<Pt> Spots = new List<Pt>();

        public List<Pt> Mobs = new List<Pt>();
        public string[] CreatureNames = new string[0];
        public HashSet<int> Bosses = new HashSet<int>();
        public string[] CreatureIcons = new string[0];
        public string[] CreatureSquare = new string[0];
        public string[] CreatureZone = new string[0];
        public string[] CreatureAlt = new string[0];
        public string[] CreatureLevel = new string[0];

        public List<Pt> Items = new List<Pt>();
        public string[] ItemNames = new string[0];
        public string[] ItemIcons = new string[0];
        public string[] ItemZone = new string[0];
        public string[] ItemAlt = new string[0];
        public string[] ItemCat = new string[0];

        public List<Pt> Treasures = new List<Pt>();
        public string[] TreasureNames = new string[0];
        public string[] TreasureIcons = new string[0];
        public string[] TreasureFull = new string[0];
        public string[] TreasureText = new string[0];
        public string[] TreasureCat = new string[0];
        public string[] TreasureZone = new string[0];

        public List<Pt> Statues = new List<Pt>();
        public string[] StatueNames = new string[0];
        public string[] StatueLabel = new string[0];
        public string[] StatueZone = new string[0];

        public List<Pt> Npcs = new List<Pt>();
        public string[] NpcNames = new string[0];
        public string[] NpcLabel = new string[0];
        public string[] NpcZone = new string[0];
        public NpcShop[] NpcShops = new NpcShop[0];

        public string[] Zones = new string[0];
        public List<Region> Regions = new List<Region>();

        public Dictionary<string, int> ItemPlaceTotal = new Dictionary<string, int>();
        public int ItemPlaceSum;

        public string[] BuffIcons()
        {
            List<string> seen = new List<string>();
            for (int r = 0; r < ItemCat.Length; r++)
            {
                if (ItemCat[r] != "buff") continue;
                string ic = r < ItemIcons.Length ? ItemIcons[r] : null;
                if (string.IsNullOrEmpty(ic)) continue;
                if (!seen.Contains(ic)) seen.Add(ic);
            }
            return seen.ToArray();
        }
    }

    public sealed class Data
    {
        public string AssetRoot;
        public double K = 0.5;
        public List<WorldMeta> Maps = new List<WorldMeta>();
        public Dictionary<string, World> Worlds = new Dictionary<string, World>();
        public List<Species> Species = new List<Species>();

        public World W(string key)
        {
            World w;
            return Worlds.TryGetValue(key, out w) ? w : null;
        }

        public static Data Load(string dataPath, string assetRoot)
        {
            string raw = File.ReadAllText(dataPath, Encoding.UTF8);
            int i = raw.IndexOf("=", StringComparison.Ordinal);
            int j = raw.LastIndexOf(';');
            if (i >= 0 && j > i) raw = raw.Substring(i + 1, j - i - 1);

            JavaScriptSerializer ser = new JavaScriptSerializer();
            ser.MaxJsonLength = int.MaxValue;
            Dictionary<string, object> root = ser.DeserializeObject(raw) as Dictionary<string, object>;
            if (root == null) throw new InvalidDataException("Atlas data is not a JSON object.");

            Data d = new Data();
            d.AssetRoot = assetRoot;
            object kv;
            if (root.TryGetValue("k", out kv)) d.K = J.D(kv, 0.5);
            object[] spAll = J.A(root.ContainsKey("species") ? root["species"] : null);

            object[] sp = spAll;
            if (sp != null)
            {
                object[] spi = J.A(root.ContainsKey("speciesIcons") ? root["speciesIcons"] : null);
                for (int n = 0; n < sp.Length; n++)
                {
                    Dictionary<string, object> m = J.M(sp[n]);
                    if (m == null) continue;
                    Species s = new Species();
                    s.GlobalIndex = n;
                    s.Name = J.S(m.ContainsKey("name") ? m["name"] : null);
                    s.Map = J.S(m.ContainsKey("map") ? m["map"] : null);
                    s.Kind = J.S(m.ContainsKey("kind") ? m["kind"] : null);
                    s.Zone = J.S(m.ContainsKey("zone") ? m["zone"] : null);
                    s.ColorHex = J.S(m.ContainsKey("color") ? m["color"] : null);
                    s.N = J.D(m.ContainsKey("n") ? m["n"] : null, 0);
                    s.Icon = (spi != null && n < spi.Length) ? J.S(spi[n]) : null;
                    try { s.Color = string.IsNullOrEmpty(s.ColorHex) ? Color.Gray : ColorTranslator.FromHtml(s.ColorHex); }
                    catch { s.Color = Color.Gray; }
                    d.Species.Add(s);
                }
            }

            object[] maps = J.A(root.ContainsKey("maps") ? root["maps"] : null);
            if (maps == null) maps = new object[0];
            for (int q = 0; q < maps.Length; q++)
            {
                Dictionary<string, object> m = J.M(maps[q]);
                if (m == null) continue;
                WorldMeta wm = new WorldMeta();
                wm.Key = J.S(m.ContainsKey("key") ? m["key"] : null);
                wm.Title = J.S(m.ContainsKey("title") ? m["title"] : null);
                wm.Img = J.S(m.ContainsKey("img") ? m["img"] : null);
                wm.W = (int)J.D(m.ContainsKey("w") ? m["w"] : null, 0);
                wm.H = (int)J.D(m.ContainsKey("h") ? m["h"] : null, 0);
                if (string.IsNullOrEmpty(wm.Key)) continue;
                d.Maps.Add(wm);
                d.Worlds[wm.Key] = BuildWorld(root, wm, d);
            }

            return d;
        }

        private static List<Pt> ReadPts(Dictionary<string, object> root, string table, string key)
        {
            List<Pt> outp = new List<Pt>();
            Dictionary<string, object> t = J.Sub(root, table);
            object v = null;
            if (t != null) t.TryGetValue(key, out v);
            object[] a = J.A(v);
            if (a == null) return outp;
            for (int i = 0; i < a.Length; i++)
            {
                object[] p = J.A(a[i]);
                if (p == null || p.Length < 3) continue;
                string pz = p.Length > 3 ? J.S(p[3]) : null;
                outp.Add(new Pt(J.D(p[0], 0), J.D(p[1], 0), (int)J.D(p[2], 0), pz));
            }
            return outp;
        }

        private static NpcStack ReadStack(Dictionary<string, object> m)
        {
            if (m == null) return null;
            NpcStack s = new NpcStack();
            s.Name = J.S(J.Val(m, "name"));
            s.Icon = J.S(J.Val(m, "icon"));
            s.Sym = J.S(J.Val(m, "sym"));
            s.Id = (int)J.D(J.Val(m, "id"), 0);
            s.Qty = (int)J.D(J.Val(m, "qty"), 0);
            if (s.Qty <= 0) s.Qty = 1;
            return s;
        }

        private static NpcTrade ReadTrade(Dictionary<string, object> m)
        {
            if (m == null) return null;
            NpcTrade t = new NpcTrade();
            t.Tabs = J.StrArray(m, "tabs");
            t.Currency = ReadStack(J.M(J.Val(m, "currency")));
            object[] items = J.A(J.Val(m, "items"));
            if (items != null)
            {
                for (int i = 0; i < items.Length; i++)
                {
                    Dictionary<string, object> it = J.M(items[i]);
                    if (it == null) continue;
                    NpcTradeItem x = new NpcTradeItem();
                    x.Name = J.S(J.Val(it, "name"));
                    x.Icon = J.S(J.Val(it, "icon"));
                    x.Sym = J.S(J.Val(it, "sym"));
                    x.Id = (int)J.D(J.Val(it, "id"), 0);
                    x.Price = (int)J.D(J.Val(it, "price"), 0);
                    x.Tab = (int)J.D(J.Val(it, "tab"), 0);
                    if (!string.IsNullOrEmpty(x.Name)) t.Items.Add(x);
                }
            }
            if (t.Items.Count == 0) return null;
            return t;
        }

        private static NpcOffer ReadOffer(Dictionary<string, object> m)
        {
            if (m == null) return null;
            NpcOffer o = new NpcOffer();
            object[] cost = J.A(J.Val(m, "cost"));
            object[] get = J.A(J.Val(m, "get"));
            if (cost != null)
                for (int i = 0; i < cost.Length; i++)
                {
                    NpcStack s = ReadStack(J.M(cost[i]));
                    if (s != null && !string.IsNullOrEmpty(s.Name)) o.Cost.Add(s);
                }
            if (get != null)
                for (int i = 0; i < get.Length; i++)
                {
                    NpcStack s = ReadStack(J.M(get[i]));
                    if (s != null && !string.IsNullOrEmpty(s.Name)) o.Get.Add(s);
                }
            if (o.Cost.Count == 0 && o.Get.Count == 0) return null;
            return o;
        }

        private static NpcExchange ReadExchange(Dictionary<string, object> m)
        {
            if (m == null) return null;
            NpcExchange e = new NpcExchange();
            e.Menu = J.S(J.Val(m, "menu"));
            e.Label = J.S(J.Val(m, "label"));
            if (string.IsNullOrEmpty(e.Label)) e.Label = "Exchange";
            object[] sets = J.A(J.Val(m, "sets"));
            if (sets != null)
                for (int i = 0; i < sets.Length; i++)
                {
                    NpcOffer o = ReadOffer(J.M(sets[i]));
                    if (o != null) e.Sets.Add(o);
                }
            if (e.Sets.Count == 0) return null;
            return e;
        }

        private static string StackLine(NpcStack s)
        {
            if (s == null) return "";
            string n = s.Name ?? "";
            if (s.Qty > 1) n += " x" + s.Qty.ToString(CultureInfo.InvariantCulture);
            return n;
        }

        public static string ShopSearchText(string raw, string label, NpcShop shop)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(label ?? raw); sb.Append(" \u00b7 ");
            sb.Append(raw ?? ""); sb.Append(" \u00b7 npc \u00b7 npcs \u00b7 trade \u00b7 exchange \u00b7 shop \u00b7 vendor \u00b7 trader");
            if (shop == null) return sb.ToString();
            if (shop.Trade != null)
            {
                if (shop.Trade.Currency != null && !string.IsNullOrEmpty(shop.Trade.Currency.Name))
                { sb.Append(" \u00b7 "); sb.Append(shop.Trade.Currency.Name); }
                for (int i = 0; i < shop.Trade.Tabs.Length; i++)
                { sb.Append(" \u00b7 "); sb.Append(shop.Trade.Tabs[i]); }
                for (int i = 0; i < shop.Trade.Items.Count; i++)
                { sb.Append(" \u00b7 "); sb.Append(shop.Trade.Items[i].Name); }
            }
            for (int e = 0; e < shop.Exchange.Count; e++)
            {
                NpcExchange ex = shop.Exchange[e];
                if (!string.IsNullOrEmpty(ex.Label)) { sb.Append(" \u00b7 "); sb.Append(ex.Label); }
                for (int i = 0; i < ex.Sets.Count; i++)
                {
                    for (int c = 0; c < ex.Sets[i].Cost.Count; c++)
                    { sb.Append(" \u00b7 "); sb.Append(ex.Sets[i].Cost[c].Name); }
                    for (int g = 0; g < ex.Sets[i].Get.Count; g++)
                    { sb.Append(" \u00b7 "); sb.Append(ex.Sets[i].Get[g].Name); }
                }
            }
            return sb.ToString();
        }

        private static NpcShop[] ReadShops(Dictionary<string, object> tbl, string key, int count)
        {
            NpcShop[] outp = new NpcShop[count];
            if (tbl == null) return outp;
            object v = null;
            if (!tbl.TryGetValue(key, out v)) return outp;
            object[] a = J.A(v);
            if (a == null) return outp;
            int n = Math.Min(count, a.Length);
            for (int i = 0; i < n; i++)
            {
                Dictionary<string, object> m = J.M(a[i]);
                if (m == null) continue;
                NpcShop s = new NpcShop();
                s.Script = J.S(J.Val(m, "script"));
                s.Trade = ReadTrade(J.M(J.Val(m, "trade")));
                object[] exs = J.A(J.Val(m, "exchange"));
                if (exs != null)
                    for (int e = 0; e < exs.Length; e++)
                    {
                        NpcExchange ex = ReadExchange(J.M(exs[e]));
                        if (ex != null) s.Exchange.Add(ex);
                    }
                bool trade = s.Trade != null && s.Trade.Items.Count > 0;
                bool exchange = s.Exchange.Count > 0;
                s.Kind = trade && exchange ? "Trade + Exchange" :
                         trade ? "Trade" : exchange ? "Exchange" : "Shop";
                string tag = null;
                if (s.Trade != null && s.Trade.Currency != null) tag = s.Trade.Currency.Name;
                if (string.IsNullOrEmpty(tag) && s.Exchange.Count > 0) tag = s.Exchange[0].Label;
                if (string.IsNullOrEmpty(tag)) tag = s.Kind;
                s.Tag = tag;
                s.Text = ShopSearchText(null, null, s);
                outp[i] = s;
            }
            return outp;
        }

        private static World BuildWorld(Dictionary<string, object> root, WorldMeta wm, Data d)
        {
            World w = new World();
            w.Meta = wm;
            string k = wm.Key;

            foreach (Species s in d.Species)
                if (s.Map == k) w.Species.Add(s);

            w.Spots = ReadPts(root, "spots", k);
            w.Mobs = ReadPts(root, "mobs", k);
            w.Items = ReadPts(root, "items", k);
            w.Treasures = ReadPts(root, "treasures", k);
            w.Statues = ReadPts(root, "statues", k);
            w.Npcs = ReadPts(root, "npcs", k);

            w.CreatureNames = J.StrArray(J.Sub(root, "creatureNames"), k);
            w.CreatureIcons = J.StrArray(J.Sub(root, "creatureIcons"), k);
            w.CreatureSquare = J.StrArray(J.Sub(root, "creatureSquare"), k);
            w.CreatureZone = J.StrArray(J.Sub(root, "creatureZone"), k);
            w.CreatureAlt = J.StrArray(J.Sub(root, "creatureNamesAlt"), k);
            w.CreatureLevel = J.StrArray(J.Sub(root, "creatureLevel"), k);

            Dictionary<string, object> cb = J.Sub(root, "creatureBoss");
            object cbv = null;
            if (cb != null) cb.TryGetValue(k, out cbv);
            object[] cba = J.A(cbv);
            if (cba != null)
                for (int i = 0; i < cba.Length; i++) w.Bosses.Add((int)J.D(cba[i], 0));

            w.ItemNames = J.StrArray(J.Sub(root, "itemNames"), k);
            w.ItemIcons = J.StrArray(J.Sub(root, "itemIcons"), k);
            w.ItemZone = J.StrArray(J.Sub(root, "itemZone"), k);
            w.ItemAlt = J.StrArray(J.Sub(root, "itemNamesAlt"), k);
            w.ItemCat = J.StrArray(J.Sub(root, "itemCat"), k);

            w.TreasureNames = J.StrArray(J.Sub(root, "treasureNames"), k);
            w.TreasureIcons = J.StrArray(J.Sub(root, "treasureIcons"), k);
            w.TreasureFull = J.StrArray(J.Sub(root, "treasureFull"), k);
            w.TreasureText = J.StrArray(J.Sub(root, "treasureText"), k);
            w.TreasureCat = J.StrArray(J.Sub(root, "treasureCat"), k);
            w.TreasureZone = J.StrArray(J.Sub(root, "treasureZone"), k);

            w.StatueNames = J.StrArray(J.Sub(root, "statueNames"), k);
            w.StatueLabel = J.StrArray(J.Sub(root, "statueLabel"), k);
            w.StatueZone = J.StrArray(J.Sub(root, "statueZone"), k);

            w.NpcNames = J.StrArray(J.Sub(root, "npcNames"), k);
            w.NpcLabel = J.StrArray(J.Sub(root, "npcLabel"), k);
            w.NpcZone = J.StrArray(J.Sub(root, "npcZone"), k);
            w.NpcShops = ReadShops(J.Sub(root, "npcShops"), k, w.NpcNames.Length);
            for (int i = 0; i < w.NpcShops.Length; i++)
            {
                if (w.NpcShops[i] == null) continue;
                string label = i < w.NpcLabel.Length ? w.NpcLabel[i] : null;
                if (string.IsNullOrEmpty(label)) label = w.NpcNames[i];
                w.NpcShops[i].Text = ShopSearchText(w.NpcNames[i], label, w.NpcShops[i]);
            }

            w.Zones = J.StrArray(J.Sub(root, "zones"), k);

            Dictionary<string, object> rg = J.Sub(root, "regions");
            object rgv = null;
            if (rg != null) rg.TryGetValue(k, out rgv);
            object[] rga = J.A(rgv);
            if (rga != null)
                for (int i = 0; i < rga.Length; i++)
                {
                    Dictionary<string, object> m = J.M(rga[i]);
                    if (m == null) continue;
                    Region r = new Region();
                    r.Title = J.S(m.ContainsKey("title") ? m["title"] : null);
                    r.X = J.D(m.ContainsKey("ix") ? m["ix"] : null, 0);
                    r.Y = J.D(m.ContainsKey("iy") ? m["iy"] : null, 0);
                    r.Big = m.ContainsKey("big") ? J.D(m["big"], 0) != 0 : false;
                    w.Regions.Add(r);
                }

            HashSet<string> seen = new HashSet<string>();
            for (int i = 0; i < w.Items.Count; i++)
            {
                Pt p = w.Items[i];
                string cat = (p.Ref < w.ItemCat.Length && w.ItemCat[p.Ref] != null) ? w.ItemCat[p.Ref] : "quest";
                string kk = cat + "|" + p.X.ToString("R", CultureInfo.InvariantCulture)
                                 + "," + p.Y.ToString("R", CultureInfo.InvariantCulture);
                if (!seen.Add(kk)) continue;
                int tv;
                w.ItemPlaceTotal.TryGetValue(cat, out tv);
                w.ItemPlaceTotal[cat] = tv + 1;
                w.ItemPlaceSum++;
            }

            return w;
        }
    }

    #endregion

    #region Filter / search engine

    public static class K
    {
        public const string C = "@c:";
        public const string I = "@i:";
        public const string T = "@t:";
        public const string S = "@s:";
        public const string N = "@n:";
    }

    public sealed class Row
    {
        public string Key, Label, Group, Sw, Icon, Tag, Alt, Lvl, Zone, Cat, Text;
        public NpcShop Shop;
        public string RawName;
        public Color Color;
        public int N;
        public int Depth;
        public bool Stray;
    }

    public sealed class Bucket
    {
        public string Title;
        public string Cls;
        public bool Head;
        public List<Row> Rows = new List<Row>();
        public int N;
    }

    public sealed class Group
    {
        public string Title;
        public string UnitSing, UnitPlur;
        public List<Row> Rows = new List<Row>();
        public List<Bucket> Buckets = new List<Bucket>();
        public int Total;
    }

    public sealed class XRow
    {
        public string Key, Label, Group, Zone, Cat, Lvl, Alt, Text, Map;
    }

    public sealed class Hit
    {
        public string Kind;
        public string Label, Lvl, Alt, Zone, Icon, Full, Text, Cat;
        public NpcShop Shop;
        public double X, Y, WX, WZ;
        public bool Boss;
        public Color Color;
    }

    public sealed class State
    {
        public Data D;
        public string MapKey;
        public string MapGroup = "all";
        public string Query = "";
        public bool ShowSp = true, ShowPl = true;
        public Dictionary<string, bool> On = new Dictionary<string, bool>();
        public int FilterGen = 1;
        public int IcoGen = 1;

        public string[] Terms = new string[0];
        public bool Riddle = true;
        public bool Daily;

        public Dictionary<string, bool> Match = new Dictionary<string, bool>();
        public List<Group> Groups = new List<Group>();
        public Dictionary<string, List<XRow>> XIdx;
        public Dictionary<string, List<XRow>> XByMap = new Dictionary<string, List<XRow>>();
        public HashSet<string> AllKeys = new HashSet<string>();

        public World W { get { return D == null || MapKey == null ? null : D.W(MapKey); } }
        public WorldMeta Meta { get { return W == null ? null : W.Meta; } }

        public void SetQuery(string q)
        {
            Query = q ?? "";
            if (Query.Length > 120) Query = Query.Substring(0, 120);
            Query = Query.Trim();
            BuildTerms();
            InvalidateMatch();
        }

        private void BuildTerms()
        {
            List<string> t = new List<string>();
            string low = Query.ToLowerInvariant();
            string[] parts = low.Split('|');
            for (int i = 0; i < parts.Length; i++)
            {
                string s = parts[i].Trim();
                if (s.Length > 0) t.Add(s);
            }
            Terms = t.ToArray();
            Riddle = RiddleOn(Terms);
            Daily = Query.Trim().ToLowerInvariant() == "daily quest";
        }

        private static readonly string[] LAYER_TITLES = {
            "Ore","Plant","Bosses","Monsters","Items","Treasure Maps","Statues","Traders","NPCs" };

        public static bool RiddleOn(string[] terms)
        {
            for (int i = 0; i < terms.Length; i++)
            {
                for (int g = 0; g < LAYER_TITLES.Length; g++)
                    if (LAYER_TITLES[g].ToLowerInvariant().Contains(terms[i])) return false;
            }
            return true;
        }

        public void InvalidateMatch()
        {
            Match = new Dictionary<string, bool>();
        }

        public void SetOn(string key, bool val)
        {
            bool old;
            On.TryGetValue(key, out old);
            if (old == val) return;
            On[key] = val;
            FilterGen++;
        }

        public bool IsOn(string key)
        {
            if (key == null) return false;
            bool v;
            return On.TryGetValue(key, out v) && v;
        }

        public bool Shown(string key)
        {
            if (!IsOn(key)) return false;
            if (Terms.Length == 0) return true;
            bool m;
            return Match.TryGetValue(key, out m) && m;
        }

        public static readonly string[] GROUPS = { "all", "main", "dungeons", "fields", "event", "farm", "platin", "royal" };

        public static readonly Dictionary<string, string> GRP_TITLE = new Dictionary<string, string> {
            {"main","Main Maps"},{"dungeons","Dungeons"},{"fields","Others"},{"event","Event"},
            {"farm","Farm Maps"},{"platin","Platin Level Maps"},{"royal","Royal Level Maps"} };

        private static readonly string[] G_MAIN = { "cardia","cave","kogarashi","rifloatingcity","wdmadrigal" };
        private static readonly string[] G_FIELDS = {
            "wdheaven01","wdheaven02","wdheaven03","wdheaven04","wdheaven05","wdvolcane","wdvolcanered",
            "wdcisland","wdkebaras","wdrartesia","wddarkrartesia","duflmas","wdeden","wdoldkebaras",
            "wdaurania","duflvenux","wdveruxlevel01","amethyst","dreadcove","wdoasisdesert","dudadk",
            "duflmasnew","dusatemple","wdguildwar" };
        private static readonly string[] G_DUNGEONS = {
            "dupremiummap","duominous","duominous_1","durustia","durustia_1","dubehamah","dukalgas",
            "dudreadfulcave","dusanpres","duherneos","ducontaminatedtrails","dumysticdungeon",
            "dudarkforest","duwaterdungeon","duflcavenew","dumelind","duholytemplemap","dupandora",
            "dugolemtemple","dutrigremruins","duethralgarden","duskydungeon01","dutempleofthedeep",
            "duveruxwinter","duveruxfiresmoke","duumbala01","dupiratecave","duelysianisland",
            "wdveruxjpinfinitycastle","dudracmaus","duveruxruins" };
        private static readonly string[] G_EVENT = { "wdpokeevent01","wdcandyworld","duhalloween","wdeasterevent","wdhalloweenfarm" };

        private static readonly string[] TIER_UNTIL = {
            "dudreadfulcave","ducontaminatedtrails","dugolemtemple","dutempleofthedeep","duveruxruins" };
        private static readonly string[] TIER_TITLE = { "Beginner","Shiny","Black","Soulreaver","Royal" };

        private static readonly string[] FARM_ORDER = {
            "wdheaven03","duflmas","wdoasisdesert","dreadcove","wddarkrartesia","amethyst","rifloatingcity" };
        private static readonly Dictionary<string, string> FARM_LABEL = new Dictionary<string, string> {
            {"wdheaven03","Penya Beginner"},{"duflmas","Penya Beginner 2"},{"wdoasisdesert","Penya Midgame"},
            {"dreadcove","Penya Endgame"},{"wddarkrartesia","ECP Beginner"},{"amethyst","ECP Midgame"},
            {"rifloatingcity","ECP Endgame"} };

        private static readonly Dictionary<string, string[]> DAILY = new Dictionary<string, string[]> {
            {"wdoldkebaras", new[]{"Monsters","Bosses","Items"}},
            {"wdaurania",    new[]{"Monsters","Items"}},
            {"cardia",       new[]{"Monsters","Bosses","Items"}},
            {"kogarashi",    new[]{"Monsters","Items"}} };

        private static readonly Dictionary<string, string> CATW = new Dictionary<string, string> {
            {"omega glyph","Omega Glyph"},{"lilly","Lilly"},{"quest","Quest"} };
        private static readonly Dictionary<string, string> CATTITLE = new Dictionary<string, string> {
            {"buff","World Buff Spawn"} };

        public static readonly string[] ITEM_SETS = { "omega glyph","lilly","quest","buff" };
        public static readonly string[] TRE_SETS = { "surface","cave","omega","cove" };

        public static readonly string[] GUIDE_KEYS = { "statue","(ore)","(plant)","omega","astral crystal|celestial map","treasure","daily quest" };
        public static readonly string[] GUIDE_LABEL = { "Statues","Ore","Plant","Omega","Astral","Treasure","Daily Quest" };

        public static readonly string[] MINING_SETS = {
            "adamantite (ore)|aquamarine (ore)|argentum (ore)|copper (ore)",
            "crusium (ore)|crystalwing (ore)|duckharukon (ore)|dureline (ore)",
            "emeralnd (ore)|fairy swing (ore)|hearthstone (ore)|iron (ore)",
            "lifestone (ore)|manastone (ore)|mithril (ore)|moonlightstone (ore)",
            "obsidian (ore)|opal (ore)|perl (ore)|regency (ore)",
            "retion (ore)|ruby (ore)|runecarore (ore)|sapphire (ore)",
            "adamantite (ore)|aquamarine (ore)|spiritstone (ore)|thunderboltstone (ore)" };
        public static readonly string[] MINING_LABEL = {
            "Adamantite, Aquamarine, Argentum, Copper",
            "Crusium, Crystalwing, Duckharukon, Dureline",
            "Emeralnd, Fairy Swing, Hearthstone, Iron",
            "Lifestone, Manastone, Mithril, Moonlightstone",
            "Obsidian, Opal, Perl, Regency",
            "Retion, Ruby, Runecarore, Sapphire",
            "Adamantite, Aquamarine, Spiritstone, Thunderboltstone" };

        public static bool IsDailyWorld(string map, string groupTitle)
        {
            string[] allow;
            return DAILY.TryGetValue(map, out allow) && Array.IndexOf(allow, groupTitle) >= 0;
        }

        public static string GroupOf(string k)
        {
            if (Array.IndexOf(G_MAIN, k) >= 0) return "main";
            if (Array.IndexOf(G_FIELDS, k) >= 0) return "fields";
            if (Array.IndexOf(G_DUNGEONS, k) >= 0) return "dungeons";
            if (Array.IndexOf(G_EVENT, k) >= 0) return "event";
            return "unsorted";
        }

        public List<KeyValuePair<string, List<WorldMeta>>> DungeonTiers()
        {
            List<WorldMeta> list = VisibleWorlds("dungeons");
            List<KeyValuePair<string, List<WorldMeta>>> outp = new List<KeyValuePair<string, List<WorldMeta>>>();
            int start = 0;
            for (int t = 0; t < TIER_UNTIL.Length; t++)
            {
                int end = -1;
                for (int i = start; i < list.Count; i++)
                    if (list[i].Key == TIER_UNTIL[t]) { end = i; break; }
                if (end < start) end = list.Count - 1;
                List<WorldMeta> slice = list.GetRange(start, end - start + 1);
                if (slice.Count > 0)
                    outp.Add(new KeyValuePair<string, List<WorldMeta>>(TIER_TITLE[t], slice));
                start = end + 1;
            }
            if (start < list.Count)
                outp.Add(new KeyValuePair<string, List<WorldMeta>>(
                    TIER_TITLE[TIER_TITLE.Length - 1], list.GetRange(start, list.Count - start)));
            return outp;
        }

        public List<WorldMeta> VisibleWorlds(string grp)
        {
            List<WorldMeta> outp = new List<WorldMeta>();
            if (grp == "farm")
            {
                for (int i = 0; i < FARM_ORDER.Length; i++)
                {
                    WorldMeta w = Find(FARM_ORDER[i]);
                    if (w != null) outp.Add(w);
                }
                return outp;
            }
            if (grp == "platin" || grp == "royal")
            {
                bool want = grp == "platin";
                List<WorldMeta> hits = new List<WorldMeta>();
                for (int i = 0; i < D.Maps.Count; i++)
                {
                    WorldMeta wm = D.Maps[i];
                    World w = D.W(wm.Key);
                    if (w == null) continue;
                    if (TierMinMax(w, want) >= 0) hits.Add(wm);
                }
                hits.Sort(delegate (WorldMeta a, WorldMeta b)
                {
                    int ta = TierMinMax(D.W(a.Key), want), tb = TierMinMax(D.W(b.Key), want);
                    if (ta != tb) return ta.CompareTo(tb);
                    return string.Compare(a.Title, b.Title, StringComparison.CurrentCultureIgnoreCase);
                });
                return hits;
            }
            for (int i = 0; i < D.Maps.Count; i++)
            {
                WorldMeta wm = D.Maps[i];
                if (grp == "all" || GroupOf(wm.Key) == grp) outp.Add(wm);
            }
            return outp;
        }

        public WorldMeta Find(string key)
        {
            for (int i = 0; i < D.Maps.Count; i++) if (D.Maps[i].Key == key) return D.Maps[i];
            return null;
        }

        public static int TierMinMax(World w, bool platin)
        {
            if (w == null || w.CreatureLevel == null) return -1;
            int min = int.MaxValue;
            for (int i = 0; i < w.CreatureLevel.Length; i++)
            {
                string s = w.CreatureLevel[i];
                if (string.IsNullOrEmpty(s)) continue;
                if (platin)
                {
                    if (s.Length < 2 || s[0] != 'P') continue;
                }
                else
                {
                    if (s.Length < 3 || s[0] != 'R' || s[1] != 'L') continue;
                    if (w.Meta.Key == "wdmadrigal") continue;
                }
                int n;
                if (int.TryParse(s.Substring(platin ? 1 : 2), NumberStyles.Integer, CultureInfo.InvariantCulture, out n))
                    if (n < min) min = n;
            }
            return min == int.MaxValue ? -1 : min;
        }

        public static string TierText(World w, bool platin)
        {
            int min = int.MaxValue, max = -1;
            for (int i = 0; i < w.CreatureLevel.Length; i++)
            {
                string s = w.CreatureLevel[i];
                if (string.IsNullOrEmpty(s)) continue;
                if (platin) { if (s.Length < 2 || s[0] != 'P') continue; }
                else { if (s.Length < 3 || s[0] != 'R' || s[1] != 'L') continue; if (w.Meta.Key == "wdmadrigal") continue; }
                int n;
                if (!int.TryParse(s.Substring(platin ? 1 : 2), NumberStyles.Integer, CultureInfo.InvariantCulture, out n)) continue;
                if (n < min) min = n;
                if (n > max) max = n;
            }
            if (max < 0) return "";
            return min == max
                ? string.Format(CultureInfo.InvariantCulture, "{0} {1}", platin ? "P" : "RL", min)
                : string.Format(CultureInfo.InvariantCulture, "{0} {1}-{2}", platin ? "P" : "RL", min, max);
        }

        public static string FarmLabel(string key)
        {
            string s;
            return FARM_LABEL.TryGetValue(key, out s) ? s : "";
        }

        public static bool IsOreFilter(string query)
        {
            if (string.IsNullOrEmpty(query)) return false;
            string q = query.Trim();
            if (string.Equals(q, "(ore)", StringComparison.OrdinalIgnoreCase)) return true;
            for (int i = 0; i < MINING_SETS.Length; i++)
                if (string.Equals(q, MINING_SETS[i], StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        public static string CatTitle(string cat)
        {
            string s;
            return CATTITLE.TryGetValue(cat ?? "", out s) ? s : cat;
        }

        public static string CatWord(string cat)
        {
            string s;
            return CATW.TryGetValue(cat ?? "", out s) ? s : cat;
        }


        public static string LvlText(string code)
        {
            if (string.IsNullOrEmpty(code)) return "";
            if (code[0] == 'P') return "<P " + code.Substring(1) + ">";
            if (code.StartsWith("RL", StringComparison.Ordinal)) return "<RL. " + code.Substring(2) + ">";
            return "";
        }

        private static int LvlNum(string lvl)
        {
            if (string.IsNullOrEmpty(lvl)) return -1;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < lvl.Length; i++) if (char.IsDigit(lvl[i])) sb.Append(lvl[i]);
            int n;
            return int.TryParse(sb.ToString(), out n) ? n : -1;
        }

        private static string Str(string[] a, int i)
        {
            return (a != null && i >= 0 && i < a.Length) ? a[i] : null;
        }

        private static Dictionary<int, int> Counts(List<Pt> pts)
        {
            Dictionary<int, int> c = new Dictionary<int, int>();
            for (int i = 0; i < pts.Count; i++)
            {
                int r;
                c.TryGetValue(pts[i].Ref, out r);
                c[pts[i].Ref] = r + 1;
            }
            return c;
        }

        public void SeedSwitches()
        {
            for (int i = 0; i < Groups.Count; i++)
                for (int r = 0; r < Groups[i].Rows.Count; r++)
                {
                    string k = Groups[i].Rows[r].Key;
                    if (!On.ContainsKey(k)) On[k] = true;
                }
        }

        public void BuildGroups()
        {
            Groups = new List<Group>();
            InvalidateMatch();
            World w = W;
            if (w == null) return;
            Data d = D;

            BuildSpeciesGroup(w, d, "ore", "Ore", "species", "species");
            BuildSpeciesGroup(w, d, "plant", "Plant", "species", "species");
            BuildCreatureGroups(w);
            BuildItemGroup(w);
            BuildTreasureGroup(w);
            BuildStatueGroup(w);
            BuildNpcGroup(w);
            SeedSwitches();
            SyncMatch();
        }

        private void Finish(Group g, string[] zones, string[] cats, Dictionary<string, int> place, int placeSum)
        {
            if (g.Rows.Count == 0) return;
            g.Total = 0;
            if (place != null)
            {
                for (int i = 0; i < cats.Length; i++)
                {
                    int n;
                    if (place.TryGetValue(cats[i], out n)) g.Total += n;
                }
            }
            else
            {
                for (int i = 0; i < g.Rows.Count; i++) g.Total += g.Rows[i].N;
            }
            Layout(g, zones, cats, place);
            Groups.Add(g);
        }

        private void Layout(Group g, string[] zones, string[] cats, Dictionary<string, int> place)
        {
            g.Buckets = new List<Bucket>();
            List<string> sets = new List<string>();
            for (int i = 0; i < cats.Length; i++)
            {
                bool has = false;
                for (int r = 0; r < g.Rows.Count; r++) if (g.Rows[r].Cat == cats[i]) { has = true; break; }
                if (has) sets.Add(cats[i]);
            }

            if (sets.Count > 1)
            {
                for (int i = 0; i < sets.Count; i++)
                {
                    List<Row> mine = new List<Row>();
                    for (int r = 0; r < g.Rows.Count; r++) if (g.Rows[r].Cat == sets[i]) mine.Add(g.Rows[r]);
                    if (mine.Count == 0) continue;
                    Bucket head = new Bucket();
                    head.Title = sets[i]; head.Cls = "cat"; head.Head = true; head.Rows = mine;
                    head.N = place != null ? Lookup(place, sets[i]) : -1;
                    g.Buckets.Add(head);
                    ByZone(g, mine, zones, 3);
                }
            }
            else
            {
                ByZone(g, g.Rows, zones, 2);
            }

            Bucket open = null;
            for (int b = 0; b < g.Buckets.Count; b++)
            {
                Bucket bk = g.Buckets[b];
                if (bk.Title == null)
                {
                    for (int r = 0; r < bk.Rows.Count; r++)
                    {
                        bk.Rows[r].Depth = open == null ? 0 : (open.Title != null ? 3 : 2);
                        bk.Rows[r].Stray = true;
                    }
                }
                else
                {
                    open = bk;
                    for (int r = 0; r < bk.Rows.Count; r++) bk.Rows[r].Depth = bk.Head ? 2 : 2;
                }
            }
        }

        private static int Lookup(Dictionary<string, int> d, string k)
        {
            int v;
            return d != null && d.TryGetValue(k ?? "", out v) ? v : -1;
        }

        private void ByZone(Group g, List<Row> list, string[] zones, int depth)
        {
            if (zones == null || zones.Length == 0)
            {
                Bucket b = new Bucket();
                b.Cls = "zone"; b.Rows = list; b.N = -1;
                g.Buckets.Add(b);
                return;
            }
            for (int z = 0; z < zones.Length; z++)
            {
                List<Row> sub = new List<Row>();
                for (int i = 0; i < list.Count; i++) if (list[i].Zone == zones[z]) sub.Add(list[i]);
                if (sub.Count == 0) continue;
                Bucket b = new Bucket();
                b.Title = zones[z]; b.Cls = "zone"; b.Rows = sub; b.N = -1;
                g.Buckets.Add(b);
            }
            List<Row> stray = new List<Row>();
            for (int i = 0; i < list.Count; i++) if (Array.IndexOf(zones, list[i].Zone) < 0) stray.Add(list[i]);
            if (stray.Count > 0)
            {
                Bucket b = new Bucket();
                b.Cls = "zone"; b.Rows = stray; b.N = -1;
                g.Buckets.Add(b);
            }
            if (g.Buckets.Count == 0)
            {
                Bucket b = new Bucket();
                b.Cls = "zone"; b.Rows = list; b.N = -1;
                g.Buckets.Add(b);
            }
        }

        private void BuildSpeciesGroup(World w, Data d, string kind, string title, string us, string up)
        {
            List<Row> rows = new List<Row>();
            Dictionary<int, int> cnt = Counts(w.Spots);
            for (int i = 0; i < w.Species.Count; i++)
            {
                Species s = w.Species[i];
                if (s.Kind != kind) continue;
                Row r = new Row();
                r.Key = s.Name; r.Label = s.Name; r.Group = title;
                r.Sw = s.Kind; r.Color = s.Color; r.Icon = s.Icon;
                r.Zone = string.IsNullOrEmpty(s.Zone) ? null : s.Zone;
                int n; cnt.TryGetValue(s.GlobalIndex, out n);
                r.N = n;
                rows.Add(r);
            }
            if (rows.Count == 0) return;
            rows.Sort(delegate (Row a, Row b) { return string.Compare(a.Label, b.Label, StringComparison.CurrentCultureIgnoreCase); });
            Group g = new Group();
            g.Title = title; g.UnitSing = us; g.UnitPlur = up; g.Rows = rows;
            Finish(g, w.Zones, new string[0], null, 0);
        }

        private void BuildCreatureGroups(World w)
        {
            Dictionary<int, int> cnt = Counts(w.Mobs);
            List<Row> rows = new List<Row>();
            for (int refI = 0; refI < w.CreatureNames.Length; refI++)
            {
                string n = w.CreatureNames[refI];
                int c;
                if (!cnt.TryGetValue(refI, out c) || c == 0) continue;
                bool boss = w.Bosses.Contains(refI);
                Row r = new Row();
                r.Key = K.C + n; r.Label = n; r.Group = boss ? "Bosses" : "Monsters";
                r.Sw = boss ? "boss" : "mob";
                r.Color = boss ? ColorTranslator.FromHtml("#ff6ea8") : ColorTranslator.FromHtml("#78dcff");
                r.Tag = boss ? "Boss" : "";
                string ic = Str(w.CreatureIcons, refI);
                if (string.IsNullOrEmpty(ic)) ic = Str(w.CreatureSquare, refI);
                r.Icon = ic;
                r.Alt = Str(w.CreatureAlt, refI);
                r.Lvl = Str(w.CreatureLevel, refI);
                r.Zone = Str(w.CreatureZone, refI);
                r.N = c;
                rows.Add(r);
            }
            rows.Sort(delegate (Row a, Row b)
            {
                int x = LvlNum(a.Lvl), y = LvlNum(b.Lvl);
                if (x < 0 && y < 0) return string.Compare(a.Label, b.Label, StringComparison.CurrentCultureIgnoreCase);
                if (x < 0) return 1;
                if (y < 0) return -1;
                if (x != y) return x.CompareTo(y);
                return string.Compare(a.Label, b.Label, StringComparison.CurrentCultureIgnoreCase);
            });

            List<Row> bosses = new List<Row>();
            List<Row> plain = new List<Row>();
            for (int i = 0; i < rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(rows[i].Tag)) bosses.Add(rows[i]); else plain.Add(rows[i]);
            }

            if (bosses.Count > 0)
            {
                Group g = new Group();
                g.Title = "Bosses"; g.UnitSing = "boss"; g.UnitPlur = "bosses"; g.Rows = bosses;
                Finish(g, w.Zones, new string[0], null, 0);
            }
            if (plain.Count > 0)
            {
                Group g = new Group();
                g.Title = "Monsters"; g.UnitSing = "creature"; g.UnitPlur = "creatures"; g.Rows = plain;
                Finish(g, w.Zones, new string[0], null, 0);
            }
        }

        private void BuildItemGroup(World w)
        {
            Dictionary<int, int> cnt = Counts(w.Items);
            List<Row> rows = new List<Row>();
            for (int refI = 0; refI < w.ItemNames.Length; refI++)
            {
                string n = w.ItemNames[refI];
                int c;
                if (!cnt.TryGetValue(refI, out c) || c == 0) continue;
                Row r = new Row();
                r.Key = K.I + n; r.Label = n; r.Group = "Items";
                r.Sw = "item"; r.Color = ColorTranslator.FromHtml("#ffb64a");
                r.Icon = Str(w.ItemIcons, refI);
                r.Alt = Str(w.ItemAlt, refI);
                r.Zone = Str(w.ItemZone, refI);
                string cat = Str(w.ItemCat, refI);
                r.Cat = string.IsNullOrEmpty(cat) ? "quest" : cat;
                r.N = c;
                rows.Add(r);
            }
            if (rows.Count == 0) return;
            rows.Sort(delegate (Row a, Row b) { return string.Compare(a.Label, b.Label, StringComparison.CurrentCultureIgnoreCase); });
            Group g = new Group();
            g.Title = "Items"; g.UnitSing = "kind"; g.UnitPlur = "kinds"; g.Rows = rows;
            Finish(g, w.Zones, ITEM_SETS, w.ItemPlaceTotal, w.ItemPlaceSum);
        }

        private void BuildTreasureGroup(World w)
        {
            Dictionary<int, int> cnt = Counts(w.Treasures);
            List<Row> rows = new List<Row>();
            for (int refI = 0; refI < w.TreasureNames.Length; refI++)
            {
                string n = w.TreasureNames[refI];
                int c;
                if (!cnt.TryGetValue(refI, out c) || c == 0) continue;
                Row r = new Row();
                r.Key = K.T + n; r.Label = n; r.Group = "Treasure Maps";
                r.Sw = "treasure"; r.Color = ColorTranslator.FromHtml("#d4953a");
                r.Icon = Str(w.TreasureIcons, refI);
                r.Zone = Str(w.TreasureZone, refI);
                string cat = Str(w.TreasureCat, refI);
                r.Cat = string.IsNullOrEmpty(cat) ? "surface" : cat;
                r.Text = Str(w.TreasureText, refI);
                r.N = c;
                rows.Add(r);
            }
            if (rows.Count == 0) return;
            rows.Sort(delegate (Row a, Row b) { return string.Compare(a.Label, b.Label, StringComparison.CurrentCultureIgnoreCase); });
            Group g = new Group();
            g.Title = "Treasure Maps"; g.UnitSing = "map"; g.UnitPlur = "maps"; g.Rows = rows;
            Finish(g, w.Zones, TRE_SETS, null, 0);
        }

        private void BuildStatueGroup(World w)
        {
            Dictionary<int, int> cnt = Counts(w.Statues);
            List<Row> rows = new List<Row>();
            for (int refI = 0; refI < w.StatueNames.Length; refI++)
            {
                string n = w.StatueNames[refI];
                int c;
                if (!cnt.TryGetValue(refI, out c) || c == 0) continue;
                Row r = new Row();
                r.Key = K.S + n; r.Group = "Statues";
                string lb = Str(w.StatueLabel, refI);
                r.Label = string.IsNullOrEmpty(lb) ? n : lb;
                r.Sw = "statue"; r.Color = ColorTranslator.FromHtml("#4d9de0");
                r.Zone = Str(w.StatueZone, refI);
                r.N = c;
                rows.Add(r);
            }
            if (rows.Count == 0) return;
            rows.Sort(delegate (Row a, Row b)
            {
                int c = string.Compare(a.Zone ?? "", b.Zone ?? "", StringComparison.CurrentCultureIgnoreCase);
                if (c != 0) return c;
                return string.Compare(a.Label, b.Label, StringComparison.CurrentCultureIgnoreCase);
            });
            Group g = new Group();
            g.Title = "Statues"; g.UnitSing = "statue"; g.UnitPlur = "statues"; g.Rows = rows;
            Finish(g, w.Zones, new string[0], null, 0);
        }

        public static string NpcArea(World w, Pt p, int refI)
        {
            string z = p.Zone;
            if (string.IsNullOrEmpty(z)) z = Str(w.NpcZone, refI);
            return string.IsNullOrEmpty(z) ? null : z;
        }

        private static int ZoneOrder(string[] zones, string z)
        {
            if (zones == null) return 0;
            int i = Array.IndexOf(zones, z);
            return i < 0 ? zones.Length : i;
        }

        private static List<string> NpcWords(string s)
        {
            List<string> outp = new List<string>();
            StringBuilder b = new StringBuilder();
            char prev = '\0';
            for (int i = 0; i < (s ?? "").Length; i++)
            {
                char c = s[i];
                if (char.IsUpper(c) && (char.IsLower(prev) || char.IsDigit(prev))) b.Append(' ');
                if (char.IsLetterOrDigit(c)) b.Append(char.ToLowerInvariant(c));
                else b.Append(' ');
                prev = c;
            }
            string[] parts = b.ToString().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++) outp.Add(parts[i]);
            return outp;
        }

        private static string NpcDistinctToken(string name, string label, List<string> names)
        {
            HashSet<string> skip = new HashSet<string>();
            skip.Add("npc"); skip.Add("vendor"); skip.Add("shop"); skip.Add("trader");
            skip.Add("the"); skip.Add("and"); skip.Add("of");
            List<string> mine = NpcWords(name);
            List<List<string>> others = new List<List<string>>();
            List<string> labelWords = NpcWords(label);
            for (int i = 0; i < labelWords.Count; i++) skip.Add(labelWords[i]);
            for (int i = 0; i < names.Count; i++)
                if (names[i] != name) others.Add(NpcWords(names[i]));
            List<string> only = new List<string>();
            for (int i = 0; i < mine.Count; i++)
            {
                if (skip.Contains(mine[i])) continue;
                bool seen = false;
                for (int j = 0; j < others.Count; j++)
                {
                    if (others[j].Contains(mine[i])) { seen = true; break; }
                }
                if (!seen) only.Add(mine[i]);
            }
            if (only.Count == 0) return "";
            string w = only[only.Count - 1];
            return char.ToUpperInvariant(w[0]) + w.Substring(1);
        }

        private static void AnnotateDuplicateTraders(List<Row> rows)
        {
            Dictionary<string, List<Row>> byLabel = new Dictionary<string, List<Row>>();
            for (int i = 0; i < rows.Count; i++)
            {
                string key = rows[i].Label ?? "";
                List<Row> list;
                if (!byLabel.TryGetValue(key, out list)) { list = new List<Row>(); byLabel[key] = list; }
                list.Add(rows[i]);
            }
            foreach (KeyValuePair<string, List<Row>> kv in byLabel)
            {
                List<Row> group = kv.Value;
                if (group.Count < 2) continue;
                HashSet<string> tags = new HashSet<string>();
                HashSet<string> zones = new HashSet<string>();
                HashSet<string> kinds = new HashSet<string>();
                List<string> raws = new List<string>();
                for (int i = 0; i < group.Count; i++)
                {
                    Row r = group[i];
                    tags.Add(r.Shop == null ? "" : (r.Shop.Tag ?? ""));
                    zones.Add(r.Zone ?? "");
                    kinds.Add(r.Shop == null ? "" : (r.Shop.Kind ?? ""));
                    raws.Add(r.RawName ?? "");
                }
                for (int i = 0; i < group.Count; i++)
                {
                    Row r = group[i];
                    string tag = "";
                    if (tags.Count > 1) tag = r.Shop == null ? "" : (r.Shop.Tag ?? "");
                    else if (zones.Count > 1) tag = r.Zone ?? "";
                    else if (kinds.Count > 1) tag = r.Shop == null ? "" : (r.Shop.Kind ?? "");
                    if (string.IsNullOrEmpty(tag)) tag = NpcDistinctToken(r.RawName, r.Label, raws);
                    if (string.IsNullOrEmpty(tag)) tag = (i + 1).ToString(CultureInfo.InvariantCulture);
                    if (r.Shop != null && tag != r.Shop.Kind)
                        r.Tag = r.Shop.Kind + " \u00b7 " + tag;
                }
            }
        }

        private static Row NpcRow(string raw, string label, NpcShop shop, string zone, int n)
        {
            Row r = new Row();
            r.Key = K.N + raw; r.Label = label; r.RawName = raw;
            r.Group = shop != null ? "Traders" : "NPCs";
            r.Sw = shop != null ? "trader" : "npc";
            r.Color = ColorTranslator.FromHtml(shop != null ? "#ffc857" : "#3fd16a");
            r.Shop = shop;
            r.Zone = zone; r.N = n;
            r.Tag = shop != null ? shop.Kind : "";
            r.Text = shop != null ? shop.Text : raw;
            return r;
        }

        private void BuildNpcGroup(World w)
        {
            Dictionary<int, int> cnt = Counts(w.Npcs);
            List<Row> traders = new List<Row>();
            List<Row> plain = new List<Row>();
            for (int refI = 0; refI < w.NpcNames.Length; refI++)
            {
                string n = w.NpcNames[refI];
                int c;
                if (!cnt.TryGetValue(refI, out c) || c == 0) continue;
                string label = Str(w.NpcLabel, refI);
                if (string.IsNullOrEmpty(label)) label = n;
                NpcShop shop = refI < w.NpcShops.Length ? w.NpcShops[refI] : null;
                if (shop != null) shop.Text = Data.ShopSearchText(n, label, shop);

                Dictionary<string, int> areas = new Dictionary<string, int>();
                for (int p = 0; p < w.Npcs.Count; p++)
                {
                    if (w.Npcs[p].Ref != refI) continue;
                    string az = NpcArea(w, w.Npcs[p], refI);
                    if (az == null) continue;
                    int ac;
                    areas.TryGetValue(az, out ac);
                    areas[az] = ac + 1;
                }
                List<string> areaNames = new List<string>(areas.Keys);
                areaNames.Sort(delegate (string a, string b)
                {
                    int ia = ZoneOrder(w.Zones, a), ib = ZoneOrder(w.Zones, b);
                    if (ia != ib) return ia.CompareTo(ib);
                    return string.Compare(a, b, StringComparison.CurrentCultureIgnoreCase);
                });

                List<Row> target = shop != null ? traders : plain;
                if (areaNames.Count > 1)
                {
                    for (int a = 0; a < areaNames.Count; a++)
                        target.Add(NpcRow(n, label, shop, areaNames[a], areas[areaNames[a]]));
                }
                else
                {
                    string zone = areaNames.Count == 1 ? areaNames[0] : Str(w.NpcZone, refI);
                    target.Add(NpcRow(n, label, shop, zone, c));
                }
            }

            AnnotateDuplicateTraders(traders);
            Comparison<Row> byZoneLabel = delegate (Row a, Row b)
            {
                int c = string.Compare(a.Zone ?? "", b.Zone ?? "", StringComparison.CurrentCultureIgnoreCase);
                if (c != 0) return c;
                return string.Compare(a.Label, b.Label, StringComparison.CurrentCultureIgnoreCase);
            };
            traders.Sort(byZoneLabel);
            plain.Sort(byZoneLabel);

            if (traders.Count > 0)
            {
                Group gt = new Group();
                gt.Title = "Traders"; gt.UnitSing = "trader"; gt.UnitPlur = "traders"; gt.Rows = traders;
                Finish(gt, w.Zones, new string[0], null, 0);
            }
            if (plain.Count > 0)
            {
                Group g = new Group();
                g.Title = "NPCs"; g.UnitSing = "npc"; g.UnitPlur = "npcs"; g.Rows = plain;
                Finish(g, w.Zones, new string[0], null, 0);
            }
        }

        private static bool LvlHit(Row r, string term)
        {
            if (string.IsNullOrEmpty(r.Lvl)) return false;
            string nq = term;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < nq.Length; i++)
            {
                char c = nq[i];
                if (c == '<' || c == '>' || c == '.' || char.IsWhiteSpace(c)) continue;
                sb.Append(c);
            }
            if (sb.Length == 0) return false;
            return string.Equals(r.Lvl, sb.ToString(), StringComparison.OrdinalIgnoreCase);
        }

        public void HitMap(Group g, Dictionary<string, bool> outp)
        {
            if (Daily)
            {
                bool allow = IsDailyWorld(MapKey, g.Title);
                for (int i = 0; i < g.Rows.Count; i++) outp[g.Rows[i].Key] = allow;
                return;
            }
            if (Terms.Length == 0)
            {
                for (int i = 0; i < g.Rows.Count; i++) outp[g.Rows[i].Key] = true;
                return;
            }
            string titleLow = (g.Title ?? "").ToLowerInvariant();
            bool titleHit = false;
            for (int i = 0; i < Terms.Length; i++) if (titleLow.Contains(Terms[i])) { titleHit = true; break; }

            HashSet<string> under = null;
            if (!titleHit)
                for (int b = 0; b < g.Buckets.Count; b++)
                {
                    Bucket bk = g.Buckets[b];
                    if (bk.Title == null) continue;
                    string bt = bk.Title.ToLowerInvariant();
                    bool hit = false;
                    for (int i = 0; i < Terms.Length && !hit; i++) if (bt.Contains(Terms[i])) hit = true;
                    if (!hit) continue;
                    if (under == null) under = new HashSet<string>();
                    for (int r = 0; r < bk.Rows.Count; r++) under.Add(bk.Rows[r].Key);
                }

            for (int i = 0; i < g.Rows.Count; i++)
            {
                Row r = g.Rows[i];
                if (titleHit || (under != null && under.Contains(r.Key))) { outp[r.Key] = true; continue; }
                string label = (r.Label ?? r.Key).ToLowerInvariant();
                string alt = (r.Alt ?? "").ToLowerInvariant();
                string text = (r.Text ?? "").ToLowerInvariant();
                bool ok = false;
                for (int t = 0; t < Terms.Length && !ok; t++)
                {
                    string term = Terms[t];
                    if (label.Contains(term) || alt.Contains(term) || LvlHit(r, term)) { ok = true; break; }
                    if (Riddle && text.Contains(term)) { ok = true; break; }
                }
                outp[r.Key] = ok;
            }
        }

        public void SyncMatch()
        {
            if (Terms.Length == 0) return;
            for (int i = 0; i < Groups.Count; i++)
            {
                Dictionary<string, bool> hit = new Dictionary<string, bool>();
                HitMap(Groups[i], hit);
                foreach (KeyValuePair<string, bool> kv in hit) Match[kv.Key] = kv.Value;
            }
        }

        public List<Row> AllRows()
        {
            List<Row> outp = new List<Row>();
            for (int i = 0; i < Groups.Count; i++)
            {
                Dictionary<string, bool> hit = new Dictionary<string, bool>();
                HitMap(Groups[i], hit);
                for (int r = 0; r < Groups[i].Rows.Count; r++)
                {
                    Row row = Groups[i].Rows[r];
                    bool v;
                    if (hit.TryGetValue(row.Key, out v) && v) outp.Add(row);
                }
            }
            return outp;
        }

        public List<Pt> PinsFor(string key)
        {
            List<Pt> outp = new List<Pt>();
            World w = W;
            if (w == null) return outp;
            if (key.StartsWith(K.C, StringComparison.Ordinal))
            {
                string n = key.Substring(3);
                int refI = Array.IndexOf(w.CreatureNames, n);
                if (refI < 0) return outp;
                for (int i = 0; i < w.Mobs.Count; i++) if (w.Mobs[i].Ref == refI) outp.Add(w.Mobs[i]);
            }
            else if (key.StartsWith(K.I, StringComparison.Ordinal))
            {
                string n = key.Substring(3);
                int refI = Array.IndexOf(w.ItemNames, n);
                if (refI < 0) return outp;
                for (int i = 0; i < w.Items.Count; i++) if (w.Items[i].Ref == refI) outp.Add(w.Items[i]);
            }
            else if (key.StartsWith(K.T, StringComparison.Ordinal))
            {
                string n = key.Substring(3);
                int refI = Array.IndexOf(w.TreasureNames, n);
                if (refI < 0) return outp;
                for (int i = 0; i < w.Treasures.Count; i++) if (w.Treasures[i].Ref == refI) outp.Add(w.Treasures[i]);
            }
            else if (key.StartsWith(K.S, StringComparison.Ordinal))
            {
                string n = key.Substring(3);
                int refI = Array.IndexOf(w.StatueNames, n);
                if (refI < 0) return outp;
                for (int i = 0; i < w.Statues.Count; i++) if (w.Statues[i].Ref == refI) outp.Add(w.Statues[i]);
            }
            else if (key.StartsWith(K.N, StringComparison.Ordinal))
            {
                string n = key.Substring(3);
                int refI = Array.IndexOf(w.NpcNames, n);
                if (refI < 0) return outp;
                for (int i = 0; i < w.Npcs.Count; i++) if (w.Npcs[i].Ref == refI) outp.Add(w.Npcs[i]);
            }
            else
            {
                for (int i = 0; i < D.Species.Count; i++)
                {
                    if (D.Species[i].Name != key) continue;
                    int gi = D.Species[i].GlobalIndex;
                    for (int p = 0; p < w.Spots.Count; p++) if (w.Spots[p].Ref == gi) outp.Add(w.Spots[p]);
                    break;
                }
            }
            return outp;
        }

        public List<Pt> PinsFor(string key, string zone)
        {
            List<Pt> all = PinsFor(key);
            if (string.IsNullOrEmpty(zone) || !key.StartsWith(K.N, StringComparison.Ordinal)) return all;
            World w = W;
            if (w == null) return all;
            string n = key.Substring(3);
            int refI = Array.IndexOf(w.NpcNames, n);
            if (refI < 0) return all;
            List<Pt> outp = new List<Pt>();
            for (int i = 0; i < all.Count; i++)
            {
                string z = NpcArea(w, all[i], refI);
                if (z == zone) outp.Add(all[i]);
            }
            return outp.Count > 0 ? outp : all;
        }

        public static List<Pt> PinsForOn(string mapKey, string key, Data d)
        {
            List<Pt> outp = new List<Pt>();
            World w = d.W(mapKey);
            if (w == null) return outp;
            string nm = key;
            List<Pt> src = null;
            string[] names = null;
            if (key.StartsWith(K.C, StringComparison.Ordinal)) { nm = key.Substring(3); src = w.Mobs; names = w.CreatureNames; }
            else if (key.StartsWith(K.I, StringComparison.Ordinal)) { nm = key.Substring(3); src = w.Items; names = w.ItemNames; }
            else if (key.StartsWith(K.T, StringComparison.Ordinal)) { nm = key.Substring(3); src = w.Treasures; names = w.TreasureNames; }
            else if (key.StartsWith(K.S, StringComparison.Ordinal)) { nm = key.Substring(3); src = w.Statues; names = w.StatueNames; }
            else if (key.StartsWith(K.N, StringComparison.Ordinal)) { nm = key.Substring(3); src = w.Npcs; names = w.NpcNames; }
            if (src != null)
            {
                int refI = Array.IndexOf(names, nm);
                if (refI < 0) return outp;
                for (int i = 0; i < src.Count; i++) if (src[i].Ref == refI) outp.Add(src[i]);
                return outp;
            }
            for (int i = 0; i < d.Species.Count; i++)
            {
                if (d.Species[i].Name != nm) continue;
                int gi = d.Species[i].GlobalIndex;
                for (int p = 0; p < w.Spots.Count; p++) if (w.Spots[p].Ref == gi) outp.Add(w.Spots[p]);
                break;
            }
            return outp;
        }

        public List<Pt> PinsForZone(string group, string zone)
        {
            List<Pt> outp = new List<Pt>();
            for (int i = 0; i < Groups.Count; i++)
            {
                Group g = Groups[i];
                if (!string.IsNullOrEmpty(group) && g.Title != group) continue;
                for (int r = 0; r < g.Rows.Count; r++)
                {
                    Row row = g.Rows[r];
                    if (!string.IsNullOrEmpty(zone) && row.Zone != zone) continue;
                    List<Pt> pins = PinsFor(row.Key, row.Zone);
                    for (int p = 0; p < pins.Count; p++) outp.Add(pins[p]);
                }
            }
            return outp;
        }

        public static List<string> KeysForLabel(string key, string label)
        {
            List<string> outp = new List<string>();
            if (key.StartsWith(K.S, StringComparison.Ordinal)) { outp.Add(key); return outp; }
            outp.Add(key);
            return outp;
        }

        public void EnsureXIdx()
        {
            if (XIdx != null) return;
            XIdx = new Dictionary<string, List<XRow>>();
            XByMap = new Dictionary<string, List<XRow>>();
            for (int mi = 0; mi < D.Maps.Count; mi++)
            {
                string mk = D.Maps[mi].Key;
                World w = D.W(mk);
                if (w == null) continue;

                HashSet<int> bosses = w.Bosses;
                Dictionary<int, int> cm = Counts(w.Mobs);
                for (int i = 0; i < w.CreatureNames.Length; i++)
                {
                    if (cm.ContainsKey(i) == false || cm[i] == 0) continue;
                    AddX(mk, K.C + w.CreatureNames[i], w.CreatureNames[i],
                        bosses.Contains(i) ? "Bosses" : "Monsters", Str(w.CreatureZone, i), null,
                        Str(w.CreatureLevel, i), Str(w.CreatureAlt, i), null);
                }
                Dictionary<int, int> ic = Counts(w.Items);
                for (int i = 0; i < w.ItemNames.Length; i++)
                {
                    if (!ic.ContainsKey(i) || ic[i] == 0) continue;
                    string cat = Str(w.ItemCat, i); if (string.IsNullOrEmpty(cat)) cat = "quest";
                    AddX(mk, K.I + w.ItemNames[i], w.ItemNames[i], "Items", Str(w.ItemZone, i), cat, null, Str(w.ItemAlt, i), null);
                }
                Dictionary<int, int> tc = Counts(w.Treasures);
                for (int i = 0; i < w.TreasureNames.Length; i++)
                {
                    if (!tc.ContainsKey(i) || tc[i] == 0) continue;
                    string cat = Str(w.TreasureCat, i); if (string.IsNullOrEmpty(cat)) cat = "surface";
                    AddX(mk, K.T + w.TreasureNames[i], w.TreasureNames[i], "Treasure Maps", Str(w.TreasureZone, i), cat, null, null, Str(w.TreasureText, i));
                }
                Dictionary<int, int> sc2 = Counts(w.Statues);
                for (int i = 0; i < w.StatueNames.Length; i++)
                {
                    if (!sc2.ContainsKey(i) || sc2[i] == 0) continue;
                    string lb = Str(w.StatueLabel, i); if (string.IsNullOrEmpty(lb)) lb = w.StatueNames[i];
                    AddX(mk, K.S + w.StatueNames[i], lb, "Statues", Str(w.StatueZone, i), null, null, null, null);
                }
                Dictionary<int, int> nc = Counts(w.Npcs);
                for (int i = 0; i < w.NpcNames.Length; i++)
                {
                    if (!nc.ContainsKey(i) || nc[i] == 0) continue;
                    NpcShop shop = i < w.NpcShops.Length ? w.NpcShops[i] : null;
                    string label = Str(w.NpcLabel, i);
                    if (string.IsNullOrEmpty(label)) label = w.NpcNames[i];
                    AddX(mk, K.N + w.NpcNames[i], label,
                         shop != null ? "Traders" : "NPCs", Str(w.NpcZone, i), null, null, null,
                         shop != null ? shop.Text : w.NpcNames[i]);
                }
            }
            for (int si = 0; si < D.Species.Count; si++)
            {
                Species s = D.Species[si];
                World w = D.W(s.Map);
                if (w == null) continue;
                Dictionary<int, int> cnt = Counts(w.Spots);
                if (!cnt.ContainsKey(s.GlobalIndex) || cnt[s.GlobalIndex] == 0) continue;
                AddX(s.Map, s.Name, s.Name, s.Kind == "ore" ? "Ore" : "Plant",
                     string.IsNullOrEmpty(s.Zone) ? null : s.Zone, null, null, null, null);
            }
            foreach (KeyValuePair<string, List<XRow>> kv in XIdx)
            {
                AllKeys.Add(kv.Key);
                if (!On.ContainsKey(kv.Key)) On[kv.Key] = true;
            }
        }

        private void AddX(string map, string key, string label, string group, string zone, string cat, string lvl, string alt, string text)
        {
            XRow r = new XRow();
            r.Map = map; r.Key = key; r.Label = label; r.Group = group;
            r.Zone = zone; r.Cat = cat; r.Lvl = lvl; r.Alt = alt; r.Text = text;

            List<XRow> list;
            if (!XIdx.TryGetValue(key, out list)) { list = new List<XRow>(); XIdx[key] = list; }
            list.Add(r);

            List<XRow> byMap;
            if (!XByMap.TryGetValue(map, out byMap)) { byMap = new List<XRow>(); XByMap[map] = byMap; }
            byMap.Add(r);
        }

        private bool XHit(XRow r, string mapHit)
        {
            if (Daily) return IsDailyWorld(r.Map, r.Group);
            if (Terms.Length == 0) return false;
            if (mapHit != null) return true;
            string label = (r.Label ?? "").ToLowerInvariant();
            string zone = (r.Zone ?? "").ToLowerInvariant();
            string catw = (CatWord(r.Cat) ?? "").ToLowerInvariant();
            string grp = (r.Group ?? "").ToLowerInvariant();
            string alt = (r.Alt ?? "").ToLowerInvariant();
            string text = (r.Text ?? "").ToLowerInvariant();
            for (int i = 0; i < Terms.Length; i++)
            {
                string t = Terms[i];
                if (label.Contains(t) || zone.Contains(t) || catw.Contains(t) || grp.Contains(t)) return true;
                if (Riddle && text.Contains(t)) return true;
                if (!string.IsNullOrEmpty(r.Lvl))
                {
                    string nq = t.Replace("<", "").Replace(">", "").Replace(".", "").Replace(" ", "");
                    if (nq.Length > 0 && string.Equals(r.Lvl, nq, StringComparison.OrdinalIgnoreCase)) return true;
                }
                if (!string.IsNullOrEmpty(alt) && alt.Contains(t)) return true;
            }
            return false;
        }

        public List<KeyValuePair<string, List<XRow>>> OtherWorlds(int limitPerWorld)
        {
            List<KeyValuePair<string, List<XRow>>> outp = new List<KeyValuePair<string, List<XRow>>>();
            if (Terms.Length == 0) return outp;
            EnsureXIdx();

            for (int mi = 0; mi < D.Maps.Count; mi++)
            {
                string mk = D.Maps[mi].Key;
                if (mk == MapKey) continue;
                WorldMeta wm = D.Maps[mi];
                string titleLow = (wm.Title ?? "").ToLowerInvariant();
                string mapHit = null;
                for (int i = 0; i < Terms.Length; i++) if (titleLow.Contains(Terms[i])) { mapHit = mk; break; }

                List<XRow> pool;
                if (!XByMap.TryGetValue(mk, out pool)) continue;

                List<XRow> hits = new List<XRow>();
                HashSet<string> seenKeys = new HashSet<string>();
                for (int i = 0; i < pool.Count; i++)
                {
                    XRow r = pool[i];
                    if (r.Key == MapKey) continue;
                    if (!seenKeys.Add(r.Key)) continue;
                    if (XHit(r, mapHit)) hits.Add(r);
                }
                if (hits.Count > 0) outp.Add(new KeyValuePair<string, List<XRow>>(mk, hits));
            }
            return outp;
        }

        public List<string> GlobalMatchingKeys()
        {
            List<string> outp = new List<string>();
            if (Terms.Length == 0)
            {
                foreach (string k in On.Keys) outp.Add(k);
                return outp;
            }
            EnsureXIdx();
            foreach (KeyValuePair<string, List<XRow>> kv in XIdx)
            {
                for (int i = 0; i < kv.Value.Count; i++)
                    if (XHit(kv.Value[i], null)) { outp.Add(kv.Key); break; }
            }
            return outp;
        }

        public int KeysOn
        {
            get
            {
                int n = 0;
                foreach (KeyValuePair<string, bool> kv in On) if (kv.Value) n++;
                return n;
            }
        }

        public int[] StatusCounts()
        {
            World w = W;
            int[] outp = new int[7];
            if (w == null) return outp;

            HashSet<int> spSeen = new HashSet<int>();
            for (int i = 0; i < w.Spots.Count; i++)
            {
                Pt p = w.Spots[i];
                if (p.Ref < 0 || p.Ref >= D.Species.Count) continue;
                if (!Shown(D.Species[p.Ref].Name)) continue;
                outp[0]++;
                spSeen.Add(p.Ref);
            }
            outp[1] = spSeen.Count;

            if (ShowSp)
            {
                for (int i = 0; i < w.Mobs.Count; i++)
                    if (Shown(K.C + Str(w.CreatureNames, w.Mobs[i].Ref))) outp[2]++;
                HashSet<string> seen = new HashSet<string>();
                for (int i = 0; i < w.Items.Count; i++)
                {
                    Pt p = w.Items[i];
                    if (!Shown(K.I + Str(w.ItemNames, p.Ref))) continue;
                    string pk = p.X.ToString("R", CultureInfo.InvariantCulture) + "," + p.Y.ToString("R", CultureInfo.InvariantCulture);
                    if (!seen.Add(pk)) continue;
                    outp[3]++;
                }
                for (int i = 0; i < w.Treasures.Count; i++)
                    if (Shown(K.T + Str(w.TreasureNames, w.Treasures[i].Ref))) outp[4]++;
                for (int i = 0; i < w.Statues.Count; i++)
                    if (Shown(K.S + Str(w.StatueNames, w.Statues[i].Ref))) outp[5]++;
                for (int i = 0; i < w.Npcs.Count; i++)
                    if (Shown(K.N + Str(w.NpcNames, w.Npcs[i].Ref))) outp[6]++;
            }
            return outp;
        }

        public void OnlyThis(string keep)
        {
            List<string> keys = new List<string>(On.Keys);
            for (int i = 0; i < keys.Count; i++) SetOn(keys[i], keys[i] == keep);
        }

        public int SetAllGlobal(bool on)
        {
            List<string> keys = GlobalMatchingKeys();
            for (int i = 0; i < keys.Count; i++) SetOn(keys[i], on);
            return keys.Count;
        }

        public int SetAllHere(bool on)
        {
            List<Row> rows = AllRows();
            for (int i = 0; i < rows.Count; i++) SetOn(rows[i].Key, on);
            return rows.Count;
        }
    }

    #endregion

    #region Icon cache with silhouette rims

    internal static class Icons
    {
        private static Dictionary<string, Bitmap> cache = new Dictionary<string, Bitmap>();
        private static Dictionary<string, Bitmap> rims = new Dictionary<string, Bitmap>();
        private static Dictionary<string, long> rimBytes = new Dictionary<string, long>();
        private static long bytes = 0;
        private static readonly object gate = new object();
        private static int gen = 1;

        private const long BUDGET = 120L * 1024L * 1024L;

        internal static int Gen { get { return gen; } }
        internal static int RimCount { get { lock (gate) { return rims.Count; } } }

        internal static void Clear()
        {
            lock (gate)
            {
                cache = new Dictionary<string, Bitmap>();
                rims = new Dictionary<string, Bitmap>();
                rimBytes = new Dictionary<string, long>();
                bytes = 0;
            }
        }

        internal static string Full(Data d, string rel)
        {
            if (string.IsNullOrEmpty(rel)) return null;
            if (Path.IsPathRooted(rel)) return rel;
            return Path.Combine(d.AssetRoot, rel.Replace('/', Path.DirectorySeparatorChar));
        }

        internal static Bitmap Img(Data d, string rel)
        {
            if (string.IsNullOrEmpty(rel)) return null;
            string key = rel;
            lock (gate)
            {
                Bitmap b;
                if (cache.TryGetValue(key, out b)) return b;
            }
            string p = Full(d, rel);
            Bitmap got = null;
            try
            {
                if (p != null && File.Exists(p))
                {
                    using (FileStream fs = new FileStream(p, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (Image src = Image.FromStream(fs))
                    {
                        Bitmap tmp = new Bitmap(src);
                        got = new Bitmap(tmp);
                        tmp.Dispose();
                    }
                }
            }
            catch { got = null; }
            lock (gate) { cache[key] = got; gen++; }
            return got;
        }

        internal static int Round4(double n)
        {
            int i = (int)Math.Round(n / 4.0) * 4;
            return i < 4 ? 4 : i;
        }

        internal static Bitmap Rim(Data d, string rel, int n, Color col)
        {
            n = Round4(n);
            string key = rel + "#" + n + "#" + col.ToArgb();
            lock (gate)
            {
                Bitmap b;
                if (rims.TryGetValue(key, out b)) return b;
                while (bytes > BUDGET && rims.Count > 0)
                {
                    List<string> drop = new List<string>(rimBytes.Keys);
                    for (int i = 0; i < drop.Count / 4; i++)
                    {
                        long sz;
                        if (rimBytes.TryGetValue(drop[i], out sz)) bytes -= sz;
                        rims.Remove(drop[i]);
                        rimBytes.Remove(drop[i]);
                    }
                }
            }

            Bitmap body = Img(d, rel);
            if (body == null) { lock (gate) { rims[key] = null; } return null; }

            int bw = body.Width, bh = body.Height;
            float s = (float)n / (bw > bh ? bw : bh);
            int dw = Math.Max(1, (int)Math.Round(bw * s)), dh = Math.Max(1, (int)Math.Round(bh * s));
            if (dw > n) dw = n;
            if (dh > n) dh = n;
            int ox = (n - dw) / 2, oy = (n - dh) / 2;

            Bitmap scaled = null;
            try
            {
                scaled = new Bitmap(dw, dh, PixelFormat.Format32bppArgb);
                using (Graphics sg = Graphics.FromImage(scaled))
                {
                    sg.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    sg.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    sg.CompositingQuality = CompositingQuality.HighQuality;
                    sg.DrawImage(body, new Rectangle(0, 0, dw, dh));
                }
            }
            catch { if (scaled != null) { scaled.Dispose(); scaled = null; } }
            Bitmap src = scaled;
            if (src == null)
            {
                dw = Math.Min(bw, n); dh = Math.Min(bh, n);
                ox = (n - dw) / 2; oy = (n - dh) / 2;
                src = body;
            }

            int w = n + 2;
            bool[] mask = new bool[w * w];
            bool[] body2 = new bool[w * w];
            for (int y = 0; y < dh; y++)
                for (int x = 0; x < dw; x++)
                {
                    Color c;
                    try { c = src.GetPixel(x, y); }
                    catch { continue; }
                    if (c.A <= 32) continue;
                    body2[(y + oy + 1) * w + (x + ox + 1)] = true;
                }
            for (int i = 0; i < w * w; i++) mask[i] = body2[i];
            for (int y = 0; y < w; y++)
                for (int x = 0; x < w; x++)
                {
                    int i = y * w + x;
                    if (mask[i]) continue;
                    bool near = false;
                    for (int dy = -1; dy <= 1 && !near; dy++)
                        for (int dx = -1; dx <= 1; dx++)
                        {
                            int nx = x + dx, ny = y + dy;
                            if (nx < 0 || ny < 0 || nx >= w || ny >= w) continue;
                            if (!body2[ny * w + nx]) continue;
                            near = true; break;
                        }
                    mask[i] = near;
                }

            Bitmap outp = new Bitmap(w, w, PixelFormat.Format32bppArgb);
            using (Graphics gg = Graphics.FromImage(outp))
            {
                gg.Clear(Color.Transparent);
                gg.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gg.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gg.CompositingQuality = CompositingQuality.HighQuality;
                gg.DrawImage(src, new Rectangle(ox + 1, oy + 1, dw, dh));
                using (SolidBrush br = new SolidBrush(Color.FromArgb(255, col)))
                {
                    for (int y = 0; y < w; y++)
                        for (int x = 0; x < w; x++)
                        {
                            int i = y * w + x;
                            if (!mask[i] || body2[i]) continue;
                            gg.FillRectangle(br, x, y, 1, 1);
                        }
                }
            }
            if (scaled != null) scaled.Dispose();
            body2 = null;
            lock (gate)
            {
                long bmpBytes = (long)w * w * 4;
                rims[key] = outp;
                rimBytes[key] = bmpBytes;
                bytes += bmpBytes;
            }
            return outp;
        }
    }

    #endregion

    #region A PNG reader for the big maps

    internal static class Png
    {
        private static uint BE32(byte[] b, int i)
        {
            return ((uint)b[i] << 24) | ((uint)b[i + 1] << 16) | ((uint)b[i + 2] << 8) | (uint)b[i + 3];
        }

        private static bool Tag(byte[] b, int i, string t)
        {
            return b[i] == t[0] && b[i + 1] == t[1] && b[i + 2] == t[2] && b[i + 3] == t[3];
        }

        internal static bool Size(byte[] data, out int w, out int h)
        {
            w = 0; h = 0;
            if (data == null || data.Length < 33) return false;
            if (data[0] != 137 || data[1] != 80 || data[2] != 78 || data[3] != 71) return false;
            if (!Tag(data, 12, "IHDR")) return false;
            w = (int)BE32(data, 16);
            h = (int)BE32(data, 20);
            return w > 0 && h > 0;
        }

        internal static Bitmap Load(byte[] data, int maxDim, out int outW, out int outH)
        {
            outW = 0; outH = 0;
            int w, h;
            if (!Size(data, out w, out h)) return null;

            int depth = 0, color = 0, interlace = 0;
            byte[] pal = null, trns = null;
            byte[] comp = null;
            try
            {
                MemoryStream idat = new MemoryStream();
                int p = 8;
                while (p + 8 <= data.Length)
                {
                    int len = (int)BE32(data, p);
                    if (len < 0 || p + 8 + len + 4 > data.Length) { idat.Dispose(); return null; }
                    int d = p + 8;
                    if (Tag(data, p + 4, "IHDR"))
                    {
                        depth = data[d + 8]; color = data[d + 9]; interlace = data[d + 12];
                    }
                    else if (Tag(data, p + 4, "PLTE"))
                    {
                        pal = new byte[len]; Buffer.BlockCopy(data, d, pal, 0, len);
                    }
                    else if (Tag(data, p + 4, "tRNS"))
                    {
                        trns = new byte[len]; Buffer.BlockCopy(data, d, trns, 0, len);
                    }
                    else if (Tag(data, p + 4, "IDAT")) idat.Write(data, d, len);
                    else if (Tag(data, p + 4, "IEND")) break;
                    p = d + len + 4;
                }
                comp = idat.ToArray();
                idat.Dispose();
            }
            catch { return null; }

            if (interlace != 0) return null;
            if (color != 0 && color != 2 && color != 3 && color != 4 && color != 6) return null;
            if (depth != 8 && depth != 16 && !(color == 3 && depth >= 1 && depth <= 8)) return null;
            if (color == 3 && pal == null) return null;

            int channels = color == 0 ? 1 : color == 2 ? 3 : color == 3 ? 1 : color == 4 ? 2 : 4;
            int bpr = (w * channels * depth + 7) / 8;
            int unit = Math.Max(1, channels * depth / 8);

            int scale = 1;
            int longSide = Math.Max(w, h);
            if (maxDim > 0 && longSide > maxDim) scale = (longSide + maxDim - 1) / maxDim;
            int ow = Math.Max(1, (w + scale - 1) / scale);
            int oh = Math.Max(1, (h + scale - 1) / scale);

            Bitmap bmp = null;
            BitmapData bits = null;
            try
            {
                bmp = new Bitmap(ow, oh, PixelFormat.Format32bppArgb);
                bits = bmp.LockBits(new Rectangle(0, 0, ow, oh), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);

                int z = 0;
                if (comp.Length > 2 && (comp[0] & 0x0F) == 8 &&
                    ((((int)comp[0] << 8) | comp[1]) % 31) == 0) z = 2;

                byte[] prev = new byte[bpr];
                byte[] cur = new byte[bpr];
                byte[] line = new byte[ow * 4];
                int row = 0;

                using (MemoryStream ms = new MemoryStream(comp, z, comp.Length - z))
                using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress))
                {
                    for (int y = 0; y < h; y++)
                    {
                        int filter = ds.ReadByte();
                        if (filter < 0) throw new InvalidDataException("The picture ends early.");
                        if (!Read(ds, cur)) throw new InvalidDataException("The picture ends early.");
                        Unfilter(filter, cur, prev, unit);

                        if (y % scale == 0 && row < oh)
                        {
                            WriteRow(cur, w, scale, ow, line, depth, color, channels, pal, trns);
                            Marshal.Copy(line, 0, IntPtr.Add(bits.Scan0, row * bits.Stride), ow * 4);
                            row++;
                        }
                        byte[] swap = prev; prev = cur; cur = swap;
                    }
                }
                bmp.UnlockBits(bits);
                bits = null;
                outW = ow; outH = oh;
                return bmp;
            }
            catch
            {
                if (bits != null) { try { bmp.UnlockBits(bits); } catch { } }
                if (bmp != null) bmp.Dispose();
                return null;
            }
        }

        private static bool Read(DeflateStream ds, byte[] buf)
        {
            int got = 0;
            while (got < buf.Length)
            {
                int n = ds.Read(buf, got, buf.Length - got);
                if (n <= 0) return false;
                got += n;
            }
            return true;
        }

        private static void Unfilter(int filter, byte[] cur, byte[] prev, int unit)
        {
            switch (filter)
            {
                case 0: break;
                case 1:
                    for (int i = unit; i < cur.Length; i++) cur[i] = (byte)(cur[i] + cur[i - unit]);
                    break;
                case 2:
                    for (int i = 0; i < cur.Length; i++) cur[i] = (byte)(cur[i] + prev[i]);
                    break;
                case 3:
                    for (int i = 0; i < cur.Length; i++)
                    {
                        int a = i >= unit ? cur[i - unit] : 0;
                        cur[i] = (byte)(cur[i] + ((a + prev[i]) >> 1));
                    }
                    break;
                case 4:
                    for (int i = 0; i < cur.Length; i++)
                    {
                        int a = i >= unit ? cur[i - unit] : 0;
                        int b = prev[i];
                        int c = i >= unit ? prev[i - unit] : 0;
                        int pp = a + b - c;
                        int pa = Math.Abs(pp - a), pb = Math.Abs(pp - b), pc = Math.Abs(pp - c);
                        int pr = (pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c);
                        cur[i] = (byte)(cur[i] + pr);
                    }
                    break;
                default: break;
            }
        }

        private static void WriteRow(byte[] cur, int w, int scale, int ow, byte[] line,
                                     int depth, int color, int channels, byte[] pal, byte[] trns)
        {
            int step = depth == 16 ? 2 : 1;
            for (int i = 0; i < ow; i++)
            {
                int sx = Math.Min(w - 1, i * scale);
                int r = 0, g = 0, b = 0, a = 255;
                if (color == 3)
                {
                    int idx;
                    if (depth == 8) idx = cur[sx];
                    else
                    {
                        int per = 8 / depth;
                        int bi = sx / per;
                        int shift = 8 - depth * (sx % per + 1);
                        idx = (cur[bi] >> shift) & ((1 << depth) - 1);
                    }
                    int pi = idx * 3;
                    if (pi + 2 < pal.Length) { r = pal[pi]; g = pal[pi + 1]; b = pal[pi + 2]; }
                    if (trns != null && idx < trns.Length) a = trns[idx];
                }
                else if (color == 0 || color == 4)
                {
                    int at = sx * channels * step;
                    r = g = b = cur[at];
                    if (color == 4) a = cur[at + step];
                }
                else
                {
                    int at = sx * channels * step;
                    r = cur[at]; g = cur[at + step]; b = cur[at + 2 * step];
                    if (color == 6) a = cur[at + 3 * step];
                }
                int o = i * 4;
                line[o] = (byte)b; line[o + 1] = (byte)g; line[o + 2] = (byte)r; line[o + 3] = (byte)a;
            }
        }
    }

    #endregion

    #region Canvas

    public sealed class Canvas : Control
    {
        private const int MARGIN = 200;

        public State S;

        public double Vx, Vy, Sc = 1;
        public bool ViewChosen;

        private int picW, picH;
        public double PicW { get { return picW > 0 ? picW : (S != null && S.Meta != null ? S.Meta.W : 0); } }
        public double PicH { get { return picH > 0 ? picH : (S != null && S.Meta != null ? S.Meta.H : 0); } }
        public double Ppx { get { WorldMeta m = S == null ? null : S.Meta; return (m != null && m.W > 0 && PicW > 0) ? PicW / m.W : 1.0; } }
        public double Ppy { get { WorldMeta m = S == null ? null : S.Meta; return (m != null && m.H > 0 && PicH > 0) ? PicH / m.H : 1.0; } }

        private void LoadPictureMetrics(WorldMeta m)
        {
            picW = 0; picH = 0;
            if (m == null || S == null || S.D == null) return;
            try
            {
                string p = Icons.Full(S.D, m.Img);
                if (string.IsNullOrEmpty(p) || !File.Exists(p)) return;
                byte[] head = new byte[24];
                using (FileStream fs = new FileStream(p, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (fs.Read(head, 0, 24) < 24) return;
                }
                if (head[0] != 0x89 || head[1] != 0x50 || head[2] != 0x4E || head[3] != 0x47) return;
                int w = (head[16] << 24) | (head[17] << 16) | (head[18] << 8) | head[19];
                int h = (head[20] << 24) | (head[21] << 16) | (head[22] << 8) | head[23];
                if (w > 0 && h > 0) { picW = w; picH = h; }
            }
            catch { picW = 0; picH = 0; }
        }

        public string BackgroundStatus = "";

        public event EventHandler ViewChanged;
        public event EventHandler HoverChanged;
        public event EventHandler Picked;

        private Bitmap layer;
        private string layerKey = "";
        private double layerSc, layerVx, layerVy;
        private int layerW, layerH;

        private Image bg;
        private string bgKey = "";

        private Hit hover;
        public Hit Hover { get { return hover; } }
        private Point downAt = Point.Empty;
        private bool dragging, dragMoved;
        private Point lastAt;

        public int BuffRotT;
        private Timer buffTimer;

        private Color CBack = Color.FromArgb(10, 14, 19);
        private Color CFg = Color.FromArgb(231, 238, 247);
        private Color CMut = Color.FromArgb(132, 150, 172);
        private Color CLine = Color.FromArgb(38, 51, 69);

        public void SetBack(Color c)
        {
            if (c == CBack) return;
            CBack = c;
            BackColor = c;
            InvalidateLayer();
            CancelPreview();
            Invalidate();
        }

        public Canvas()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
            BackColor = CBack;
            DoubleBuffered = true;
            Cursor = Cursors.Hand;
            TabStop = true;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (buffTimer != null) { buffTimer.Stop(); buffTimer.Dispose(); buffTimer = null; }
                if (settleTimer != null) { settleTimer.Stop(); settleTimer.Dispose(); settleTimer = null; }
                lock (bgGate)
                {
                    if (bg != null) { bg.Dispose(); bg = null; }
                    if (bgMem != null) { bgMem.Dispose(); bgMem = null; }
                }
                if (layer != null) { layer.Dispose(); layer = null; }
            }
            base.Dispose(disposing);
        }

        public void SetWorld(string key, bool keepView)
        {
            string prev = S.MapKey;
            S.MapKey = key;
            if (bgKey != key)
            {
                lock (bgGate)
                {
                    if (bg != null) { bg.Dispose(); if (bgMem != null) bgMem.Dispose(); }
                    bg = null; bgMem = null;
                }
                bgKey = key;
            }
            LoadPictureMetrics(S.Meta);
            if (!keepView || prev != key || !ViewChosen) Fit();
            InvalidateLayer();
            CancelPreview();
            UpdateBuffTimer();
            OnViewChanged();
        }

        public double FitScale()
        {
            WorldMeta m = S.Meta;
            if (m == null || Width <= 0 || Height <= 0) return 1;
            if (PicW <= 0 || PicH <= 0) return 1;
            return Math.Min((double)Width / PicW, (double)Height / PicH);
        }

        public void Fit()
        {
            WorldMeta m = S.Meta;
            if (m == null) return;
            Sc = FitScale();
            Vx = PicW / 2.0;
            Vy = PicH / 2.0;
            ViewChosen = true;
            InvalidateLayer();
            CancelPreview();
            Invalidate();
            OnViewChanged();
        }

        public void ClampView()
        {
            WorldMeta m = S.Meta;
            if (m == null || PicW <= 0 || PicH <= 0) return;
            double halfW = Width / 2.0 / Sc, halfH = Height / 2.0 / Sc;
            double pw = PicW, ph = PicH;
            const double keep = 60;
            if (pw + keep * 2 < halfW * 2) { Vx = pw / 2.0; }
            else
            {
                double minX = Math.Min(pw / 2.0, halfW - keep), maxX = Math.Max(pw / 2.0, pw - halfW + keep);
                Vx = Math.Min(maxX, Math.Max(minX, Vx));
            }
            if (ph + keep * 2 < halfH * 2) { Vy = ph / 2.0; }
            else
            {
                double minY = Math.Min(ph / 2.0, halfH - keep), maxY = Math.Max(ph / 2.0, ph - halfH + keep);
                Vy = Math.Min(maxY, Math.Max(minY, Vy));
            }
        }

        public void ZoomAt(double px, double py, double factor)
        {
            double ix = ((px - Width / 2.0) / Sc + Vx) / Ppx;
            double iy = ((py - Height / 2.0) / Sc + Vy) / Ppy;
            Sc = Math.Min(24, Math.Max(0.012, Sc * factor));
            Vx = ix * Ppx - (px - Width / 2.0) / Sc;
            Vy = iy * Ppy - (py - Height / 2.0) / Sc;
            ClampView();
            InvalidateLayerAfterMotion();
            Invalidate();
            OnViewChanged();
        }

        public void FitPoints(List<Pt> pts)
        {
            if (pts == null || pts.Count == 0 || Width <= 0 || Height <= 0) return;
            double x0 = double.MaxValue, y0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue;
            for (int i = 0; i < pts.Count; i++)
            {
                Pt p = pts[i];
                if (p.X < x0) x0 = p.X; if (p.X > x1) x1 = p.X;
                if (p.Y < y0) y0 = p.Y; if (p.Y > y1) y1 = p.Y;
            }
            double bw = Math.Max(200, (x1 - x0) * Ppx) + 140, bh = Math.Max(200, (y1 - y0) * Ppy) + 140;
            Sc = Math.Min(8, Math.Min(Width / bw, Height / bh));
            Vx = (x0 + x1) / 2.0 * Ppx;
            Vy = (y0 + y1) / 2.0 * Ppy;
            ViewChosen = true;
            ClampView();
            InvalidateLayer();
            CancelPreview();
            Invalidate();
            OnViewChanged();
        }

        public void CentreOn(double ix, double iy)
        {
            Vx = ix * Ppx; Vy = iy * Ppy; ViewChosen = true;
            ClampView();
            InvalidateLayer();
            CancelPreview();
            Invalidate();
            OnViewChanged();
        }

        public void ToWorld(double ix, double iy, out double wx, out double wz)
        {
            WorldMeta m = S.Meta;
            double h = m == null ? 0 : m.H;
            wx = ix / S.D.K;
            wz = (h - iy) / S.D.K;
        }

        private void OnViewChanged()
        {
            if (ViewChanged != null) ViewChanged(this, EventArgs.Empty);
        }

        private const int BG_FULL_PIXELS = 4 * 1024 * 1024;
        private const int BG_MAX_DIM = 2048;

        private bool bgTried;
        private bool bgFailed;
        public bool BackgroundReady { get { return bgTried; } }

        private Image Background()
        {
            WorldMeta m = S.Meta;
            if (m == null) return null;
            lock (bgGate)
            {
                if (bg != null && bgKey == m.Key) return bg;
                if (bg != null) { bg.Dispose(); if (bgMem != null) bgMem.Dispose(); bg = null; bgMem = null; }
                if (bgFailed && bgKey == m.Key) return null;
            }
            bgKey = m.Key;
            bgFailed = false;
            string p = Icons.Full(S.D, m.Img);
            try
            {
                if (string.IsNullOrEmpty(p) || !File.Exists(p))
                {
                    bg = null;
                    BackgroundStatus = m.Img + " is missing.";
                    bgTried = true;
                    return null;
                }

                byte[] bytes = File.ReadAllBytes(p);

                int pw, ph;
                if (Png.Size(bytes, out pw, out ph) && pw > 0 && ph > 0) { picW = pw; picH = ph; }
                if (pw > 0 && ph > 0 && (long)pw * ph > BG_FULL_PIXELS)
                {
                    int rw, rh;
                    Bitmap small = Png.Load(bytes, BG_MAX_DIM, out rw, out rh);
                    if (small != null)
                    {
                        lock (bgGate) { bg = small; bgMem = null; }
                        BackgroundStatus = "";
                        bgTried = true;
                        return bg;
                    }
                }

                MemoryStream ms = new MemoryStream(bytes);
                Image img = Image.FromStream(ms);
                picW = img.Width; picH = img.Height;
                lock (bgGate) { bg = img; bgMem = ms; }
                BackgroundStatus = "";
            }
            catch (Exception ex)
            {
                bg = null;
                bgFailed = true;
                BackgroundStatus = m.Img + " could not be read: " + ex.Message;
            }
            bgTried = true;
            return bg;
        }

        private readonly object bgGate = new object();
        private MemoryStream bgMem;

        private string LayerKeyNow()
        {
            return S.MapKey + "|" + S.ShowSp + "|" + S.ShowPl + "|" + S.FilterGen
                 + "|" + Icons.Gen + "|" + BuffRotT;
        }

        public void InvalidateLayer()
        {
            layerKey = "";
        }

        public void InvalidateLayerAfterMotion()
        {
            layerKey = "";
            ArmSettle();
        }

        private bool settling;
        private Timer settleTimer;

        private void ArmSettle()
        {
            settling = true;
            if (settleTimer == null)
            {
                settleTimer = new Timer();
                settleTimer.Interval = 170;
                settleTimer.Tick += delegate
                {
                    settleTimer.Stop();
                    settling = false;
                    layerKey = "";
                    Invalidate();
                };
            }
            settleTimer.Stop();
            settleTimer.Start();
        }

        public void CancelPreview()
        {
            settling = false;
            if (settleTimer != null) settleTimer.Stop();
        }

        private void EnsureLayer()
        {
            if (Width <= 0 || Height <= 0) return;
            int lw = Width + MARGIN * 2, lh = Height + MARGIN * 2;
            string lk = LayerKeyNow();
            bool sizeOk = layer != null && layerW == lw && layerH == lh;
            bool panOk = layer != null
                         && Math.Abs((Vx - layerVx) * layerSc) <= MARGIN
                         && Math.Abs((Vy - layerVy) * layerSc) <= MARGIN;
            if (sizeOk && panOk && layerKey == lk && layerSc == Sc) return;

            if (layer == null || layerW != lw || layerH != lh)
            {
                if (layer != null) layer.Dispose();
                layer = new Bitmap(lw, lh, PixelFormat.Format32bppArgb);
                layerW = lw; layerH = lh;
            }
            layerSc = Sc; layerVx = Vx; layerVy = Vy; layerKey = lk;

            using (Graphics g = Graphics.FromImage(layer))
                RenderScene(g);
        }

        private float LX(double ix) { return (float)((ix * Ppx - Vx) * Sc + Width / 2.0 + MARGIN); }
        private float LY(double iy) { return (float)((iy * Ppy - Vy) * Sc + Height / 2.0 + MARGIN); }
        private double IX(float lx) { return ((lx - (Width / 2.0 + MARGIN)) / Sc + Vx) / Ppx; }
        private double IY(float ly) { return ((ly - (Height / 2.0 + MARGIN)) / Sc + Vy) / Ppy; }

        private double SpotR { get { return Math.Max(1.7, Math.Min(6, 3.1 * Sc)); } }
        private double PinR { get { return Sc > 1 ? 4.4 : 3.0; } }
        private double BigR { get { return Sc > 1 ? 4.6 : 3.2; } }

        private void RenderScene(Graphics g)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.Clear(CBack);

            Image back = Background();
            WorldMeta m = S.Meta;
            if (back != null && m != null)
            {
				var box = new RectangleF(LX(0), LY(0), (float)(PicW * Sc), (float)(PicH * Sc));
                g.DrawImage(back, box);
                using (Pen p = new Pen(Color.FromArgb(41, 90, 176, 255), 1f))
                g.DrawRectangle(p, box.X, box.Y, box.Width, box.Height);
            }

            if (S.ShowPl) DrawRegions(g);

            if (S.ShowSp)
            {
                if (Sc >= 0.11) DrawMobs(g);
                DrawTreasures(g);
                DrawStatues(g);
                DrawNpcs(g);
                DrawSpots(g);
            }
        }

        private void DrawRegions(Graphics g)
        {
            World w = S.W;
            if (w == null) return;
            float fs = (float)Math.Max(9.5, Math.Min(17, 9 + Sc * 2.4));
            if (Sc < 0.09 && fs < 12) return;
            using (Font f = new Font("Segoe UI", fs, FontStyle.Bold, GraphicsUnit.Pixel))
            using (SolidBrush stroke = new SolidBrush(Color.FromArgb(209, 0, 0, 0)))
            using (SolidBrush fill = new SolidBrush(Color.FromArgb(240, 226, 238, 250)))
            using (GraphicsPath path = new GraphicsPath())
            {
                for (int i = 0; i < w.Regions.Count; i++)
                {
                    Region r = w.Regions[i];
                    if (!r.Big && Sc < 0.16) continue;
                    string t = (r.Title ?? "").Replace("\r", "");
                    if (t.Length == 0) continue;
                    float cx = LX(r.X), cy = LY(r.Y);
                    SizeF sz = g.MeasureString(t, f);
                    path.Reset();
                    path.AddRectangle(new RectangleF(cx - sz.Width / 2, cy - sz.Height / 2, sz.Width, sz.Height));
                    g.FillRectangle(stroke, cx - sz.Width / 2 - 1, cy - sz.Height / 2 - 1, sz.Width + 2, sz.Height + 2);
                    g.DrawString(t, f, fill, cx - sz.Width / 2, cy - sz.Height / 2);
                }
            }
        }

        private struct Slot
        {
            public float X, Y, Size;
            public Bitmap Rim;
        }

        private struct Rect2
        {
            public float X, Y, Side;
            public Color C;
        }

        private class Batch
        {
            public List<Slot> slots = new List<Slot>();
            public List<Rect2> rects = new List<Rect2>();

            public void Add(Bitmap r, float cx, float cy, int n)
            {
                Slot s = new Slot();
                s.X = cx; s.Y = cy; s.Size = n; s.Rim = r;
                slots.Add(s);
            }

            public void AddRect(float x, float y, float side, Color c)
            {
                Rect2 r = new Rect2();
                r.X = x; r.Y = y; r.Side = side; r.C = c;
                rects.Add(r);
            }

            public void Flush(Graphics g)
            {
                for (int i = 0; i < slots.Count; i++)
                {
                    Slot s = slots[i];
                    if (s.Rim == null) continue;
                    try
                    {
                        float rs = s.Rim.Width;
                        g.DrawImage(s.Rim, new RectangleF(s.X - rs / 2f, s.Y - rs / 2f, rs, rs));
                    }
                    catch { }
                }
                for (int i = 0; i < rects.Count; i++)
                {
                    Rect2 r = rects[i];
                    using (SolidBrush b = new SolidBrush(r.C))
                        g.FillRectangle(b, r.X, r.Y, r.Side, r.Side);
                }
            }
        }

        private int IconSize(int glyph)
        {
            double v;
            switch (glyph)
            {
                case 0: v = Math.Max(12, Math.Min(52, 6 + Sc * 16)); break;
                case 1: v = Math.Max(11, Math.Min(50, 5 + Sc * 16)); break;
                case 2: v = Math.Max(15, Math.Min(64, 7 + Sc * 21)); break;
                case 3: v = Math.Max(19, Math.Min(78, 10 + Sc * 25)); break;
                default: v = Math.Max(31, Math.Min(126, (14 + Sc * 35) * 1.157)); break;
            }
            if (glyph == 4) return Math.Min(72, Icons.Round4(v));
            return Math.Min(48, Icons.Round4(v));
        }

        private static readonly Color RimMob = Color.FromArgb(255, 0x33, 0x33, 0x33);
        private static readonly Color RimBoss = Color.FromArgb(255, 0xff, 0x5a, 0x4d);
        private static readonly Color RimItem = Color.FromArgb(255, 0xff, 0xdc, 0x4d);
        private static readonly Color RimOre = Color.FromArgb(255, 0x8a, 0x90, 0x99);
        private static readonly Color RimPlant = Color.FromArgb(255, 0x5a, 0x4b, 0x70);
        private static readonly Color RimTre = Color.FromArgb(255, 0xd4, 0x95, 0x3a);
        private static readonly Color RimStat = Color.FromArgb(255, 0x4d, 0x9d, 0xe0);

        private int BuffIconIndex(double x, double y, int n)
        {
            long a = (long)Math.Round(x) * 92837111L;
            long b = (long)Math.Round(y) * 689287499L;
            uint h = (uint)((a ^ b) & 0xFFFFFFFFL);
            int idx = (int)(h % (uint)n);
            idx = ((idx + BuffRotT) % n + n) % n;
            return idx;
        }

        private void DrawMobs(Graphics g)
        {
            World w = S.W;
            if (w == null) return;
            bool far = Sc < 0.5;
            double ds = Math.Max(2.5, Math.Min(6, Sc * 10));

            Batch buffs = new Batch(), mobs = new Batch(), bosses = new Batch(), items = new Batch();
            string[] buffIcons = w.BuffIcons();
            int nBuff = buffIcons.Length;

            for (int i = 0; i < w.Mobs.Count; i++)
            {
                Pt p = w.Mobs[i];
                bool boss = w.Bosses.Contains(p.Ref);
                string name = Str(w.CreatureNames, p.Ref);
                if (!S.Shown(K.C + name)) continue;
                float cx = LX(p.X), cy = LY(p.Y);
                if (boss)
                {
                    if (far) { bosses.AddRect(cx - (float)ds / 2, cy - (float)ds / 2, (float)ds, RimBoss); continue; }
                    string ic = Str(w.CreatureIcons, p.Ref);
                    bool square = string.IsNullOrEmpty(ic);
                    if (square) ic = Str(w.CreatureSquare, p.Ref);
                    int n = IconSize(2);
                    if (square) n = Math.Max(8, Icons.Round4(n * 0.85));
                    Bitmap rim = string.IsNullOrEmpty(ic) ? null : Icons.Rim(S.D, ic, n, RimBoss);
                    if (rim != null) bosses.Add(rim, cx, cy, n);
                    else Fallback(g, cx, cy, PinR, Color.FromArgb(242, 0xff, 0x6e, 0xa8), true);
                }
                else
                {
                    if (far) { mobs.AddRect(cx - (float)ds / 2, cy - (float)ds / 2, (float)ds, RimMob); continue; }
                    string ic = Str(w.CreatureIcons, p.Ref);
                    bool square = string.IsNullOrEmpty(ic);
                    if (square) ic = Str(w.CreatureSquare, p.Ref);
                    int n = IconSize(1);
                    if (square) n = Math.Max(8, Icons.Round4(n * 0.85));
                    Bitmap rim = string.IsNullOrEmpty(ic) ? null : Icons.Rim(S.D, ic, n, RimMob);
                    if (rim != null) mobs.Add(rim, cx, cy, n);
                    else Fallback(g, cx, cy, PinR, Color.FromArgb(200, 0xff, 0x3b, 0x30), false);
                }
            }

            HashSet<string> drawn = new HashSet<string>();
            for (int i = 0; i < w.Items.Count; i++)
            {
                Pt p = w.Items[i];
                string cat = Str(w.ItemCat, p.Ref);
                if (string.IsNullOrEmpty(cat)) cat = "quest";
                bool isBuff = cat == "buff";
                if (!S.Shown(K.I + Str(w.ItemNames, p.Ref))) continue;
                if (!drawn.Add(p.X.ToString("R", CultureInfo.InvariantCulture) + "," + p.Y.ToString("R", CultureInfo.InvariantCulture))) continue;
                float cx = LX(p.X), cy = LY(p.Y);
                Batch ib = isBuff ? buffs : items;
                string ic = Str(w.ItemIcons, p.Ref);
                if (isBuff && nBuff >= 2) ic = buffIcons[BuffIconIndex(p.X, p.Y, nBuff)];
                if (far) { ib.AddRect(cx - (float)ds / 2, cy - (float)ds / 2, (float)ds, RimItem); continue; }
                Bitmap rim = string.IsNullOrEmpty(ic) ? null : Icons.Rim(S.D, ic, IconSize(0), RimItem);
                if (rim != null) ib.Add(rim, cx, cy, IconSize(0));
                else Fallback(g, cx, cy, PinR * 1.45f, Color.FromArgb(245, 0xff, 0xb6, 0x4a), false);
            }

            buffs.Flush(g);
            mobs.Flush(g);
            bosses.Flush(g);
            items.Flush(g);
        }

        private void DrawSpots(Graphics g)
        {
            World w = S.W;
            if (w == null) return;
            bool far = Sc < 0.5;
            double ds = Math.Max(2.5, Math.Min(6, Sc * 10));
            double r = SpotR;
            Batch ore = new Batch(), plant = new Batch();
            using (SolidBrush halo = new SolidBrush(Color.FromArgb(153, 0, 0, 0)))
            {
                for (int i = 0; i < w.Spots.Count; i++)
                {
                    Pt p = w.Spots[i];
                    int gi = p.Ref;
                    if (gi < 0 || gi >= S.D.Species.Count) continue;
                    Species s = S.D.Species[gi];
                    if (!S.Shown(s.Name)) continue;
                    float cx = LX(p.X), cy = LY(p.Y);
                    if (far)
                    {
                        (s.Kind == "ore" ? ore : plant).AddRect(cx - (float)ds / 2, cy - (float)ds / 2, (float)ds,
                            s.Kind == "ore" ? RimOre : RimPlant);
                        continue;
                    }
                    int n = Icons.Round4(Math.Max(6, Math.Min(26, 2.2 + Sc * 8)));
                    Batch b = s.Kind == "ore" ? ore : plant;
                    Color rc = s.Kind == "ore" ? RimOre : RimPlant;
                    float hr = (float)(r + 1.15);
                    Bitmap pic = string.IsNullOrEmpty(s.Icon) ? null : Icons.Rim(S.D, s.Icon, n, rc);
                    if (pic != null)
                    {
                        b.Add(pic, cx, cy, n);
                    }
                    else if (s.Kind == "ore")
                    {
                        g.FillEllipse(halo, cx - hr, cy - hr, hr * 2, hr * 2);
                        using (SolidBrush br = new SolidBrush(Color.FromArgb(230, s.Color)))
                            g.FillEllipse(br, cx - (float)r, cy - (float)r, (float)r * 2, (float)r * 2);
                    }
                    else
                    {
                        g.FillEllipse(halo, cx - hr, cy - hr, hr * 2, hr * 2);
                        Diamond(g, cx, cy, (float)(r * 1.3f), (float)(r * 1.05f), s.Color);
                    }
                }
            }
            ore.Flush(g);
            plant.Flush(g);
        }

        private void DrawTreasures(Graphics g)
        {
            World w = S.W;
            if (w == null) return;
            bool far = Sc < 0.5;
            double ds = Math.Max(2.5, Math.Min(6, Sc * 10));
            Batch b = new Batch();
            for (int i = 0; i < w.Treasures.Count; i++)
            {
                Pt p = w.Treasures[i];
                if (!S.Shown(K.T + Str(w.TreasureNames, p.Ref))) continue;
                float cx = LX(p.X), cy = LY(p.Y);
                if (far) { b.AddRect(cx - (float)ds / 2, cy - (float)ds / 2, (float)ds, RimTre); continue; }
                string ic = Str(w.TreasureIcons, p.Ref);
                if (far) { b.AddRect(cx - (float)ds / 2, cy - (float)ds / 2, (float)ds, RimTre); continue; }
                int n = IconSize(3);
                Bitmap rim = string.IsNullOrEmpty(ic) ? null : Icons.Rim(S.D, ic, n, RimTre);
                if (rim != null) b.Add(rim, cx, cy, n);
                else Diamond(g, cx, cy, (float)(PinR * 1.45f), (float)(PinR * 1.45f), ColorTranslator.FromHtml("#d4953a"));
            }
            b.Flush(g);
        }

        private void DrawStatues(Graphics g)
        {
            World w = S.W;
            if (w == null) return;
            bool far = Sc < 0.5;
            double ds = Math.Max(2.5, Math.Min(6, Sc * 10));
            Batch b = new Batch();
            string ic = "shared/statues/statue.png";
            int n = IconSize(4);
            Bitmap rim = far ? null : Icons.Rim(S.D, ic, n, RimStat);
            for (int i = 0; i < w.Statues.Count; i++)
            {
                Pt p = w.Statues[i];
                if (!S.Shown(K.S + Str(w.StatueNames, p.Ref))) continue;
                float cx = LX(p.X), cy = LY(p.Y);
                if (far) { b.AddRect(cx - (float)ds / 2, cy - (float)ds / 2, (float)ds, RimStat); continue; }
                if (rim != null) b.Add(rim, cx, cy, n);
                else Diamond(g, cx, cy, (float)(BigR * 1.45f), (float)(BigR * 1.45f), ColorTranslator.FromHtml("#4d9de0"));
            }
            b.Flush(g);
        }

        private void DrawNpcs(Graphics g)
        {
            World w = S.W;
            if (w == null) return;
            bool far = Sc < 0.5;
            double ds = Math.Max(2.5, Math.Min(6, Sc * 10));
            Color npcCol = ColorTranslator.FromHtml("#3fd16a");
            Color traderCol = ColorTranslator.FromHtml("#ffc857");
            Batch b = new Batch();
            for (int i = 0; i < w.Npcs.Count; i++)
            {
                Pt p = w.Npcs[i];
                if (!S.Shown(K.N + Str(w.NpcNames, p.Ref))) continue;
                Color c = p.Ref < w.NpcShops.Length && w.NpcShops[p.Ref] != null ? traderCol : npcCol;
                float cx = LX(p.X), cy = LY(p.Y);
                if (far) { b.AddRect(cx - (float)ds / 2, cy - (float)ds / 2, (float)ds, c); continue; }
                Diamond(g, cx, cy, (float)(BigR * 1.45f), (float)(BigR * 1.45f), c);
            }
            b.Flush(g);
        }

        private void Fallback(Graphics g, float cx, float cy, double r, Color fill, bool ring)
        {
            if (ring)
            {
                using (Pen p = new Pen(fill, 1.6f))
                    g.DrawEllipse(p, cx - (float)r * 2.1f, cy - (float)r * 2.1f, (float)r * 4.2f, (float)r * 4.2f);
            }
            using (SolidBrush b = new SolidBrush(fill))
                g.FillEllipse(b, cx - (float)r, cy - (float)r, (float)r * 2, (float)r * 2);
        }

        private void Diamond(Graphics g, float cx, float cy, float ry, float rx, Color c)
        {
            PointF[] pts = new PointF[] {
                new PointF(cx, cy - ry), new PointF(cx + rx, cy),
                new PointF(cx, cy + ry), new PointF(cx - rx, cy) };
            using (SolidBrush b = new SolidBrush(Color.FromArgb(245, c)))
                g.FillPolygon(b, pts);
        }

        private void UpdateBuffTimer()
        {
            World w = S.W;
            bool want = w != null && S.ShowSp && w.BuffIcons().Length >= 2;
            if (want && buffTimer == null)
            {
                buffTimer = new Timer();
                buffTimer.Interval = 2400;
                buffTimer.Tick += delegate (object s, EventArgs e)
                {
                    BuffRotT++;
                    InvalidateLayer();
                    CancelPreview();
                    Invalidate();
                };
            }
            if (want) buffTimer.Start();
            else if (buffTimer != null) buffTimer.Stop();
        }

        public void RefreshFilters()
        {
            UpdateBuffTimer();
            InvalidateLayer();
            CancelPreview();
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (S == null || S.D == null || S.Meta == null) return;
            if (settling && layer != null && layerSc > 0)
            {
                float k = (float)(Sc / layerSc);
                float dx = (float)((Vx - layerVx) * Sc);
                float dy = (float)((Vy - layerVy) * Sc);
                bool usable = Math.Abs(dx) <= Width && Math.Abs(dy) <= Height
                           && k > 0.3f && k < 3.3f;
                if (usable)
                {
                    var dest = new RectangleF(
                        (Width / 2f) - (Width / 2f) * k - dx,
                        (Height / 2f) - (Height / 2f) * k - dy,
                        Width * k, Height * k);
                    e.Graphics.DrawImage(layer, dest,
                        new Rectangle(MARGIN, MARGIN, Width, Height), GraphicsUnit.Pixel);
                    DrawScale(e.Graphics);
                    DrawCursorPos(e.Graphics);
                    return;
                }
            }
            EnsureLayer();
            if (layer != null)
                e.Graphics.DrawImage(layer, new Rectangle(0, 0, Width, Height),
                                     new Rectangle(MARGIN, MARGIN, Width, Height), GraphicsUnit.Pixel);
            DrawScale(e.Graphics);
            DrawCursorPos(e.Graphics);
        }

        private void DrawScale(Graphics g)
        {
            double pxm = Sc * S.D.K * Ppx;
            if (pxm <= 0) return;
            double[] steps = { 1, 2, 5, 10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000 };
            double pick = steps[0];
            for (int i = 0; i < steps.Length; i++) { if (steps[i] * pxm <= 150) pick = steps[i]; }
            double len = pick * pxm;
            if (len < 8) return;
            int x0 = Width - (int)len - 18, y0 = Height - 16;
            using (Pen p = new Pen(Color.FromArgb(230, 231, 238, 247), 2f))
            {
                g.DrawLine(p, x0, y0, x0 + (int)len, y0);
                g.DrawLine(p, x0, y0 - 4, x0, y0 + 4);
                g.DrawLine(p, x0 + (int)len, y0 - 4, x0 + (int)len, y0 + 4);
            }
            string lab = pick.ToString(CultureInfo.InvariantCulture) + " m";
            using (Font f = new Font("Segoe UI", 8.5f))
            using (SolidBrush b = new SolidBrush(CFg))
                g.DrawString(lab, f, b, x0, y0 - 20);
        }

        private void DrawCursorPos(Graphics g)
        {
            if (!CursorVisible) return;
            double wx, wz;
            ToWorld(IX(MouseX + MARGIN), IY(MouseY + MARGIN), out wx, out wz);
            string s = string.Format(CultureInfo.InvariantCulture, "x {0:0}   z {1:0}", wx, wz);
            using (Font f = new Font("Consolas", 8.5f))
            using (SolidBrush b = new SolidBrush(CMut))
                g.DrawString(s, f, b, 12, Height - 20);
        }

        public bool CursorVisible = true;

        public Hit Pick(int cx, int cy)
        {
            World w = S.W;
            if (w == null) return null;
            double lx = cx + MARGIN, ly = cy + MARGIN;

            Hit best = null;
            double bestScore = double.MaxValue;
            Nearest(w.Spots, lx, ly, "species", ref best, ref bestScore);
            Nearest(w.Treasures, lx, ly, "treasure", ref best, ref bestScore);
            Nearest(w.Statues, lx, ly, "statue", ref best, ref bestScore);
            Nearest(w.Npcs, lx, ly, "npc", ref best, ref bestScore);
            Nearest(w.Items, lx, ly, "item", ref best, ref bestScore);
            if (Sc >= 0.11) Nearest(w.Mobs, lx, ly, "mob", ref best, ref bestScore);
            return best;
        }

        private double HitRadius(string kind, int refI)
        {
            const double Pad = 3;
            if (Sc < 0.5)
            {
                double ds = Math.Max(2.5, Math.Min(6, Sc * 10));
                return Math.Max(7.5, ds / 2.0 + 3);
            }
            if (kind == "species")
            {
                int n = Icons.Round4(Math.Max(6, Math.Min(26, 2.2 + Sc * 8)));
                return Math.Max(7.5, n / 2.0 + Pad);
            }
            if (kind == "treasure") return IconSize(3) / 2.0 + Pad;
            if (kind == "statue") return IconSize(4) / 2.0 + Pad;
            if (kind == "npc") return Math.Max(7, BigR * 1.45 + Pad);
            if (kind == "item") return Math.Max(7, IconSize(0) / 2.0 + Pad);
            World w = S.W;
            bool boss = w != null && w.Bosses.Contains(refI);
            return Math.Max(6.5, IconSize(boss ? 2 : 1) / 2.0 + Pad);
        }

        private void Nearest(List<Pt> pts, double lx, double ly, string kind, ref Hit best, ref double bestScore)
        {
            World w = S.W;
            for (int i = 0; i < pts.Count; i++)
            {
                Pt p = pts[i];
                float x = LX(p.X), y = LY(p.Y);
                double dx = x - lx, dy = y - ly;
                double d2 = dx * dx + dy * dy;
                double radius = HitRadius(kind, p.Ref);
                if (radius <= 0) continue;
                double score = d2 / (radius * radius);
                if (score > 1 || score >= bestScore) continue;
                string name;
                string key;
                if (kind == "species")
                {
                    int gi = p.Ref;
                    if (gi < 0 || gi >= S.D.Species.Count) continue;
                    Species s = S.D.Species[gi];
                    name = s.Name; key = s.Name;
                    if (!S.Shown(key)) continue;
                }
                else if (kind == "treasure") { name = Str(w.TreasureNames, p.Ref); key = K.T + name; if (!S.Shown(key)) continue; }
                else if (kind == "statue") { name = Str(w.StatueNames, p.Ref); key = K.S + name; if (!S.Shown(key)) continue; }
                else if (kind == "npc") { name = Str(w.NpcNames, p.Ref); key = K.N + name; if (!S.Shown(key)) continue; }
                else if (kind == "item") { name = Str(w.ItemNames, p.Ref); key = K.I + name; if (!S.Shown(key)) continue; }
                else { name = Str(w.CreatureNames, p.Ref); key = K.C + name; if (!S.Shown(key)) continue; }

                double wx, wz;
                ToWorld(p.X, p.Y, out wx, out wz);
                Hit h = new Hit();
                h.Kind = kind; h.Label = name; h.X = p.X; h.Y = p.Y; h.WX = wx; h.WZ = wz;
                if (kind == "species")
                {
                    Species s = S.D.Species[p.Ref];
                    h.Zone = string.IsNullOrEmpty(s.Zone) ? null : s.Zone;
                    h.Cat = s.Kind;
                    h.Icon = s.Icon;
                    h.Color = s.Color;
                }
                else if (kind == "treasure")
                {
                    h.Full = Str(w.TreasureFull, p.Ref);
                    h.Icon = Str(w.TreasureIcons, p.Ref);
                    h.Text = Str(w.TreasureText, p.Ref);
                    h.Zone = Str(w.TreasureZone, p.Ref);
                    h.Label = name;
                }
                else if (kind == "statue")
                {
                    h.Label = Str(w.StatueLabel, p.Ref);
                    if (string.IsNullOrEmpty(h.Label)) h.Label = name;
                    h.Zone = Str(w.StatueZone, p.Ref);
                    h.Icon = "shared/statues/statue.png";
                }
                else if (kind == "npc")
                {
                    string shown = Str(w.NpcLabel, p.Ref);
                    if (!string.IsNullOrEmpty(shown)) h.Label = shown;
                    h.Zone = State.NpcArea(w, p, p.Ref);
                    NpcShop shop = p.Ref < w.NpcShops.Length ? w.NpcShops[p.Ref] : null;
                    h.Shop = shop;
                    if (shop != null)
                    {
                        h.Color = ColorTranslator.FromHtml("#ffc857");
                        h.Alt = shop.Kind;
                    }
                    else h.Color = ColorTranslator.FromHtml("#3fd16a");
                }
                else if (kind == "item")
                {
                    h.Zone = Str(w.ItemZone, p.Ref);
                    h.Alt = Str(w.ItemAlt, p.Ref);
                    h.Icon = Str(w.ItemIcons, p.Ref);
                    h.Color = ColorTranslator.FromHtml("#ffb64a");
                }
                else
                {
                    h.Zone = Str(w.CreatureZone, p.Ref);
                    h.Lvl = Str(w.CreatureLevel, p.Ref);
                    h.Alt = Str(w.CreatureAlt, p.Ref);
                    h.Icon = Str(w.CreatureIcons, p.Ref);
                    if (string.IsNullOrEmpty(h.Icon)) h.Icon = Str(w.CreatureSquare, p.Ref);
                    h.Boss = w.Bosses.Contains(p.Ref);
                    h.Color = h.Boss ? ColorTranslator.FromHtml("#ff6ea8") : ColorTranslator.FromHtml("#ff3b30");
                }
                best = h; bestScore = score;
            }
        }

        public static string Str(string[] a, int i)
        {
            return (a != null && i >= 0 && i < a.Length) ? a[i] : null;
        }

        public int MouseX, MouseY;

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            MouseX = e.X; MouseY = e.Y;
            if (e.Button == MouseButtons.Left)
            {
                downAt = new Point(e.X, e.Y);
                lastAt = downAt;
                dragging = true; dragMoved = false;
                Cursor = Cursors.SizeAll;
                Capture = true;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            MouseX = e.X; MouseY = e.Y;
            if (dragging)
            {
                int dx = e.X - lastAt.X, dy = e.Y - lastAt.Y;
                if (dx != 0 || dy != 0)
                {
                    Vx -= dx / Sc; Vy -= dy / Sc;
                    lastAt = new Point(e.X, e.Y);
                    if (Math.Abs(e.X - downAt.X) + Math.Abs(e.Y - downAt.Y) > 3) dragMoved = true;
                    ClampView();
                    InvalidateLayerAfterMotion();
                    Invalidate();
                    OnViewChanged();
                }
                return;
            }
            Hit h = Pick(e.X, e.Y);
            bool same = (h == null && hover == null) ||
                        (h != null && hover != null && h.Kind == hover.Kind && h.Label == hover.Label &&
                         Math.Abs(h.X - hover.X) < 0.001 && Math.Abs(h.Y - hover.Y) < 0.001);
            if (!same)
            {
                hover = h;
                if (HoverChanged != null) HoverChanged(this, EventArgs.Empty);
            }
            Cursor = h != null ? Cursors.Hand : Cursors.Hand;
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left) return;
            dragging = false;
            Capture = false;
            Cursor = Cursors.Hand;
            if (dragMoved) return;
            Hit h = Pick(e.X, e.Y);
            if (Picked != null) Picked(this, new AtlasPickEventArgs(h));
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (hover != null)
            {
                hover = null;
                if (HoverChanged != null) HoverChanged(this, EventArgs.Empty);
            }
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            ZoomAt(e.X, e.Y, e.Delta < 0 ? 1 / 1.18 : 1.18);
        }

        protected override bool IsInputKey(Keys keyData)
        {
            return keyData == Keys.Left || keyData == Keys.Right || keyData == Keys.Up ||
                   keyData == Keys.Down || base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            Keys k = e.KeyCode;
            if (k == Keys.F) { Fit(); e.Handled = true; }
            else if (k == Keys.Add || k == Keys.Oemplus || k == Keys.D1) { ZoomAt(Width / 2.0, Height / 2.0, 1.4); e.Handled = true; }
            else if (k == Keys.Subtract || k == Keys.OemMinus || k == Keys.D0) { ZoomAt(Width / 2.0, Height / 2.0, 1 / 1.4); e.Handled = true; }
            else if (k == Keys.Left) { Vx -= 90 / Sc; ClampView(); InvalidateLayerAfterMotion(); Invalidate(); e.Handled = true; }
            else if (k == Keys.Right) { Vx += 90 / Sc; ClampView(); InvalidateLayerAfterMotion(); Invalidate(); e.Handled = true; }
            else if (k == Keys.Up) { Vy -= 90 / Sc; ClampView(); InvalidateLayerAfterMotion(); Invalidate(); e.Handled = true; }
            else if (k == Keys.Down) { Vy += 90 / Sc; ClampView(); InvalidateLayerAfterMotion(); Invalidate(); e.Handled = true; }
            else if (k == Keys.D1 || (k >= Keys.D1 && k <= Keys.D9))
            {
                int idx = k - Keys.D1;
                if (S != null && S.D != null && idx < S.D.Maps.Count)
                {
                    WorldChange(this, new AtlasPickEventArgs(S.D.Maps[idx].Key));
                    e.Handled = true;
                }
            }
        }

        public event EventHandler<AtlasPickEventArgs> WorldChange;
    }

    public class AtlasPickEventArgs : EventArgs
    {
        public Hit Hit;
        public string Key;
        public string Group;
        public string Zone;
        public int More;
        public AtlasPickEventArgs(Hit h) { Hit = h; }
        public AtlasPickEventArgs(string key) { Key = key; }
    }

    #endregion

    #region Sidebar

    public sealed class Sidebar : Control
    {
        public State S;

        private static Font KeepAlive(Font f)
        {
            GC.SuppressFinalize(f);
            return f;
        }

        private static readonly Font FTag = KeepAlive(new Font("Segoe UI", 7f));
        private static readonly Font FZone = KeepAlive(new Font("Segoe UI", 8f));
        private static readonly Font FAlt = KeepAlive(new Font("Segoe UI", 8.2f, FontStyle.Italic));
        private static readonly Font FRow = KeepAlive(new Font("Segoe UI", 8.7f));
        private static readonly Font FHead = KeepAlive(new Font("Segoe UI", 8.6f, FontStyle.Regular, GraphicsUnit.Point));
        private static readonly Font FHeadBold = KeepAlive(new Font("Segoe UI", 8.6f, FontStyle.Bold, GraphicsUnit.Point));
        private static readonly Font FMono = KeepAlive(new Font("Consolas", 7.6f));

        public event EventHandler<AtlasPickEventArgs> RowActivated;
        public event EventHandler<AtlasPickEventArgs> ZoneActivated;
        public event EventHandler<AtlasPickEventArgs> WorldActivated;
        public event EventHandler FilterChanged;
        public event EventHandler FoldsChanged;

        private const int RH = 22;
        private const int HEAD = 24;
        private const int XROW = 21;

        public const int PAD_L = 8;
        public const int INDENT = 16;
        public const int ARROW_W = 12;
        public const int ARROW_HIT = 12;
        public const int BOX = 14;
        public const int BOX_HEAD = 13;
        public const int BOX_GAP = 7;
        public const int ICON = 22;
        public const int ICON_SMALL = 16;
        public const int ICON_GAP = 7;
        public const int RAIL = 9;
        public const int RAIL_GAP = 10;
        public const int TEXT_H = 13;

        public const int K_GROUP = 0, K_CAT = 1, K_ZONE = 2, K_ROW = 3,
                         K_XGRP = 4, K_XWORLD = 5, K_XROW = 6, K_EMPTY = 7;

        public sealed class Item
        {
            public int Kind;
            public string Key;
            public string Label;
            public string Group;
            public string Zone;
            public string Lvl, Alt, Icon, BTag, Tag;
            public Color Color;
            public int N, Total, Indent;
            public bool Tri;
            public bool Stray;
            public int More;
        }

        public List<Item> Items = new List<Item>();
        private Dictionary<string, bool> gOpen = new Dictionary<string, bool>();
        private Dictionary<string, bool> hOpen = new Dictionary<string, bool>();
        private Dictionary<string, bool> xOpen = new Dictionary<string, bool>();
        private Dictionary<string, Dictionary<string, bool>> hits = new Dictionary<string, Dictionary<string, bool>>();
        private int contentH;
        private int scroll;
        private int hoverIdx = -1;
        private bool hoverRail;
        private bool draggingThumb;
        private int thumbGrab;

        private Color CPanel, CPanel2, CLine, CFg, CMut, CDim, CAcc;

        public Sidebar()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
            CPanel = Color.FromArgb(17, 24, 35);
            CPanel2 = Color.FromArgb(24, 33, 46);
            CLine = Color.FromArgb(38, 51, 69);
            CFg = Color.FromArgb(231, 238, 247);
            CMut = Color.FromArgb(132, 150, 172);
            CDim = Color.FromArgb(93, 110, 131);
            CAcc = Color.FromArgb(90, 176, 255);
            BackColor = CPanel;
        }

        public void SetState(State s)
        {
            S = s;
            Relayout();
        }

        public void SetPalette(Color panel, Color panel2, Color line, Color fg, Color mut, Color dim, Color acc)
        {
            CPanel = panel; CPanel2 = panel2; CLine = line;
            CFg = fg; CMut = mut; CDim = dim; CAcc = acc;
            BackColor = panel;
            Invalidate();
        }

        public void Relayout()
        {
            Items = new List<Item>();
            hits = new Dictionary<string, Dictionary<string, bool>>();
            contentH = 0;
            if (S == null) { Invalidate(); return; }

            List<KeyValuePair<string, List<XRow>>> xw = S.OtherWorlds(8);
            if (xw.Count > 0)
            {
                int total = 0;
                for (int i = 0; i < xw.Count; i++) total += xw[i].Value.Count;
                Item g = new Item();
                g.Kind = K_XGRP; g.Key = "Other worlds"; g.Label = "Other worlds";
                g.N = total; Items.Add(g); contentH += HEAD;
                if (IsOpen(xOpen, "Other worlds"))
                {
                    for (int i = 0; i < xw.Count; i++)
                    {
                        string mk = xw[i].Key;
                        WorldMeta wm = S.Find(mk);
                        Item wh = new Item();
                        wh.Kind = K_XWORLD; wh.Key = mk; wh.Label = wm == null ? mk : wm.Title;
                        wh.N = xw[i].Value.Count; Items.Add(wh); contentH += HEAD;
                        if (IsOpen(xOpen, mk))
                        {
                            int lim = Math.Min(8, xw[i].Value.Count);
                            for (int r = 0; r < lim; r++)
                            {
                                XRow xr = xw[i].Value[r];
                                Item it = new Item();
                                it.Kind = K_XROW; it.Key = mk; it.Label = xr.Label;
                                it.Group = xr.Group; it.Zone = xr.Zone; it.Lvl = xr.Lvl;
                                Items.Add(it); contentH += XROW;
                            }
                            if (xw[i].Value.Count > lim)
                            {
                                Item m = new Item();
                                m.Kind = K_EMPTY;
                                m.Key = mk;
                                m.More = xw[i].Value.Count - lim;
                                m.Label = "... and " + m.More + " more on this world";
                                Items.Add(m); contentH += XROW;
                            }
                        }
                    }
                }
            }

            for (int gi = 0; gi < S.Groups.Count; gi++)
            {
                Group g = S.Groups[gi];
                Dictionary<string, bool> hit = new Dictionary<string, bool>();
                S.HitMap(g, hit);
                hits[g.Title] = hit;

                int vis = 0, visN = 0;
                for (int i = 0; i < g.Rows.Count; i++)
                {
                    if (hit[g.Rows[i].Key]) { vis++; visN += g.Rows[i].N; }
                }

                Item head = new Item();
                head.Kind = K_GROUP; head.Key = g.Title; head.Label = g.Title;
                head.N = vis; head.Total = visN;
                head.Tri = TriOf(g.Rows, hit);
                head.BTag = vis == 1 ? g.UnitSing : g.UnitPlur;
                Items.Add(head); contentH += HEAD;

                if (!IsOpen(gOpen, g.Title)) continue;

                bool inSet = false;
                bool setClosed = false;
                for (int b = 0; b < g.Buckets.Count; b++)
                {
                    Bucket bk = g.Buckets[b];
                    string bkey = g.Title + "/" + (bk.Title ?? "~");
                    if (bk.Title != null)
                    {
                        int bv = 0, bvn = 0;
                        for (int r = 0; r < bk.Rows.Count; r++)
                            if (hit[bk.Rows[r].Key]) { bv++; bvn += bk.Rows[r].N; }
                        if (bv == 0)
                        {
                            if (bk.Cls == "cat") setClosed = true;
                            continue;
                        }
                        Item h = new Item();
                        h.Kind = bk.Cls == "cat" ? K_CAT : K_ZONE;
                        h.Key = bkey;
                        h.Label = bk.Cls == "cat" ? State.CatTitle(bk.Title) : bk.Title;
                        h.Group = g.Title;
                        h.N = bv; h.Total = bvn;
                        h.Indent = bk.Cls == "cat" ? 1 : (inSet ? 3 : 2);
                        h.Tri = TriOf(bk.Rows, hit);

                        if (bk.Cls == "cat")
                        {
                            Items.Add(h); contentH += HEAD;
                            setClosed = !IsOpen(hOpen, bkey);
                            inSet = !setClosed;
                            if (setClosed) continue;
                            if (bk.Head) continue;
                        }
                        else
                        {
                            if (setClosed) continue;
                            Items.Add(h); contentH += HEAD;
                            if (!IsOpen(hOpen, bkey)) continue;
                        }
                    }
                    for (int r = 0; r < bk.Rows.Count; r++)
                    {
                        Row rw = bk.Rows[r];
                        if (!hit[rw.Key]) continue;
                        Item it = new Item();
                        it.Kind = K_ROW; it.Key = rw.Key; it.Label = rw.Label;
                        it.Group = g.Title; it.Zone = rw.Zone;
                        it.Lvl = rw.Lvl; it.Alt = rw.Alt; it.Icon = rw.Icon;
                        it.Tag = rw.Tag;
                        it.Color = rw.Color; it.N = rw.N;
                        it.Stray = rw.Stray;
                        it.Indent = bk.Title == null ? 1 : (inSet ? 4 : 3);
                        Items.Add(it); contentH += RH;
                    }
                }
            }

            if (Items.Count == 0)
            {
                Item e = new Item();
                e.Kind = K_EMPTY; e.Label = "No species have been recorded here yet.";
                Items.Add(e); contentH += RH;
            }

            scroll = Math.Max(0, Math.Min(scroll, Math.Max(0, contentH - Height)));
            Invalidate();
        }

        private static bool IsOpen(Dictionary<string, bool> d, string k)
        {
            if (k == null) return true;
            bool v;
            return !d.TryGetValue(k, out v) || v;
        }

        private bool TriOf(List<Row> rows, Dictionary<string, bool> hit)
        {
            int on = 0, tot = 0;
            for (int i = 0; i < rows.Count; i++)
            {
                bool v;
                if (!hit.TryGetValue(rows[i].Key, out v) || !v) continue;
                tot++;
                if (SIsOn(rows[i].Key)) on++;
            }
            if (tot == 0) return false;
            return on == tot;
        }

        private bool SIsOn(string key)
        {
            return S != null && S.IsOn(key);
        }


        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(CPanel);
            if (Items.Count == 0) return;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            int y = -scroll;
            for (int i = 0; i < Items.Count; i++)
            {
                Item it = Items[i];
                int h = HeightOf(it);
                if (y + h >= 0 && y <= Height) PaintItem(g, it, y, i == hoverIdx);
                y += h;
                if (y > Height) break;
            }

            if (contentH > Height)
            {
                int ty, th;
                ScrollGeom(out ty, out th);
                using (SolidBrush track = new SolidBrush(Color.FromArgb(16, CLine)))
                    g.FillRectangle(track, Width - RAIL, 0, RAIL, Height);
                Color thumb = draggingThumb ? CAcc
                            : hoverRail ? Color.FromArgb(120, 140, 172)
                            : Color.FromArgb(49, 65, 90);
                using (SolidBrush b = new SolidBrush(thumb))
                    g.FillRectangle(b, Width - RAIL, ty, RAIL, th);
            }
        }

        private void ScrollGeom(out int top, out int height)
        {
            top = 0;
            height = 0;
            if (contentH <= Height || Height <= 0) return;
            height = (int)Math.Max(30L, (long)Height * Height / Math.Max(1, contentH));
            if (height > Height) height = Height;
            int maxTop = Height - height;
            int maxScroll = Math.Max(1, contentH - Height);
            top = (int)((long)scroll * maxTop / maxScroll);
            if (top < 0) top = 0;
            if (top > maxTop) top = maxTop;
        }

        private int ScrollForTop(int top)
        {
            int railTop, railH;
            ScrollGeom(out railTop, out railH);
            int maxTop = Height - railH;
            if (maxTop <= 0) return 0;
            int maxScroll = Math.Max(0, contentH - Height);
            int v = (int)((long)top * maxScroll / maxTop);
            if (v < 0) v = 0;
            if (v > maxScroll) v = maxScroll;
            return v;
        }

        private bool OverRail(int x)
        {
            return contentH > Height && x >= Width - RAIL;
        }

        private bool OverThumb(int y)
        {
            int ty, th;
            ScrollGeom(out ty, out th);
            return th > 0 && y >= ty && y <= ty + th;
        }

        private int HeightOf(Item it)
        {
            switch (it.Kind)
            {
                case K_ROW: return RH;
                case K_XROW: return XROW;
                case K_EMPTY: return RH;
                default: return HEAD;
            }
        }

        private void PaintItem(Graphics g, Item it, int y, bool hot)
        {
            int leftPad = PAD_L + it.Indent * INDENT;
            bool head = it.Kind == K_GROUP || it.Kind == K_CAT || it.Kind == K_ZONE || it.Kind == K_XGRP || it.Kind == K_XWORLD;

            if (hot && !head)
            {
                using (SolidBrush b = new SolidBrush(CPanel2))
                    g.FillRectangle(b, 0, y, Width - RAIL_GAP, RH);
                using (SolidBrush b = new SolidBrush(CAcc))
                    g.FillRectangle(b, 0, y, 2, RH);
            }
            if (it.Kind == K_CAT)
            {
                using (Pen p = new Pen(Color.FromArgb(55, CLine)))
                    g.DrawLine(p, PAD_L, y, Width - RAIL_GAP, y);
            }

            string text = it.Label;
            string sub = null;

            if (it.Kind == K_GROUP)
                sub = it.N + " " + (it.BTag ?? "") + " \u00B7 " + it.Total;
            else if (it.Kind == K_CAT || it.Kind == K_ZONE || it.Kind == K_XWORLD)
                sub = it.N.ToString(CultureInfo.InvariantCulture);
            else if (it.Kind == K_ROW)
                sub = it.N.ToString(CultureInfo.InvariantCulture);

            int tx = leftPad;
            int arrowW = ARROW_W;

            int tagW = 0;
            if ((it.Kind == K_XROW && !string.IsNullOrEmpty(it.Group)) ||
                (it.Kind == K_ROW && !string.IsNullOrEmpty(it.Tag)))
            {
                Font ft = FTag;
                string tt = it.Kind == K_XROW ? it.Group : it.Tag;
                tagW = (int)g.MeasureString(tt, ft).Width + 9;
            }
            int zoneW = 0;
            if (it.Kind == K_XROW && !string.IsNullOrEmpty(it.Zone))
            {
                Font fz = FZone;
                zoneW = (int)g.MeasureString(it.Zone, fz).Width + 8;
            }

            if (head && (it.Kind == K_GROUP || it.Kind == K_CAT || it.Kind == K_ZONE || it.Kind == K_XGRP || it.Kind == K_XWORLD))
            {
                bool open = it.Kind == K_GROUP ? IsOpen(gOpen, it.Key)
                          : it.Kind == K_XGRP ? IsOpen(xOpen, it.Key)
                          : IsOpen(hOpen, it.Key);
                DrawArrow(g, leftPad, y + 7, open, CMut);
                tx += arrowW;
            }
            else tx += arrowW;

            if (it.Kind != K_XROW && it.Kind != K_EMPTY && it.Kind != K_XGRP && it.Kind != K_XWORLD)
                tx = DrawBox(g, tx, y + (head ? 6 : 5), head ? BOX_HEAD : BOX,
                             head ? it.Tri : SIsOn(it.Key));

            bool rail = contentH > Height;
            int rightEdge = Width - (rail ? RAIL + RAIL_GAP : RAIL_GAP);
            int countW = 0;
            if (sub != null)
            {
                using (Font fc = new Font("Consolas", 7.6f))
                    countW = (int)g.MeasureString(sub, fc).Width + 10;
                int room = rightEdge - leftPad - 30;
                if (countW > room) countW = room;
            }

            if (it.Kind == K_ROW && !string.IsNullOrEmpty(it.Icon))
            {
                int isz = it.Indent >= 4 ? ICON_SMALL : ICON;
                Bitmap b = Icons.Img(S.D, it.Icon);
                if (b != null)
                {
                    float s = (float)isz / (b.Width > b.Height ? b.Width : b.Height);
                    int dw = Math.Max(1, (int)(b.Width * s)), dh = Math.Max(1, (int)(b.Height * s));
                    g.DrawImage(b, tx, y + (RH - dh) / 2, dw, dh);
                    tx += isz + ICON_GAP;
                }
            }
            else if (it.Kind == K_ROW)
            {
                using (SolidBrush b = new SolidBrush(SOn(it.Key) ? it.Color : Fade(it.Color)))
                {
                    if (it.Group == "Items" || it.Group == "Treasure Maps" || it.Group == "Statues" || it.Group == "NPCs" || it.Group == "Traders")
                    {
                        var c = new Point[] { new Point(tx + 7, y + 5), new Point(tx + 15, y + 12), new Point(tx + 7, y + 19), new Point(tx - 1, y + 12) };
                        g.FillPolygon(b, c);
                    }
                    else if (it.Group == "Ore" || it.Group == "Plant")
                        g.FillEllipse(b, tx + 2, y + 6, 11, 11);
                    else
                        g.FillEllipse(b, tx + 2, y + 6, 11, 11);
                }
                if (it.Group == "Bosses")
                {
                    using (Pen p = new Pen(Color.FromArgb(255, 110, 168), 2f))
                        g.DrawEllipse(p, tx + 1, y + 5, 13, 13);
                }
                tx += 18;
            }

            if (tagW > 0)
            {
                Font ft = FTag;
                string tt = it.Kind == K_XROW ? it.Group : it.Tag;
                Color tc = it.Kind == K_ROW ? it.Color : CDim;
                using (Pen p = new Pen(CLine))
                using (SolidBrush b = new SolidBrush(tc))
                {
                    var tr = new RectangleF(tx, y + 4, tagW, 14);
                    g.DrawRectangle(p, tr.X, tr.Y, tr.Width, tr.Height);
                    g.DrawString(tt, ft, b, tr.X + 4, tr.Y + 1);
                }
                tx += tagW + 6;
            }

            int altW = 0;
            if (it.Kind == K_ROW && !string.IsNullOrEmpty(it.Alt))
            {
                Font fa = FAlt;
                altW = (int)g.MeasureString(it.Alt, fa).Width + 6;
            }

            Font fh = head ? (it.Kind == K_CAT ? FHeadBold : FHead) : FRow;
            Font fn = FRow;
            Font fm = FMono;
            {
                Color tc = head ? (it.Kind == K_CAT ? CAcc : CMut) : CFg;
                if (!head && it.Kind == K_ROW && !SIsOn(it.Key)) tc = CDim;

                int lvlW = 0;
                if (it.Kind == K_ROW && !string.IsNullOrEmpty(it.Lvl))
                {
                    string lt = State.LvlText(it.Lvl);
                    if (lt.Length > 0) lvlW = (int)g.MeasureString(lt, fm).Width + 6;
                }

                int avail = rightEdge - countW - lvlW - altW - zoneW - tx;
                if (avail < 10) avail = 10;
                string shown = Ellipsis(g, text, fn, avail);

                int cy = y + (HeightOf(it) - TEXT_H) / 2 - 1;
                if (!SOn(it.Key) && (it.Kind == K_ROW)) { tc = CDim; }
                using (SolidBrush tb = new SolidBrush(tc))
                {
                    if (head) g.DrawString(shown, fh, tb, tx, cy);
                    else g.DrawString(shown, fn, tb, tx, cy);
                }

                if (zoneW > 0)
                {
                    Font fz = FZone;
                    using (SolidBrush b = new SolidBrush(CDim))
                        g.DrawString(Ellipsis(g, it.Zone, fz, zoneW), fz, b, rightEdge - countW - zoneW, cy);
                }

                if (it.Kind == K_ROW && !string.IsNullOrEmpty(it.Alt))
                {
                    Font fa = FAlt;
                    using (SolidBrush b = new SolidBrush(CMut))
                        g.DrawString(Ellipsis(g, it.Alt, fa, altW - 2), fa, b, tx + avail, cy);
                }
                if (lvlW > 0)
                {
                    using (SolidBrush b = new SolidBrush(CMut))
                        g.DrawString(State.LvlText(it.Lvl), fm, b, tx + avail + altW, cy + 1);
                }
                if (sub != null)
                {
                    using (SolidBrush b = new SolidBrush(CDim))
                        g.DrawString(Ellipsis(g, sub, fm, countW - 6), fm, b, rightEdge - countW + 6, cy + 1);
                }
            }
        }

        private bool SOn(string key)
        {
            return S != null && S.IsOn(key);
        }

        private static Color Fade(Color c)
        {
            return Color.FromArgb(70, c);
        }

        private static string Ellipsis(Graphics g, string s, Font f, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (g.MeasureString(s, f).Width <= max) return s;
            while (s.Length > 1 && g.MeasureString(s + "...", f).Width > max) s = s.Substring(0, s.Length - 1);
            return s + "...";
        }

        private void DrawArrow(Graphics g, int x, int y, bool open, Color c)
        {
            using (Pen p = new Pen(c, 1.7f))
            {
                p.StartCap = LineCap.Round;
                p.EndCap = LineCap.Round;
                if (open)
                {
                    g.DrawLine(p, x, y - 1, x + 4, y + 3);
                    g.DrawLine(p, x + 4, y + 3, x + 8, y - 1);
                }
                else
                {
                    g.DrawLine(p, x + 2, y - 3, x + 6, y + 1);
                    g.DrawLine(p, x + 6, y + 1, x + 2, y + 5);
                }
            }
        }

        private int DrawBox(Graphics g, int x, int y, int size, bool on)
        {
            var r = new Rectangle(x, y, size, size);
            using (SolidBrush bg = new SolidBrush(on ? CAcc : Color.FromArgb(13, 18, 26)))
                g.FillRectangle(bg, r);
            using (Pen p = new Pen(on ? CAcc : CLine))
                g.DrawRectangle(p, r);
            if (on)
            {
                using (Pen p = new Pen(Color.FromArgb(10, 14, 19), 1.7f))
                {
                    g.DrawLine(p, x + size * 0.24f, y + size * 0.52f, x + size * 0.43f, y + size * 0.71f);
                    g.DrawLine(p, x + size * 0.43f, y + size * 0.71f, x + size * 0.78f, y + size * 0.30f);
                }
            }
            return x + size + BOX_GAP;
        }

        private int IndexAt(int y)
        {
            int cy = y + scroll;
            for (int i = 0; i < Items.Count; i++)
            {
                int h = HeightOf(Items[i]);
                if (cy < h) return i;
                cy -= h;
            }
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (draggingThumb)
            {
                scroll = ScrollForTop(e.Y - thumbGrab);
                Invalidate();
                return;
            }
            bool rail = OverRail(e.X);
            if (rail != hoverRail) { hoverRail = rail; Invalidate(); }
            if (rail)
            {
                if (hoverIdx >= 0) { hoverIdx = -1; Invalidate(); }
                Cursor = OverThumb(e.Y) ? Cursors.SizeNS : Cursors.Default;
                return;
            }
            int i = IndexAt(e.Y);
            if (i != hoverIdx)
            {
                hoverIdx = i;
                Cursor = i >= 0 ? Cursors.Hand : Cursors.Default;
                Invalidate();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (draggingThumb) return;
            if (hoverIdx >= 0 || hoverRail)
            {
                hoverIdx = -1;
                hoverRail = false;
                Invalidate();
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;

            if (OverRail(e.X))
            {
                int ty, th;
                ScrollGeom(out ty, out th);
                Focus();
                draggingThumb = true;
                if (OverThumb(e.Y))
                {
                    thumbGrab = e.Y - ty;
                }
                else
                {
                    int page = Math.Max(RH, Height - RH);
                    int max = Math.Max(0, contentH - Height);
                    scroll = Math.Max(0, Math.Min(scroll + (e.Y < ty ? -page : page), max));
                    ScrollGeom(out ty, out th);
                    thumbGrab = th / 2;
                }
                Capture = true;
                Cursor = Cursors.SizeNS;
                Invalidate();
                return;
            }

            int i = IndexAt(e.Y);
            if (i < 0) return;
            Item it = Items[i];
            int leftPad = PAD_L + it.Indent * INDENT;
            bool head = it.Kind == K_GROUP || it.Kind == K_CAT || it.Kind == K_ZONE || it.Kind == K_XGRP || it.Kind == K_XWORLD;

            if (head && e.X - leftPad <= ARROW_HIT)
            {
                ToggleFold(it);
                Relayout();
                if (FoldsChanged != null) FoldsChanged(this, EventArgs.Empty);
                return;
            }

            int tx = leftPad + ARROW_W;
            bool hasBox = it.Kind != K_XROW && it.Kind != K_EMPTY && it.Kind != K_XGRP && it.Kind != K_XWORLD;
            if (hasBox)
            {
                int bw = head ? BOX_HEAD : BOX;
                if (e.X - tx <= bw + BOX_GAP)
                {
                    if (it.Kind == K_GROUP || it.Kind == K_CAT || it.Kind == K_ZONE) SetHeading(it, !it.Tri);
                    else SetRow(it.Key, !SIsOn(it.Key));
                    return;
                }
            }

            if (it.Kind == K_ROW)
            {
                if (!SIsOn(it.Key)) SetRow(it.Key, true);
                if (RowActivated != null)
                {
                    AtlasPickEventArgs a = new AtlasPickEventArgs(it.Key);
                    a.Group = it.Group;
                    a.Zone = it.Zone;
                    RowActivated(this, a);
                }
            }
            else if (it.Kind == K_XROW)
            {
                if (RowActivated != null) RowActivated(this, new AtlasPickEventArgs(it.Key + "|" + it.Label));
            }
            else if (it.Kind == K_ZONE)
            {
                if (ZoneActivated != null)
                {
                    AtlasPickEventArgs a = new AtlasPickEventArgs(it.Key);
                    a.Group = it.Group;
                    a.Zone = it.Label;
                    ZoneActivated(this, a);
                }
            }
            else if (it.Kind == K_XWORLD)
            {
                xOpen[it.Key] = true;
                Relayout();
                if (FoldsChanged != null) FoldsChanged(this, EventArgs.Empty);
            }
            else if (it.Kind == K_EMPTY && it.More > 0)
            {
                if (WorldActivated != null)
                {
                    AtlasPickEventArgs a = new AtlasPickEventArgs(it.Key);
                    a.More = it.More;
                    WorldActivated(this, a);
                }
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left) return;
            if (!draggingThumb) return;
            draggingThumb = false;
            Capture = false;
            Cursor = OverRail(e.X) ? Cursors.SizeNS : Cursors.Default;
            Invalidate();
        }

        private void ToggleFold(Item it)
        {
            if (it.Kind == K_GROUP) Flip(gOpen, it.Key);
            else if (it.Kind == K_XGRP) Flip(xOpen, it.Key);
            else if (it.Kind == K_XWORLD) Flip(xOpen, it.Key);
            else Flip(hOpen, it.Key);
        }

        private static void Flip(Dictionary<string, bool> d, string k)
        {
            bool v;
            bool open = !d.TryGetValue(k, out v) || v;
            d[k] = !open;
        }

        private void SetRow(string key, bool on)
        {
            S.SetOn(key, on);
            Relayout();
            if (FilterChanged != null) FilterChanged(this, EventArgs.Empty);
        }

        private void SetHeading(Item it, bool on)
        {
            Group g = null;
            for (int i = 0; i < S.Groups.Count; i++) if (S.Groups[i].Title == (it.Kind == K_GROUP ? it.Key : it.Group)) g = S.Groups[i];
            if (g == null) return;
            Dictionary<string, bool> hit;
            if (!hits.TryGetValue(g.Title, out hit)) return;

            string hkey = it.Key;
            if (it.Kind != K_GROUP)
            {
                foreach (Bucket b in g.Buckets)
                {
                    string bk = g.Title + "/" + (b.Title ?? "~");
                    if (bk != hkey) continue;
                    for (int r = 0; r < b.Rows.Count; r++)
                    {
                        bool v;
                        if (!hit.TryGetValue(b.Rows[r].Key, out v) || !v) continue;
                        S.SetOn(b.Rows[r].Key, on);
                    }
                }
            }
            else
            {
                for (int r = 0; r < g.Rows.Count; r++)
                {
                    bool v;
                    if (hit.TryGetValue(g.Rows[r].Key, out v) && v) S.SetOn(g.Rows[r].Key, on);
                }
            }
            Relayout();
            if (FilterChanged != null) FilterChanged(this, EventArgs.Empty);
        }

        public void FoldAll(bool open)
        {
            if (open)
            {
                gOpen.Clear();
                hOpen.Clear();
                xOpen.Clear();
            }
            else
            {
                for (int i = 0; i < Items.Count; i++)
                {
                    Item it = Items[i];
                    if (it.Kind == K_GROUP) gOpen[it.Key] = false;
                    else if (it.Kind == K_CAT || it.Kind == K_ZONE) hOpen[it.Key] = false;
                    else if (it.Kind == K_XGRP || it.Kind == K_XWORLD) xOpen[it.Key] = false;
                }
            }
            Relayout();
            if (FoldsChanged != null) FoldsChanged(this, EventArgs.Empty);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            int max = Math.Max(0, contentH - Height);
            int next = scroll - (e.Delta / 120) * RH * 2;
            if (next < 0) next = 0;
            if (next > max) next = max;
            if (next != scroll) { scroll = next; Invalidate(); }
        }
    }

    #endregion

    #region Desktop integration

    public static class Cue
    {
        private const int EM_SETCUEBANNER = 0x1501;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        public static void Set(Control box, string text)
        {
            if (box == null || box.IsDisposed || !box.IsHandleCreated) return;
            try { SendMessage(box.Handle, EM_SETCUEBANNER, new IntPtr(1), text ?? ""); }
            catch { }
        }
    }

    public static class Cache
    {
        public static void Clear()
        {
            Icons.Clear();
        }
    }

    #endregion
}
'@
$script:AtlasCoreReferences = @('System.Drawing', 'System.Windows.Forms', 'System.Web.Extensions')
$script:MarketCoreSource = @'
using System;
using System.Windows.Forms;
namespace Entropia
{
    public class HistHoverForm : Form
    {
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x00000080;
                return cp;
            }
        }
        public HistHoverForm()
        {
            ShowInTaskbar = false;
            TopMost = true;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
        }
    }
}
'@
$script:MarketCoreReferences = @('System.Windows.Forms')

function Initialize-DashboardAtlasTypes
{
	[CmdletBinding()]
	param()
	if (([System.Management.Automation.PSTypeName]'Atlas.Canvas').Type) { return $true }
	Add-Type -ReferencedAssemblies $script:AtlasCoreReferences -WarningAction SilentlyContinue -TypeDefinition $script:AtlasCoreSource -ErrorAction Stop
	return $true
}

function Initialize-DashboardMarketTypes
{
	[CmdletBinding()]
	param()
	if ('Entropia.HistHoverForm' -as [type]) { return $true }
	Add-Type -ReferencedAssemblies $script:MarketCoreReferences -WarningAction SilentlyContinue -TypeDefinition $script:MarketCoreSource -ErrorAction Stop
	return $true
}
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