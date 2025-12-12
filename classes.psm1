<# classes.psm1
    .SYNOPSIS
        Provides a library of custom C# classes compiled in-memory for use within the Entropia Dashboard application. 
		These classes offer low-level system integration, INI file handling, and custom-themed UI controls.

    .DESCRIPTION
        This PowerShell module serves as a foundational library by defining and compiling a suite of custom C# classes directly into the PowerShell session using `Add-Type`. These classes provide essential low-level functionalities and custom UI elements that are not natively available in PowerShell or .NET, enabling the rest of the application to perform advanced tasks.

        The provided classes fall into several categories:

        1.  **Low-Level System Integration (`Custom.Native`, `Custom.MouseHookManager`):**
            *   **`Native`:** A comprehensive static utility class that exposes a wide array of Windows API functions through P/Invoke (`DllImport`). Its capabilities include advanced window management (moving, resizing, changing Z-order), checking window states (minimized, responsive), simulating user input (mouse and keyboard events), and providing the necessary components for an efficient, low-CPU UI message loop.
            *   **`MouseHookManager`:** Provides a simple interface to install and remove a system-wide low-level mouse hook, allowing the application to intercept and react to mouse events globally.

        2.  **Configuration Management (`Custom.IniFile`):**
            *   A self-contained class for parsing and generating INI-formatted files. It reads INI files into structured `OrderedDictionary` objects and can write them back to disk, forming the core of the application's configuration persistence mechanism.

        3.  **External Library Integration (`Custom.Ftool`):**
            *   A specialized class that acts as a bridge to an external `ftool.dll` library. It uses P/Invoke to call a specific function within the DLL, demonstrating how the application integrates with pre-compiled native components.

        4.  **Custom "Dark Mode" UI Controls:**
            *   A collection of owner-drawn Windows Forms controls designed to give the application a consistent, modern, dark-themed appearance. By overriding the `OnPaint` method, these classes take full control of their rendering.
            *   Includes `DarkComboBox`, `DarkTabControl`, a `Toggle` switch (based on `CheckBox`), and a `TextProgressBar`.

        Upon import, the module immediately compiles this C# source code, making these classes available for instantiation and use by other application modules.
#>

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
#endregion Global Configuration

#region C# Class Definitions
    $classes = @"
using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Management.Automation; 

namespace Custom
{
    public static class MouseHookManager {
        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static HookProc hookDelegate; // GC protector
        public static IntPtr HookId = IntPtr.Zero;

        public static void Start(HookProc proc) {
            hookDelegate = proc;
            HookId = SetWindowsHookEx(14, proc, GetModuleHandle(null), 0);
        }

        public static void Stop() {
            if (HookId != IntPtr.Zero) {
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
        // ---------------------------------------------------------
        // Original Source A Imports
        // ---------------------------------------------------------
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
        
        [DllImport("user32.dll", SetLastError=true)]
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

        // ---------------------------------------------------------
        // ADDED: Missing Window Manipulation from Source B
        // ---------------------------------------------------------
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
            // Check if running on 64-bit (8 bytes) or 32-bit (4 bytes)
            return IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, nIndex) : GetWindowLong32(hWnd, nIndex);
        }

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern IntPtr SetWindowLong32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        public static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            return IntPtr.Size == 8 ? SetWindowLongPtr64(hWnd, nIndex, dwNewLong) : SetWindowLong32(hWnd, nIndex, dwNewLong);
        }
        // ---------------------------------------------------------

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
            if (Native.IsWindowMinimized(windowHandle)) Native.ShowWindow(windowHandle, Native.SW_RESTORE);
            return Native.SetForegroundWindow(windowHandle);
        }

        public static bool SendToBack(object windowIdentifier)
        {
            IntPtr windowHandle = GetWindowHandle(windowIdentifier);
            if (windowHandle == IntPtr.Zero) return false;
            return Native.ShowWindow(windowHandle, Native.SW_MINIMIZE);
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
                IntPtr sendResult = Native.GetWindowResponse(hWnd, Native.WM_NULL, IntPtr.Zero, IntPtr.Zero, Native.SMTO_ABORTIFHUNG, timeout, out result);
                return sendResult != IntPtr.Zero;
            }
            catch { return false; }
        }

        public static Task<bool> ResponsiveAsync(IntPtr hWnd, uint timeout = 100)
        {
            var tcs = new TaskCompletionSource<bool>();
            ThreadPool.QueueUserWorkItem(state => {
                try { tcs.SetResult(Responsive(hWnd, timeout)); }
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
        
        // GetWindow commands
        public const uint GW_HWNDNEXT = 2;
        public const uint GW_HWNDPREV = 3;

        // Window States
        public const int SW_HIDE = 0;       // Added from B
        public const int SW_SHOW = 5;       // Added from B
        public const int SW_MAXIMIZE = 3;   // Added from B
        public const int SW_MINIMIZE = 6;
        public const int SW_RESTORE = 9;

        public const uint QS_ALLINPUT = 0x04FF;
        public const uint PM_REMOVE = 0x0001;
        
        // Positioning Flags
        public const uint SWP_NOSIZE = 0x0001;      
        public const uint SWP_NOMOVE = 0x0002;      
        public const uint SWP_NOZORDER = 0x0004;    
        public const uint SWP_NOACTIVATE = 0x0010;
        public const uint SWP_FRAMECHANGED = 0x0020; // <--- ADD THIS
        public const int SWP_SHOWWINDOW = 0x0040;   
        
        // Other constants added from B
        public const int ERROR_TIMEOUT = 1460;
        public const uint WAIT_TIMEOUT = 258;
        public const int GWL_EXSTYLE = -20;
        public const uint WS_EX_TOPMOST = 0x00000008;

        public const byte VK_MENU = 0x12; 
        public const uint KEYEVENTF_EXTENDEDKEY = 0x0001; 
        public const uint KEYEVENTF_KEYUP = 0x0002; 

        [Flags]
        public enum WindowPositionOptions : uint 
        { 
            NoZOrderChange = SWP_NOZORDER,   // Added from B
            DoNotActivate = SWP_NOACTIVATE,
            MakeVisible = SWP_SHOWWINDOW     // Added from B
        }
    }

    public static class Ftool
    {
        [DllImport(@"$($global:DashboardConfig.Paths.FtoolDLL)", EntryPoint = "fnPostMessage", CallingConvention = CallingConvention.StdCall)]
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
                File.WriteAllLines(filePath, lines);
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
            this.SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            this.DrawMode = DrawMode.OwnerDrawFixed;
            this.DropDownStyle = ComboBoxStyle.DropDownList;
            this.FlatStyle = FlatStyle.Flat;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Rectangle bounds = this.ClientRectangle;
            using (SolidBrush backBrush = new SolidBrush(this.Enabled ? BackgroundColor : SystemColors.ControlDark))
                g.FillRectangle(backBrush, bounds);

            if (this.DropDownStyle == ComboBoxStyle.DropDownList)
            {
                string text = this.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    Rectangle textRect = new Rectangle(bounds.Left - 8, bounds.Top, bounds.Width, bounds.Height);
                    TextRenderer.DrawText(g, text, this.Font, textRect, this.Enabled ? TextColor : SystemColors.GrayText, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }
            }
            using (Pen borderPen = new Pen(this.Enabled ? BorderColor : SystemColors.ControlDarkDark))
                g.DrawRectangle(borderPen, bounds.Left, bounds.Top, bounds.Width - 2, bounds.Height);

            Rectangle buttonRect = new Rectangle(bounds.Width, bounds.Top, bounds.Left, bounds.Height);
            Point center = new Point(buttonRect.Left + buttonRect.Width / 2, buttonRect.Top + buttonRect.Height / 2);
            Point[] arrowPoints = new Point[] { new Point(center.X - 5, center.Y - 2), new Point(center.X + 5, center.Y - 2), new Point(center.X, center.Y + 3) };
            using (SolidBrush arrowBrush = new SolidBrush(this.Enabled ? ArrowColor : SystemColors.GrayText))
                g.FillPolygon(arrowBrush, arrowPoints);
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= this.Items.Count) return;
            Graphics g = e.Graphics;
            Rectangle bounds = e.Bounds;
            Color currentBackColor = ((e.State & DrawItemState.Selected) == DrawItemState.Selected) ? SelectedItemBackColor : DropDownBackColor;
            Color currentForeColor = ((e.State & DrawItemState.Selected) == DrawItemState.Selected) ? SelectedItemForeColor : TextColor;

            using (SolidBrush backBrush = new SolidBrush(currentBackColor))
                g.FillRectangle(backBrush, bounds.Left, bounds.Top, bounds.Width, bounds.Height);

            string itemText = this.Items[e.Index].ToString();
            Rectangle textBounds = new Rectangle(bounds.Left, bounds.Top, bounds.Width, bounds.Height);
            TextRenderer.DrawText(g, itemText, this.Font, textBounds, currentForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    public class DarkTabControl : TabControl
    {
        public DarkTabControl()
        {
            this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            this.DrawMode = TabDrawMode.OwnerDrawFixed;
            this.SizeMode = TabSizeMode.Fixed;
            this.ItemSize = new Size(120, 30);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(Color.FromArgb(30, 30, 30)))
            {
                e.Graphics.FillRectangle(b, this.ClientRectangle);
            }
            
            for (int i = 0; i < this.TabCount; i++)
            {
                Rectangle tabRect = this.GetTabRect(i);
                bool isSelected = (this.SelectedIndex == i);
                
                Color bg = isSelected ? Color.FromArgb(50, 50, 50) : Color.FromArgb(30, 30, 30);
                
                using (SolidBrush b = new SolidBrush(bg))
                {
                    e.Graphics.FillRectangle(b, tabRect);
                }
                
                string text = this.TabPages[i].Text;
                TextRenderer.DrawText(e.Graphics, text, this.Font, tabRect, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                
                if (isSelected)
                {
                    using (Pen p = new Pen(Color.FromArgb(0, 120, 215), 2))
                    {
                        e.Graphics.DrawLine(p, tabRect.Left, tabRect.Bottom - 1, tabRect.Right, tabRect.Bottom - 1);
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
            this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
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
                path.AddArc(d, d, r, r, 90, 180);
                path.AddArc(this.Width - r - d, d, r, r, -90, 180);
                path.CloseFigure();
                using(var brush = new SolidBrush(this.Checked ? onColor : offColor)) e.Graphics.FillPath(brush, path);

                r = this.Height - 6;
                var rect = this.Checked ? new Rectangle(this.Width - r - 3, 3, r, r) : new Rectangle(3, 3, r, r);
                e.Graphics.FillEllipse(Brushes.White, rect);
                
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
                    e.Graphics.DrawString(this.Text, this.Font, shadowBrush, shadowBounds, sf);
                    e.Graphics.DrawString(this.Text, this.Font, textBrush, boundsF, sf);
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
                this.Invalidate(); // Force redraw when text changes
            }
        }
        
        public TextProgressBar()
        {
            CustomText = "";
            this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle rect = ClientRectangle;
            Graphics g = e.Graphics;

            ProgressBarRenderer.DrawHorizontalBar(g, rect);
            
            if (this.Value > 0)
            {
                Rectangle clip = new Rectangle(rect.X, rect.Y, (int)Math.Round(((float)this.Value / this.Maximum) * rect.Width), rect.Height);
                
                // Draw default chunks manually to control color better or use renderer
                using (SolidBrush b = new SolidBrush(Color.FromArgb(0, 120, 215)))
                {
                    g.FillRectangle(b, clip);
                }
            }

            using (Font f = new Font("Segoe UI", 8))
            {
                string text = string.IsNullOrEmpty(CustomText) ? "" : CustomText;
                SizeF len = g.MeasureString(text, f);
                Point location = new Point((int)((rect.Width / 2) - (len.Width / 2)), (int)((rect.Height / 2) - (len.Height / 2)));
                
                // Draw text shadow
                g.DrawString(text, f, Brushes.Black, location.X + 1, location.Y + 1);
                // Draw text
                g.DrawString(text, f, Brushes.White, location);
            }
        }
    }
}
"@ 
#endregion Step: Define C# Class String

#region Module Initialization
        function Initialize-ClassesModule
        {
            [CmdletBinding()]
            param()
            try
            {
                Write-Verbose "Initializing system integration classes..." -ForegroundColor Cyan
                $addTypeArgs = @{
                    TypeDefinition       = $classes
                    Language             = $script:CompilationOptions.Language
                    ReferencedAssemblies = $script:CompilationOptions.ReferencedAssemblies
                    WarningAction        = $script:CompilationOptions.WarningAction
                    ErrorAction          = 'Stop'
                }
                Add-Type @addTypeArgs
                Write-Verbose "System integration classes initialized successfully." -ForegroundColor Green
                return $true
            }
            catch
            {
                Write-Error "Failed to initialize/compile system integration classes. Error: $($_.Exception.Message)"
                return $false
            }
        }
    Initialize-ClassesModule
#endregion Module Initialization

#region Module Exports
        Export-ModuleMember -Function Initialize-ClassesModule
#endregion Module Exports