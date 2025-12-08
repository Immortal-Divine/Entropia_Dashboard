<# classes.psm1
    .SYNOPSIS
        Core System Integration Module for Entropia Dashboard.
    .DESCRIPTION
        This module provides low-level system access capabilities for the Entropia Dashboard:
        - Window management (position, focus, minimize, restore)
        - Process detection and monitoring
        - Asynchronous window state checking
        - Configuration file handling through INI files
        - External DLL integration for ftool interaction
        - Thread-safe UI operations

        Classes Defined (Compiled from C#):
        - Native: Windows API access for window and process management via P/Invoke.
        - Ftool: Game-specific DLL integration for enhanced functionality.
        - IniFile: Configuration file reading and writing.
        - DarkComboBox: Custom styled ComboBox for dark theme UI.
    .NOTES
        Author: Immortal / Divine
        Version: 1.2.1
        Requires: PowerShell 5.1+, .NET Framework 4.5+, Administrator rights

        Documentation Standards Followed:
        - Module Level Documentation: Synopsis, Description, Notes.
        - Function Level Documentation: Synopsis, Parameter Descriptions, Output Specifications.
        - Code Organization: Logical grouping using #region / #endregion. Functions organized by workflow.
        - Step Documentation: Code blocks enclosed in '#region Step: Description' / '#endregion Step: Description'.
        - Variable Definitions: Inline comments describing the purpose of significant variables.
        - Error Handling: Comprehensive try/catch/finally blocks with error logging and user notification.

        This module compiles C# code at runtime using Add-Type. Changes to the C# source
        require restarting the PowerShell session or re-importing the module.
#>

#region Global Configuration
    #region Step: Define Referenced Assemblies
        # $script:ReferencedAssemblies: Specifies .NET assemblies required by the C# classes.
        $script:ReferencedAssemblies = @(
            'System.Windows.Forms',
            'System.Drawing',
			'System.Management.Automation'
        )
    #endregion Step: Define Referenced Assemblies

    #region Step: Define Compilation Options
        # $script:CompilationOptions: Hashtable storing parameters for the Add-Type cmdlet used to compile the C# classes.
        $script:CompilationOptions = @{
            Language             = 'CSharp'
            ReferencedAssemblies = $script:ReferencedAssemblies
            WarningAction        = 'SilentlyContinue' # Suppress compilation warnings
        }
    #endregion Step: Define Compilation Options
#endregion Global Configuration

#region C# Class Definitions
    <#
    .SYNOPSIS
        C# class definitions for system integration with Windows APIs and .NET Framework.
    .DESCRIPTION
        Contains the following class definitions embedded within a PowerShell here-string:
        - Native: P/Invoke definitions for Windows API functions (window management, process info, etc.).
        - Ftool: Integration with the FTool DLL for game automation via P/Invoke.
        - IniFile: Class for reading and writing configuration data from/to INI files.
        - DarkComboBox: Custom-styled ComboBox control inheriting from System.Windows.Forms.ComboBox for dark UI themes.
        - HotkeyHandler: Inherits from NativeWindow to intercept WM_HOTKEY messages for a specific form.
    .NOTES
        These classes are compiled at runtime using Add-Type within the Initialize-ClassesModule function.
        The C# code uses P/Invoke extensively for low-level system access.
        The Ftool class path depends on the $global:DashboardConfig.Paths.FtoolDLL variable being set correctly elsewhere.
    #>
    #region Step: Define C# Class String
        # $classes: Here-string containing the C# source code for Native, Ftool, IniFile, DarkComboBox, and HotkeyHandler classes.
        # This string is passed to Add-Type for compilation.
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
	using System.Management.Automation; // Added for ScriptBlock

	namespace Custom
	{
	/// <summary>
	/// Windows API functions for window and process management via P/Invoke (Platform Invoke).
	/// Provides static methods to interact with the Windows operating system at a low level.
	/// </summary>
	public static class Native
	{
		// Position window: Sets the size, position, and Z order of a window.
		[DllImport("user32.dll", EntryPoint = "SetWindowPos", SetLastError = true)]
		public static extern bool PositionWindow(
			IntPtr windowHandle,      // Handle to the window
			IntPtr insertAfterHandle, // Handle to the window to precede the positioned window in the Z order
			int X,                    // New position of the left side of the window
			int Y,                    // New position of the top of the window
			int width,                // New width of the window
			int height,               // New height of the window
			uint flags); // Window sizing and positioning flags

		// Get window size: Retrieves the dimensions of the bounding rectangle of the specified window.
		[DllImport("user32.dll", EntryPoint = "GetWindowRect")]
		public static extern bool GetWindowRect(
			IntPtr windowHandle, // Handle to the window
			out RECT rect);      // Pointer to a structure that receives the screen coordinates of the window

		// Change window state: Sets the specified window's show state.
		[DllImport("user32.dll", EntryPoint = "ShowWindow")]
		public static extern bool ShowWindow(
			IntPtr windowHandle, // Handle to the window
			int nCmdShow);       // Controls how the window is to be shown (e.g., SW_MINIMIZE, SW_RESTORE)

		// Focus window: Brings the thread that created the specified window into the foreground and activates the window.
		[DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
		public static extern bool SetForegroundWindow(
			IntPtr windowHandle); // Handle to the window

		// Get focused window: Retrieves a handle to the foreground window (the window with which the user is currently working).
		[DllImport("user32.dll", EntryPoint = "GetForegroundWindow")]
		public static extern IntPtr GetForegroundWindow();

		// Check if a window handle is valid
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsWindow(IntPtr hWnd);

		// Get Thread Process ID: Retrieves the identifier of the thread that created the specified window and, optionally, the identifier of the process that created the window.
		[DllImport("user32.dll", SetLastError = true)]
		public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

		// Get active window (alternative name): Retrieves the window handle to the active window attached to the calling thread's message queue.
		[DllImport("user32.dll", EntryPoint = "GetActiveWindow")]
		public static extern IntPtr GetActiveWindowHandle();

		// Check visibility: Determines the visibility state of the specified window.
		[DllImport("user32.dll", EntryPoint = "IsWindowVisible")]
		public static extern bool IsWindowActive( // Note: Renamed in C# for clarity, maps to IsWindowVisible API
			IntPtr windowHandle); // Handle to the window

		// Check if minimized: Determines whether the specified window is minimized (iconic).
		[DllImport("user32.dll", EntryPoint = "IsIconic")]
		public static extern bool IsWindowMinimized(
			IntPtr windowHandle); // Handle to the window

		// Change title: Changes the text of the specified window's title bar (if it has one).
		[DllImport("user32.dll", EntryPoint = "SetWindowText", SetLastError = true, CharSet = CharSet.Auto)]
		public static extern bool SetWindowTitle(
			IntPtr windowHandle, // Handle to the window
			string windowTitle); // The new title text

		// Test responsiveness: Sends the specified message to one or more windows, waiting for a response or timeout.
		// Used here to check if a window is responding (not hung).
		[DllImport("user32.dll", EntryPoint = "SendMessageTimeout", SetLastError = true)]
		public static extern IntPtr GetWindowResponse( // Note: Renamed in C# for clarity, maps to SendMessageTimeout API
			IntPtr hWnd,         // Handle to the window whose window procedure will receive the message
			uint Msg,            // The message to be sent (WM_NULL is often used for responsiveness checks)
			IntPtr wParam,       // Additional message-specific information
			IntPtr lParam,       // Additional message-specific information
			uint fuFlags,        // How to send the message (e.g., SMTO_ABORTIFHUNG)
			uint uTimeout,       // The duration of the time-out period, in milliseconds
			out IntPtr lpdwResult); // Receives the result of the message processing

		// Free memory: Empties the working set of the specified process.
		[DllImport("psapi.dll", EntryPoint = "EmptyWorkingSet")]
		public static extern bool EmptyWorkingSet(IntPtr hProcess); // Handle to the process

		// Wait for events: Waits until one or all of the specified objects are in the signaled state or the time-out interval elapses.
		// Can also wait for specific types of messages.
		[DllImport("user32.dll", EntryPoint = "MsgWaitForMultipleObjects", SetLastError = true)]
		public static extern uint AsyncExecution( // Note: Renamed in C# for clarity, maps to MsgWaitForMultipleObjects API
			uint nCount,          // The number of object handles in the array pointed to by pHandles
			IntPtr[] pHandles,    // An array of object handles
			bool bWaitAll,        // If TRUE, the function returns when the state of all objects is signaled
			uint dwMilliseconds,  // The time-out interval, in milliseconds
			uint dwWakeMask);     // The types of input events for which to wait (e.g., QS_ALLINPUT)

		// Check messages: Dispatches incoming nonqueued messages, checks the thread message queue for a posted message, and retrieves the message (if any exist).
		[DllImport("user32.dll", EntryPoint = "PeekMessage", SetLastError = true)]
		public static extern bool PeekMessage(
			out MSG lpMsg,         // Pointer to an MSG structure that receives message information
			IntPtr hWnd,           // A handle to the window whose messages are to be retrieved (NULL for thread messages)
			uint wMsgFilterMin,    // The value of the first message in the range of messages to be examined
			uint wMsgFilterMax,    // The value of the last message in the range of messages to be examined
			uint wRemoveMsg);      // Specifies how messages are to be handled (e.g., PM_REMOVE)

		// Process key messages: Translates virtual-key messages into character messages.
		[DllImport("user32.dll", EntryPoint = "TranslateMessage" )]
		public static extern bool TranslateMessage(
			ref MSG lpMsg); // Pointer to an MSG structure that contains message information retrieved from the calling thread's message queue

		// Send message to system: Dispatches a message to a window procedure.
		[DllImport("user32.dll", EntryPoint = "DispatchMessage" )]
		public static extern IntPtr DispatchMessage(
			ref MSG lpMsg); // Pointer to the structure that contains the message

		// Release mouse: Releases mouse capture from a window in the current thread.
		[DllImport("user32.dll", EntryPoint = "ReleaseCapture" )]
		public static extern bool ReleaseCapture();

		// Send window message: Sends the specified message to a window or windows.
		[DllImport("user32.dll", EntryPoint = "SendMessage" )]
		public static extern IntPtr SendMessage(
			IntPtr hWnd,    // Handle to the window whose window procedure will receive the message
			int Msg,        // The message to be sent
			int wParam,     // Additional message-specific information
			int lParam);    // Additional message-specific information

		/// <summary>
		/// Helper method to get a window handle (IntPtr) from either an existing handle or a process ID.
		/// </summary>
		/// <param name="windowIdentifier">Either an IntPtr window handle or an int process ID.</param>
		/// <returns>The window handle (IntPtr) or IntPtr.Zero if not found.</returns>
		public static IntPtr GetWindowHandle(object windowIdentifier)
		{
			if (windowIdentifier is IntPtr)
			{
				return (IntPtr)windowIdentifier;
			}
			if (windowIdentifier is int)
			{
				return FindWindowByProcessId((int)windowIdentifier);
			}
			return IntPtr.Zero; // Return zero if identifier type is invalid
		}

		/// <summary>
		/// Finds the main window handle for a given process ID.
		/// </summary>
		/// <param name="processId">The ID of the process.</param>
		/// <returns>The main window handle (IntPtr) or IntPtr.Zero if the process is not found or has no main window.</returns>
		public static IntPtr FindWindowByProcessId(int processId)
		{
			// Direct access using System.Diagnostics.Process for speed
			try
			{
				Process process = Process.GetProcessById(processId);
				return process.MainWindowHandle; // Can be IntPtr.Zero if no main window
			}
			catch (ArgumentException) // Catch if process ID is not found
			{
				return IntPtr.Zero;
			}
		}

		/// <summary>
		/// Brings the specified window to the foreground. If minimized, it restores it first.
		/// </summary>
		/// <param name="windowIdentifier">Either an IntPtr window handle or an int process ID.</param>
		/// <returns>True if the window was successfully brought to the front, false otherwise.</returns>
		public static bool BringToFront(object windowIdentifier)
		{
			IntPtr windowHandle = GetWindowHandle(windowIdentifier);
			if (windowHandle == IntPtr.Zero)
			{
				return false; // Window not found
			}
			// If the window is minimized, restore it first
			if (Native.IsWindowMinimized(windowHandle))
			{
				Native.ShowWindow(windowHandle, Native.SW_RESTORE);
			}
			// Set the window to the foreground
			return Native.SetForegroundWindow(windowHandle);
		}

		/// <summary>
		/// P/Invoke definitions for mouse-related actions.
		/// </summary>
		[DllImport("user32.dll")]
		public static extern bool ClientToScreen(IntPtr hWnd,ref Point lpPoint); // Converts client coordinates to screen coordinates
		[DllImport("user32.dll",SetLastError=true)]
		public static extern bool SetCursorPos(int x,int y); // Moves the cursor to the specified screen coordinates
		[DllImport("user32.dll")]
		public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo); // Synthesizes mouse motion and button clicks

		/// <summary>
		/// Minimizes the specified window.
		/// </summary>
		/// <param name="windowIdentifier">Either an IntPtr window handle or an int process ID.</param>
		/// <returns>True if the window was successfully minimized, false otherwise.</returns>
		public static bool SendToBack(object windowIdentifier) // Note: Name implies sending back, but action is minimization
		{
			IntPtr windowHandle = GetWindowHandle(windowIdentifier);
			if (windowHandle == IntPtr.Zero)
			{
				return false; // Window not found
			}
			// Minimize the window
			return Native.ShowWindow(windowHandle, Native.SW_MINIMIZE);
		}

		/// <summary>
		/// Checks if the specified window is currently minimized.
		/// </summary>
		/// <param name="windowIdentifier">Either an IntPtr window handle or an int process ID.</param>
		/// <returns>True if the window is minimized, false otherwise or if not found.</returns>
		public static bool IsMinimized(object windowIdentifier)
		{
			IntPtr windowHandle = GetWindowHandle(windowIdentifier);
			if (windowHandle == IntPtr.Zero)
			{
				return false; // Window not found, assume not minimized
			}
			return Native.IsWindowMinimized(windowHandle);
		}

		/// <summary>
		/// Checks if a window is responding by sending it a WM_NULL message with a timeout.
		/// </summary>
		/// <param name="hWnd">The handle of the window to check.</param>
		/// <param name="timeout">Timeout in milliseconds to wait for a response.</param>
		/// <returns>True if the window responded within the timeout, false otherwise.</returns>
		public static bool Responsive(IntPtr hWnd, uint timeout = 100)
		{
			try
			{
				if (hWnd == IntPtr.Zero) return false; // Invalid handle
				IntPtr result;
				// Send WM_NULL message and wait for a response or timeout
				IntPtr sendResult = Native.GetWindowResponse(
					hWnd,
					Native.WM_NULL,          // Message to send (no operation)
					IntPtr.Zero,             // wParam
					IntPtr.Zero,             // lParam
					Native.SMTO_ABORTIFHUNG, // Flags: Abort if the window is hung
					timeout,                 // Timeout duration
					out result);             // Receives the result

				// Success is indicated by a non-zero return value from SendMessageTimeout
				return sendResult != IntPtr.Zero;
			}
			catch // Catch potential exceptions during the API call
			{
				return false; // Assume not responsive if an error occurs
			}
		}

		/// <summary>
		/// Asynchronously checks if a window is responding.
		/// Wraps the synchronous Responsive method in a Task.
		/// </summary>
		/// <param name="hWnd">The handle of the window to check.</param>
		/// <param name="timeout">Timeout in milliseconds.</param>
		/// <returns>A Task that resolves to true if responsive, false otherwise.</returns>
		public static Task<bool> ResponsiveAsync(IntPtr hWnd, uint timeout = 100)
		{
			var tcs = new TaskCompletionSource<bool>();
			// Queue the responsiveness check to run on a thread pool thread
			ThreadPool.QueueUserWorkItem(state =>
				{
					try
					{
						bool isResponsive = Responsive(hWnd, timeout);
						tcs.SetResult(isResponsive); // Set the task result
					}
					catch (Exception ex)
					{
						tcs.SetException(ex); // Set exception if the check fails
					}
				});
			return tcs.Task; // Return the task
		}

		/// <summary>
		/// Structure to hold window coordinates (used by GetWindowRect).
		/// </summary>
		[StructLayout(LayoutKind.Sequential)]
		public struct RECT
		{
			public int Left;   // Specifies the x-coordinate of the upper-left corner of the rectangle.
			public int Top;    // Specifies the y-coordinate of the upper-left corner of the rectangle.
			public int Right;  // Specifies the x-coordinate of the lower-right corner of the rectangle.
			public int Bottom; // Specifies the y-coordinate of the lower-right corner of the rectangle.
		}

		/// <summary>
		/// Structure containing message information from a thread's message queue (used by PeekMessage, etc.).
		/// </summary>
		[StructLayout(LayoutKind.Sequential)]
		public struct MSG
		{
			public IntPtr hwnd;      // A handle to the window whose window procedure receives the message.
			public uint message;     // The message identifier.
			public UIntPtr wParam;   // Additional information about the message. Depends on the message value.
			public IntPtr lParam;    // Additional information about the message. Depends on the message value.
			public uint time;        // The time at which the message was posted.
			public System.Drawing.Point pt; // The cursor position, in screen coordinates, when the message was posted.
		}

		// --- Constants used by the Native methods ---
		        // Window Handles
				public static readonly IntPtr TopWindowHandle = new IntPtr(0); // HWND_TOP
				public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
				public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
		
				// Window Messages
				public const uint WM_NULL = 0x0000; // Null message, often used for responsiveness checks
		
				// Error Codes
				public const int ERROR_TIMEOUT = 1460; // Timeout error code
		
				// SendMessageTimeout Flags
				public const uint SMTO_ABORTIFHUNG = 0x0002; // Do not wait if the target thread is hung
		
				// ShowWindow Commands (nCmdShow parameter)
				public const int SW_HIDE = 0;       // Hides the window and activates another window.
				public const int SW_MINIMIZE = 6;   // Minimizes the specified window and activates the next top-level window in the Z order.
				public const int SW_RESTORE = 9;    // Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position.
				public const int SW_SHOW = 5;       // Activates the window and displays it in its current size and position.
				public const int SW_MAXIMIZE = 3;   // Maximizes the specified window.
		
				// Queue Status Flags (dwWakeMask for MsgWaitForMultipleObjects)
				public const uint QS_ALLINPUT = 0x04FF; // Any message is in the queue.
		
				// PeekMessage Flags (wRemoveMsg parameter)
				public const uint PM_REMOVE = 0x0001; // Messages are removed from the queue after processing by PeekMessage.
		
				// Wait Constants (Return value for MsgWaitForMultipleObjects)
				public const uint WAIT_TIMEOUT = 258; // The time-out interval elapsed, and the object's state is nonsignaled.
		
				// SetWindowPos Flags (Combined with WindowPositionOptions enum)
				public const uint SWP_NOSIZE = 0x0001;
				public const uint SWP_NOMOVE = 0x0002;
						public const uint SWP_NOZORDER = 0x0004;   // Retains the current Z order (ignores insertAfterHandle).
						public const int SWP_NOACTIVATE = 0x0010; // Does not activate the window.
						public const int SWP_SHOWWINDOW = 0x0040; // Displays the window.
				
						// For GetWindowLong
						public const int GWL_EXSTYLE = -20;
						public const uint WS_EX_TOPMOST = 0x00000008;
				
						[DllImport("user32.dll", SetLastError = true)]
						public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);
						
						[DllImport("user32.dll", EntryPoint="GetWindowLong")]
						public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
				
						// Keyboard Event Constants
						public const byte VK_MENU = 0x12; // ALT key
						public const uint KEYEVENTF_EXTENDEDKEY = 0x0001; // Key down flag
						public const uint KEYEVENTF_KEYUP = 0x0002; // Key up flag

		/// <summary>
		/// Flags for the SetWindowPos function (PositionWindow method).
		/// </summary>
		[Flags]
		public enum WindowPositionOptions : uint
		{
			NoZOrderChange = SWP_NOZORDER,   // Retains the current Z order (ignores the hWndInsertAfter parameter).
			DoNotActivate = SWP_NOACTIVATE, // Does not activate the window. If this flag is not set, the window is activated and moved to the top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter parameter).
			MakeVisible = SWP_SHOWWINDOW    // Displays the window.
		}
	}

	/// <summary>
	/// Provides access to functions within the Ftool.dll library via P/Invoke.
	/// Used for game-specific automation tasks. Requires Ftool.dll to be present.
	/// The path is dynamically determined by $global:DashboardConfig.Paths.FtoolDLL.
	/// </summary>
	public static class Ftool
	{
		// Post message using Ftool: Sends a message via the Ftool DLL's specific mechanism.
		// The exact path to Ftool.dll is interpolated from the PowerShell global variable at compile time.
		[DllImport(@"$($global:DashboardConfig.Paths.FtoolDLL)", EntryPoint = "fnPostMessage", CallingConvention = CallingConvention.StdCall)]
		public static extern void fnPostMessage(
			IntPtr hWnd,    // Handle to the target window
			int msg,        // Message identifier
			int wParam,     // Message parameter
			int lParam);    // Message parameter
	}

	/// <summary>
	/// Class to handle reading from and writing to INI configuration files.
	/// Provides methods to parse sections, keys, and values.
	/// Uses OrderedDictionary to preserve key order within sections.
	/// Includes retry logic for file access.
	/// </summary>
	public class IniFile
	{
		private string filePath; // Stores the full path to the INI file

		/// <summary>
		/// Initializes a new instance of the IniFile class.
		/// </summary>
		/// <param name="filePath">The path to the INI file to be managed.</param>
		public IniFile(string filePath)
		{
			this.filePath = filePath;
		}

		/// <summary>
		/// Reads the entire INI file and returns its structure as a nested OrderedDictionary.
		/// Outer dictionary keys are section names, values are OrderedDictionaries of key-value pairs.
		/// </summary>
		/// <returns>An OrderedDictionary representing the INI file content.</returns>
		public OrderedDictionary ReadIniFile()
		{
			OrderedDictionary config = new OrderedDictionary(); // Use OrderedDictionary to maintain order

			if (!File.Exists(filePath))
			{
				Console.WriteLine("Warning: INI file not found at: " + filePath);
				return config; // Return empty dictionary if file doesn't exist
			}

			try
			{
				string currentSection = null;
				string[] lines = File.ReadAllLines(filePath); // Read all lines from the file

				foreach (string line in lines)
				{
					string trimmedLine = line.Trim(); // Remove leading/trailing whitespace

					// Skip empty lines and comments (lines starting with ';' or '#')
					if (string.IsNullOrWhiteSpace(trimmedLine) || trimmedLine.StartsWith(";") || trimmedLine.StartsWith("#"))
					{
						continue;
					}

					// Check if the line defines a section header (e.g., "[SectionName]")
					if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
					{
						currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2).Trim(); // Extract section name
						if (!config.Contains(currentSection))
						{
							// Add the section as a new OrderedDictionary if it doesn't exist
							config[currentSection] = new OrderedDictionary();
						}
						continue; // Move to the next line after processing section header
					}

					// Process key-value pairs (e.g., "Key = Value")
					int equalsPos = trimmedLine.IndexOf("=");
					if (equalsPos > 0 && currentSection != null) // Ensure '=' exists and we are within a section
					{
						string key = trimmedLine.Substring(0, equalsPos).Trim(); // Extract key
						string value = trimmedLine.Substring(equalsPos + 1).Trim(); // Extract value

						// Remove surrounding quotes from the value if present
						if (value.StartsWith("\"") && value.EndsWith("\""))
						{
							value = value.Substring(1, value.Length - 2);
						}

						// Add the key-value pair to the current section's dictionary
						((OrderedDictionary)config[currentSection])[key] = value;
					}
				}

				return config; // Return the populated dictionary
			}
			catch (Exception ex)
			{
				Console.WriteLine("Error reading INI file '" + filePath + "': " + ex.Message);
				return new OrderedDictionary(); // Return empty dictionary on error
			}
		}

		/// <summary>
		/// Reads a specific section from the INI file.
		/// </summary>
		/// <param name="section">The name of the section to read.</param>
		/// <returns>An OrderedDictionary containing the key-value pairs for the specified section, or an empty dictionary if the section is not found.</returns>
		public OrderedDictionary ReadSection(string section)
		{
			OrderedDictionary config = ReadIniFile(); // Read the whole file first

			if (config.Contains(section))
			{
				return (OrderedDictionary)config[section]; // Return the specific section's dictionary
			}

			// Return an empty dictionary if the section doesn't exist
			return new OrderedDictionary();
		}

		/// <summary>
		/// Writes a single key-value pair to a specified section in the INI file.
		/// If the section or key does not exist, they will be created.
		/// If the key exists, its value will be updated.
		/// </summary>
		/// <param name="section">The name of the section.</param>
		/// <param name="key">The name of the key.</param>
		/// <param name="value">The value to write.</param>
		public void WriteValue(string section, string key, string value)
		{
			// Ensure the directory and file exist before writing
			EnsureFileExists();

			// Read the current content of the file
			List<string> content = new List<string>();
			if (File.Exists(filePath))
			{
				content.AddRange(File.ReadAllLines(filePath));
			}

			// --- Find or create the section ---
			bool sectionFound = false;
			int sectionIndex = -1; // Index of the line containing the section header

			for (int i = 0; i < content.Count; i++)
			{
				if (content[i].Trim().Equals("[" + section + "]", StringComparison.OrdinalIgnoreCase)) // Case-insensitive section match
				{
					sectionFound = true;
					sectionIndex = i;
					break;
				}
			}

			// If section wasn't found, add it at the end of the file
			if (!sectionFound)
			{
				// Add a blank line before the new section if the file is not empty
				if (content.Count > 0 && !string.IsNullOrWhiteSpace(content[content.Count - 1]))
				{
					content.Add("");
				}
				content.Add("[" + section + "]");
				sectionIndex = content.Count - 1; // Update section index to the newly added line
			}

			// --- Find or add the key within the section ---
			bool keyFound = false;
			int keyIndex = -1; // Index of the line containing the key

			// Search for the key only within the target section
			for (int i = sectionIndex + 1; i < content.Count; i++)
			{
				string currentLineTrimmed = content[i].Trim();
				// Stop searching if we hit the start of another section or end of file
				if (currentLineTrimmed.StartsWith("[") && currentLineTrimmed.EndsWith("]"))
				{
					break;
				}

				// Check if the current line starts with the key followed by '=' (allowing for spaces)
				if (currentLineTrimmed.StartsWith(key + " ", StringComparison.OrdinalIgnoreCase) ||
					currentLineTrimmed.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase))
				{
					// More precise check to avoid partial matches (e.g., "Key1" matching "Key")
					int equalsPos = currentLineTrimmed.IndexOf("=");
					if (equalsPos > 0 && currentLineTrimmed.Substring(0, equalsPos).Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
					{
						keyFound = true;
						keyIndex = i;
						break;
					}
				}
			}

			// Format the value: Add quotes if it contains spaces and isn't already quoted
			if (value.Contains(" ") && !(value.StartsWith("\"") && value.EndsWith("\"")))
			{
				value = "\"" + value + "\"";
			}

			// Construct the new line for the key-value pair
			string newLine = key + " = " + value;

			// Update the existing key line or insert the new key line
			if (keyFound)
			{
				content[keyIndex] = newLine; // Replace the existing line
			}
			else
			{
				// Insert the new key-value pair immediately after the section header
				content.Insert(sectionIndex + 1, newLine);
			}

			// Write the modified content back to the file using retry logic
			string[] contentArray = content.ToArray();
			RetryFileOperation(() => { File.WriteAllLines(filePath, contentArray); });
		}

		/// <summary>
		/// Writes all key-value pairs from an OrderedDictionary to a specified section in the INI file.
		/// Calls WriteValue for each entry in the dictionary.
		/// </summary>
		/// <param name="section">The name of the section.</param>
		/// <param name="data">An OrderedDictionary containing the key-value pairs for the section.</param>
		public void WriteSection(string section, OrderedDictionary data)
		{
			foreach (DictionaryEntry entry in data)
			{
				// Ensure keys and values are converted to strings before writing
				WriteValue(section, entry.Key.ToString(), entry.Value.ToString());
			}
		}

		/// <summary>
		/// Writes an entire configuration (represented by a nested OrderedDictionary) to the INI file.
		/// This method effectively overwrites the existing file content with the provided configuration.
		/// </summary>
		/// <param name="config">An OrderedDictionary where keys are section names and values are OrderedDictionaries of key-value pairs.</param>
		public void WriteIniFile(OrderedDictionary config)
		{
			// Ensure the file exists and is empty (overwrite)
			EnsureFileExists(true);

			List<string> lines = new List<string>();
			bool firstSection = true;

			// Iterate through each section in the configuration dictionary
			foreach (DictionaryEntry sectionEntry in config)
			{
				string section = sectionEntry.Key.ToString();
				OrderedDictionary sectionData = (OrderedDictionary)sectionEntry.Value;

				// Add a blank line before sections (except the first one)
				if (!firstSection)
				{
					lines.Add("");
				}
				lines.Add("[" + section + "]"); // Add the section header
				firstSection = false;

				// Iterate through key-value pairs within the section
				foreach (DictionaryEntry kvpEntry in sectionData)
				{
					string key = kvpEntry.Key.ToString();
					string value = kvpEntry.Value.ToString();

					// Format value (add quotes if needed)
					if (value.Contains(" ") && !(value.StartsWith("\"") && value.EndsWith("\"")))
					{
						value = "\"" + value + "\"";
					}
					lines.Add(key + " = " + value); // Add the key-value line
				}
			}

			// Write all constructed lines to the file using retry logic
			RetryFileOperation(() => { File.WriteAllLines(filePath, lines); });
		}

		/// <summary>
		/// Ensures that the directory for the INI file exists, creating it if necessary.
		/// Optionally creates or clears the INI file itself.
		/// </summary>
		/// <param name="overwrite">If true, the file will be created (or cleared if it exists). If false, the file is only created if it doesn't exist.</param>
		private void EnsureFileExists(bool overwrite = false)
		{
			try
			{
				string directory = Path.GetDirectoryName(filePath);

				// Create directory if it doesn't exist
				if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
				{
					Directory.CreateDirectory(directory);
				}

				// Create or clear the file if overwrite is true or if the file doesn't exist
				if (overwrite || !File.Exists(filePath))
				{
					// Use retry logic for the file creation/clearing operation
					RetryFileOperation(() => { File.WriteAllText(filePath, ""); });
				}
			}
			catch (Exception ex)
			{
				// Log error if directory/file creation fails
				Console.WriteLine("Error ensuring INI file exists '" + filePath + "': " + ex.Message);
				// Depending on requirements, might re-throw or handle differently
			}
		}

		/// <summary>
		/// Attempts to perform a file operation, retrying several times with delays if an IOException occurs (e.g., file locked).
		/// </summary>
		/// <param name="operation">The file operation (Action delegate) to perform.</param>
		/// <exception cref="IOException">Throws an IOException if the operation fails after all retry attempts.</exception>
		private void RetryFileOperation(Action operation)
		{
			int maxAttempts = 5;      // Maximum number of retry attempts
			int retryDelayMs = 100;   // Delay between retries in milliseconds

			int attempts = 0;
			while (attempts < maxAttempts)
			{
				try
				{
					operation(); // Attempt the file operation
					return;      // Success, exit the method
				}
				catch (IOException ex) // Catch only IOExceptions, likely related to file access
				{
					attempts++;
					if (attempts >= maxAttempts)
					{
						// Throw a detailed exception if max attempts are reached
						throw new IOException("Failed to access INI file " + filePath + " after " + maxAttempts + " attempts. Last error: " + ex.Message);
					}
					// Wait before the next retry
					Thread.Sleep(retryDelayMs);
				}
				// Other exceptions (e.g., UnauthorizedAccessException) are not caught here and will propagate immediately.
			}
		}
	}

	/// <summary>
	/// Custom ComboBox control with a dark theme appearance.
	/// Overrides painting methods to draw custom background, border, text, and dropdown arrow.
	/// Inherits from System.Windows.Forms.ComboBox.
	/// </summary>
	public class DarkComboBox : ComboBox
	{
		// Windows Messages constants used in WndProc
		private const int WM_PAINT = 0x000F;      // Message sent when the window needs to be painted
		private const int WM_NCPAINT = 0x0085;    // Message sent to paint the non-client area (border, title bar)
		private const int WM_ERASEBKGND = 0x0014; // Message sent when the window background must be erased

		// P/Invoke definitions for drawing
		[DllImport("user32.dll")]
		static extern IntPtr GetWindowDC(IntPtr hWnd); // Gets the device context (DC) for the entire window, including title bar, menus, and scroll bars.
		[DllImport("user32.dll")]
		static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC); // Releases a device context.
		[DllImport("user32.dll")]
		static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase); // Invalidates a rectangular area of a window, adding it to the update region.

		// --- Customizable Colors for the Dark Theme ---
		private Color BackgroundColor = Color.FromArgb(40, 40, 40);     // Dark grey background
		private Color TextColor = Color.FromArgb(240, 240, 240);   // Light grey/white text
		private Color BorderColor = Color.FromArgb(65, 65, 70);     // Slightly lighter grey border
		private Color ArrowColor = Color.FromArgb(240, 240, 240);    // White arrow
		private Color SelectedItemBackColor = Color.FromArgb(0, 120, 200); // Blue highlight for selected item background
		private Color SelectedItemForeColor = Color.FromArgb(240, 240, 240);   // Light grey/white text
		private Color DropDownBackColor = Color.FromArgb(40, 40, 40);      // Background color of the dropdown list

		/// <summary>
		/// Constructor for the DarkComboBox. Sets necessary control styles for custom painting.
		/// </summary>
		public DarkComboBox() : base()
		{
			// Set control styles to enable custom painting and reduce flicker
			this.SetStyle(ControlStyles.UserPaint |             // Control paints itself rather than the OS doing so.
						ControlStyles.AllPaintingInWmPaint |  // Ignore WM_ERASEBKGND to reduce flicker.
						ControlStyles.OptimizedDoubleBuffer | // Paint via a buffer to reduce flicker.
						ControlStyles.ResizeRedraw,           // Redraw when resized.
						true);
			this.DrawMode = DrawMode.OwnerDrawFixed; // Specify that items are drawn manually and are of fixed height
			this.DropDownStyle = ComboBoxStyle.DropDownList; // Prevent user text input, only allow selection
			this.FlatStyle = FlatStyle.Flat; // Use flat style as base for custom drawing
		}

		/// <summary>
		/// Overrides the OnPaint method to handle custom drawing of the ComboBox control itself (the main box).
		/// </summary>
		/// <param name="e">PaintEventArgs containing graphics context and clipping rectangle.</param>
		protected override void OnPaint(PaintEventArgs e)
		{
			Graphics g = e.Graphics;
			Rectangle bounds = this.ClientRectangle;

			// 1. Draw the background
			using (SolidBrush backBrush = new SolidBrush(this.Enabled ? BackgroundColor : SystemColors.ControlDark)) // Use disabled color if needed
			{
				g.FillRectangle(backBrush, bounds);
			}

			// 2. Draw the selected item text (if applicable)
			if (this.DropDownStyle == ComboBoxStyle.DropDownList || this.DropDownStyle == ComboBoxStyle.DropDown)
			{
				string text = this.Text; // Use Text property which reflects selected item or typed text
				if (!string.IsNullOrEmpty(text))
				{
					// Define text rectangle, leaving space for border and dropdown button
					Rectangle textRect = new Rectangle(bounds.Left - 8, bounds.Top, bounds.Width, bounds.Height);
					TextFormatFlags flags = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter;

					// Use appropriate text color based on enabled state
					Color currentTextColor = this.Enabled ? TextColor : SystemColors.GrayText;
					TextRenderer.DrawText(g, text, this.Font, textRect, currentTextColor, flags);
				}
			}

			// 3. Draw the border
			using (Pen borderPen = new Pen(this.Enabled ? BorderColor : SystemColors.ControlDarkDark))
			{
				g.DrawRectangle(borderPen, bounds.Left, bounds.Top, bounds.Width - 2, bounds.Height);
			}

			// 4. Draw the dropdown button background
			Rectangle buttonRect = new Rectangle(bounds.Width, bounds.Top, bounds.Left, bounds.Height);
			using (SolidBrush buttonBrush = new SolidBrush(this.Enabled ? BackgroundColor : SystemColors.ControlDark))
			{
				g.FillRectangle(buttonBrush, buttonRect);
			}

			// 5. Draw the dropdown arrow
			Point center = new Point(buttonRect.Left + buttonRect.Width / 2, buttonRect.Top + buttonRect.Height / 2);
			Point[] arrowPoints = new Point[] {
				new Point(center.X - 5, center.Y - 2), // Top-left
				new Point(center.X + 5, center.Y - 2), // Top-right
				new Point(center.X, center.Y + 3)      // Bottom-center
			};
			using (SolidBrush arrowBrush = new SolidBrush(this.Enabled ? ArrowColor : SystemColors.GrayText))
			{
				g.FillPolygon(arrowBrush, arrowPoints);
			}
		}

		/// <summary>
		/// Overrides the window procedure to handle specific paint-related messages for custom drawing.
		/// This helps ensure consistent appearance, especially for the non-client area (border).
		/// </summary>
		/// <param name="m">The Windows Message.</param>
		protected override void WndProc(ref Message m)
		{
			// Intercept paint messages to ensure custom drawing
			if (m.Msg == WM_PAINT)
			{
				base.WndProc(ref m); // Let the base control handle the paint first
				// Now, perform our custom painting over it, especially the border and button
				using (Graphics g = Graphics.FromHwnd(this.Handle))
				{
					Rectangle bounds = this.ClientRectangle;
					// Redraw border
					using (Pen borderPen = new Pen(BorderColor))
					{
						g.DrawRectangle(borderPen, 0, 0, this.Width - 1, this.Height - 1);
					}
					// Redraw dropdown button area (background and arrow)
					Rectangle buttonRect = new Rectangle(this.Width - 20, 0, 20, this.Height);
					using (SolidBrush buttonBrush = new SolidBrush(BackgroundColor)) { g.FillRectangle(buttonBrush, buttonRect); }
					Point center = new Point(buttonRect.Left + buttonRect.Width / 2, buttonRect.Top + buttonRect.Height / 2);
					Point[] arrowPoints = new Point[] { new Point(center.X - 5, center.Y - 2), new Point(center.X + 5, center.Y - 2), new Point(center.X, center.Y + 3) };
					using (SolidBrush arrowBrush = new SolidBrush(ArrowColor)) { g.FillPolygon(arrowBrush, arrowPoints); }
					// Draw separator line
					using (Pen sepPen = new Pen(BorderColor)) { g.DrawLine(sepPen, buttonRect.Left -1, 0, buttonRect.Left -1, bounds.Bottom); }
				}
				m.Result = IntPtr.Zero; // Indicate message was handled
			}
			else if (m.Msg == WM_ERASEBKGND) 
			{
				// Prevent default background erasing to reduce flicker
				m.Result = (IntPtr)1; // Indicate we handled it
			}
			else
			{
				base.WndProc(ref m); // Handle other messages normally
			}
		}

		/// <summary>
		/// Overrides the OnDrawItem method to handle custom drawing of individual items in the dropdown list.
		/// </summary>
		/// <param name="e">DrawItemEventArgs containing item index, state, graphics context, and bounds.</param>
		protected override void OnDrawItem(DrawItemEventArgs e)
		{
			if (e.Index < 0 || e.Index >= this.Items.Count) return; // Check for valid index

			Graphics g = e.Graphics;
			Rectangle bounds = e.Bounds;

			// Determine background and text colors based on item state (selected, focused, etc.)
			Color currentBackColor;
			Color currentForeColor;

			if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
			{
				currentBackColor = SelectedItemBackColor; // Use highlight color for selected item background
				currentForeColor = SelectedItemForeColor; // Use highlight color for selected item text
			}
			else
			{
				currentBackColor = DropDownBackColor; // Use default dropdown background color
				currentForeColor = TextColor;         // Use default text color
			}

			// Draw the item background
			using (SolidBrush backBrush = new SolidBrush(currentBackColor))
			{
				g.FillRectangle(backBrush, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
			}

			// Draw the item text
			string itemText = this.Items[e.Index].ToString();
			TextFormatFlags flags = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter; // Align center, center vertically
			Rectangle textBounds = new Rectangle(bounds.Left, bounds.Top, bounds.Width, bounds.Height);

			TextRenderer.DrawText(g, itemText, this.Font, textBounds, currentForeColor, flags);

			// Draw focus rectangle if the item has focus (optional, but good for accessibility)
			if ((e.State & DrawItemState.Focus) == DrawItemState.Focus && (e.State & DrawItemState.NoFocusRect) == 0)
			{
				// ControlPaint.DrawFocusRectangle(g, bounds, currentForeColor, currentBackColor); // Standard focus rect
				// Or a custom focus indicator:
				using (Pen focusPen = new Pen(SelectedItemBackColor, 1)) // Use selection color for focus border
				{
					focusPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
					g.DrawRectangle(focusPen, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
				}
			}

			// Call base method if needed, though usually not required with full custom draw
			// base.OnDrawItem(e);
		}

		// Override OnMeasureItem if DrawMode is OwnerDrawVariable
		// protected override void OnMeasureItem(MeasureItemEventArgs e) { ... }

		// Override background color property if needed
		// public override Color BackColor { get => BackgroundColor; set { BackgroundColor = value; Invalidate(); } }
	}

    public class Toggle : CheckBox
    {
        public Toggle()
        {
            this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
            this.Padding = new Padding(1);
            this.Appearance = Appearance.Button; // Important to hide default checkbox appearance
            this.FlatStyle = FlatStyle.Flat;
            this.FlatAppearance.BorderSize = 0;
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            // Base painting
            base.OnPaint(e);
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            
            Color backColor = this.Parent != null ? this.Parent.BackColor : SystemColors.Control;
            e.Graphics.Clear(backColor);

            // Colors
            Color onColor = Color.FromArgb(35, 175, 75); // Green
            Color offColor = Color.FromArgb(210, 45, 45); // Red

            using (var path = new System.Drawing.Drawing2D.GraphicsPath())
            {
                var d = this.Padding.All;
                var r = this.Height - 2 * d;
                path.AddArc(d, d, r, r, 90, 180);
                path.AddArc(this.Width - r - d, d, r, r, -90, 180);
                path.CloseFigure();

                // Fill background track
                using(var brush = new SolidBrush(this.Checked ? onColor : offColor))
                {
                    e.Graphics.FillPath(brush, path);
                }

                // Draw thumb
                r = this.Height - 6;
                var rect = this.Checked ? new Rectangle(this.Width - r - 3, 3, r, r) : new Rectangle(3, 3, r, r);
                e.Graphics.FillEllipse(Brushes.White, rect);
                
				// Draw Text on thumb (legacy) -- replaced by centered, contrast-aware drawing below
				// TextRenderer.DrawText(e.Graphics, this.Text, this.Font, rect, Color.Black, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

				// Draw the symbol centered over the control with a small shadow to ensure
				// visibility against both the red and green track backgrounds.
				Color trackColor = this.Checked ? onColor : offColor;
				// Calculate luminance to pick a contrasting text color
				double luminance = (0.2126 * trackColor.R) + (0.7152 * trackColor.G) + (0.0722 * trackColor.B);
				Color textColor = (luminance < 140) ? Color.White : Color.Black;

				using (var shadowBrush = new SolidBrush(Color.FromArgb(140, 0, 0, 0))) // semi-transparent shadow
				using (var textBrush = new SolidBrush(textColor))
				{
					var sf = new System.Drawing.StringFormat();
					sf.Alignment = System.Drawing.StringAlignment.Center;
					sf.LineAlignment = System.Drawing.StringAlignment.Center;

					var boundsF = new RectangleF(0, 0, this.Width, this.Height);

					// Draw shadow slightly offset
					var shadowOffset = 1;
					var shadowBounds = new RectangleF(boundsF.X + shadowOffset, boundsF.Y + shadowOffset, boundsF.Width, boundsF.Height);
					e.Graphics.DrawString(this.Text, this.Font, shadowBrush, shadowBounds, sf);

					// Draw main text
					e.Graphics.DrawString(this.Text, this.Font, textBrush, boundsF, sf);
				}
            }
        }
	}
	} // namespace Custom

"@ # End of C# here-string
    #endregion Step: Define C# Class String
#endregion C# Class Definitions

#region Module Initialization
    #region Function: Initialize-ClassesModule
        function Initialize-ClassesModule
        {
            <#
            .SYNOPSIS
                Initializes the module by compiling the embedded C# classes using Add-Type.
            .DESCRIPTION
                This function takes the C# code defined in the $classes variable and compiles it
                into the current PowerShell session using the Add-Type cmdlet. This makes the
                Native, Ftool, IniFile, and DarkComboBox classes available for use within the session.
                It uses the configuration stored in $script:CompilationOptions.
            .PARAMETER
                None. This function does not accept any parameters.
            .OUTPUTS
                [bool] Returns $true if the C# classes were compiled successfully, $false otherwise.
            .NOTES
                This function is typically called automatically when the module is imported.
                It includes basic error handling to report compilation failures.
                Requires appropriate .NET Framework version and referenced assemblies to be available.
            #>
            [CmdletBinding()]
            param()

            #region Step: Compile C# Classes with Error Handling
                try
                {
                    #region Step: Display Initialization Message
                        Write-Verbose "Initializing system integration classes (Native, Ftool, IniFile, DarkComboBox)..." -ForegroundColor Cyan
                    #endregion Step: Display Initialization Message

                    #region Step: Define Add-Type Arguments
                        # Prepare arguments for Add-Type based on script-level configuration
                        $addTypeArgs = @{
                            TypeDefinition       = $classes # The C# source code
                            Language             = $script:CompilationOptions.Language
                            ReferencedAssemblies = $script:CompilationOptions.ReferencedAssemblies
                            WarningAction        = $script:CompilationOptions.WarningAction
                            ErrorAction          = 'Stop' # Ensure compilation errors throw terminating exceptions
                        }
                    #endregion Step: Define Add-Type Arguments

                    #region Step: Compile C# Code
                        # Execute Add-Type to compile the classes into the current session
                        Add-Type @addTypeArgs
                    #endregion Step: Compile C# Code

                    #region Step: Handle Success
                        Write-Verbose "System integration classes initialized successfully." -ForegroundColor Green
                        return $true
                    #endregion Step: Handle Success
                }
                catch
                {
                    #region Step: Handle Errors
                        # Log detailed error information if compilation fails
                        Write-Error "Failed to initialize/compile system integration classes. Error: $($_.Exception.Message)"
                        # Optionally log more details: $_ | Format-List * -Force | Out-String | Write-Error
                        return $false
                    #endregion Step: Handle Errors
                }
            #endregion Step: Compile C# Classes with Error Handling
        }
    #endregion Function: Initialize-ClassesModule

    #region Step: Compile Classes on Module Import
        # Automatically attempt to compile the classes when this module is imported.
        # The result ($true/$false) is implicitly returned but usually not captured here.
        Initialize-ClassesModule
    #endregion Step: Compile Classes on Module Import
#endregion Module Initialization

#region Module Exports
    #region Step: Export Public Functions
        # Export the functions intended for use by other modules or the main script.
        Export-ModuleMember -Function Initialize-ClassesModule
#endregion Module Exports