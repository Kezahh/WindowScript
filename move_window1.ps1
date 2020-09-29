Add-Type @"
    using System;
    using System.Runtime.InteropServices;

	public struct RECT {
		public int Left;
		public int Top;
		public int Right;
		public int Bottom;
	}
	
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

    public class Win32 { 
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
		
		[DllImport("user32.dll")]
		public static extern IntPtr GetForegroundWindow();
		
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);
		
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
		
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool GetWindowInfo(IntPtr hWnd, ref WINDOWINFO lpWind);
	}
	
	public class DWM {
		[DllImport("dwmapi.dll")]
		public static extern int DwmGetWindowAttribute(IntPtr hwnd, uint dwAttribute, out RECT pvAttribute, uint cbAttribute);
    }
	
	public class Marshal {
		[DllImport("InteropServices.dll")]
		public static extern int SizeOf(Type t);
	}
"@

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Windows.Forms.SendKeys]::SendWait("%{TAB}")

do {
	$window_handle = [Win32]::GetForegroundWindow()
	$process_name = get-process | where-object { $_.mainwindowhandle -eq $window_handle } | select processName
} while ($process_name -eq $null)

$client_rect = new-object RECT
$window_rect = new-object RECT
$window_info = new-object WINDOWINFO

[void][Win32]::GetClientRect($window_handle, [ref]$client_rect)
[void][Win32]::GetWindowRect($window_handle, [ref]$window_rect)
[void][Win32]::GetWindowInfo($window_handle, [ref]$window_info)

echo "client rect is " ($client_rect.right - $client_rect.left) ($client_rect.bottom - $client_rect.top)
echo "window rect is " ($window_rect.right - $window_rect.left) ($window_rect.bottom - $window_rect.top)
echo $window_info

New-Variable -Name DWMA_EXTENDED_FRAME_BOUNDS -Option Constant -Value 0x09

$attribute_rect = new-object RECT
$attribute_rect_size = [System.Runtime.InteropServices.Marshal]::SizeOf([System.Type][RECT])

echo "window_handle is " $window_handle
echo "frame bounds is "

[void][DWM]::DwmGetWindowAttribute($window_handle, $DWMA_EXTENDED_FRAME_BOUNDS, [ref]$attribute_rect, $attribute_rect_size)
$attribute_rect

pause
