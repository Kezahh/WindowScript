$ALT_TAB_WINDOW_HANDLE = "65950"

Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class Win32 { 
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
        
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);
    }
    
    public class DWM
    {
        [DllImport("dwmapi.dll")]
        public static extern int DwmGetWindowAttribute(IntPtr hwnd, uint dwAttribute, out RECT pvAttribute, uint cbAttribute);
    }
    
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
"@

# position 0 (left), 1 (center), 2 (right)
# x_pos is the left.
# y_pos is the top
$position = $args[0]
$x_pos = 1706*$position
$y_pos = 0

# Alt-tab. Foreground window is currently powershell. Alt-tab to foreground
# window before powershell started.
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Windows.Forms.SendKeys]::SendWait("%{TAB}")

# loop until alt-tab window is no longer foreground
$processName = $null
do {
    $window_handle = [Win32]::GetForegroundWindow()
    $processName = get-process | where-object { $_.mainwindowhandle -eq $window_handle } | select processName
} while ($processName -eq $null)

$window_rect = new-object RECT
$client_rect = new-object RECT
$frame_rect = new-object RECT
$frame_rect_size = [System.Runtime.InteropServices.Marshal]::SizeOf([System.Type][RECT])
New-Variable -Name DWMA_EXTENDED_FRAME_BOUNDS -Option Constant -Value 0x09
[void][Win32]::GetWindowRect($window_handle, [ref]$window_rect)
[void][Win32]::GetClientRect($window_handle, [ref]$client_rect)
[void][DWM]::DwmGetWindowAttribute($window_handle, $DWMA_EXTENDED_FRAME_BOUNDS, [ref]$frame_rect, $frame_rect_size)


$border_width = (($window_rect.Right - $window_rect.Left) - ($frame_rect.Right - $frame_rect.Left))
$border_height = (($window_rect.Bottom - $window_rect.Top) - ($frame_rect.Bottom - $frame_rect.Top))

$target_width = 1706
$target_height = 1400

[void][Win32]::MoveWindow($window_handle, $x_pos + ($window_rect.left - $frame_rect.left), 0, $target_width + $border_width , $target_height + $border_height, $true)

echo ("Client left is " + $client_rect.left + " right is " + $client_rect.right)
echo ("window left is " + $window_rect.left + " right is " + $window_rect.right)
echo ("frame left is " + $frame_rect.left + " right is " + $frame_rect.right)