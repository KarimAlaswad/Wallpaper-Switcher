$wallpaperExe = "D:\Downloads\wallpaper_engine\wallpaper64.exe"
$profileMinimized = "All"
$profileActive = "Dark"
$pollInterval = 1
$debug = $false
$stabilityCount = 2

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

public class WindowChecker {
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    [DllImport("user32.dll")]
    public static extern IntPtr GetShellWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct WINDOWPLACEMENT {
        public int length;
        public int flags;
        public int showCmd;
        public System.Drawing.Point ptMinPosition;
        public System.Drawing.Point ptMaxPosition;
        public System.Drawing.Rectangle rcNormalPosition;
    }

    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;
    public const int WS_VISIBLE = 0x10000000;
    public const int WS_EX_TOOLWINDOW = 0x00000080;
    public const int WS_EX_APPWINDOW = 0x00040000;
    public const int SW_MINIMIZE = 6;
    public const int SW_SHOWMINIMIZED = 2;
    public const int SW_SHOWMAXIMIZED = 3;
    public const int SW_SHOWNORMAL = 1;
    public const int SW_HIDE = 0;

    public static bool IsAnyWindowNotMinimized() {
        IntPtr shellWindow = GetShellWindow();
        IntPtr progman = FindWindow("Progman", null);
        IntPtr workerW = FindWindow("WorkerW", null);
        bool[] hasVisibleNonMinimized = { false };
        string[] excludeTitles = { 
            "Windows Input Experience", 
            "Microsoft Text Input", 
            "Touch Keyboard", 
            "Task Switching",
            "Taskbar",
            "Cortana",
            "Start",
            "Program Manager"
        };
        
        string[] excludeClasses = {
            "Shell_TrayWnd",
            "Shell_SecondaryTrayWnd",
            "Windows.UI.Core.CoreWindow",
            "ApplicationFrameWindow",
            "Windows.UI.Shell.WindowManagerProxy",
            "Progman",
            "WorkerW",
            "DV2ControlHost",
            "TaskbarThumbnailWnd",
            "ThumbnailWnd"
        };
        
        EnumWindows((hWnd, lParam) => {
            if (hWnd == IntPtr.Zero) return true;
            if (hWnd == shellWindow) return true;
            if (hWnd == progman) return true;
            if (hWnd == workerW) return true;

            int length = GetWindowTextLength(hWnd);
            if (length == 0) return true;

            int style = GetWindowLong(hWnd, GWL_STYLE);
            int exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);

            if ((style & WS_VISIBLE) == 0) return true;
            if ((exStyle & WS_EX_TOOLWINDOW) != 0) return true;

            StringBuilder className = new StringBuilder(256);
            GetClassName(hWnd, className, className.Capacity);
            string classStr = className.ToString();

            foreach (string ex in excludeClasses) {
                if (classStr == ex) return true;
            }

            StringBuilder sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            string title = sb.ToString();

            foreach (string ex in excludeTitles) {
                if (title.Contains(ex)) return true;
            }

            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(typeof(WINDOWPLACEMENT));
            GetWindowPlacement(hWnd, ref placement);

            if (placement.showCmd == SW_SHOWNORMAL || placement.showCmd == SW_SHOWMAXIMIZED) {
                hasVisibleNonMinimized[0] = true;
                return false;
            }

            return true;
        }, IntPtr.Zero);

        return hasVisibleNonMinimized[0];
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
}
"@ -ReferencedAssemblies System.Drawing

$lastState = $null
$stableCount = 0
$currentProfile = $null
$pendingProfile = $null
$pendingSince = $null
$activeDelaySeconds = 5

Write-Host "Wallpaper Profile Switcher Started"
Write-Host "Monitoring window states every $pollInterval second(s)..."
Write-Host "Minimized profile: $profileMinimized"
Write-Host "Active profile: $profileActive"
Write-Host ""

while ($true) {
    $anyWindowActive = [WindowChecker]::IsAnyWindowNotMinimized()
    $newState = if ($anyWindowActive) { "active" } else { "minimized" }

    if ($pendingProfile -ne $null) {
        $pendingState = if ($pendingProfile -eq $profileActive) { "active" } else { "minimized" }
        
        $delay = $activeDelaySeconds
        
        if ($newState -ne $pendingState) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') Pending cancelled - state changed to $newState"
            $pendingProfile = $null
            $pendingSince = $null
        } elseif ($pendingSince -ne $null -and ((Get-Date) - $pendingSince).TotalSeconds -ge $delay) {
            if ($pendingProfile -ne $currentProfile) {
                Write-Host "$(Get-Date -Format 'HH:mm:ss') Delay complete: Applying profile $pendingProfile"
                & $wallpaperExe -control openProfile -profile $pendingProfile
                $currentProfile = $pendingProfile
            }
            $pendingProfile = $null
            $pendingSince = $null
        }
    }

    if ($newState -ne $lastState) {
        $stableCount = 1
        $lastState = $newState
    } else {
        $stableCount++
    }

    if ($stableCount -ge $stabilityCount -and $pendingProfile -eq $null) {
        $desiredProfile = if ($newState -eq "active") { $profileActive } else { $profileMinimized }
        
        if ($desiredProfile -ne $currentProfile) {
            $pendingProfile = $desiredProfile
            $pendingSince = Get-Date
            Write-Host "$(Get-Date -Format 'HH:mm:ss') State stable: $newState → Waiting $activeDelaySeconds seconds before applying $desiredProfile..."
        }
        
        $stableCount = 0
    }

    Start-Sleep -Seconds $pollInterval
}
