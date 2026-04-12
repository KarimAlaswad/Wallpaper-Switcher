$wallpaperExe = "D:\Downloads\wallpaper_engine\wallpaper64.exe"
$profileMinimized = "All"
$profileActive = "Dark"
$pollInterval = 1
$debug = $false
$stabilityCount = 2

# Configuration for monitored apps
$monitoredApps = @("zen.exe", "Code.exe", "Discord.exe", "Files.exe", "Everything.exe", "qbittorrent.exe", "MRA.exe", "WindowsTerminal.exe", "explorer.exe")
$excludeClasses = @("MozillaDialogClass", "Progman")

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

public class WindowChecker {
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromWindow(IntPtr hWnd, uint dwFlags);

    [DllImport("user32.dll")]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

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
        public Point ptMinPosition;
        public Point ptMaxPosition;
        public Rectangle rcNormalPosition;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MONITORINFO {
        public int cbSize;
        public Rectangle rcMonitor;
        public Rectangle rcWork;
        public int dwFlags;
    }

    public const int MONITOR_DEFAULTTONULL = 0;
    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;
    public const int WS_VISIBLE = 0x10000000;
    public const int WS_EX_TOOLWINDOW = 0x00000080;
    public const int SW_SHOWNORMAL = 1;
    public const int SW_SHOWMAXIMIZED = 3;

    public static bool IsAnyMonitoredWindowVisible(string[] monitoredApps, string[] excludeClasses) {
        IntPtr shellWindow = GetShellWindow();
        bool[] found = { false };
        Rectangle primaryBounds = Screen.PrimaryScreen.Bounds;

        EnumWindows((hWnd, lParam) => {
            if (hWnd == IntPtr.Zero || hWnd == shellWindow) return true;
            if ((GetWindowLong(hWnd, GWL_STYLE) & WS_VISIBLE) == 0) return true;
            if ((GetWindowLong(hWnd, GWL_EXSTYLE) & WS_EX_TOOLWINDOW) != 0) return true;

            int pid;
            GetWindowThreadProcessId(hWnd, out pid);
            string processName = "";
            try { processName = Process.GetProcessById(pid).ProcessName + ".exe"; } catch { }
            
            StringBuilder className = new StringBuilder(256);
            GetClassName(hWnd, className, className.Capacity);
            string classStr = className.ToString();

            // Handle explorer.exe restriction: only CabinetWClass
            if (processName.Equals("explorer.exe", StringComparison.OrdinalIgnoreCase)) {
                if (!classStr.Equals("CabinetWClass", StringComparison.OrdinalIgnoreCase)) return true;
            } else {
                bool isMonitored = false;
                foreach (string app in monitoredApps) { if (processName.Equals(app, StringComparison.OrdinalIgnoreCase)) isMonitored = true; }
                if (!isMonitored) return true;
            }

            foreach (string ex in excludeClasses) { if (classStr == ex) return true; }

            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(typeof(WINDOWPLACEMENT));
            GetWindowPlacement(hWnd, ref placement);
            if (placement.showCmd != SW_SHOWNORMAL && placement.showCmd != SW_SHOWMAXIMIZED) return true;

            IntPtr hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONULL);
            if (hMonitor != IntPtr.Zero) {
                MONITORINFO mi = new MONITORINFO();
                mi.cbSize = Marshal.SizeOf(typeof(MONITORINFO));
                if (GetMonitorInfo(hMonitor, ref mi) && mi.rcMonitor == primaryBounds) {
                    found[0] = true;
                    return false;
                }
            }

            return true;
        }, IntPtr.Zero);

        return found[0];
    }
}
"@ -ReferencedAssemblies System.Drawing, System.Windows.Forms

$lastState = $null
$stableCount = 0
$currentProfile = $null
$pendingProfile = $null
$pendingSince = $null
$activeDelaySeconds = 3

while ($true) {
    $anyWindowActive = [WindowChecker]::IsAnyMonitoredWindowVisible($monitoredApps, $excludeClasses)
    $newState = if ($anyWindowActive) { "active" } else { "minimized" }

    if ($pendingProfile -ne $null) {
        $pendingState = if ($pendingProfile -eq $profileActive) { "active" } else { "minimized" }
        
        if ($newState -ne $pendingState) {
            Write-Host "Cancelled: $newState"
            $pendingProfile = $null
            $pendingSince = $null
        } elseif ($pendingSince -ne $null -and ((Get-Date) - $pendingSince).TotalSeconds -ge $activeDelaySeconds) {
            if ($pendingProfile -ne $currentProfile) {
                Write-Host "Applied: $pendingProfile"
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
            Write-Host "Stable: $newState -> Wait $activeDelaySeconds s for $desiredProfile"
        }
        
        $stableCount = 0
    }

    Start-Sleep -Seconds $pollInterval
}
