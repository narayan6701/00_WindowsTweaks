# ─────────────────────────────────────────────────────────────
# Script: Switch-WindowsTheme.ps1
# Purpose: Automatically switch Windows 11 theme based on system time and battery level.
# Behavior:
#   - Dark Theme: Active between 2:00 PM and 7:00 AM, or if battery is below 20%
#   - Light Theme: Active between 7:00 AM and 2:00 PM, only if battery is above 20%
#   - Avoids unnecessary registry writes
#   - Broadcasts WM_SETTINGCHANGE to refresh UI without restarting Explorer
# ─────────────────────────────────────────────────────────────

# ── Function: Get-BatteryPercentage
# Retrieves the current battery percentage using WMI.
# Returns null if no battery is detected (e.g., desktop PC).
function Get-BatteryPercentage {
    $battery = Get-WmiObject -Class Win32_Battery
    if ($battery -ne $null) {
        return $battery.EstimatedChargeRemaining
    } else {
        return $null
    }
}

# ── Function: Set-WindowsTheme
# Applies the desired theme (Light = 1, Dark = 0) to both system and apps.
# Only updates registry if current values differ to avoid redundant writes.
function Set-WindowsTheme {
    param([int]$ThemeValue)

    # Registry path for theme personalization
    $regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'

    # Get current theme values
    $current = Get-ItemProperty -Path $regPath

    # Apply theme only if values differ
    if ($current.SystemUsesLightTheme -ne $ThemeValue -or $current.AppsUseLightTheme -ne $ThemeValue) {
        Set-ItemProperty -Path $regPath -Name 'AppsUseLightTheme' -Value $ThemeValue
        Set-ItemProperty -Path $regPath -Name 'SystemUsesLightTheme' -Value $ThemeValue
        Write-Output "Theme changed to $ThemeValue"
    } else {
        Write-Output "Theme already set to $ThemeValue. No change made."
    }
}

# ── Function: Send-SettingChangeBroadcast
# Sends a WM_SETTINGCHANGE broadcast to all windows to refresh theme settings.
# This avoids restarting Explorer and ensures UI updates where supported.
function Send-SettingChangeBroadcast {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class NativeMethods {
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    public static extern int SendMessageTimeout(
        IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam,
        uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
}
"@

    # Constants for broadcast
    $HWND_BROADCAST = [IntPtr]0xffff
    $WM_SETTINGCHANGE = 0x001A
    $SMTO_ABORTIFHUNG = 0x0002
    $result = [UIntPtr]::Zero

    # Send broadcast message to all top-level windows
    [NativeMethods]::SendMessageTimeout(
        $HWND_BROADCAST,
        $WM_SETTINGCHANGE,
        [UIntPtr]::Zero,
        "ImmersiveColorSet",
        $SMTO_ABORTIFHUNG,
        100,
        [ref]$result
    ) | Out-Null
}

# ── Main Logic Block ──────────────────────────────────────────

# Get current hour (0–23) to determine time-based theme
$currentHour = (Get-Date).Hour

# Get battery percentage (null if no battery)
$batteryPercent = Get-BatteryPercentage

# ── Decision: Apply Dark Theme
# If time is between 2 PM and 7 AM, or battery is below 20%
if (($currentHour -ge 14) -or ($currentHour -lt 7)) {
    Set-WindowsTheme -ThemeValue 0
    Write-Output "Switched to Dark Theme (Time: $currentHour)"
}

# ── Decision: Apply Light Theme
# If time is between 7 AM and 2 PM and battery is healthy
elseif (($currentHour -ge 7) -and ($currentHour -lt 14)) {
    if ($batteryPercent -eq $null -or $batteryPercent -gt 20) {
        Set-WindowsTheme -ThemeValue 1
        Write-Output "Switched to Light Theme (Time: $currentHour, Battery: $batteryPercent`%)"
    } else {
        # Battery is low, override to Dark Theme
        Set-WindowsTheme -ThemeValue 0
        Write-Output "Battery is low ($batteryPercent`%). Stayed in Dark Theme."
    }
}

# ── Final Step: Broadcast theme change to refresh UI
Send-SettingChangeBroadcast