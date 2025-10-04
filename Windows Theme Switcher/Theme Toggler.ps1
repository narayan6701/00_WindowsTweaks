# Switch-WindowsTheme.ps1
# This script switches Windows 11 theme based on system time and battery percentage,
# and sends a system broadcast to refresh the UI.

# Function to get battery percentage
function Get-BatteryPercentage {
    $battery = Get-WmiObject -Class Win32_Battery
    if ($battery -ne $null) {
        return $battery.EstimatedChargeRemaining
    } else {
        return $null
    }
}

# Function to set Windows theme: Light (1) or Dark (0)
function Set-WindowsTheme {
    param([int]$ThemeValue)
    $regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
    Set-ItemProperty -Path $regPath -Name 'AppsUseLightTheme' -Value $ThemeValue
    Set-ItemProperty -Path $regPath -Name 'SystemUsesLightTheme' -Value $ThemeValue
}

# Function to broadcast WM_SETTINGCHANGE to all windows
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

    $HWND_BROADCAST = [IntPtr]0xffff
    $WM_SETTINGCHANGE = 0x001A
    $SMTO_ABORTIFHUNG = 0x0002
    $result = [UIntPtr]::Zero

    # Send the message
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

# Main logic
$currentHour = (Get-Date).Hour
$batteryPercent = Get-BatteryPercentage

if (($currentHour -ge 14) -or ($currentHour -lt 7)) {
    Set-WindowsTheme -ThemeValue 0 # Dark Theme
    Write-Output "Switched to Dark Theme (Time: $currentHour)"
} elseif (($currentHour -ge 7) -and ($currentHour -lt 14)) {
    if ($batteryPercent -eq $null -or $batteryPercent -gt 20) {
        Set-WindowsTheme -ThemeValue 1 # Light Theme
        Write-Output "Switched to Light Theme (Time: $currentHour, Battery: $batteryPercent`%)"
    } else {
        Set-WindowsTheme -ThemeValue 0 # Battery low, Dark Theme
        Write-Output "Battery is low ($batteryPercent`%). Stayed in Dark Theme."
    }
} else {
    Set-WindowsTheme -ThemeValue 0 # Fallback, Dark Theme
    Write-Output "Fallback: Switched to Dark Theme."
}

# Broadcast to refresh theme setting
Send-SettingChangeBroadcast

# Note:
# - This broadcast is less disruptive than restarting Explorer.
# - Some apps may not update theme until restarted.