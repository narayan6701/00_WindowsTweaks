@echo off

setlocal

set registryPath="HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
for /f "tokens=2*" %%A in ('reg query %registryPath% /v AppsUseLightTheme 2^>nul') do set AppsUseLightTheme=%%B
for /f "tokens=2*" %%A in ('reg query %registryPath% /v SystemUsesLightTheme 2^>nul') do set SystemUsesLightTheme=%%B

if "%AppsUseLightTheme%"=="0x1" (
    set newAppsUseLightTheme=0
    set newSystemUsesLightTheme=0
) else (
    set newAppsUseLightTheme=1
    set newSystemUsesLightTheme=1
)

reg add %registryPath% /v AppsUseLightTheme /t REG_DWORD /d %newAppsUseLightTheme% /f
reg add %registryPath% /v SystemUsesLightTheme /t REG_DWORD /d %newSystemUsesLightTheme% /f

taskkill /f /im explorer.exe
start explorer.exe

endlocal
@echo on