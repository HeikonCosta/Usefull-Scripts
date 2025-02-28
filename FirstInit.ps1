# Set windows on Dark Mode
$themePath = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
Set-ItemProperty -Path $themePath -Name 'SystemUsesLightTheme' -Value 0 -Type Dword -Force
Set-ItemProperty -Path $themePath -Name 'AppsUseLightTheme' -Value 0 -Type Dword -Force

# Set Taskbar alignment to Left
$registryPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
$taskbarAlignment = 'TaskbarAl'
Set-ItemProperty -Path $registryPath -Name $taskbarAlignment -Value 0 -Type Dword -Force -ErrorAction Stop

# Set windows appearance to best appearance settings
$visualEffectsPath = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects'
Start-Process -FilePath 'reg' -ArgumentList "ADD $visualEffectsPath /v VisualFXSetting /t REG_DWORD /d 1 /f" -NoNewWindow -Wait

# Set mouse speed to 5
Function Set-MouseSpeed {
    [CmdletBinding()]
    param (
        [ValidateRange(1, 20)]
        [int] $Value
    )

    $winApi = Add-Type -Name user32 -Namespace tq84 -PassThru -MemberDefinition '
        [DllImport("user32.dll")]
        public static extern bool SystemParametersInfo(
            uint uiAction,
            uint uiParam,
            uint pvParam,
            uint fWinIni
        );
    '

    $SPI_SETMOUSESPEED = 0x0071
    $MouseSpeedRegPath = 'HKCU:\Control Panel\Mouse'
    Write-Verbose "MouseSensitivity before WinAPI call: $((Get-ItemProperty $MouseSpeedRegPath).MouseSensitivity)"

    try {
        $null = $winApi::SystemParametersInfo($SPI_SETMOUSESPEED, 0, $Value, 0)
        Set-ItemProperty -Path $MouseSpeedRegPath -Name 'MouseSensitivity' -Value $Value -ErrorAction Stop
        Write-Verbose "MouseSensitivity after WinAPI call: $((Get-ItemProperty $MouseSpeedRegPath).MouseSensitivity)"
    }
    catch {
        Write-Error "Failed to set mouse speed: $_"
    }
}

Set-MouseSpeed -Value 5 -Verbose

# Restart Explorer task
Stop-Process -Name 'explorer' -Force
Start-Process -FilePath 'explorer'

# Set notification config and show it
# (Include additional code here if needed)