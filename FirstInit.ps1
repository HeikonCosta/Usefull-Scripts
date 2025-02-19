# Set windows on Dark Mode
Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name SystemUsesLightTheme -Value 0 -Type Dword -Force;
Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name AppsUseLightTheme -Value 0 -Type Dword -Force;

# Set Taskbar alignment to Left
$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$Al = "TaskbarAl"
$value = "0"
New-ItemProperty -Path $registryPath -Name $Al -Value $value -PropertyType DWORD -Force -ErrorAction Ignore

# Set windows appearance to best appearance settings
REG ADD HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects /v VisualFXSetting /t REG_DWORD /d 1 /f

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
    $MouseSpeedRegPath = 'hkcu:\Control Panel\Mouse'
    Write-Verbose "MouseSensitivity before WinAPI call: $((Get-ItemProperty $MouseSpeedRegPath).MouseSensitivity)"

    $null = $winApi::SystemParametersInfo($SPI_SETMOUSESPEED, 0, $Value, 0)

    # Calling SystemParametersInfo() does not permanently store the modification
    # of the mouse speed. It needs to be changed in the registry as well
    Set-ItemProperty $MouseSpeedRegPath -Name MouseSensitivity -Value $Value

    Write-Verbose "MouseSensitivity after WinAPI call: $((Get-ItemProperty $MouseSpeedRegPath).MouseSensitivity)"
}

Set-MouseSpeed -Value 5 -Verbose

# Restart Explorer task
Stop-Process -Name explorer -Force; Start-Process explorer

#Set notification config and show it