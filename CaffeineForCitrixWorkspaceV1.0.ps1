<#
.SYNOPSIS
This script checks and ensures that specific registry keys for Citrix ICA Client are correctly set and simulates activity to keep Citrix sessions alive.

.DESCRIPTION
The script performs the following actions:
1. Checks if Citrix ICA Client is installed.
2. Verifies if specific registry keys are set correctly. If not, it sets them.
3. If the Citrix ICA Client is installed and running, it simulates key presses to prevent sessions from timing out. By default, it uses the F15 key (key code 126), which typically does not interfere with user activities.
4. The script ensures it runs with administrative privileges, and if not, it restarts itself with elevated rights.
5. A Mutex is used to prevent multiple instances of the script from running simultaneously.
6. Logs all actions and registry changes to a log file located in the user's temporary folder.

.PARAMETER SetRegistryKeys
Switch parameter to set registry keys if they are not correctly configured.

.PARAMETER Interval
Specifies the interval for the keep-alive function in milliseconds. The default is 15000 milliseconds.

.PARAMETER Keystroke
Specifies the key code to simulate during the keep-alive function. The default is 126 (F15 key).

.EXAMPLE
.\CitrixKeepAlive.ps1

Runs the script, checks the registry settings, and simulates activity in Citrix sessions if necessary using the default F15 key.

.EXAMPLE
.\CitrixKeepAlive.ps1 -SetRegistryKeys

Runs the script and ensures that the specified registry keys are set before checking Citrix sessions.

.EXAMPLE
.\CitrixKeepAlive.ps1 -Keystroke 16

Runs the script using the Shift key (key code 16) as the simulated key press for keeping the session alive.

.NOTES
This script requires administrative privileges to modify registry settings.
The log file is saved in the user's temporary folder.
#>

param (
    [switch]$SetRegistryKeys,
    [int]$Interval = 15000,  # Interval for the keep-alive function in milliseconds
    [int]$Keystroke = 126  # Default keystroke is F15 key
)

# Function to ensure the script is running with admin rights
function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        # Restart the script with elevated privileges
        $cmd = "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
        Start-Process $cmd -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SetRegistryKeys" -Verb RunAs -Wait
        exit
    }
}

# Function to check if registry keys are correctly set
function Check-RegistryKeys {
    param (
        [string]$path,
        [hashtable]$keys
    )
    
    foreach ($key in $keys.Keys) {
        $currentValue = Get-ItemProperty -Path $path -Name $key -ErrorAction SilentlyContinue
        if ($currentValue.$key -ne $keys[$key]) {
            return $false
        }
    }
    return $true
}

# Function to set a registry key if the value is different
function Set-RegistryKeyIfDifferent {
    param (
        [string]$path,
        [string]$name,
        [object]$value
    )
    try {
        $currentValue = Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue
        if ($currentValue.$name -ne $value) {
            Set-ItemProperty -Path $path -Name $name -Value $value -Force
            Write-Output "Registry value $name set to $value in $path"
        } else {
            Write-Output "Registry value $name in $path is already correctly set."
        }
    } catch {
        Write-Error "An error occurred while setting the registry value $name in $path : $_"
    }
}

# Function to ensure registry keys are present and correctly set
function Ensure-RegistryKeys {
    param (
        [string]$basePath,
        [string]$ccmPath,
        [hashtable]$keys
    )
    if (Test-Path $basePath) {
        if (-not (Test-Path $ccmPath)) {
            New-Item -Path $ccmPath -Force
            Write-Output "$ccmPath key created."
        }
        foreach ($key in $keys.Keys) {
            Set-RegistryKeyIfDifferent -path $ccmPath -name $key -value $keys[$key]
        }
    } else {
        Write-Output "$basePath does not exist, CCM values will not be set."
    }
}

# Start logging
$logPath = Join-Path $env:TEMP "CitrixRegistryChanges.log"
Start-Transcript -Path $logPath -Append -NoClobber

if ($SetRegistryKeys) {
    Ensure-Admin

    # Define necessary registry keys and values
    $keys = @{
        "AllowSimulationAPI" = 1
        "AllowLiveMonitoring" = 1
    }

    # Ensure registry keys are set correctly
    Ensure-RegistryKeys -basePath "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client" -ccmPath "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client\CCM" -keys $keys
    Ensure-RegistryKeys -basePath "HKLM:\Software\Citrix\ICA Client" -ccmPath "HKLM:\Software\Citrix\ICA Client\CCM" -keys $keys

    Stop-Transcript
    exit
}

# Check if Citrix ICA Client is installed
$icaClientPath = Join-Path $env:ProgramFiles "Citrix\ICA Client\wfica32.exe"

if (Test-Path $icaClientPath) {
    Write-Output "Citrix ICA Client found."

    $keys = @{
        "AllowSimulationAPI" = 1
        "AllowLiveMonitoring" = 1
    }

    # Check if the registry settings are correct
    $settingsCorrect = $true

    if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client") {
        if (-not (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client\CCM")) {
            $settingsCorrect = $false
        } else {
            $settingsCorrect = $settingsCorrect -and (Check-RegistryKeys -path "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client\CCM" -keys $keys)
        }
    }

    if (Test-Path "HKLM:\Software\Citrix\ICA Client") {
        if (-not (Test-Path "HKLM:\Software\Citrix\ICA Client\CCM")) {
            $settingsCorrect = $false
        } else {
            $settingsCorrect = $settingsCorrect -and (Check-RegistryKeys -path "HKLM:\Software\Citrix\ICA Client\CCM" -keys $keys)
        }
    }

    if (-not $settingsCorrect) {
        # Dispose of the mutex before restarting the script
        if ($mutex) {
            $mutex.Dispose()
        }
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SetRegistryKeys" -Verb RunAs -Wait
    }

    # Start Mutex to prevent duplicate execution
    $mutexName = "CaffeineForWorkspaceMutex"
    $mutex = New-Object -TypeName System.Threading.Mutex($true, $mutexName)

    try {
        if (-not $mutex.WaitOne(0, $false)) {
            [System.Windows.Forms.MessageBox]::Show("The application is already running.", "Caffeine for Citrix Workspace", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
            exit
        }

        # Load Citrix ICA Client COM objects
        [System.Reflection.Assembly]::LoadFile("$env:ProgramFiles\Citrix\ICA Client\WfIcaLib.dll") | Out-Null
        $ICO = New-Object WFICALib.ICAClientClass
        $ICO.OutputMode = [WFICALib.OutputMode]::OutputModeNormal

        do {
            $EnumHandle = $ICO.EnumerateCCMSessions()
            $NumSessions = $ICO.GetEnumNameCount($EnumHandle)

            Write-Output "Active sessions: $NumSessions"

            for ($index = 0; $index -lt $NumSessions; $index++) {
                $sessionid = $ICO.GetEnumNameByIndex($EnumHandle, $index)
                Write-Output "Simulating keepalive for session: $sessionid"
                $ICO.StartMonitoringCCMSession($sessionid, $true)
                $ICO.Session.Keyboard.SendKeyDown($Keystroke)  # Simulate the specified key press
                $ICO.StopMonitoringCCMSession($sessionid)
            }

            $ICO.CloseEnumHandle($EnumHandle) | Out-Null
        } until (Start-Sleep -Seconds ($Interval / 1000))
    }
    finally {
        if ($mutex) {
            $mutex.ReleaseMutex()
            $mutex.Dispose()
        }
    }
} else {
    Write-Output "Citrix ICA Client is not installed. Please install it first."
    Stop-Transcript
    exit
}

# Stop logging
Stop-Transcript
