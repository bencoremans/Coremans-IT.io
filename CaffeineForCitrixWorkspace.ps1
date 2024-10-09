<#
.SYNOPSIS
Keeps Citrix Workspace sessions alive by simulating user activity and ensures necessary registry keys are set.

.DESCRIPTION
This script performs the following actions:
1. Checks if the Citrix ICA Client (Workspace) is installed.
2. Verifies and sets specific registry keys required for session simulation.
3. If registry keys are missing, starts an elevated process to set them.
4. Waits for the registry keys to be set before proceeding.
5. Simulates keystrokes in active Citrix sessions to prevent them from timing out.
6. Provides a system tray icon for user interaction, allowing pause, resume, and exit actions.
7. Implements logging of actions and registry changes to a log file with different log levels.
8. Ensures only one instance of the script runs at a time using a mutex.

.PARAMETER SetRegistryKeys
Internal switch parameter used to set registry keys with elevated privileges. Do not use this parameter manually.

.EXAMPLE
# Run the script normally to keep Citrix sessions alive.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\CitrixKeepAlive.ps1"

.NOTES
- **Important Notes:**
  - **Administrative Privileges:** The script requires administrative privileges to modify registry settings. It will elevate when necessary.
  - **Citrix Workspace Installation:** Ensure that the Citrix Workspace app is installed on your system, and that the `WfIcaLib.dll` file is accessible.
  - **Log File Location:** The log file is saved in the user's temporary folder (e.g., `%TEMP%\CitrixKeepAlive.log`).
  - **Configuration File:** A `config.json` file must be present in the same directory as the script, containing configuration settings.
  - **Script Execution Policy:** The script uses `-ExecutionPolicy Bypass` to ensure it runs even if the execution policy is restricted.
  - **System Tray Icon:** The script runs with a system tray icon for interaction. Use the icon to pause, resume, or exit the script.
  - **Single Instance Enforcement:** The script ensures only one instance runs at a time.

#>

param (
    [switch]$SetRegistryKeys
)

# Add necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to load settings from configuration file
function Load-Config {
    param (
        [string]$configPath
    )
    if (Test-Path $configPath) {
        try {
            $json = Get-Content $configPath -Raw | ConvertFrom-Json
            return $json
        } catch {
            Write-Error "Error reading the configuration file: $_"
            exit
        }
    } else {
        Write-Error "Configuration file not found at $configPath"
        exit
    }
}

# Function for logging with different levels
function Write-Log {
    param (
        [string]$message,
        [string]$level = "Info",
        [bool]$WriteToConsole = $true
    )
    $allowedLevels = @("Info", "Warning", "Error")
    if ($allowedLevels -contains $level) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] [$level] $message"
        Add-Content -Path $logPath -Value $logEntry
        if ($WriteToConsole) {
            try {
                if ($level -eq "Error") {
                    Write-Error $message
                } elseif ($level -eq "Warning") {
                    Write-Warning $message
                } else {
                    Write-Output $message
                }
            } catch [System.Management.Automation.PipelineStoppedException] {
                # Suppress the exception when the pipeline is stopped
            }
        }
    }
}

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
            New-ItemProperty -Path $path -Name $name -Value $value -PropertyType DWord -Force | Out-Null
            Write-Log "Registry key $name set to $value in $path" "Info"
        } else {
            Write-Log "Registry key $name in $path is already correctly set." "Info"
        }
    } catch {
        Write-Log "An error occurred while setting the registry key $name in $path. Error: $_" "Error"
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
            New-Item -Path $ccmPath -Force | Out-Null
            Write-Log "Registry key $ccmPath created." "Info"
        }
        foreach ($key in $keys.Keys) {
            Set-RegistryKeyIfDifferent -path $ccmPath -name $key -value $keys[$key]
        }
    } else {
        Write-Log "$basePath does not exist; CCM values will not be set." "Warning"
    }
}

# Function to check if registry keys are correctly set
function Are-RegistryKeysSet {
    param (
        [hashtable]$keys
    )
    $settingsCorrect = $true

    if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client") {
        if (-not (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client\CCM")) {
            return $false
        } else {
            $settingsCorrect = $settingsCorrect -and (Check-RegistryKeys -path "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client\CCM" -keys $keys)
        }
    }

    if (Test-Path "HKLM:\Software\Citrix\ICA Client") {
        if (-not (Test-Path "HKLM:\Software\Citrix\ICA Client\CCM")) {
            return $false
        } else {
            $settingsCorrect = $settingsCorrect -and (Check-RegistryKeys -path "HKLM:\Software\Citrix\ICA Client\CCM" -keys $keys)
        }
    }

    return $settingsCorrect
}

# Load configuration
$configPath = Join-Path -Path (Split-Path -Parent $PSCommandPath) -ChildPath "config.json"
$config = Load-Config -configPath $configPath

# Retrieve settings from configuration
$Interval = $config.Interval
$Keystroke = $config.Keystroke
$LogLevel = $config.LogLevel
$IconPath = $config.IconPath

# Start logging
$logPath = Join-Path $env:TEMP "CitrixKeepAlive.log"
Write-Log "Script started." "Info"

if ($SetRegistryKeys) {
    Ensure-Admin

    # Define necessary registry keys and values
    $keys = @{
        "AllowSimulationAPI" = 1
        "AllowLiveMonitoring" = 1
    }

    # Ensure registry keys are set correctly
    Ensure-RegistryKeys -basePath "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client" `
                        -ccmPath "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client\CCM" `
                        -keys $keys
    Ensure-RegistryKeys -basePath "HKLM:\Software\Citrix\ICA Client" `
                        -ccmPath "HKLM:\Software\Citrix\ICA Client\CCM" `
                        -keys $keys

    Write-Log "Registry keys have been set." "Info"
    Write-Log "Script ended." "Info"
    exit
}

# Check if Citrix ICA Client is installed
$icaClientPath = Join-Path $env:ProgramFiles "Citrix\ICA Client\wfica32.exe"

if (Test-Path $icaClientPath) {
    Write-Log "Citrix ICA Client found." "Info"

    $keys = @{
        "AllowSimulationAPI" = 1
        "AllowLiveMonitoring" = 1
    }

    # Check if the registry settings are correct
    $settingsCorrect = Are-RegistryKeysSet -keys $keys

    if (-not $settingsCorrect) {
        # Start a new elevated process to set the registry keys
        Write-Log "Registry keys are not correctly set. Starting elevated process to set registry keys." "Warning"

        # Start the elevated process to set registry keys
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SetRegistryKeys" -Verb RunAs -Wait

        # Wait for registry keys to be set
        $maxWaitTime = 60  # seconds
        $waitInterval = 5  # seconds
        $elapsedTime = 0

        while ($elapsedTime -lt $maxWaitTime) {
            # Re-check the registry settings
            $settingsCorrect = Are-RegistryKeysSet -keys $keys

            if ($settingsCorrect) {
                Write-Log "Registry keys have been set correctly." "Info"
                break
            }

            Start-Sleep -Seconds $waitInterval
            $elapsedTime += $waitInterval
        }

        if (-not $settingsCorrect) {
            Write-Log "Registry keys were not set correctly after waiting. Exiting script." "Error"
            [System.Windows.Forms.MessageBox]::Show("Registry keys could not be set. Please run the script as administrator to set registry keys.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Log "Script ended." "Info"
            exit
        }
    }

    # Start Mutex to prevent duplicate execution
    $mutexName = "CaffeineForWorkspaceMutex"
    $mutex = New-Object -TypeName System.Threading.Mutex($false, $mutexName)
    $hasHandle = $false

    try {
        try {
            $hasHandle = $mutex.WaitOne(0, $false)
        } catch [System.Threading.AbandonedMutexException] {
            $hasHandle = $true
        }

        if (-not $hasHandle) {
            Write-Log "The application is already running." "Warning"
            [System.Windows.Forms.MessageBox]::Show("The application is already running.", "Caffeine for Citrix Workspace", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
            Write-Log "Script ended." "Info"
            exit
        }

        # Implement system tray icon
        if (-not (Test-Path $IconPath)) {
            Write-Log "Icon file not found at $IconPath. Please check the configuration." "Error"
            [System.Windows.Forms.MessageBox]::Show("Icon file not found. Please check the configuration.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Log "Script ended." "Info"
            exit
        }

        $trayIcon = New-Object System.Windows.Forms.NotifyIcon
        $trayIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($IconPath)
        $trayIcon.Visible = $true
        $trayIcon.Text = "Caffeine for Citrix Workspace"

        # Add context menu to the system tray icon
        $contextMenu = New-Object System.Windows.Forms.ContextMenu
        $exitMenuItem = New-Object System.Windows.Forms.MenuItem "Exit"
        $pauseMenuItem = New-Object System.Windows.Forms.MenuItem "Pause"
        $resumeMenuItem = New-Object System.Windows.Forms.MenuItem "Resume"
        $openLogMenuItem = New-Object System.Windows.Forms.MenuItem "Open Log File"
        $resumeMenuItem.Enabled = $false
        $contextMenu.MenuItems.Add($pauseMenuItem)
        $contextMenu.MenuItems.Add($resumeMenuItem)
        $contextMenu.MenuItems.Add($openLogMenuItem)
        $contextMenu.MenuItems.Add($exitMenuItem)
        $trayIcon.ContextMenu = $contextMenu

        # Handle Exit menu item click
        $exitMenuItem.Add_Click({
            try {
                $global:scriptRunning = $false
                $trayIcon.Visible = $false
                $trayIcon.Dispose()
                # Close the form to exit the message loop
                $form.Close()
            } catch {
                # Suppress any exceptions
            }
        })

        # Handle Pause menu item click
        $pauseMenuItem.Add_Click({
            try {
                $global:keepAlivePaused = $true
                $pauseMenuItem.Enabled = $false
                $resumeMenuItem.Enabled = $true
                Write-Log "Keep-alive function paused by user." "Info" $false
            } catch {
                # Suppress any exceptions
            }
        })

        # Handle Resume menu item click
        $resumeMenuItem.Add_Click({
            try {
                $global:keepAlivePaused = $false
                $pauseMenuItem.Enabled = $true
                $resumeMenuItem.Enabled = $false
                Write-Log "Keep-alive function resumed by user." "Info" $false
            } catch {
                # Suppress any exceptions
            }
        })

        # Handle Open Log File menu item click
        $openLogMenuItem.Add_Click({
            try {
                if (Test-Path $logPath) {
                    Invoke-Item $logPath
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Log file not found.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } catch {
                # Suppress any exceptions
            }
        })

        $global:scriptRunning = $true
        $global:keepAlivePaused = $false

        # Start the keep-alive function asynchronously
        $keepAliveJob = Start-Job -ScriptBlock {
            param($Interval, $Keystroke, $logPath, $LogLevel)

            # Set global variables in job scope
            $global:scriptRunning = $true
            $global:keepAlivePaused = $false

            # Define the Write-Log function in the job
            function Write-Log {
                param (
                    [string]$message,
                    [string]$level = "Info",
                    [bool]$WriteToConsole = $true
                )
                $allowedLevels = @("Info", "Warning", "Error")
                if ($allowedLevels -contains $level) {
                    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    $logEntry = "[$timestamp] [$level] $message"
                    Add-Content -Path $logPath -Value $logEntry
                    if ($WriteToConsole) {
                        try {
                            if ($level -eq "Error") {
                                Write-Error $message
                            } elseif ($level -eq "Warning") {
                                Write-Warning $message
                            } else {
                                Write-Output $message
                            }
                        } catch [System.Management.Automation.PipelineStoppedException] {
                            # Suppress the exception when the pipeline is stopped
                        }
                    }
                }
            }

            # Define the Keep-Alive function in the job
            function Keep-Alive {
                param (
                    [int]$Interval,
                    [int]$Keystroke
                )
                $wfIcaLibPath = Join-Path $env:ProgramFiles "Citrix\ICA Client\WfIcaLib.dll"
                if (-not (Test-Path $wfIcaLibPath)) {
                    Write-Log "WfIcaLib.dll not found at $wfIcaLibPath. Please ensure the Citrix ICA Client is properly installed." "Error" $false
                    [System.Windows.Forms.MessageBox]::Show("WfIcaLib.dll not found. Please ensure the Citrix ICA Client is properly installed.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    exit
                }

                try {
                    # Load Citrix ICA Client assembly
                    [System.Reflection.Assembly]::LoadFile($wfIcaLibPath) | Out-Null

                    $ICO = New-Object WFICALib.ICAClientClass
                    $ICO.OutputMode = [WFICALib.OutputMode]::OutputModeNormal

                    do {
                        if ($global:keepAlivePaused) {
                            Start-Sleep -Seconds 1
                            continue
                        }

                        $EnumHandle = $ICO.EnumerateCCMSessions()
                        $NumSessions = $ICO.GetEnumNameCount($EnumHandle)

                        Write-Log "Active sessions: $NumSessions" "Info" $false

                        for ($index = 0; $index -lt $NumSessions; $index++) {
                            $sessionid = $ICO.GetEnumNameByIndex($EnumHandle, $index)
                            Write-Log "Simulating keepalive for session: $sessionid" "Info" $false

                            try {
                                # Create a new ICAClient object for the session
                                $SessionICO = New-Object WFICALib.ICAClientClass
                                $SessionICO.OutputMode = [WFICALib.OutputMode]::OutputModeNormal
                                $SessionICO.StartMonitoringCCMSession($sessionid, $true)

                                # Send the keystroke to the session
                                $SessionICO.Session.Keyboard.SendKeyDown($Keystroke)
                                Start-Sleep -Milliseconds 100
                                $SessionICO.Session.Keyboard.SendKeyUp($Keystroke)

                                # Stop monitoring the session
                                $SessionICO.StopMonitoringCCMSession($sessionid)

                                # Release the session object
                                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($SessionICO)
                            } catch {
                                Write-Log "Failed to simulate keystroke for session $sessionid. Error: $_" "Warning" $false
                            }
                        }

                        [void]$ICO.CloseEnumHandle($EnumHandle)

                        Start-Sleep -Milliseconds $Interval
                    } while ($global:scriptRunning)
                } catch {
                    Write-Log "An error occurred in the keep-alive function. Error: $_" "Error" $false
                } finally {
                    if ($ICO) {
                        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ICO)
                    }
                }
            }

            # Call Keep-Alive function
            Keep-Alive -Interval $Interval -Keystroke $Keystroke
        } -ArgumentList $Interval, $Keystroke, $logPath, $LogLevel

        # Create a hidden form to start the message loop
        $form = New-Object System.Windows.Forms.Form
        $form.ShowInTaskbar = $false
        $form.WindowState = 'Minimized'
        [void]$form.Show()
        [void]$form.Hide()

        # Run the application to process Windows messages
        [System.Windows.Forms.Application]::Run($form)

        # Cleanup code after form is closed
        if ($keepAliveJob) {
            Stop-Job -Job $keepAliveJob
            Remove-Job -Job $keepAliveJob
        }

    }
    finally {
        try {
            if ($hasHandle) {
                $mutex.ReleaseMutex()
            }
            $mutex.Dispose()
            # Ensure the system tray icon is cleaned up upon exit
            if ($trayIcon) {
                $trayIcon.Visible = $false
                $trayIcon.Dispose()
            }
            Write-Log "Script ended." "Info"
        } catch {
            # Suppress any exceptions
        }
    }
} else {
    Write-Log "Citrix ICA Client is not installed. Please install it first." "Error"
    [System.Windows.Forms.MessageBox]::Show("Citrix ICA Client is not installed. Please install it first.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    Write-Log "Script ended." "Info"
    exit
}
