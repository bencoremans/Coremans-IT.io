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

Version history:
  v2.0 - Fixes: pause/resume via named EventWaitHandle, MessageBox from main thread only,
         dual ProgramFiles path check (x86 + x64), LogLevel filtering, Ensure-Admin path fix,
         log rotation at 5 MB.
#>

param (
    [switch]$SetRegistryKeys
)

# Add necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ─────────────────────────────────────────────────────────────────────────────
# FIX #1: Pause/Resume via named EventWaitHandle (cross-process/job shared state)
# $global:keepAlivePaused did not work across the job boundary because the main
# thread and the background job each have their own memory space. A named
# EventWaitHandle is a kernel object shared between both by name.
# ─────────────────────────────────────────────────────────────────────────────
$pauseEventName = "CitrixKeepAlive_Pause"
$pauseEvent = [System.Threading.EventWaitHandle]::new(
    $false,                                          # initially: not signaled (= active)
    [System.Threading.EventResetMode]::ManualReset,  # must be reset manually
    $pauseEventName
)

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

# ─────────────────────────────────────────────────────────────────────────────
# FIX #4: LogLevel filtering — $LogLevel was read from config but never used.
# Added: levelOrder lookup so that "Warning" also passes "Error", etc.
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param (
        [string]$message,
        [string]$level = "Info",
        [bool]$WriteToConsole = $true
    )
    $allowedLevels = @("Info", "Warning", "Error")
    if ($allowedLevels -contains $level) {

        # Only log if the message level is equal to or higher than the configured level
        $levelOrder = @{ "Info" = 0; "Warning" = 1; "Error" = 2 }
        $configuredLevel = if ($script:LogLevel -and $levelOrder.ContainsKey($script:LogLevel)) {
            $script:LogLevel
        } else {
            "Info"  # fallback if LogLevel is not configured
        }

        if ($levelOrder[$level] -ge $levelOrder[$configuredLevel]) {
            # ─────────────────────────────────────────────────────────────────
            # FIX #6: Log rotation — prevent the log file from growing indefinitely.
            # If the log exceeds 5 MB, rename it to .log.bak and start a new file.
            # ─────────────────────────────────────────────────────────────────
            if (Test-Path $logPath) {
                $logSize = (Get-Item $logPath).Length
                if ($logSize -gt 5MB) {
                    $bakPath = "$logPath.bak"
                    # Overwrite any existing .bak file
                    Move-Item -Path $logPath -Destination $bakPath -Force
                }
            }

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
}

# ─────────────────────────────────────────────────────────────────────────────
# FIX #5: Ensure-Admin used a hardcoded SysWOW64 path.
# On 64-bit systems running a 64-bit PowerShell session, SysWOW64\powershell.exe
# is the 32-bit version — unnecessary and potentially problematic.
# Using $PSHOME ensures the same PowerShell executable is used for re-launch.
# ─────────────────────────────────────────────────────────────────────────────
function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        # Re-launch the script with elevated privileges using the current PowerShell executable
        $cmd = Join-Path $PSHOME "powershell.exe"
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
$Interval   = $config.Interval
$Keystroke  = $config.Keystroke
$LogLevel   = $config.LogLevel   # now actually used in Write-Log (FIX #4)
$IconPath   = $config.IconPath

# Make LogLevel available as a script-scope variable for Write-Log
$script:LogLevel = $LogLevel

# Start logging
$logPath = Join-Path $env:TEMP "CitrixKeepAlive.log"
Write-Log "Script started." "Info"

if ($SetRegistryKeys) {
    Ensure-Admin

    # Define necessary registry keys and values
    $keys = @{
        "AllowSimulationAPI"  = 1
        "AllowLiveMonitoring" = 1
    }

    # Ensure registry keys are set correctly
    Ensure-RegistryKeys -basePath "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client" `
                        -ccmPath  "HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client\CCM" `
                        -keys $keys
    Ensure-RegistryKeys -basePath "HKLM:\Software\Citrix\ICA Client" `
                        -ccmPath  "HKLM:\Software\Citrix\ICA Client\CCM" `
                        -keys $keys

    Write-Log "Registry keys have been set." "Info"
    Write-Log "Script ended." "Info"
    exit
}

# ─────────────────────────────────────────────────────────────────────────────
# FIX #3: Hardcoded ProgramFiles path for wfica32.exe.
# On 64-bit Windows, the Citrix client is typically installed in Program Files (x86),
# not Program Files. Both locations are now checked; the first one found is used.
# ─────────────────────────────────────────────────────────────────────────────
$citrixPaths = @(
    (Join-Path $env:ProgramFiles "Citrix\ICA Client\wfica32.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Citrix\ICA Client\wfica32.exe")
)
$icaClientPath = $citrixPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($icaClientPath) {
    Write-Log "Citrix ICA Client found at: $icaClientPath" "Info"

    # Determine the correct path for WfIcaLib.dll (same directory as wfica32.exe)
    $icaClientDir = Split-Path -Parent $icaClientPath

    $keys = @{
        "AllowSimulationAPI"  = 1
        "AllowLiveMonitoring" = 1
    }

    # Check if the registry settings are correct
    $settingsCorrect = Are-RegistryKeysSet -keys $keys

    if (-not $settingsCorrect) {
        # Registry keys are missing — start an elevated process to set them
        Write-Log "Registry keys are not correctly set. Starting elevated process to set registry keys." "Warning"

        # Start the elevated process to set registry keys
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SetRegistryKeys" -Verb RunAs -Wait

        # Wait for registry keys to be set
        $maxWaitTime  = 60  # seconds
        $waitInterval = 5   # seconds
        $elapsedTime  = 0

        while ($elapsedTime -lt $maxWaitTime) {
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
            # ─────────────────────────────────────────────────────────────────
            # FIX #2: MessageBox shown from the main thread (not the background job).
            # The main thread has WinForms loaded and an STA context, so
            # MessageBox::Show() is safe here. In the job it would fail or hang.
            # ─────────────────────────────────────────────────────────────────
            [System.Windows.Forms.MessageBox]::Show(
                "Registry keys could not be set. Please run the script as administrator to set registry keys.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            Write-Log "Script ended." "Info"
            exit
        }
    }

    # Start Mutex to prevent duplicate execution
    $mutexName = "CaffeineForWorkspaceMutex"
    $mutex     = New-Object -TypeName System.Threading.Mutex($false, $mutexName)
    $hasHandle = $false

    try {
        try {
            $hasHandle = $mutex.WaitOne(0, $false)
        } catch [System.Threading.AbandonedMutexException] {
            $hasHandle = $true
        }

        if (-not $hasHandle) {
            Write-Log "The application is already running." "Warning"
            [System.Windows.Forms.MessageBox]::Show(
                "The application is already running.",
                "Caffeine for Citrix Workspace",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Exclamation
            )
            Write-Log "Script ended." "Info"
            exit
        }

        # Implement system tray icon
        if (-not (Test-Path $IconPath)) {
            Write-Log "Icon file not found at $IconPath. Please check the configuration." "Error"
            [System.Windows.Forms.MessageBox]::Show(
                "Icon file not found. Please check the configuration.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            Write-Log "Script ended." "Info"
            exit
        }

        $trayIcon         = New-Object System.Windows.Forms.NotifyIcon
        $trayIcon.Icon    = [System.Drawing.Icon]::ExtractAssociatedIcon($IconPath)
        $trayIcon.Visible = $true
        $trayIcon.Text    = "Caffeine for Citrix Workspace"

        # Add context menu to the system tray icon
        $contextMenu      = New-Object System.Windows.Forms.ContextMenu
        $exitMenuItem     = New-Object System.Windows.Forms.MenuItem "Exit"
        $pauseMenuItem    = New-Object System.Windows.Forms.MenuItem "Pause"
        $resumeMenuItem   = New-Object System.Windows.Forms.MenuItem "Resume"
        $openLogMenuItem  = New-Object System.Windows.Forms.MenuItem "Open Log File"
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
                $form.Close()
            } catch {
                # Suppress any exceptions
            }
        })

        # ─────────────────────────────────────────────────────────────────────
        # FIX #1 (continued): Pause/Resume via named EventWaitHandle.
        # Instead of $global:keepAlivePaused (which does not work across the job
        # boundary), Set()/Reset() are called on the shared kernel object.
        # The background job calls WaitOne(0) to check whether the event is set.
        # ─────────────────────────────────────────────────────────────────────
        $pauseMenuItem.Add_Click({
            try {
                # Signal the event -> job will pause
                $pauseEvent.Set()
                $pauseMenuItem.Enabled   = $false
                $resumeMenuItem.Enabled  = $true
                Write-Log "Keep-alive function paused by user." "Info" $false
            } catch {
                # Suppress any exceptions
            }
        })

        $resumeMenuItem.Add_Click({
            try {
                # Reset the event -> job resumes the keep-alive loop
                $pauseEvent.Reset()
                $pauseMenuItem.Enabled   = $true
                $resumeMenuItem.Enabled  = $false
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
                    [System.Windows.Forms.MessageBox]::Show(
                        "Log file not found.",
                        "Information",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
                }
            } catch {
                # Suppress any exceptions
            }
        })

        $global:scriptRunning = $true

        # ─────────────────────────────────────────────────────────────────────
        # Start the keep-alive function as a background job.
        # Parameters are passed explicitly; the job runs in its own scope.
        # ─────────────────────────────────────────────────────────────────────
        $keepAliveJob = Start-Job -ScriptBlock {
            param($Interval, $Keystroke, $logPath, $LogLevel, $pauseEventName, $icaClientDir)

            # ─────────────────────────────────────────────────────────────────
            # FIX #4 (in job): Same LogLevel filtering as in the main thread.
            # ─────────────────────────────────────────────────────────────────
            function Write-Log {
                param (
                    [string]$message,
                    [string]$level = "Info",
                    [bool]$WriteToConsole = $true
                )
                $allowedLevels = @("Info", "Warning", "Error")
                if ($allowedLevels -contains $level) {
                    $levelOrder = @{ "Info" = 0; "Warning" = 1; "Error" = 2 }
                    $configuredLevel = if ($LogLevel -and $levelOrder.ContainsKey($LogLevel)) {
                        $LogLevel
                    } else {
                        "Info"
                    }

                    if ($levelOrder[$level] -ge $levelOrder[$configuredLevel]) {
                        # FIX #6 (in job): Log rotation also applies inside the job
                        if (Test-Path $logPath) {
                            $logSize = (Get-Item $logPath).Length
                            if ($logSize -gt 5MB) {
                                $bakPath = "$logPath.bak"
                                Move-Item -Path $logPath -Destination $bakPath -Force
                            }
                        }

                        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        $logEntry  = "[$timestamp] [$level] $message"
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
            }

            function Keep-Alive {
                param (
                    [int]$Interval,
                    [int]$Keystroke
                )

                # ─────────────────────────────────────────────────────────────
                # FIX #3 (in job): WfIcaLib.dll is looked up in icaClientDir
                # (passed as parameter), so both 32-bit and 64-bit paths work.
                # ─────────────────────────────────────────────────────────────
                $wfIcaLibPath = Join-Path $icaClientDir "WfIcaLib.dll"

                if (-not (Test-Path $wfIcaLibPath)) {
                    Write-Log "WfIcaLib.dll not found at $wfIcaLibPath. Please ensure the Citrix ICA Client is properly installed." "Error" $false
                    # ─────────────────────────────────────────────────────────
                    # FIX #2: No MessageBox from the job — it has no UI context.
                    # The error is written to the log file via Write-Log.
                    # The main thread picks this up via job state monitoring.
                    # ─────────────────────────────────────────────────────────
                    return
                }

                try {
                    # Load Citrix ICA Client assembly
                    [System.Reflection.Assembly]::LoadFile($wfIcaLibPath) | Out-Null

                    $ICO = New-Object WFICALib.ICAClientClass
                    $ICO.OutputMode = [WFICALib.OutputMode]::OutputModeNormal

                    # ─────────────────────────────────────────────────────────
                    # FIX #1 (in job): Open the named EventWaitHandle created by
                    # the main thread. WaitOne(0) returns immediately: $true if
                    # the event is signaled (= paused), $false otherwise.
                    # ─────────────────────────────────────────────────────────
                    $jobPauseEvent = [System.Threading.EventWaitHandle]::OpenExisting($pauseEventName)

                    do {
                        # Non-blocking check: are we paused?
                        if ($jobPauseEvent.WaitOne(0)) {
                            Start-Sleep -Seconds 1
                            continue
                        }

                        $EnumHandle  = $ICO.EnumerateCCMSessions()
                        $NumSessions = $ICO.GetEnumNameCount($EnumHandle)

                        Write-Log "Active sessions: $NumSessions" "Info" $false

                        for ($index = 0; $index -lt $NumSessions; $index++) {
                            $sessionid = $ICO.GetEnumNameByIndex($EnumHandle, $index)
                            Write-Log "Simulating keepalive for session: $sessionid" "Info" $false

                            try {
                                $SessionICO = New-Object WFICALib.ICAClientClass
                                $SessionICO.OutputMode = [WFICALib.OutputMode]::OutputModeNormal
                                $SessionICO.StartMonitoringCCMSession($sessionid, $true)

                                $SessionICO.Session.Keyboard.SendKeyDown($Keystroke)
                                Start-Sleep -Milliseconds 100
                                $SessionICO.Session.Keyboard.SendKeyUp($Keystroke)

                                $SessionICO.StopMonitoringCCMSession($sessionid)
                                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($SessionICO)
                            } catch {
                                Write-Log "Failed to simulate keystroke for session $sessionid. Error: $_" "Warning" $false
                            }
                        }

                        [void]$ICO.CloseEnumHandle($EnumHandle)
                        Start-Sleep -Milliseconds $Interval

                    } while ($true)  # Job runs until Stop-Job is called

                } catch {
                    Write-Log "An error occurred in the keep-alive function. Error: $_" "Error" $false
                } finally {
                    if ($ICO) {
                        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ICO)
                    }
                }
            }

            Keep-Alive -Interval $Interval -Keystroke $Keystroke

        } -ArgumentList $Interval, $Keystroke, $logPath, $LogLevel, $pauseEventName, $icaClientDir

        # Create a hidden form to start the message loop
        $form = New-Object System.Windows.Forms.Form
        $form.ShowInTaskbar = $false
        $form.WindowState   = 'Minimized'
        [void]$form.Show()
        [void]$form.Hide()

        # ─────────────────────────────────────────────────────────────────────
        # FIX #2 (continued): Monitor job output from the main thread.
        # Errors from the job are caught here and shown as a MessageBox,
        # because the main thread has STA/UI context and the job does not.
        # ─────────────────────────────────────────────────────────────────────
        $jobMonitorTimer          = New-Object System.Windows.Forms.Timer
        $jobMonitorTimer.Interval = 5000  # check every 5 seconds
        $jobMonitorTimer.Add_Tick({
            if ($keepAliveJob -and $keepAliveJob.State -eq 'Failed') {
                $errorMsg = ($keepAliveJob.ChildJobs[0].JobStateInfo.Reason.Message)
                Write-Log "Background job failed: $errorMsg" "Error" $false
                [System.Windows.Forms.MessageBox]::Show(
                    "The keep-alive job stopped unexpectedly:`n$errorMsg",
                    "Citrix KeepAlive - Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                $jobMonitorTimer.Stop()
            }
        })
        $jobMonitorTimer.Start()

        # Run the application to process Windows messages
        [System.Windows.Forms.Application]::Run($form)

        # Cleanup after form is closed
        $jobMonitorTimer.Stop()
        $jobMonitorTimer.Dispose()

        if ($keepAliveJob) {
            Stop-Job    -Job $keepAliveJob
            Remove-Job  -Job $keepAliveJob
        }

    } finally {
        try {
            if ($hasHandle) {
                $mutex.ReleaseMutex()
            }
            $mutex.Dispose()

            # Clean up the named EventWaitHandle
            $pauseEvent.Close()

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
    Write-Log "Citrix ICA Client is not installed or not found in standard paths. Please install it first." "Error"
    [System.Windows.Forms.MessageBox]::Show(
        "Citrix ICA Client is not installed. Please install it first.",
        "Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    Write-Log "Script ended." "Info"
    exit
}
