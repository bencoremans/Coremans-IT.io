# Citrix KeepAlive Script

## Overview

This PowerShell script keeps Citrix Workspace sessions alive by simulating user activity. It ensures that your Citrix sessions do not time out due to inactivity, which can be especially helpful during long-running tasks or when working remotely.

The script performs the following actions:

1. **Citrix ICA COM Check**: Reads both registry views and validates the registered OCX and adjacent `WFICALib.dll`. Automatically restarts in matching Windows PowerShell when necessary.
2. **Registry Key Verification and Setting**: Checks for specific registry keys required for session simulation and sets them if necessary.
3. **Elevation for Registry Changes**: If registry keys are missing, starts an elevated process to set them.
4. **Session Simulation**: Simulates keystrokes in active Citrix sessions to prevent them from timing out.
5. **System Tray Icon**: Provides a system tray icon for user interaction, allowing pause, resume, and exit actions.
6. **Logging**: Implements logging with configurable log levels. The log file is automatically rotated when it exceeds 5 MB (renamed to `.log.bak`).
7. **Single Instance Enforcement**: Ensures only one instance of the script runs at a time using a mutex.

## Requirements

- **Operating System**: Windows with Windows PowerShell 5.1 (x86/x64). Native ARM64 has not been validated.
- **Citrix Workspace App**: Ensure that the Citrix Workspace app is installed and that the `WfIcaLib.dll` file is accessible.
- **Administrative Privileges**: Required to set necessary registry keys (the script will elevate privileges when needed).
- **Configuration File**: A `config.json` file with appropriate settings (see [Configuration](#configuration)).

## Installation

### 1. Download the Script

Download `CitrixKeepAlive.ps1` and save it to a desired location on your computer.

### 2. Create a Configuration File

In the same directory as the script, create a `config.json` file with the following content:

```json
{
    "Interval": 15000,
    "Keystroke": 126,
    "LogLevel": "Info",
    "IconPath": "C:\\Path\\To\\Your\\Icon.ico"
}
```

- **Interval**: The interval in milliseconds between simulated keystrokes (e.g., `15000` for 15 seconds).
- **Keystroke**: The virtual key code of the keystroke to simulate (e.g., `126`).
- **LogLevel**: The minimum level to log — `"Info"`, `"Warning"`, or `"Error"`. Messages below the configured level are suppressed.
- **IconPath**: The path to the icon file to use for the system tray icon. Replace `"C:\\Path\\To\\Your\\Icon.ico"` with the actual path to your icon file.

### 3. Adjust Execution Policy (If Necessary)

The script uses `-ExecutionPolicy Bypass` to run even if the execution policy is restricted. If you encounter issues, you may need to adjust your PowerShell execution policy:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Usage

### Running the Script

Start from Command Prompt or your existing shortcut:

```cmd
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\To\CitrixKeepAlive.ps1"
```

Replace the script path with your actual location. The script now reads ICA CLSID `{238F6F83-B8B4-11CF-8771-00A024541EE3}` in both registry views and automatically restarts in matching Windows PowerShell if needed. Existing 64-bit shortcuts remain valid.

The registered OCX and adjacent `WFICALib.dll` must both exist. If both architectures are usable, the current architecture is retained. Missing or stale registrations produce a repair message. The script does not use an unrelated DLL from another installation or the Legacy folder as a fallback.

A 32-bit process uses `Sysnative` to launch 64-bit Windows PowerShell; the reverse direction uses `SysWOW64`. A restart guard prevents loops. The script records the selected registration and host in its log. Fatal background-job errors display a dialog, change the tray status to stopped and disable Pause/Resume.

Run KeepAlive as the same Windows user as Citrix, without manually elevating the main process. The script requests elevation for registry setup when needed.

### Running the Script at Startup (Optional)

To have the script run automatically when you log in:

1. **Create a Batch File**

   Create a batch file named `StartCitrixKeepAlive.bat` with the following content:

   ```cmd
   @echo off
   "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\To\CitrixKeepAlive.ps1"
   ```

   Replace `"C:\Path\To\CitrixKeepAlive.ps1"` with the actual path to your script.

2. **Add to Startup Folder**

   - Press `Win + R`, type `shell:startup`, and press **Enter** to open the Startup folder.
   - Copy your `StartCitrixKeepAlive.bat` file into this folder.

### Interacting with the Script

- **System Tray Icon**: The script runs silently in the background with a system tray icon for interaction.
- **Context Menu Options**:
  - **Pause**: Temporarily pause the keep-alive functionality.
  - **Resume**: Resume the keep-alive functionality if paused.
  - **Open Log File**: Open the log file for review.
  - **Exit**: Exit the script.

## Configuration

The `config.json` file must be in the same directory as the script and include the following settings:

- **Interval**: (Integer) The interval in milliseconds between simulated keystrokes.
- **Keystroke**: (Integer) The virtual key code of the keystroke to simulate.
- **LogLevel**: (String) The minimum log level to record — `"Info"`, `"Warning"`, or `"Error"`.
- **IconPath**: (String) The file path to the icon used for the system tray icon.

## Important Notes

- **Administrative Privileges**: The script may prompt for administrative privileges to set registry keys. Accept the prompt to allow the script to configure the necessary settings.
- **Citrix Workspace Installation**: Detection uses the registered ICA COM server location, including custom installation directories.
- **Registry Keys**: Only the view matching the selected Citrix COM server is configured. On 64-bit Windows, 32-bit Citrix uses the first path below and 64-bit Citrix the second. Elevation uses the same Windows PowerShell executable as the main process:

  ```registry
  [HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Citrix\ICA Client\CCM]
  "AllowSimulationAPI"=dword:00000001
  "AllowLiveMonitoring"=dword:00000001

  [HKEY_LOCAL_MACHINE\SOFTWARE\Citrix\ICA Client\CCM]
  "AllowSimulationAPI"=dword:00000001
  "AllowLiveMonitoring"=dword:00000001
  ```

- **Logging**: The script logs its actions to `%TEMP%\CitrixKeepAlive.log`. When the log file exceeds 5 MB, it is automatically renamed to `CitrixKeepAlive.log.bak` and a new log file is started.
- **Script Execution Policy**: Uses `-ExecutionPolicy Bypass` to ensure it runs even if the execution policy is restricted.
- **Single Instance Enforcement**: Ensures only one instance runs at a time using a mutex.
- **Console Window Visibility**: The script runs with the console window hidden. If the console window appears, ensure you're using the `-WindowStyle Hidden` parameter.

## Troubleshooting

- **Console Window Visible**: If the PowerShell console window remains visible, make sure to include the `-WindowStyle Hidden` parameter when starting the script.
- **Elevated Process Console Window**: A console window may briefly appear when the script elevates to set registry keys. This is normal and should close quickly.
- **Script Does Not Start**: Ensure that the script and the `config.json` file are in the same directory, and that the paths specified are correct.
- **No System Tray Icon**: Verify that the `IconPath` in your `config.json` is correct and points to a valid `.ico` file.
- **No Usable Citrix ICA COM Registration**: Inspect the paths in the log. Repair Citrix Workspace if the registration is absent or the registered OCX or adjacent `WFICALib.dll` is missing.
- **Pause/Resume Not Working**: Pause and Resume use a named kernel EventWaitHandle (`CitrixKeepAlive_Pause`) to communicate between the main thread and the background job. Ensure no other instance of the script is running that may have claimed the same handle name.

- **COM Error `80040154` / Class Not Registered**: Check the ICA class `{238F6F83-B8B4-11CF-8771-00A024541EE3}` under `HKEY_CLASSES_ROOT\CLSID` in both registry views. Host architecture is now selected automatically. If activation still fails, inspect the selected paths and host in the log and repair the Citrix installation. A present `WFICALib.dll` does not guarantee that the COM class is registered for the running process.
- **Active sessions: 0**: No sessions were returned by the Citrix API. Open a Citrix desktop or application. If a session was already open when the CCM registry values were changed, fully sign out of that Citrix session and start it again, then check the log. `Simulating keepalive for session` logs an attempt; check for subsequent warnings and verify that the session remains active beyond its usual idle timeout.

## Validation

Run `powershell.exe -NoProfile -File .\CitrixKeepAlive.Tests.ps1` from this directory. The dependency-free checks parse the script, test architecture selection using temporary files, cover missing/stale files and host paths, and verify that missing or invalid interop assemblies fail the actual background job. They do not modify the registry or launch the tray UI.

Live Windows/Citrix validation is still required: launch from both x86 and x64 PowerShell, confirm the selected registration and successful COM activation in the log, then check session enumeration and keepalive behavior. Also check elevation and tray error reporting on Windows.

## Version History

- **v2.2** — COM-registration-based detection, automatic Windows PowerShell architecture selection with restart guard, registry setup in the selected view, and visible fatal background-job errors.
- **v2.1** — Include `ProgramW6432` in client detection, ignore missing/duplicate program folders, and document matching PowerShell architecture to Citrix COM registration, including 64-bit startup and troubleshooting.
- **v2.0** — Fixes: pause/resume via named EventWaitHandle (cross-process), MessageBox from main thread only, dual `ProgramFiles` path check (x86 + x64), LogLevel filtering, `Ensure-Admin` path fix using `$PSHOME`, log rotation at 5 MB.
- **v1.0** — Initial release.

## Contributing

Contributions are welcome! If you have suggestions for improvements or encounter any issues, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).

## Disclaimer

This script is provided as-is without any warranty. Use at your own risk. Ensure you understand what the script does before running it, and back up any important data.
