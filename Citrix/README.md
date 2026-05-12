# Citrix KeepAlive Script

## Overview

This PowerShell script keeps Citrix Workspace sessions alive by simulating user activity. It ensures that your Citrix sessions do not time out due to inactivity, which can be especially helpful during long-running tasks or when working remotely.

The script performs the following actions:

1. **Citrix ICA Client Check**: Verifies if the Citrix ICA Client (Workspace) is installed. Both `Program Files` and `Program Files (x86)` are checked automatically.
2. **Registry Key Verification and Setting**: Checks for specific registry keys required for session simulation and sets them if necessary.
3. **Elevation for Registry Changes**: If registry keys are missing, starts an elevated process to set them.
4. **Session Simulation**: Simulates keystrokes in active Citrix sessions to prevent them from timing out.
5. **System Tray Icon**: Provides a system tray icon for user interaction, allowing pause, resume, and exit actions.
6. **Logging**: Implements logging with configurable log levels. The log file is automatically rotated when it exceeds 5 MB (renamed to `.log.bak`).
7. **Single Instance Enforcement**: Ensures only one instance of the script runs at a time using a mutex.

## Requirements

- **Operating System**: Windows with PowerShell installed.
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

Open Command Prompt (`cmd.exe`) and run the following command:

```cmd
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\To\CitrixKeepAlive.ps1"
```

**Replace** `"C:\Path\To\CitrixKeepAlive.ps1"` with the actual path to the script.

### Running the Script at Startup (Optional)

To have the script run automatically when you log in:

1. **Create a Batch File**

   Create a batch file named `StartCitrixKeepAlive.bat` with the following content:

   ```cmd
   @echo off
   powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\To\CitrixKeepAlive.ps1"
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
- **Citrix Workspace Installation**: The script checks both `Program Files` and `Program Files (x86)` for `wfica32.exe`. Ensure that Citrix Workspace is installed on your system.
- **Registry Keys**: The script checks and sets the following registry keys required for session simulation:

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
- **Citrix ICA Client Not Found**: The script searches for `wfica32.exe` in both `Program Files` and `Program Files (x86)`. If it is still not found, verify that Citrix Workspace is installed correctly and that `wfica32.exe` exists in the Citrix ICA Client directory.
- **Pause/Resume Not Working**: Pause and Resume use a named kernel EventWaitHandle (`CitrixKeepAlive_Pause`) to communicate between the main thread and the background job. Ensure no other instance of the script is running that may have claimed the same handle name.

## Version History

- **v2.0** — Fixes: pause/resume via named EventWaitHandle (cross-process), MessageBox from main thread only, dual `ProgramFiles` path check (x86 + x64), LogLevel filtering, `Ensure-Admin` path fix using `$PSHOME`, log rotation at 5 MB.
- **v1.0** — Initial release.

## Contributing

Contributions are welcome! If you have suggestions for improvements or encounter any issues, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).

## Disclaimer

This script is provided as-is without any warranty. Use at your own risk. Ensure you understand what the script does before running it, and back up any important data.
