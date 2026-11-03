# Citrix KeepAlive

Citrix KeepAlive is a Windows PowerShell tray application that attempts to keep Citrix Workspace sessions active by sending a configured key through the Citrix ICA COM interface. It processes each session returned by the local Citrix client.

It does not change server-side timeout policies or guarantee that sessions remain connected.

## Requirements

- Windows with Windows PowerShell 5.1 (x86/x64) and Windows Forms. Native ARM64 has not been validated.
- Citrix Workspace with a registered ICA COM server and an adjacent `WFICALib.dll`.
- The script, a valid `config.json` and an accessible icon file.
- Permission to configure the machine registry values if they are not already set.

Run on the computer running Citrix Workspace, under the same Windows user as the Citrix session. Start normally; the script requests elevation for registry setup when needed. Windows PowerShell is the application runtime, not PowerShell 7.

## Installation

Download these files from this directory:

| File | Purpose |
|---|---|
| [CitrixKeepAlive.ps1](CitrixKeepAlive.ps1) | Application |
| [config.json](config.json) | Required configuration |
| [caf.ico](caf.ico) | Tray icon |
| [CitrixKeepAlive.Tests.ps1](CitrixKeepAlive.Tests.ps1) | Optional development checks |

For the examples below, put the first three files in `C:\ProgramData\Citrix`. The supplied configuration already uses that icon location:

```json
{
    "Interval": 15000,
    "Keystroke": 126,
    "LogLevel": "Info",
    "IconPath": "C:\\ProgramData\\Citrix\\caf.ico"
}
```

Another folder is fine: adjust the shortcut and `IconPath`. The script always reads `config.json` from its own directory, regardless of the working directory.

### Configuration

| Setting | Meaning |
|---|---|
| `Interval` | Delay in milliseconds after processing all sessions. Default configuration: 15000 (15 seconds). Processing time is additional. Use a positive integer. |
| `Keystroke` | Integer key code passed to the Citrix keyboard API. Default configuration: 126. Each key-down is followed by key-up after approximately 100 ms. |
| `LogLevel` | Minimum level: `Info`, `Warning` or `Error`. Missing/unrecognized levels fall back to `Info`. |
| `IconPath` | Icon file path. Use an absolute path so it also works after relaunch or Restart. |

Supply all four settings. JSON parsing is checked, but the script does not perform complete schema or range validation. Settings are read only at startup; use **Restart** after changes.

Relative icon paths depend on the process working directory. They are not resolved relative to the script, and environment variables embedded in the configured string are not expanded.

## Starting the application

### Windows shortcut

Create a shortcut with the following properties.

**Target / Doel:**

```text
%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\ProgramData\Citrix\CitrixKeepAlive.ps1"
```

**Start in / Beginnen in:**

```text
"C:\ProgramData\Citrix"
```

Optionally choose `C:\ProgramData\Citrix\caf.ico` as the shortcut icon. Leave **Run as administrator** unchecked.

The working directory is the folder, without the script filename. Configuration lookup does not depend on it. Use an absolute icon path because restarted processes do not explicitly preserve the shortcut's working directory.

Separate x86/x64 shortcuts are not required: the script selects matching Windows PowerShell when needed. Existing 64-bit shortcuts remain valid.

### Command Prompt

```cmd
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\ProgramData\Citrix\CitrixKeepAlive.ps1"
```

For startup diagnostics, add `-NoExit` and remove `-WindowStyle Hidden` when running from Command Prompt.

The command requests execution-policy Bypass for that process. It does not change the saved user or machine execution policy or guarantee execution under centrally enforced restrictions. No permanent execution-policy change is part of installation.

### Start at sign-in

Press **Win + R**, enter `shell:startup`, and copy the shortcut into that folder. A batch file is not required. KeepAlive starts when that Windows user signs in.

## Tray menu

The normal tooltip is **Caffeine for Citrix Workspace**. Right-click the icon:

| Item | Behavior |
|---|---|
| **Pause** | Pauses subsequent keepalive cycles. A cycle already in progress may finish. |
| **Resume** | Resumes after a pause. |
| **Open Log File** | Opens the current log in its associated application. |
| **Clear Log** | Empties the current log if present. Logging continues, so new entries may appear immediately. The backup log and pause state are unchanged. |
| **Restart** | Stops the background job, releases resources and launches a fresh instance using the current Windows PowerShell executable. Reloads configuration and Citrix detection. The previous paused state is not retained. |
| **Exit** | Stops the background job and closes KeepAlive; it does not sign out of Citrix. |

Restart remains available after a fatal background-job failure. Clear Log does not recover a failed job. A named mutex prevents a second active instance within the Windows session.

Restart does not download updates: replace the local script first to load a newer version.

## Detection and registry setup

1. Read ICA CLSID `{238F6F83-B8B4-11CF-8771-00A024541EE3}` under `HKEY_CLASSES_ROOT\CLSID`, subkey `InprocServer32`, in the 32-bit view and, on 64-bit Windows, the 64-bit view.
2. Check that the registered server file and adjacent `WFICALib.dll` exist. Prefer the current architecture when both registrations pass the file checks; otherwise use the available alternative.
3. Relaunch in matching Windows PowerShell if required. A 32-bit process uses `Sysnative` to launch 64-bit PowerShell; the reverse uses `SysWOW64`. A restart guard prevents loops.
4. Check the CCM values in the selected architecture's machine registry view. Request elevation if necessary, using the same Windows PowerShell executable, then verify the values.
5. Load the DLL and create the COM object in the background job. Log successful activation before enumerating sessions.

File checks do not prove COM activation; that happens in the background job. The script no longer scans standard Program Files folders and does not fall back to an unrelated or Legacy DLL.

The required DWORD values are `AllowSimulationAPI = 1` and `AllowLiveMonitoring = 1`:

| Selected client | CCM registry path |
|---|---|
| 64-bit client on 64-bit Windows | `HKLM\SOFTWARE\Citrix\ICA Client\CCM` |
| 32-bit client on 64-bit Windows | `HKLM\SOFTWARE\WOW6432Node\Citrix\ICA Client\CCM` |
| 32-bit client on 32-bit Windows | `HKLM\SOFTWARE\Citrix\ICA Client\CCM` |

Only the selected view is configured. The base ICA Client key must exist; the script can create its CCM subkey. Exit does not remove these values.

`-SetRegistryKeys` and `-ArchitectureRestarted` are internal switches. Do not add them to your shortcut.

## Logs and troubleshooting

Log location: `%TEMP%\CitrixKeepAlive.log`. Before writing an eligible entry, an existing log larger than 5 MB is rotated to `CitrixKeepAlive.log.bak`, replacing the previous backup. Use `LogLevel` set to `Info` for startup/session diagnostics.

| Message or symptom | Explanation / action |
|---|---|
| Missing/invalid configuration; no new log | Configuration is read before logging starts. Check the JSON beside the script and launch visibly to see errors. |
| No usable Citrix ICA COM registration | Inspect the registered paths in the log. Investigate or repair Workspace if registration, OCX or adjacent DLL is missing. |
| Class not registered / `80040154` | Automatic host selection completed, but activation failed. Inspect the selected view/path and investigate the installation. |
| Registry values could not be set | Check elevation, permissions and the base ICA Client key. After setup returns, the main process rechecks for up to 60 seconds. |
| Icon file not found | Check the absolute `IconPath` and make sure `caf.ico` was downloaded. |
| Active sessions: 0 | No sessions were returned. Open a Citrix desktop/application as the same user. If CCM values were just changed, fully sign out of the existing Citrix session and reconnect before checking again. |
| Simulating keepalive for session | An attempt is starting; this is not a success confirmation. Check subsequent warnings and test beyond the usual timeout. |
| Failed to simulate keystroke | That session's attempt failed. Other sessions and later cycles can still be processed. |
| Citrix KeepAlive - stopped | A fatal error or unexpected job completion was detected. A dialog appears and Pause/Resume are disabled. Investigate the log, then Restart. |
| Clear Log error | The file could not be cleared. Check permissions or an application locking the file, then retry. |
| Script ended during setup | The elevated helper logs its own start/end. Later entries may show the main instance continuing normally. |

The tray checks job state every five seconds. A failure to launch the restarted process also produces a dialog.

## Validation

Run the optional checks from this directory:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\CitrixKeepAlive.Tests.ps1
```

They cover syntax, registration selection with temporary files, missing/stale files, host executable paths and fatal job propagation with missing/invalid assemblies. They do not modify the registry, launch the tray UI or test a real Citrix connection.

The automated checks passed under PowerShell 7.4.6 on Linux. Separately, Windows testing confirmed 64-bit COM activation and session enumeration with Citrix client file version 26.10.0.12. The user also confirmed Restart and Clear Log work.

Actual automatic architecture switching, native ARM64 and prolonged idle-timeout behavior have not been confirmed by those checks. Validate on your own Windows/Citrix installation.

## Version history

- **v2.3** — Restart and Clear Log tray actions.
- **v2.2** — COM-based detection, automatic host architecture selection, matching-view registry setup and visible fatal job errors.
- **v2.1** — Added ProgramW6432 to the former path-based detection; superseded by v2.2.
- **v2.0** — Cross-process pause/resume, main-thread error dialogs, expanded path checks, log-level filtering, elevation path correction and log rotation.
- **v1.0** — Initial release.

## Contributing

Use this repository's issues and pull requests. Include relevant log entries and Citrix/PowerShell versions when reporting problems.

## License

The previous README identified the project as MIT-licensed, but no LICENSE file containing the license text is currently included in this repository.
