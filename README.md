# utils

Personal collection of Windows utilities.

## gameorganizer

Windows GUI for managing Steam and GOG installs on NTFS.

Lives under [`gameorganizer/`](gameorganizer/).

### Running

Double-click `gameorganizer/gameorganizer.bat`. The launcher is:

```
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0gameorganizer.ps1"
```

`-ExecutionPolicy Bypass` is scoped to that one invocation and does not
change machine policy.

### Configuration

On first run a setup wizard asks for your library, archive, and scratch
paths. Steam, GOG, and defrag are optional — uncheck what you don't use.
Paths are stored in `%APPDATA%\gameorganizer\config.json` and can be
edited from the config bar at the top of the main window. **Clear Saved
Settings** wipes that folder and reopens the wizard.

Keep the archive folder on the same volume as your libraries — toggle
is a rename within a volume and a full copy across volumes.

### What it does

- Toggle games between active (in place) and archived (moved, junction-linked).
- Detect orphan archives and broken junctions.
- Apply NTFS compression (LZX, XPRESS 16K/8K/4K) or decompress via `compact.exe`.
- Defrag by copying a game to a scratch folder and back via robocopy.

### Requirements

- Windows 10 / 11 on NTFS
- Windows PowerShell 5.1
- Administrator for junction creation and LZX compression (right-click
  the `.bat` → Run as administrator)
