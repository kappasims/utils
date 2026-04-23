# utils

Personal collection of Windows utilities.

## GameOrganizer

A Windows GUI for managing Steam and GOG game installs: archive games behind
directory junctions, toggle them between active and archived, and apply NTFS
compression (LZX / XPRESS) or run a copy-cycle defrag — all from a single
DataGridView.

Lives under [`GameOrganizer/`](GameOrganizer/).

### Running

1. Clone or copy the repo anywhere on a Windows machine.
2. Keep `GameOrganizer/GameOrganizer.ps1` and `GameOrganizer/GameOrganizer.bat`
   **side-by-side** in the same folder (the launcher resolves the script with
   `%~dp0`).
3. Double-click `GameOrganizer/GameOrganizer.bat` to start the GUI.

`GameOrganizer.bat` launches PowerShell with:

```
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0GameOrganizer.ps1"
```

`-ExecutionPolicy Bypass` is scoped to that single invocation and does not
change your machine-wide policy, so there is nothing to pre-configure.

### Configuration

Open `GameOrganizer/GameOrganizer.ps1` and edit the `CONFIGURATION` block at
the top:

| Variable          | Purpose                                                         |
|-------------------|-----------------------------------------------------------------|
| `$steamPath`      | Your Steam library's `steamapps\common` folder                  |
| `$gogPath`        | Your GOG Games folder                                           |
| `$archivedPath`   | Where archived games are moved (junctions point here)           |
| `$defragTempPath` | Scratch folder for the defrag copy-cycle                        |
| `$markerFileName` | Per-game hidden marker file recording the active compression    |

Keep `$archivedPath` on the **same volume** as your library paths — toggling
state is a rename when they share a volume, and a full copy when they don't.
The GUI warns you at startup if this is misconfigured.

### Permissions

The app uses features that normally require elevation:

- **Directory junctions** (`mklink /J` via `cmd /c`) — used by the
  Active/Archived toggle. Creating or removing a junction in a system-managed
  folder usually needs Administrator.
- **NTFS compression** (`compact.exe /exe:lzx` and friends) — LZX and the
  XPRESS variants are only available on NTFS volumes on Windows 10 / 11.
- **Robocopy during defrag** — no elevation needed on user-writable paths,
  but it does consume disk I/O and SSD writes.

Recommended: right-click `GameOrganizer.bat` → **Run as administrator** (or
pin a shortcut with "Run as administrator" checked in Properties → Advanced).

### Features

- **Active / Archived toggle** — moves a game directory to `$archivedPath`
  and replaces the original path with a junction, so Steam/GOG still see it.
  Toggles back by removing the junction and moving the data home.
- **Cleanup detection** — rows highlight "Orphaned Archive" (archive exists
  but no junction) and "Broken Link" (junction points at nothing) with a
  one-click fix.
- **Compression** — per-row dropdown for `LZX`, `XPRESS16K/8K/4K`, or
  `None (decompress)`. Runs `compact.exe` in a background runspace and
  streams per-file progress back into the status bar.
- **Defrag** — copies the game to `$defragTempPath` and back via robocopy,
  giving you a contiguous layout on the destination drive.
- **Locking overlay** — a faded spinner covers the grid during any refresh
  so you cannot click a row mid-scan.

### Requirements

- Windows 10 or 11
- NTFS volumes (required for compression and junctions)
- Windows PowerShell 5.1 (ships with Windows) — not tested on PowerShell 7+
