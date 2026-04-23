# utils

Personal collection of Windows utilities.

## GameOrganizer

Windows GUI for managing Steam and GOG installs on NTFS.

Lives under [`GameOrganizer/`](GameOrganizer/).

### What it does

- Toggle games between active (in place) and archived (moved, junction-linked).
- Detect orphan archives and broken junctions.
- Apply NTFS compression (LZX, XPRESS 16K/8K/4K) or decompress via `compact.exe`. See [About the compression](#about-the-compression).
- Defrag by copying a game to a scratch folder and back via robocopy.

### Requirements

- Windows 10 / 11 on NTFS
- Windows PowerShell 5.1
- Administrator for junction creation and LZX compression (right-click
  the `.bat` -> Run as administrator)

### Prerequisites

`gameorganizer.bat` resolves its script via `%~dp0`, so both files must
sit next to each other. Simplest setup is to clone the repo and run from
the checkout:

```
git clone https://github.com/kappasims/utils.git C:\Tools\utils
C:\Tools\utils\GameOrganizer\gameorganizer.bat
```

or copy the folder anywhere, as long as `gameorganizer.ps1` and
`gameorganizer.bat` stay together.

### Running

Double-click `GameOrganizer/gameorganizer.bat`. The launcher is:

```
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0gameorganizer.ps1"
```

`-ExecutionPolicy Bypass` is scoped to that one invocation and does not
change machine policy.

### Configuration

On first run a setup wizard asks for your library, archive zone, and
scratch paths. Steam, GOG, and defrag are optional; uncheck what you
don't use. The archive is expressed as **volume + zone** (drive
dropdown + subpath) so you can move it across drives without retyping
the whole path. Paths are stored in `%APPDATA%\GameOrganizer\config.json`
and can be edited from the config bar at the top of the main window.
**Clear Saved Settings** wipes that folder and reopens the wizard.

Keep the archive zone on the same volume as your libraries, as toggle
is a rename within a volume and a full copy across volumes.

<a id="about-the-compression"></a>

### About the compression

The Compress action runs `compact.exe /exe:<algo>`, which applies
Windows 10/11's per-file transparent compression (originally built for
"Compact OS" / WimBoot). Each file is compressed once at rest and
decompressed on read by the NTFS driver. This is a different code path
from classic "NTFS compression" (the Advanced Attributes checkbox):

- It works on any NTFS volume regardless of whether the volume has
  traditional compression enabled; the file-level attribute is what
  triggers decompression, not a folder flag.
- It avoids the pathological read-amplification and memory-mapped-file
  problems that make classic NTFS compression a poor fit for game data.
- Trade-off: reads are fine but any rewrite drops the file back to
  uncompressed, so this is suited to mostly-static install data, not
  save files or shader caches.

Algorithms in rough order of ratio vs. cost:

- **XPRESS4K / 8K / 16K**: fast, modest ratio. 16K is the usual pick.
- **LZX**: best ratio, considerably slower to compress. Runtime reads
  are still cheap.

**Shadow copies (VSS):** compressing a game rewrites every file, so if
VSS is enabled on the volume its shadow storage will grow to keep the
prior (uncompressed) content. On large libraries this can fill the VSS
quota and silently discard older restore points. The Defrag action has
the same effect for the same reason. If you rely on VSS, run compress
or defrag outside your snapshot window, or increase `vssadmin resize
shadowstorage` first.
