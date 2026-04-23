Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- DEFAULT CONFIGURATION (user overrides saved to %APPDATA%\GameOrganizer\config.json) ---
$script:steamPath      = "F:\SteamLibrary\steamapps\common"
$script:gogPath        = "F:\GOG Games"
$script:archivedPath   = "F:\Archived"
$script:defragTempPath = "E:\TempDefrag"
$markerFileName        = ".gameorganizer_compression"   # per-game marker so we can show state fast

$script:configDir  = Join-Path $env:APPDATA "GameOrganizer"
$script:configFile = Join-Path $script:configDir "config.json"

# If the user points at a Steam library root (contains steamapps\common) or at
# the steamapps folder itself, descend to the common folder. Otherwise return
# the input unchanged -- an invalid/nonexistent path is the user's problem.
function Resolve-SteamPath {
    param([string]$path)
    if (-not $path) { return $path }
    if (-not (Test-Path $path)) { return $path }

    # Already pointing at <lib>\steamapps\common
    $leaf   = Split-Path $path -Leaf
    $parent = Split-Path $path -Parent
    if ($leaf -ieq "common" -and $parent -and (Split-Path $parent -Leaf) -ieq "steamapps") {
        return $path
    }

    # <lib>\steamapps\common beneath the selection (library root)
    $c1 = Join-Path $path "steamapps\common"
    if (Test-Path $c1) { return $c1 }

    # <lib>\steamapps selected -- descend into its common folder
    $c2 = Join-Path $path "common"
    if ($leaf -ieq "steamapps" -and (Test-Path $c2)) { return $c2 }

    return $path
}

function Load-AppConfig {
    if (Test-Path $script:configFile) {
        try {
            $c = Get-Content $script:configFile -Raw -ErrorAction Stop | ConvertFrom-Json
            if ($c.SteamPath)      { $script:steamPath      = Resolve-SteamPath ([string]$c.SteamPath) }
            if ($c.GogPath)        { $script:gogPath        = [string]$c.GogPath }
            if ($c.ArchivedPath)   { $script:archivedPath   = [string]$c.ArchivedPath }
            if ($c.DefragTempPath) { $script:defragTempPath = [string]$c.DefragTempPath }
        } catch {}
    }
}

function Save-AppConfig {
    if (-not (Test-Path $script:configDir)) {
        New-Item -ItemType Directory -Path $script:configDir | Out-Null
    }
    $obj = [ordered]@{
        SteamPath      = $script:steamPath
        GogPath        = $script:gogPath
        ArchivedPath   = $script:archivedPath
        DefragTempPath = $script:defragTempPath
    }
    $obj | ConvertTo-Json | Set-Content -Path $script:configFile -Force
}

# ============================================================
# FIRST-RUN SETUP WIZARD
# Shown when %APPDATA%\GameOrganizer\config.json does not exist, and
# on demand when the user clicks "Clear Saved Settings". Each optional
# feature (Steam, GOG, Defrag) has a "I don't use this" checkbox.
# Returns $true if the user saved, $false on cancel.
# ============================================================
function Show-SetupWizard {
    param([hashtable]$preset)

    $defaults = @{
        SteamPath      = if ($preset -and $preset.SteamPath)      { $preset.SteamPath }      else { "C:\Program Files (x86)\Steam\steamapps\common" }
        GogPath        = if ($preset -and $preset.GogPath)        { $preset.GogPath }        else { "C:\GOG Games" }
        ArchivedPath   = if ($preset -and $preset.ArchivedPath)   { $preset.ArchivedPath }   else { "C:\GO_Archive" }
        DefragTempPath = if ($preset -and $preset.DefragTempPath) { $preset.DefragTempPath } else { "C:\GO_DefragTemp" }
    }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Game Organizer -- Setup"
    $f.Size = New-Object System.Drawing.Size(660, 620)
    $f.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false

    $padX    = 16
    $fieldW  = 540
    $btnW    = 32
    $rowBtnX = $padX + $fieldW + 6

    $intro = New-Object System.Windows.Forms.Label
    $intro.Text = "Configure where Game Organizer looks for games and where it parks archived data. Uncheck optional sections you don't need. Settings are stored in $($script:configDir) and can be changed later from the main window."
    $intro.Location = New-Object System.Drawing.Point($padX, 12)
    $intro.Size = New-Object System.Drawing.Size(($fieldW + $btnW + 6), 50)
    $intro.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $rows = @{}

    function Add-WizardSection {
        param(
            [ref]$y,
            [string]$key,
            [string]$title,
            [string]$help,
            [string]$initial,
            [bool]$optional,
            [string]$dialogDesc
        )
        $row = @{}
        $cy  = $y.Value

        if ($optional) {
            $cb = New-Object System.Windows.Forms.CheckBox
            $cb.Text = $title
            $cb.Checked = $true
            $cb.Location = New-Object System.Drawing.Point($padX, $cy)
            $cb.Size = New-Object System.Drawing.Size(($fieldW + $btnW), 22)
            $cb.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $row.CheckBox = $cb
            $cy += 24
        } else {
            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text = $title
            $lbl.Location = New-Object System.Drawing.Point($padX, $cy)
            $lbl.Size = New-Object System.Drawing.Size(($fieldW + $btnW), 20)
            $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $row.Label = $lbl
            $cy += 22
        }

        $hl = New-Object System.Windows.Forms.Label
        $hl.Text = $help
        $hl.Location = New-Object System.Drawing.Point($padX, $cy)
        $hl.Size = New-Object System.Drawing.Size(($fieldW + $btnW), 38)
        $hl.ForeColor = [System.Drawing.Color]::DimGray
        $row.Help = $hl
        $cy += 42

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Text = $initial
        $tb.Location = New-Object System.Drawing.Point($padX, $cy)
        $tb.Size = New-Object System.Drawing.Size($fieldW, 24)
        $row.TextBox = $tb

        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = "..."
        $btn.Size = New-Object System.Drawing.Size($btnW, 24)
        $btn.Location = New-Object System.Drawing.Point($rowBtnX, $cy)
        $btn.Add_Click({
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = $dialogDesc
            if ($tb.Text -and (Test-Path $tb.Text)) { $dlg.SelectedPath = $tb.Text }
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $tb.Text = $dlg.SelectedPath
            }
        }.GetNewClosure())
        $row.Button = $btn

        $cy += 32
        $y.Value = $cy
        $rows[$key] = $row
    }

    $y = 68

    Add-WizardSection ([ref]$y) -key "Steam" -optional $true `
        -title "I use Steam" `
        -help "Pick your Steam library root (the folder that contains steamapps\common) or the common folder itself -- we'll descend automatically." `
        -initial $defaults.SteamPath `
        -dialogDesc "Steam library folder"

    Add-WizardSection ([ref]$y) -key "Gog" -optional $true `
        -title "I use GOG" `
        -help "Pick the folder that holds your GOG game installs (often called 'GOG Games')." `
        -initial $defaults.GogPath `
        -dialogDesc "GOG Games folder"

    Add-WizardSection ([ref]$y) -key "Archive" -optional $false `
        -title "Archive folder" `
        -help "Where games move when archived; the original path becomes a junction. Keep this on the same volume as your libraries, otherwise toggling becomes a full copy instead of a rename." `
        -initial $defaults.ArchivedPath `
        -dialogDesc "Archive folder"

    Add-WizardSection ([ref]$y) -key "Defrag" -optional $true `
        -title "Enable defrag" `
        -help "Scratch folder for the copy-cycle defrag. Any drive with free space works. Uncheck if you don't plan to defrag." `
        -initial $defaults.DefragTempPath `
        -dialogDesc "Defrag scratch folder"

    # Enable/disable children when the section checkbox toggles
    $syncEnabled = {
        foreach ($k in @("Steam","Gog","Defrag")) {
            $r = $rows[$k]
            if ($r.CheckBox) {
                $on = $r.CheckBox.Checked
                $r.TextBox.Enabled = $on
                $r.Button.Enabled  = $on
                $r.Help.ForeColor  = if ($on) { [System.Drawing.Color]::DimGray } else { [System.Drawing.Color]::LightGray }
            }
        }
    }.GetNewClosure()
    foreach ($k in @("Steam","Gog","Defrag")) {
        $rows[$k].CheckBox.Add_CheckedChanged($syncEnabled)
    }

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancel.Location = New-Object System.Drawing.Point(($rowBtnX - 112), ($y + 10))
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Save && Continue"
    $btnOK.Size = New-Object System.Drawing.Size(130, 28)
    $btnOK.Location = New-Object System.Drawing.Point(($rowBtnX - 18), ($y + 10))
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $f.Controls.Add($intro)
    foreach ($k in @("Steam","Gog","Archive","Defrag")) {
        $r = $rows[$k]
        if ($r.CheckBox) { $f.Controls.Add($r.CheckBox) }
        if ($r.Label)    { $f.Controls.Add($r.Label) }
        $f.Controls.Add($r.Help)
        $f.Controls.Add($r.TextBox)
        $f.Controls.Add($r.Button)
    }
    $f.Controls.Add($btnCancel)
    $f.Controls.Add($btnOK)
    $f.AcceptButton = $btnOK
    $f.CancelButton = $btnCancel
    $f.ClientSize   = New-Object System.Drawing.Size(($fieldW + $btnW + 6 + ($padX * 2)), ($y + 50))

    if ($f.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $false }

    $script:steamPath      = if ($rows["Steam"].CheckBox.Checked)  { Resolve-SteamPath ($rows["Steam"].TextBox.Text.Trim()) } else { "" }
    $script:gogPath        = if ($rows["Gog"].CheckBox.Checked)    { $rows["Gog"].TextBox.Text.Trim() }                      else { "" }
    $script:archivedPath   = $rows["Archive"].TextBox.Text.Trim()
    $script:defragTempPath = if ($rows["Defrag"].CheckBox.Checked) { $rows["Defrag"].TextBox.Text.Trim() }                   else { "" }
    Save-AppConfig
    return $true
}

$configExisted = Test-Path $script:configFile
Load-AppConfig
if (-not $configExisted) {
    if (-not (Show-SetupWizard)) {
        # User cancelled first-run setup -- nothing to do
        return
    }
}

# Ensure archive/defrag dirs exist (only if configured)
if ($script:archivedPath   -and -not (Test-Path $script:archivedPath))   { New-Item -ItemType Directory -Path $script:archivedPath   -Force | Out-Null }
if ($script:defragTempPath -and -not (Test-Path $script:defragTempPath)) { New-Item -ItemType Directory -Path $script:defragTempPath -Force | Out-Null }

# Cross-volume warning only runs when both a library and the archive actually exist
function Test-CrossVolume {
    $libRoot = $null
    foreach ($p in @($script:steamPath, $script:gogPath)) {
        if ($p -and (Test-Path $p)) { $libRoot = (Get-Item $p).PSDrive.Name; break }
    }
    if (-not $libRoot) { return }
    if (-not ($script:archivedPath -and (Test-Path $script:archivedPath))) { return }
    $arcRoot = (Get-Item $script:archivedPath).PSDrive.Name
    if ($libRoot -ne $arcRoot) {
        [System.Windows.Forms.MessageBox]::Show(
            "Warning: archive path ($($script:archivedPath)) is on a different volume than the library paths. Toggle operations will be slow (full copy instead of rename).",
            "Cross-volume config", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
}
Test-CrossVolume

# ============================================================
# 1. MAIN FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Game Organizer: Steam & GOG Manager"
$form.Size = New-Object System.Drawing.Size(1050, 600)
$form.StartPosition = "CenterScreen"
$form.MinimumSize  = New-Object System.Drawing.Size(900, 500)

# ============================================================
# 2. DATAGRIDVIEW
# ============================================================
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10, 80)
$grid.Size     = New-Object System.Drawing.Size(1010, 360)
$grid.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor `
                 [System.Windows.Forms.AnchorStyles]::Left -bor `
                 [System.Windows.Forms.AnchorStyles]::Right -bor `
                 [System.Windows.Forms.AnchorStyles]::Bottom
$grid.AllowUserToAddRows    = $false
$grid.AllowUserToDeleteRows = $false
$grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$grid.AutoSizeColumnsMode   = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.MultiSelect           = $false
$grid.RowHeadersVisible     = $false
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 250)
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

function Add-TextColumn {
    param($name, $header, $fillWeight = 100, $readOnly = $true)
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $name
    $c.HeaderText = $header
    $c.ReadOnly = $readOnly
    $c.FillWeight = $fillWeight
    $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $grid.Columns.Add($c) | Out-Null
}

Add-TextColumn "GameName"    "Game"          220
Add-TextColumn "Platform"    "Platform"      60
Add-TextColumn "Status"      "Status"        110
Add-TextColumn "SizeMB"      "Size (MB)"     80
Add-TextColumn "Compression" "Compression"   90
Add-TextColumn "Cleanup"     "Cleanup"       110

# Compression algorithm dropdown (per-row selector)
$algoCol = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$algoCol.Name = "Algo"
$algoCol.HeaderText = "Algorithm"
$algoCol.FillWeight = 80
$algoCol.Items.AddRange(@("LZX", "XPRESS16K", "XPRESS8K", "XPRESS4K", "None (decompress)"))
$algoCol.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$grid.Columns.Add($algoCol) | Out-Null

# Single header-less actions column; we custom-paint two buttons inside each cell
# and dispatch via click X-position (see CellPainting / CellMouseClick below).
$actionCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$actionCol.Name = "Actions"
$actionCol.HeaderText = ""
$actionCol.ReadOnly = $true
$actionCol.FillWeight = 140
$actionCol.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$actionCol.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::White
$actionCol.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
$grid.Columns.Add($actionCol) | Out-Null

# ============================================================
# 3. STATUS STRIP (progress + message at bottom)
# ============================================================
$statusStrip    = New-Object System.Windows.Forms.StatusStrip
$statusLabel    = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusLabel.Spring = $true
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusProgress = New-Object System.Windows.Forms.ToolStripProgressBar
$statusProgress.Size = New-Object System.Drawing.Size(200, 16)
$statusProgress.Visible = $false
[void]$statusStrip.Items.Add($statusLabel)
[void]$statusStrip.Items.Add($statusProgress)
$form.Controls.Add($statusStrip)

function Set-Status {
    param([string]$text, [int]$pct = -1, [bool]$busy = $false)
    $statusLabel.Text = $text
    if ($busy) {
        $statusProgress.Visible = $true
        if ($pct -lt 0) {
            $statusProgress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        } else {
            $statusProgress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
            $statusProgress.Value = [Math]::Min(100, [Math]::Max(0, $pct))
        }
    } else {
        $statusProgress.Visible = $false
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# ============================================================
# 4. HELPERS
# ============================================================
function Get-MarkerPath { param($gameRoot) Join-Path $gameRoot $markerFileName }

function Get-CompressionState {
    param($gameRoot)
    $mp = Get-MarkerPath $gameRoot
    if (Test-Path $mp) {
        try {
            $algo = (Get-Content $mp -First 1 -ErrorAction Stop).Trim()
            if ($algo) { return $algo }
        } catch {}
    }
    return "(uncompressed)"
}

function Set-CompressionState {
    param($gameRoot, $algo)
    $mp = Get-MarkerPath $gameRoot
    if ($algo -eq "None (decompress)" -or [string]::IsNullOrWhiteSpace($algo)) {
        if (Test-Path $mp) { Remove-Item $mp -Force -ErrorAction SilentlyContinue }
    } else {
        if (Test-Path $mp) { Remove-Item $mp -Force -ErrorAction SilentlyContinue }
        Set-Content -Path $mp -Value $algo -Force
        $mi = Get-Item -Path $mp -Force -ErrorAction SilentlyContinue
        if ($mi) { $mi.Attributes = 'Hidden' }
    }
}

function Get-FolderSizeMB {
    param($path)
    if (-not (Test-Path $path)) { return 0 }
    try {
        $bytes = (Get-ChildItem -Path $path -Recurse -File -Force -ErrorAction SilentlyContinue |
                  Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if ($null -eq $bytes) { return 0 }
        return [Math]::Round($bytes / 1MB, 0)
    } catch { return 0 }
}

function Resolve-RealPath {
    # For archived items (junctions), return the real archive path. For active items, return the source path itself.
    param($sourcePath, $platform, $gameName, $isArchived)
    if ($isArchived) {
        return (Join-Path $archivedPath ("{0}_{1}" -f $platform, $gameName))
    }
    return $sourcePath
}

# ============================================================
# 5. BUSY OVERLAY (faded panel + rotating-dot spinner over the grid)
# ============================================================
$script:overlayText  = "Working..."
$script:spinnerAngle = 0

$overlayPanel = New-Object System.Windows.Forms.Panel
$overlayPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 245, 248)
$overlayPanel.Visible = $false
$overlayPanel.Anchor = $grid.Anchor
$overlayPanel.Location = $grid.Location
$overlayPanel.Size     = $grid.Size
$overlayPanel.Cursor   = [System.Windows.Forms.Cursors]::WaitCursor

$overlayPanel.Add_Paint({
    param($sender, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $cx = [int]($sender.Width / 2)
    $cy = [int]($sender.Height / 2)
    $r  = 32

    # 12-dot rotating ring; trailing dots fade out
    for ($i = 0; $i -lt 12; $i++) {
        $angle = ([double]($script:spinnerAngle + ($i * 30))) * [Math]::PI / 180.0
        $x = $cx + [Math]::Cos($angle) * $r
        $y = $cy + [Math]::Sin($angle) * $r
        $alpha = [int](40 + (215 * ($i / 11.0)))
        $color = [System.Drawing.Color]::FromArgb($alpha, 60, 90, 170)
        $brush = New-Object System.Drawing.SolidBrush($color)
        $g.FillEllipse($brush, [single]($x - 5), [single]($y - 5), 10, 10)
        $brush.Dispose()
    }

    $font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $size = $g.MeasureString($script:overlayText, $font)
    $g.DrawString(
        $script:overlayText, $font,
        [System.Drawing.Brushes]::DimGray,
        [single]($cx - $size.Width / 2),
        [single]($cy + $r + 14)
    )
    $font.Dispose()
})

# Swallow clicks so the grid underneath stays "locked"
$overlayPanel.Add_Click({ })

$spinnerTimer = New-Object System.Windows.Forms.Timer
$spinnerTimer.Interval = 60
$spinnerTimer.Add_Tick({
    $script:spinnerAngle = ($script:spinnerAngle + 20) % 360
    $overlayPanel.Invalidate()
})

function Show-Overlay {
    param([string]$text = "Working...")
    $script:overlayText = $text
    $overlayPanel.Location = $grid.Location
    $overlayPanel.Size     = $grid.Size
    $overlayPanel.Visible  = $true
    $overlayPanel.BringToFront()
    $spinnerTimer.Start()
    $overlayPanel.Invalidate()
    [System.Windows.Forms.Application]::DoEvents()
}

function Hide-Overlay {
    $spinnerTimer.Stop()
    $overlayPanel.Visible = $false
}

# ============================================================
# 5b. LOAD GAMES
# ============================================================
function Load-Games {
    Show-Overlay "Scanning libraries..."
    Set-Status "Scanning libraries..." -busy $true
    $grid.Rows.Clear()
    [System.Windows.Forms.Application]::DoEvents()

    $platforms = @(
        @{ Path = $script:steamPath; Name = "Steam" },
        @{ Path = $script:gogPath;   Name = "GOG"   }
    )

    $seen = 0
    foreach ($p in $platforms) {
        if (-not (Test-Path $p.Path)) { continue }
        Get-ChildItem -Path $p.Path -Directory -Force -ErrorAction SilentlyContinue | ForEach-Object {
            $isJunction    = $_.Attributes -match "ReparsePoint"
            $archiveTarget = Join-Path $archivedPath ("{0}_{1}" -f $p.Name, $_.Name)
            $realPath      = if ($isJunction) { $archiveTarget } else { $_.FullName }

            $status  = if ($isJunction) { "Archived" } else { "Active" }
            $cleanup = "No"
            $color   = [System.Drawing.Color]::Black

            if ($isJunction) {
                if (-not (Test-Path $archiveTarget)) {
                    $cleanup = "Broken Link"
                    $color   = [System.Drawing.Color]::Red
                } else {
                    $color = [System.Drawing.Color]::DimGray
                }
            } else {
                if (Test-Path $archiveTarget) {
                    $cleanup = "Orphaned Archive"
                    $color   = [System.Drawing.Color]::DarkOrange
                }
            }

            $script:overlayText = "Scanning $($p.Name): $($_.Name)"

            $sizeMB = if ($cleanup -eq "Broken Link") { 0 } else { Get-FolderSizeMB $realPath }
            $compState = if ($cleanup -eq "Broken Link") { "-" } else { Get-CompressionState $realPath }

            $rowIdx = $grid.Rows.Add(
                $_.Name, $p.Name, $status, $sizeMB, $compState, $cleanup,
                "LZX", ""
            )
            $row = $grid.Rows[$rowIdx]
            $row.Tag = @{
                FullPath   = $_.FullName
                RealPath   = $realPath
                Platform   = $p.Name
                GameName   = $_.Name
                IsArchived = $isJunction
            }
            for ($i = 0; $i -lt $row.Cells.Count; $i++) {
                $row.Cells[$i].Style.ForeColor = $color
            }

            $seen++
            # Pump the message loop so the spinner animates and the UI stays responsive
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    Hide-Overlay
    Set-Status "Ready. $($grid.Rows.Count) game(s) listed."
}

# ============================================================
# 6. BACKGROUND OPERATION RUNNER (runspace-based)
#
#   Runs a scriptblock in a separate runspace (same process, shared
#   state possible but we keep it isolated). Supports:
#     - OnComplete callback (gets result or error)
#     - OnProgress callback (gets streamed output lines from a
#       thread-safe queue, polled by the UI timer)
#
#   The scriptblock receives a ConcurrentQueue as its first arg
#   (named $OutQueue). It can push progress lines into it:
#       $OutQueue.Enqueue("some progress text")
#   The UI timer drains the queue each tick and calls OnProgress
#   on the UI thread.
# ============================================================
Add-Type -AssemblyName System.Collections

$script:currentOp = $null  # holds @{ PS = PowerShell instance; Handle = AsyncResult; Runspace = Runspace; Queue = ConcurrentQueue; OnComplete; OnProgress }

$jobTimer = New-Object System.Windows.Forms.Timer
$jobTimer.Interval = 200
$jobTimer.Add_Tick({
    if ($null -eq $script:currentOp) {
        $jobTimer.Stop()
        return
    }
    $op = $script:currentOp

    # Drain progress queue -> UI
    if ($op.Queue -and $op.OnProgress) {
        $line = $null
        while ($op.Queue.TryDequeue([ref]$line)) {
            try { & $op.OnProgress $line } catch {}
        }
    }

    if ($op.Handle.IsCompleted) {
        $jobTimer.Stop()
        $result = $null; $err = $null
        try {
            $result = $op.PS.EndInvoke($op.Handle)
            if ($op.PS.HadErrors) {
                $errs = $op.PS.Streams.Error | ForEach-Object { $_.ToString() }
                if ($errs) { $err = ($errs -join "`n") }
            }
        } catch {
            $err = $_.Exception.Message
        } finally {
            try { $op.PS.Dispose() } catch {}
            try { $op.Runspace.Close(); $op.Runspace.Dispose() } catch {}
            $script:currentOp = $null
            $btnToggle.Enabled  = $true
            $btnFix.Enabled     = $true
            $btnRefresh.Enabled = $true
            $grid.Enabled       = $true
            if ($op.OnComplete) { & $op.OnComplete $result $err }
        }
    }
})

function Start-BackgroundOperation {
    param(
        [ScriptBlock]$Script,
        [object[]]$ArgumentList = @(),
        [string]$StatusText,
        [ScriptBlock]$OnComplete,
        [ScriptBlock]$OnProgress   # optional: receives each streamed line
    )
    if ($script:currentOp) {
        [System.Windows.Forms.MessageBox]::Show("Another operation is already running.", "Busy") | Out-Null
        return
    }
    $btnToggle.Enabled  = $false
    $btnFix.Enabled     = $false
    $btnRefresh.Enabled = $false
    $grid.Enabled       = $false
    Set-Status $StatusText -busy $true

    # Thread-safe queue for streaming progress lines
    $queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'

    # Build a runspace; pre-populate with the queue so the script can see it as $OutQueue
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('OutQueue', $queue)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Script)
    foreach ($a in $ArgumentList) { [void]$ps.AddArgument($a) }

    $handle = $ps.BeginInvoke()

    $script:currentOp = @{
        PS         = $ps
        Handle     = $handle
        Runspace   = $rs
        Queue      = $queue
        OnComplete = $OnComplete
        OnProgress = $OnProgress
    }
    $jobTimer.Start()
}

# ============================================================
# 7. TOGGLE (active <-> archived)
# ============================================================
$btnToggle = New-Object System.Windows.Forms.Button
$btnToggle.Text = "Toggle Active/Archived"
$btnToggle.Location = New-Object System.Drawing.Point(10, 450)
$btnToggle.Size = New-Object System.Drawing.Size(180, 35)
$btnToggle.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnToggle.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $row  = $grid.SelectedRows[0]
    $data = $row.Tag
    if ($row.Cells["Cleanup"].Value -ne "No") {
        [System.Windows.Forms.MessageBox]::Show("Resolve the cleanup issue first.", "Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    $sourcePath    = $data.FullPath
    $gameName      = $data.GameName
    $platform      = $data.Platform
    $isArchived    = $data.IsArchived
    $archiveTarget = Join-Path $archivedPath ("{0}_{1}" -f $platform, $gameName)

    $script = {
        param($sourcePath, $archiveTarget, $isArchived)
        if (-not $isArchived) {
            Move-Item -Path $sourcePath -Destination $archiveTarget -Force
            New-Item -ItemType Junction -Path $sourcePath -Target $archiveTarget | Out-Null
        } else {
            cmd /c rd """$sourcePath""" | Out-Null
            Move-Item -Path $archiveTarget -Destination $sourcePath -Force
        }
        return "ok"
    }

    Start-BackgroundOperation -Script $script `
        -ArgumentList @($sourcePath, $archiveTarget, $isArchived) `
        -StatusText "Toggling state for $gameName..." `
        -OnComplete {
            param($result, $err)
            if ($err) {
                [System.Windows.Forms.MessageBox]::Show("Toggle failed: $err", "Error") | Out-Null
            }
            Load-Games
        }
})

# ============================================================
# 8. FIX CLEANUP
# ============================================================
$btnFix = New-Object System.Windows.Forms.Button
$btnFix.Text = "Fix Cleanup Issue"
$btnFix.Location = New-Object System.Drawing.Point(200, 450)
$btnFix.Size = New-Object System.Drawing.Size(160, 35)
$btnFix.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnFix.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $row  = $grid.SelectedRows[0]
    $data = $row.Tag
    $sourcePath    = $data.FullPath
    $gameName      = $data.GameName
    $platform      = $data.Platform
    $archiveTarget = Join-Path $archivedPath ("{0}_{1}" -f $platform, $gameName)
    $state         = $row.Cells["Cleanup"].Value

    if ($state -eq "Orphaned Archive") {
        $resp = [System.Windows.Forms.MessageBox]::Show(
            "Delete orphaned archive data for '$gameName'?", "Clean Up",
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($resp -eq 'Yes') {
            Remove-Item -Path $archiveTarget -Recurse -Force
            Load-Games
        }
    } elseif ($state -eq "Broken Link") {
        $resp = [System.Windows.Forms.MessageBox]::Show(
            "Delete broken junction for '$gameName'?", "Clean Up",
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($resp -eq 'Yes') {
            cmd /c rd """$sourcePath""" | Out-Null
            Load-Games
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No cleanup needed.", "Info") | Out-Null
    }
})

# ============================================================
# 9. REFRESH
# ============================================================
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh"
$btnRefresh.Location = New-Object System.Drawing.Point(370, 450)
$btnRefresh.Size = New-Object System.Drawing.Size(100, 35)
$btnRefresh.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnRefresh.Add_Click({ Load-Games })

# Clear Saved Settings -- wipes %APPDATA%\GameOrganizer and re-runs the wizard
$btnClearSettings = New-Object System.Windows.Forms.Button
$btnClearSettings.Text = "Clear Saved Settings"
$btnClearSettings.Size = New-Object System.Drawing.Size(160, 35)
$btnClearSettings.Location = New-Object System.Drawing.Point(860, 450)
$btnClearSettings.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$btnClearSettings.Add_Click({
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Delete all saved settings at $($script:configDir)?`n`nThis does not touch any game files, just Game Organizer's own config.",
        "Clear Saved Settings",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    if (Test-Path $script:configDir) {
        try { Remove-Item $script:configDir -Recurse -Force -ErrorAction Stop }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Could not remove $($script:configDir): $($_.Exception.Message)", "Error") | Out-Null
            return
        }
    }

    # Reset in-memory and re-run the wizard so the user can pick fresh values
    $script:steamPath = ""; $script:gogPath = ""; $script:archivedPath = ""; $script:defragTempPath = ""
    if (Show-SetupWizard) {
        if ($script:archivedPath   -and -not (Test-Path $script:archivedPath))   { New-Item -ItemType Directory -Path $script:archivedPath   -Force | Out-Null }
        if ($script:defragTempPath -and -not (Test-Path $script:defragTempPath)) { New-Item -ItemType Directory -Path $script:defragTempPath -Force | Out-Null }
        $script:suppressReload = $true
        if ($script:steamRow) { $script:steamRow.TextBox.Text = $script:steamPath }
        if ($script:gogRow)   { $script:gogRow.TextBox.Text   = $script:gogPath }
        $script:suppressReload = $false
        Load-Games
    } else {
        # User cancelled re-setup; nothing to reload from, so show empty grid
        $grid.Rows.Clear()
        Set-Status "No configuration. Use Clear Saved Settings to run setup again."
    }
})

# ============================================================
# 10. ACTION COLUMN: custom-paint two buttons per cell, dispatch on click X
# ============================================================
$btnFontCache = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# Split a cell rect into left (Defrag) and right (Compress) button rectangles.
function Get-ActionButtonRects {
    param([System.Drawing.Rectangle]$cellBounds)
    $pad   = 4
    $gap   = 6
    $halfW = [Math]::Max(40, [int](($cellBounds.Width - ($pad * 2) - $gap) / 2))
    $btnH  = [Math]::Max(18, $cellBounds.Height - ($pad * 2))
    $y     = $cellBounds.Y + [int](($cellBounds.Height - $btnH) / 2)
    $left  = New-Object System.Drawing.Rectangle ($cellBounds.X + $pad), $y, $halfW, $btnH
    $right = New-Object System.Drawing.Rectangle ($cellBounds.X + $pad + $halfW + $gap), $y, $halfW, $btnH
    return @($left, $right)
}

$grid.Add_CellPainting({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    if ($e.ColumnIndex -ne $grid.Columns["Actions"].Index) { return }

    # Let the grid paint background/selection/border first
    $parts = [System.Windows.Forms.DataGridViewPaintParts]::Background `
        -bor [System.Windows.Forms.DataGridViewPaintParts]::Border `
        -bor [System.Windows.Forms.DataGridViewPaintParts]::SelectionBackground
    $e.Paint($e.ClipBounds, $parts)

    $rects = Get-ActionButtonRects $e.CellBounds
    [System.Windows.Forms.ButtonRenderer]::DrawButton(
        $e.Graphics, $rects[0], "Defrag", $btnFontCache, $false,
        [System.Windows.Forms.VisualStyles.PushButtonState]::Normal)
    [System.Windows.Forms.ButtonRenderer]::DrawButton(
        $e.Graphics, $rects[1], "Compress", $btnFontCache, $false,
        [System.Windows.Forms.VisualStyles.PushButtonState]::Normal)

    $e.Handled = $true
})

$grid.Add_CellMouseClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    if ($e.ColumnIndex -ne $grid.Columns["Actions"].Index) { return }
    $row = $grid.Rows[$e.RowIndex]
    $data = $row.Tag
    if (-not $data) { return }

    # CellMouseEventArgs.X/Y are relative to the cell; compute which half was clicked
    $cellWidth = $grid.Columns[$e.ColumnIndex].Width
    $action = if ($e.X -lt ($cellWidth / 2)) { "Defrag" } else { "Compress" }

    $platform   = $data.Platform
    $gameName   = $data.GameName
    $isArchived = $data.IsArchived
    $realPath   = Resolve-RealPath $data.FullPath $platform $gameName $isArchived

    # ------------- DEFRAG -------------
    if ($action -eq "Defrag") {
        $resp = [System.Windows.Forms.MessageBox]::Show(
            "Defrag '$gameName' via copy cycle to $defragTempPath? This will take a while and consumes SSD writes.",
            "Confirm Defrag",
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($resp -ne 'Yes') { return }

        $script:defragPhase = "starting"
        $script:defragFiles = 0
        $script:gameName    = $gameName

        $script = {
            param($realPath, $tempRoot, $gameName)

            $tempTarget = Join-Path $tempRoot $gameName
            $oldName    = [System.IO.Path]::GetFileName($realPath) + ".old"
            $oldPath    = Join-Path (Split-Path $realPath) $oldName

            if (Test-Path $tempTarget) { Remove-Item $tempTarget -Recurse -Force }
            if (Test-Path $oldPath)    { Remove-Item $oldPath -Recurse -Force }

            function Invoke-Robocopy {
                param($src, $dst, $phaseLabel)
                $OutQueue.Enqueue("PHASE:$phaseLabel")
                # /E recurse incl empty; /COPY:DAT data+attrs+timestamps; /R:2 retry 2x; /W:2 wait 2s;
                # /NP no percent; /NDL no dir listing; /BYTES show in bytes;
                # /NJH /NJS no job header/summary (keeps output compact).
                # We do want per-file lines to count progress.
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName  = "robocopy.exe"
                $psi.Arguments = "`"$src`" `"$dst`" /E /COPY:DAT /R:2 /W:2 /NP /NDL /NJH /NJS"
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError  = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow  = $true
                $proc = [System.Diagnostics.Process]::Start($psi)
                while (-not $proc.StandardOutput.EndOfStream) {
                    $line = $proc.StandardOutput.ReadLine()
                    if ($line) { $OutQueue.Enqueue($line) }
                }
                [void]$proc.StandardError.ReadToEnd()
                $proc.WaitForExit()
                # robocopy exit codes 0-7 are success variants; >=8 is an actual failure
                if ($proc.ExitCode -ge 8) {
                    throw "robocopy failed (exit $($proc.ExitCode)) copying $src -> $dst"
                }
            }

            # Phase 1: realPath -> tempTarget
            Invoke-Robocopy $realPath $tempTarget "Copying to temp"
            # Phase 2: rename realPath -> realPath.old (atomic on same volume)
            $OutQueue.Enqueue("PHASE:Renaming original to .old")
            Rename-Item -Path $realPath -NewName $oldName -Force
            # Phase 3: tempTarget -> realPath
            Invoke-Robocopy $tempTarget $realPath "Copying back (defragmented)"
            # Phase 4: cleanup
            $OutQueue.Enqueue("PHASE:Cleaning up")
            Remove-Item -Path $oldPath -Recurse -Force
            Remove-Item -Path $tempTarget -Recurse -Force
            return "ok"
        }

        $progressHandler = {
            param($line)
            if ($line -match '^PHASE:(.+)$') {
                $script:defragPhase = $matches[1]
                Set-Status ("Defrag {0}: {1} ({2} files)" -f $script:gameName, $script:defragPhase, $script:defragFiles) -busy $true
            }
            elseif ($line -match '^\s+(\d+)\s+') {
                # robocopy per-file line starts with whitespace + size
                $script:defragFiles++
                if (($script:defragFiles % 20) -eq 0) {
                    Set-Status ("Defrag {0}: {1} ({2} files)" -f $script:gameName, $script:defragPhase, $script:defragFiles) -busy $true
                }
            }
        }

        Start-BackgroundOperation -Script $script `
            -ArgumentList @($realPath, $defragTempPath, $gameName) `
            -StatusText "Defragmenting $gameName..." `
            -OnProgress $progressHandler `
            -OnComplete {
                param($result, $err)
                if ($err) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Defrag failed: $err`nCheck for '$($script:gameName).old' folder - your data may still be there.",
                        "Error") | Out-Null
                } else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Defrag complete: $($script:defragFiles) files copied.",
                        "Done") | Out-Null
                }
                Load-Games
            }
        return
    }

    # ------------- COMPRESS -------------
    if ($action -eq "Compress") {
        $algo = $row.Cells["Algo"].Value
        if (-not $algo) { $algo = "LZX" }

        $confirmMsg = if ($algo -eq "None (decompress)") {
            "Decompress '$gameName'? This rewrites all files uncompressed."
        } else {
            "Compress '$gameName' with $algo? Slow algorithms (LZX) can take 30+ min on large games."
        }
        $resp = [System.Windows.Forms.MessageBox]::Show(
            $confirmMsg, "Confirm Compress",
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($resp -ne 'Yes') { return }

        # Counters for progress -- updated via OnProgress on UI thread
        $script:progFiles = 0
        $script:progDir   = ""
        $script:gameName  = $gameName
        $script:algo      = $algo

        $script = {
            param($realPath, $algo, $markerName)

            $markerPath = Join-Path $realPath $markerName

            # Build compact args
            if ($algo -eq "None (decompress)") {
                $cargs = @('/u', '/s', '/i', "$realPath\*")
            } else {
                $exeArg = switch ($algo) {
                    "LZX"       { "/exe:lzx" }
                    "XPRESS16K" { "/exe:xpress16k" }
                    "XPRESS8K"  { "/exe:xpress8k" }
                    "XPRESS4K"  { "/exe:xpress4k" }
                    default     { "/exe:lzx" }
                }
                $cargs = @('/c', '/s', '/i', $exeArg, "$realPath\*")
            }

            # Start compact.exe with redirected stdout so we can stream line by line
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName  = "compact.exe"
            $psi.Arguments = ($cargs | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' '
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow  = $true

            $proc = [System.Diagnostics.Process]::Start($psi)

            # Stream stdout line by line into the UI queue
            while (-not $proc.StandardOutput.EndOfStream) {
                $line = $proc.StandardOutput.ReadLine()
                if ($null -ne $line -and $line.Length -gt 0) {
                    $OutQueue.Enqueue($line)
                }
            }
            # Drain any stderr
            $stderr = $proc.StandardError.ReadToEnd()
            $proc.WaitForExit()

            if ($proc.ExitCode -ne 0 -and $stderr) {
                # Compact uses non-zero exit for some benign cases too; only throw if we got stderr content
                throw "compact exited with code $($proc.ExitCode): $stderr"
            }

            # Write/remove marker
            if ($algo -eq "None (decompress)") {
                if (Test-Path $markerPath) { Remove-Item $markerPath -Force }
            } else {
                # Delete any existing (possibly hidden) marker first so Set-Content
                # always writes a clean file, then re-apply the Hidden attribute.
                if (Test-Path $markerPath) { Remove-Item $markerPath -Force }
                Set-Content -Path $markerPath -Value $algo -Force
                $mItem = Get-Item -Path $markerPath -Force -ErrorAction SilentlyContinue
                if ($mItem) { $mItem.Attributes = 'Hidden' }
            }
            return "ok"
        }

        # OnProgress parses Compact's output lines and updates status.
        # Compact emits lines like:
        #   "Compressing files in F:\...\game\data\"
        #   "   foo.dat        1234567 :    234567 = 5.3 to 1"
        # and a summary at the end. We count files compressed and show current dir.
        $progressHandler = {
            param($line)
            if ($line -match '^\s*Compressing files in\s+(.+?)\s*$') {
                $script:progDir = $matches[1]
                Set-Status ("Compressing {0} ({1} files done) - {2}" -f $script:gameName, $script:progFiles, $script:progDir) -busy $true
            }
            elseif ($line -match '^\s*Uncompressing files in\s+(.+?)\s*$') {
                $script:progDir = $matches[1]
                Set-Status ("Decompressing {0} ({1} files done) - {2}" -f $script:gameName, $script:progFiles, $script:progDir) -busy $true
            }
            elseif ($line -match '\[OK\]\s*$') {
                # Per-file success line -- compact.exe tags both compressed and decompressed files with "[OK]"
                $script:progFiles++
                if (($script:progFiles % 25) -eq 0) {
                    Set-Status ("Compressing {0} ({1} files) - {2}" -f $script:gameName, $script:progFiles, $script:progDir) -busy $true
                }
            }
            elseif ($line -match 'files? within') {
                # Summary line at the end: "123 files within 45 directories were compressed."
                Set-Status ("{0}: {1}" -f $script:gameName, $line.Trim()) -busy $true
            }
        }

        Start-BackgroundOperation -Script $script `
            -ArgumentList @($realPath, $algo, $markerFileName) `
            -StatusText "Compressing $gameName with $algo..." `
            -OnProgress $progressHandler `
            -OnComplete {
                param($result, $err)
                if ($err) {
                    [System.Windows.Forms.MessageBox]::Show("Compression failed: $err", "Error") | Out-Null
                } else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "$($script:gameName) : $($script:algo) - $($script:progFiles) files processed.",
                        "Compress") | Out-Null
                }
                Load-Games
            }
        return
    }
})

# ============================================================
# 11. CONFIG PANEL (library paths, editable from the GUI)
# ============================================================
function New-PathRow {
    param(
        [string]$label,
        [int]$y,
        [string]$initial,
        [scriptblock]$onCommit,
        [string]$dialogDesc
    )

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $label
    $lbl.Location = New-Object System.Drawing.Point(0, ($y + 4))
    $lbl.Size = New-Object System.Drawing.Size(52, 22)
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(55, $y)
    $tb.Size = New-Object System.Drawing.Size(915, 24)
    $tb.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                 [System.Windows.Forms.AnchorStyles]::Left -bor `
                 [System.Windows.Forms.AnchorStyles]::Right
    $tb.Text = $initial

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "..."
    $btn.Size = New-Object System.Drawing.Size(32, 24)
    $btn.Location = New-Object System.Drawing.Point(975, $y)
    $btn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

    # Commit on Leave (blur) or Enter key
    $commit = {
        $new = $tb.Text.Trim()
        & $onCommit $new
    }.GetNewClosure()

    $tb.Add_Leave($commit)
    $tb.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true
            $s.Parent.Focus() | Out-Null  # triggers Leave on $tb
        }
    })

    $btn.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = $dialogDesc
        if ($tb.Text -and (Test-Path $tb.Text)) { $dlg.SelectedPath = $tb.Text }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $tb.Text = $dlg.SelectedPath
            & $onCommit $dlg.SelectedPath
        }
    }.GetNewClosure())

    return @{ Label = $lbl; TextBox = $tb; Button = $btn }
}

$cfgPanel = New-Object System.Windows.Forms.Panel
$cfgPanel.Location = New-Object System.Drawing.Point(10, 8)
$cfgPanel.Size     = New-Object System.Drawing.Size(1010, 66)
$cfgPanel.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor `
                     [System.Windows.Forms.AnchorStyles]::Left -bor `
                     [System.Windows.Forms.AnchorStyles]::Right

# Reloads on commit only when the path actually changed; saves config either way
# so an explicitly-cleared path is remembered as "not configured."
$script:suppressReload = $false

$script:steamRow = New-PathRow -label "Steam:" -y 4 -initial $script:steamPath `
    -dialogDesc "Select Steam library root or steamapps\common folder" `
    -onCommit {
        param($newPath)
        $resolved = Resolve-SteamPath $newPath
        # If we descended, reflect the jump back into the textbox
        if ($script:steamRow -and $script:steamRow.TextBox.Text -ne $resolved) {
            $script:steamRow.TextBox.Text = $resolved
        }
        if ($resolved -ne $script:steamPath) {
            $script:steamPath = $resolved
            Save-AppConfig
            if (-not $script:suppressReload) { Load-Games }
        }
    }

$script:gogRow = New-PathRow -label "GOG:" -y 34 -initial $script:gogPath `
    -dialogDesc "Select GOG Games folder" `
    -onCommit {
        param($newPath)
        if ($newPath -ne $script:gogPath) {
            $script:gogPath = $newPath
            Save-AppConfig
            if (-not $script:suppressReload) { Load-Games }
        }
    }

$steamRow = $script:steamRow
$gogRow   = $script:gogRow

$cfgPanel.Controls.AddRange(@(
    $steamRow.Label, $steamRow.TextBox, $steamRow.Button,
    $gogRow.Label,   $gogRow.TextBox,   $gogRow.Button
))

# ============================================================
# 12. ADD CONTROLS & SHOW
# ============================================================
$form.Controls.Add($cfgPanel)
$form.Controls.Add($grid)
$form.Controls.Add($overlayPanel)
$form.Controls.Add($btnToggle)
$form.Controls.Add($btnFix)
$form.Controls.Add($btnRefresh)
$form.Controls.Add($btnClearSettings)
$overlayPanel.BringToFront()

# Fire the initial scan after the form becomes visible so the overlay/spinner renders
$form.Add_Shown({ Load-Games })

$form.Add_FormClosing({
    if ($script:currentOp) {
        try { $script:currentOp.PS.Stop()    } catch {}
        try { $script:currentOp.PS.Dispose() } catch {}
        try { $script:currentOp.Runspace.Close() ; $script:currentOp.Runspace.Dispose() } catch {}
        $script:currentOp = $null
    }
    $jobTimer.Stop()
    $jobTimer.Dispose()
    $spinnerTimer.Stop()
    $spinnerTimer.Dispose()
})
$form.ShowDialog() | Out-Null
