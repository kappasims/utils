Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- CONFIGURATION ---
$steamPath      = "F:\SteamLibrary\steamapps\common"
$gogPath        = "F:\GOG Games"
$archivedPath   = "F:\Archived"
$defragTempPath = "E:\TempDefrag"
$markerFileName = ".zonemover_compression"   # per-game marker so we can show state fast
# ---------------------

# Ensure required directories exist
if (-not (Test-Path $archivedPath))   { New-Item -ItemType Directory -Path $archivedPath   | Out-Null }
if (-not (Test-Path $defragTempPath)) { New-Item -ItemType Directory -Path $defragTempPath | Out-Null }

# Warn if archive is on a different volume than sources (toggle would become a slow copy)
$srcRoot = (Get-Item $steamPath).PSDrive.Name
$arcRoot = (Get-Item $archivedPath).PSDrive.Name
if ($srcRoot -ne $arcRoot) {
    [System.Windows.Forms.MessageBox]::Show(
        "Warning: archive path ($archivedPath) is on a different volume than the library paths. Toggle operations will be slow (full copy instead of rename).",
        "Cross-volume config", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}

# ============================================================
# 1. MAIN FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Zone Mover: Steam & GOG Manager"
$form.Size = New-Object System.Drawing.Size(1050, 600)
$form.StartPosition = "CenterScreen"
$form.MinimumSize  = New-Object System.Drawing.Size(900, 500)

# ============================================================
# 2. DATAGRIDVIEW
# ============================================================
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10, 10)
$grid.Size     = New-Object System.Drawing.Size(1010, 430)
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

# Action button columns
$defragCol = New-Object System.Windows.Forms.DataGridViewButtonColumn
$defragCol.Name = "Defrag"
$defragCol.HeaderText = "Defrag"
$defragCol.Text = "Defrag"
$defragCol.UseColumnTextForButtonValue = $true
$defragCol.FillWeight = 60
$grid.Columns.Add($defragCol) | Out-Null

$compressCol = New-Object System.Windows.Forms.DataGridViewButtonColumn
$compressCol.Name = "Compress"
$compressCol.HeaderText = "Compress"
$compressCol.Text = "Apply"
$compressCol.UseColumnTextForButtonValue = $true
$compressCol.FillWeight = 60
$grid.Columns.Add($compressCol) | Out-Null

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
# 5. LOAD GAMES
# ============================================================
function Load-Games {
    Set-Status "Scanning libraries..." -busy $true
    $grid.Rows.Clear()

    $platforms = @(
        @{ Path = $steamPath; Name = "Steam" },
        @{ Path = $gogPath;   Name = "GOG"   }
    )

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

            $sizeMB = if ($cleanup -eq "Broken Link") { 0 } else { Get-FolderSizeMB $realPath }
            $compState = if ($cleanup -eq "Broken Link") { "-" } else { Get-CompressionState $realPath }

            $rowIdx = $grid.Rows.Add(
                $_.Name, $p.Name, $status, $sizeMB, $compState, $cleanup,
                "LZX", "Defrag", "Apply"
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
        }
    }
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

# ============================================================
# 10. CELLCLICK (Defrag + Compress, merged)
# ============================================================
$grid.Add_CellClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $row = $grid.Rows[$e.RowIndex]
    $data = $row.Tag
    if (-not $data) { return }

    $platform   = $data.Platform
    $gameName   = $data.GameName
    $isArchived = $data.IsArchived
    $realPath   = Resolve-RealPath $data.FullPath $platform $gameName $isArchived

    # ------------- DEFRAG -------------
    if ($e.ColumnIndex -eq $grid.Columns["Defrag"].Index) {
        $resp = [System.Windows.Forms.MessageBox]::Show(
            "Defrag '$gameName' via copy cycle to $defragTempPath? This will take a while and consumes SSD writes.",
            "Confirm Defrag",
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($resp -ne 'Yes') { return }

        $script:defragPhase = "starting"
        $script:defragFiles = 0

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
                Set-Status ("Defrag {0}: {1} ({2} files)" -f $gameName, $script:defragPhase, $script:defragFiles) -busy $true
            }
            elseif ($line -match '^\s+(\d+)\s+') {
                # robocopy per-file line starts with whitespace + size
                $script:defragFiles++
                if (($script:defragFiles % 20) -eq 0) {
                    Set-Status ("Defrag {0}: {1} ({2} files)" -f $gameName, $script:defragPhase, $script:defragFiles) -busy $true
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
                        "Defrag failed: $err`nCheck for '$gameName.old' folder - your data may still be there.",
                        "Error") | Out-Null
                } else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Defrag complete: $script:defragFiles files copied.",
                        "Done") | Out-Null
                }
                Load-Games
            }
        return
    }

    # ------------- COMPRESS -------------
    if ($e.ColumnIndex -eq $grid.Columns["Compress"].Index) {
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

        # Counters for progress — updated via OnProgress on UI thread
        $script:progFiles = 0
        $script:progDir   = ""

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
                Set-Status ("Compressing {0} ({1} files done) - {2}" -f $gameName, $script:progFiles, $script:progDir) -busy $true
            }
            elseif ($line -match '^\s*Uncompressing files in\s+(.+?)\s*$') {
                $script:progDir = $matches[1]
                Set-Status ("Decompressing {0} ({1} files done) - {2}" -f $gameName, $script:progFiles, $script:progDir) -busy $true
            }
            elseif ($line -match '\[OK\]\s*$') {
                # Per-file success line — compact.exe tags both compressed and decompressed files with "[OK]"
                $script:progFiles++
                if (($script:progFiles % 25) -eq 0) {
                    Set-Status ("Compressing {0} ({1} files) - {2}" -f $gameName, $script:progFiles, $script:progDir) -busy $true
                }
            }
            elseif ($line -match 'files? within') {
                # Summary line at the end: "123 files within 45 directories were compressed."
                Set-Status ("{0}: {1}" -f $gameName, $line.Trim()) -busy $true
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
                        "$gameName : $algo - $script:progFiles files processed.",
                        "Compress") | Out-Null
                }
                Load-Games
            }
        return
    }
})

# ============================================================
# 11. ADD CONTROLS & SHOW
# ============================================================
$form.Controls.Add($grid)
$form.Controls.Add($btnToggle)
$form.Controls.Add($btnFix)
$form.Controls.Add($btnRefresh)

Load-Games
$form.Add_FormClosing({
    if ($script:currentOp) {
        try { $script:currentOp.PS.Stop()    } catch {}
        try { $script:currentOp.PS.Dispose() } catch {}
        try { $script:currentOp.Runspace.Close() ; $script:currentOp.Runspace.Dispose() } catch {}
        $script:currentOp = $null
    }
    $jobTimer.Stop()
    $jobTimer.Dispose()
})
$form.ShowDialog() | Out-Null
