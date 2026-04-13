Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- CONFIGURATION ---
$steamPath = "F:\SteamLibrary\steamapps\common"
$gogPath = "F:\GOG Games"
$archivedPath = "F:\Archived"
$defragTempPath = "E:\TempDefrag"
# ---------------------

# Ensure archive directory exists
if (-not (Test-Path $archivedPath)) {
    New-Item -ItemType Directory -Path $archivedPath | Out-Null
}

# 1. Main Form Setup
$form = New-Object System.Windows.Forms.Form
$form.Text = "Zone Mover: Steam & GOG Manager"
$form.Size = New-Object System.Drawing.Size(680, 450)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# 2. DataGridView Setup
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10, 10)
$grid.Size = New-Object System.Drawing.Size(640, 340)
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.ReadOnly = $true
$grid.MultiSelect = $false

# Add columns
$col1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$col1.Name = "GameName"
$col1.HeaderText = "Game Name"
$col1.ReadOnly = $true
$grid.Columns.Add($col1) | Out-Null

$col2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$col2.Name = "Platform"
$col2.HeaderText = "Platform"
$col2.ReadOnly = $true
$grid.Columns.Add($col2) | Out-Null

$col3 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$col3.Name = "Status"
$col3.HeaderText = "Status"
$col3.ReadOnly = $true
$grid.Columns.Add($col3) | Out-Null

$col4 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$col4.Name = "Cleanup"
$col4.HeaderText = "Needs Cleanup?"
$col4.ReadOnly = $true
$grid.Columns.Add($col4) | Out-Null

$btnCol = New-Object System.Windows.Forms.DataGridViewButtonColumn
$btnCol.Name = "Defrag"
$btnCol.HeaderText = "Defrag"
$btnCol.Text = "Defrag"
$btnCol.UseColumnTextForButtonValue = $true
$grid.Columns.Add($btnCol) | Out-Null

# 3. Core Logic: Scan Directories and Populate UI
function Load-Games {
    $grid.Rows.Clear()
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $processFolders = {
        param($path, $platform)
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Directory | ForEach-Object {
                $isJunction = $_.Attributes -match "ReparsePoint"
                $archiveTarget = Join-Path $archivedPath "$platform`_$($_.Name)"

                $status = if ($isJunction) { "Archived (Zone 3)" } else { "Active (Zone 1)" }
                $cleanup = "No"
                $color = [System.Drawing.Color]::Black

                # Logic for the Cleanup column
                if ($isJunction) {
                    if (-not (Test-Path $archiveTarget)) {
                        $cleanup = "Yes (Broken Link)"
                        $color = [System.Drawing.Color]::Red
                    } else {
                        $color = [System.Drawing.Color]::Gray
                    }
                } else {
                    if (Test-Path $archiveTarget) {
                        $cleanup = "Yes (Orphaned Archive)"
                        $color = [System.Drawing.Color]::DarkOrange
                    }
                }

                $row = New-Object System.Windows.Forms.DataGridViewRow
                $row.CreateCells($grid, $_.Name, $platform, $status, $cleanup, "Defrag")
                $row.Tag = @{
                    FullPath = $_.FullName
                    Platform = $platform
                    GameName = $_.Name
                    IsArchived = $isJunction
                    IsActive = -not $isJunction
                }

                # Color the entire row
                for ($i = 0; $i -lt $row.Cells.Count; $i++) {
                    $row.Cells[$i].Style.ForeColor = $color
                }

                $grid.Rows.Add($row) | Out-Null
            }
        }
    }

    &$processFolders $steamPath "Steam"
    &$processFolders $gogPath "GOG"

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

# 4. Toggle Button Logic
$btnToggle = New-Object System.Windows.Forms.Button
$btnToggle.Text = "Toggle State"
$btnToggle.Location = New-Object System.Drawing.Point(10, 360)
$btnToggle.Size = New-Object System.Drawing.Size(130, 35)
$btnToggle.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }

    $row = $grid.SelectedRows[0]
    $data = $row.Tag
    $sourcePath = $data.FullPath
    $gameName = $data.GameName
    $platform = $data.Platform
    $isArchived = $data.IsArchived
    $archiveTarget = Join-Path $archivedPath "$platform`_$gameName"

    if ($grid.Rows[$row.Index].Cells[3].Value -ne "No") {
        [System.Windows.Forms.MessageBox]::Show("Please resolve the cleanup issue for this game before toggling its state.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $btnToggle.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        if (-not $isArchived) {
            Move-Item -Path $sourcePath -Destination $archiveTarget -Force
            New-Item -ItemType Junction -Path $sourcePath -Target $archiveTarget | Out-Null
        } else {
            cmd /c rd "`"$sourcePath`""
            Move-Item -Path $archiveTarget -Destination $sourcePath -Force
        }
        Load-Games
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error moving files: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $btnToggle.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# 5. Fix Cleanup Issue Button
$btnFix = New-Object System.Windows.Forms.Button
$btnFix.Text = "Fix Cleanup Issue"
$btnFix.Location = New-Object System.Drawing.Point(150, 360)
$btnFix.Size = New-Object System.Drawing.Size(130, 35)
$btnFix.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }

    $row = $grid.SelectedRows[0]
    $data = $row.Tag
    $sourcePath = $data.FullPath
    $gameName = $data.GameName
    $platform = $data.Platform
    $cleanupState = $grid.Rows[$row.Index].Cells[3].Value
    $archiveTarget = Join-Path $archivedPath "$platform`_$gameName"

    if ($cleanupState -match "Orphaned Archive") {
        $msg = [System.Windows.Forms.MessageBox]::Show("Delete the orphaned archive data for '$gameName' to free up space?", "Clean Up", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($msg -eq 'Yes') {
            Remove-Item -Path $archiveTarget -Recurse -Force
            Load-Games
        }
    } elseif ($cleanupState -match "Broken Link") {
        $msg = [System.Windows.Forms.MessageBox]::Show("The junction point for '$gameName' is broken (target missing). Delete the dead link?", "Clean Up", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($msg -eq 'Yes') {
            cmd /c rd "`"$sourcePath`""
            Load-Games
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No cleanup needed for this game.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# 6. Defrag Button (Per-Row via DataGridView Button Column)
$grid.Add_CellClick({
    if ($_.ColumnIndex -eq $grid.Columns["Defrag"].Index) {
        $row = $grid.Rows[$_.RowIndex]
        $data = $row.Tag
        $sourcePath = $data.FullPath
        $gameName = $data.GameName
        $platform = $data.Platform
        $isArchived = $data.IsArchived

        $msg = [System.Windows.Forms.MessageBox]::Show("Copy '$gameName' to E:\ for defragmentation? This may take a while.", "Confirm Defrag", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($msg -ne 'Yes') { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $grid.Enabled = $false

        try {
            if (-not (Test-Path $defragTempPath)) {
                New-Item -ItemType Directory -Path $defragTempPath | Out-Null
            }

            $tempTarget = Join-Path $defragTempPath $gameName

            # If archived, defrag the archive location, not the junction
            if ($isArchived) {
                $defragSource = Join-Path $archivedPath "$platform`_$gameName"
            } else {
                $defragSource = $sourcePath
            }

            Copy-Item -Path $defragSource -Destination $tempTarget -Recurse -Force
            Remove-Item -Path $defragSource -Recurse -Force
            Copy-Item -Path $tempTarget -Destination $defragSource -Recurse -Force
            Remove-Item -Path $tempTarget -Recurse -Force

            [System.Windows.Forms.MessageBox]::Show("Defragmentation complete for '$gameName'.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Load-Games
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error during defragmentation: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $grid.Enabled = $true
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
})

# 7. Refresh Button
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh List"
$btnRefresh.Location = New-Object System.Drawing.Point(290, 360)
$btnRefresh.Size = New-Object System.Drawing.Size(130, 35)
$btnRefresh.Add_Click({ Load-Games })

# Combine and Execute
$form.Controls.Add($grid)
$form.Controls.Add($btnToggle)
$form.Controls.Add($btnFix)
$form.Controls.Add($btnRefresh)

Load-Games
$form.ShowDialog() | Out-Null