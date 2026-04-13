Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- CONFIGURATION ---
$steamPath = "F:\SteamLibrary\steamapps\common"
$gogPath = "F:\GOG Games"
$archivedPath = "F:\Archived"
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

# 2. ListView (Data Grid) Setup
$listView = New-Object System.Windows.Forms.ListView
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Location = New-Object System.Drawing.Point(10, 10)
$listView.Size = New-Object System.Drawing.Size(640, 340)
$listView.Columns.Add("Game Name", 240) | Out-Null
$listView.Columns.Add("Platform", 90) | Out-Null
$listView.Columns.Add("Status", 120) | Out-Null
$listView.Columns.Add("Needs Cleanup?", 160) | Out-Null

# 3. Core Logic: Scan Directories and Populate UI
function Load-Games {
    $listView.Items.Clear()
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $processFolders = {
        param($path, $platform)
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Directory | ForEach-Object {
                $isJunction = $_.Attributes -match "ReparsePoint"
                $archiveTarget = Join-Path $archivedPath "$platform`_$($_.Name)"
                
                $status = if ($isJunction) { "Archived (Zone 3)" } else { "Active (Zone 1)" }
                $cleanup = "No"
                
                $item = New-Object System.Windows.Forms.ListViewItem($_.Name)
                
                # Logic for the Cleanup column
                if ($isJunction) {
                    if (-not (Test-Path $archiveTarget)) {
                        $cleanup = "Yes (Broken Link)"
                        $item.ForeColor = [System.Drawing.Color]::Red
                    } else {
                        $item.ForeColor = [System.Drawing.Color]::Gray
                    }
                } else {
                    if (Test-Path $archiveTarget) {
                        $cleanup = "Yes (Orphaned Archive)"
                        $item.ForeColor = [System.Drawing.Color]::DarkOrange
                    } else {
                        $item.ForeColor = [System.Drawing.Color]::Black
                    }
                }

                $item.SubItems.Add($platform) | Out-Null
                $item.SubItems.Add($status) | Out-Null
                $item.SubItems.Add($cleanup) | Out-Null
                $item.Tag = $_.FullName # Store the full physical path hidden in the tag
                
                $listView.Items.Add($item) | Out-Null
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
    if ($listView.SelectedItems.Count -eq 0) { return }
    
    $selected = $listView.SelectedItems[0]
    $sourcePath = $selected.Tag
    $gameName = $selected.Text
    $platform = $selected.SubItems[1].Text
    $isArchived = $selected.SubItems[2].Text -match "Archived"
    $archiveTarget = Join-Path $archivedPath "$platform`_$gameName" 
    
    if ($selected.SubItems[3].Text -ne "No") {
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
            cmd /c rd "`"$sourcePath`"" # Delete junction silently
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
    if ($listView.SelectedItems.Count -eq 0) { return }
    
    $selected = $listView.SelectedItems[0]
    $sourcePath = $selected.Tag
    $gameName = $selected.Text
    $platform = $selected.SubItems[1].Text
    $cleanupState = $selected.SubItems[3].Text
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

# 6. Refresh Button
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh List"
$btnRefresh.Location = New-Object System.Drawing.Point(290, 360)
$btnRefresh.Size = New-Object System.Drawing.Size(130, 35)
$btnRefresh.Add_Click({ Load-Games })

# Combine and Execute
$form.Controls.Add($listView)
$form.Controls.Add($btnToggle)
$form.Controls.Add($btnFix)
$form.Controls.Add($btnRefresh)

Load-Games
$form.ShowDialog() | Out-Null