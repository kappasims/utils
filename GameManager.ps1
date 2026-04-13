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
$form.Size = New-Object System.Drawing.Size(550, 450)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# 2. ListView (Data Grid) Setup
$listView = New-Object System.Windows.Forms.ListView
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Location = New-Object System.Drawing.Point(10, 10)
$listView.Size = New-Object System.Drawing.Size(510, 340)
$listView.Columns.Add("Game Name", 260) | Out-Null
$listView.Columns.Add("Platform", 100) | Out-Null
$listView.Columns.Add("Status", 120) | Out-Null

# 3. Core Logic: Scan Directories and Populate UI
function Load-Games {
    $listView.Items.Clear()
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $processFolders = {
        param($path, $platform)
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Directory | ForEach-Object {
                $isJunction = $_.Attributes -match "ReparsePoint"
                $status = if ($isJunction) { "Archived (Zone 3)" } else { "Active (Zone 1)" }
                
                $item = New-Object System.Windows.Forms.ListViewItem($_.Name)
                $item.SubItems.Add($platform) | Out-Null
                $item.SubItems.Add($status) | Out-Null
                $item.Tag = $_.FullName # Store the full physical path hidden in the tag
                
                # Visual distinction
                if ($isJunction) { $item.ForeColor = [System.Drawing.Color]::Gray }
                else { $item.ForeColor = [System.Drawing.Color]::Black }
                
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
$btnToggle.Size = New-Object System.Drawing.Size(150, 35)
$btnToggle.Add_Click({
    if ($listView.SelectedItems.Count -eq 0) { return }
    
    $selected = $listView.SelectedItems[0]
    $sourcePath = $selected.Tag
    $gameName = $selected.Text
    $platform = $selected.SubItems[1].Text
    $isArchived = $selected.SubItems[2].Text -match "Archived"
    
    # Prefix the platform to the archive folder to prevent naming collisions (e.g. if you own the same game on both)
    $archiveTarget = Join-Path $archivedPath "$platform`_$gameName" 
    
    $btnToggle.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        if (-not $isArchived) {
            # ACTIVE -> ARCHIVED
            Move-Item -Path $sourcePath -Destination $archiveTarget -Force
            New-Item -ItemType Junction -Path $sourcePath -Target $archiveTarget | Out-Null
        } else {
            # ARCHIVED -> ACTIVE
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

# 5. Refresh Button
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh List"
$btnRefresh.Location = New-Object System.Drawing.Point(170, 360)
$btnRefresh.Size = New-Object System.Drawing.Size(150, 35)
$btnRefresh.Add_Click({ Load-Games })

# Combine and Execute
$form.Controls.Add($listView)
$form.Controls.Add($btnToggle)
$form.Controls.Add($btnRefresh)

Load-Games
$form.ShowDialog() | Out-Null