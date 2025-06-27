<#
.SYNOPSIS
    Windows Utility Toolkit - A GUI for common system administration and cleanup tasks.
    Combines tools for removing bloatware, resetting caches, and unlocking folders.

.WARNING
    This toolkit contains DESTRUCTIVE operations. Removing system components like Edge or OneDrive
    is IRREVERSIBLE and can lead to unexpected behavior. Use with extreme caution. The author
    is not responsible for any data loss or system damage.
#>

# --- Required Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Script Globals ---
$script:foundPids = @()
$script:handleExePath = $null
$script:scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# --- Core Functions ---
function Write-Log {
    param([string]$Message, [System.Drawing.Color]$Color = [System.Drawing.Color]::Black, [bool]$Bold = $false)
    $fontStyle = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $logBox.SelectionStart = $logBox.TextLength; $logBox.SelectionLength = 0
    $logBox.SelectionColor = $Color
    $logBox.SelectionFont = New-Object System.Drawing.Font($logBox.Font.FontFamily, $logBox.Font.Size, $fontStyle)
    $logBox.AppendText("$(Get-Date -Format 'HH:mm:ss') - $Message`n")
    $logBox.SelectionColor = $logBox.ForeColor; $logBox.SelectionFont = $logBox.Font
    $logBox.ScrollToCaret()
}

function Confirm-Action {
    param([string]$ActionTitle, [string]$ActionMessage)
    $confirmResult = [System.Windows.Forms.MessageBox]::Show($ActionMessage, $ActionTitle, 'YesNo', 'Warning', 'Button2')
    return $confirmResult -eq 'Yes'
}

function Invoke-FolderBrowser {
    param([string]$Title = "Select a folder")
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Title
    $dialog.ShowNewFolderButton = $true
    if ($dialog.ShowDialog() -eq 'OK') {
        return $dialog.SelectedPath
    }
    return $null
}

function Invoke-NukeEdge {
    if (-not (Confirm-Action "Confirm Edge Removal" "WARNING: This will PERMANENTLY remove Microsoft Edge and is irreversible.`n`nThis can affect applications that rely on the Edge WebView2 runtime.`n`nAre you absolutely sure you want to proceed?")) {
        Write-Log "Edge removal cancelled by user." "Blue"; return
    }
    Write-Log "--- Starting Microsoft Edge Annihilation ---" "DarkRed" -Bold $true
    $edgePath = "$env:ProgramFiles(x86)\Microsoft"
    Write-Log "Terminating all Edge-related processes..."; Stop-Process -Name "msedge", "msedgewebview2", "MicrosoftEdgeUpdate", "mms_utility" -Force -ErrorAction SilentlyContinue
    foreach ($folder in @("Edge", "EdgeCore", "EdgeUpdate")) {
        $fullPath = Join-Path $edgePath $folder
        if (Test-Path $fullPath) {
            Write-Log "Taking ownership and deleting '$folder' folder..."
            try { takeown /f $fullPath /r /d y > $null; icacls $fullPath /grant "Administrators:F" /t /c /q > $null; Remove-Item -Path $fullPath -Recurse -Force -ErrorAction Stop; Write-Log "Successfully deleted '$fullPath'." "Green" } 
            catch { Write-Log "ERROR deleting folder '$fullPath': $($_.Exception.Message)" "Red" }
        } else { Write-Log "Folder '$folder' not found, skipping." }
    }
    Write-Log "Cleaning up scheduled tasks and services..."; Unregister-ScheduledTask -TaskName "MicrosoftEdgeUpdateTaskMachineCore", "MicrosoftEdgeUpdateTaskMachineUA" -Confirm:$false -ErrorAction SilentlyContinue
    sc.exe delete edgeupdate > $null 2>&1; sc.exe delete edgeupdatem > $null 2>&1
    Write-Log "Applying registry block to prevent reinstallation..."
    $regPath = "HKLM:\SOFTWARE\Microsoft\EdgeUpdate"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    Set-ItemProperty -Path $regPath -Name "DoNotUpdateToEdgeWithChromium" -Value 1 -Type DWord -Force
    Write-Log "Deleting Start Menu & Desktop shortcuts..."
    Remove-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk", "$env:PUBLIC\Desktop\Microsoft Edge.lnk" -Force -ErrorAction SilentlyContinue
    Write-Log "--- Edge Annihilation Complete. A reboot is recommended. ---" "Green" -Bold $true
}

function Invoke-RemoveOneDrive {
    if (-not (Confirm-Action "Confirm OneDrive Removal" "WARNING: This will PERMANENTLY uninstall OneDrive and delete its local data.`n`nFiles already synced to the cloud will NOT be deleted from your Microsoft account.`n`nAre you absolutely sure you want to proceed?")) {
        Write-Log "OneDrive removal cancelled by user." "Blue"; return
    }
    Write-Log "--- Starting OneDrive Eradication ---" "DarkRed" -Bold $true
    Write-Log "Terminating OneDrive process..."; Stop-Process -Name "OneDrive" -Force -ErrorAction SilentlyContinue
    Write-Log "Running OneDrive uninstaller..."
    $setupPath64 = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"; $setupPath32 = "$env:SystemRoot\System32\OneDriveSetup.exe"
    if (Test-Path $setupPath64) { Start-Process -FilePath $setupPath64 -ArgumentList "/uninstall" -Wait } 
    elseif (Test-Path $setupPath32) { Start-Process -FilePath $setupPath32 -ArgumentList "/uninstall" -Wait } 
    else { Write-Log "OneDriveSetup.exe not found, it may already be uninstalled." }
    Write-Log "Deleting remaining OneDrive folders..."
    @("$env:UserProfile\OneDrive", "$env:LocalAppData\Microsoft\OneDrive", "$env:ProgramData\Microsoft OneDrive", "C:\OneDriveTemp") | ForEach-Object { if (Test-Path $_) { Remove-Item -Path $_ -Recurse -Force -ErrorAction SilentlyContinue; Write-Log "Removed folder: $_" } }
    Write-Log "Removing OneDrive from File Explorer sidebar..."; Remove-Item -Path "Registry::HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}", "Registry::HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Applying registry blocks to prevent reinstallation..."; $policyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive"; if (-not (Test-Path $policyPath)) { New-Item -Path $policyPath -Force | Out-Null }; Set-ItemProperty -Path $policyPath -Name "DisableFileSyncNGSC" -Value 1 -Type DWord -Force
    Write-Log "--- OneDrive Eradication Complete. ---" "Green" -Bold $true
}

function Invoke-ResetIconCache {
    if (-not (Confirm-Action "Confirm Icon Cache Reset" "This will briefly kill and restart Windows Explorer, causing your taskbar and desktop icons to disappear and reappear.`n`nSave any work in open File Explorer windows before proceeding.`n`nContinue?")) {
        Write-Log "Icon Cache reset cancelled by user." "Blue"; return
    }
    Write-Log "--- Resetting Icon Cache ---" "DarkBlue" -Bold $true
    try {
        Write-Log "Terminating Windows Explorer..."; Stop-Process -Name explorer -Force
        $iconCachePath = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
        Write-Log "Deleting iconcache files in '$iconCachePath'..."; Get-ChildItem -Path $iconCachePath -Filter "iconcache*.db" | Remove-Item -Force
        Write-Log "Cache files deleted." "Green"
    } catch { Write-Log "An error occurred during cleanup: $($_.Exception.Message)" "Red"
    } finally { Write-Log "Restarting Windows Explorer..."; Start-Process explorer.exe; Write-Log "--- Icon Cache Reset Complete! ---" "Green" -Bold $true }
}

$script:defenderRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet"; $script:defenderRegName = "SubmitSamplesConsent"
function Update-DefenderButton {
    $currentValue = Get-ItemProperty -Path $script:defenderRegPath -Name $script:defenderRegName -ErrorAction SilentlyContinue
    if ($currentValue -and $currentValue.($script:defenderRegName) -eq 2) { $defenderStatusLabel.Text = "Current Status: Fix Applied (Automatic sample submission is disabled)."; $defenderStatusLabel.ForeColor = "DarkGreen"; $toggleDefenderButton.Text = "Revert to Windows Default" } 
    else { $defenderStatusLabel.Text = "Current Status: Windows Default (May automatically submit samples)."; $defenderStatusLabel.ForeColor = "DarkRed"; $toggleDefenderButton.Text = "Apply Fix (Disable Automatic Submission)" }
}
function Invoke-ToggleDefenderFix {
    $currentValue = Get-ItemProperty -Path $script:defenderRegPath -Name $script:defenderRegName -ErrorAction SilentlyContinue
    if ($currentValue -and $currentValue.($script:defenderRegName) -eq 2) {
        Write-Log "Reverting Defender setting to default (Value: 0)..."; Set-ItemProperty -Path $script:defenderRegPath -Name $script:defenderRegName -Value 0 -Type DWord -Force; Write-Log "Defender setting reverted." "Green"
    } else {
        Write-Log "Applying Defender fix (Value: 2)..."; if (-not (Test-Path $script:defenderRegPath)) { New-Item -Path $script:defenderRegPath -Force | Out-Null }; Set-ItemProperty -Path $script:defenderRegPath -Name $script:defenderRegName -Value 2 -Type DWord -Force; Write-Log "Defender fix applied." "Green"
    }
    Update-DefenderButton
}

function Find-HandleExe {
    foreach($exeName in @("handle64.exe", "handle.exe")) {
        $localPath = Join-Path $script:scriptDir "Force_Close_Folder\$exeName"
        if(Test-Path $localPath) { $script:handleExePath = $localPath; Write-Log "Found '$exeName' in script sub-directory." "Green"; return }
        $pathResult = Get-Command $exeName -ErrorAction SilentlyContinue
        if ($pathResult) { $script:handleExePath = $pathResult.Source; Write-Log "Found '$exeName' in system PATH." "Green"; return }
    }
    $msg = "ERROR: 'handle64.exe' or 'handle.exe' not found in '.\Force_Close_Folder' or system PATH."; Write-Log $msg "Red"; [System.Windows.Forms.MessageBox]::Show($msg, "Prerequisite Missing", "OK", "Error"); $script:handleExePath = $null
}

# --- GUI Element Definitions ---
$form = New-Object System.Windows.Forms.Form; $form.Text = "Windows Utility Toolkit"; $form.Size = New-Object System.Drawing.Size(700, 650); $form.MinimumSize = $form.Size; $form.StartPosition = 'CenterScreen'
$logBox = New-Object System.Windows.Forms.RichTextBox; $logBox.Location = New-Object System.Drawing.Point(10, 480); $logBox.Size = New-Object System.Drawing.Size(665, 120); $logBox.Anchor = 'Bottom', 'Left', 'Right'; $logBox.Font = New-Object System.Drawing.Font("Consolas", 9); $logBox.ReadOnly = $true; $form.Controls.Add($logBox)
$tabControl = New-Object System.Windows.Forms.TabControl; $tabControl.Location = New-Object System.Drawing.Point(10, 10); $tabControl.Size = New-Object System.Drawing.Size(665, 460); $tabControl.Anchor = 'Top', 'Left', 'Right', 'Bottom'; $form.Controls.Add($tabControl)

# Tab 1: System Cleanup
$cleanupTab = New-Object System.Windows.Forms.TabPage; $cleanupTab.Text = "System Cleanup"; $tabControl.Controls.Add($cleanupTab)
$edgeGroup = New-Object System.Windows.Forms.GroupBox; $edgeGroup.Text = "Microsoft Edge"; $edgeGroup.Size = New-Object System.Drawing.Size(630, 180); $edgeGroup.Location = New-Object System.Drawing.Point(10, 10); $cleanupTab.Controls.Add($edgeGroup)
$edgeLabel = New-Object System.Windows.Forms.Label; $edgeLabel.Text = "This tool will forcefully and permanently remove Microsoft Edge, its updater, and related components. This action is IRREVERSIBLE and might affect apps that depend on the Edge WebView2 runtime. Proceed with extreme caution."; $edgeLabel.Location = New-Object System.Drawing.Point(15, 25); $edgeLabel.Size = New-Object System.Drawing.Size(600, 70); $edgeGroup.Controls.Add($edgeLabel)
$edgeButton = New-Object System.Windows.Forms.Button; $edgeButton.Text = "Permanently Remove Microsoft Edge"; $edgeButton.Location = New-Object System.Drawing.Point(15, 110); $edgeButton.Size = New-Object System.Drawing.Size(250, 40); $edgeButton.BackColor = "MistyRose"; $edgeButton.add_Click({ $edgeButton.Enabled = $false; try { Invoke-NukeEdge } finally { $edgeButton.Enabled = $true } }); $edgeGroup.Controls.Add($edgeButton)
$oneDriveGroup = New-Object System.Windows.Forms.GroupBox; $oneDriveGroup.Text = "Microsoft OneDrive"; $oneDriveGroup.Size = New-Object System.Drawing.Size(630, 180); $oneDriveGroup.Location = New-Object System.Drawing.Point(10, 200); $cleanupTab.Controls.Add($oneDriveGroup)
$oneDriveLabel = New-Object System.Windows.Forms.Label; $oneDriveLabel.Text = "This tool will uninstall OneDrive, remove it from the File Explorer sidebar, and attempt to block it from reinstalling. This does NOT delete files from your cloud storage, only from this PC. This action is IRREVERSIBLE."; $oneDriveLabel.Location = New-Object System.Drawing.Point(15, 25); $oneDriveLabel.Size = New-Object System.Drawing.Size(600, 70); $oneDriveGroup.Controls.Add($oneDriveLabel)
$oneDriveButton = New-Object System.Windows.Forms.Button; $oneDriveButton.Text = "Permanently Remove OneDrive"; $oneDriveButton.Location = New-Object System.Drawing.Point(15, 110); $oneDriveButton.Size = New-Object System.Drawing.Size(250, 40); $oneDriveButton.BackColor = "MistyRose"; $oneDriveButton.add_Click({ $oneDriveButton.Enabled = $false; try { Invoke-RemoveOneDrive } finally { $oneDriveButton.Enabled = $true } }); $oneDriveGroup.Controls.Add($oneDriveButton)

# Tab 2: System Tools
$toolsTab = New-Object System.Windows.Forms.TabPage; $toolsTab.Text = "System Tools"; $tabControl.Controls.Add($toolsTab); $toolsTab.Add_Enter({ Update-DefenderButton })
$iconCacheGroup = New-Object System.Windows.Forms.GroupBox; $iconCacheGroup.Text = "Icon Cache"; $iconCacheGroup.Size = New-Object System.Drawing.Size(630, 150); $iconCacheGroup.Location = New-Object System.Drawing.Point(10, 10); $toolsTab.Controls.Add($iconCacheGroup)
$iconCacheLabel = New-Object System.Windows.Forms.Label; $iconCacheLabel.Text = "If your desktop or folder icons are corrupted or displaying incorrectly, resetting the icon cache can often fix the issue. This will temporarily restart Windows Explorer."; $iconCacheLabel.Location = New-Object System.Drawing.Point(15, 25); $iconCacheLabel.Size = New-Object System.Drawing.Size(600, 45); $iconCacheGroup.Controls.Add($iconCacheLabel)
$iconCacheButton = New-Object System.Windows.Forms.Button; $iconCacheButton.Text = "Reset Icon Cache"; $iconCacheButton.Location = New-Object System.Drawing.Point(15, 80); $iconCacheButton.Size = New-Object System.Drawing.Size(200, 40); $iconCacheButton.add_Click({ $iconCacheButton.Enabled = $false; try { Invoke-ResetIconCache } finally { $iconCacheButton.Enabled = $true } }); $iconCacheGroup.Controls.Add($iconCacheButton)
$defenderGroup = New-Object System.Windows.Forms.GroupBox; $defenderGroup.Text = "Windows Defender Tweak"; $defenderGroup.Size = New-Object System.Drawing.Size(630, 180); $defenderGroup.Location = New-Object System.Drawing.Point(10, 170); $toolsTab.Controls.Add($defenderGroup)
$defenderLabel = New-Object System.Windows.Forms.Label; $defenderLabel.Text = "This setting controls 'Automatic sample submission' in Windows Defender. The fix disables it and gets rid of the annoying notifications, which can enhance peace and privacy. You can apply the fix or revert it to the Windows default."; $defenderLabel.Location = New-Object System.Drawing.Point(15, 25); $defenderLabel.Size = New-Object System.Drawing.Size(600, 45); $defenderGroup.Controls.Add($defenderLabel)
$defenderStatusLabel = New-Object System.Windows.Forms.Label; $defenderStatusLabel.Text = "Current Status: Checking..."; $defenderStatusLabel.Location = New-Object System.Drawing.Point(15, 80); $defenderStatusLabel.AutoSize = $true; $defenderStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $defenderGroup.Controls.Add($defenderStatusLabel)
$toggleDefenderButton = New-Object System.Windows.Forms.Button; $toggleDefenderButton.Text = "Apply/Revert"; $toggleDefenderButton.Location = New-Object System.Drawing.Point(15, 110); $toggleDefenderButton.Size = New-Object System.Drawing.Size(280, 40); $toggleDefenderButton.add_Click({ $toggleDefenderButton.Enabled = $false; try { Invoke-ToggleDefenderFix } finally { $toggleDefenderButton.Enabled = $true } }); $defenderGroup.Controls.Add($toggleDefenderButton)

# Tab 3: Folder Unlocker
$unlockerTab = New-Object System.Windows.Forms.TabPage; $unlockerTab.Text = "Folder Unlocker"; $tabControl.Controls.Add($unlockerTab); $unlockerTab.Add_Enter({ if(-not $script:handleExePath) { Find-HandleExe } })
$folderLabel = New-Object System.Windows.Forms.Label; $folderLabel.Text = "Target Folder:"; $folderLabel.Location = New-Object System.Drawing.Point(15, 20); $folderLabel.AutoSize = $true; $unlockerTab.Controls.Add($folderLabel)
$selectedFolderPathLabel = New-Object System.Windows.Forms.Label; $selectedFolderPathLabel.Text = "No folder selected."; $selectedFolderPathLabel.Location = New-Object System.Drawing.Point(110, 20); $selectedFolderPathLabel.Size = New-Object System.Drawing.Size(510, 23); $selectedFolderPathLabel.BorderStyle = 'FixedSingle'; $selectedFolderPathLabel.AutoEllipsis = $true; $unlockerTab.Controls.Add($selectedFolderPathLabel)
$browseButton = New-Object System.Windows.Forms.Button; $browseButton.Text = "Browse Folder..."; $browseButton.Location = New-Object System.Drawing.Point(15, 55); $browseButton.Size = New-Object System.Drawing.Size(130, 30); $unlockerTab.Controls.Add($browseButton)
$findButton = New-Object System.Windows.Forms.Button; $findButton.Text = "Find Locking Processes"; $findButton.Location = New-Object System.Drawing.Point(155, 55); $findButton.Size = New-Object System.Drawing.Size(160, 30); $findButton.Enabled = $false; $unlockerTab.Controls.Add($findButton)
$processListBox = New-Object System.Windows.Forms.ListBox; $processListBox.Location = New-Object System.Drawing.Point(15, 95); $processListBox.Size = New-Object System.Drawing.Size(605, 230); $processListBox.Font = New-Object System.Drawing.Font("Consolas", 9); $processListBox.ScrollAlwaysVisible = $true; $processListBox.HorizontalScrollbar = $true; $processListBox.SelectionMode = 'None'; $unlockerTab.Controls.Add($processListBox)
$killButton = New-Object System.Windows.Forms.Button; $killButton.Text = "Kill All Found Processes"; $killButton.Location = New-Object System.Drawing.Point(15, 335); $killButton.Size = New-Object System.Drawing.Size(605, 40); $killButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $killButton.BackColor = 'LightCoral'; $killButton.Enabled = $false; $unlockerTab.Controls.Add($killButton)

# --- Event Handlers for Folder Unlocker Tab ---
$browseButton.Add_Click({
    $selectedPath = Invoke-FolderBrowser -Title "Select the folder to unlock"
    if ($selectedPath) {
        $selectedFolderPathLabel.Text = $selectedPath
        $processListBox.Items.Clear(); $killButton.Enabled = $false; $script:foundPids = @()
        Write-Log "Folder selected: $selectedPath"; $findButton.Enabled = $true
    }
})
$findButton.Add_Click({
    $targetFolder = $selectedFolderPathLabel.Text
    $processListBox.Items.Clear(); $killButton.Enabled = $false; $script:foundPids = @()
    if (-not $script:handleExePath) { Write-Log "Cannot search: handle.exe is missing." "Red"; Find-HandleExe; return }
    if (-not (Test-Path $targetFolder -PathType Container)) { Write-Log "Error: Please select a valid folder first." "Red"; return }
    Write-Log "Searching for locking processes in '$targetFolder'..." "Blue"
    $findButton.Enabled = $false; $browseButton.Enabled = $false; $form.Cursor = 'WaitCursor'
    try {
        $handleOutput = & "$($script:handleExePath)" -accepteula -nobanner "$targetFolder" 2>&1
        if ($handleOutput -match "No matching handles found") { Write-Log "Success: No locking processes found." "Green"; return }
        $regex = '^\S+\s+pid:\s+(\d+)\s+type:\s+(?:File|Dir)\s+.*'
        $pidsFound = $handleOutput | Select-String -Pattern $regex | ForEach-Object { $_.Matches[0].Groups[1].Value } | Sort-Object -Unique
        $script:foundPids = $pidsFound
        if ($script:foundPids.Count -eq 0) {
            Write-Log "Could not parse any locking processes from handle.exe output." "Orange"
            $processListBox.Items.Add("--- Raw Handle.exe Output ---"); $handleOutput | ForEach-Object { $processListBox.Items.Add($_) }
        } else {
            Write-Log "Found $($script:foundPids.Count) process(es) locking files." "DarkOrange"
            $processListBox.Items.Add("PID `t Process Name"); $processListBox.Items.Add("--- `t -------------")
            foreach ($processId in $script:foundPids) {
                $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
                $processName = if ($process) { $process.ProcessName } else { "<Unknown/Access Denied>" }
                $processListBox.Items.Add("$processId `t $processName")
            }
            $killButton.Enabled = $true
        }
    } catch { Write-Log "An error occurred during find operation: $($_.Exception.Message)" "Red"
    } finally { if($selectedFolderPathLabel.Text -ne "No folder selected.") { $findButton.Enabled = $true }; $browseButton.Enabled = $true; $form.Cursor = 'Default' }
})
$killButton.Add_Click({
    if ($script:foundPids.Count -eq 0) { Write-Log "No processes selected to kill." "Orange"; return }
    if (-not (Confirm-Action "Confirm Force Kill" "DANGER! You are about to FORCE-CLOSE $($script:foundPids.Count) process(es).`n`nThis can cause data loss and system instability.`n`nARE YOU ABSOLUTELY SURE?")) {
        Write-Log "Kill operation cancelled by user." "Blue"; return
    }
    Write-Log "--- Attempting to kill $($script:foundPids.Count) processes... ---" "DarkRed"
    $killButton.Enabled = $false; $browseButton.Enabled = $false; $findButton.Enabled = $false; $form.Cursor = 'WaitCursor'
    $killedCount = 0; $failedCount = 0
    foreach ($processId in $script:foundPids) {
        try {
            $proc = Get-Process -Id $processId -ErrorAction SilentlyContinue
            Stop-Process -Id $processId -Force -ErrorAction Stop
            Write-Log "Successfully stopped $($proc.Name) (PID: $processId)" "Green"; $killedCount++
        } catch { Write-Log "FAILED to stop PID $processId. Error: $($_.Exception.Message -replace '[\r\n]',' ')" "Red"; $failedCount++ }
    }
    Write-Log "Kill operation finished. Killed: $killedCount, Failed: $failedCount." "Blue"
    $script:foundPids = @(); $form.Cursor = 'Default'; $browseButton.Enabled = $true
    if($selectedFolderPathLabel.Text -ne "No folder selected.") { $findButton.Enabled = $true }
    $processListBox.TopIndex = $processListBox.Items.Count - 1
    [System.Windows.Forms.MessageBox]::Show("Operation complete.`nKilled: $killedCount`nFailed: $failedCount`n`nYou can now try accessing the folder again.", "Operation Complete", "OK", "Information")
})

# Tab 4: About
$aboutTab = New-Object System.Windows.Forms.TabPage; $aboutTab.Text = "About"; $tabControl.Controls.Add($aboutTab)
$aboutLabel = New-Object System.Windows.Forms.Label; $aboutLabel.Text = "Windows Utility Toolkit`n`n`n`nWARNING:`nThese tools perform powerful and often irreversible actions. Use them at your own risk. The developer assumes no liability for data loss or system damage resulting from the use of this software."; $aboutLabel.Location = New-Object System.Drawing.Point(15, 20); $aboutLabel.Size = New-Object System.Drawing.Size(620, 300); $aboutTab.Controls.Add($aboutLabel)

# --- Show Form ---
Write-Log "GUI Loaded. Running as Administrator. Ready for commands." "Green" -Bold $true
$form.ShowDialog() | Out-Null