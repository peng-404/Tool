# ==========================================
# 0. Elevate Privileges (Admin rights required)
# ==========================================
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Load .NET Forms Assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==========================================
# Phase 1: System Native Cleanup (Safest foundation)
# ==========================================
Clear-Host
Write-Host "--- Starting Phase 1: System Native Optimization ---" -ForegroundColor Cyan
Write-Host ">> Cleaning WinSxS components and update residues, please wait (may take several minutes)..." -ForegroundColor Gray
Start-Process "dism.exe" "/Online /Cleanup-Image /StartComponentCleanup" -Wait -NoNewWindow
Start-Process "cleanmgr.exe" "/d c /VERYLOWDISK" -Wait -NoNewWindow

# ==========================================
# Phase 2: Full-Scope Path Scanning
# ==========================================
Write-Host ">> Scanning third-party startups and deep hidden caches..." -ForegroundColor Gray
$startupList = @()
$cacheList = @()

# 1. Scan startup items (Capture: rogue software, popup viruses, bundled programs)
$startupItems = Get-CimInstance Win32_StartupCommand
$regex = '(?i)([a-z]:\\[^\*?"<>\|]+\.(?:exe|bat|cmd|vbs|dll|ps1))'
foreach ($item in $startupItems) {
    if ($item.Command -match $regex) {
        $path = $matches[1]
        # Exclude Microsoft native core paths
        if (($path -notmatch "(?i)^C:\\Windows\\") -and (Test-Path $path)) { $startupList += $path }
    }
}

# 2. Scan common and developer environment caches
$cachePaths = @(
    # --- Common Software Caches ---
    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\Cache_Data", 
    "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\Cache_Data", 
    "$env:USERPROFILE\Documents\WeChat Files\Applet",                      
    "$env:USERPROFILE\Documents\WeChat Files\FileStorage\Cache",          
    "$env:APPDATA\Tencent\QQ\Temp",                                        
    "$env:TEMP",                                                           
    # --- Developer Environment ---
    "$env:USERPROFILE\.cache\huggingface", "$env:USERPROFILE\.cache\torch", 
    "$env:LOCALAPPDATA\pip\Cache", "$env:USERPROFILE\.conda\pkgs",
    "$env:APPDATA\Obsidian\Cache", "$env:LOCALAPPDATA\npm-cache"
)
foreach ($cp in $cachePaths) { if (Test-Path $cp) { $cacheList += $cp } }

# ==========================================
# Phase 3: Generate AI Review TXT
# ==========================================
$txtPath = "$env:USERPROFILE\Desktop\AI-Review-List.txt"
$header = @"
# === Windows Deep Clean AI Review List ===
# Instructions:
# 1. Copy all content below to a large language model (e.g., Gemini, ChatGPT, etc.).
# 2. Paste the AI-filtered [pure path list] back into this file.
# 3. Save (Ctrl+S) and close Notepad, the script will proceed to final confirmation.

[SECTION_STARTUP_ITEMS] (Suspected rogue software/viruses/adware)
"@
$header | Out-File -FilePath $txtPath -Encoding UTF8
($startupList | Select-Object -Unique) | Out-File -FilePath $txtPath -Append -Encoding UTF8

"`n[SECTION_APP_CACHES] (High-frequency app and environment redundant caches)" | Out-File -FilePath $txtPath -Append -Encoding UTF8
($cacheList | Select-Object -Unique) | Out-File -FilePath $txtPath -Append -Encoding UTF8

# ==========================================
# Phase 4: Interactive Prompt and AI Integration
# ==========================================
$msg = "[AI Assisted Decision Phase]`n`n1. 'AI-Review-List.txt' has been generated on desktop.`n2. Copy its content to an AI model for diagnosis.`n3. Paste confirmed junk/rogue paths back and save.`n`nAfter closing Notepad, the final verification window will appear!"
[System.Windows.Forms.MessageBox]::Show($msg, "Anti-Mistake Review", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

Start-Process "notepad.exe" -ArgumentList $txtPath -Wait

# ==========================================
# Phase 5: Read AI Results and Visual Final Confirmation
# ==========================================
if (Test-Path $txtPath) {
    $finalDeleteList = Get-Content $txtPath | Where-Object { $_.Trim() -ne "" -and $_ -notmatch "^#" -and $_ -notmatch "\[SECTION_" }
    
    if ($finalDeleteList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No valid paths in the list, operation cancelled.", "Notice")
        Remove-Item $txtPath -Force
        Exit
    }

    $confirmForm = New-Object System.Windows.Forms.Form
    $confirmForm.Text = "Final Check: Items to be shredded"
    $confirmForm.Size = New-Object System.Drawing.Size(700, 500)
    $confirmForm.StartPosition = "CenterScreen"
    $confirmForm.FormBorderStyle = "FixedDialog"
    $confirmForm.MaximizeBox = $false

    $cLabel = New-Object System.Windows.Forms.Label
    $cLabel.Text = "⚠️ Please final confirm the following $($finalDeleteList.Count) targets. Once you click [Start Shredding], files cannot be recovered!"
    $cLabel.Location = New-Object System.Drawing.Point(20, 15)
    $cLabel.Size = New-Object System.Drawing.Size(650, 30)
    $cLabel.ForeColor = [System.Drawing.Color]::Red
    $cLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei", 10, [System.Drawing.FontStyle]::Bold)
    $confirmForm.Controls.Add($cLabel)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(20, 50)
    $listBox.Size = New-Object System.Drawing.Size(640, 320)
    $listBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    foreach ($line in $finalDeleteList) { $listBox.Items.Add($line) | Out-Null }
    $confirmForm.Controls.Add($listBox)

    $btnFinal = New-Object System.Windows.Forms.Button
    $btnFinal.Text = "⚡ Confirm and Start Shredding"
    $btnFinal.Location = New-Object System.Drawing.Point(250, 390)
    $btnFinal.Size = New-Object System.Drawing.Size(200, 45)
    $btnFinal.BackColor = [System.Drawing.Color]::DarkRed
    $btnFinal.ForeColor = [System.Drawing.Color]::White
    $btnFinal.Font = New-Object System.Drawing.Font("Microsoft YaHei", 10, [System.Drawing.FontStyle]::Bold)

    $isConfirmed = $false
    $btnFinal.Add_Click({
        $script:isConfirmed = $true
        $confirmForm.Close()
    })
    $confirmForm.Controls.Add($btnFinal)

    $confirmForm.ShowDialog() | Out-Null

    # ==========================================
    # Phase 6: Execute Final Shredding (with progress bar)
    # ==========================================
    if ($isConfirmed) {
        Write-Host "`n[!] Executing ultimate shredding process..." -ForegroundColor Red
        $successCount = 0
        $totalItems = $finalDeleteList.Count
        $currentIndex = 0

        foreach ($target in $finalDeleteList) {
            $target = $target.Trim()
            $currentIndex++
            
            if (Test-Path $target) {
                $percent = [math]::Round(($currentIndex / $totalItems) * 100)
                $name = [System.IO.Path]::GetFileName($target)

                Write-Progress -Activity "⚡ Ultimate Shredding in Progress" -Status "Shredding [$currentIndex/$totalItems]: $name" -PercentComplete $percent -CurrentOperation "Overall Progress: $percent%"

                try {
                    $processName = [System.IO.Path]::GetFileNameWithoutExtension($target)
                    Stop-Process -Name $processName -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Milliseconds 200
                    
                    Remove-Item -Path $target -Force -Recurse -ErrorAction SilentlyContinue
                    $successCount++
                    Write-Host " [Success] Shredded: $target" -ForegroundColor DarkGray
                } catch {
                    Write-Host " [Skipped] Cannot delete (may be system locked): $target" -ForegroundColor Yellow
                }
            }
        }
        
        Write-Progress -Activity "⚡ Ultimate Shredding in Progress" -Completed
        Remove-Item $txtPath -Force
        [System.Windows.Forms.MessageBox]::Show("System cleanup completed successfully!`nSuccessfully shredded $successCount redundant/rogue targets.", "Task Completed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("Operation cancelled, no verified files were deleted.", "Terminated")
        Remove-Item $txtPath -Force
    }
}