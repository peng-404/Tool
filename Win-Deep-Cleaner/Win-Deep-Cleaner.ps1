# ==========================================
# 0. Elevate Privileges (Admin rights required)
# ==========================================
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    try {
        Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -ErrorAction Stop
        exit
    } catch {
        Write-Host "Win Deep Cleaner requires administrator privileges. Please relaunch it and approve the UAC prompt." -ForegroundColor Yellow
        exit 1
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:DesktopPath = [Environment]::GetFolderPath("Desktop")
$script:ReviewFilePath = Join-Path $script:DesktopPath "AI-Review-List.txt"
$script:ManifestPath = Join-Path $script:DesktopPath "AI-Review-Metadata.json"
$script:ReportRoot = Join-Path $script:DesktopPath "Win-Deep-Cleaner-Reports"
$script:BackupRoot = Join-Path $script:DesktopPath "Win-Deep-Cleaner-Backups"

$script:AccentColor = [System.Drawing.Color]::FromArgb(34, 87, 74)
$script:DangerColor = [System.Drawing.Color]::FromArgb(136, 32, 32)
$script:CanvasColor = [System.Drawing.Color]::FromArgb(245, 240, 231)
$script:CardColor = [System.Drawing.Color]::FromArgb(255, 251, 245)
$script:TextColor = [System.Drawing.Color]::FromArgb(36, 36, 36)
$script:MutedColor = [System.Drawing.Color]::FromArgb(106, 96, 84)

function Write-UiLog {
    param(
        [System.Windows.Forms.TextBox]$LogBox,
        [string]$Message
    )

    if (-not $LogBox) {
        return
    }

    $timestamp = Get-Date -Format "HH:mm:ss"
    $LogBox.AppendText("[$timestamp] $Message`r`n")
    $LogBox.SelectionStart = $LogBox.TextLength
    $LogBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-TaggedControlsState {
    param(
        [System.Windows.Forms.Control]$Parent,
        [bool]$Enabled
    )

    foreach ($child in $Parent.Controls) {
        if ($child.Tag -eq "action-button") {
            $child.Enabled = $Enabled
        }

        if ($child.Controls.Count -gt 0) {
            Set-TaggedControlsState -Parent $child -Enabled $Enabled
        }
    }
}

function Set-UiBusyState {
    param(
        [System.Windows.Forms.Form]$Form,
        [bool]$IsBusy
    )

    $Form.UseWaitCursor = $IsBusy
    Set-TaggedControlsState -Parent $Form -Enabled (-not $IsBusy)
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-ProgressState {
    param(
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.Label]$ActivityLabel,
        [string]$StatusText,
        [string]$ActivityText,
        [ValidateSet("Idle", "Marquee", "Step")] [string]$Mode = "Idle",
        [int]$Current = 0,
        [int]$Maximum = 1
    )

    if ($StatusLabel) {
        $StatusLabel.Text = $StatusText
    }

    if ($ActivityLabel) {
        $ActivityLabel.Text = $ActivityText
    }

    if ($ProgressBar) {
        switch ($Mode) {
            "Marquee" {
                $ProgressBar.Style = "Marquee"
                $ProgressBar.MarqueeAnimationSpeed = 25
            }
            "Step" {
                if ($Maximum -lt 1) {
                    $Maximum = 1
                }

                $ProgressBar.Style = "Continuous"
                $ProgressBar.MarqueeAnimationSpeed = 0
                $ProgressBar.Value = 0
                $ProgressBar.Maximum = $Maximum
                $ProgressBar.Value = [Math]::Min([Math]::Max($Current, 0), $Maximum)
            }
            default {
                $ProgressBar.Style = "Continuous"
                $ProgressBar.MarqueeAnimationSpeed = 0
                $ProgressBar.Value = 0
                $ProgressBar.Maximum = 1
            }
        }
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        [void](New-Item -ItemType Directory -Path $Path -Force)
    }
}

function Get-NormalizedPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    try {
        $fullPath = [System.IO.Path]::GetFullPath($Path)
    } catch {
        $fullPath = $Path
    }

    return $fullPath.Trim().TrimEnd('\').ToLowerInvariant()
}

function Get-PathGuard {
    param([string]$Path)

    $normalized = Get-NormalizedPath -Path $Path
    if (-not $normalized) {
        return [pscustomobject]@{
            Blocked = $true
            Reason = "Path is empty or invalid"
        }
    }

    $systemDriveRoot = Get-NormalizedPath -Path "$env:SystemDrive\"
    $windowsRoot = Get-NormalizedPath -Path $env:windir
    $driverRoots = @(
        (Join-Path $env:windir "System32\drivers"),
        (Join-Path $env:windir "System32\DriverStore"),
        (Join-Path $env:SystemDrive "Drivers")
    ) | ForEach-Object { Get-NormalizedPath -Path $_ }

    $otherProtectedRoots = @(
        (Join-Path $env:SystemDrive "System Volume Information"),
        (Join-Path $env:SystemDrive '$Recycle.Bin'),
        (Join-Path $env:ProgramData "Microsoft\Windows")
    ) | Where-Object { $_ } | ForEach-Object { Get-NormalizedPath -Path $_ }

    if ($normalized -eq (Get-NormalizedPath -Path $env:SystemDrive)) {
        return [pscustomobject]@{
            Blocked = $true
            Reason = "System drive root is blocked"
        }
    }

    if ($normalized -eq $systemDriveRoot.TrimEnd('\').ToLowerInvariant()) {
        return [pscustomobject]@{
            Blocked = $true
            Reason = "System drive root is blocked"
        }
    }

    if ($normalized -eq $windowsRoot -or $normalized.StartsWith("$windowsRoot\")) {
        return [pscustomobject]@{
            Blocked = $true
            Reason = "Windows system directory is blocked"
        }
    }

    foreach ($driverRoot in $driverRoots) {
        if ($driverRoot -and ($normalized -eq $driverRoot -or $normalized.StartsWith("$driverRoot\"))) {
            return [pscustomobject]@{
                Blocked = $true
                Reason = "Driver-related directory is blocked"
            }
        }
    }

    foreach ($protectedRoot in $otherProtectedRoots) {
        if ($protectedRoot -and ($normalized -eq $protectedRoot -or $normalized.StartsWith("$protectedRoot\"))) {
            return [pscustomobject]@{
                Blocked = $true
                Reason = "Protected system directory is blocked"
            }
        }
    }

    return [pscustomobject]@{
        Blocked = $false
        Reason = ""
    }
}

function Get-PathFromLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $null
    }

    $cleaned = $Line.Trim()
    $cleaned = $cleaned -replace '^[\-\*\u2022]+\s*', ''
    $cleaned = $cleaned -replace '^\d+[\.\)]\s*', ''
    $cleaned = $cleaned.Trim('"', "'", '`')
    $cleaned = $cleaned -replace '[\u201C\u201D\u2018\u2019]', ''

    if ($cleaned -match '(?i)([a-z]:\\.*)$') {
        $path = $matches[1].Trim()
        $path = $path.Trim('"', "'", '`')
        $path = $path -replace '[\u201C\u201D\u2018\u2019]', ''
        $path = $path -replace '["''`]+$', ''
        $path = $path -replace '[,\uFF0C;\uFF1B\u3002]+$', ''
        return $path
    }

    return $null
}

function Format-Bytes {
    param([nullable[long]]$Bytes)

    if (-not $Bytes -or $Bytes -le 0) {
        return "-"
    }

    $size = [double]$Bytes
    $units = @("B", "KB", "MB", "GB", "TB")
    $unitIndex = 0
    while ($size -ge 1024 -and $unitIndex -lt ($units.Count - 1)) {
        $size = $size / 1024
        $unitIndex++
    }

    return "{0:N1} {1}" -f $size, $units[$unitIndex]
}

function Get-ShortDisplayPath {
    param(
        [string]$Path,
        [int]$MaxLength = 72
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or $Path.Length -le $MaxLength) {
        return $Path
    }

    $tailLength = [Math]::Max(20, $MaxLength - 4)
    return "...\" + $Path.Substring($Path.Length - $tailLength)
}

function Format-DateValue {
    param([Nullable[datetime]]$Value)

    if (-not $Value) {
        return "-"
    }

    return $Value.ToString("yyyy-MM-dd")
}

function Get-LastActivityTime {
    param($Item)

    if (-not $Item) {
        return $null
    }

    $lastWrite = $Item.LastWriteTime
    $lastAccess = $Item.LastAccessTime

    if ($lastAccess -and $lastAccess -gt $lastWrite) {
        return $lastAccess
    }

    return $lastWrite
}

function Get-ItemKindLabel {
    param($Item)

    if (-not $Item) {
        return "Unknown"
    }

    if ($Item.PSIsContainer) {
        return "Folder"
    }

    return "File"
}

function Get-FileSizeIfAvailable {
    param($Item)

    if ($Item -and (-not $Item.PSIsContainer)) {
        return [long]$Item.Length
    }

    return $null
}

function New-CandidateRecord {
    param(
        [string]$Path,
        [string]$CategoryLabel,
        [string]$SourceLabel,
        [string]$Reason,
        [object]$SelectedByDefault = $null
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $item = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    if (-not $item) {
        return $null
    }

    $guard = Get-PathGuard -Path $item.FullName
    $lastActivity = Get-LastActivityTime -Item $item
    $selectedState = (-not $guard.Blocked)
    if ($null -ne $SelectedByDefault) {
        $selectedState = [bool]$SelectedByDefault
    }

    if ($guard.Blocked) {
        $selectedState = $false
    }

    return [pscustomobject]@{
        Id = [guid]::NewGuid().Guid
        Path = $item.FullName
        NormalizedPath = Get-NormalizedPath -Path $item.FullName
        CategoryLabel = $CategoryLabel
        SourceLabel = $SourceLabel
        Reason = $Reason
        ItemKind = Get-ItemKindLabel -Item $item
        SizeBytes = Get-FileSizeIfAvailable -Item $item
        LastActivity = $lastActivity
        LastWriteTime = $item.LastWriteTime
        ExistsNow = $true
        Blocked = $guard.Blocked
        BlockReason = $guard.Reason
        SelectedByDefault = $selectedState
    }
}

function Merge-CandidateRecords {
    param([object[]]$Candidates)

    $map = @{}
    foreach ($candidate in $Candidates) {
        if (-not $candidate) {
            continue
        }

        $key = $candidate.NormalizedPath
        if (-not $key) {
            continue
        }

        if ($map.ContainsKey($key)) {
            $existing = $map[$key]
            $existing.SourceLabel = (($existing.SourceLabel -split ' / ') + ($candidate.SourceLabel -split ' / ') | Where-Object { $_ } | Select-Object -Unique) -join " / "
            $existing.Reason = (($existing.Reason -split ' / ') + ($candidate.Reason -split ' / ') | Where-Object { $_ } | Select-Object -Unique) -join " / "
            if (($candidate.LastActivity -and -not $existing.LastActivity) -or ($candidate.LastActivity -and $existing.LastActivity -and $candidate.LastActivity -gt $existing.LastActivity)) {
                $existing.LastActivity = $candidate.LastActivity
            }
            if (-not $existing.SizeBytes -and $candidate.SizeBytes) {
                $existing.SizeBytes = $candidate.SizeBytes
            }
            if (-not $candidate.SelectedByDefault) {
                $existing.SelectedByDefault = $false
            }
            $existing.Blocked = $existing.Blocked -or $candidate.Blocked
            if ($candidate.BlockReason) {
                $existing.BlockReason = (($existing.BlockReason -split ' / ') + ($candidate.BlockReason -split ' / ') | Where-Object { $_ } | Select-Object -Unique) -join " / "
            }
        } else {
            $map[$key] = $candidate
        }
    }

    return $map.Values | Sort-Object CategoryLabel, SourceLabel, Path
}

function Test-FileAttribute {
    param(
        $Item,
        [System.IO.FileAttributes]$Attribute
    )

    if (-not $Item) {
        return $false
    }

    return (($Item.Attributes -band $Attribute) -eq $Attribute)
}

function Test-PathStartsWith {
    param(
        [string]$Path,
        [string]$Root
    )

    $normalizedPath = Get-NormalizedPath -Path $Path
    $normalizedRoot = Get-NormalizedPath -Path $Root

    if (-not $normalizedPath -or -not $normalizedRoot) {
        return $false
    }

    if ($normalizedPath -eq $normalizedRoot) {
        return $true
    }

    return $normalizedPath.StartsWith("$normalizedRoot\")
}

function Get-RecursiveFileItems {
    param(
        [string]$Root,
        [int]$MaxDepth = 5,
        [string[]]$ExcludedDirectoryNames = @()
    )

    $items = @()
    if (-not (Test-Path -LiteralPath $Root)) {
        return $items
    }

    $excluded = @($ExcludedDirectoryNames | ForEach-Object { $_.ToLowerInvariant() })
    $stack = New-Object System.Collections.ArrayList
    [void]$stack.Add([pscustomobject]@{
        Path = $Root
        Depth = 0
    })

    while ($stack.Count -gt 0) {
        $current = $stack[$stack.Count - 1]
        $stack.RemoveAt($stack.Count - 1)

        try {
            $children = Get-ChildItem -LiteralPath $current.Path -Force -ErrorAction Stop
        } catch {
            continue
        }

        foreach ($child in $children) {
            if ($child.PSIsContainer) {
                if (Test-FileAttribute -Item $child -Attribute ([System.IO.FileAttributes]::ReparsePoint)) {
                    continue
                }

                if ($excluded -contains $child.Name.ToLowerInvariant()) {
                    continue
                }

                if ($current.Depth -lt $MaxDepth) {
                    [void]$stack.Add([pscustomobject]@{
                        Path = $child.FullName
                        Depth = ($current.Depth + 1)
                    })
                }

                continue
            }

            $items += $child
        }
    }

    return $items
}

function Test-SystemLikeFileName {
    param([string]$FileName)

    if ([string]::IsNullOrWhiteSpace($FileName)) {
        return $false
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName).ToLowerInvariant()
    return $baseName -match '^(svchost|explorer|lsass|csrss|services|winlogon|rundll32|taskhostw|conhost|spoolsv|dllhost|chrome|msedge|firefox|setup|install|update|runtime|security|defender)(?:\d+)?$'
}

function Get-StartupCandidateRecords {
    $records = @()
    $regex = '(?i)([a-z]:\\[^\*?"<>\|]+\.(?:exe|bat|cmd|vbs|dll|ps1))'

    try {
        $startupItems = Get-CimInstance Win32_StartupCommand
        foreach ($item in $startupItems) {
            if ($item.Command -match $regex) {
                $categoryLabel = "Startup Item"
                $sourceLabel = "Third-party Startup Entry"
                $reason = "Detected from the system startup command list"

                if (
                    (Test-PathStartsWith -Path $matches[1] -Root $env:TEMP) -or
                    (Test-PathStartsWith -Path $matches[1] -Root $env:APPDATA) -or
                    (Test-PathStartsWith -Path $matches[1] -Root $env:LOCALAPPDATA)
                ) {
                    $categoryLabel = "Suspicious Startup Item"
                    $sourceLabel = "Startup Entry In User-Writable Location"
                    $reason = "Startup command points to Temp or AppData, which is a common location for unwanted auto-start programs"
                }

                $record = New-CandidateRecord -Path $matches[1] -CategoryLabel $categoryLabel -SourceLabel $sourceLabel -Reason $reason -SelectedByDefault $false
                if ($record) {
                    $records += $record
                }
            }
        }
    } catch {
        throw "Failed to read startup items: $($_.Exception.Message)"
    }

    return $records
}

function Get-CacheCandidateRecords {
    $definitions = @(
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\Cache_Data"; Category = "App Cache"; Source = "Edge Browser Cache"; Reason = "Default browser cache directory" },
        @{ Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\Cache_Data"; Category = "App Cache"; Source = "Chrome Browser Cache"; Reason = "Default browser cache directory" },
        @{ Path = "$env:USERPROFILE\Documents\WeChat Files\Applet"; Category = "App Cache"; Source = "WeChat Mini Program Cache"; Reason = "Applet cache inside the WeChat files directory" },
        @{ Path = "$env:USERPROFILE\Documents\WeChat Files\FileStorage\Cache"; Category = "App Cache"; Source = "WeChat File Cache"; Reason = "WeChat FileStorage cache directory" },
        @{ Path = "$env:APPDATA\Tencent\QQ\Temp"; Category = "App Cache"; Source = "QQ Temporary Files"; Reason = "QQ temporary directory" },
        @{ Path = "$env:TEMP"; Category = "App Cache"; Source = "System Temporary Directory"; Reason = "Current user's temp directory" },
        @{ Path = "$env:USERPROFILE\.cache\huggingface"; Category = "Dev Cache"; Source = "HuggingFace Cache"; Reason = "Model download cache directory" },
        @{ Path = "$env:USERPROFILE\.cache\torch"; Category = "Dev Cache"; Source = "PyTorch Cache"; Reason = "Torch default cache directory" },
        @{ Path = "$env:LOCALAPPDATA\pip\Cache"; Category = "Dev Cache"; Source = "pip Cache"; Reason = "pip package cache directory" },
        @{ Path = "$env:USERPROFILE\.conda\pkgs"; Category = "Dev Cache"; Source = "Conda Package Cache"; Reason = "Conda pkgs cache directory" },
        @{ Path = "$env:APPDATA\Obsidian\Cache"; Category = "App Cache"; Source = "Obsidian Cache"; Reason = "Obsidian cache directory" },
        @{ Path = "$env:LOCALAPPDATA\npm-cache"; Category = "Dev Cache"; Source = "npm Cache"; Reason = "npm default cache directory" }
    )

    $records = @()
    foreach ($definition in $definitions) {
        $record = New-CandidateRecord -Path $definition.Path -CategoryLabel $definition.Category -SourceLabel $definition.Source -Reason $definition.Reason
        if ($record) {
            $records += $record
        }
    }

    return $records
}

function Get-InactiveFileCandidateRecords {
    param([int]$InactiveDays = 180)

    $archiveExtensions = @(".zip", ".rar", ".7z", ".iso", ".msi", ".exe", ".cab", ".dmp", ".log", ".tmp")
    $definitions = @(
        @{ Root = (Join-Path $env:USERPROFILE "Downloads"); Category = "Inactive Files"; Source = "Recursive Scan: Downloads"; MinSizeBytes = 50MB; Extensions = $archiveExtensions; MaxDepth = 6; InactiveDays = $InactiveDays; SelectedByDefault = $true },
        @{ Root = (Join-Path $env:USERPROFILE "Desktop"); Category = "Inactive Files"; Source = "Recursive Scan: Desktop"; MinSizeBytes = 100MB; Extensions = $archiveExtensions; MaxDepth = 4; InactiveDays = $InactiveDays; SelectedByDefault = $true },
        @{ Root = (Join-Path $env:USERPROFILE "Documents"); Category = "Inactive Files"; Source = "Recursive Scan: Documents"; MinSizeBytes = 20MB; Extensions = $archiveExtensions; MaxDepth = 6; InactiveDays = $InactiveDays; SelectedByDefault = $true },
        @{ Root = $env:TEMP; Category = "Stale Temp Files"; Source = "Recursive Scan: Temp Directory"; MinSizeBytes = 10MB; Extensions = $archiveExtensions; MaxDepth = 4; InactiveDays = 30; SelectedByDefault = $true },
        @{ Root = (Join-Path $env:USERPROFILE "Videos"); Category = "Large Old Files"; Source = "Optional Large File Scan: Videos"; MinSizeBytes = 1GB; Extensions = @(); MaxDepth = 4; InactiveDays = 365; SelectedByDefault = $false }
    )

    $records = @()
    foreach ($definition in $definitions) {
        if (-not (Test-Path -LiteralPath $definition.Root)) {
            continue
        }

        $files = Get-RecursiveFileItems -Root $definition.Root -MaxDepth $definition.MaxDepth

        foreach ($file in $files) {
            $lastActivity = Get-LastActivityTime -Item $file
            if (-not $lastActivity) {
                continue
            }

            $cutoff = (Get-Date).AddDays(-[int]$definition.InactiveDays)
            $extension = $file.Extension.ToLowerInvariant()
            $isInterestingExtension = $definition.Extensions -contains $extension
            $isLargeFile = $file.Length -ge $definition.MinSizeBytes

            if ($lastActivity -le $cutoff -and ($isInterestingExtension -or $isLargeFile)) {
                $reason = "File in a user-safe scan area is inactive, last activity $(Format-DateValue -Value $lastActivity)"
                $record = New-CandidateRecord -Path $file.FullName -CategoryLabel $definition.Category -SourceLabel $definition.Source -Reason $reason -SelectedByDefault $definition.SelectedByDefault
                if ($record) {
                    $records += $record
                }
            }
        }
    }

    return $records
}

function Get-SuspiciousFileCandidateRecords {
    param([int]$RecentDays = 30)

    $definitions = @(
        @{ Root = (Join-Path $env:USERPROFILE "Downloads"); Source = "Suspicious File Scan: Downloads"; MaxDepth = 6; Excluded = @() },
        @{ Root = (Join-Path $env:USERPROFILE "Desktop"); Source = "Suspicious File Scan: Desktop"; MaxDepth = 4; Excluded = @() },
        @{ Root = (Join-Path $env:USERPROFILE "Documents"); Source = "Suspicious File Scan: Documents"; MaxDepth = 6; Excluded = @() },
        @{ Root = $env:TEMP; Source = "Suspicious File Scan: Temp Directory"; MaxDepth = 4; Excluded = @() },
        @{ Root = (Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup"); Source = "Suspicious File Scan: User Startup Folder"; MaxDepth = 2; Excluded = @() }
    )

    $suspiciousExtensions = @(".exe", ".dll", ".bat", ".cmd", ".vbs", ".vbe", ".js", ".jse", ".wsf", ".wsh", ".ps1", ".scr", ".com", ".pif", ".jar", ".lnk")
    $doubleExtensionPattern = '(?i)\.(pdf|docx?|xlsx?|pptx?|txt|jpg|jpeg|png|gif|zip|rar|7z)\.(exe|scr|js|jse|vbs|vbe|bat|cmd|com|pif|ps1)$'
    $recentCutoff = (Get-Date).AddDays(-$RecentDays)
    $records = @()

    foreach ($definition in $definitions) {
        if (-not (Test-Path -LiteralPath $definition.Root)) {
            continue
        }

        $files = Get-RecursiveFileItems -Root $definition.Root -MaxDepth $definition.MaxDepth -ExcludedDirectoryNames $definition.Excluded
        foreach ($file in $files) {
            $extension = $file.Extension.ToLowerInvariant()
            $looksLikeDoubleExtension = $file.Name -match $doubleExtensionPattern
            if (($suspiciousExtensions -notcontains $extension) -and -not $looksLikeDoubleExtension) {
                continue
            }

            $reasons = @()
            $lastActivity = Get-LastActivityTime -Item $file

            if ($looksLikeDoubleExtension) {
                $reasons += "Double-extension executable or script name"
            }

            if (Test-FileAttribute -Item $file -Attribute ([System.IO.FileAttributes]::Hidden)) {
                $reasons += "Hidden executable or script file"
            }

            if (Test-SystemLikeFileName -FileName $file.Name) {
                $reasons += "System-like file name inside a user-writable location"
            }

            if ($lastActivity -and $lastActivity -ge $recentCutoff) {
                $reasons += "Recently created or active executable/script in a user profile location"
            }

            if (Test-PathStartsWith -Path $file.FullName -Root $env:TEMP) {
                $reasons += "Executable or script stored in Temp"
            }

            if (Test-PathStartsWith -Path $file.FullName -Root (Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup")) {
                $reasons += "Executable or script stored in the user Startup folder"
            }

            if ($file.Length -le 2MB -and $extension -in @(".exe", ".dll", ".scr", ".com")) {
                $reasons += "Small executable payload in a user-writable location"
            }

            if ($reasons.Count -eq 0) {
                continue
            }

            $reason = ($reasons | Select-Object -Unique) -join " / "
            $record = New-CandidateRecord -Path $file.FullName -CategoryLabel "Suspicious Item" -SourceLabel $definition.Source -Reason $reason -SelectedByDefault $false
            if ($record) {
                $records += $record
            }
        }
    }

    return $records
}

function Wait-ExternalProcess {
    param(
        [string]$FilePath,
        [string]$ArgumentList,
        [System.Windows.Forms.TextBox]$LogBox,
        [string]$StartMessage,
        [string]$EndMessage,
        [bool]$ContinueOnError = $false
    )

    try {
        Write-UiLog -LogBox $LogBox -Message $StartMessage
        $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -WindowStyle Normal -ErrorAction Stop

        while (-not $process.HasExited) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 250
        }

        if ($process.ExitCode -ne 0) {
            $message = "$EndMessage Exit code: $($process.ExitCode)"
            if ($ContinueOnError) {
                Write-UiLog -LogBox $LogBox -Message "Warning: $message"
                return $process.ExitCode
            }

            throw "$FilePath failed with exit code $($process.ExitCode)"
        }

        Write-UiLog -LogBox $LogBox -Message $EndMessage
        return $process.ExitCode
    } catch {
        if ($ContinueOnError) {
            Write-UiLog -LogBox $LogBox -Message "Warning: $FilePath failed, but the program will continue. $($_.Exception.Message)"
            return $null
        }

        throw
    }
}

function Save-ReviewArtifacts {
    param([object[]]$Candidates)

    $safeCandidates = @($Candidates | Where-Object { -not $_.Blocked })
    $blockedCandidates = @($Candidates | Where-Object { $_.Blocked })

    $header = @(
        "# === Windows Deep Clean AI Review List ==="
        "# This is the candidate list generated in phase one. You can send it to an AI model or review it manually."
        "# Phase two will show a checkbox-based confirmation screen. Only checked and confirmed items will be processed."
        "# You can delete lines you do not want, or keep the full tagged line."
        "# Phase two automatically extracts the path at the end of each line."
        "# Suspicious items are heuristic findings, not a malware verdict."
        "#"
        "# Auto-blocked high-risk paths: $($blockedCandidates.Count)"
        "# They will not be processed in phase two."
        ""
    )

    $header | Out-File -FilePath $script:ReviewFilePath -Encoding UTF8

    foreach ($group in ($safeCandidates | Group-Object CategoryLabel)) {
        "[SECTION_$($group.Name)]" | Out-File -FilePath $script:ReviewFilePath -Append -Encoding UTF8
        foreach ($candidate in $group.Group) {
            $line = "[Source: $($candidate.SourceLabel)] [Type: $($candidate.ItemKind)] [Last Activity: $(Format-DateValue -Value $candidate.LastActivity)] [Size: $(Format-Bytes -Bytes $candidate.SizeBytes)] [Reason: $($candidate.Reason)] $($candidate.Path)"
            $line | Out-File -FilePath $script:ReviewFilePath -Append -Encoding UTF8
        }
        "" | Out-File -FilePath $script:ReviewFilePath -Append -Encoding UTF8
    }

    $manifest = [pscustomobject]@{
        CreatedAt = (Get-Date).ToString("s")
        ReviewFilePath = $script:ReviewFilePath
        SafeCount = $safeCandidates.Count
        BlockedCount = $blockedCandidates.Count
        Candidates = $Candidates
    }

    $manifest | ConvertTo-Json -Depth 6 | Out-File -FilePath $script:ManifestPath -Encoding UTF8

    return [pscustomobject]@{
        SafeCandidates = $safeCandidates
        BlockedCandidates = $blockedCandidates
    }
}

function Load-CandidateManifest {
    if (-not (Test-Path -LiteralPath $script:ManifestPath)) {
        return @()
    }

    try {
        $raw = Get-Content -LiteralPath $script:ManifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return @()
    }

    $items = @()
    foreach ($candidate in $raw.Candidates) {
        $sizeBytes = $null
        if ($candidate.SizeBytes) {
            $sizeBytes = [long]$candidate.SizeBytes
        }

        $lastActivity = $null
        if ($candidate.LastActivity) {
            $lastActivity = [datetime]$candidate.LastActivity
        }

        $lastWriteTime = $null
        if ($candidate.LastWriteTime) {
            $lastWriteTime = [datetime]$candidate.LastWriteTime
        }

        $items += [pscustomobject]@{
            Id = $candidate.Id
            Path = $candidate.Path
            NormalizedPath = Get-NormalizedPath -Path $candidate.Path
            CategoryLabel = $candidate.CategoryLabel
            SourceLabel = $candidate.SourceLabel
            Reason = $candidate.Reason
            ItemKind = $candidate.ItemKind
            SizeBytes = $sizeBytes
            LastActivity = $lastActivity
            LastWriteTime = $lastWriteTime
            ExistsNow = Test-Path -LiteralPath $candidate.Path
            Blocked = [bool]$candidate.Blocked
            BlockReason = $candidate.BlockReason
            SelectedByDefault = [bool]$candidate.SelectedByDefault
        }
    }

    return $items
}

function Get-ReviewedPaths {
    if (-not (Test-Path -LiteralPath $script:ReviewFilePath)) {
        return @()
    }

    try {
        $lines = Get-Content -LiteralPath $script:ReviewFilePath -ErrorAction Stop
    } catch {
        return @()
    }

    $paths = foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed -and $trimmed -notmatch '^#' -and $trimmed -notmatch '^\[SECTION_') {
            $path = Get-PathFromLine -Line $trimmed
            if ($path) {
                $path
            }
        }
    }

    return $paths | Select-Object -Unique
}

function New-AdHocReviewedCandidate {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{
            Id = [guid]::NewGuid().Guid
            Path = $Path
            NormalizedPath = Get-NormalizedPath -Path $Path
            CategoryLabel = "User Added"
            SourceLabel = "Path Added Manually"
            Reason = "This path came from the review file but is not present in the phase-one metadata"
            ItemKind = "Unknown"
            SizeBytes = $null
            LastActivity = $null
            LastWriteTime = $null
            ExistsNow = $false
            Blocked = $false
            BlockReason = ""
            SelectedByDefault = $true
        }
    }

    $record = New-CandidateRecord -Path $Path -CategoryLabel "User Added" -SourceLabel "Path Added Manually" -Reason "This path came from the review file but is not present in the phase-one metadata"
    return $record
}

function Resolve-ReviewedCandidates {
    param(
        [string[]]$ReviewedPaths,
        [object[]]$ManifestCandidates
    )

    $manifestMap = @{}
    foreach ($candidate in $ManifestCandidates) {
        if ($candidate.NormalizedPath) {
            $manifestMap[$candidate.NormalizedPath] = $candidate
        }
    }

    $resolved = @()
    $blockedFromReview = @()

    foreach ($path in ($ReviewedPaths | Select-Object -Unique)) {
        $normalized = Get-NormalizedPath -Path $path
        if (-not $normalized) {
            continue
        }

        if ($manifestMap.ContainsKey($normalized)) {
            $candidate = $manifestMap[$normalized]
        } else {
            $candidate = New-AdHocReviewedCandidate -Path $path
        }

        $guard = Get-PathGuard -Path $candidate.Path
        $candidate.Blocked = $guard.Blocked
        $candidate.BlockReason = $guard.Reason
        $candidate.ExistsNow = Test-Path -LiteralPath $candidate.Path

        if ($candidate.Blocked) {
            $blockedFromReview += $candidate
        } else {
            $resolved += $candidate
        }
    }

    return [pscustomobject]@{
        Allowed = $resolved
        Blocked = $blockedFromReview
    }
}

function Stop-TargetProcessIfNeeded {
    param([object]$Candidate)

    if (-not $Candidate.ExistsNow) {
        return
    }

    if ($Candidate.ItemKind -ne "File") {
        return
    }

    $extension = [System.IO.Path]::GetExtension($Candidate.Path).ToLowerInvariant()
    if ($extension -notin @(".exe", ".bat", ".cmd", ".ps1", ".vbs")) {
        return
    }

    $processName = [System.IO.Path]::GetFileName($Candidate.Path)
    $normalizedTargetPath = Get-NormalizedPath -Path $Candidate.Path
    if (-not $processName -or -not $normalizedTargetPath) {
        return
    }

    try {
        $escapedProcessName = $processName.Replace("'", "''")
        $matchingProcesses = Get-CimInstance Win32_Process -Filter "Name = '$escapedProcessName'" -ErrorAction Stop
    } catch {
        return
    }

    foreach ($process in $matchingProcesses) {
        $processPath = Get-NormalizedPath -Path $process.ExecutablePath
        if (-not $processPath -or $processPath -ne $normalizedTargetPath) {
            continue
        }

        try {
            Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop
            Start-Sleep -Milliseconds 150
        } catch {
            continue
        }
    }
}

function Backup-Target {
    param(
        [object]$Candidate,
        [string]$SessionBackupPath,
        [int]$Index
    )

    $slot = Join-Path $SessionBackupPath ("item-{0:D3}" -f $Index)
    Ensure-Directory -Path $slot
    Copy-Item -LiteralPath $Candidate.Path -Destination $slot -Recurse -Force -ErrorAction Stop
    return $slot
}

function Remove-TargetWithMode {
    param(
        [object]$Candidate,
        [bool]$UseRecycleBin
    )

    $item = Get-Item -LiteralPath $Candidate.Path -Force -ErrorAction Stop
    if ($UseRecycleBin) {
        if ($item.PSIsContainer) {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
                $item.FullName,
                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
            )
        } else {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                $item.FullName,
                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
            )
        }
    } else {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
    }
}

function Save-DeletionReport {
    param(
        [string]$SessionId,
        [string]$ReviewSnapshotPath,
        [string]$BackupPath,
        [bool]$UseRecycleBin,
        [object[]]$Entries,
        [int]$BlockedCount
    )

    Ensure-Directory -Path $script:ReportRoot
    $reportBase = Join-Path $script:ReportRoot ("Cleanup-Report-{0}" -f $SessionId)
    $txtPath = "$reportBase.txt"
    $jsonPath = "$reportBase.json"

    $successCount = @($Entries | Where-Object { $_.Result -eq "Success" }).Count
    $skippedCount = @($Entries | Where-Object { $_.Result -ne "Success" }).Count
    $backupPathText = "-"
    if ($BackupPath) {
        $backupPathText = $BackupPath
    }

    $deletionModeText = "PermanentDelete"
    if ($UseRecycleBin) {
        $deletionModeText = "RecycleBin"
    }

    $lines = @(
        "Win Deep Cleaner Execution Report"
        "Session: $SessionId"
        "CreatedAt: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
        "ReviewFile: $ReviewSnapshotPath"
        "BackupPath: $backupPathText"
        "DeletionMode: $deletionModeText"
        "BlockedByGuard: $BlockedCount"
        "SuccessCount: $successCount"
        "SkippedCount: $skippedCount"
        ""
        "Details:"
    )

    foreach ($entry in $Entries) {
        $lines += "[{0}] [{1}] [{2}] {3} | {4}" -f $entry.Result, $entry.SourceLabel, $entry.CategoryLabel, $entry.Path, $entry.Message
    }

    $lines | Out-File -FilePath $txtPath -Encoding UTF8

    [pscustomobject]@{
        SessionId = $SessionId
        CreatedAt = (Get-Date).ToString("s")
        ReviewFile = $ReviewSnapshotPath
        BackupPath = $BackupPath
        UseRecycleBin = $UseRecycleBin
        BlockedByGuard = $BlockedCount
        Entries = $Entries
    } | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8

    return [pscustomobject]@{
        TextPath = $txtPath
        JsonPath = $jsonPath
    }
}

function Show-ExecutionPlannerDialog {
    param(
        [object[]]$Candidates,
        [object[]]$BlockedCandidates
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Phase Two: Review and Execute"
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.ClientSize = New-Object System.Drawing.Size(980, 640)
    $dialog.BackColor = $script:CardColor
    $dialog.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Phase Two Confirmation"
    $title.Location = New-Object System.Drawing.Point(24, 18)
    $title.Size = New-Object System.Drawing.Size(260, 28)
    $title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 15, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $script:DangerColor
    $dialog.Controls.Add($title)

    $summary = New-Object System.Windows.Forms.Label
    $summary.Text = "These are the items you kept in the review file. You can uncheck any item here. High-risk paths are blocked automatically and will not be executed."
    $summary.Location = New-Object System.Drawing.Point(24, 50)
    $summary.Size = New-Object System.Drawing.Size(920, 24)
    $summary.ForeColor = $script:TextColor
    $dialog.Controls.Add($summary)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(24, 88)
    $listView.Size = New-Object System.Drawing.Size(932, 390)
    $listView.View = "Details"
    $listView.CheckBoxes = $true
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.HideSelection = $false
    [void]$listView.Columns.Add("Source", 170)
    [void]$listView.Columns.Add("Category", 110)
    [void]$listView.Columns.Add("Type", 70)
    [void]$listView.Columns.Add("Last Activity", 95)
    [void]$listView.Columns.Add("Size", 90)
    [void]$listView.Columns.Add("Status", 90)
    [void]$listView.Columns.Add("Path", 290)
    $dialog.Controls.Add($listView)

    foreach ($candidate in $Candidates) {
        $statusText = "Missing"
        if ($candidate.ExistsNow) {
            $statusText = "Ready"
        }

        $item = New-Object System.Windows.Forms.ListViewItem($candidate.SourceLabel)
        [void]$item.SubItems.Add($candidate.CategoryLabel)
        [void]$item.SubItems.Add($candidate.ItemKind)
        [void]$item.SubItems.Add((Format-DateValue -Value $candidate.LastActivity))
        [void]$item.SubItems.Add((Format-Bytes -Bytes $candidate.SizeBytes))
        [void]$item.SubItems.Add($statusText)
        [void]$item.SubItems.Add($candidate.Path)
        $item.Tag = $candidate
        $item.Checked = $candidate.SelectedByDefault
        if (-not $candidate.ExistsNow) {
            $item.ForeColor = [System.Drawing.Color]::Gray
            $item.Checked = $false
        }
        [void]$listView.Items.Add($item)
    }

    $guardLabel = New-Object System.Windows.Forms.Label
    $guardLabel.Location = New-Object System.Drawing.Point(24, 488)
    $guardLabel.Size = New-Object System.Drawing.Size(932, 40)
    $guardLabel.ForeColor = $script:MutedColor
    if ($BlockedCandidates.Count -gt 0) {
        $blockedPreview = (($BlockedCandidates | Select-Object -First 3).Path) -join " ; "
        $guardLabel.Text = "$($BlockedCandidates.Count) high-risk path(s) were blocked and will not be executed. Example: $blockedPreview"
    } else {
        $guardLabel.Text = "No high-risk paths were detected in this phase-two run."
    }
    $dialog.Controls.Add($guardLabel)

    $backupCheck = New-Object System.Windows.Forms.CheckBox
    $backupCheck.Text = "Back up items before deletion"
    $backupCheck.Location = New-Object System.Drawing.Point(24, 538)
    $backupCheck.Size = New-Object System.Drawing.Size(250, 24)
    $backupCheck.Checked = $true
    $dialog.Controls.Add($backupCheck)

    $recycleCheck = New-Object System.Windows.Forms.CheckBox
    $recycleCheck.Text = "Prefer moving items to the Recycle Bin"
    $recycleCheck.Location = New-Object System.Drawing.Point(292, 538)
    $recycleCheck.Size = New-Object System.Drawing.Size(180, 24)
    $recycleCheck.Checked = $true
    $dialog.Controls.Add($recycleCheck)

    $reportLabel = New-Object System.Windows.Forms.Label
    $reportLabel.Text = "A TXT and JSON report will be generated automatically after execution."
    $reportLabel.Location = New-Object System.Drawing.Point(24, 566)
    $reportLabel.Size = New-Object System.Drawing.Size(420, 20)
    $reportLabel.ForeColor = $script:MutedColor
    $dialog.Controls.Add($reportLabel)

    $selectedLabel = New-Object System.Windows.Forms.Label
    $selectedLabel.Location = New-Object System.Drawing.Point(620, 538)
    $selectedLabel.Size = New-Object System.Drawing.Size(160, 20)
    $selectedLabel.ForeColor = $script:TextColor
    $dialog.Controls.Add($selectedLabel)

    $updateSelectedLabel = {
        $checkedCount = @($listView.Items | Where-Object { $_.Checked }).Count
        $selectedLabel.Text = "Selected: $checkedCount item(s)"
    }

    $listView.Add_ItemChecked({
        & $updateSelectedLabel
    })

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = "Select All Ready Items"
    $btnSelectAll.Location = New-Object System.Drawing.Point(620, 566)
    $btnSelectAll.Size = New-Object System.Drawing.Size(108, 32)
    $btnSelectAll.BackColor = [System.Drawing.Color]::FromArgb(228, 221, 210)
    $btnSelectAll.FlatStyle = "Flat"
    $btnSelectAll.Add_Click({
        foreach ($item in $listView.Items) {
            if ($item.SubItems[5].Text -eq "Ready") {
                $item.Checked = $true
            }
        }
    })
    $dialog.Controls.Add($btnSelectAll)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = "Clear Selection"
    $btnClear.Location = New-Object System.Drawing.Point(742, 566)
    $btnClear.Size = New-Object System.Drawing.Size(90, 32)
    $btnClear.BackColor = [System.Drawing.Color]::FromArgb(228, 221, 210)
    $btnClear.FlatStyle = "Flat"
    $btnClear.Add_Click({
        foreach ($item in $listView.Items) {
            $item.Checked = $false
        }
    })
    $dialog.Controls.Add($btnClear)

    $btnConfirm = New-Object System.Windows.Forms.Button
    $btnConfirm.Text = "Execute"
    $btnConfirm.Location = New-Object System.Drawing.Point(846, 566)
    $btnConfirm.Size = New-Object System.Drawing.Size(110, 32)
    $btnConfirm.BackColor = $script:DangerColor
    $btnConfirm.ForeColor = [System.Drawing.Color]::White
    $btnConfirm.FlatStyle = "Flat"
    $dialog.Controls.Add($btnConfirm)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(846, 604)
    $btnCancel.Size = New-Object System.Drawing.Size(110, 24)
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.Add_Click({
        $dialog.Tag = $null
        $dialog.Close()
    })
    $dialog.Controls.Add($btnCancel)

    $btnConfirm.Add_Click({
        $selected = @()
        foreach ($item in $listView.Items) {
            if ($item.Checked) {
                $selected += $item.Tag
            }
        }

        if ($selected.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No executable items are selected yet.",
                "Nothing Selected",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $dialog.Tag = [pscustomobject]@{
            SelectedCandidates = $selected
            CreateBackup = $backupCheck.Checked
            UseRecycleBin = $recycleCheck.Checked
        }
        $dialog.Close()
    })

    & $updateSelectedLabel
    [void]$dialog.ShowDialog()
    return $dialog.Tag
}

function Invoke-ReviewedDeletion {
    param(
        [object[]]$SelectedCandidates,
        [bool]$CreateBackup,
        [bool]$UseRecycleBin,
        [int]$BlockedCount,
        [System.Windows.Forms.TextBox]$LogBox,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.Label]$ActivityLabel
    )

    if ($SelectedCandidates.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "There are no items to process in this run.",
            "No Tasks",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $sessionId = Get-Date -Format "yyyyMMdd-HHmmss"
    $sessionBackupPath = $null
    if ($CreateBackup) {
        Ensure-Directory -Path $script:BackupRoot
        $sessionBackupPath = Join-Path $script:BackupRoot ("Backup-{0}" -f $sessionId)
        Ensure-Directory -Path $sessionBackupPath
    }

    $entries = @()
    $deletionModeLabel = "Permanent Delete"
    if ($UseRecycleBin) {
        $deletionModeLabel = "Recycle Bin"
    }

    Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase Two Running" -ActivityText "Initializing execution plan..." -Mode Step -Current 0 -Maximum $SelectedCandidates.Count
    Write-UiLog -LogBox $LogBox -Message "Phase two started: $($SelectedCandidates.Count) selected item(s), deletion mode $deletionModeLabel."

    for ($i = 0; $i -lt $SelectedCandidates.Count; $i++) {
        $candidate = $SelectedCandidates[$i]
        $candidate.ExistsNow = Test-Path -LiteralPath $candidate.Path
        $guard = Get-PathGuard -Path $candidate.Path
        $candidate.Blocked = $guard.Blocked
        $candidate.BlockReason = $guard.Reason

        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Processing: $($i + 1) / $($SelectedCandidates.Count)" -ActivityText "Working on: $(Get-ShortDisplayPath -Path $candidate.Path)" -Mode Step -Current ($i + 1) -Maximum $SelectedCandidates.Count

        if ($candidate.Blocked) {
            $entries += [pscustomobject]@{
                Path = $candidate.Path
                SourceLabel = $candidate.SourceLabel
                CategoryLabel = $candidate.CategoryLabel
                Result = "Blocked"
                Message = $candidate.BlockReason
                BackupPath = ""
            }
            Write-UiLog -LogBox $LogBox -Message "Blocked: $($candidate.Path) ($($candidate.BlockReason))"
            continue
        }

        if (-not $candidate.ExistsNow) {
            $entries += [pscustomobject]@{
                Path = $candidate.Path
                SourceLabel = $candidate.SourceLabel
                CategoryLabel = $candidate.CategoryLabel
                Result = "Skipped"
                Message = "Path does not exist"
                BackupPath = ""
            }
            Write-UiLog -LogBox $LogBox -Message "Skipped: $($candidate.Path) (path does not exist)"
            continue
        }

        $backupPath = ""
        try {
            if ($CreateBackup) {
                $backupPath = Backup-Target -Candidate $candidate -SessionBackupPath $sessionBackupPath -Index ($i + 1)
                Write-UiLog -LogBox $LogBox -Message "Backed up: $($candidate.Path) -> $backupPath"
            }

            Stop-TargetProcessIfNeeded -Candidate $candidate
            Remove-TargetWithMode -Candidate $candidate -UseRecycleBin $UseRecycleBin

            $resultMessage = "Permanently deleted"
            if ($UseRecycleBin) {
                $resultMessage = "Moved to Recycle Bin"
            }

            $entries += [pscustomobject]@{
                Path = $candidate.Path
                SourceLabel = $candidate.SourceLabel
                CategoryLabel = $candidate.CategoryLabel
                Result = "Success"
                Message = $resultMessage
                BackupPath = $backupPath
            }
            Write-UiLog -LogBox $LogBox -Message "Processed: $($candidate.Path)"
        } catch {
            $entries += [pscustomobject]@{
                Path = $candidate.Path
                SourceLabel = $candidate.SourceLabel
                CategoryLabel = $candidate.CategoryLabel
                Result = "Skipped"
                Message = $_.Exception.Message
                BackupPath = $backupPath
            }
            Write-UiLog -LogBox $LogBox -Message "Skipped: $($candidate.Path) ($($_.Exception.Message))"
        }
    }

    Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase Two Complete" -ActivityText "Generating execution report..." -Mode Step -Current $SelectedCandidates.Count -Maximum $SelectedCandidates.Count
    $report = Save-DeletionReport -SessionId $sessionId -ReviewSnapshotPath $script:ReviewFilePath -BackupPath $sessionBackupPath -UseRecycleBin $UseRecycleBin -Entries $entries -BlockedCount $BlockedCount

    $successCount = @($entries | Where-Object { $_.Result -eq "Success" }).Count
    $skippedCount = @($entries | Where-Object { $_.Result -ne "Success" }).Count

    Start-Process "notepad.exe" -ArgumentList $report.TextPath | Out-Null
    Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase Two Complete" -ActivityText "Report created: $(Get-ShortDisplayPath -Path $report.TextPath)" -Mode Idle
    [System.Windows.Forms.MessageBox]::Show(
        "Phase two is complete.`r`nSucceeded: $successCount`r`nSkipped / Not Executed: $skippedCount`r`n`r`nReport: $($report.TextPath)",
        "Execution Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Update-ReviewStatus {
    param(
        [System.Windows.Forms.Label]$Label,
        [System.Windows.Forms.Button]$OpenButton,
        [System.Windows.Forms.Button]$Phase2Button
    )

    if (Test-Path -LiteralPath $script:ReviewFilePath) {
        $manifestState = ", metadata missing"
        if (Test-Path -LiteralPath $script:ManifestPath) {
            $manifestState = ", metadata available"
        }

        $Label.Text = "Review file: $script:ReviewFilePath$manifestState"
        $OpenButton.Enabled = $true
        $Phase2Button.Enabled = $true
    } else {
        $Label.Text = "Review file: not generated yet"
        $OpenButton.Enabled = $false
        $Phase2Button.Enabled = $false
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Win Deep Cleaner"
$form.StartPosition = "CenterScreen"
$form.ClientSize = New-Object System.Drawing.Size(980, 680)
$form.MinimumSize = New-Object System.Drawing.Size(980, 680)
$form.BackColor = $script:CanvasColor
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)

$hero = New-Object System.Windows.Forms.Panel
$hero.Location = New-Object System.Drawing.Point(0, 0)
$hero.Size = New-Object System.Drawing.Size(980, 118)
$hero.BackColor = $script:AccentColor
$form.Controls.Add($hero)

$title = New-Object System.Windows.Forms.Label
$title.Text = "Win Deep Cleaner"
$title.Location = New-Object System.Drawing.Point(28, 20)
$title.Size = New-Object System.Drawing.Size(360, 34)
$title.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 20, [System.Drawing.FontStyle]::Bold)
$title.ForeColor = [System.Drawing.Color]::White
$hero.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "An AI-assisted Windows cleanup review tool: scan safe user areas first, then review and execute with backup and reports."
$subtitle.Location = New-Object System.Drawing.Point(30, 60)
$subtitle.Size = New-Object System.Drawing.Size(720, 24)
$subtitle.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
$subtitle.ForeColor = [System.Drawing.Color]::FromArgb(231, 240, 236)
$hero.Controls.Add($subtitle)

$phase1Card = New-Object System.Windows.Forms.Panel
$phase1Card.Location = New-Object System.Drawing.Point(28, 144)
$phase1Card.Size = New-Object System.Drawing.Size(442, 194)
$phase1Card.BackColor = $script:CardColor
$phase1Card.BorderStyle = "FixedSingle"
$form.Controls.Add($phase1Card)

$phase1Title = New-Object System.Windows.Forms.Label
$phase1Title.Text = "Phase One: Scan and Generate Review List"
$phase1Title.Location = New-Object System.Drawing.Point(18, 16)
$phase1Title.Size = New-Object System.Drawing.Size(260, 28)
$phase1Title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 13, [System.Drawing.FontStyle]::Bold)
$phase1Title.ForeColor = $script:AccentColor
$phase1Card.Controls.Add($phase1Title)

$phase1Text = New-Object System.Windows.Forms.Label
$phase1Text.Text = "Scans startup items, app caches, safe user folders, stale temp files, and suspicious executables in user-writable locations. High-risk paths are blocked automatically."
$phase1Text.Location = New-Object System.Drawing.Point(18, 52)
$phase1Text.Size = New-Object System.Drawing.Size(398, 66)
$phase1Text.ForeColor = $script:TextColor
$phase1Card.Controls.Add($phase1Text)

$btnPhase1 = New-Object System.Windows.Forms.Button
$btnPhase1.Text = "Start Phase One"
$btnPhase1.Location = New-Object System.Drawing.Point(18, 132)
$btnPhase1.Size = New-Object System.Drawing.Size(180, 40)
$btnPhase1.BackColor = $script:AccentColor
$btnPhase1.ForeColor = [System.Drawing.Color]::White
$btnPhase1.FlatStyle = "Flat"
$btnPhase1.Tag = "action-button"
$phase1Card.Controls.Add($btnPhase1)

$phase2Card = New-Object System.Windows.Forms.Panel
$phase2Card.Location = New-Object System.Drawing.Point(500, 144)
$phase2Card.Size = New-Object System.Drawing.Size(452, 194)
$phase2Card.BackColor = $script:CardColor
$phase2Card.BorderStyle = "FixedSingle"
$form.Controls.Add($phase2Card)

$phase2Title = New-Object System.Windows.Forms.Label
$phase2Title.Text = "Phase Two: Review and Execute"
$phase2Title.Location = New-Object System.Drawing.Point(18, 16)
$phase2Title.Size = New-Object System.Drawing.Size(250, 28)
$phase2Title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 13, [System.Drawing.FontStyle]::Bold)
$phase2Title.ForeColor = $script:DangerColor
$phase2Card.Controls.Add($phase2Title)

$phase2Text = New-Object System.Windows.Forms.Label
$phase2Text.Text = "After you edit the TXT file, phase two loads the remaining candidates into a checkbox list and offers backup, Recycle Bin, and audit report options."
$phase2Text.Location = New-Object System.Drawing.Point(18, 52)
$phase2Text.Size = New-Object System.Drawing.Size(412, 52)
$phase2Text.ForeColor = $script:TextColor
$phase2Card.Controls.Add($phase2Text)

$btnPhase2 = New-Object System.Windows.Forms.Button
$btnPhase2.Text = "Start Phase Two"
$btnPhase2.Location = New-Object System.Drawing.Point(18, 132)
$btnPhase2.Size = New-Object System.Drawing.Size(180, 40)
$btnPhase2.BackColor = $script:DangerColor
$btnPhase2.ForeColor = [System.Drawing.Color]::White
$btnPhase2.FlatStyle = "Flat"
$btnPhase2.Tag = "action-button"
$phase2Card.Controls.Add($btnPhase2)

$statusCard = New-Object System.Windows.Forms.Panel
$statusCard.Location = New-Object System.Drawing.Point(28, 360)
$statusCard.Size = New-Object System.Drawing.Size(924, 126)
$statusCard.BackColor = $script:CardColor
$statusCard.BorderStyle = "FixedSingle"
$form.Controls.Add($statusCard)

$reviewStatusTitle = New-Object System.Windows.Forms.Label
$reviewStatusTitle.Text = "Review File Status"
$reviewStatusTitle.Location = New-Object System.Drawing.Point(18, 16)
$reviewStatusTitle.Size = New-Object System.Drawing.Size(180, 24)
$reviewStatusTitle.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 11, [System.Drawing.FontStyle]::Bold)
$reviewStatusTitle.ForeColor = $script:TextColor
$statusCard.Controls.Add($reviewStatusTitle)

$reviewPathLabel = New-Object System.Windows.Forms.Label
$reviewPathLabel.Location = New-Object System.Drawing.Point(18, 44)
$reviewPathLabel.Size = New-Object System.Drawing.Size(610, 22)
$reviewPathLabel.ForeColor = $script:MutedColor
$statusCard.Controls.Add($reviewPathLabel)

$btnOpenReview = New-Object System.Windows.Forms.Button
$btnOpenReview.Text = "Open Review File"
$btnOpenReview.Location = New-Object System.Drawing.Point(666, 18)
$btnOpenReview.Size = New-Object System.Drawing.Size(112, 34)
$btnOpenReview.BackColor = [System.Drawing.Color]::FromArgb(228, 221, 210)
$btnOpenReview.FlatStyle = "Flat"
$btnOpenReview.Tag = "action-button"
$statusCard.Controls.Add($btnOpenReview)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh Status"
$btnRefresh.Location = New-Object System.Drawing.Point(792, 18)
$btnRefresh.Size = New-Object System.Drawing.Size(112, 34)
$btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(228, 221, 210)
$btnRefresh.FlatStyle = "Flat"
$btnRefresh.Tag = "action-button"
$statusCard.Controls.Add($btnRefresh)

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Text = "Waiting"
$progressLabel.Location = New-Object System.Drawing.Point(666, 60)
$progressLabel.Size = New-Object System.Drawing.Size(238, 20)
$progressLabel.ForeColor = $script:MutedColor
$statusCard.Controls.Add($progressLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(18, 72)
$progressBar.Size = New-Object System.Drawing.Size(610, 10)
$progressBar.Style = "Continuous"
$statusCard.Controls.Add($progressBar)

$activityLabel = New-Object System.Windows.Forms.Label
$activityLabel.Text = "Current task: waiting to start"
$activityLabel.Location = New-Object System.Drawing.Point(18, 90)
$activityLabel.Size = New-Object System.Drawing.Size(886, 22)
$activityLabel.ForeColor = $script:TextColor
$statusCard.Controls.Add($activityLabel)

$logCard = New-Object System.Windows.Forms.Panel
$logCard.Location = New-Object System.Drawing.Point(28, 510)
$logCard.Size = New-Object System.Drawing.Size(924, 140)
$logCard.BackColor = $script:CardColor
$logCard.BorderStyle = "FixedSingle"
$form.Controls.Add($logCard)

$logTitle = New-Object System.Windows.Forms.Label
$logTitle.Text = "Activity Log"
$logTitle.Location = New-Object System.Drawing.Point(18, 14)
$logTitle.Size = New-Object System.Drawing.Size(160, 24)
$logTitle.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 11, [System.Drawing.FontStyle]::Bold)
$logTitle.ForeColor = $script:TextColor
$logCard.Controls.Add($logTitle)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(18, 44)
$logBox.Size = New-Object System.Drawing.Size(886, 78)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::FromArgb(252, 249, 244)
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logCard.Controls.Add($logBox)

$btnPhase1.Add_Click({
    try {
        $phase1TotalSteps = 7
        Set-UiBusyState -Form $form -IsBusy $true
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Preparing phase one..." -Mode Step -Current 0 -Maximum $phase1TotalSteps
        Write-UiLog -LogBox $logBox -Message "Phase one started: running system cleanup and scanning safe user areas for cleanup and suspicious candidates."

        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 1/${phase1TotalSteps}: running DISM cleanup..." -Mode Marquee
        Wait-ExternalProcess -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup" -LogBox $logBox -StartMessage "Running DISM to clean WinSxS components and update residue. This may take several minutes." -EndMessage "DISM finished." -ContinueOnError $true
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 1/$phase1TotalSteps complete: DISM finished." -Mode Step -Current 1 -Maximum $phase1TotalSteps

        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 2/${phase1TotalSteps}: running Disk Cleanup..." -Mode Marquee
        Wait-ExternalProcess -FilePath "cleanmgr.exe" -ArgumentList "/d c /VERYLOWDISK" -LogBox $logBox -StartMessage "Running cleanmgr for system disk cleanup." -EndMessage "Disk Cleanup finished." -ContinueOnError $true
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 2/$phase1TotalSteps complete: Disk Cleanup finished." -Mode Step -Current 2 -Maximum $phase1TotalSteps

        Write-UiLog -LogBox $logBox -Message "Scanning third-party startup items."
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 3/${phase1TotalSteps}: scanning startup items..." -Mode Marquee
        $startupCandidates = Get-StartupCandidateRecords
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 3/$phase1TotalSteps complete: found $($startupCandidates.Count) startup candidate(s)." -Mode Step -Current 3 -Maximum $phase1TotalSteps

        Write-UiLog -LogBox $logBox -Message "Scanning app caches and developer caches."
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 4/${phase1TotalSteps}: scanning app and developer caches..." -Mode Marquee
        $cacheCandidates = Get-CacheCandidateRecords
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 4/$phase1TotalSteps complete: found $($cacheCandidates.Count) cache candidate(s)." -Mode Step -Current 4 -Maximum $phase1TotalSteps

        Write-UiLog -LogBox $logBox -Message "Scanning safe user folders recursively for inactive installers, stale temp files, and large old files."
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 5/${phase1TotalSteps}: scanning inactive files..." -Mode Marquee
        $inactiveFileCandidates = Get-InactiveFileCandidateRecords -InactiveDays 180
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 5/$phase1TotalSteps complete: found $($inactiveFileCandidates.Count) inactive file candidate(s)." -Mode Step -Current 5 -Maximum $phase1TotalSteps

        Write-UiLog -LogBox $logBox -Message "Scanning safe user locations for suspicious executables and scripts."
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 6/${phase1TotalSteps}: scanning suspicious items..." -Mode Marquee
        $suspiciousCandidates = Get-SuspiciousFileCandidateRecords -RecentDays 30
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 6/$phase1TotalSteps complete: found $($suspiciousCandidates.Count) suspicious candidate(s)." -Mode Step -Current 6 -Maximum $phase1TotalSteps

        $allCandidates = Merge-CandidateRecords -Candidates ($startupCandidates + $cacheCandidates + $inactiveFileCandidates + $suspiciousCandidates)
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Running" -ActivityText "Step 7/${phase1TotalSteps}: creating review file and metadata..." -Mode Marquee
        $artifactInfo = Save-ReviewArtifacts -Candidates $allCandidates

        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Complete" -ActivityText "Review list created with $($artifactInfo.SafeCandidates.Count) candidate(s)." -Mode Idle
        Update-ReviewStatus -Label $reviewPathLabel -OpenButton $btnOpenReview -Phase2Button $btnPhase2

        Write-UiLog -LogBox $logBox -Message "Phase one complete: $($artifactInfo.SafeCandidates.Count) reviewable candidate(s), $($artifactInfo.BlockedCandidates.Count) high-risk path(s) blocked."
        Write-UiLog -LogBox $logBox -Message "Review file: $script:ReviewFilePath"
        Write-UiLog -LogBox $logBox -Message "Metadata file: $script:ManifestPath"
        Start-Process "notepad.exe" -ArgumentList $script:ReviewFilePath | Out-Null

        [System.Windows.Forms.MessageBox]::Show(
            "Phase one is complete.`r`n`r`n1. The review TXT and metadata file were created on the Desktop.`r`n2. Candidates may include cleanup items and suspicious items.`r`n3. Suspicious items are heuristic findings, not a malware verdict.`r`n4. Save the TXT, then return to the main window and start phase two.",
            "Phase One Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase One Failed" -ActivityText "Interrupted: $($_.Exception.Message)" -Mode Idle
        Write-UiLog -LogBox $logBox -Message "Phase one failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Phase one failed: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        Set-UiBusyState -Form $form -IsBusy $false
        Update-ReviewStatus -Label $reviewPathLabel -OpenButton $btnOpenReview -Phase2Button $btnPhase2
    }
})

$btnOpenReview.Add_Click({
    if (Test-Path -LiteralPath $script:ReviewFilePath) {
        Start-Process "notepad.exe" -ArgumentList $script:ReviewFilePath | Out-Null
        Write-UiLog -LogBox $logBox -Message "Review file opened."
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "There is no review file yet. Please run phase one first.",
            "No File",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
})

$btnRefresh.Add_Click({
    Update-ReviewStatus -Label $reviewPathLabel -OpenButton $btnOpenReview -Phase2Button $btnPhase2
    Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Waiting" -ActivityText "Status refreshed. Waiting for the next action." -Mode Idle
    Write-UiLog -LogBox $logBox -Message "Review file status refreshed."
})

$btnPhase2.Add_Click({
    try {
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Preparing Phase Two" -ActivityText "Reading the review file..." -Mode Marquee
        $reviewedPaths = Get-ReviewedPaths
        if ($reviewedPaths.Count -eq 0) {
            Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Waiting" -ActivityText "No executable paths were found. Please save your review results first." -Mode Idle
            [System.Windows.Forms.MessageBox]::Show(
                "No paths were read from the review TXT. Please complete phase one first and keep the lines you want to process.",
                "Nothing To Execute",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Preparing Phase Two" -ActivityText "Resolving reviewed items and applying safety rules..." -Mode Marquee
        $manifestCandidates = Load-CandidateManifest
        $resolution = Resolve-ReviewedCandidates -ReviewedPaths $reviewedPaths -ManifestCandidates $manifestCandidates
        if ($resolution.Allowed.Count -eq 0) {
            Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Waiting" -ActivityText "No executable candidates remain. They may all be blocked or missing." -Mode Idle
            [System.Windows.Forms.MessageBox]::Show(
                "All paths in the review file were blocked by safety rules or no longer exist. Nothing can be executed.",
                "No Executable Items",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        Write-UiLog -LogBox $logBox -Message "Phase two ready: $($resolution.Allowed.Count) executable candidate(s) loaded, $($resolution.Blocked.Count) high-risk path(s) blocked."
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Waiting for Confirmation" -ActivityText "Loaded $($resolution.Allowed.Count) candidate(s). Confirm the items in the checkbox window." -Mode Idle
        $plan = Show-ExecutionPlannerDialog -Candidates $resolution.Allowed -BlockedCandidates $resolution.Blocked
        if (-not $plan) {
            Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase Two Cancelled" -ActivityText "You cancelled the checkbox confirmation. No deletion was performed." -Mode Idle
            Write-UiLog -LogBox $logBox -Message "Phase two cancelled. Nothing was executed."
            return
        }

        Set-UiBusyState -Form $form -IsBusy $true
        Invoke-ReviewedDeletion -SelectedCandidates $plan.SelectedCandidates -CreateBackup $plan.CreateBackup -UseRecycleBin $plan.UseRecycleBin -BlockedCount $resolution.Blocked.Count -LogBox $logBox -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel
    } catch {
        Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Phase Two Failed" -ActivityText "Interrupted: $($_.Exception.Message)" -Mode Idle
        Write-UiLog -LogBox $logBox -Message "Phase two failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Phase two failed: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        Set-UiBusyState -Form $form -IsBusy $false
        Update-ReviewStatus -Label $reviewPathLabel -OpenButton $btnOpenReview -Phase2Button $btnPhase2
    }
})

Update-ReviewStatus -Label $reviewPathLabel -OpenButton $btnOpenReview -Phase2Button $btnPhase2
Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Waiting" -ActivityText "Current task: run phase one to generate the review list." -Mode Idle
Write-UiLog -LogBox $logBox -Message "UI started. Run phase one to generate the review list, edit the TXT file, then start phase two."

[void]$form.ShowDialog()
