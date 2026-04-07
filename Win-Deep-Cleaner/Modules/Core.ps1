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

function Get-ActionDisplayLabel {
    param([string]$ActionType)

    switch ($ActionType) {
        "DisableStartup" { return "Disable Startup" }
        "PartitionAnalysis" { return "Analyze" }
        default { return "Delete" }
    }
}

function Get-FileCompanyName {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return ""
    }

    try {
        return ([System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path)).CompanyName
    } catch {
        return ""
    }
}

function Get-ExecutablePathFromCommand {
    param([string]$Command)

    if ([string]::IsNullOrWhiteSpace($Command)) {
        return $null
    }

    $patterns = @(
        '(?i)"([a-z]:\\[^"]+\.(?:exe|bat|cmd|vbs|ps1|com))"',
        '(?i)([a-z]:\\[^\s,;]+\.(?:exe|bat|cmd|vbs|ps1|com))'
    )

    foreach ($pattern in $patterns) {
        if ($Command -match $pattern) {
            return $matches[1]
        }
    }

    return $null
}

function Get-ShortcutTargetPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($Path)
        if ($shortcut.TargetPath) {
            return $shortcut.TargetPath
        }
    } catch {
        return $null
    }

    return $null
}

function Get-RegistryPropertyTypeName {
    param([Microsoft.Win32.RegistryValueKind]$ValueKind)

    switch ($ValueKind) {
        "String" { return "String" }
        "ExpandString" { return "ExpandString" }
        "Binary" { return "Binary" }
        "DWord" { return "DWord" }
        "MultiString" { return "MultiString" }
        "QWord" { return "QWord" }
        default { return "String" }
    }
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

    if (-not $Item) {
        return $null
    }

    if (-not $Item.PSIsContainer) {
        return [long]$Item.Length
    }

    try {
        $total = 0L
        Get-ChildItem -LiteralPath $Item.FullName -Force -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
            $total += [long]$_.Length
        }
        return $total
    } catch {
        return $null
    }

    return $null
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
