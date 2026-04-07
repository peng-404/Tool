function Stop-TargetProcessIfNeeded {
    param([object]$Candidate)

    if ($Candidate.ActionType -ne "Delete") {
        return
    }

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

function Get-UniqueDestinationPath {
    param(
        [string]$Directory,
        [string]$Name
    )

    $candidatePath = Join-Path $Directory $Name
    if (-not (Test-Path -LiteralPath $candidatePath)) {
        return $candidatePath
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $extension = [System.IO.Path]::GetExtension($Name)

    for ($index = 1; $index -lt 5000; $index++) {
        $alternative = Join-Path $Directory ("{0}-{1}{2}" -f $baseName, $index, $extension)
        if (-not (Test-Path -LiteralPath $alternative)) {
            return $alternative
        }
    }

    throw "Unable to create a unique destination path for $Name"
}

function Backup-Candidate {
    param(
        [object]$Candidate,
        [string]$SessionBackupPath,
        [int]$Index
    )

    $slot = Join-Path $SessionBackupPath ("item-{0:D3}" -f $Index)
    Ensure-Directory -Path $slot

    $metadata = [pscustomobject]@{
        Id = $Candidate.Id
        ActionType = $Candidate.ActionType
        SourceLabel = $Candidate.SourceLabel
        CategoryLabel = $Candidate.CategoryLabel
        Path = $Candidate.Path
        StartupEntryKind = $Candidate.StartupEntryKind
        StartupEntryName = $Candidate.StartupEntryName
        StartupEntryLocation = $Candidate.StartupEntryLocation
        StartupRegistryDisplayPath = $Candidate.StartupRegistryDisplayPath
        StartupRegistryValueName = $Candidate.StartupRegistryValueName
        StartupFolderItemPath = $Candidate.StartupFolderItemPath
        StartupCommand = $Candidate.StartupCommand
        StartupTargetPath = $Candidate.StartupTargetPath
    }
    $metadata | ConvertTo-Json -Depth 5 | Out-File -FilePath (Join-Path $slot "candidate.json") -Encoding UTF8

    if ($Candidate.ActionType -eq "DisableStartup") {
        if ($Candidate.StartupEntryKind -eq "StartupFolderFile" -and (Test-Path -LiteralPath $Candidate.StartupFolderItemPath)) {
            Copy-Item -LiteralPath $Candidate.StartupFolderItemPath -Destination $slot -Force -ErrorAction Stop
        }
    } else {
        Copy-Item -LiteralPath $Candidate.Path -Destination $slot -Recurse -Force -ErrorAction Stop
    }

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

function Disable-StartupCandidate {
    param(
        [object]$Candidate,
        [string]$SessionId
    )

    if ($Candidate.StartupEntryKind -eq "RegistryValue") {
        if (-not (Test-StartupRegistryValueExists -RegistryPath $Candidate.StartupRegistryPsPath -ValueName $Candidate.StartupRegistryValueName)) {
            throw "Startup registry value no longer exists"
        }

        $registryItem = Get-Item -LiteralPath $Candidate.StartupRegistryPsPath -ErrorAction Stop
        $valueData = $registryItem.GetValue($Candidate.StartupRegistryValueName, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        $valueKind = $registryItem.GetValueKind($Candidate.StartupRegistryValueName)
        $disabledPath = Join-Path $Candidate.StartupRegistryPsPath "WinDeepCleanerDisabled"
        if (-not (Test-Path -LiteralPath $disabledPath)) {
            [void](New-Item -Path $disabledPath -Force)
        }

        $disabledName = "{0}__{1}" -f $Candidate.StartupRegistryValueName, $SessionId
        New-ItemProperty -Path $disabledPath -Name $disabledName -Value $valueData -PropertyType (Get-RegistryPropertyTypeName -ValueKind $valueKind) -Force | Out-Null
        Remove-ItemProperty -LiteralPath $Candidate.StartupRegistryPsPath -Name $Candidate.StartupRegistryValueName -Force -ErrorAction Stop
        return "Startup registry value disabled: $($Candidate.StartupRegistryDisplayPath) [$($Candidate.StartupRegistryValueName)]"
    }

    if ($Candidate.StartupEntryKind -eq "StartupFolderFile") {
        if (-not (Test-Path -LiteralPath $Candidate.StartupFolderItemPath)) {
            throw "Startup folder item no longer exists"
        }

        Ensure-Directory -Path $script:StartupDisabledRoot
        $sessionPath = Join-Path $script:StartupDisabledRoot ("Disabled-{0}" -f $SessionId)
        Ensure-Directory -Path $sessionPath
        $destination = Get-UniqueDestinationPath -Directory $sessionPath -Name ([System.IO.Path]::GetFileName($Candidate.StartupFolderItemPath))
        Move-Item -LiteralPath $Candidate.StartupFolderItemPath -Destination $destination -Force -ErrorAction Stop
        return "Startup folder item moved out of Startup: $destination"
    }

    throw "Startup entry metadata is incomplete"
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

    $lines = @(
        "Win Deep Cleaner Execution Report"
        "Session: $SessionId"
        "CreatedAt: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
        "ReviewFile: $ReviewSnapshotPath"
        "BackupPath: $(if ($BackupPath) { $BackupPath } else { '-' })"
        "DeleteMode: $(if ($UseRecycleBin) { 'RecycleBin' } else { 'PermanentDelete' })"
        "BlockedByGuard: $BlockedCount"
        "SuccessCount: $successCount"
        "SkippedCount: $skippedCount"
        ""
        "Details:"
    )

    foreach ($entry in $Entries) {
        $lines += "[{0}] [{1}] [{2}] [{3}] {4} | {5}" -f $entry.Result, $entry.Action, $entry.SourceLabel, $entry.CategoryLabel, $entry.Path, $entry.Message
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
    Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase Two Running" -ActivityText "Initializing execution plan..." -Mode Step -Current 0 -Maximum $SelectedCandidates.Count
    Write-UiLog -LogBox $LogBox -Message "Phase two started: $($SelectedCandidates.Count) selected item(s). Delete mode: $(if ($UseRecycleBin) { 'Recycle Bin' } else { 'Permanent Delete' }). Startup items will be disabled."

    for ($i = 0; $i -lt $SelectedCandidates.Count; $i++) {
        $candidate = $SelectedCandidates[$i]
        Refresh-CandidateExecutionState -Candidate $candidate

        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Processing: $($i + 1) / $($SelectedCandidates.Count)" -ActivityText "Working on: $(Get-ShortDisplayPath -Path $candidate.Path)" -Mode Step -Current ($i + 1) -Maximum $SelectedCandidates.Count

        if ($candidate.Blocked) {
            $entries += [pscustomobject]@{
                Action = Get-ActionDisplayLabel -ActionType $candidate.ActionType
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
                Action = Get-ActionDisplayLabel -ActionType $candidate.ActionType
                Path = $candidate.Path
                SourceLabel = $candidate.SourceLabel
                CategoryLabel = $candidate.CategoryLabel
                Result = "Skipped"
                Message = if ($candidate.ActionType -eq "DisableStartup") { "Startup entry no longer exists" } else { "Path does not exist" }
                BackupPath = ""
            }
            Write-UiLog -LogBox $LogBox -Message "Skipped: $($candidate.Path) ($(if ($candidate.ActionType -eq 'DisableStartup') { 'startup entry no longer exists' } else { 'path does not exist' }))"
            continue
        }

        $backupPath = ""
        try {
            if ($CreateBackup) {
                $backupPath = Backup-Candidate -Candidate $candidate -SessionBackupPath $sessionBackupPath -Index ($i + 1)
                Write-UiLog -LogBox $LogBox -Message "Backed up: $($candidate.Path) -> $backupPath"
            }

            if ($candidate.ActionType -eq "DisableStartup") {
                $message = Disable-StartupCandidate -Candidate $candidate -SessionId $sessionId
            } else {
                Stop-TargetProcessIfNeeded -Candidate $candidate
                Remove-TargetWithMode -Candidate $candidate -UseRecycleBin $UseRecycleBin
                $message = if ($UseRecycleBin) { "Moved to Recycle Bin" } else { "Permanently deleted" }
            }

            $entries += [pscustomobject]@{
                Action = Get-ActionDisplayLabel -ActionType $candidate.ActionType
                Path = $candidate.Path
                SourceLabel = $candidate.SourceLabel
                CategoryLabel = $candidate.CategoryLabel
                Result = "Success"
                Message = $message
                BackupPath = $backupPath
            }
            Write-UiLog -LogBox $LogBox -Message "Processed: $($candidate.Path)"
        } catch {
            $entries += [pscustomobject]@{
                Action = Get-ActionDisplayLabel -ActionType $candidate.ActionType
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
