function Get-PartitionFriendlyType {
    param($Partition)

    if (-not $Partition) {
        return "-"
    }

    if ($Partition.Type) {
        return [string]$Partition.Type
    }

    if ($Partition.GptType) {
        return [string]$Partition.GptType
    }

    if ($Partition.MbrType) {
        return [string]$Partition.MbrType
    }

    return "-"
}

function Get-PartitionLabel {
    param($Partition)

    if (-not $Partition) {
        return "-"
    }

    $label = if ($Partition.DriveLetter) { "$($Partition.DriveLetter):" } else { "NoDriveLetter" }
    return "Partition $($Partition.PartitionNumber) ($label)"
}

function Get-PartitionVolumeInfo {
    param($Partition)

    if (-not $Partition) {
        return $null
    }

    try {
        if ($Partition.DriveLetter) {
            return Get-Volume -DriveLetter $Partition.DriveLetter -ErrorAction Stop
        }
    } catch {
        return $null
    }

    return $null
}

function Get-PartitionAnalysisData {
    $requiredCommands = @("Get-Partition", "Get-Disk", "Get-PartitionSupportedSize")
    foreach ($commandName in $requiredCommands) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
            throw "The Storage PowerShell cmdlets are not available on this system."
        }
    }

    $systemDriveLetter = $env:SystemDrive.TrimEnd(':', '\')
    $systemPartition = Get-Partition -DriveLetter $systemDriveLetter -ErrorAction Stop
    $disk = Get-Disk -Number $systemPartition.DiskNumber -ErrorAction Stop
    $partitions = @(Get-Partition -DiskNumber $systemPartition.DiskNumber -ErrorAction Stop | Sort-Object Offset)
    $supportedSize = Get-PartitionSupportedSize -DiskNumber $systemPartition.DiskNumber -PartitionNumber $systemPartition.PartitionNumber -ErrorAction Stop
    $systemIndex = -1
    for ($i = 0; $i -lt $partitions.Count; $i++) {
        if ($partitions[$i].PartitionNumber -eq $systemPartition.PartitionNumber) {
            $systemIndex = $i
            break
        }
    }
    $rightNeighbor = if ($systemIndex -ge 0 -and $systemIndex -lt ($partitions.Count - 1)) { $partitions[$systemIndex + 1] } else { $null }
    $leftNeighbor = if ($systemIndex -gt 0) { $partitions[$systemIndex - 1] } else { $null }
    $currentSize = [long]$systemPartition.Size
    $maxSupportedSize = [long]$supportedSize.SizeMax
    $expandableNowBytes = [Math]::Max($maxSupportedSize - $currentSize, 0)

    $partitionDetails = foreach ($partition in $partitions) {
        $volume = Get-PartitionVolumeInfo -Partition $partition
        [pscustomobject]@{
            PartitionNumber = $partition.PartitionNumber
            DriveLetter = if ($partition.DriveLetter) { "$($partition.DriveLetter):" } else { "" }
            SizeBytes = [long]$partition.Size
            OffsetBytes = [long]$partition.Offset
            Type = Get-PartitionFriendlyType -Partition $partition
            IsSystemPartition = ($partition.PartitionNumber -eq $systemPartition.PartitionNumber)
            FileSystem = if ($volume) { $volume.FileSystem } else { "" }
            VolumeLabel = if ($volume) { $volume.FileSystemLabel } else { "" }
        }
    }

    $guidance = @()
    if ($expandableNowBytes -gt 0) {
        $guidance += "C: can be extended immediately by up to $(Format-Bytes -Bytes $expandableNowBytes) because the system reports adjacent extendable space."
    } else {
        $guidance += "C: cannot be extended immediately with the current partition layout."
    }

    if ($rightNeighbor) {
        $rightLabel = Get-PartitionLabel -Partition $rightNeighbor
        $rightType = Get-PartitionFriendlyType -Partition $rightNeighbor
        $guidance += "The partition to the right of C: is $rightLabel, type $rightType."

        if ($rightNeighbor.DriveLetter) {
            $guidance += "If this is on the same physical disk and you shrink it, Windows still usually needs the resulting unallocated space to be directly adjacent to C: before C: can grow."
        } else {
            $guidance += "Because the partition to the right has no drive letter, it is often a recovery, OEM, or reserved partition. This commonly blocks simple C: extension in Windows Disk Management."
        }
    } else {
        $guidance += "There is no partition to the right of C: on this disk."
    }

    if ($leftNeighbor) {
        $guidance += "The partition to the left of C: is $(Get-PartitionLabel -Partition $leftNeighbor). Space from the left side does not help C: with normal Windows extend operations."
    }

    $guidance += "Only space from another partition on the same physical disk can potentially be reassigned to C:. Space on a different disk cannot be directly added to the system partition."
    $guidance += "This tool only analyzes and reports. It does not shrink, move, or extend partitions automatically."

    return [pscustomobject]@{
        CreatedAt = (Get-Date).ToString("s")
        SystemDrive = "${systemDriveLetter}:"
        DiskNumber = $disk.Number
        DiskFriendlyName = $disk.FriendlyName
        DiskPartitionStyle = $disk.PartitionStyle
        DiskSizeBytes = [long]$disk.Size
        SystemPartitionNumber = $systemPartition.PartitionNumber
        SystemPartitionSizeBytes = $currentSize
        SystemPartitionSupportedMaxBytes = $maxSupportedSize
        ExpandableNowBytes = $expandableNowBytes
        CanExtendNow = ($expandableNowBytes -gt 0)
        RightNeighbor = if ($rightNeighbor) { $rightNeighbor } else { $null }
        LeftNeighbor = if ($leftNeighbor) { $leftNeighbor } else { $null }
        Guidance = $guidance
        Partitions = $partitionDetails
    }
}

function Save-PartitionAnalysisReport {
    param([object]$Analysis)

    Ensure-Directory -Path $script:PartitionReportRoot
    $sessionId = Get-Date -Format "yyyyMMdd-HHmmss"
    $basePath = Join-Path $script:PartitionReportRoot ("Partition-Analysis-{0}" -f $sessionId)
    $txtPath = "$basePath.txt"
    $jsonPath = "$basePath.json"

    $lines = @(
        "Win Deep Cleaner Partition Analysis"
        "CreatedAt: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "SystemDrive: $($Analysis.SystemDrive)"
        "DiskNumber: $($Analysis.DiskNumber)"
        "DiskFriendlyName: $($Analysis.DiskFriendlyName)"
        "DiskPartitionStyle: $($Analysis.DiskPartitionStyle)"
        "DiskSize: $(Format-Bytes -Bytes $Analysis.DiskSizeBytes)"
        "CurrentCSize: $(Format-Bytes -Bytes $Analysis.SystemPartitionSizeBytes)"
        "MaxSupportedCSize: $(Format-Bytes -Bytes $Analysis.SystemPartitionSupportedMaxBytes)"
        "CanExtendNow: $($Analysis.CanExtendNow)"
        "ExpandableNow: $(Format-Bytes -Bytes $Analysis.ExpandableNowBytes)"
        ""
        "Guidance:"
    )

    foreach ($item in $Analysis.Guidance) {
        $lines += "- $item"
    }

    $lines += ""
    $lines += "Partitions on the same disk:"
    foreach ($partition in $Analysis.Partitions) {
        $lines += "[Partition $($partition.PartitionNumber)] Drive=$($partition.DriveLetter) Size=$(Format-Bytes -Bytes $partition.SizeBytes) Type=$($partition.Type) FileSystem=$($partition.FileSystem) Label=$($partition.VolumeLabel)"
    }

    $lines | Out-File -FilePath $txtPath -Encoding UTF8
    $Analysis | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8

    return [pscustomobject]@{
        TextPath = $txtPath
        JsonPath = $jsonPath
    }
}

function Invoke-PartitionAnalysisWorkflow {
    param(
        [System.Windows.Forms.Form]$Form,
        [System.Windows.Forms.TextBox]$LogBox,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.Label]$ActivityLabel
    )

    try {
        Set-UiBusyState -Form $Form -IsBusy $true
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Partition Analysis Running" -ActivityText "Inspecting disk and partition layout..." -Mode Marquee
        Write-UiLog -LogBox $LogBox -Message "Partition analysis started: checking whether C: can be extended safely."

        $analysis = Get-PartitionAnalysisData
        $report = Save-PartitionAnalysisReport -Analysis $analysis

        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Partition Analysis Complete" -ActivityText "Report created: $(Get-ShortDisplayPath -Path $report.TextPath)" -Mode Idle
        Write-UiLog -LogBox $LogBox -Message "Partition analysis complete. Report: $($report.TextPath)"
        Start-Process "notepad.exe" -ArgumentList $report.TextPath | Out-Null

        $summary = if ($analysis.CanExtendNow) {
            "C: can be extended immediately by up to $(Format-Bytes -Bytes $analysis.ExpandableNowBytes)."
        } else {
            "C: cannot be extended immediately. Review the generated report for the blocking partition layout."
        }

        [System.Windows.Forms.MessageBox]::Show(
            "$summary`r`n`r`nReport: $($report.TextPath)",
            "Partition Analysis Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Partition Analysis Failed" -ActivityText "Interrupted: $($_.Exception.Message)" -Mode Idle
        Write-UiLog -LogBox $LogBox -Message "Partition analysis failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Partition analysis failed: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        Set-UiBusyState -Form $Form -IsBusy $false
    }
}
