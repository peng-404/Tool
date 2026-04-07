function Update-ReviewStatus {
    param(
        [System.Windows.Forms.Label]$Label,
        [System.Windows.Forms.Button]$OpenButton,
        [System.Windows.Forms.Button]$Phase2Button
    )

    if (Test-Path -LiteralPath $script:ReviewFilePath) {
        $manifestState = if (Test-Path -LiteralPath $script:ManifestPath) { ", metadata available" } else { ", metadata missing" }
        $Label.Text = "Review file: $script:ReviewFilePath$manifestState"
        $OpenButton.Enabled = $true
        $Phase2Button.Enabled = $true
    } else {
        $Label.Text = "Review file: not generated yet"
        $OpenButton.Enabled = $false
        $Phase2Button.Enabled = $false
    }
}

function Invoke-PhaseOneWorkflow {
    param(
        [System.Windows.Forms.Form]$Form,
        [System.Windows.Forms.TextBox]$LogBox,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.Label]$ActivityLabel,
        [System.Windows.Forms.Label]$ReviewStatusLabel,
        [System.Windows.Forms.Button]$OpenReviewButton,
        [System.Windows.Forms.Button]$Phase2Button
    )

    try {
        $phase1TotalSteps = 6
        Set-UiBusyState -Form $Form -IsBusy $true
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Preparing phase one..." -Mode Step -Current 0 -Maximum $phase1TotalSteps
        Write-UiLog -LogBox $LogBox -Message "Phase one started: running system cleanup and collecting candidate items."

        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 1/${phase1TotalSteps}: running DISM cleanup..." -Mode Marquee
        Wait-ExternalProcess -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup" -LogBox $LogBox -StartMessage "Running DISM to clean WinSxS components and update residue. This may take several minutes." -EndMessage "DISM finished." -ContinueOnError $true
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 1/$phase1TotalSteps complete: DISM finished." -Mode Step -Current 1 -Maximum $phase1TotalSteps

        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 2/${phase1TotalSteps}: running Disk Cleanup..." -Mode Marquee
        Wait-ExternalProcess -FilePath "cleanmgr.exe" -ArgumentList "/d c /VERYLOWDISK" -LogBox $LogBox -StartMessage "Running cleanmgr for system disk cleanup." -EndMessage "Disk Cleanup finished." -ContinueOnError $true
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 2/$phase1TotalSteps complete: Disk Cleanup finished." -Mode Step -Current 2 -Maximum $phase1TotalSteps

        Write-UiLog -LogBox $LogBox -Message "Scanning startup registry entries and startup folder items."
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 3/${phase1TotalSteps}: scanning startup items..." -Mode Marquee
        $startupCandidates = Get-StartupCandidateRecords
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 3/$phase1TotalSteps complete: found $($startupCandidates.Count) startup candidate(s)." -Mode Step -Current 3 -Maximum $phase1TotalSteps

        Write-UiLog -LogBox $LogBox -Message "Scanning app caches and developer caches."
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 4/${phase1TotalSteps}: scanning app and developer caches..." -Mode Marquee
        $cacheCandidates = Get-CacheCandidateRecords
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 4/$phase1TotalSteps complete: found $($cacheCandidates.Count) cache candidate(s)." -Mode Step -Current 4 -Maximum $phase1TotalSteps

        Write-UiLog -LogBox $LogBox -Message "Scanning inactive installers and large files in Downloads, Desktop, and Documents."
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 5/${phase1TotalSteps}: scanning inactive files..." -Mode Marquee
        $inactiveFileCandidates = Get-InactiveFileCandidateRecords -InactiveDays 180
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 5/$phase1TotalSteps complete: found $($inactiveFileCandidates.Count) inactive file candidate(s)." -Mode Step -Current 5 -Maximum $phase1TotalSteps

        $allCandidates = Merge-CandidateRecords -Candidates ($startupCandidates + $cacheCandidates + $inactiveFileCandidates)
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Running" -ActivityText "Step 6/${phase1TotalSteps}: creating review file and metadata..." -Mode Marquee
        $artifactInfo = Save-ReviewArtifacts -Candidates $allCandidates

        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Complete" -ActivityText "Review list created with $($artifactInfo.SafeCandidates.Count) candidate(s)." -Mode Idle
        Update-ReviewStatus -Label $ReviewStatusLabel -OpenButton $OpenReviewButton -Phase2Button $Phase2Button

        Write-UiLog -LogBox $LogBox -Message "Phase one complete: $($artifactInfo.SafeCandidates.Count) reviewable candidate(s), $($artifactInfo.BlockedCandidates.Count) high-risk or protected item(s) blocked."
        Write-UiLog -LogBox $LogBox -Message "Review file: $script:ReviewFilePath"
        Write-UiLog -LogBox $LogBox -Message "Metadata file: $script:ManifestPath"
        Start-Process "notepad.exe" -ArgumentList $script:ReviewFilePath | Out-Null

        [System.Windows.Forms.MessageBox]::Show(
            "Phase one is complete.`r`n`r`n1. The review TXT and metadata file were created on the Desktop.`r`n2. Every candidate includes a source label and action type.`r`n3. You can send the TXT to an AI model or review it manually.`r`n4. Save the TXT, then return to the main window and start phase two.",
            "Phase One Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase One Failed" -ActivityText "Interrupted: $($_.Exception.Message)" -Mode Idle
        Write-UiLog -LogBox $LogBox -Message "Phase one failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Phase one failed: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        Set-UiBusyState -Form $Form -IsBusy $false
        Update-ReviewStatus -Label $ReviewStatusLabel -OpenButton $OpenReviewButton -Phase2Button $Phase2Button
    }
}

function Invoke-PhaseTwoWorkflow {
    param(
        [System.Windows.Forms.Form]$Form,
        [System.Windows.Forms.TextBox]$LogBox,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.Label]$ActivityLabel,
        [System.Windows.Forms.Label]$ReviewStatusLabel,
        [System.Windows.Forms.Button]$OpenReviewButton,
        [System.Windows.Forms.Button]$Phase2Button
    )

    try {
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Preparing Phase Two" -ActivityText "Reading the review file..." -Mode Marquee
        $reviewedSelections = Get-ReviewedSelections
        if ($reviewedSelections.Count -eq 0) {
            Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Waiting" -ActivityText "No review selections were found. Please save your review results first." -Mode Idle
            [System.Windows.Forms.MessageBox]::Show(
                "No selections were read from the review TXT. Please complete phase one first and keep the lines you want to process.",
                "Nothing To Execute",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Preparing Phase Two" -ActivityText "Resolving reviewed items and applying safety rules..." -Mode Marquee
        $manifestCandidates = Load-CandidateManifest
        $resolution = Resolve-ReviewedCandidates -ReviewedSelections $reviewedSelections -ManifestCandidates $manifestCandidates
        if ($resolution.Allowed.Count -eq 0) {
            Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Waiting" -ActivityText "No executable candidates remain. They may all be blocked, protected, or missing." -Mode Idle
            [System.Windows.Forms.MessageBox]::Show(
                "All review items were blocked by safety rules, marked as protected startup items, or no longer exist. Nothing can be executed.",
                "No Executable Items",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        Write-UiLog -LogBox $LogBox -Message "Phase two ready: $($resolution.Allowed.Count) executable candidate(s) loaded, $($resolution.Blocked.Count) high-risk or protected item(s) blocked."
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Waiting for Confirmation" -ActivityText "Loaded $($resolution.Allowed.Count) candidate(s). Confirm the items in the checkbox window." -Mode Idle
        $plan = Show-ExecutionPlannerDialog -Candidates $resolution.Allowed -BlockedCandidates $resolution.Blocked
        if (-not $plan) {
            Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase Two Cancelled" -ActivityText "You cancelled the checkbox confirmation. No action was performed." -Mode Idle
            Write-UiLog -LogBox $LogBox -Message "Phase two cancelled. Nothing was executed."
            return
        }

        Set-UiBusyState -Form $Form -IsBusy $true
        Invoke-ReviewedDeletion -SelectedCandidates $plan.SelectedCandidates -CreateBackup $plan.CreateBackup -UseRecycleBin $plan.UseRecycleBin -BlockedCount $resolution.Blocked.Count -LogBox $LogBox -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel
    } catch {
        Set-ProgressState -ProgressBar $ProgressBar -StatusLabel $StatusLabel -ActivityLabel $ActivityLabel -StatusText "Phase Two Failed" -ActivityText "Interrupted: $($_.Exception.Message)" -Mode Idle
        Write-UiLog -LogBox $LogBox -Message "Phase two failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Phase two failed: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        Set-UiBusyState -Form $Form -IsBusy $false
        Update-ReviewStatus -Label $ReviewStatusLabel -OpenButton $OpenReviewButton -Phase2Button $Phase2Button
    }
}
