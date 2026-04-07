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
    $summary.Text = "These are the items you kept in the review file. You can uncheck any item here. File candidates will be deleted; startup candidates will be disabled."
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
    [void]$listView.Columns.Add("Category", 105)
    [void]$listView.Columns.Add("Action", 95)
    [void]$listView.Columns.Add("Type", 85)
    [void]$listView.Columns.Add("Last Activity", 95)
    [void]$listView.Columns.Add("Size", 90)
    [void]$listView.Columns.Add("Status", 90)
    [void]$listView.Columns.Add("Path", 245)
    $dialog.Controls.Add($listView)

    foreach ($candidate in $Candidates) {
        $item = New-Object System.Windows.Forms.ListViewItem($candidate.SourceLabel)
        [void]$item.SubItems.Add($candidate.CategoryLabel)
        [void]$item.SubItems.Add((Get-ActionDisplayLabel -ActionType $candidate.ActionType))
        [void]$item.SubItems.Add($candidate.ItemKind)
        [void]$item.SubItems.Add((Format-DateValue -Value $candidate.LastActivity))
        [void]$item.SubItems.Add((Format-Bytes -Bytes $candidate.SizeBytes))
        [void]$item.SubItems.Add((if ($candidate.ExistsNow) { "Ready" } else { "Missing" }))
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
        $guardLabel.Text = "$($BlockedCandidates.Count) high-risk or protected item(s) were blocked and will not be executed. Example: $blockedPreview"
    } else {
        $guardLabel.Text = "No high-risk or protected items were detected in this phase-two run."
    }
    $dialog.Controls.Add($guardLabel)

    $backupCheck = New-Object System.Windows.Forms.CheckBox
    $backupCheck.Text = "Back up items before delete / disable"
    $backupCheck.Location = New-Object System.Drawing.Point(24, 538)
    $backupCheck.Size = New-Object System.Drawing.Size(250, 24)
    $backupCheck.Checked = $true
    $dialog.Controls.Add($backupCheck)

    $recycleCheck = New-Object System.Windows.Forms.CheckBox
    $recycleCheck.Text = "Prefer moving deletions to the Recycle Bin"
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
            if ($item.SubItems[6].Text -eq "Ready") {
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
                "No items are selected yet.",
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

function Start-WinDeepCleanerApp {
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
    $subtitle.Text = "Generate reviewable cleanup candidates, disable non-essential startup items, and analyze whether C: can grow safely."
    $subtitle.Location = New-Object System.Drawing.Point(30, 60)
    $subtitle.Size = New-Object System.Drawing.Size(820, 24)
    $subtitle.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(231, 240, 236)
    $hero.Controls.Add($subtitle)

    $phase1Card = New-Object System.Windows.Forms.Panel
    $phase1Card.Location = New-Object System.Drawing.Point(28, 144)
    $phase1Card.Size = New-Object System.Drawing.Size(292, 194)
    $phase1Card.BackColor = $script:CardColor
    $phase1Card.BorderStyle = "FixedSingle"
    $form.Controls.Add($phase1Card)

    $phase1Title = New-Object System.Windows.Forms.Label
    $phase1Title.Text = "Phase One"
    $phase1Title.Location = New-Object System.Drawing.Point(18, 16)
    $phase1Title.Size = New-Object System.Drawing.Size(150, 28)
    $phase1Title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 13, [System.Drawing.FontStyle]::Bold)
    $phase1Title.ForeColor = $script:AccentColor
    $phase1Card.Controls.Add($phase1Title)

    $phase1Text = New-Object System.Windows.Forms.Label
    $phase1Text.Text = "Scan startup items, browser and app caches, developer caches, crash dumps, and deeper inactive files under Downloads / Desktop / Documents."
    $phase1Text.Location = New-Object System.Drawing.Point(18, 52)
    $phase1Text.Size = New-Object System.Drawing.Size(246, 66)
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
    $phase2Card.Location = New-Object System.Drawing.Point(344, 144)
    $phase2Card.Size = New-Object System.Drawing.Size(292, 194)
    $phase2Card.BackColor = $script:CardColor
    $phase2Card.BorderStyle = "FixedSingle"
    $form.Controls.Add($phase2Card)

    $phase2Title = New-Object System.Windows.Forms.Label
    $phase2Title.Text = "Phase Two"
    $phase2Title.Location = New-Object System.Drawing.Point(18, 16)
    $phase2Title.Size = New-Object System.Drawing.Size(150, 28)
    $phase2Title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 13, [System.Drawing.FontStyle]::Bold)
    $phase2Title.ForeColor = $script:DangerColor
    $phase2Card.Controls.Add($phase2Title)

    $phase2Text = New-Object System.Windows.Forms.Label
    $phase2Text.Text = "Review and confirm each item. File candidates are deleted; startup candidates are disabled, never auto-deleted."
    $phase2Text.Location = New-Object System.Drawing.Point(18, 52)
    $phase2Text.Size = New-Object System.Drawing.Size(246, 66)
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

    $partitionCard = New-Object System.Windows.Forms.Panel
    $partitionCard.Location = New-Object System.Drawing.Point(660, 144)
    $partitionCard.Size = New-Object System.Drawing.Size(292, 194)
    $partitionCard.BackColor = $script:CardColor
    $partitionCard.BorderStyle = "FixedSingle"
    $form.Controls.Add($partitionCard)

    $partitionTitle = New-Object System.Windows.Forms.Label
    $partitionTitle.Text = "C: Partition Analysis"
    $partitionTitle.Location = New-Object System.Drawing.Point(18, 16)
    $partitionTitle.Size = New-Object System.Drawing.Size(200, 28)
    $partitionTitle.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 13, [System.Drawing.FontStyle]::Bold)
    $partitionTitle.ForeColor = $script:AccentColor
    $partitionCard.Controls.Add($partitionTitle)

    $partitionText = New-Object System.Windows.Forms.Label
    $partitionText.Text = "Inspect whether C: can grow, which partition blocks it, and whether the layout is safe enough for a manual resize plan."
    $partitionText.Location = New-Object System.Drawing.Point(18, 52)
    $partitionText.Size = New-Object System.Drawing.Size(246, 66)
    $partitionText.ForeColor = $script:TextColor
    $partitionCard.Controls.Add($partitionText)

    $btnPartition = New-Object System.Windows.Forms.Button
    $btnPartition.Text = "Analyze C: Layout"
    $btnPartition.Location = New-Object System.Drawing.Point(18, 132)
    $btnPartition.Size = New-Object System.Drawing.Size(180, 40)
    $btnPartition.BackColor = [System.Drawing.Color]::FromArgb(68, 110, 97)
    $btnPartition.ForeColor = [System.Drawing.Color]::White
    $btnPartition.FlatStyle = "Flat"
    $btnPartition.Tag = "action-button"
    $partitionCard.Controls.Add($btnPartition)

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
        Invoke-PhaseOneWorkflow -Form $form -LogBox $logBox -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -ReviewStatusLabel $reviewPathLabel -OpenReviewButton $btnOpenReview -Phase2Button $btnPhase2
    })

    $btnPhase2.Add_Click({
        Invoke-PhaseTwoWorkflow -Form $form -LogBox $logBox -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -ReviewStatusLabel $reviewPathLabel -OpenReviewButton $btnOpenReview -Phase2Button $btnPhase2
    })

    $btnPartition.Add_Click({
        Invoke-PartitionAnalysisWorkflow -Form $form -LogBox $logBox -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel
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

    Update-ReviewStatus -Label $reviewPathLabel -OpenButton $btnOpenReview -Phase2Button $btnPhase2
    Set-ProgressState -ProgressBar $progressBar -StatusLabel $progressLabel -ActivityLabel $activityLabel -StatusText "Waiting" -ActivityText "Current task: run phase one, phase two, or partition analysis." -Mode Idle
    Write-UiLog -LogBox $logBox -Message "UI started. Run phase one to generate the review list, phase two to execute confirmed actions, or analyze the C: partition layout."

    [void]$form.ShowDialog()
}
