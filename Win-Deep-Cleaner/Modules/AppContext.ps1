function Initialize-WinDeepCleanerContext {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName Microsoft.VisualBasic
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $script:DesktopPath = [Environment]::GetFolderPath("Desktop")
    $script:ReviewFilePath = Join-Path $script:DesktopPath "AI-Review-List.txt"
    $script:ManifestPath = Join-Path $script:DesktopPath "AI-Review-Metadata.json"
    $script:ReportRoot = Join-Path $script:DesktopPath "Win-Deep-Cleaner-Reports"
    $script:BackupRoot = Join-Path $script:DesktopPath "Win-Deep-Cleaner-Backups"
    $script:StartupDisabledRoot = Join-Path $script:DesktopPath "Win-Deep-Cleaner-Disabled-Startup"
    $script:PartitionReportRoot = Join-Path $script:DesktopPath "Win-Deep-Cleaner-Partition-Reports"

    $script:AccentColor = [System.Drawing.Color]::FromArgb(34, 87, 74)
    $script:DangerColor = [System.Drawing.Color]::FromArgb(136, 32, 32)
    $script:CanvasColor = [System.Drawing.Color]::FromArgb(245, 240, 231)
    $script:CardColor = [System.Drawing.Color]::FromArgb(255, 251, 245)
    $script:TextColor = [System.Drawing.Color]::FromArgb(36, 36, 36)
    $script:MutedColor = [System.Drawing.Color]::FromArgb(106, 96, 84)
}
