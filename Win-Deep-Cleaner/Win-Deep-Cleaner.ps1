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

$moduleRoot = Join-Path $PSScriptRoot "Modules"
$moduleFiles = @(
    "AppContext.ps1",
    "Core.ps1",
    "Candidates.ps1",
    "Actions.ps1",
    "Partition.ps1",
    "Workflow.ps1",
    "UI.ps1"
)

foreach ($moduleFile in $moduleFiles) {
    $modulePath = Join-Path $moduleRoot $moduleFile
    if (-not (Test-Path -LiteralPath $modulePath)) {
        throw "Required module file is missing: $modulePath"
    }

    . $modulePath
}

Initialize-WinDeepCleanerContext
Start-WinDeepCleanerApp
