function Test-StartupCandidateProtection {
    param([object]$Candidate)

    $combined = @(
        $Candidate.StartupEntryName
        $Candidate.StartupEntryLocation
        $Candidate.StartupCommand
        $Candidate.StartupTargetPath
        $Candidate.CompanyName
    ) -join " "

    $combined = $combined.ToLowerInvariant()
    $company = ($Candidate.CompanyName | ForEach-Object { $_.ToLowerInvariant() })

    if ($Candidate.StartupTargetPath) {
        $targetGuard = Get-PathGuard -Path $Candidate.StartupTargetPath
        if ($targetGuard.Blocked) {
            return [pscustomobject]@{
                Protected = $true
                Reason = "Startup target is inside a protected Windows or driver path"
            }
        }
    }

    if ($company -like "*microsoft*") {
        return [pscustomobject]@{
            Protected = $true
            Reason = "Likely a Microsoft / Windows startup item"
        }
    }

    $protectedKeywords = @(
        "defender",
        "securityhealth",
        "security",
        "firewall",
        "intel",
        "realtek",
        "nvidia",
        "geforce",
        "radeon",
        "amd",
        "qualcomm",
        "broadcom",
        "mediatek",
        "killer",
        "synaptics",
        "elan",
        "touchpad",
        "bluetooth",
        "wifi",
        "wireless",
        "ethernet",
        "lan",
        "wlan",
        "audio",
        "sound",
        "dolby",
        "dts",
        "hotkey"
    )

    foreach ($keyword in $protectedKeywords) {
        if ($combined -like "*$keyword*") {
            return [pscustomobject]@{
                Protected = $true
                Reason = "Likely required by hardware, audio, network, security, or system integration"
            }
        }
    }

    return [pscustomobject]@{
        Protected = $false
        Reason = ""
    }
}

function New-CandidateRecord {
    param(
        [string]$Path,
        [string]$CategoryLabel,
        [string]$SourceLabel,
        [string]$Reason,
        [string]$ActionType = "Delete",
        [string]$MergeKey = "",
        [bool]$AllowMissingPath = $false
    )

    if (-not $AllowMissingPath -and -not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $item = $null
    if (Test-Path -LiteralPath $Path) {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }

    if (-not $AllowMissingPath -and -not $item) {
        return $null
    }

    $resolvedPath = if ($item) { $item.FullName } else { $Path }
    $guard = Get-PathGuard -Path $resolvedPath
    $lastActivity = Get-LastActivityTime -Item $item

    return [pscustomobject]@{
        Id = [guid]::NewGuid().Guid
        Path = $resolvedPath
        NormalizedPath = Get-NormalizedPath -Path $resolvedPath
        CategoryLabel = $CategoryLabel
        SourceLabel = $SourceLabel
        Reason = $Reason
        ItemKind = Get-ItemKindLabel -Item $item
        SizeBytes = Get-FileSizeIfAvailable -Item $item
        LastActivity = $lastActivity
        LastWriteTime = if ($item) { $item.LastWriteTime } else { $null }
        ExistsNow = (Test-Path -LiteralPath $resolvedPath)
        Blocked = $guard.Blocked
        BlockReason = $guard.Reason
        SelectedByDefault = (-not $guard.Blocked)
        ActionType = $ActionType
        MergeKey = if ($MergeKey) { $MergeKey } else { Get-NormalizedPath -Path $resolvedPath }
        CompanyName = if ($item -and -not $item.PSIsContainer) { Get-FileCompanyName -Path $item.FullName } else { "" }
        StartupEntryKind = ""
        StartupEntryName = ""
        StartupEntryLocation = ""
        StartupRegistryPsPath = ""
        StartupRegistryDisplayPath = ""
        StartupRegistryValueName = ""
        StartupFolderItemPath = ""
        StartupCommand = ""
        StartupTargetPath = ""
    }
}

function New-StartupCandidateRecord {
    param(
        [string]$Path,
        [string]$SourceLabel,
        [string]$Reason,
        [string]$StartupEntryKind,
        [string]$StartupEntryName,
        [string]$StartupEntryLocation,
        [string]$StartupCommand,
        [string]$StartupTargetPath = "",
        [string]$StartupRegistryPsPath = "",
        [string]$StartupRegistryDisplayPath = "",
        [string]$StartupRegistryValueName = "",
        [string]$StartupFolderItemPath = ""
    )

    $mergeKey = "startup::{0}::{1}::{2}" -f $StartupEntryKind, $StartupEntryLocation, $StartupEntryName
    $record = New-CandidateRecord -Path $Path -CategoryLabel "Startup Item" -SourceLabel $SourceLabel -Reason $Reason -ActionType "DisableStartup" -MergeKey $mergeKey -AllowMissingPath $true
    if (-not $record) {
        return $null
    }

    $record.ItemKind = "Startup Entry"
    $record.Blocked = $false
    $record.BlockReason = ""
    $record.ExistsNow = $true
    $record.StartupEntryKind = $StartupEntryKind
    $record.StartupEntryName = $StartupEntryName
    $record.StartupEntryLocation = $StartupEntryLocation
    $record.StartupRegistryPsPath = $StartupRegistryPsPath
    $record.StartupRegistryDisplayPath = $StartupRegistryDisplayPath
    $record.StartupRegistryValueName = $StartupRegistryValueName
    $record.StartupFolderItemPath = $StartupFolderItemPath
    $record.StartupCommand = $StartupCommand
    $record.StartupTargetPath = $StartupTargetPath
    if ($StartupTargetPath -and (Test-Path -LiteralPath $StartupTargetPath)) {
        $record.CompanyName = Get-FileCompanyName -Path $StartupTargetPath
    }

    $protection = Test-StartupCandidateProtection -Candidate $record
    if ($protection.Protected) {
        $record.Blocked = $true
        $record.BlockReason = $protection.Reason
        $record.SelectedByDefault = $false
    } else {
        $record.SelectedByDefault = $true
    }

    return $record
}

function Merge-CandidateRecords {
    param([object[]]$Candidates)

    $map = @{}
    foreach ($candidate in $Candidates) {
        if (-not $candidate) {
            continue
        }

        $key = if ($candidate.MergeKey) { $candidate.MergeKey } else { $candidate.NormalizedPath }
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

function Get-StartupCandidateRecords {
    $records = @()

    $registryDefinitions = @(
        @{ PsPath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"; DisplayPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"; Source = "Startup Registry Entry (Current User)" },
        @{ PsPath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"; DisplayPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce"; Source = "Startup Registry Entry (Current User)" },
        @{ PsPath = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run"; DisplayPath = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run"; Source = "Startup Registry Entry (All Users)" },
        @{ PsPath = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce"; DisplayPath = "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce"; Source = "Startup Registry Entry (All Users)" },
        @{ PsPath = "Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"; DisplayPath = "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"; Source = "Startup Registry Entry (All Users)" },
        @{ PsPath = "Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce"; DisplayPath = "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce"; Source = "Startup Registry Entry (All Users)" }
    )

    foreach ($definition in $registryDefinitions) {
        if (-not (Test-Path -LiteralPath $definition.PsPath)) {
            continue
        }

        try {
            $registryItem = Get-Item -LiteralPath $definition.PsPath -ErrorAction Stop
            foreach ($valueName in $registryItem.GetValueNames()) {
                $command = [string]$registryItem.GetValue($valueName)
                if ([string]::IsNullOrWhiteSpace($command)) {
                    continue
                }

                $targetPath = Get-ExecutablePathFromCommand -Command $command
                if (-not $targetPath) {
                    continue
                }

                $reason = "Startup value '$valueName' from $($definition.DisplayPath)"
                $record = New-StartupCandidateRecord -Path $targetPath -SourceLabel $definition.Source -Reason $reason -StartupEntryKind "RegistryValue" -StartupEntryName $valueName -StartupEntryLocation $definition.DisplayPath -StartupCommand $command -StartupTargetPath $targetPath -StartupRegistryPsPath $definition.PsPath -StartupRegistryDisplayPath $definition.DisplayPath -StartupRegistryValueName $valueName
                if ($record) {
                    $records += $record
                }
            }
        } catch {
            continue
        }
    }

    $startupFolders = @(
        @{ Path = [Environment]::GetFolderPath("Startup"); Source = "Startup Folder Item (Current User)" },
        @{ Path = [Environment]::GetFolderPath("CommonStartup"); Source = "Startup Folder Item (All Users)" }
    )

    foreach ($definition in $startupFolders) {
        if ([string]::IsNullOrWhiteSpace($definition.Path) -or -not (Test-Path -LiteralPath $definition.Path)) {
            continue
        }

        try {
            $items = Get-ChildItem -LiteralPath $definition.Path -Force -File -ErrorAction Stop
        } catch {
            continue
        }

        foreach ($item in $items) {
            $targetPath = if ($item.Extension -ieq ".lnk") { Get-ShortcutTargetPath -Path $item.FullName } else { $item.FullName }
            if (-not $targetPath) {
                $targetPath = $item.FullName
            }

            $reason = "Startup folder item '$($item.Name)' in $($definition.Path)"
            $record = New-StartupCandidateRecord -Path $item.FullName -SourceLabel $definition.Source -Reason $reason -StartupEntryKind "StartupFolderFile" -StartupEntryName $item.Name -StartupEntryLocation $definition.Path -StartupCommand $targetPath -StartupTargetPath $targetPath -StartupFolderItemPath $item.FullName
            if ($record) {
                $records += $record
            }
        }
    }

    return $records
}

function Get-CacheCandidateRecords {
    function Get-DefinitionPaths {
        param($Definition)

        $paths = @()

        if ($Definition.Path) {
            if (Test-Path -LiteralPath $Definition.Path) {
                try {
                    $paths += (Get-Item -LiteralPath $Definition.Path -Force -ErrorAction Stop).FullName
                } catch {
                }
            }
        }

        if ($Definition.Pattern) {
            try {
                $paths += (Get-ChildItem -Path $Definition.Pattern -Force -ErrorAction Stop | Select-Object -ExpandProperty FullName)
            } catch {
            }
        }

        return @($paths | Where-Object { $_ } | Select-Object -Unique)
    }

    $definitions = @(
        @{ Pattern = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\Cache\Cache_Data"; Category = "App Cache"; Source = "Edge Browser Cache"; Reason = "Browser profile cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\Code Cache"; Category = "App Cache"; Source = "Edge Code Cache"; Reason = "Browser code cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\GPUCache"; Category = "App Cache"; Source = "Edge GPU Cache"; Reason = "Browser GPU cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\Service Worker\CacheStorage"; Category = "App Cache"; Source = "Edge Service Worker Cache"; Reason = "Browser service worker cache storage" },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Crashpad"; Category = "App Cache"; Source = "Edge Crashpad Cache"; Reason = "Browser crash reporting cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\Google\Chrome\User Data\*\Cache\Cache_Data"; Category = "App Cache"; Source = "Chrome Browser Cache"; Reason = "Browser profile cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\Google\Chrome\User Data\*\Code Cache"; Category = "App Cache"; Source = "Chrome Code Cache"; Reason = "Browser code cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\Google\Chrome\User Data\*\GPUCache"; Category = "App Cache"; Source = "Chrome GPU Cache"; Reason = "Browser GPU cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\Google\Chrome\User Data\*\Service Worker\CacheStorage"; Category = "App Cache"; Source = "Chrome Service Worker Cache"; Reason = "Browser service worker cache storage" },
        @{ Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Crashpad"; Category = "App Cache"; Source = "Chrome Crashpad Cache"; Reason = "Browser crash reporting cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\*\Cache\Cache_Data"; Category = "App Cache"; Source = "Brave Browser Cache"; Reason = "Browser profile cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\*\Code Cache"; Category = "App Cache"; Source = "Brave Code Cache"; Reason = "Browser code cache directory" },
        @{ Pattern = "$env:APPDATA\Mozilla\Firefox\Profiles\*\cache2"; Category = "App Cache"; Source = "Firefox Cache"; Reason = "Firefox profile cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\Packages\Mozilla.Firefox*\LocalCache\Local\Mozilla\Firefox\Profiles\*\cache2"; Category = "App Cache"; Source = "Firefox Store Cache"; Reason = "Firefox Store edition profile cache directory" },
        @{ Path = "$env:USERPROFILE\Documents\WeChat Files\Applet"; Category = "App Cache"; Source = "WeChat Mini Program Cache"; Reason = "Applet cache inside the WeChat files directory" },
        @{ Path = "$env:USERPROFILE\Documents\WeChat Files\FileStorage\Cache"; Category = "App Cache"; Source = "WeChat File Cache"; Reason = "WeChat FileStorage cache directory" },
        @{ Path = "$env:APPDATA\Tencent\QQ\Temp"; Category = "App Cache"; Source = "QQ Temporary Files"; Reason = "QQ temporary directory" },
        @{ Path = "$env:APPDATA\Slack\Cache"; Category = "App Cache"; Source = "Slack Cache"; Reason = "Slack cache directory" },
        @{ Path = "$env:APPDATA\Slack\Code Cache"; Category = "App Cache"; Source = "Slack Code Cache"; Reason = "Slack code cache directory" },
        @{ Path = "$env:APPDATA\Slack\GPUCache"; Category = "App Cache"; Source = "Slack GPU Cache"; Reason = "Slack GPU cache directory" },
        @{ Path = "$env:APPDATA\Slack\Service Worker\CacheStorage"; Category = "App Cache"; Source = "Slack Service Worker Cache"; Reason = "Slack service worker cache storage" },
        @{ Path = "$env:APPDATA\discord\Cache"; Category = "App Cache"; Source = "Discord Cache"; Reason = "Discord cache directory" },
        @{ Path = "$env:APPDATA\discord\Code Cache"; Category = "App Cache"; Source = "Discord Code Cache"; Reason = "Discord code cache directory" },
        @{ Path = "$env:APPDATA\discord\GPUCache"; Category = "App Cache"; Source = "Discord GPU Cache"; Reason = "Discord GPU cache directory" },
        @{ Path = "$env:APPDATA\discord\Service Worker\CacheStorage"; Category = "App Cache"; Source = "Discord Service Worker Cache"; Reason = "Discord service worker cache storage" },
        @{ Path = "$env:APPDATA\Microsoft\Teams\Cache"; Category = "App Cache"; Source = "Teams Cache"; Reason = "Classic Teams cache directory" },
        @{ Path = "$env:APPDATA\Microsoft\Teams\Code Cache"; Category = "App Cache"; Source = "Teams Code Cache"; Reason = "Classic Teams code cache directory" },
        @{ Path = "$env:APPDATA\Microsoft\Teams\GPUCache"; Category = "App Cache"; Source = "Teams GPU Cache"; Reason = "Classic Teams GPU cache directory" },
        @{ Path = "$env:APPDATA\Microsoft\Teams\Service Worker\CacheStorage"; Category = "App Cache"; Source = "Teams Service Worker Cache"; Reason = "Classic Teams service worker cache storage" },
        @{ Path = "$env:APPDATA\Code\Cache"; Category = "App Cache"; Source = "VS Code Cache"; Reason = "VS Code cache directory" },
        @{ Path = "$env:APPDATA\Code\CachedData"; Category = "App Cache"; Source = "VS Code Cached Data"; Reason = "VS Code cached data directory" },
        @{ Path = "$env:APPDATA\Code\GPUCache"; Category = "App Cache"; Source = "VS Code GPU Cache"; Reason = "VS Code GPU cache directory" },
        @{ Path = "$env:APPDATA\Code\Service Worker\CacheStorage"; Category = "App Cache"; Source = "VS Code Service Worker Cache"; Reason = "VS Code service worker cache storage" },
        @{ Path = "$env:LOCALAPPDATA\CrashDumps"; Category = "App Cache"; Source = "Application Crash Dumps"; Reason = "Per-user crash dump directory" },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportArchive"; Category = "App Cache"; Source = "Windows Error Reporting Archive"; Reason = "Archived error reports for the current user" },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportQueue"; Category = "App Cache"; Source = "Windows Error Reporting Queue"; Reason = "Queued error reports for the current user" },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\WER\Temp"; Category = "App Cache"; Source = "Windows Error Reporting Temp"; Reason = "Temporary error report directory for the current user" },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Category = "App Cache"; Source = "Windows Thumbnail Cache"; Reason = "Explorer thumbnail and icon cache directory" },
        @{ Path = "$env:TEMP"; Category = "App Cache"; Source = "System Temporary Directory"; Reason = "Current user's temp directory" },
        @{ Path = "$env:USERPROFILE\.cache\huggingface"; Category = "Dev Cache"; Source = "HuggingFace Cache"; Reason = "Model download cache directory" },
        @{ Path = "$env:USERPROFILE\.cache\torch"; Category = "Dev Cache"; Source = "PyTorch Cache"; Reason = "Torch default cache directory" },
        @{ Path = "$env:LOCALAPPDATA\pip\Cache"; Category = "Dev Cache"; Source = "pip Cache"; Reason = "pip package cache directory" },
        @{ Path = "$env:USERPROFILE\.conda\pkgs"; Category = "Dev Cache"; Source = "Conda Package Cache"; Reason = "Conda pkgs cache directory" },
        @{ Path = "$env:USERPROFILE\.nuget\packages"; Category = "Dev Cache"; Source = "NuGet Global Packages Cache"; Reason = "NuGet global package cache directory" },
        @{ Path = "$env:LOCALAPPDATA\NuGet\Cache"; Category = "Dev Cache"; Source = "NuGet HTTP Cache"; Reason = "NuGet HTTP cache directory" },
        @{ Path = "$env:APPDATA\Obsidian\Cache"; Category = "App Cache"; Source = "Obsidian Cache"; Reason = "Obsidian cache directory" },
        @{ Path = "$env:LOCALAPPDATA\npm-cache"; Category = "Dev Cache"; Source = "npm Cache"; Reason = "npm default cache directory" },
        @{ Path = "$env:LOCALAPPDATA\Yarn\Cache"; Category = "Dev Cache"; Source = "Yarn Cache"; Reason = "Yarn cache directory" },
        @{ Path = "$env:USERPROFILE\.yarn\berry\cache"; Category = "Dev Cache"; Source = "Yarn Berry Cache"; Reason = "Yarn Berry package cache directory" },
        @{ Path = "$env:LOCALAPPDATA\pnpm-store"; Category = "Dev Cache"; Source = "pnpm Store"; Reason = "pnpm store directory" },
        @{ Path = "$env:USERPROFILE\.pnpm-store"; Category = "Dev Cache"; Source = "pnpm Store"; Reason = "pnpm store directory" },
        @{ Path = "$env:USERPROFILE\.gradle\caches"; Category = "Dev Cache"; Source = "Gradle Cache"; Reason = "Gradle cache directory" },
        @{ Path = "$env:USERPROFILE\.m2\repository"; Category = "Dev Cache"; Source = "Maven Local Repository"; Reason = "Maven local repository cache directory" },
        @{ Path = "$env:USERPROFILE\.cargo\registry\cache"; Category = "Dev Cache"; Source = "Cargo Registry Cache"; Reason = "Cargo registry cache directory" },
        @{ Path = "$env:USERPROFILE\.cargo\git\db"; Category = "Dev Cache"; Source = "Cargo Git Cache"; Reason = "Cargo git dependency cache directory" },
        @{ Path = "$env:USERPROFILE\.android\cache"; Category = "Dev Cache"; Source = "Android SDK Cache"; Reason = "Android SDK cache directory" },
        @{ Path = "$env:LOCALAPPDATA\go-build"; Category = "Dev Cache"; Source = "Go Build Cache"; Reason = "Go build cache directory" },
        @{ Path = "$env:LOCALAPPDATA\ms-playwright"; Category = "Dev Cache"; Source = "Playwright Browser Cache"; Reason = "Playwright browser download cache directory" },
        @{ Pattern = "$env:LOCALAPPDATA\JetBrains\*\caches"; Category = "Dev Cache"; Source = "JetBrains IDE Cache"; Reason = "JetBrains IDE cache directory" }
    )

    $records = @()
    foreach ($definition in $definitions) {
        foreach ($resolvedPath in (Get-DefinitionPaths -Definition $definition)) {
            $record = New-CandidateRecord -Path $resolvedPath -CategoryLabel $definition.Category -SourceLabel $definition.Source -Reason $definition.Reason
            if ($record) {
                $records += $record
            }
        }
    }

    return $records
}

function Get-FilesWithDepthLimit {
    param(
        [string]$Root,
        [int]$MaxDepth
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        return @()
    }

    $results = New-Object System.Collections.Generic.List[object]

    function Add-FilesRecursively {
        param(
            [string]$CurrentPath,
            [int]$CurrentDepth,
            [int]$MaxDepth,
            [System.Collections.Generic.List[object]]$Results
        )

        try {
            foreach ($file in (Get-ChildItem -LiteralPath $CurrentPath -Force -File -ErrorAction Stop)) {
                $Results.Add($file)
            }
        } catch {
        }

        if ($CurrentDepth -ge $MaxDepth) {
            return
        }

        try {
            $directories = Get-ChildItem -LiteralPath $CurrentPath -Force -Directory -ErrorAction Stop
        } catch {
            return
        }

        foreach ($directory in $directories) {
            Add-FilesRecursively -CurrentPath $directory.FullName -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth -Results $Results
        }
    }

    Add-FilesRecursively -CurrentPath $Root -CurrentDepth 0 -MaxDepth $MaxDepth -Results $results
    return @($results)
}

function Get-InactiveFileCandidateRecords {
    param(
        [int]$InactiveDays = 180,
        [int]$MaxDepth = 2
    )

    $cutoff = (Get-Date).AddDays(-$InactiveDays)
    $interestingExtensions = @(".zip", ".rar", ".7z", ".iso", ".msi", ".exe", ".cab", ".dmp", ".log", ".tmp", ".bak")
    $definitions = @(
        @{ Root = (Join-Path $env:USERPROFILE "Downloads"); Category = "Inactive Files"; Source = "Old Installers / Archives in Downloads"; MinSizeBytes = 50MB; Extensions = $interestingExtensions; Depth = $MaxDepth },
        @{ Root = (Join-Path $env:USERPROFILE "Desktop"); Category = "Inactive Files"; Source = "Old Installers / Large Files on Desktop"; MinSizeBytes = 100MB; Extensions = $interestingExtensions; Depth = 1 },
        @{ Root = (Join-Path $env:USERPROFILE "Documents"); Category = "Inactive Files"; Source = "Old Installers / Archives in Documents"; MinSizeBytes = 20MB; Extensions = $interestingExtensions; Depth = 1 }
    )

    $records = @()
    foreach ($definition in $definitions) {
        if (-not (Test-Path -LiteralPath $definition.Root)) {
            continue
        }

        $files = Get-FilesWithDepthLimit -Root $definition.Root -MaxDepth $definition.Depth

        foreach ($file in $files) {
            $lastActivity = Get-LastActivityTime -Item $file
            $extension = $file.Extension.ToLowerInvariant()
            $isInterestingExtension = $definition.Extensions -contains $extension
            $isLargeFile = $file.Length -ge $definition.MinSizeBytes

            if ($lastActivity -le $cutoff -and ($isInterestingExtension -or $isLargeFile)) {
                $relativePath = $file.FullName.Substring($definition.Root.Length).TrimStart('\')
                $reason = "File inactive for six months within $($definition.Root), last activity $(Format-DateValue -Value $lastActivity), relative path $relativePath"
                $record = New-CandidateRecord -Path $file.FullName -CategoryLabel $definition.Category -SourceLabel $definition.Source -Reason $reason
                if ($record) {
                    $records += $record
                }
            }
        }
    }

    return $records
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
        "# Phase two prefers the [Id: ...] tag and falls back to the path at the end of the line."
        "#"
        "# Auto-blocked high-risk or protected items: $($blockedCandidates.Count)"
        "# They will not be processed in phase two."
        ""
    )

    $header | Out-File -FilePath $script:ReviewFilePath -Encoding UTF8

    foreach ($group in ($safeCandidates | Group-Object CategoryLabel)) {
        "[SECTION_$($group.Name)]" | Out-File -FilePath $script:ReviewFilePath -Append -Encoding UTF8
        foreach ($candidate in ($group.Group | Sort-Object @{ Expression = { if ($_.SizeBytes) { [long]$_.SizeBytes } else { 0L } }; Descending = $true }, Path)) {
            $line = "[Id: $($candidate.Id)] [Action: $(Get-ActionDisplayLabel -ActionType $candidate.ActionType)] [Source: $($candidate.SourceLabel)] [Type: $($candidate.ItemKind)] [Last Activity: $(Format-DateValue -Value $candidate.LastActivity)] [Size: $(Format-Bytes -Bytes $candidate.SizeBytes)] $($candidate.Path)"
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
        $items += [pscustomobject]@{
            Id = $candidate.Id
            Path = $candidate.Path
            NormalizedPath = Get-NormalizedPath -Path $candidate.Path
            CategoryLabel = $candidate.CategoryLabel
            SourceLabel = $candidate.SourceLabel
            Reason = $candidate.Reason
            ItemKind = $candidate.ItemKind
            SizeBytes = if ($candidate.SizeBytes) { [long]$candidate.SizeBytes } else { $null }
            LastActivity = if ($candidate.LastActivity) { [datetime]$candidate.LastActivity } else { $null }
            LastWriteTime = if ($candidate.LastWriteTime) { [datetime]$candidate.LastWriteTime } else { $null }
            ExistsNow = Test-Path -LiteralPath $candidate.Path
            Blocked = [bool]$candidate.Blocked
            BlockReason = $candidate.BlockReason
            SelectedByDefault = [bool]$candidate.SelectedByDefault
            ActionType = if ($candidate.ActionType) { $candidate.ActionType } else { "Delete" }
            MergeKey = $candidate.MergeKey
            CompanyName = $candidate.CompanyName
            StartupEntryKind = $candidate.StartupEntryKind
            StartupEntryName = $candidate.StartupEntryName
            StartupEntryLocation = $candidate.StartupEntryLocation
            StartupRegistryPsPath = $candidate.StartupRegistryPsPath
            StartupRegistryDisplayPath = $candidate.StartupRegistryDisplayPath
            StartupRegistryValueName = $candidate.StartupRegistryValueName
            StartupFolderItemPath = $candidate.StartupFolderItemPath
            StartupCommand = $candidate.StartupCommand
            StartupTargetPath = $candidate.StartupTargetPath
        }
    }

    return $items
}

function Get-ReviewIdFromLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $null
    }

    if ($Line -match '\[Id:\s*([^\]]+)\]') {
        return $matches[1].Trim()
    }

    return $null
}

function Get-ReviewedSelections {
    if (-not (Test-Path -LiteralPath $script:ReviewFilePath)) {
        return @()
    }

    try {
        $lines = Get-Content -LiteralPath $script:ReviewFilePath -ErrorAction Stop
    } catch {
        return @()
    }

    $selections = foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed -and $trimmed -notmatch '^#' -and $trimmed -notmatch '^\[SECTION_') {
            $id = Get-ReviewIdFromLine -Line $trimmed
            $path = Get-PathFromLine -Line $trimmed
            if ($id -or $path) {
                [pscustomobject]@{
                    Id = $id
                    Path = $path
                }
            }
        }
    }

    $map = @{}
    foreach ($selection in $selections) {
        $key = if ($selection.Id) {
            "id::$($selection.Id)"
        } elseif ($selection.Path) {
            "path::$(Get-NormalizedPath -Path $selection.Path)"
        } else {
            continue
        }

        if (-not $map.ContainsKey($key)) {
            $map[$key] = $selection
        }
    }

    return $map.Values
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
            ActionType = "Delete"
            MergeKey = (Get-NormalizedPath -Path $Path)
            CompanyName = ""
            StartupEntryKind = ""
            StartupEntryName = ""
            StartupEntryLocation = ""
            StartupRegistryPsPath = ""
            StartupRegistryDisplayPath = ""
            StartupRegistryValueName = ""
            StartupFolderItemPath = ""
            StartupCommand = ""
            StartupTargetPath = ""
        }
    }

    $record = New-CandidateRecord -Path $Path -CategoryLabel "User Added" -SourceLabel "Path Added Manually" -Reason "This path came from the review file but is not present in the phase-one metadata"
    return $record
}

function Test-StartupRegistryValueExists {
    param(
        [string]$RegistryPath,
        [string]$ValueName
    )

    if ([string]::IsNullOrWhiteSpace($RegistryPath) -or -not (Test-Path -LiteralPath $RegistryPath)) {
        return $false
    }

    try {
        $item = Get-Item -LiteralPath $RegistryPath -ErrorAction Stop
        return $item.GetValueNames() -contains $ValueName
    } catch {
        return $false
    }
}

function Refresh-CandidateExecutionState {
    param([object]$Candidate)

    if ($Candidate.ActionType -eq "DisableStartup") {
        $Candidate.Blocked = $false
        $Candidate.BlockReason = ""

        if ($Candidate.StartupEntryKind -eq "RegistryValue") {
            $Candidate.ExistsNow = Test-StartupRegistryValueExists -RegistryPath $Candidate.StartupRegistryPsPath -ValueName $Candidate.StartupRegistryValueName
        } elseif ($Candidate.StartupEntryKind -eq "StartupFolderFile") {
            $Candidate.ExistsNow = Test-Path -LiteralPath $Candidate.StartupFolderItemPath
        } else {
            $Candidate.ExistsNow = $false
        }

        return
    }

    $guard = Get-PathGuard -Path $Candidate.Path
    $Candidate.Blocked = $guard.Blocked
    $Candidate.BlockReason = $guard.Reason
    $Candidate.ExistsNow = Test-Path -LiteralPath $Candidate.Path
}

function Resolve-ReviewedCandidates {
    param(
        [object[]]$ReviewedSelections,
        [object[]]$ManifestCandidates
    )

    $manifestIdMap = @{}
    $manifestPathMap = @{}
    foreach ($candidate in $ManifestCandidates) {
        if ($candidate.Id) {
            $manifestIdMap[$candidate.Id] = $candidate
        }

        if ($candidate.NormalizedPath) {
            $manifestPathMap[$candidate.NormalizedPath] = $candidate
        }
    }

    $resolved = @()
    $blockedFromReview = @()

    foreach ($selection in $ReviewedSelections) {
        $candidate = $null

        if ($selection.Id -and $manifestIdMap.ContainsKey($selection.Id)) {
            $candidate = $manifestIdMap[$selection.Id]
        } elseif ($selection.Path) {
            $normalized = Get-NormalizedPath -Path $selection.Path
            if ($normalized -and $manifestPathMap.ContainsKey($normalized)) {
                $candidate = $manifestPathMap[$normalized]
            } else {
                $candidate = New-AdHocReviewedCandidate -Path $selection.Path
            }
        }

        if (-not $candidate) {
            continue
        }

        Refresh-CandidateExecutionState -Candidate $candidate

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
