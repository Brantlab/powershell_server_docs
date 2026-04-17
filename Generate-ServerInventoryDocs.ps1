param(
    [Parameter(Mandatory = $true)]
    [string]$SearchBase,

    [string]$OutputPath = ".\docs\servers",

    [string]$DomainController,

    [string]$ServerFilter = "*",

    [int]$MaxServers = 0,

    [switch]$IncludeDisabled,

    [switch]$SkipIndex,

    [System.Management.Automation.PSCredential]$Credential
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-DirectoryIfNeeded {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -Path $Path -ItemType Directory -Force
    }
}

function ConvertTo-GiB {
    param(
        [double]$Bytes
    )

    if ($Bytes -le 0) {
        return 0
    }

    return [math]::Round($Bytes / 1GB, 2)
}

function ConvertTo-MarkdownSafe {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return "Unknown"
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "Unknown"
    }

    return ($text -replace "\|", "\\|").Trim()
}

function Get-ServerTargets {
    param(
        [string]$SearchBase,
        [string]$DomainController,
        [string]$ServerFilter,
        [bool]$IncludeDisabled,
        [int]$MaxServers
    )

    Import-Module ActiveDirectory -ErrorAction Stop

    $adParams = @{
        SearchBase = $SearchBase
        Filter     = "*"
        Properties = @(
            "DNSHostName",
            "Enabled",
            "OperatingSystem",
            "OperatingSystemVersion",
            "LastLogonDate",
            "IPv4Address"
        )
    }

    if ($DomainController) {
        $adParams.Server = $DomainController
    }

    $servers = Get-ADComputer @adParams |
        Where-Object { $_.Name -like $ServerFilter } |
        Where-Object { $_.OperatingSystem -match "Server" }

    if (-not $IncludeDisabled) {
        $servers = $servers | Where-Object Enabled
    }

    $servers = $servers | Sort-Object Name

    if ($MaxServers -gt 0) {
        $servers = $servers | Select-Object -First $MaxServers
    }

    return $servers
}

function Get-RemoteData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [System.Management.Automation.PSCredential]$Credential
    )

    $sessionParams = @{
        ComputerName = $ComputerName
    }

    if ($Credential) {
        $sessionParams.Credential = $Credential
    }

    $os = Get-CimInstance @sessionParams -ClassName Win32_OperatingSystem
    $computerSystem = Get-CimInstance @sessionParams -ClassName Win32_ComputerSystem
    $bios = Get-CimInstance @sessionParams -ClassName Win32_BIOS
    $logicalDisks = Get-CimInstance @sessionParams -ClassName Win32_LogicalDisk -Filter "DriveType = 3"
    $network = Get-CimInstance @sessionParams -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True"
    $hotFixes = Get-CimInstance @sessionParams -ClassName Win32_QuickFixEngineering

    [pscustomobject]@{
        OperatingSystem = $os
        ComputerSystem  = $computerSystem
        Bios            = $bios
        LogicalDisks    = $logicalDisks
        Network         = $network
        HotFixes        = $hotFixes
    }
}

function Get-ServerSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.ActiveDirectory.Management.ADComputer]$Computer,

        [System.Management.Automation.PSCredential]$Credential
    )

    $online = $false
    $errorMessage = $null
    $remoteData = $null

    try {
        $online = Test-Connection -ComputerName $Computer.Name -Count 1 -Quiet -ErrorAction Stop
    }
    catch {
        $online = $false
    }

    if ($online) {
        try {
            $remoteData = Get-RemoteData -ComputerName $Computer.Name -Credential $Credential
        }
        catch {
            $errorMessage = $_.Exception.Message
        }
    }
    else {
        $errorMessage = "Host did not respond to ping."
    }

    $latestHotfix = $null
    if ($remoteData -and $remoteData.HotFixes) {
        $latestHotfix = $remoteData.HotFixes |
            Sort-Object -Property InstalledOn -Descending |
            Select-Object -First 1
    }

    [pscustomobject]@{
        Name              = $Computer.Name
        DnsHostName       = $Computer.DNSHostName
        Enabled           = $Computer.Enabled
        ADOperatingSystem = $Computer.OperatingSystem
        ADOSVersion       = $Computer.OperatingSystemVersion
        LastLogonDate     = $Computer.LastLogonDate
        IPv4Address       = $Computer.IPv4Address
        Online            = $online
        ErrorMessage      = $errorMessage
        RemoteData        = $remoteData
        LatestHotfix      = $latestHotfix
        GeneratedAt       = Get-Date
    }
}

function Get-StatusColor {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Snapshot
    )

    if (-not $Snapshot.Online) {
        return "Red"
    }

    if ($Snapshot.ErrorMessage) {
        return "Orange"
    }

    return "Green"
}

function Get-StatusLabel {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Snapshot
    )

    if (-not $Snapshot.Online) {
        return "Offline"
    }

    if ($Snapshot.ErrorMessage) {
        return "Needs Attention"
    }

    return "Online"
}

function Get-AdmonitionType {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Snapshot
    )

    if (-not $Snapshot.Online) {
        return "danger"
    }

    if ($Snapshot.ErrorMessage) {
        return "warning"
    }

    return "success"
}

function Get-DomainRoleName {
    param(
        [AllowNull()]
        [int]$DomainRole
    )

    switch ($DomainRole) {
        0 { return "Standalone Workstation" }
        1 { return "Member Workstation" }
        2 { return "Standalone Server" }
        3 { return "Member Server" }
        4 { return "Backup Domain Controller" }
        5 { return "Primary Domain Controller" }
        default { return "Unknown" }
    }
}

function Get-PatchAgeSummary {
    param(
        [AllowNull()]
        [object]$LatestHotFix,

        [datetime]$AsOf
    )

    if (-not $LatestHotFix -or -not $LatestHotFix.InstalledOn) {
        return [pscustomobject]@{
            Text           = "No patch date available"
            AdmonitionType = "warning"
        }
    }

    $age = (New-TimeSpan -Start (Get-Date $LatestHotFix.InstalledOn) -End $AsOf).Days

    if ($age -le 30) {
        return [pscustomobject]@{
            Text           = "$age day(s) since latest visible patch"
            AdmonitionType = "success"
        }
    }

    if ($age -le 60) {
        return [pscustomobject]@{
            Text           = "$age day(s) since latest visible patch"
            AdmonitionType = "warning"
        }
    }

    return [pscustomobject]@{
        Text           = "$age day(s) since latest visible patch"
        AdmonitionType = "danger"
    }
}

function Get-DriveRiskSummary {
    param(
        [AllowNull()]
        [object[]]$Drives
    )

    if (-not $Drives) {
        return [pscustomobject]@{
            Text           = "No storage data collected"
            AdmonitionType = "warning"
        }
    }

    $criticalDrives = @()
    foreach ($drive in $Drives) {
        if ($drive.Size -le 0) {
            continue
        }

        $freePct = ($drive.FreeSpace / $drive.Size) * 100
        if ($freePct -lt 10) {
            $criticalDrives += $drive.DeviceID
        }
    }

    if ($criticalDrives.Count -gt 0) {
        return [pscustomobject]@{
            Text           = "Low disk space detected on: $($criticalDrives -join ', ')"
            AdmonitionType = "danger"
        }
    }

    return [pscustomobject]@{
        Text           = "No drives below 10% free space"
        AdmonitionType = "success"
    }
}

function New-MetricCards {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Items
    )

    $lines = @(
        '<div class="grid cards" markdown>',
        ''
    )

    foreach ($item in $Items) {
        $lines += "-   **$($item.Title)**"
        $lines += ""
        $lines += "    $($item.Value)"
        if ($item.Note) {
            $lines += ""
            $lines += "    $($item.Note)"
        }
        $lines += ""
    }

    $lines += "</div>"
    return ($lines -join [Environment]::NewLine)
}

function New-DriveTable {
    param(
        [AllowNull()]
        [object[]]$Drives
    )

    if (-not $Drives) {
        return (@(
            "| Drive | Label | Size (GiB) | Free (GiB) | Free % |",
            "| --- | --- | ---: | ---: | ---: |",
            "| n/a | n/a | 0 | 0 | 0% |"
        ) -join [Environment]::NewLine)
    }

    $lines = @(
        "| Drive | Label | Size (GiB) | Free (GiB) | Free % |",
        "| --- | --- | ---: | ---: | ---: |"
    )

    foreach ($drive in $Drives | Sort-Object DeviceID) {
        $sizeGiB = ConvertTo-GiB -Bytes $drive.Size
        $freeGiB = ConvertTo-GiB -Bytes $drive.FreeSpace
        $freePct = if ($drive.Size -gt 0) { [math]::Round(($drive.FreeSpace / $drive.Size) * 100, 1) } else { 0 }

        $lines += "| $(ConvertTo-MarkdownSafe $drive.DeviceID) | $(ConvertTo-MarkdownSafe $drive.VolumeName) | $sizeGiB | $freeGiB | $freePct% |"
    }

    return ($lines -join [Environment]::NewLine)
}

function New-NetworkList {
    param(
        [AllowNull()]
        [object[]]$Adapters
    )

    if (-not $Adapters) {
        return "- No active network adapters discovered."
    }

    $lines = foreach ($adapter in $Adapters) {
        $ips = @($adapter.IPAddress | Where-Object { $_ -match "^\d{1,3}(\.\d{1,3}){3}$" }) -join ", "
        $gws = @($adapter.DefaultIPGateway) -join ", "
        $dns = @($adapter.DNSServerSearchOrder) -join ", "
        "- **$(ConvertTo-MarkdownSafe $adapter.Description)**: IP $ips, Gateway $gws, DNS $dns"
    }

    return ($lines -join [Environment]::NewLine)
}

function New-HotfixTable {
    param(
        [AllowNull()]
        [object[]]$HotFixes
    )

    if (-not $HotFixes) {
        return (@(
            "| Installed On | Hotfix ID | Description |",
            "| --- | --- | --- |",
            "| n/a | n/a | No hotfix inventory available |"
        ) -join [Environment]::NewLine)
    }

    $lines = @(
        "| Installed On | Hotfix ID | Description |",
        "| --- | --- | --- |"
    )

    foreach ($hotfix in $HotFixes | Sort-Object InstalledOn -Descending | Select-Object -First 10) {
        $installedOn = if ($hotfix.InstalledOn) { (Get-Date $hotfix.InstalledOn).ToString("yyyy-MM-dd") } else { "Unknown" }
        $lines += "| $installedOn | $(ConvertTo-MarkdownSafe $hotfix.HotFixID) | $(ConvertTo-MarkdownSafe $hotfix.Description) |"
    }

    return ($lines -join [Environment]::NewLine)
}

function Get-LatestPatchText {
    param(
        [AllowNull()]
        [object]$HotFix
    )

    if (-not $HotFix) {
        return "Unknown"
    }

    $installedOn = if ($HotFix.InstalledOn) { (Get-Date $HotFix.InstalledOn).ToString("yyyy-MM-dd") } else { "Unknown" }
    return "$($HotFix.HotFixID) on $installedOn"
}

function New-ServerMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Snapshot
    )

    $statusColor = Get-StatusColor -Snapshot $Snapshot
    $generatedAt = $Snapshot.GeneratedAt.ToString("yyyy-MM-dd HH:mm:ss zzz")
    $lastLogon = if ($Snapshot.LastLogonDate) { (Get-Date $Snapshot.LastLogonDate).ToString("yyyy-MM-dd HH:mm") } else { "Unknown" }

    $remoteData = $Snapshot.RemoteData
    $os = if ($remoteData) { $remoteData.OperatingSystem } else { $null }
    $computerSystem = if ($remoteData) { $remoteData.ComputerSystem } else { $null }
    $bios = if ($remoteData) { $remoteData.Bios } else { $null }
    $hotfixes = if ($remoteData) { $remoteData.HotFixes } else { $null }
    $drives = if ($remoteData) { $remoteData.LogicalDisks } else { $null }
    $network = if ($remoteData) { $remoteData.Network } else { $null }

    $memoryGiB = if ($computerSystem.TotalPhysicalMemory) {
        [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
    }
    else {
        "Unknown"
    }

    $uptime = if ($os.LastBootUpTime) {
        New-TimeSpan -Start $os.LastBootUpTime -End $Snapshot.GeneratedAt
    }
    else {
        $null
    }

    $uptimeText = if ($uptime) {
        "{0}d {1}h {2}m" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
    }
    else {
        "Unknown"
    }

    $latestPatchText = Get-LatestPatchText -HotFix $Snapshot.LatestHotfix
    $patchAgeSummary = Get-PatchAgeSummary -LatestHotFix $Snapshot.LatestHotfix -AsOf $Snapshot.GeneratedAt
    $driveRiskSummary = Get-DriveRiskSummary -Drives $drives
    $statusLabel = Get-StatusLabel -Snapshot $Snapshot
    $admonitionType = Get-AdmonitionType -Snapshot $Snapshot

    $totalDiskGiB = if ($drives) { [math]::Round((($drives | Measure-Object -Property Size -Sum).Sum) / 1GB, 2) } else { 0 }
    $freeDiskGiB = if ($drives) { [math]::Round((($drives | Measure-Object -Property FreeSpace -Sum).Sum) / 1GB, 2) } else { 0 }
    $freeDiskPct = if ($totalDiskGiB -gt 0) { [math]::Round(($freeDiskGiB / $totalDiskGiB) * 100, 1) } else { 0 }
    $adapterCount = @($network).Count
    $domainRole = if ($computerSystem) { Get-DomainRoleName -DomainRole $computerSystem.DomainRole } else { "Unknown" }
    $metricCards = New-MetricCards -Items @(
        [pscustomobject]@{ Title = "Status"; Value = $statusLabel; Note = "Reachable: $(if ($Snapshot.Online) { 'Yes' } else { 'No' })" },
        [pscustomobject]@{ Title = "Operating System"; Value = (ConvertTo-MarkdownSafe $(if ($os.Caption) { $os.Caption } else { $Snapshot.ADOperatingSystem })); Note = "Version: $(ConvertTo-MarkdownSafe $(if ($os.Version) { $os.Version } else { $Snapshot.ADOSVersion }))" },
        [pscustomobject]@{ Title = "Patch Level"; Value = (ConvertTo-MarkdownSafe $latestPatchText); Note = $patchAgeSummary.Text },
        [pscustomobject]@{ Title = "Storage"; Value = "$freeDiskGiB GiB free of $totalDiskGiB GiB"; Note = $driveRiskSummary.Text },
        [pscustomobject]@{ Title = "Memory"; Value = "$memoryGiB GiB"; Note = "Role: $domainRole" },
        [pscustomobject]@{ Title = "Connectivity"; Value = "$adapterCount active adapter(s)"; Note = "AD Last Logon: $lastLogon" }
    )

    $lines = @(
        "---",
        "title: $($Snapshot.Name)",
        "tags:",
        "  - server",
        "  - inventory",
        "---",
        "",
        "# $($Snapshot.Name)",
        "",
        "!!! $admonitionType ""Server Status: $statusLabel""",
        "    Generated: $generatedAt",
        "    ",
        "    DNS Name: $(ConvertTo-MarkdownSafe $Snapshot.DnsHostName)",
        "    ",
        "    Latest Patch Seen: $(ConvertTo-MarkdownSafe $latestPatchText)",
        "",
        $metricCards,
        "",
        "## Inventory Summary",
        "",
        "| Item | Value |",
        "| --- | --- |",
        "| DNS Name | $(ConvertTo-MarkdownSafe $Snapshot.DnsHostName) |",
        "| Status | $statusLabel |",
        "| Reachable | $(if ($Snapshot.Online) { 'Yes' } else { 'No' }) |",
        "| Active Directory Enabled | $(if ($Snapshot.Enabled) { 'Yes' } else { 'No' }) |",
        "| Operating System | $(ConvertTo-MarkdownSafe $(if ($os.Caption) { $os.Caption } else { $Snapshot.ADOperatingSystem })) |",
        "| OS Version | $(ConvertTo-MarkdownSafe $(if ($os.Version) { $os.Version } else { $Snapshot.ADOSVersion })) |",
        "| Last Boot | $(if ($os.LastBootUpTime) { (Get-Date $os.LastBootUpTime).ToString('yyyy-MM-dd HH:mm') } else { 'Unknown' }) |",
        "| Uptime | $uptimeText |",
        "| Latest Patch Seen | $(ConvertTo-MarkdownSafe $latestPatchText) |",
        "| Total Disk | $totalDiskGiB GiB |",
        "| Free Disk | $freeDiskGiB GiB ($freeDiskPct%) |",
        "| Memory | $memoryGiB GiB |",
        "| Domain Role | $domainRole |",
        "| AD Last Logon | $lastLogon |",
        "| IPv4 From AD | $(ConvertTo-MarkdownSafe $Snapshot.IPv4Address) |",
        "| Manufacturer / Model | $(ConvertTo-MarkdownSafe $computerSystem.Manufacturer) / $(ConvertTo-MarkdownSafe $computerSystem.Model) |",
        "| Serial Number | $(ConvertTo-MarkdownSafe $bios.SerialNumber) |"
    )

    if ($Snapshot.ErrorMessage) {
        $lines += @(
            "",
            "!!! warning ""Collection Notes""",
            "    $($Snapshot.ErrorMessage)"
        )
    }

    $lines += @(
        "",
        "!!! $($driveRiskSummary.AdmonitionType) ""Storage Health""",
        "    $($driveRiskSummary.Text)",
        "",
        "!!! $($patchAgeSummary.AdmonitionType) ""Patch Freshness""",
        "    $($patchAgeSummary.Text)",
        "",
        "## Storage",
        "",
        (New-DriveTable -Drives $drives),
        "",
        "??? info ""Network Details""",
        "",
        (New-NetworkList -Adapters $network),
        "",
        "??? info ""Recent Hotfixes""",
        "",
        (New-HotfixTable -HotFixes $hotfixes)
    )

    return ($lines -join [Environment]::NewLine)
}

function New-IndexMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Snapshots,

        [Parameter(Mandatory = $true)]
        [string]$SearchBase
    )

    $generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"
    $onlineCount = @($Snapshots | Where-Object Online).Count
    $attentionCount = @($Snapshots | Where-Object ErrorMessage).Count
    $offlineCount = @($Snapshots | Where-Object { -not $_.Online }).Count
    $indexCards = New-MetricCards -Items @(
        [pscustomobject]@{ Title = "Servers Found"; Value = "$($Snapshots.Count)"; Note = "Search Base: $SearchBase" },
        [pscustomobject]@{ Title = "Reachable"; Value = "$onlineCount"; Note = "Responded to ping and inventory collection attempted" },
        [pscustomobject]@{ Title = "Needs Attention"; Value = "$attentionCount"; Note = "Collection issue or partial data detected" },
        [pscustomobject]@{ Title = "Offline"; Value = "$offlineCount"; Note = "Did not respond during collection window" }
    )

    $lines = @(
        "---",
        "title: Server Inventory",
        "tags:",
        "  - server",
        "  - inventory",
        "---",
        "",
        "# Server Inventory",
        "",
        "!!! info ""Inventory Run""",
        "    Search base: $SearchBase",
        "    ",
        "    Generated: $generatedAt",
        "",
        $indexCards,
        "",
        "## Server List",
        "",
        "| Server | Status | Operating System | Latest Patch Seen | Free Disk |",
        "| --- | --- | --- | --- | --- |"
    )

    foreach ($snapshot in $Snapshots | Sort-Object Name) {
        $status = if (-not $snapshot.Online) {
            "Offline"
        }
        elseif ($snapshot.ErrorMessage) {
            "Needs Attention"
        }
        else {
            "Online"
        }

        $os = if ($snapshot.RemoteData -and $snapshot.RemoteData.OperatingSystem -and $snapshot.RemoteData.OperatingSystem.Caption) {
            $snapshot.RemoteData.OperatingSystem.Caption
        }
        else {
            $snapshot.ADOperatingSystem
        }

        $drives = if ($snapshot.RemoteData) { $snapshot.RemoteData.LogicalDisks } else { $null }
        $totalDiskGiB = if ($drives) { [math]::Round((($drives | Measure-Object -Property Size -Sum).Sum) / 1GB, 2) } else { 0 }
        $freeDiskGiB = if ($drives) { [math]::Round((($drives | Measure-Object -Property FreeSpace -Sum).Sum) / 1GB, 2) } else { 0 }
        $freePct = if ($totalDiskGiB -gt 0) { [math]::Round(($freeDiskGiB / $totalDiskGiB) * 100, 1) } else { 0 }
        $patch = Get-LatestPatchText -HotFix $snapshot.LatestHotfix

        $lines += "| [$($snapshot.Name)]($($snapshot.Name).md) | $status | $(ConvertTo-MarkdownSafe $os) | $(ConvertTo-MarkdownSafe $patch) | $freeDiskGiB GiB ($freePct%) |"
    }

    return ($lines -join [Environment]::NewLine)
}

New-DirectoryIfNeeded -Path $OutputPath

$servers = Get-ServerTargets `
    -SearchBase $SearchBase `
    -DomainController $DomainController `
    -ServerFilter $ServerFilter `
    -IncludeDisabled:$IncludeDisabled `
    -MaxServers $MaxServers

if (-not $servers) {
    throw "No servers were found under '$SearchBase' using filter '$ServerFilter'."
}

$snapshots = foreach ($server in $servers) {
    Write-Host "Collecting $($server.Name)..."
    $snapshot = Get-ServerSnapshot -Computer $server -Credential $Credential
    $markdown = New-ServerMarkdown -Snapshot $snapshot
    $filePath = Join-Path -Path $OutputPath -ChildPath ("{0}.md" -f $snapshot.Name)
    Set-Content -Path $filePath -Value $markdown -Encoding UTF8
    $snapshot
}

if (-not $SkipIndex) {
    $indexMarkdown = New-IndexMarkdown -Snapshots $snapshots -SearchBase $SearchBase
    Set-Content -Path (Join-Path -Path $OutputPath -ChildPath "index.md") -Value $indexMarkdown -Encoding UTF8
}

Write-Host "Generated $(@($snapshots).Count) server report(s) in '$OutputPath'."
