param(
    [Parameter(Mandatory = $true)]
    [string]$SearchBase,

    [string]$OutputPath = ".\docs\servers",

    [string]$DomainController,

    [string]$ServerFilter = "*",

    [int]$MaxServers = 0,

    [switch]$IncludeDisabled,

    [switch]$SkipIndex,

    [switch]$Push,

    [System.Management.Automation.PSCredential]$Credential
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$generatorPath = Join-Path -Path $scriptPath -ChildPath "Generate-ServerInventoryDocs.ps1"

if (-not (Test-Path -LiteralPath $generatorPath)) {
    throw "Could not find $generatorPath"
}

& $generatorPath `
    -SearchBase $SearchBase `
    -OutputPath $OutputPath `
    -DomainController $DomainController `
    -ServerFilter $ServerFilter `
    -MaxServers $MaxServers `
    -IncludeDisabled:$IncludeDisabled `
    -SkipIndex:$SkipIndex `
    -Credential $Credential

$status = git status --porcelain

if (-not $status) {
    Write-Host "No documentation changes detected."
    exit 0
}

git add $OutputPath
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"
git commit -m "Nightly server inventory update $timestamp"

if ($Push) {
    git push
}
