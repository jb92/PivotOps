# Builds, registers and launches the packaged PivotOps WinUI shell for local development.
#   pwsh -File scripts/run-desktop.ps1 -Configuration Debug -Architecture x64
#
# The app declares <WindowsPackageType>MSIX</WindowsPackageType>, so the built .exe cannot be
# started directly: PasswordVault and ApplicationData.Current require package identity. Instead
# the loose build output is registered from its AppxManifest.xml (requires Developer Mode).

[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [ValidateSet('x64', 'x86', 'ARM64')]
    [string]$Architecture = 'x64',
    [switch]$SkipBuild,
    [switch]$NoLaunch
)

$ErrorActionPreference = 'Stop'
$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$project = Join-Path $root 'desktop\PivotOps.Desktop\PivotOps.Desktop.csproj'
$rid = "win-$($Architecture.ToLower())"

if (-not $SkipBuild) {
    dotnet build $project -c $Configuration -r $rid -p:Platform=$Architecture
    if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
}

$manifest = Get-ChildItem -Path (Join-Path $root 'desktop\PivotOps.Desktop\bin') -Recurse -Filter 'AppxManifest.xml' -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -like "*\$Configuration\*\$rid\*" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $manifest) {
    throw "No AppxManifest.xml found for $Configuration/$rid. Build the project first."
}

# Add-AppxPackage ships in a Windows PowerShell module; PowerShell 7 needs the compatibility shim.
if ($PSVersionTable.PSEdition -eq 'Core') {
    Import-Module Appx -UseWindowsPowerShell -WarningAction SilentlyContinue
}

Write-Host "Registering $($manifest.FullName)" -ForegroundColor Cyan
Add-AppxPackage -Register $manifest.FullName

$xml = [xml](Get-Content -LiteralPath $manifest.FullName)
$identityName = $xml.Package.Identity.Name
$applicationId = @($xml.Package.Applications.Application)[0].Id
$package = Get-AppxPackage -Name $identityName
if (-not $package) { throw "Package '$identityName' is not registered." }

$aumid = "$($package.PackageFamilyName)!$applicationId"
Write-Host "AUMID: $aumid" -ForegroundColor Cyan

if (-not $NoLaunch) {
    Start-Process "shell:AppsFolder\$aumid"
}
