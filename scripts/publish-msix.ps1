# Builds the PivotOps web payload and produces an MSIX package for the WinUI shell.
#   pwsh -File scripts/publish-msix.ps1 -Architecture x64
#
# The resulting .msix is written to desktop/PivotOps.Desktop/AppPackages.
# For a Store submission, first replace Identity/Publisher in Package.appxmanifest
# with the values assigned in Partner Center, then upload the unsigned package.

[CmdletBinding()]
param(
    [ValidateSet('x64', 'x86', 'ARM64')]
    [string]$Architecture = 'x64',
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$root = Resolve-Path (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) '..')
$project = Join-Path $root 'desktop\PivotOps.Desktop\PivotOps.Desktop.csproj'
$rid = "win-$($Architecture.ToLower())"

Push-Location $root
try {
    npm run build:desktop
    if ($LASTEXITCODE -ne 0) { throw 'Web payload build failed.' }

    dotnet publish $project `
        -c $Configuration `
        -r $rid `
        -p:Platform=$Architecture `
        -p:SelfContained=true `
        -p:WindowsAppSDKSelfContained=true `
        -p:GenerateAppxPackageOnBuild=true `
        -p:UapAppxPackageBuildMode=StoreUpload `
        -p:AppxPackageSigningEnabled=false
    if ($LASTEXITCODE -ne 0) { throw 'MSIX packaging failed.' }
}
finally {
    Pop-Location
}

Get-ChildItem (Join-Path $root 'desktop\PivotOps.Desktop\AppPackages') -Recurse -Include *.msix, *.msixupload |
    Select-Object -ExpandProperty FullName
