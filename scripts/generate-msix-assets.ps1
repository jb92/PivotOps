# Generates the MSIX tile/logo assets required by Package.appxmanifest
# from the existing add-in logo. Run from the repository root:
#   pwsh -File scripts/generate-msix-assets.ps1

[CmdletBinding()]
param(
    [string]$Source,
    [string]$OutputDir
)

Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Source) { $Source = Join-Path $root '..\assets\logo-300.png' }
if (-not $OutputDir) { $OutputDir = Join-Path $root '..\desktop\PivotOps.Desktop\Assets' }

$Source = (Resolve-Path $Source).Path
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$OutputDir = (Resolve-Path $OutputDir).Path

# name = width x height; the logo is centred and scaled to fit.
$targets = @(
    @{ Name = 'StoreLogo.png';                                  W = 50;  H = 50 },
    @{ Name = 'Square44x44Logo.png';                             W = 44;  H = 44 },
    @{ Name = 'Square44x44Logo.targetsize-24_altform-unplated.png'; W = 24; H = 24 },
    @{ Name = 'Square150x150Logo.png';                           W = 150; H = 150 },
    @{ Name = 'Wide310x150Logo.png';                             W = 310; H = 150 },
    @{ Name = 'SplashScreen.png';                                W = 620; H = 300 }
)

$image = [System.Drawing.Image]::FromFile($Source)
try {
    foreach ($t in $targets) {
        $bitmap = New-Object System.Drawing.Bitmap($t.W, $t.H)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.Clear([System.Drawing.Color]::Transparent)

            $scale = [Math]::Min($t.W / $image.Width, $t.H / $image.Height)
            $w = [int]($image.Width * $scale)
            $h = [int]($image.Height * $scale)
            $graphics.DrawImage($image, [int](($t.W - $w) / 2), [int](($t.H - $h) / 2), $w, $h)

            $path = Join-Path $OutputDir $t.Name
            $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
            Write-Host "wrote $path"
        }
        finally {
            $graphics.Dispose()
            $bitmap.Dispose()
        }
    }
}
finally {
    $image.Dispose()
}
