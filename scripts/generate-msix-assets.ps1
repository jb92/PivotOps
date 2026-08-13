# Generates the MSIX tile/logo assets and the window/exe icon required by
# Package.appxmanifest from the existing add-in logo. Run from the repository root:
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

function New-ScaledBitmap {
    param([System.Drawing.Image]$Image, [int]$Width, [int]$Height)

    $bitmap = New-Object System.Drawing.Bitmap($Width, $Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $scale = [Math]::Min($Width / $Image.Width, $Height / $Image.Height)
        $w = [int]($Image.Width * $scale)
        $h = [int]($Image.Height * $scale)
        $graphics.DrawImage($Image, [int](($Width - $w) / 2), [int](($Height - $h) / 2), $w, $h)
    }
    finally {
        $graphics.Dispose()
    }

    return $bitmap
}

# System.Drawing cannot write multi-resolution icons, so the ICO container is
# assembled by hand around PNG-compressed frames (supported since Windows Vista).
function Write-Icon {
    param([System.Drawing.Image]$Image, [string]$Path, [int[]]$Sizes)

    $frames = foreach ($size in $Sizes) {
        $bitmap = New-ScaledBitmap -Image $Image -Width $size -Height $size
        $stream = New-Object System.IO.MemoryStream
        try {
            $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
            [pscustomobject]@{ Size = $size; Bytes = $stream.ToArray() }
        }
        finally {
            $stream.Dispose()
            $bitmap.Dispose()
        }
    }

    $file = [System.IO.File]::Create($Path)
    $writer = New-Object System.IO.BinaryWriter($file)
    try {
        $writer.Write([UInt16]0)              # reserved
        $writer.Write([UInt16]1)              # type: icon
        $writer.Write([UInt16]$frames.Count)

        $offset = 6 + (16 * $frames.Count)
        foreach ($frame in $frames) {
            $dimension = if ($frame.Size -ge 256) { 0 } else { $frame.Size }
            $writer.Write([Byte]$dimension)   # width
            $writer.Write([Byte]$dimension)   # height
            $writer.Write([Byte]0)            # palette size
            $writer.Write([Byte]0)            # reserved
            $writer.Write([UInt16]1)          # colour planes
            $writer.Write([UInt16]32)         # bits per pixel
            $writer.Write([UInt32]$frame.Bytes.Length)
            $writer.Write([UInt32]$offset)
            $offset += $frame.Bytes.Length
        }

        foreach ($frame in $frames) { $writer.Write($frame.Bytes) }
    }
    finally {
        $writer.Dispose()
        $file.Dispose()
    }

    Write-Host "wrote $Path"
}

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
        $bitmap = New-ScaledBitmap -Image $image -Width $t.W -Height $t.H
        try {
            $path = Join-Path $OutputDir $t.Name
            $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
            Write-Host "wrote $path"
        }
        finally {
            $bitmap.Dispose()
        }
    }

    Write-Icon -Image $image -Path (Join-Path $OutputDir 'PivotOps.ico') -Sizes @(16, 20, 24, 32, 40, 48, 64, 128, 256)
}
finally {
    $image.Dispose()
}
