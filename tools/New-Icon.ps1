# tools\New-Icon.ps1
# Генерирует assets\calendar.ico (иконка календаря) без внешних файлов.
# Формат стандартный для Windows: BMP-записи для 16–128 px и PNG для 256 px.
# Запуск:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\New-Icon.ps1

[CmdletBinding()]
param([string]$OutPath)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$projectRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutPath) { $OutPath = Join-Path $projectRoot 'assets\calendar.ico' }

$sizes = 16, 20, 24, 32, 40, 48, 64, 96, 128, 256

function New-CalendarBitmap([int]$s) {
    $bmp = New-Object System.Drawing.Bitmap($s, $s, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.Clear([System.Drawing.Color]::Transparent)

    $u = [double]$s / 32.0            # единица масштаба: 1 = 1/32 стороны иконки
    function P([double]$v) { [single]([math]::Round($v * $u, 2)) }

    $m = P 2.0                        # внешний отступ
    $w = [single]($s - 2.0 * $m)
    $h = [single]($s - 2.0 * $m)
    $r = P 3.0                        # радиус скругления

    # Белая «страница» календаря с лёгкой тенью-обводкой.
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($m, $m, $r * 2, $r * 2, 180, 90)
    $path.AddArc($m + $w - $r * 2, $m, $r * 2, $r * 2, 270, 90)
    $path.AddArc($m + $w - $r * 2, $m + $h - $r * 2, $r * 2, $r * 2, 0, 90)
    $path.AddArc($m, $m + $h - $r * 2, $r * 2, $r * 2, 90, 90)
    $path.CloseFigure()

    $shadow = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60, 0, 0, 0))
    $g.FillPath($shadow, $path)
    $shadow.Dispose()

    $bg = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $g.FillPath($bg, $path)
    $bg.Dispose()

    $border = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(210, 205, 205), [single]([math]::Max(1.0, (P 0.8))))
    $g.DrawPath($border, $path)
    $border.Dispose()

    # Красная «шапка» календаря.
    $headH = P 8.5
    $head = New-Object System.Drawing.Drawing2D.GraphicsPath
    $head.AddArc($m, $m, $r * 2, $r * 2, 180, 90)
    $head.AddArc($m + $w - $r * 2, $m, $r * 2, $r * 2, 270, 90)
    $head.AddLine($m + $w, $m + $headH, $m, $m + $headH)
    $head.CloseFigure()
    $g.SetClip($head)
    $red = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(214, 45, 45))
    $g.FillRectangle($red, [single]$m, [single]$m, $w, $headH)
    $red.Dispose()
    $g.ResetClip()
    $head.Dispose()

    # Кольца подшивки.
    $ringW = [single]([math]::Max(1.0, (P 1.2)))
    $ringC = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(90, 30, 30, 30), $ringW)
    $ringC.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $ringC.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    foreach ($x in @(($m + $w * 0.28), ($m + $w * 0.72))) {
        $g.DrawLine($ringC, [single]$x, [single](P 0.8), [single]$x, [single](P 6.5))
    }
    $ringC.Dispose()

    # Сетка дней («выходные» серые), если сторона позволяет.
    if ($s -ge 24) {
        $cols = 4; $rows = 3
        $gridTop = $m + $headH + (P 1.6)
        $gridBottom = $m + $h - (P 2.2)
        $cellW = $w / $cols
        $cellH = ($gridBottom - $gridTop) / $rows
        $dot = [single]([math]::Min($cellW, $cellH) * 0.30)
        $work = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(70, 70, 80))
        $rest = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200, 205, 215))
        for ($ry = 0; $ry -lt $rows; $ry++) {
            for ($cx = 0; $cx -lt $cols; $cx++) {
                $cxx = [single]($m + $cellW * ($cx + 0.5))
                $cyy = [single]($gridTop + $cellH * ($ry + 0.5))
                $b = if ($cx -eq $cols - 1 -or $ry -eq $rows - 1) { $rest } else { $work }
                $g.FillEllipse($b, $cxx - $dot / 2, $cyy - $dot / 2, $dot, $dot)
            }
        }
        $work.Dispose(); $rest.Dispose()
    }

    $g.Dispose()
    return $bmp
}

function Get-PngBytes([System.Drawing.Bitmap]$bmp) {
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bytes = $ms.ToArray()
    $ms.Dispose()
    return , $bytes
}

# BMP-запись: BITMAPINFOHEADER + XOR (BGRA, снизу вверх) + AND-маска (1bpp).
function Get-DibBytes([System.Drawing.Bitmap]$bmp) {
    $s = $bmp.Width
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)
    $bw.Write([uint32]40)
    $bw.Write([int32]$s)
    $bw.Write([int32]($s * 2))      # высота = XOR + AND
    $bw.Write([uint16]1)
    $bw.Write([uint16]32)
    $bw.Write([uint32]0)            # BI_RGB
    $bw.Write([uint32]($s * $s * 4))
    $bw.Write([int32]0); $bw.Write([int32]0)
    $bw.Write([uint32]0); $bw.Write([uint32]0)
    for ($y = $s - 1; $y -ge 0; $y--) {
        for ($x = 0; $x -lt $s; $x++) {
            $c = $bmp.GetPixel($x, $y)
            $bw.Write([byte]$c.B); $bw.Write([byte]$c.G); $bw.Write([byte]$c.R); $bw.Write([byte]$c.A)
        }
    }
    $rowBytes = [int]([math]::Floor(($s + 31) / 32) * 4)
    $bw.Write((New-Object byte[] ($rowBytes * $s)))
    $bw.Flush()
    $bytes = $ms.ToArray()
    $bw.Dispose(); $ms.Dispose()
    return , $bytes
}

$entries = @()
foreach ($s in $sizes) {
    $bmp = New-CalendarBitmap $s
    if ($s -ge 256) { $data = Get-PngBytes $bmp } else { $data = Get-DibBytes $bmp }
    $bmp.Dispose()
    $entries += , @{ Size = $s; Data = $data }
}

$dir = Split-Path -Parent $OutPath
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$fs = [System.IO.File]::Create($OutPath)
$bw = New-Object System.IO.BinaryWriter($fs)
# ICONDIR
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$entries.Count)
$offset = 6 + 16 * $entries.Count
foreach ($e in $entries) {
    $dim = if ($e.Size -ge 256) { 0 } else { $e.Size }
    $bw.Write([byte]$dim); $bw.Write([byte]$dim)
    $bw.Write([byte]0); $bw.Write([byte]0)       # палитра / резерв
    $bw.Write([uint16]1); $bw.Write([uint16]32)  # плоскости / бит на пиксель
    $bw.Write([uint32]$e.Data.Length)
    $bw.Write([uint32]$offset)
    $offset += $e.Data.Length
}
foreach ($e in $entries) { $bw.Write($e.Data) }
$bw.Flush(); $bw.Dispose(); $fs.Dispose()

Write-Host "Иконка создана: $OutPath"
Write-Host ("Размеры: " + ($sizes -join ', '))
