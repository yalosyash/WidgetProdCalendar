# tools\New-Preview.ps1
# Рисует assets\preview.png — картинку виджета для README, без запуска самого
# виджета и без снимка экрана: те же функции отрисовки, что и в приложении.
# Запуск:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\New-Preview.ps1

[CmdletBinding()]
param([string]$OutPath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $projectRoot 'lib\App.ps1')

if (-not $OutPath) { $OutPath = Join-Path $projectRoot 'assets\preview.png' }

# --- содержимое окна ---------------------------------------------------------
# Масштаб 2: картинка должна быть читаемой в README, где она уменьшается.
$ui = New-CalendarUi -Scale 2.0
$codes = ConvertTo-CodesTable (Read-JsonFile $script:OfflineFile)
$today = (Get-Date).Date
$year = $today.Year
if (-not $codes.ContainsKey("$year")) { $year = ($codes.Keys | ForEach-Object { [int]$_ } | Sort-Object)[-1] }
$codeStr = [string]$codes["$year"]
$quarter = [int][Math]::Floor(($today.Month - 1) / 3) + 1

$months = New-Object System.Collections.ArrayList
foreach ($m in (Get-QuarterMonths -Year $year -Quarter $quarter)) {
    [void]$months.Add([pscustomobject]@{
            Year = $year; Month = $m; Codes = $codeStr
            Stat = (Get-MonthStat -Codes $codeStr -Month $m -Year $year)
        })
}

$size = Get-CalendarSurfaceSize -Ui $ui -Rows 1
$ui.Surface.ClientSize = New-Object System.Drawing.Size($size[0], $size[1])
Draw-CalendarSurface -Ui $ui -Months $months.ToArray() `
    -HeaderTitle ("{0} · {1} квартал" -f $year, $script:CcQuarter[$quarter - 1]) `
    -FooterLeft 'данные: офлайн' `
    -FooterRight "$([char]0x25B2) развернуть · ‹ › или колесо — кварталы" `
    -HoverKey '' -ShowArrows $true

# --- тень виджета (как в Update-WidgetShadow) --------------------------------
$w = $ui.Surface.Width; $h = $ui.Surface.Height
$m = [int]$ui.ShadowMargin; $radius = [int]$ui.WindowRadius
$shadow = New-Object System.Drawing.Bitmap(($w + $m * 2), ($h + $m * 2 + [int]$ui.ShadowY), [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$sg = [System.Drawing.Graphics]::FromImage($shadow)
$sg.SmoothingMode = 'AntiAlias'
$sg.Clear([System.Drawing.Color]::Transparent)
Add-SoftShadow -Graphics $sg -X $m -Y $m -W $w -H $h -Margin $m -Radius $radius `
    -OffsetY ([int]$ui.ShadowY) -Color $ui.C_ShadowSoft -Peak 16.0
$card = New-RoundPath $m $m $w $h $radius
$sg.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
$clear = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Transparent)
$sg.FillPath($clear, $card); $clear.Dispose(); $card.Dispose()
$sg.Dispose()

# --- подсказка ---------------------------------------------------------------
# Берём сегодняшний день, если он в кадре: подсказка показывает и статус дня,
# и её оформление. Иначе — первый день недели с нерабочим статусом.
$dayHits = @($ui.Hits | Where-Object { $_.Kind -eq 'day' })
$tipHit = $dayHits | Where-Object { $_.Tip.Badge -eq 'Сегодня' } | Select-Object -First 1
if ($null -eq $tipHit) { $tipHit = $dayHits | Where-Object { $_.Tip.Tone -eq 'off' } | Select-Object -First 1 }
if ($null -eq $tipHit) { $tipHit = $dayHits | Select-Object -First 1 }
if ($null -eq $tipHit) { throw 'В разметке нет дней — подсказку для скриншота не построить.' }
$tipBmp = New-CalendarTooltipBitmap -Ui $ui -Tip $tipHit.Tip

# Подсказку ставим сбоку от виджета, а не поверх: на скриншоте не должно быть
# закрытых дней и подписей.
$pad = 34
$side = [int](26 * $ui.Scale)
$tipX = $m * 2 + $w + $side
$tipY = $m + $tipHit.Rect.Top - [int]($tipBmp.Height / 2) + [int]($tipHit.Rect.Height / 2)
if ($tipY -lt $pad) { $tipY = $pad }

$cw = $pad + $tipX + $tipBmp.Width + $pad
$ch = [Math]::Max(($m * 2 + $h), ($tipY + $tipBmp.Height)) + $pad * 2
$comp = New-Object System.Drawing.Bitmap($cw, $ch)
$cg = [System.Drawing.Graphics]::FromImage($comp)
$cg.SmoothingMode = 'AntiAlias'
# Фон нейтральный, светлее виджета: тень по краям окна должна быть видна.
$cg.Clear([System.Drawing.Color]::FromArgb(0xE9, 0xEB, 0xF1))
$cg.DrawImage($shadow, $pad, $pad)
# У настоящего окна углы срезает регион (Set-RoundFormRegion), а bitmap окна
# прямоугольный. Обрезаем его тем же скруглением — иначе на скриншоте углы
# будут прямыми, а из-под них полезет фон виджета.
$clip = New-RoundPath ($pad + $m) ($pad + $m) $w $h $radius
$cg.SetClip($clip)
$cg.DrawImage($ui.Bmp, ($pad + $m), ($pad + $m))
$cg.ResetClip()
$clip.Dispose()
$cg.DrawImage($tipBmp, ($pad + $tipX), ($pad + $tipY))
$cg.Dispose()

$dir = Split-Path -Parent $OutPath
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$comp.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)

$comp.Dispose(); $shadow.Dispose(); $tipBmp.Dispose()
$ui.Bmp.Dispose(); $ui.Surface.Dispose(); $ui.TipForm.Dispose(); $ui.ShadowForm.Dispose()

Write-Host "Скриншот создан: $OutPath"
Write-Host ("Размер: {0}x{1}" -f $cw, $ch)
