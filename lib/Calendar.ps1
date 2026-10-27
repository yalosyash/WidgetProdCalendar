# lib\Calendar.ps1
# Современная отрисовка: шапка с навигацией по кварталам, карточки месяцев
# (3 в ряд, при развороте — 2 ряда = 2 квартала), нижняя строка со сводкой.
# Данные НЕ подтягивает сам: получает готовые месяцы (коды + посчитанную
# статистику) и рисует. Поэтому не зависит от Data.ps1.

$script:CcMonthFull   = @('Январь','Февраль','Март','Апрель','Май','Июнь',
                          'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь')
$script:CcMonthShort  = @('янв','фев','мар','апр','май','июн','июл','авг','сен','окт','ноя','дек')
$script:CcWeekdays    = @('пн','вт','ср','чт','пт','сб','вс')
$script:CcWeekdaysFull= @('понедельник','вторник','среда','четверг','пятница','суббота','воскресенье')
$script:CcQuarter     = @('I','II','III','IV')
$script:CultureRu     = [System.Globalization.CultureInfo]::GetCultureInfo('ru-RU')

function Fmt1 { param([double]$v) return $v.ToString('0.0', $script:CultureRu) }

function New-RoundPath {
    param([double]$X, [double]$Y, [double]$W, [double]$H, [double]$R)
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    if ($R -le 0 -or $W -le 0 -or $H -le 0) {
        $p.AddRectangle((New-Object System.Drawing.RectangleF($X, $Y, [Math]::Max($W,1), [Math]::Max($H,1))))
        return $p
    }
    $d = $R * 2
    $p.AddArc($X, $Y, $d, $d, 180, 90)
    $p.AddArc(($X + $W - $d), $Y, $d, $d, 270, 90)
    $p.AddArc(($X + $W - $d), ($Y + $H - $d), $d, $d, 0, 90)
    $p.AddArc($X, ($Y + $H - $d), $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function New-CalendarUi {
    param([double]$Scale = 1.0)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $S = $Scale
    function Scl([double]$v) { [int][Math]::Round($v * $S) }

    $ui = @{ Scale = $S }
    $ui.CellW   = Scl 32; $ui.CellH = Scl 27
    $ui.CardPadX= Scl 12; $ui.CardPadY = Scl 10
    $ui.TitleH  = Scl 26; $ui.WkH = Scl 20
    $ui.GridW   = $ui.CellW * 7
    $ui.CardW   = $ui.GridW + 2 * $ui.CardPadX
    $ui.GridH   = $ui.CellH * 6
    $ui.StatH   = Scl 30
    $ui.CardH   = $ui.CardPadY + $ui.TitleH + $ui.WkH + $ui.GridH + $ui.StatH + $ui.CardPadY
    $ui.GapX    = Scl 12; $ui.GapY = Scl 12
    $ui.Pad     = Scl 16; $ui.Spacer = Scl 10
    $ui.HeaderH = Scl 50; $ui.FooterH = Scl 18
    $ui.Cols    = 3
    $ui.ButtonSize = Scl 34; $ui.BtnGap = Scl 8; $ui.Radius = Scl 10; $ui.CellRadius = Scl 7
    $ui.FormWidth  = $ui.Pad * 2 + $ui.Cols * $ui.CardW + ($ui.Cols - 1) * $ui.GapX
    $ui.Rows    = 1

    # палитра
    $ui.C_Bg        = [System.Drawing.Color]::FromArgb(0xF6,0xF7,0xFA)
    $ui.C_CardBg    = [System.Drawing.Color]::White
    $ui.C_CardBorder= [System.Drawing.Color]::FromArgb(0xE6,0xE8,0xEE)
    $ui.C_Shadow    = [System.Drawing.Color]::FromArgb(0xEC,0xEE,0xF3)
    $ui.C_Text      = [System.Drawing.Color]::FromArgb(0x1F,0x24,0x30)
    $ui.C_Muted     = [System.Drawing.Color]::FromArgb(0xA2,0xA8,0xB4)
    $ui.C_Red       = [System.Drawing.Color]::FromArgb(0xE5,0x48,0x4D)
    $ui.C_RedBg     = [System.Drawing.Color]::FromArgb(0xFE,0xF1,0xF1)
    $ui.C_Amber     = [System.Drawing.Color]::FromArgb(0xD9,0x82,0x0B)
    $ui.C_AmberBg   = [System.Drawing.Color]::FromArgb(0xFF,0xF7,0xEA)
    $ui.C_Accent    = [System.Drawing.Color]::FromArgb(0x4F,0x46,0xE5)
    $ui.C_AccentSoft= [System.Drawing.Color]::FromArgb(0xEE,0xF0,0xFE)
    $ui.C_Footer    = [System.Drawing.Color]::FromArgb(0x7A,0x81,0x94)
    $ui.C_Divider   = [System.Drawing.Color]::FromArgb(0xEF,0xF1,0xF5)

    # шрифты
    $ui.FTitle    = New-Object System.Drawing.Font('Segoe UI', [float](13.5 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FMonth    = New-Object System.Drawing.Font('Segoe UI', [float](11.5 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FWk       = New-Object System.Drawing.Font('Segoe UI', [float](8.5 * $S))
    $ui.FDay      = New-Object System.Drawing.Font('Segoe UI', [float](10 * $S))
    $ui.FDayBold  = New-Object System.Drawing.Font('Segoe UI', [float](10 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FStat     = New-Object System.Drawing.Font('Segoe UI', [float](7.8 * $S))
    $ui.FGlyph    = New-Object System.Drawing.Font('Segoe UI', [float](12 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FSmall    = New-Object System.Drawing.Font('Segoe UI', [float](8 * $S))

    # форма
    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'None'
    $form.ShowInTaskbar = $false
    $form.TopMost = $true
    $form.StartPosition = 'Manual'
    $form.KeyPreview = $true
    $form.BackColor = $ui.C_Bg

    $ui.Surface = $form
    $ui.Tip = New-Object System.Windows.Forms.ToolTip
    $ui.Bmp = $null
    $ui.Hits = @()
    return $ui
}

function Get-CalendarSurfaceSize {
    param($Ui, [int]$Rows)
    $w = $Ui.Pad * 2 + $Ui.Cols * $Ui.CardW + ($Ui.Cols - 1) * $Ui.GapX
    $h = $Ui.Pad + $Ui.HeaderH + $Ui.Spacer + $Rows * $Ui.CardH + ($Rows - 1) * $Ui.GapY + $Ui.Spacer + $Ui.FooterH + $Ui.Pad
    return @($w, $h)
}

function Draw-CalendarSurface {
    param(
        [Parameter(Mandatory)]$Ui,
        [Parameter(Mandatory)][array]$Months,
        [string]$HeaderTitle = '',
        [string]$FooterLeft = '',
        [string]$FooterRight = '',
        [string]$HoverKey = ''
    )

    $S = $Ui.Scale
    $rows = [int][Math]::Ceiling($Months.Count / $Ui.Cols)
    if ($rows -lt 1) { $rows = 1 }
    $Ui.Rows = $rows

    $w = $Ui.Surface.Width
    $h = $Ui.Surface.Height
    $bmp = New-Object System.Drawing.Bitmap([Math]::Max($w,1), [Math]::Max($h,1))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'ClearTypeGridFit'
    $g.Clear($Ui.C_Bg)

    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
    $sfL = New-Object System.Drawing.StringFormat
    $sfL.Alignment = 'Near'; $sfL.LineAlignment = 'Center'
    $sfR = New-Object System.Drawing.StringFormat
    $sfR.Alignment = 'Far'; $sfR.LineAlignment = 'Center'

    $brText   = New-Object System.Drawing.SolidBrush($Ui.C_Text)
    $brMuted  = New-Object System.Drawing.SolidBrush($Ui.C_Muted)
    $brRed    = New-Object System.Drawing.SolidBrush($Ui.C_Red)
    $brAmber  = New-Object System.Drawing.SolidBrush($Ui.C_Amber)
    $brAccent = New-Object System.Drawing.SolidBrush($Ui.C_Accent)
    $brWhite  = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $brFoot   = New-Object System.Drawing.SolidBrush($Ui.C_Footer)
    $hits = New-Object System.Collections.ArrayList

    # ---------- шапка ----------
    $btn = $Ui.ButtonSize
    $hy = $Ui.Pad + [int](($Ui.HeaderH - $btn) / 2)
    $prevRect = New-Object System.Drawing.Rectangle($Ui.Pad, $hy, $btn, $btn)
    $nextRect = New-Object System.Drawing.Rectangle(($w - $Ui.Pad - $btn), $hy, $btn, $btn)
    $expRect  = New-Object System.Drawing.Rectangle(($nextRect.X - $Ui.BtnGap - $btn), $hy, $btn, $btn)
    [void]$hits.Add(@{ Kind = 'prev';   Rect = $prevRect })
    [void]$hits.Add(@{ Kind = 'next';   Rect = $nextRect })
    [void]$hits.Add(@{ Kind = 'expand'; Rect = $expRect })

    foreach ($b in @(
        @{ R = $prevRect; Key = 'prev';   G = [char]0x2039 },
        @{ R = $nextRect; Key = 'next';   G = [char]0x203A },
        @{ R = $expRect;  Key = 'expand'; G = $(if ($rows -gt 1) { [char]0x25BC } else { [char]0x25B2 }) }
    )) {
        $rp = New-RoundPath $b.R.X $b.R.Y $b.R.Width $b.R.Height $Ui.Radius
        if ($HoverKey -eq $b.Key) {
            $hb = New-Object System.Drawing.SolidBrush($Ui.C_AccentSoft)
            $g.FillPath($hb, $rp); $hb.Dispose()
        }
        $g.DrawString([string]$b.G, $Ui.FGlyph, $brAccent, (New-Object System.Drawing.RectangleF($b.R.X, $b.R.Y, $b.R.Width, $b.R.Height)), $sf)
        $rp.Dispose()
    }

    $titleRect = New-Object System.Drawing.RectangleF(($prevRect.Right + $Ui.BtnGap), $Ui.Pad, ($expRect.X - $Ui.BtnGap - $prevRect.Right - $Ui.BtnGap), $Ui.HeaderH)
    $g.DrawString($HeaderTitle, $Ui.FTitle, $brText, $titleRect, $sf)

    # ---------- карточки месяцев ----------
    $today = (Get-Date).Date
    for ($i = 0; $i -lt $Months.Count; $i++) {
        $mo = $Months[$i]
        $col = $i % $Ui.Cols
        $row = [int][Math]::Floor($i / $Ui.Cols)
        $cx = $Ui.Pad + $col * ($Ui.CardW + $Ui.GapX)
        $cy = $Ui.Pad + $Ui.HeaderH + $Ui.Spacer + $row * ($Ui.CardH + $Ui.GapY)

        # тень
        $shp = New-RoundPath ($cx) ($cy + 2) $Ui.CardW $Ui.CardH $Ui.Radius
        $shb = New-Object System.Drawing.SolidBrush($Ui.C_Shadow)
        $g.FillPath($shb, $shp); $shb.Dispose(); $shp.Dispose()

        # карточка
        $cp = New-RoundPath $cx $cy $Ui.CardW $Ui.CardH $Ui.Radius
        $cb = New-Object System.Drawing.SolidBrush($Ui.C_CardBg)
        $g.FillPath($cb, $cp); $cb.Dispose()
        $penB = New-Object System.Drawing.Pen($Ui.C_CardBorder)
        $g.DrawPath($penB, $cp); $penB.Dispose(); $cp.Dispose()

        # заголовок месяца
        $mx = $cx + $Ui.CardPadX
        $my = $cy + $Ui.CardPadY
        $g.DrawString(("{0} {1}" -f $script:CcMonthFull[$mo.Month - 1], $mo.Year), $Ui.FMonth, $brText,
            (New-Object System.Drawing.RectangleF($mx, $my, $Ui.GridW, $Ui.TitleH)), $sfL)

        # карточка целиком — как область тултипа (дни добавляются позже и перекрывают)
        [void]$hits.Add(@{
            Kind = 'card'
            Rect = (New-Object System.Drawing.Rectangle($cx, $cy, $Ui.CardW, $Ui.CardH))
            Text = (Build-MonthTooltip -Year $mo.Year -Month $mo.Month -Stat $mo.Stat)
        })

        # дни недели
        $wy = $my + $Ui.TitleH
        for ($c = 0; $c -lt 7; $c++) {
            $wr = New-Object System.Drawing.RectangleF(($mx + $c * $Ui.CellW), $wy, $Ui.CellW, $Ui.WkH)
            $br = if ($c -ge 5) { $brRed } else { $brMuted }
            $g.DrawString($script:CcWeekdays[$c], $Ui.FWk, $br, $wr, $sf)
        }
        $penD = New-Object System.Drawing.Pen($Ui.C_Divider)
        $g.DrawLine($penD, $mx, ($wy + $Ui.WkH - 1), ($mx + $Ui.GridW), ($wy + $Ui.WkH - 1)); $penD.Dispose()

        $gridY = $wy + $Ui.WkH

        if ($null -eq $mo.Codes) {
            $g.DrawString('загрузка...', $Ui.FDay, $brMuted,
                (New-Object System.Drawing.RectangleF($mx, $gridY, $Ui.GridW, $Ui.GridH)), $sf)
        } else {
            $offset = 0
            for ($mm = 1; $mm -lt $mo.Month; $mm++) { $offset += [DateTime]::DaysInMonth($mo.Year, $mm) }
            $first = New-Object System.DateTime($mo.Year, $mo.Month, 1)
            $startCol = (([int]$first.DayOfWeek + 6) % 7)
            $days = [DateTime]::DaysInMonth($mo.Year, $mo.Month)

            for ($d = 1; $d -le $days; $d++) {
                $pos = $startCol + ($d - 1)
                $r = [int][Math]::Floor($pos / 7)
                $c = $pos % 7
                $x = $mx + $c * $Ui.CellW
                $y = $gridY + $r * $Ui.CellH
                $code = [string]$mo.Codes[$offset + ($d - 1)]
                $isToday = ($mo.Year -eq $today.Year -and $mo.Month -eq $today.Month -and $d -eq $today.Day)

                $rectF = New-Object System.Drawing.RectangleF($x, $y, $Ui.CellW, $Ui.CellH)
                $label = "$d"
                $font = $Ui.FDay
                $brush = $brText

                if ($isToday) {
                    $tp = New-RoundPath ($x + 2) ($y + 2) ($Ui.CellW - 4) ($Ui.CellH - 4) $Ui.CellRadius
                    $tb = New-Object System.Drawing.SolidBrush($Ui.C_Accent)
                    $g.FillPath($tb, $tp); $tb.Dispose(); $tp.Dispose()
                    $brush = $brWhite; $font = $Ui.FDayBold
                } elseif ($code -eq '1') {
                    $tp = New-RoundPath ($x + 2) ($y + 2) ($Ui.CellW - 4) ($Ui.CellH - 4) $Ui.CellRadius
                    $tb = New-Object System.Drawing.SolidBrush($Ui.C_RedBg)
                    $g.FillPath($tb, $tp); $tb.Dispose(); $tp.Dispose()
                    $brush = $brRed
                } elseif ($code -eq '2') {
                    $tp = New-RoundPath ($x + 2) ($y + 2) ($Ui.CellW - 4) ($Ui.CellH - 4) $Ui.CellRadius
                    $tb = New-Object System.Drawing.SolidBrush($Ui.C_AmberBg)
                    $g.FillPath($tb, $tp); $tb.Dispose(); $tp.Dispose()
                    $brush = $brAmber
                    $label = "$d*"
                }

                $g.DrawString($label, $font, $brush, $rectF, $sf)
                [void]$hits.Add(@{
                    Kind = 'day'
                    Rect = (New-Object System.Drawing.Rectangle($x, $y, $Ui.CellW, $Ui.CellH))
                    Text = (Build-DayTooltip -Year $mo.Year -Month $mo.Month -Day $d -Code $code)
                })
            }
        }

        # статистика месяца
        if ($null -ne $mo.Stat) {
            $stx = $mx
            $sty = $gridY + $Ui.GridH + [int]($Ui.Scale * 1)
            $half = [int]($Ui.StatH / 2)
            $l1 = "Рабочих {0}  ·  Выходных {1}" -f $mo.Stat.Work, $mo.Stat.Off
            $l2 = "{0}  ·  {1}  ·  {2}" -f (Fmt1 $mo.Stat.Hours40), (Fmt1 $mo.Stat.Hours36), (Fmt1 $mo.Stat.Hours24)
            $g.DrawString($l1, $Ui.FStat, $brMuted, (New-Object System.Drawing.RectangleF($stx, $sty, $Ui.GridW, $half)), $sf)
            $g.DrawString($l2, $Ui.FStat, $brText,  (New-Object System.Drawing.RectangleF($stx, ($sty + $half), $Ui.GridW, $half)), $sf)
        }
    }

    # ---------- футер ----------
    $fy = $h - $Ui.Pad - $Ui.FooterH
    $g.DrawString($FooterLeft, $Ui.FSmall, $brFoot, (New-Object System.Drawing.RectangleF($Ui.Pad, $fy, ($w/2), $Ui.FooterH)), $sfL)
    $g.DrawString($FooterRight, $Ui.FSmall, $brFoot, (New-Object System.Drawing.RectangleF(($w/2), $fy, ($w/2 - $Ui.Pad), $Ui.FooterH)), $sfR)

    # ---------- готово ----------
    $brText.Dispose(); $brMuted.Dispose(); $brRed.Dispose(); $brAmber.Dispose()
    $brAccent.Dispose(); $brWhite.Dispose(); $brFoot.Dispose()
    $sf.Dispose(); $sfL.Dispose(); $sfR.Dispose()
    $g.Dispose()

    if ($Ui.Bmp) { $Ui.Bmp.Dispose() }
    $Ui.Bmp = $bmp
    $Ui.Surface.BackgroundImage = $bmp
    $Ui.Hits = $hits.ToArray()
}

function Build-DayTooltip {
    param([int]$Year, [int]$Month, [int]$Day, [string]$Code)
    $dt = New-Object System.DateTime($Year, $Month, $Day)
    $wd = $script:CcWeekdaysFull[([int]$dt.DayOfWeek + 6) % 7]
    $ds = "{0} {1} {2}, {3}" -f $Day, $script:CcMonthFull[$Month - 1].ToLower(), $Year, $wd
    switch ($Code) {
        '0' { return "$ds — рабочий день" }
        '1' { return "$ds — нерабочий (выходной/праздник)" }
        '2' { return "$ds — сокращённый день, −1 час" }
        default { return $ds }
    }
}

function Build-MonthTooltip {
    param([int]$Year, [int]$Month, $Stat)
    $name = "{0} {1}" -f $script:CcMonthFull[$Month - 1], $Year
    if ($null -eq $Stat) { return $name }
    return ("{0}: рабочих {1}, выходных {2}`r`nНорма: {3} / {4} / {5}" -f `
        $name, $Stat.Work, $Stat.Off, (Fmt1 $Stat.Hours40), (Fmt1 $Stat.Hours36), (Fmt1 $Stat.Hours24))
}
