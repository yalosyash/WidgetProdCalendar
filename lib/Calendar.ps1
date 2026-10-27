# lib\Calendar.ps1
# Построение интерфейса календаря (WinForms) и его отрисовка.
# Интерфейс умышленно рисуется на одном Panel (полный контроль над цветами и
# подсветкой «сегодня»), кнопки навигации — отдельные Button-ы.

# Месяцы (дублируем, чтобы файл был самодостаточен при отдельном dot-source).
$script:CcMonthFull = @('Январь','Февраль','Март','Апрель','Май','Июнь',
                        'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь')
$script:CcWeekdays = @('Пн','Вт','Ср','Чт','Пт','Сб','Вс')
$script:CcWeekdaysFull = @('понедельник','вторник','среда','четверг','пятница','суббота','воскресенье')

function New-CalendarUi {
    [CmdletBinding()]
    param(
        [double]$Scale = 1.0,
        [double]$TodayOpacity = 1.0
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $S = $Scale
    # Имя Scl (не sc!) — 'sc' является алиасом Set-Content и перебил бы функцию.
    function Scl([double]$v) { [int][Math]::Round($v * $S) }

    $cellW = Scl 46; $cellH = Scl 34
    $gridW = $cellW * 7
    $headH = Scl 40            # шапка с месяцем/годом
    $wkH   = Scl 22            # строка Пн..Вс
    $gridH = $wkH + $cellH * 6
    $sumH  = Scl 92            # панель сводки
    $srcH  = Scl 20            # строка источника
    $pad   = Scl 10

    $w = $gridW + $pad * 2
    $h = $headH + $pad + $gridH + $pad + $sumH + $srcH + $pad

    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'None'
    $form.ShowInTaskbar = $false
    $form.TopMost = $true
    $form.StartPosition = 'Manual'
    $form.KeyPreview = $true
    $form.ClientSize = New-Object System.Drawing.Size($w, $h)
    $form.BackColor = [System.Drawing.Color]::White

    # --- Шапка --------------------------------------------------------------
    $header = New-Object System.Windows.Forms.Panel
    $header.Location = New-Object System.Drawing.Point($pad, 0)
    $header.Size = New-Object System.Drawing.Size($gridW, $headH)
    $header.BackColor = [System.Drawing.Color]::FromArgb(0xF3, 0xF6, 0xFA)
    $form.Controls.Add($header)

    $btnFont = New-Object System.Drawing.Font('Segoe UI', [float](11 * $S), [System.Drawing.FontStyle]::Bold)
    $btnPrev = New-Object System.Windows.Forms.Button
    $btnPrev.Text = [char]0x2039
    $btnPrev.Font = $btnFont
    $btnPrev.FlatStyle = 'Flat'
    $btnPrev.FlatAppearance.BorderSize = 0
    $btnPrev.Size = New-Object System.Drawing.Size($headH, $headH)
    $btnPrev.Location = New-Object System.Drawing.Point(0, 0)
    $header.Controls.Add($btnPrev)

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = [char]0x203A
    $btnNext.Font = $btnFont
    $btnNext.FlatStyle = 'Flat'
    $btnNext.FlatAppearance.BorderSize = 0
    $btnNext.Size = New-Object System.Drawing.Size($headH, $headH)
    $btnNext.Location = New-Object System.Drawing.Point(($gridW - $headH), 0)
    $header.Controls.Add($btnNext)

    $titleLbl = New-Object System.Windows.Forms.Label
    $titleLbl.TextAlign = 'MiddleCenter'
    $titleLbl.Font = New-Object System.Drawing.Font('Segoe UI', [float](12 * $S), [System.Drawing.FontStyle]::Bold)
    $titleLbl.ForeColor = [System.Drawing.Color]::FromArgb(0x1F, 0x30, 0x50)
    $titleLbl.BackColor = [System.Drawing.Color]::Transparent
    $titleLbl.Location = New-Object System.Drawing.Point($headH, 0)
    $titleLbl.Size = New-Object System.Drawing.Size(($gridW - 2 * $headH), $headH)
    $header.Controls.Add($titleLbl)

    # --- Сетка --------------------------------------------------------------
    $grid = New-Object System.Windows.Forms.Panel
    $grid.Location = New-Object System.Drawing.Point($pad, ($headH + $pad))
    $grid.Size = New-Object System.Drawing.Size($gridW, $gridH)
    $grid.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($grid)

    # --- Сводка -------------------------------------------------------------
    $summary = New-Object System.Windows.Forms.Label
    $summary.Location = New-Object System.Drawing.Point($pad, ($headH + $pad + $gridH + $pad))
    $summary.Size = New-Object System.Drawing.Size($gridW, $sumH)
    $summary.Font = New-Object System.Drawing.Font('Segoe UI', [float](9 * $S))
    $summary.ForeColor = [System.Drawing.Color]::FromArgb(0x22, 0x22, 0x22)
    $summary.BackColor = [System.Drawing.Color]::FromArgb(0xFA, 0xFA, 0xFA)
    $form.Controls.Add($summary)

    # --- Источник данных ----------------------------------------------------
    $sourceLbl = New-Object System.Windows.Forms.Label
    $sourceLbl.Location = New-Object System.Drawing.Point($pad, ($headH + $pad + $gridH + $pad + $sumH))
    $sourceLbl.Size = New-Object System.Drawing.Size($gridW, $srcH)
    $sourceLbl.Font = New-Object System.Drawing.Font('Segoe UI', [float](7.5 * $S))
    $sourceLbl.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($sourceLbl)

    $tip = New-Object System.Windows.Forms.ToolTip

    return @{
        Form     = $form
        Header   = $header
        Title    = $titleLbl
        BtnPrev  = $btnPrev
        BtnNext  = $btnNext
        Grid     = $grid
        Summary  = $summary
        Source   = $sourceLbl
        Tip      = $tip
        Scale    = $S
        CellW    = $cellW
        CellH    = $cellH
        HeadH    = $headH
        WkH      = $wkH
        Pad      = $pad
    }
}

function Draw-CalendarGrid {
    param(
        [Parameter(Mandatory)]$Ui,
        [Parameter(Mandatory)][int]$Year,
        [Parameter(Mandatory)][int]$Month,
        [Parameter(Mandatory)][string]$Codes
    )

    $S = $Ui.Scale
    $cw = $Ui.CellW; $ch = $Ui.CellH; $wkH = $Ui.WkH
    $gridW = $cw * 7
    $today = (Get-Date).Date

    $bmp = New-Object System.Drawing.Bitmap($Ui.Grid.Width, $Ui.Grid.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'ClearTypeGridFit'
    $g.Clear([System.Drawing.Color]::White)

    $fWk     = New-Object System.Drawing.Font('Segoe UI', [float](8 * $S), [System.Drawing.FontStyle]::Bold)
    $fDay    = New-Object System.Drawing.Font('Consolas', [float](10.5 * $S))
    $fDayBold= New-Object System.Drawing.Font('Consolas', [float](10.5 * $S), [System.Drawing.FontStyle]::Bold)
    $fShort  = New-Object System.Drawing.Font('Consolas', [float](10.5 * $S))
    $brushWk = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0x80, 0x80, 0x80))
    $brushDay= New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0x20, 0x20, 0x20))
    $brushOff= New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0xC0, 0x00, 0x00))
    $brushSh = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0xB2, 0x5E, 0x00))

    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Alignment = 'Center'
    $fmt.LineAlignment = 'Center'

    # строка дней недели
    for ($c = 0; $c -lt 7; $c++) {
        $rect = New-Object System.Drawing.RectangleF(($c * $cw), 0, $cw, $wkH)
        $col = if ($c -ge 5) { [System.Drawing.Color]::FromArgb(0xC0,0x00,0x00) } else { [System.Drawing.Color]::FromArgb(0x70,0x70,0x70) }
        $b = New-Object System.Drawing.SolidBrush($col)
        $g.DrawString($script:CcWeekdays[$c], $fWk, $b, $rect, $fmt)
        $b.Dispose()
    }

    # первый день месяца относительно понедельника
    $first = New-Object System.DateTime($Year, $Month, 1)
    $startCol = (([int]$first.DayOfWeek + 6) % 7)
    $daysInMonth = [DateTime]::DaysInMonth($Year, $Month)
    $offsetInYear = 0
    for ($m = 1; $m -lt $Month; $m++) { $offsetInYear += [DateTime]::DaysInMonth($Year, $m) }

    for ($d = 1; $d -le $daysInMonth; $d++) {
        $pos = $startCol + ($d - 1)
        $row = [int][Math]::Floor($pos / 7)
        $col = $pos % 7
        $x = $col * $cw
        $y = $wkH + $row * $ch
        $rect = New-Object System.Drawing.Rectangle($x, $y, $cw, $ch)
        $rectF = New-Object System.Drawing.RectangleF($x, $y, $cw, $ch)

        $code = $Codes[$offsetInYear + ($d - 1)]
        $isWeekendCol = ($col -ge 5)
        $isToday = ($Year -eq $today.Year -and $Month -eq $today.Month -and $d -eq $today.Day)

        $bg = [System.Drawing.Color]::White
        $fg = $brushDay
        switch ($code) {
            '1' { $bg = [System.Drawing.Color]::FromArgb(0xFD,0xE7,0xE7); $fg = $brushOff }
            '2' { $bg = [System.Drawing.Color]::FromArgb(0xFF,0xF2,0xD9); $fg = $brushSh }
        }
        if ($code -eq '0' -and $isWeekendCol) { $fg = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0xC0,0x00,0x00)) }

        $bgBrush = New-Object System.Drawing.SolidBrush($bg)
        $g.FillRectangle($bgBrush, $rect)
        $bgBrush.Dispose()

        # сетка ячейки
        $penGrid = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0xEC,0xEC,0xEC))
        $g.DrawRectangle($penGrid, $rect)
        $penGrid.Dispose()

        $label = "$d"
        $font = $fDay
        if ($code -eq '2') { $label = "$d*" }
        if ($isToday) { $font = $fDayBold }

        if ($isToday) {
            $penT = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0x00,0x78,0xD7), [float](2 * $S))
            $inner = New-Object System.Drawing.Rectangle(($x+1), ($y+1), ($cw-2), ($ch-2))
            $g.DrawRectangle($penT, $inner)
            $penT.Dispose()
        }
        $g.DrawString($label, $font, $fg, $rectF, $fmt)
        if ($code -eq '0' -and $isWeekendCol) { $fg.Dispose() }
    }

    $g.Dispose()
    $Ui.Grid.BackgroundImage = $bmp
    $Ui.Grid.Tag = @{ Year = $Year; Month = $Month; Codes = $Codes; StartCol = $startCol; Days = $daysInMonth }
    $Ui.Grid.Invalidate()
}

function Build-DayTooltip {
    param([int]$Year, [int]$Month, [int]$Day, [string]$Code)
    $dt = New-Object System.DateTime($Year, $Month, $Day)
    $wd = $script:CcWeekdaysFull[([int]$dt.DayOfWeek + 6) % 7]
    $dateStr = "{0} {1} {2}, {3}" -f $Day, ($script:CcMonthFull[$Month-1].ToLower()), $Year, $wd
    switch ($Code) {
        '0' { return "$dateStr - рабочий день" }
        '1' { return "$dateStr - нерабочий (выходной/праздник)" }
        '2' { return "$dateStr - сокращённый день, -1 час" }
        default { return "$dateStr" }
    }
}
