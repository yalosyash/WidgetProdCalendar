# lib\Calendar.ps1
# Современная отрисовка: шапка с навигацией по кварталам, карточки месяцев
# (3 в ряд, при развороте — 2 ряда = 2 квартала), нижняя строка со сводкой.
# Данные НЕ подтягивает сам: получает готовые месяцы (коды + посчитанную
# статистику) и рисует. Поэтому не зависит от Data.ps1.

$script:CcMonthFull = @('Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь')
$script:CcMonthShort = @('янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек')
$script:CcWeekdays = @('пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс')
$script:CcWeekdaysFull = @('понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье')
$script:CcQuarter = @('I', 'II', 'III', 'IV')
$script:CultureRu = [System.Globalization.CultureInfo]::GetCultureInfo('ru-RU')

function Fmt1 { param([double]$v) return $v.ToString('0.0', $script:CultureRu) }

function New-RoundPath {
    param([double]$X, [double]$Y, [double]$W, [double]$H, [double]$R)
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    if ($R -le 0 -or $W -le 0 -or $H -le 0) {
        $p.AddRectangle((New-Object System.Drawing.RectangleF($X, $Y, [Math]::Max($W, 1), [Math]::Max($H, 1))))
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

function Set-TooltipWindowStyles {
    # Окно подсказки не должно забирать активацию: иначе показ подсказки
    # вызывает Deactivate у основного окна и виджет сам себя закрывает.
    # WS_EX_NOACTIVATE — не активируется, WS_EX_TOOLWINDOW — нет в Alt+Tab,
    # WS_EX_TRANSPARENT — не перехватывает клики.
    #
    # Важно: WinForms-овский Form.Show() активирует окно даже с WS_EX_NOACTIVATE,
    # поэтому подсказку показываем и прячем нативно (см. Show/HideNativeTip).
    param([Parameter(Mandatory)]$Form)
    if ($null -eq $Form -or $Form.IsDisposed -or -not $Form.IsHandleCreated) { return }
    if (-not ('PcNative.TipWin' -as [type])) {
        Add-Type -Namespace PcNative -Name TipWin -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
[DllImport("user32.dll", SetLastError = true)] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
[DllImport("user32.dll", SetLastError = true)] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
'@
    }
    $GWL_EXSTYLE = -20
    $WS_EX_NOACTIVATE = 0x08000000
    $WS_EX_TOOLWINDOW = 0x00000080
    $WS_EX_TRANSPARENT = 0x00000020
    $ex = [PcNative.TipWin]::GetWindowLong($Form.Handle, $GWL_EXSTYLE)
    [void][PcNative.TipWin]::SetWindowLong($Form.Handle, $GWL_EXSTYLE,
        ($ex -bor $WS_EX_NOACTIVATE -bor $WS_EX_TOOLWINDOW -bor $WS_EX_TRANSPARENT))
}

function Show-NativeTipWindow {
    # Показ/перемещение окна подсказки без активации: SetWindowPos с SWP_NOACTIVATE.
    # WinForms-овский Show() тут не годится — он активирует окно и виджет закрывается.
    param([Parameter(Mandatory)]$Form, [Parameter(Mandatory)][int]$X, [Parameter(Mandatory)][int]$Y)
    if ($null -eq $Form -or $Form.IsDisposed) { return }
    [void]$Form.Handle
    $HWND_TOPMOST = [IntPtr](-1)
    $SWP_NOACTIVATE = 0x0010
    $SWP_SHOWWINDOW = 0x0040
    [void][PcNative.TipWin]::SetWindowPos($Form.Handle, $HWND_TOPMOST, $X, $Y, $Form.Width, $Form.Height,
        ($SWP_NOACTIVATE -bor $SWP_SHOWWINDOW))
}

function Hide-NativeTipWindow {
    param([Parameter(Mandatory)]$Form)
    if ($null -eq $Form -or $Form.IsDisposed -or -not $Form.IsHandleCreated) { return }
    [void][PcNative.TipWin]::ShowWindow($Form.Handle, 0)   # SW_HIDE
}

function Set-RoundFormRegion {
    # Скругление углов произвольного окна: безрамочная форма на Windows 10 сама
    # углы не скругляет, поэтому вырезаем регион через CreateRoundRectRgn.
    param(
        [Parameter(Mandatory)]$Form,
        [int]$Radius = 12
    )
    if ($null -eq $Form -or $Form.IsDisposed) { return }
    if (-not ('PcNative.Gdi32' -as [type])) {
        Add-Type -Namespace PcNative -Name Gdi32 -MemberDefinition @'
[DllImport("gdi32.dll")] public static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
[DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr hObject);
'@
    }
    $w = $Form.Width; $h = $Form.Height
    if ($w -le 0 -or $h -le 0) { return }
    $d = [int][Math]::Max(2, $Radius * 2)
    $rgn = [PcNative.Gdi32]::CreateRoundRectRgn(0, 0, $w + 1, $h + 1, $d, $d)
    if ($rgn -eq [IntPtr]::Zero) { return }
    try {
        $old = $Form.Region
        $Form.Region = [System.Drawing.Region]::FromHrgn($rgn)
        if ($null -ne $old) { $old.Dispose() }
    }
    finally {
        [void][PcNative.Gdi32]::DeleteObject($rgn)
    }
}

function Set-CalendarFormRegion {
    param($Ui)
    $radius = if ($Ui.ContainsKey('WindowRadius')) { [int]$Ui.WindowRadius } else { 12 }
    Set-RoundFormRegion -Form $Ui.Surface -Radius $radius
}

function Set-CalendarRedraw {
    # Включает/выключает перерисовку окна (WM_SETREDRAW). На время смены размера
    # перерисовку гасим: иначе Windows успевает показать старый фон, растянутый
    # на новый размер, и разворот выглядит как мерцание.
    param(
        [Parameter(Mandatory)]$Form,
        [Parameter(Mandatory)][bool]$Enable
    )
    if ($null -eq $Form -or $Form.IsDisposed -or -not $Form.IsHandleCreated) { return }
    if (-not ('PcNative.Redraw' -as [type])) {
        Add-Type -Namespace PcNative -Name Redraw -MemberDefinition @'
[DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
'@
    }
    $on = if ($Enable) { [IntPtr]1 } else { [IntPtr]0 }
    [void][PcNative.Redraw]::SendMessage($Form.Handle, 0x000B, $on, [IntPtr]::Zero)
    if ($Enable) { $Form.Invalidate(); $Form.Update() }
}

function New-CalendarUi {
    param([double]$Scale = 1.0)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $S = $Scale
    function Scl([double]$v) { [int][Math]::Round($v * $S) }

    $ui = @{ Scale = $S }
    $ui.CellW = Scl 32; $ui.CellH = Scl 27
    $ui.CardPadX = Scl 12; $ui.CardPadY = Scl 10
    $ui.TitleH = Scl 26; $ui.WkH = Scl 20
    $ui.GridW = $ui.CellW * 7
    $ui.CardW = $ui.GridW + 2 * $ui.CardPadX
    $ui.GridH = $ui.CellH * 6
    $ui.StatH = Scl 30
    $ui.CardH = $ui.CardPadY + $ui.TitleH + $ui.WkH + $ui.GridH + $ui.StatH + $ui.CardPadY
    $ui.GapX = Scl 12; $ui.GapY = Scl 12
    $ui.Pad = Scl 16; $ui.Spacer = Scl 10
    $ui.HeaderH = Scl 50; $ui.FooterH = Scl 18
    $ui.Cols = 3
    $ui.ButtonSize = Scl 34; $ui.BtnGap = Scl 8; $ui.Radius = Scl 10; $ui.CellRadius = Scl 7
    # Скругление углов окна — чуть крупнее карточек, чтобы панель читалась цельно.
    $ui.WindowRadius = Scl 12
    $ui.FormWidth = $ui.Pad * 2 + $ui.Cols * $ui.CardW + ($ui.Cols - 1) * $ui.GapX
    $ui.Rows = 1

    # палитра
    $ui.C_Bg = [System.Drawing.Color]::FromArgb(0xF6, 0xF7, 0xFA)
    $ui.C_CardBg = [System.Drawing.Color]::White
    $ui.C_CardBorder = [System.Drawing.Color]::FromArgb(0xE6, 0xE8, 0xEE)
    $ui.C_Shadow = [System.Drawing.Color]::FromArgb(0xEC, 0xEE, 0xF3)
    $ui.C_Text = [System.Drawing.Color]::FromArgb(0x1F, 0x24, 0x30)
    $ui.C_Muted = [System.Drawing.Color]::FromArgb(0xA2, 0xA8, 0xB4)
    $ui.C_Red = [System.Drawing.Color]::FromArgb(0xE5, 0x48, 0x4D)
    $ui.C_RedBg = [System.Drawing.Color]::FromArgb(0xFE, 0xF1, 0xF1)
    $ui.C_Amber = [System.Drawing.Color]::FromArgb(0xD9, 0x82, 0x0B)
    $ui.C_AmberBg = [System.Drawing.Color]::FromArgb(0xFF, 0xF7, 0xEA)
    $ui.C_Accent = [System.Drawing.Color]::FromArgb(0x4F, 0x46, 0xE5)
    $ui.C_AccentSoft = [System.Drawing.Color]::FromArgb(0xEE, 0xF0, 0xFE)
    $ui.C_Footer = [System.Drawing.Color]::FromArgb(0x7A, 0x81, 0x94)
    $ui.C_Divider = [System.Drawing.Color]::FromArgb(0xEF, 0xF1, 0xF5)

    # палитра подсказки — тёмная карточка поверх светлого фона
    $ui.C_TipBg = [System.Drawing.Color]::FromArgb(0x24, 0x28, 0x33)
    $ui.C_TipBorder = [System.Drawing.Color]::FromArgb(0x3A, 0x3F, 0x4D)
    $ui.C_TipText = [System.Drawing.Color]::FromArgb(0xF3, 0xF5, 0xFA)
    $ui.C_TipMuted = [System.Drawing.Color]::FromArgb(0x9C, 0xA4, 0xB8)
    $ui.C_TipAccent = [System.Drawing.Color]::FromArgb(0x9E, 0x9A, 0xFF)

    # шрифты
    $ui.FTitle = New-Object System.Drawing.Font('Segoe UI', [float](13.5 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FMonth = New-Object System.Drawing.Font('Segoe UI', [float](11.5 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FWk = New-Object System.Drawing.Font('Segoe UI', [float](8.5 * $S))
    $ui.FDay = New-Object System.Drawing.Font('Segoe UI', [float](10 * $S))
    $ui.FDayBold = New-Object System.Drawing.Font('Segoe UI', [float](10 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FStat = New-Object System.Drawing.Font('Segoe UI', [float](7.8 * $S))
    $ui.FGlyph = New-Object System.Drawing.Font('Segoe UI', [float](12 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FSmall = New-Object System.Drawing.Font('Segoe UI', [float](8 * $S))

    # шрифты подсказки
    $ui.FTipTitle = New-Object System.Drawing.Font('Segoe UI', [float](9.5 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FTipBody = New-Object System.Drawing.Font('Segoe UI', [float](9 * $S))
    $ui.FTipTag = New-Object System.Drawing.Font('Segoe UI', [float](8 * $S), [System.Drawing.FontStyle]::Bold)

    # размеры подсказки
    $ui.TipPadX = Scl 12; $ui.TipPadY = Scl 9
    $ui.TipRadius = Scl 8; $ui.TipGap = Scl 6

    # форма
    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'None'
    $form.ShowInTaskbar = $false
    $form.TopMost = $true
    $form.StartPosition = 'Manual'
    $form.KeyPreview = $true
    $form.BackColor = $ui.C_Bg
    # Двойная буферизация: кадр готовится целиком и отдаётся экрану одним куском.
    # Без неё видно, как перерисовывается содержимое при наведении и смене квартала.
    $dbProp = $form.GetType().GetProperty('DoubleBuffered',
        [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
    if ($dbProp) { $dbProp.SetValue($form, $true, $null) }
    # BackgroundImage не растягиваем: при несовпадении размеров растяжение давало бы
    # заметный «плывущий» кадр в момент разворота.
    $form.BackgroundImageLayout = 'None'

    $ui.Surface = $form
    $ui.TipForm = New-Object System.Windows.Forms.Form
    $ui.TipForm.FormBorderStyle = 'None'
    $ui.TipForm.ShowInTaskbar = $false
    $ui.TipForm.TopMost = $true
    $ui.TipForm.StartPosition = 'Manual'
    $ui.TipForm.BackColor = $ui.C_TipBg
    $ui.TipForm.BackgroundImageLayout = 'None'
    $tipDb = $ui.TipForm.GetType().GetProperty('DoubleBuffered',
        [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
    if ($tipDb) { $tipDb.SetValue($ui.TipForm, $true, $null) }
    # Стили окна подсказки ставим ДО первого показа: если выставить их после Show,
    # первый показ успевает забрать активацию у виджета, и тот закрывается.
    [void]$ui.TipForm.Handle
    Set-TooltipWindowStyles -Form $ui.TipForm
    $ui.Bmp = $null
    $ui.TipBmp = $null
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
        [string]$HoverKey = '',
        [bool]$ShowArrows = $true
    )

    $S = $Ui.Scale
    $rows = [int][Math]::Ceiling($Months.Count / $Ui.Cols)
    if ($rows -lt 1) { $rows = 1 }
    $Ui.Rows = $rows

    $w = $Ui.Surface.Width
    $h = $Ui.Surface.Height
    $bmp = New-Object System.Drawing.Bitmap([Math]::Max($w, 1), [Math]::Max($h, 1))
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

    $brText = New-Object System.Drawing.SolidBrush($Ui.C_Text)
    $brMuted = New-Object System.Drawing.SolidBrush($Ui.C_Muted)
    $brRed = New-Object System.Drawing.SolidBrush($Ui.C_Red)
    $brAmber = New-Object System.Drawing.SolidBrush($Ui.C_Amber)
    $brAccent = New-Object System.Drawing.SolidBrush($Ui.C_Accent)
    $brWhite = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $brFoot = New-Object System.Drawing.SolidBrush($Ui.C_Footer)
    $hits = New-Object System.Collections.ArrayList

    # ---------- шапка ----------
    $btn = $Ui.ButtonSize
    $hy = $Ui.Pad + [int](($Ui.HeaderH - $btn) / 2)
    $expandKey = $(if ($rows -gt 1) { [char]0x25BC } else { [char]0x25B2 })
    $expRect = New-Object System.Drawing.Rectangle(($w - $Ui.Pad - $btn), $hy, $btn, $btn)
    $expandTip = if ($rows -gt 1) {
        @{ Title = 'Свернуть'; Meta = 'Показать один квартал'; Badge = ''; Tone = 'plain' }
    }
    else {
        @{ Title = 'Развернуть'; Meta = 'Показать два квартала'; Badge = ''; Tone = 'plain' }
    }
    [void]$hits.Add(@{ Kind = 'expand'; Rect = $expRect; Tip = $expandTip })
    $rp = New-RoundPath $expRect.X $expRect.Y $expRect.Width $expRect.Height $Ui.Radius
    if ($HoverKey -eq 'expand') {
        $hb = New-Object System.Drawing.SolidBrush($Ui.C_AccentSoft)
        $g.FillPath($hb, $rp); $hb.Dispose()
    }
    $g.DrawString([string]$expandKey, $Ui.FGlyph, $brAccent, (New-Object System.Drawing.RectangleF($expRect.X, $expRect.Y, $expRect.Width, $expRect.Height)), $sf)
    $rp.Dispose()

    $titleLeft = $Ui.Pad
    $titleRight = $expRect.X - $Ui.BtnGap
    if ($ShowArrows) {
        $prevRect = New-Object System.Drawing.Rectangle($Ui.Pad, $hy, $btn, $btn)
        $nextRect = New-Object System.Drawing.Rectangle(($expRect.X - $Ui.BtnGap - $btn), $hy, $btn, $btn)
        [void]$hits.Add(@{ Kind = 'prev'; Rect = $prevRect
                Tip = @{ Title = 'Предыдущий квартал'; Meta = 'Колесо вверх · ←'; Badge = ''; Tone = 'plain' } 
            })
        [void]$hits.Add(@{ Kind = 'next'; Rect = $nextRect
                Tip = @{ Title = 'Следующий квартал'; Meta = 'Колесо вниз · →'; Badge = ''; Tone = 'plain' } 
            })
        foreach ($b in @(
                @{ R = $prevRect; Key = 'prev'; G = [char]0x2039 },
                @{ R = $nextRect; Key = 'next'; G = [char]0x203A }
            )) {
            $rp = New-RoundPath $b.R.X $b.R.Y $b.R.Width $b.R.Height $Ui.Radius
            if ($HoverKey -eq $b.Key) {
                $hb = New-Object System.Drawing.SolidBrush($Ui.C_AccentSoft)
                $g.FillPath($hb, $rp); $hb.Dispose()
            }
            $g.DrawString([string]$b.G, $Ui.FGlyph, $brAccent, (New-Object System.Drawing.RectangleF($b.R.X, $b.R.Y, $b.R.Width, $b.R.Height)), $sf)
            $rp.Dispose()
        }
        $titleLeft = $prevRect.Right + $Ui.BtnGap
    }
    $titleRect = New-Object System.Drawing.RectangleF($titleLeft, $Ui.Pad, ($titleRight - $titleLeft), $Ui.HeaderH)
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

        # карточка целиком — как область подсказки (дни добавляются позже и перекрывают)
        [void]$hits.Add(@{
                Kind = 'card'
                Rect = (New-Object System.Drawing.Rectangle($cx, $cy, $Ui.CardW, $Ui.CardH))
                Tip  = (Build-MonthTip -Year $mo.Year -Month $mo.Month -Stat $mo.Stat)
                Text = ("{0} {1}" -f $script:CcMonthFull[$mo.Month - 1], $mo.Year)
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
        }
        else {
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
                }
                elseif ($code -eq '1') {
                    $tp = New-RoundPath ($x + 2) ($y + 2) ($Ui.CellW - 4) ($Ui.CellH - 4) $Ui.CellRadius
                    $tb = New-Object System.Drawing.SolidBrush($Ui.C_RedBg)
                    $g.FillPath($tb, $tp); $tb.Dispose(); $tp.Dispose()
                    $brush = $brRed
                }
                elseif ($code -eq '2') {
                    $tp = New-RoundPath ($x + 2) ($y + 2) ($Ui.CellW - 4) ($Ui.CellH - 4) $Ui.CellRadius
                    $tb = New-Object System.Drawing.SolidBrush($Ui.C_AmberBg)
                    $g.FillPath($tb, $tp); $tb.Dispose(); $tp.Dispose()
                    $brush = $brAmber
                    $label = "$d*"
                }

                $g.DrawString($label, $font, $brush, $rectF, $sf)
                $dayTip = Build-DayTip -Year $mo.Year -Month $mo.Month -Day $d -Code $code
                if ($isToday) { $dayTip.Badge = 'Сегодня' }
                [void]$hits.Add(@{
                        Kind = 'day'
                        Rect = (New-Object System.Drawing.Rectangle($x, $y, $Ui.CellW, $Ui.CellH))
                        Tip  = $dayTip
                        Text = ("{0} {1} {2}" -f $d, $script:CcMonthShort[$mo.Month - 1], $mo.Year)
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
            $g.DrawString($l2, $Ui.FStat, $brText, (New-Object System.Drawing.RectangleF($stx, ($sty + $half), $Ui.GridW, $half)), $sf)
        }
    }

    # ---------- футер ----------
    $fy = $h - $Ui.Pad - $Ui.FooterH
    $g.DrawString($FooterLeft, $Ui.FSmall, $brFoot, (New-Object System.Drawing.RectangleF($Ui.Pad, $fy, ($w / 2), $Ui.FooterH)), $sfL)
    $g.DrawString($FooterRight, $Ui.FSmall, $brFoot, (New-Object System.Drawing.RectangleF(($w / 2), $fy, ($w / 2 - $Ui.Pad), $Ui.FooterH)), $sfR)

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

function Build-DayTip {
    # Структурированная подсказка дня: рисуется своим окном, а не системным ToolTip.
    param([int]$Year, [int]$Month, [int]$Day, [string]$Code)
    $dt = New-Object System.DateTime($Year, $Month, $Day)
    $wd = $script:CcWeekdaysFull[([int]$dt.DayOfWeek + 6) % 7]
    $wd = $wd.Substring(0, 1).ToUpper() + $wd.Substring(1)
    $tip = @{
        Title = ("{0} {1} {2}" -f $Day, $script:CcMonthFull[$Month - 1].ToLower(), $Year)
        Meta  = $wd
        Tone  = 'plain'
        Badge = ''
    }
    switch ($Code) {
        '0' { $tip.Badge = 'Рабочий день'; $tip.Tone = 'work' }
        '1' { $tip.Badge = 'Нерабочий · выходной'; $tip.Tone = 'off' }
        '2' { $tip.Badge = 'Сокращённый · −1 час'; $tip.Tone = 'short' }
    }
    return $tip
}

function Build-MonthTip {
    param([int]$Year, [int]$Month, $Stat)
    $tip = @{
        Title = ("{0} {1}" -f $script:CcMonthFull[$Month - 1], $Year)
        Tone  = 'plain'
        Badge = ''
        Lines = @()
    }
    if ($null -eq $Stat) {
        $tip.Badge = 'загрузка…'
        return $tip
    }
    $tip.Badge = ("{0} календарных дней" -f $Stat.Calendar)
    $tip.Lines = @(
        ("Рабочих {0}  ·  выходных {1}" -f $Stat.Work, $Stat.Off),
        ("Норма: {0} / {1} / {2} ч" -f (Fmt1 $Stat.Hours40), (Fmt1 $Stat.Hours36), (Fmt1 $Stat.Hours24))
    )
    return $tip
}

function Get-TipTone {
    # Цвет текста и подложки «пилюли» в зависимости от типа дня.
    param($Ui, [string]$Tone)
    switch ($Tone) {
        'off' { return @{ Text = [System.Drawing.Color]::FromArgb(0xFF, 0x8A, 0x8D); Fill = [System.Drawing.Color]::FromArgb(48, 0xE5, 0x48, 0x4D) } }
        'short' { return @{ Text = [System.Drawing.Color]::FromArgb(0xFF, 0xC4, 0x6B); Fill = [System.Drawing.Color]::FromArgb(48, 0xD9, 0x82, 0x0B) } }
        'work' { return @{ Text = $Ui.C_TipAccent; Fill = [System.Drawing.Color]::FromArgb(48, 0x6C, 0x63, 0xFF) } }
        default { return @{ Text = $Ui.C_TipMuted; Fill = [System.Drawing.Color]::FromArgb(40, 0x9C, 0xA4, 0xB8) } }
    }
}

function New-CalendarTooltipBitmap {
    # Современная карточка подсказки: тёмная, со скруглением, пилюлей статуса
    # и акцентной полосой слева. Отрисовывается в bitmap и отдаётся окну подсказки.
    param([Parameter(Mandatory)]$Ui, [Parameter(Mandatory)]$Tip)

    $S = $Ui.Scale
    $title = [string]$Tip.Title
    $meta = if ($Tip.ContainsKey('Meta') -and $Tip.Meta) { [string]$Tip.Meta } else { '' }
    $badge = if ($Tip.ContainsKey('Badge') -and $Tip.Badge) { [string]$Tip.Badge } else { '' }
    $lines = @()
    if ($Tip.ContainsKey('Lines') -and $Tip.Lines) { $lines = @($Tip.Lines | ForEach-Object { [string]$_ }) }
    $tone = if ($Tip.ContainsKey('Tone') -and $Tip.Tone) { [string]$Tip.Tone } else { 'plain' }
    $col = Get-TipTone -Ui $Ui -Tone $tone

    # Измеряем текст временным полотном.
    $probe = New-Object System.Drawing.Bitmap(1, 1)
    $pg = [System.Drawing.Graphics]::FromImage($probe)
    $pg.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $szTitle = $pg.MeasureString($title, $Ui.FTipTitle)
    $szMeta = if ($meta) { $pg.MeasureString($meta, $Ui.FTipBody) } else { New-Object System.Drawing.SizeF(0, 0) }
    $szBadge = if ($badge) { $pg.MeasureString($badge, $Ui.FTipTag) } else { New-Object System.Drawing.SizeF(0, 0) }
    $wLines = 0.0; $hLines = 0.0
    foreach ($l in $lines) {
        # Имя $szLine, а не $s: переменные в PowerShell без учёта регистра, и $s
        # затёр бы масштаб $S, сломав все дальнейшие расчёты размеров.
        $szLine = $pg.MeasureString($l, $Ui.FTipBody)
        $wLines = [Math]::Max($wLines, $szLine.Width)
        $hLines += $szLine.Height
    }
    $pg.Dispose(); $probe.Dispose()

    $padX = $Ui.TipPadX; $padY = $Ui.TipPadY
    $barW = [int][Math]::Max(3, $S * 3)
    $badgeH = if ($badge) { [int][Math]::Ceiling($szBadge.Height + $S * 5) } else { 0 }
    $badgeW = if ($badge) { [int][Math]::Ceiling($szBadge.Width + $S * 14) } else { 0 }
    $row2H = [int][Math]::Ceiling([Math]::Max([double]$badgeH, [double]$szMeta.Height))
    $gapRow = if ($badge -or $meta) { [int][Math]::Ceiling($S * 5) } else { 0 }
    $gapLines = if ($lines.Count -gt 0) { [int][Math]::Ceiling($S * 5) } else { 0 }

    $contentW = [Math]::Max([double]$szTitle.Width, $wLines)
    if ($badge -or $meta) {
        $row2W = [double]$badgeW + $szMeta.Width
        if ($badge -and $meta) { $row2W += $S * 8 }
        $contentW = [Math]::Max($contentW, $row2W)
    }
    $w = [int][Math]::Ceiling($contentW) + $padX * 2 + $barW
    $h = [int][Math]::Ceiling($szTitle.Height) + $gapRow + $row2H + $gapLines + [int][Math]::Ceiling($hLines) + $padY * 2
    $w = [Math]::Max($w, [int]($S * 104))
    $h = [Math]::Max($h, [int]($S * 34))

    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'ClearTypeGridFit'
    $g.Clear($Ui.C_TipBg)

    # акцентная полоса слева
    $barPath = New-RoundPath 0 ([int]($S * 4)) $barW ($h - [int]($S * 8)) ([int]($S * 2))
    $bb = New-Object System.Drawing.SolidBrush($col.Text)
    $g.FillPath($bb, $barPath); $bb.Dispose(); $barPath.Dispose()

    # рамка карточки
    $penB = New-Object System.Drawing.Pen($Ui.C_TipBorder)
    $bp = New-RoundPath 0.5 0.5 ($w - 1) ($h - 1) ($Ui.TipRadius - 1)
    $g.DrawPath($penB, $bp); $penB.Dispose(); $bp.Dispose()

    $sfC = New-Object System.Drawing.StringFormat
    $sfC.Alignment = 'Center'; $sfC.LineAlignment = 'Center'
    $sfL = New-Object System.Drawing.StringFormat
    $sfL.Alignment = 'Near'; $sfL.LineAlignment = 'Center'

    $x = $padX + $barW
    $y = $padY
    $brTitle = New-Object System.Drawing.SolidBrush($Ui.C_TipText)
    $g.DrawString($title, $Ui.FTipTitle, $brTitle, [float]$x, [float]$y)
    $brTitle.Dispose()
    $y += [int][Math]::Ceiling($szTitle.Height) + $gapRow

    if ($row2H -gt 0) {
        $rx = $x
        if ($badge) {
            $pillH = $badgeH
            $pillY = $y + [int](($row2H - $pillH) / 2)
            $pp = New-RoundPath $rx $pillY $badgeW $pillH ([int]($pillH / 2))
            $pb = New-Object System.Drawing.SolidBrush($col.Fill)
            $g.FillPath($pb, $pp); $pb.Dispose(); $pp.Dispose()
            $sb = New-Object System.Drawing.SolidBrush($col.Text)
            $g.DrawString($badge, $Ui.FTipTag, $sb,
                (New-Object System.Drawing.RectangleF($rx, $pillY, $badgeW, $pillH)), $sfC)
            $sb.Dispose()
            $rx += $badgeW + [int]($S * 8)
        }
        if ($meta) {
            $mb = New-Object System.Drawing.SolidBrush($Ui.C_TipMuted)
            $g.DrawString($meta, $Ui.FTipBody, $mb,
                (New-Object System.Drawing.RectangleF($rx, $y, ([Math]::Max($w - $rx - $padX, 1)), $row2H)), $sfL)
            $mb.Dispose()
        }
        $y += $row2H + $gapLines
    }

    if ($lines.Count -gt 0) {
        $lb = New-Object System.Drawing.SolidBrush($Ui.C_TipMuted)
        $sb2 = New-Object System.Drawing.SolidBrush($Ui.C_TipText)
        $i = 0
        foreach ($l in $lines) {
            $br = if ($i -eq $lines.Count - 1) { $sb2 } else { $lb }
            $g.DrawString($l, $Ui.FTipBody, $br, [float]$x, [float]$y)
            $y += [int][Math]::Ceiling($Ui.FTipBody.GetHeight($g))
            $i++
        }
        $lb.Dispose(); $sb2.Dispose()
    }

    $sfC.Dispose(); $sfL.Dispose()
    $g.Dispose()
    return $bmp
}

function Show-CalendarTooltip {
    # Подсказка ставится у курсора, как в системных тултипах: для элементов
    # любого размера (день, карточка, кнопка) это предсказуемее, чем привязка
    # к границам — у высокой карточки месяца «под элементом» ушло бы за экран.
    param(
        [Parameter(Mandatory)]$Ui,
        [Parameter(Mandatory)]$Tip,
        [Parameter(Mandatory)]$Surface
    )
    if ($null -eq $Tip -or -not $Tip.ContainsKey('Title') -or -not $Tip.Title) { Hide-CalendarTooltip -Ui $Ui; return }

    $bmp = New-CalendarTooltipBitmap -Ui $Ui -Tip $Tip
    $tf = $Ui.TipForm
    if ($Ui.TipBmp) { $Ui.TipBmp.Dispose() }
    $Ui.TipBmp = $bmp

    $tf.Size = New-Object System.Drawing.Size($bmp.Width, $bmp.Height)
    $tf.BackgroundImage = $bmp
    Set-RoundFormRegion -Form $tf -Radius $Ui.TipRadius

    # Точка отсчёта — курсор: смещаем карточку вниз-вправо, как системную.
    $cur = [System.Windows.Forms.Cursor]::Position
    $wa = [System.Windows.Forms.Screen]::FromControl($Surface).WorkingArea
    $px = $cur.X + $Ui.TipGap
    $py = $cur.Y + $Ui.TipGap
    # У правого/нижнего края экрана разворачиваем карточку в другую сторону.
    if (($px + $bmp.Width) -gt ($wa.Right - 2)) { $px = $cur.X - $bmp.Width - $Ui.TipGap }
    if (($py + $bmp.Height) -gt ($wa.Bottom - 2)) { $py = $cur.Y - $bmp.Height - $Ui.TipGap }
    $px = [Math]::Max($wa.Left + 2, [Math]::Min($px, $wa.Right - $bmp.Width - 2))
    $py = [Math]::Max($wa.Top + 2, [Math]::Min($py, $wa.Bottom - $bmp.Height - 2))
    # Показываем нативно, без активации — иначе виджет получит Deactivate и закроется.
    Show-NativeTipWindow -Form $tf -X $px -Y $py
}

function Test-CalendarTooltipVisible {
    # Нативный показ не обновляет свойство WinForms .Visible, поэтому спрашиваем
    # у системы, видно ли окно подсказки на самом деле.
    param($Ui)
    if ($null -eq $Ui -or -not $Ui.ContainsKey('TipForm')) { return $false }
    $tf = $Ui.TipForm
    if ($null -eq $tf -or $tf.IsDisposed -or -not $tf.IsHandleCreated) { return $false }
    return [PcNative.TipWin]::IsWindowVisible($tf.Handle)
}

function Hide-CalendarTooltip {
    param($Ui)
    if ($null -eq $Ui -or -not $Ui.ContainsKey('TipForm')) { return }
    Hide-NativeTipWindow -Form $Ui.TipForm
}
