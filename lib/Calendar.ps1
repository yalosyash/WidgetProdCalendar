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

function Fmt-Plural {
    # Склонение существительного при числе: «31 календарный день».
    param([int]$N, [string]$One, [string]$Few, [string]$Many)
    $n = [Math]::Abs($N) % 100
    if ($n -ge 11 -and $n -le 14) { return "$N $Many" }
    switch ($n % 10) {
        1 { return "$N $One" }
        2 { return "$N $Few" }
        3 { return "$N $Few" }
        4 { return "$N $Few" }
        default { return "$N $Many" }
    }
}

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

function Add-SoftShadow {
    # Мягкая тень скруглённого прямоугольника на переданном полотне: вокруг
    # прямоугольника рисуются концентрические скруглённые фигуры с малой альфой —
    # у карточки плотнее, к краю сходит на нет. Пик намеренно низкий, затухание
    # нелинейное: тень читается как лёгкая подсветка, а не как тёмная окантовка.
    # $X/$Y — левый верхний угол «карточки» на полотне; поля тени $Margin должны
    # уже помещаться в bitmap, иначе тень обрежется краем.
    param(
        [Parameter(Mandatory)]$Graphics,
        [Parameter(Mandatory)][int]$X,
        [Parameter(Mandatory)][int]$Y,
        [Parameter(Mandatory)][int]$W,
        [Parameter(Mandatory)][int]$H,
        [Parameter(Mandatory)][int]$Margin,
        [int]$Radius = 12,
        [int]$OffsetY = 0,
        [System.Drawing.Color]$Color = ([System.Drawing.Color]::FromArgb(0x1A, 0x20, 0x2E)),
        [double]$Peak = 13.0
    )
    for ($i = $Margin; $i -ge 0; $i--) {
        $alpha = [int]($Peak * [Math]::Pow(1.0 - ($i / [double]($Margin + 1)), 2.4))
        if ($alpha -le 0) { continue }
        $sc = [System.Drawing.Color]::FromArgb($alpha, $Color.R, $Color.G, $Color.B)
        $sp = New-RoundPath ($X - $i) ($Y - $i + $OffsetY) ($W + $i * 2) ($H + $i * 2) ($Radius + $i)
        $sb = New-Object System.Drawing.SolidBrush($sc)
        $Graphics.FillPath($sb, $sp); $sb.Dispose(); $sp.Dispose()
    }
}

function Set-LayeredWindowStyles {
    # Окно не должно забирать активацию; у подсказки это критично: показ подсказки
    # иначе вызывает Deactivate у основного окна и виджет сам себя закрывает.
    # WS_EX_NOACTIVATE — не активируется, WS_EX_TOOLWINDOW — нет в Alt+Tab,
    # WS_EX_TRANSPARENT — не перехватывает клики, WS_EX_LAYERED — рисуется
    # по пикселям с альфой (нужна настоящая прозрачность для тени).
    #
    # Флаги общие для подсказки и окна-тени виджета: у тени важно, чтобы клики
    # сквозь неё попадали в виджет, а сама она не появлялась в Alt+Tab.
    #
    # Важно: WinForms-овский Form.Show() активирует окно даже с WS_EX_NOACTIVATE,
    # поэтому подсказку показываем и прячем нативно (см. UpdateLayeredBitmap).
    param([Parameter(Mandatory)]$Form)
    if ($null -eq $Form -or $Form.IsDisposed -or -not $Form.IsHandleCreated) { return }
    if (-not ('PcNative.Layered' -as [type])) {
        Add-Type -Namespace PcNative -Name Layered -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
[DllImport("user32.dll", SetLastError = true)] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
[DllImport("user32.dll", SetLastError = true)] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
[DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hWnd);
[DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
[DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleDC(IntPtr hDC);
[DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr hDC);
[DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);
[DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr hObject);
[DllImport("user32.dll", SetLastError = true)] static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, ref POINT pptDst, ref SIZE psize, IntPtr hdcSrc, ref POINT pptSrc, int crKey, ref BLENDFUNCTION pblend, int dwFlags);

[StructLayout(LayoutKind.Sequential)] public struct POINT { public int x; public int y; public POINT(int X, int Y) { x = X; y = Y; } }
[StructLayout(LayoutKind.Sequential)] public struct SIZE { public int cx; public int cy; public SIZE(int X, int Y) { cx = X; cy = Y; } }
[StructLayout(LayoutKind.Sequential)] public struct BLENDFUNCTION { public byte BlendOp; public byte BlendFlags; public byte SourceConstantAlpha; public byte AlphaFormat; }

// Отдаёт окну готовый bitmap по пикселям (ULW_ALPHA): только так у подсказки
// получается настоящая полупрозрачная тень, а не прямоугольная подложка.
// HBITMAP создан вызывающим через Bitmap.GetHbitmap и здесь же удаляется.
public static bool UpdateLayeredBitmap(IntPtr hwnd, IntPtr hBitmap, int x, int y, int w, int h) {
    IntPtr screenDc = GetDC(IntPtr.Zero);
    IntPtr memDc = CreateCompatibleDC(screenDc);
    IntPtr oldBmp = SelectObject(memDc, hBitmap);
    bool ok = false;
    try {
        POINT dst = new POINT(x, y);
        POINT src = new POINT(0, 0);
        SIZE size = new SIZE(w, h);
        BLENDFUNCTION blend = new BLENDFUNCTION();
        blend.BlendOp = 0; blend.BlendFlags = 0;
        blend.SourceConstantAlpha = 255; blend.AlphaFormat = 1;   // AC_SRC_ALPHA
        ok = UpdateLayeredWindow(hwnd, screenDc, ref dst, ref size, memDc, ref src, 0, ref blend, 2); // ULW_ALPHA
    }
    finally {
        SelectObject(memDc, oldBmp);
        DeleteObject(hBitmap);
        DeleteDC(memDc);
        ReleaseDC(IntPtr.Zero, screenDc);
    }
    return ok;
}
'@
    }
    $GWL_EXSTYLE = -20
    $WS_EX_NOACTIVATE = 0x08000000
    $WS_EX_TOOLWINDOW = 0x00000080
    $WS_EX_TRANSPARENT = 0x00000020
    $WS_EX_LAYERED = 0x00080000
    $ex = [PcNative.Layered]::GetWindowLong($Form.Handle, $GWL_EXSTYLE)
    [void][PcNative.Layered]::SetWindowLong($Form.Handle, $GWL_EXSTYLE,
        ($ex -bor $WS_EX_NOACTIVATE -bor $WS_EX_TOOLWINDOW -bor $WS_EX_TRANSPARENT -bor $WS_EX_LAYERED))
}

function Show-LayeredWindow {
    # Показ окна без активации: содержимое отдаём по пикселям через
    # UpdateLayeredWindow (окно помечено WS_EX_LAYERED), поэтому тень вокруг
    # карточки получается полупрозрачной, а не прямоугольной подложкой.
    # WinForms-овский Show() тут не годится — он активирует окно и виджет закрывается.
    param(
        [Parameter(Mandatory)]$Form,
        [Parameter(Mandatory)]$Bitmap,
        [Parameter(Mandatory)][int]$X,
        [Parameter(Mandatory)][int]$Y,
        # $BelowForm — окно, ПОД которым держать это (тень виджета: она должна
        # лежать прямо под виджетом, оставаясь в полосе topmost, поэтому ставим
        # её не в самый низ Z-порядка, а сразу за окном виджета).
        $BelowForm = $null
    )
    if ($null -eq $Form -or $Form.IsDisposed -or $null -eq $Bitmap) { return }
    [void]$Form.Handle

    # GetHbitmap создаёт копию пикселей; владение передаём в UpdateLayeredBitmap,
    # которая удаляет её сама (иначе утечка GDI-объекта на каждый показ).
    $hBmp = $Bitmap.GetHbitmap([System.Drawing.Color]::FromArgb(0))
    [void][PcNative.Layered]::UpdateLayeredBitmap($Form.Handle, $hBmp, $X, $Y, $Bitmap.Width, $Bitmap.Height)

    # Позицию и размер задала ULW; здесь только показываем и ставим в Z-порядке,
    # не трогая ни геометрию, ни активацию.
    $SWP_NOSIZE = 0x0001; $SWP_NOMOVE = 0x0002; $SWP_NOACTIVATE = 0x0010; $SWP_SHOWWINDOW = 0x0040
    $after = if ($null -ne $BelowForm -and $BelowForm.IsHandleCreated) { $BelowForm.Handle } else { [IntPtr](-1) }
    [void][PcNative.Layered]::SetWindowPos($Form.Handle, $after, 0, 0, 0, 0,
        ($SWP_NOSIZE -bor $SWP_NOMOVE -bor $SWP_NOACTIVATE -bor $SWP_SHOWWINDOW))
}

function Hide-LayeredWindow {
    param([Parameter(Mandatory)]$Form)
    if ($null -eq $Form -or $Form.IsDisposed -or -not $Form.IsHandleCreated) { return }
    [void][PcNative.Layered]::ShowWindow($Form.Handle, 0)   # SW_HIDE
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
    $ui.BtnGap = Scl 8; $ui.Radius = Scl 10; $ui.CellRadius = Scl 7
    # Кнопки навигации — широкие «пилюли» с подписью: мелкий одиночный шеврон
    # не читался как кнопка. Ширину считаем по тексту при отрисовке, здесь —
    # высота, отступы и зазор между шевроном и подписью.
    $ui.ButtonH = Scl 34; $ui.BtnPadX = Scl 13; $ui.BtnGlyphGap = Scl 5
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

    # Палитра подсказки — та же светлая карточка, что и у месяцев: подсказка
    # читается как продолжение виджета, а не как отдельный тёмный элемент.
    # Рамка чуть темнее карточной: карточка висит на произвольном фоне, и
    # на светлом рабочем столе граница #E6E8EE почти не видна.
    $ui.C_TipBg = $ui.C_CardBg
    $ui.C_TipBorder = [System.Drawing.Color]::FromArgb(0xDC, 0xE0, 0xE9)
    $ui.C_TipShadow = [System.Drawing.Color]::FromArgb(0x1A, 0x20, 0x2E)
    $ui.C_TipText = $ui.C_Text
    $ui.C_TipMuted = $ui.C_Muted

    # шрифты
    $ui.FTitle = New-Object System.Drawing.Font('Segoe UI', [float](13.5 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FMonth = New-Object System.Drawing.Font('Segoe UI', [float](11.5 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FWk = New-Object System.Drawing.Font('Segoe UI', [float](8.5 * $S))
    $ui.FDay = New-Object System.Drawing.Font('Segoe UI', [float](10 * $S))
    $ui.FDayBold = New-Object System.Drawing.Font('Segoe UI', [float](10 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FStat = New-Object System.Drawing.Font('Segoe UI', [float](7.8 * $S))
    $ui.FGlyph = New-Object System.Drawing.Font('Segoe UI', [float](13 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FBtn = New-Object System.Drawing.Font('Segoe UI', [float](10 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FSmall = New-Object System.Drawing.Font('Segoe UI', [float](8 * $S))

    # Шрифты подсказки: заголовок — как у месяца, текст — как подписи в виджете.
    $ui.FTipTitle = New-Object System.Drawing.Font('Segoe UI', [float](11.5 * $S), [System.Drawing.FontStyle]::Bold)
    $ui.FTipBody = New-Object System.Drawing.Font('Segoe UI', [float](9 * $S))
    $ui.FTipTag = New-Object System.Drawing.Font('Segoe UI', [float](9 * $S), [System.Drawing.FontStyle]::Bold)

    # Размеры подсказки: радиус и отступы — как у карточки месяца.
    $ui.TipPadX = Scl 12; $ui.TipPadY = Scl 10
    $ui.TipRadius = $ui.Radius; $ui.TipGap = Scl 6
    # Тень: карточка «висит» над фоном. Размытие рисуем в самом bitmap, а вокруг
    # оставляем поля $ui.TipShadow — иначе тень упиралась бы в края окна.
    # Поля широкие, а пик альфы низкий: тень мягкая и не выбивается из светлого UI.
    $ui.TipShadow = Scl 13; $ui.TipShadowY = Scl 3

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

    # Тень виджета: отдельное слоистое окно ПОД основным. Основное окно остаётся
    # обычным (region + BackgroundImage), поэтому интерактивность и подсветка не
    # меняются, а полупрозрачную тень рисует соседнее окно (см. Update-WidgetShadow).
    $ui.ShadowMargin = Scl 14; $ui.ShadowY = Scl 4
    $ui.C_ShadowSoft = [System.Drawing.Color]::FromArgb(0x1A, 0x20, 0x2E)

    $ui.Surface = $form
    $ui.ShadowForm = New-Object System.Windows.Forms.Form
    $ui.ShadowForm.FormBorderStyle = 'None'
    $ui.ShadowForm.ShowInTaskbar = $false
    $ui.ShadowForm.TopMost = $true
    $ui.ShadowForm.StartPosition = 'Manual'
    $ui.ShadowForm.BackColor = $ui.C_Bg
    $ui.ShadowForm.BackgroundImageLayout = 'None'
    [void]$ui.ShadowForm.Handle
    Set-LayeredWindowStyles -Form $ui.ShadowForm

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
    Set-LayeredWindowStyles -Form $ui.TipForm
    $ui.Bmp = $null
    $ui.TipBmp = $null
    $ui.ShadowBmp = $null
    $ui.Hits = @()
    return $ui
}

function Update-WidgetShadow {
    # Пересобирает и показывает окно-тень под виджетом. Тень рисуется по контуру
    # скруглённого окна и живёт отдельным слоистым окном, поэтому не мешает
    # интерактивности: клики проходят сквозь (WS_EX_TRANSPARENT).
    param([Parameter(Mandatory)]$Ui, [Parameter(Mandatory)]$Form)
    if ($null -eq $Ui -or -not $Ui.ContainsKey('ShadowForm')) { return }
    $w = $Form.Width; $h = $Form.Height
    if ($w -le 0 -or $h -le 0) { return }

    $m = [int]$Ui.ShadowMargin; $radius = [int]$Ui.WindowRadius
    $bmp = New-Object System.Drawing.Bitmap(($w + $m * 2), ($h + $m * 2 + [int]$Ui.ShadowY), [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    Add-SoftShadow -Graphics $g -X $m -Y $m -W $w -H $h -Margin $m -Radius $radius `
        -OffsetY ([int]$Ui.ShadowY) -Color $Ui.C_ShadowSoft -Peak 16.0
    # Вырезаем область окна: под виджетом тень не нужна, иначе светлый фон
    # просвечивал бы через полупрозрачные пиксели и выглядел грязнее.
    $card = New-RoundPath $m $m $w $h $radius
    $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
    $clear = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Transparent)
    $g.FillPath($clear, $card); $clear.Dispose(); $card.Dispose()
    $g.Dispose()

    if ($Ui.ShadowBmp) { $Ui.ShadowBmp.Dispose() }
    $Ui.ShadowBmp = $bmp

    # Позиция — на поля тени влево/вверх от виджета, плюс вертикальный сдвиг тени.
    $px = $Form.Left - $m
    $py = $Form.Top - $m
    # -BelowForm: тень кладём сразу ПОД окно виджета, в ту же полосу topmost.
    Show-LayeredWindow -Form $Ui.ShadowForm -Bitmap $bmp -X $px -Y $py -BelowForm $Form
}

function Hide-WidgetShadow {
    param([Parameter(Mandatory)]$Ui)
    if ($null -eq $Ui -or -not $Ui.ContainsKey('ShadowForm')) { return }
    Hide-LayeredWindow -Form $Ui.ShadowForm
}

function Measure-HeaderButton {
    # Ширина кнопки-«пилюли»: шеврон + подпись + внутренние отступы.
    # Считаем по фактическому тексту, чтобы подпись «Квартал» нигде не обрезалась.
    param([Parameter(Mandatory)]$Ui, [Parameter(Mandatory)]$Graphics, [string]$Glyph = '', [string]$Label = '')
    $w = 0.0
    if ($Glyph) { $w += $Graphics.MeasureString($Glyph, $Ui.FGlyph).Width }
    if ($Glyph -and $Label) { $w += $Ui.BtnGlyphGap }
    if ($Label) { $w += $Graphics.MeasureString($Label, $Ui.FBtn).Width }
    return [int][Math]::Ceiling($w) + $Ui.BtnPadX * 2
}

function Draw-HeaderButton {
    # Кнопка шапки: скруглённая «пилюля» с постоянной подложкой (чтобы читалась
    # как кнопка ещё до наведения), крупным шевроном и подписью. Подпись — часть
    # кнопки, поэтому глазу сразу ясно, что нажимать и куда это ведёт.
    param(
        [Parameter(Mandatory)]$Ui,
        [Parameter(Mandatory)]$Graphics,
        [Parameter(Mandatory)][int]$X,
        [Parameter(Mandatory)][int]$Y,
        [Parameter(Mandatory)][int]$W,
        [Parameter(Mandatory)][int]$H,
        [string]$Glyph = '',
        [string]$Label = '',
        [bool]$GlyphLeft = $true,
        [bool]$Hover = $false,
        [Parameter(Mandatory)]$InkBrush
    )
    $path = New-RoundPath $X $Y $W $H $Ui.Radius
    # Подложка: на наведении акцентная, иначе — как у карточек, с рамкой.
    $fb = New-Object System.Drawing.SolidBrush($(if ($Hover) { $Ui.C_AccentSoft } else { $Ui.C_CardBg }))
    $Graphics.FillPath($fb, $path); $fb.Dispose()
    $pen = New-Object System.Drawing.Pen($(if ($Hover) { $Ui.C_AccentSoft } else { $Ui.C_CardBorder }))
    $Graphics.DrawPath($pen, $path); $pen.Dispose(); $path.Dispose()

    $sfMid = New-Object System.Drawing.StringFormat
    $sfMid.Alignment = 'Center'; $sfMid.LineAlignment = 'Center'

    $parts = @()
    if ($Glyph -and $Label) {
        $parts = if ($GlyphLeft) { @(@{ T = $Glyph; F = $Ui.FGlyph }, @{ T = $Label; F = $Ui.FBtn }) }
        else { @(@{ T = $Label; F = $Ui.FBtn }, @{ T = $Glyph; F = $Ui.FGlyph }) }
    }
    elseif ($Glyph) { $parts = @(@{ T = $Glyph; F = $Ui.FGlyph }) }
    elseif ($Label) { $parts = @(@{ T = $Label; F = $Ui.FBtn }) }
    if ($parts.Count -eq 0) { return }

    # Ширины считаем теми же шрифтами, что и при измерении, и центрируем группу.
    $sizes = @($parts | ForEach-Object { $Graphics.MeasureString($_.T, $_.F) })
    $groupW = 0.0
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $groupW += $sizes[$i].Width
        if ($i -lt $sizes.Count - 1 -and $parts[$i].F -ne $parts[$i + 1].F) { $groupW += $Ui.BtnGlyphGap }
    }
    $cx = $X + ($W - $groupW) / 2.0
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $Graphics.DrawString($parts[$i].T, $parts[$i].F, $InkBrush,
            (New-Object System.Drawing.RectangleF([float]$cx, [float]$Y, [float]($sizes[$i].Width + 2), [float]$H)), $sfMid)
        $cx += $sizes[$i].Width
        if ($i -lt $parts.Count - 1 -and $parts[$i].F -ne $parts[$i + 1].F) { $cx += $Ui.BtnGlyphGap }
    }
    $sfMid.Dispose()
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
    # Кнопки навигации — широкие «пилюли» с подписью, кнопка разворота — справа.
    # Ширины считаем заранее: от них зависит и область заголовка.
    $btnH = $Ui.ButtonH
    $hy = $Ui.Pad + [int](($Ui.HeaderH - $btnH) / 2)
    $expandKey = $(if ($rows -gt 1) { [char]0x25BC } else { [char]0x25B2 })
    $expandLabel = $(if ($rows -gt 1) { 'Свернуть' } else { 'Развернуть' })
    $expW = Measure-HeaderButton -Ui $Ui -Graphics $g -Glyph ([string]$expandKey) -Label $expandLabel
    $expRect = New-Object System.Drawing.Rectangle(($w - $Ui.Pad - $expW), $hy, $expW, $btnH)
    $expandTip = if ($rows -gt 1) {
        @{ Title = 'Свернуть'; Meta = 'Показать один квартал'; Badge = ''; Tone = 'plain' }
    }
    else {
        @{ Title = 'Развернуть'; Meta = 'Показать два квартала'; Badge = ''; Tone = 'plain' }
    }
    [void]$hits.Add(@{ Kind = 'expand'; Rect = $expRect; Tip = $expandTip })
    Draw-HeaderButton -Ui $Ui -Graphics $g -X $expRect.X -Y $expRect.Y -W $expRect.Width -H $expRect.Height `
        -Glyph ([string]$expandKey) -Label $expandLabel -GlyphLeft $false -Hover ($HoverKey -eq 'expand') -InkBrush $brAccent

    $titleLeft = $Ui.Pad
    $titleRight = $expRect.X - $Ui.BtnGap
    if ($ShowArrows) {
        $prevW = Measure-HeaderButton -Ui $Ui -Graphics $g -Glyph ([string][char]0x2039) -Label 'Квартал'
        $nextW = Measure-HeaderButton -Ui $Ui -Graphics $g -Glyph ([string][char]0x203A) -Label 'Квартал'
        $prevRect = New-Object System.Drawing.Rectangle($Ui.Pad, $hy, $prevW, $btnH)
        $nextRect = New-Object System.Drawing.Rectangle(($expRect.X - $Ui.BtnGap - $nextW), $hy, $nextW, $btnH)
        [void]$hits.Add(@{ Kind = 'prev'; Rect = $prevRect
                Tip = @{ Title = 'Предыдущий квартал'; Meta = 'Колесо вверх · ←'; Badge = ''; Tone = 'plain' } 
            })
        [void]$hits.Add(@{ Kind = 'next'; Rect = $nextRect
                Tip = @{ Title = 'Следующий квартал'; Meta = 'Колесо вниз · →'; Badge = ''; Tone = 'plain' } 
            })
        Draw-HeaderButton -Ui $Ui -Graphics $g -X $prevRect.X -Y $prevRect.Y -W $prevRect.Width -H $prevRect.Height `
            -Glyph ([string][char]0x2039) -Label 'Квартал' -GlyphLeft $true -Hover ($HoverKey -eq 'prev') -InkBrush $brAccent
        Draw-HeaderButton -Ui $Ui -Graphics $g -X $nextRect.X -Y $nextRect.Y -W $nextRect.Width -H $nextRect.Height `
            -Glyph ([string][char]0x203A) -Label 'Квартал' -GlyphLeft $false -Hover ($HoverKey -eq 'next') -InkBrush $brAccent
        $titleLeft = $prevRect.Right + $Ui.BtnGap
        if ($nextRect.X - $Ui.BtnGap -gt $titleLeft) { $titleRight = $nextRect.X - $Ui.BtnGap }
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
                if ($isToday) { $dayTip.Badge = 'Сегодня'; $dayTip.Tone = 'today' }
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
    $tip.Badge = (Fmt-Plural $Stat.Calendar 'календарный день' 'календарных дня' 'календарных дней')
    $tip.Lines = @(
        ("Рабочих {0}  ·  выходных {1}" -f $Stat.Work, $Stat.Off),
        ("Норма: {0} / {1} / {2} ч" -f (Fmt1 $Stat.Hours40), (Fmt1 $Stat.Hours36), (Fmt1 $Stat.Hours24))
    )
    return $tip
}

function Get-TipTone {
    # Цвет текста и подложки статуса — те же оттенки, что у дней в сетке:
    # «выходной» красный и «сокращённый» янтарный совпадают с ячейками месяца.
    # Рабочий день нейтральный: в сетке он чёрный на белом, и в подсказке ему
    # цвет не положен — акцент остаётся за событиями, а не за фоном недели.
    param($Ui, [string]$Tone)
    switch ($Tone) {
        'off' { return @{ Text = $Ui.C_Red; Fill = $Ui.C_RedBg } }
        'short' { return @{ Text = $Ui.C_Amber; Fill = $Ui.C_AmberBg } }
        # Акцент — только у «сегодня»: в сетке сегодняшний день тоже залит
        # акцентом, поэтому подсказка совпадает с виджетом.
        'today' { return @{ Text = $Ui.C_Accent; Fill = $Ui.C_AccentSoft } }
        default { return @{ Text = $Ui.C_Muted; Fill = $Ui.C_Divider } }
    }
}

function New-CalendarTooltipBitmap {
    # Карточка подсказки в стиле виджета: тот же светлый фон, рамка и скругление,
    # что у месяца, статус — «пилюлей» с мягкой подложкой и цветом дня из сетки.
    # Отрисовывается в bitmap и отдаётся окну подсказки.
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
    $w = [int][Math]::Ceiling($contentW) + $padX * 2
    $h = [int][Math]::Ceiling($szTitle.Height) + $gapRow + $row2H + $gapLines + [int][Math]::Ceiling($hLines) + $padY * 2
    $w = [Math]::Max($w, [int]($S * 104))
    $h = [Math]::Max($h, [int]($S * 34))

    # Полотно больше карточки на поля тени: тень рисуется по пикселям с альфой
    # (окно слоистое), поэтому вокруг карточки остаётся настоящая прозрачность.
    $sh = [int]$Ui.TipShadow
    $bw = $w + $sh * 2
    $bh = $h + $sh * 2 + [int]$Ui.TipShadowY

    $bmp = New-Object System.Drawing.Bitmap($bw, $bh, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'ClearTypeGridFit'
    $g.Clear([System.Drawing.Color]::Transparent)

    # Мягкая тень вокруг карточки (см. Add-SoftShadow).
    Add-SoftShadow -Graphics $g -X $sh -Y $sh -W $w -H $h -Margin $sh -Radius $Ui.TipRadius `
        -OffsetY ([int]$Ui.TipShadowY) -Color $Ui.C_TipShadow -Peak 13.0

    # Дальше вся отрисовка — в координатах карточки.
    $g.TranslateTransform([float]$sh, [float]$sh)

    # Карточка: белый фон под рамкой, чтобы сквозь неё не просвечивала тень.
    $cp = New-RoundPath 0 0 $w $h $Ui.TipRadius
    $cb = New-Object System.Drawing.SolidBrush($Ui.C_TipBg)
    $g.FillPath($cb, $cp); $cb.Dispose()
    # Рамка карточки — как у месяца; рисуется до содержимого, чтобы текст был поверх.
    $penB = New-Object System.Drawing.Pen($Ui.C_TipBorder)
    $g.DrawPath($penB, $cp); $penB.Dispose(); $cp.Dispose()

    $sfC = New-Object System.Drawing.StringFormat
    $sfC.Alignment = 'Center'; $sfC.LineAlignment = 'Center'
    $sfL = New-Object System.Drawing.StringFormat
    $sfL.Alignment = 'Near'; $sfL.LineAlignment = 'Center'

    $x = $padX
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
            $pp = New-RoundPath $rx $pillY $badgeW $pillH $Ui.CellRadius
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

    # Окно слоистое: содержимое отдаётся целиком из bitmap, регион и
    # BackgroundImage не нужны — они бы обрезали полупрозрачную тень по краям.
    $tf.Size = New-Object System.Drawing.Size($bmp.Width, $bmp.Height)

    # Точка отсчёта — курсор: смещаем карточку вниз-вправо, как системную.
    # Из смещения вычитаем поле тени, чтобы зазор считался от края карточки,
    # а не от прозрачного края полотна.
    $sh = [int]$Ui.TipShadow
    $cur = [System.Windows.Forms.Cursor]::Position
    $wa = [System.Windows.Forms.Screen]::FromControl($Surface).WorkingArea
    $px = $cur.X + $Ui.TipGap - $sh
    $py = $cur.Y + $Ui.TipGap - $sh
    # У правого/нижнего края экрана разворачиваем карточку в другую сторону.
    if (($px + $bmp.Width) -gt ($wa.Right - 2)) { $px = $cur.X - ($bmp.Width - $sh) - $Ui.TipGap }
    if (($py + $bmp.Height) -gt ($wa.Bottom - 2)) { $py = $cur.Y - ($bmp.Height - $sh) - $Ui.TipGap }
    $px = [Math]::Max($wa.Left + 2, [Math]::Min($px, $wa.Right - $bmp.Width - 2))
    $py = [Math]::Max($wa.Top + 2, [Math]::Min($py, $wa.Bottom - $bmp.Height - 2))
    # Показываем нативно, без активации — иначе виджет получит Deactivate и закроется.
    Show-LayeredWindow -Form $tf -Bitmap $bmp -X $px -Y $py
}

function Test-CalendarTooltipVisible {
    # Нативный показ не обновляет свойство WinForms .Visible, поэтому спрашиваем
    # у системы, видно ли окно подсказки на самом деле.
    param($Ui)
    if ($null -eq $Ui -or -not $Ui.ContainsKey('TipForm')) { return $false }
    $tf = $Ui.TipForm
    if ($null -eq $tf -or $tf.IsDisposed -or -not $tf.IsHandleCreated) { return $false }
    return [PcNative.Layered]::IsWindowVisible($tf.Handle)
}

function Hide-CalendarTooltip {
    param($Ui)
    if ($null -eq $Ui -or -not $Ui.ContainsKey('TipForm')) { return }
    Hide-LayeredWindow -Form $Ui.TipForm
}
