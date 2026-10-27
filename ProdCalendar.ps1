# ProdCalendar.ps1
# Точка входа виджета производственного календаря для Windows.
# Запуск:  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\ProdCalendar.ps1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Один экземпляр ---------------------------------------------------------
$script:Mutex = New-Object System.Threading.Mutex($false, 'ProdCalendarWidgetSingleInstance')
$hasHandle = $false
try { $hasHandle = $script:Mutex.WaitOne(0, $false) } catch { }
if (-not $hasHandle) {
    # Уже запущен — тихо выходим, второй значок не создаём.
    exit 0
}

# --- DPI: убираем «мыло», затем считаем масштаб -----------------------------
Add-Type -Namespace PcNative -Name User32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
[DllImport("user32.dll")] public static extern bool DestroyIcon(System.IntPtr handle);
'@
try { [void][PcNative.User32]::SetProcessDPIAware() } catch { }

$probe = New-Object System.Windows.Forms.Form
$pg = $probe.CreateGraphics()
$scale = [double]$pg.DpiX / 96.0
$pg.Dispose(); $probe.Dispose()

# --- Пути -------------------------------------------------------------------
$script:Root     = $PSScriptRoot
$script:DataLib  = Join-Path $script:Root 'lib\Data.ps1'

. (Join-Path $script:Root 'lib\Calendar.ps1')

# --- Состояние --------------------------------------------------------------
$script:Today  = (Get-Date).Date
$script:State  = @{ Year = $script:Today.Year; Month = $script:Today.Month; Codes = $null; Source = ''; Years = @() }
$script:LoadPs = $null
$script:LoadHandle = $null
$script:LoadYear = $null
$script:LastShow = [DateTime]::MinValue

# --- UI календаря -----------------------------------------------------------
$script:Ui = New-CalendarUi -Scale $scale
$form = $script:Ui.Form

function Show-Calendar {
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $x = $wa.Right - $form.Width - [int](8 * $script:Ui.Scale)
    $y = $wa.Bottom - $form.Height - [int](8 * $script:Ui.Scale)
    $form.Location = New-Object System.Drawing.Point($x, $y)
    $form.Show()
    $form.Activate()
    $form.BringToFront()
    $script:LastShow = [DateTime]::Now
}

function Hide-Calendar { $form.Hide() }

function Toggle-Calendar {
    if ($form.Visible) { Hide-Calendar } else { Show-Calendar }
}

function Update-View {
    $st = $script:State
    $year = $st.Year; $month = $st.Month

    $script:Ui.Title.Text = "{0} {1}" -f $script:CcMonthFull[$month - 1], $year

    if ($null -eq $st.Codes) {
        $script:Ui.Summary.Text = "Нет данных за $year.`r`nПроверьте подключение к интернету или добавьте год офлайн-набором."
        $script:Ui.Source.Text = "данные: нет"
        return
    }

    Draw-CalendarGrid -Ui $script:Ui -Year $year -Month $month -Codes $st.Codes

    $ms = Get-MonthStat -Codes $st.Codes -Month $month -Year $year
    $ys = Get-YearStat -Codes $st.Codes
    $script:Ui.Summary.Text = (
        "Месяц: календарных {0}, рабочих {1}, выходных {2}`r`n" +
        "Норма 40ч {3} · 36ч {4} · 24ч {5}`r`n" +
        "Год {6}: рабочих {7} из {8}`r`n" +
        "Норма года {9} / {10} / {11}"
    ) -f $ms.Calendar, $ms.Work, $ms.Off,
         $ms.Hours40, $ms.Hours36, $ms.Hours24,
         $year, $ys.Work, $ys.Calendar, $ys.Hours40, $ys.Hours36, $ys.Hours24

    $srcName = switch ($st.Source) {
        'api'     { 'isdayoff.ru' }
        'cache'   { 'кэш (последний ответ API)' }
        'offline' { "офлайн ($year)" }
        default   { $st.Source }
    }
    $script:Ui.Source.Text = "данные: $srcName"
}

# --- Загрузка данных в фоне (не блокирует трей) -----------------------------
function Start-LoadYear {
    param([int]$Year)
    if ($script:LoadPs) {
        try { $script:LoadPs.Stop(); $script:LoadPs.Dispose() } catch { }
        $script:LoadPs = $null
    }
    $ps = [System.Management.Automation.PowerShell]::Create()
    [void]$ps.AddScript({
        param($libPath, $yr)
        . $libPath
        $d = Get-DayCodes -Year $yr
        [pscustomobject]@{ Ok = $true; Year = $d.Year; Codes = $d.Codes; Source = $d.Source; Years = @(Get-AvailableYears) }
    }).AddArgument($script:DataLib).AddArgument($Year)
    $script:LoadHandle = $ps.BeginInvoke()
    $script:LoadPs = $ps
    $script:LoadYear = $Year
}

function Complete-LoadIfReady {
    if (-not $script:LoadPs) { return }
    if (-not $script:LoadHandle.IsCompleted) { return }
    $ps = $script:LoadPs
    try {
        $out = $ps.EndInvoke($script:LoadHandle)
        $ray = $ps.Streams.Error.ReadAll()
        if ($ray.Count -gt 0 -or $out.Count -eq 0) {
            $script:State.Codes = $null
            $script:State.Source = 'none'
        } else {
            $res = $out[0]
            $script:State.Year = [int]$res.Year
            $script:State.Codes = [string]$res.Codes
            $script:State.Source = [string]$res.Source
            $script:State.Years = @($res.Years)
        }
    } catch {
        $script:State.Codes = $null
        $script:State.Source = 'none'
    } finally {
        $ps.Dispose()
        $script:LoadPs = $null
        $script:LoadHandle = $null
    }
    Update-View
}

# --- Навигация --------------------------------------------------------------
function Step-Month {
    param([int]$Delta)
    $st = $script:State
    if ($null -eq $st.Codes) { return }
    $m = $st.Month + $Delta
    $y = $st.Year
    while ($m -lt 1) { $m += 12; $y -= 1 }
    while ($m -gt 12) { $m -= 12; $y += 1 }
    $minY = ($script:State.Years | Measure-Object -Minimum).Minimum
    $maxY = ($script:State.Years | Measure-Object -Maximum).Maximum
    if ($script:State.Years.Count -gt 0 -and ($y -lt $minY -or $y -gt $maxY)) { return }
    if ($y -ne $st.Year) {
        $st.Year = $y; $st.Month = $m; $st.Codes = $null
        Update-View
        Start-LoadYear -Year $y
    } else {
        $st.Month = $m
        Update-View
    }
}

$script:Ui.BtnPrev.Add_Click({ Step-Month -Delta -1 })
$script:Ui.BtnNext.Add_Click({ Step-Month -Delta 1 })

$script:Ui.Grid.Add_MouseMove({
    param($s, $e)
    $tag = $s.Tag
    if ($null -eq $tag) { return }
    $col = [int][Math]::Floor($e.X / $script:Ui.CellW)
    $row = [int][Math]::Floor(($e.Y - $script:Ui.WkH) / $script:Ui.CellH)
    if ($col -lt 0 -or $col -gt 6 -or $row -lt 0 -or $row -gt 5) {
        $script:Ui.Tip.SetToolTip($s, '')
        return
    }
    $pos = $row * 7 + $col
    $day = $pos - $tag.StartCol + 1
    if ($day -lt 1 -or $day -gt $tag.Days) { $script:Ui.Tip.SetToolTip($s, ''); return }
    $code = $tag.Codes[ (Get-YearOffset -Year $tag.Year -Month $tag.Month) + ($day - 1) ]
    $txt = Build-DayTooltip -Year $tag.Year -Month $tag.Month -Day $day -Code $code
    $script:Ui.Tip.SetToolTip($s, $txt)
})

function Get-YearOffset {
    param([int]$Year, [int]$Month)
    $off = 0
    for ($m = 1; $m -lt $Month; $m++) { $off += [DateTime]::DaysInMonth($Year, $m) }
    return $off
}

# --- Закрытие окна по Esc и потере фокуса -----------------------------------
$form.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { Hide-Calendar; $e.Handled = $true }
})
$form.Add_Deactivate({
    if ((([DateTime]::Now) - $script:LastShow).TotalMilliseconds -gt 250) { Hide-Calendar }
})

# --- Трей -------------------------------------------------------------------
$icon = $null
$icoPath = Join-Path $script:Root 'assets\calendar.ico'
if (Test-Path $icoPath) { try { $icon = New-Object System.Drawing.Icon($icoPath) } catch { } }
if ($null -eq $icon) { $icon = [System.Drawing.SystemIcons]::Application }

$tray = New-Object System.Windows.Forms.NotifyIcon
$tray.Icon = $icon
$tray.Text = 'Производственный календарь'
$tray.Visible = $true

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miCal = $menu.Items.Add('Календарь')
$miUpd = $menu.Items.Add('Обновить данные')
$miExit = $menu.Items.Add('Выход')
$miCal.Add_Click({ Show-Calendar })
$miUpd.Add_Click({ $script:State.Codes = $null; Start-LoadYear -Year $script:State.Year })
$miExit.Add_Click({
    $tray.Visible = $false
    $tray.Dispose()
    $script:ctx.ExitThread()
})
$tray.ContextMenuStrip = $menu
$tray.Add_MouseClick({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Toggle-Calendar }
})
$tray.Add_DoubleClick({ Show-Calendar })

# --- Таймер: завершение фоновой загрузки ------------------------------------
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200
$timer.Add_Tick({ Complete-LoadIfReady })
$timer.Start()

# --- Старт ------------------------------------------------------------------
Start-LoadYear -Year $script:State.Year
Update-View

$script:ctx = New-Object System.Windows.Forms.ApplicationContext
try {
    [System.Windows.Forms.Application]::Run($script:ctx)
} finally {
    $timer.Stop(); $timer.Dispose()
    if ($tray) { $tray.Visible = $false; $tray.Dispose() }
    $form.Dispose()
    if ($script:LoadPs) { try { $script:LoadPs.Dispose() } catch { } }
    try { $script:Mutex.ReleaseMutex() } catch { }
    $script:Mutex.Dispose()
}
