# lib\App.ps1
# Логика приложения: трей, окно, навигация, отрисовка, фоновая загрузка данных.
# Сам подключает Data.ps1 и Calendar.ps1. Точка входа (ProdCalendar.ps1) только
# создаёт app и запускает цикл сообщений.
#
# ВАЖНО про обработчики событий: scriptblock-и выполняются позже и в другой
# области видимости, поэтому обращаются к состоянию через $script:pc (script scope),
# а НЕ через локальные переменные функции Initialize-ProdCalendarApp.

. (Join-Path $PSScriptRoot 'Data.ps1')
. (Join-Path $PSScriptRoot 'Calendar.ps1')

function Get-YearDayOffset {
    param([int]$Year, [int]$Month)
    $off = 0
    for ($m = 1; $m -lt $Month; $m++) { $off += [DateTime]::DaysInMonth($Year, $m) }
    return $off
}

function Get-SourceDisplayName {
    param([string]$Source, [int]$Year)
    switch ($Source) {
        'api'     { return 'isdayoff.ru' }
        'cache'   { return 'кэш (последний ответ API)' }
        'offline' { return "офлайн ($Year)" }
        default   { return 'нет' }
    }
}

function Initialize-ProdCalendarApp {
    [CmdletBinding()]
    param([double]$Scale = 1.0)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $ui = New-CalendarUi -Scale $Scale
    $today = (Get-Date).Date

    # Пути фиксируем сразу (внутри отложенных scriptblock-ов $PSScriptRoot пуст).
    $dataLib  = Join-Path $PSScriptRoot 'Data.ps1'
    $projRoot = Split-Path -Parent $PSScriptRoot

    $pc = @{
        Ui       = $ui
        Form     = $ui.Form
        Scale    = $Scale
        State    = @{ Year = $today.Year; Month = $today.Month; Codes = $null; Source = ''; Years = @() }
        LoadPs   = $null
        LoadHnd  = $null
        LastShow = [DateTime]::MinValue
        Tray     = $null
        Timer    = $null
        Ctx      = New-Object System.Windows.Forms.ApplicationContext
    }
    $script:pc = $pc

    # --- показ / скрытие -----------------------------------------------------
    $pc.Show = {
        $pc = $script:pc; $form = $pc.Form
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $form.Location = New-Object System.Drawing.Point(
            ($wa.Right - $form.Width - [int](8 * $pc.Scale)),
            ($wa.Bottom - $form.Height - [int](8 * $pc.Scale)))
        $form.Show(); $form.Activate(); $form.BringToFront()
        $pc.LastShow = [DateTime]::Now
    }
    $pc.Hide = { $script:pc.Form.Hide() }
    $pc.Toggle = { $pc = $script:pc; if ($pc.Form.Visible) { & $pc.Hide } else { & $pc.Show } }

    # --- отрисовка -----------------------------------------------------------
    $pc.UpdateView = {
        $pc = $script:pc; $ui = $pc.Ui; $st = $pc.State
        $ui.Title.Text = "{0} {1}" -f $script:CcMonthFull[$st.Month - 1], $st.Year

        if ($null -eq $st.Codes) {
            $ui.Summary.Text = "Нет данных за $($st.Year).`r`nПроверьте подключение к интернету или добавьте год офлайн-набором (-Years)."
            $ui.Source.Text = 'данные: нет'
            $ui.Grid.BackgroundImage = $null
            $ui.Grid.Tag = $null
            return
        }

        Draw-CalendarGrid -Ui $ui -Year $st.Year -Month $st.Month -Codes $st.Codes

        $ms = Get-MonthStat -Codes $st.Codes -Month $st.Month -Year $st.Year
        $ys = Get-YearStat -Codes $st.Codes
        $ui.Summary.Text = (
            "Месяц: календарных {0}, рабочих {1}, выходных {2}`r`n" +
            "Норма 40ч {3} · 36ч {4} · 24ч {5}`r`n" +
            "Год {6}: рабочих {7} из {8}`r`n" +
            "Норма года {9} / {10} / {11}"
        ) -f $ms.Calendar, $ms.Work, $ms.Off,
             $ms.Hours40, $ms.Hours36, $ms.Hours24,
             $st.Year, $ys.Work, $ys.Calendar, $ys.Hours40, $ys.Hours36, $ys.Hours24

        $ui.Source.Text = 'данные: ' + (Get-SourceDisplayName -Source $st.Source -Year $st.Year)
    }

    # --- загрузка данных (в фоне) -------------------------------------------
    $pc.StartLoad = {
        param([int]$Year)
        $pc = $script:pc
        if ($pc.LoadPs) { try { $pc.LoadPs.Stop(); $pc.LoadPs.Dispose() } catch { } ; $pc.LoadPs = $null }
        $ps = [System.Management.Automation.PowerShell]::Create()
        [void]$ps.AddScript({
            param($libPath, $yr)
            . $libPath
            $d = Get-DayCodes -Year $yr
            [pscustomobject]@{ Ok = $true; Year = $d.Year; Codes = $d.Codes; Source = $d.Source; Years = @(Get-AvailableYears) }
        }).AddArgument($pc.DataLib).AddArgument($Year)
        $pc.LoadHnd = $ps.BeginInvoke()
        $pc.LoadPs = $ps
    }

    $pc.CompleteLoad = {
        $pc = $script:pc
        if (-not $pc.LoadPs) { return }
        if (-not $pc.LoadHnd.IsCompleted) { return }
        $ps = $pc.LoadPs
        try {
            $out = $ps.EndInvoke($pc.LoadHnd)
            $errs = $ps.Streams.Error.ReadAll()
            if ($errs.Count -gt 0 -or $out.Count -eq 0) {
                $pc.State.Codes = $null; $pc.State.Source = 'none'
            } else {
                $res = $out[0]
                $pc.State.Year = [int]$res.Year
                $pc.State.Codes = [string]$res.Codes
                $pc.State.Source = [string]$res.Source
                $pc.State.Years = @($res.Years)
            }
        } catch {
            $pc.State.Codes = $null; $pc.State.Source = 'none'
        } finally {
            $ps.Dispose(); $pc.LoadPs = $null; $pc.LoadHnd = $null
        }
        & $pc.UpdateView
    }

    # --- навигация -----------------------------------------------------------
    $pc.Step = {
        param([int]$Delta)
        $pc = $script:pc; $st = $pc.State
        $m = $st.Month + $Delta
        $y = $st.Year
        while ($m -lt 1) { $m += 12; $y -= 1 }
        while ($m -gt 12) { $m -= 12; $y += 1 }
        if ($st.Years.Count -gt 0) {
            $minY = ($st.Years | Measure-Object -Minimum).Minimum
            $maxY = ($st.Years | Measure-Object -Maximum).Maximum
            if ($y -lt $minY -or $y -gt $maxY) { return }
        }
        if ($y -ne $st.Year) {
            $st.Year = $y; $st.Month = $m; $st.Codes = $null
            & $pc.UpdateView
            & $pc.StartLoad $y
        } else {
            $st.Month = $m
            & $pc.UpdateView
        }
    }

    # --- события UI ----------------------------------------------------------
    $ui.BtnPrev.Add_Click({ & $script:pc.Step -Delta -1 })
    $ui.BtnNext.Add_Click({ & $script:pc.Step -Delta 1 })

    $ui.Grid.Add_MouseMove({
        param($s, $e)
        $pc = $script:pc; $tag = $s.Tag
        if ($null -eq $tag) { return }
        $col = [int][Math]::Floor($e.X / $pc.Ui.CellW)
        $row = [int][Math]::Floor(($e.Y - $pc.Ui.WkH) / $pc.Ui.CellH)
        if ($col -lt 0 -or $col -gt 6 -or $row -lt 0 -or $row -gt 5) { $pc.Ui.Tip.SetToolTip($s, ''); return }
        $day = ($row * 7 + $col) - $tag.StartCol + 1
        if ($day -lt 1 -or $day -gt $tag.Days) { $pc.Ui.Tip.SetToolTip($s, ''); return }
        $code = $tag.Codes[(Get-YearDayOffset -Year $tag.Year -Month $tag.Month) + ($day - 1)]
        $pc.Ui.Tip.SetToolTip($s, (Build-DayTooltip -Year $tag.Year -Month $tag.Month -Day $day -Code $code))
    })

    $ui.Form.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { & $script:pc.Hide; $e.Handled = $true }
    })
    $ui.Form.Add_Deactivate({
        if ((([DateTime]::Now) - $script:pc.LastShow).TotalMilliseconds -gt 250) { & $script:pc.Hide }
    })

    # --- трей ---------------------------------------------------------------
    $icon = $null
    $icoPath = Join-Path $projRoot 'assets\calendar.ico'
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
    $miCal.Add_Click({ & $script:pc.Show })
    $miUpd.Add_Click({ $script:pc.State.Codes = $null; & $script:pc.UpdateView; & $script:pc.StartLoad $script:pc.State.Year })
    $miExit.Add_Click({ $script:pc.Tray.Visible = $false; $script:pc.Tray.Dispose(); $script:pc.Ctx.ExitThread() })
    $tray.ContextMenuStrip = $menu
    $tray.Add_MouseClick({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { & $script:pc.Toggle }
    })
    $tray.Add_DoubleClick({ & $script:pc.Show })

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 200
    $timer.Add_Tick({ & $script:pc.CompleteLoad })
    $timer.Start()

    $pc.Tray = $tray
    $pc.Timer = $timer
    $pc.DataLib = $dataLib
    return $pc
}

function Stop-ProdCalendarApp {
    param($App)
    if ($null -eq $App) { return }
    try { if ($App.Timer) { $App.Timer.Stop(); $App.Timer.Dispose() } } catch { }
    try { if ($App.Tray) { $App.Tray.Visible = $false; $App.Tray.Dispose() } } catch { }
    try { if ($App.LoadPs) { $App.LoadPs.Dispose() } } catch { }
    try { if ($App.Form) { $App.Form.Dispose() } } catch { }
}
