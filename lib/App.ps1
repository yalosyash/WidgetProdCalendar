# lib\App.ps1
# Логика приложения: трей, окно, навигация ПО КВАРТАЛАМ, разворот окна вверх
# (ещё один квартал), отрисовка и фоновая загрузка данных по годам.
# Сам подключает Data.ps1 и Calendar.ps1.
#
# Обработчики событий WinForms берут состояние из $script:pc (script scope),
# потому что выполняются позже и в другой области видимости.

. (Join-Path $PSScriptRoot 'Data.ps1')
. (Join-Path $PSScriptRoot 'Calendar.ps1')

function Get-QuarterMonths {
    param([int]$Year, [int]$Quarter)
    $first = ($Quarter - 1) * 3 + 1
    # Скобки обязательны: запятая в литерале массива связывает сильнее '+'.
    return @(($first), ($first + 1), ($first + 2))
}

function Get-SourceDisplayName {
    param([string]$Source, [int]$Year)
    switch ($Source) {
        'api'     { return 'isdayoff.ru' }
        'cache'   { return 'кэш (последний ответ API)' }
        'offline' { return "офлайн ($Year)" }
        default   { return 'нет данных' }
    }
}

function Initialize-ProdCalendarApp {
    [CmdletBinding()]
    param([double]$Scale = 1.0)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $ui = New-CalendarUi -Scale $Scale
    $today = (Get-Date).Date

    $dataLib  = Join-Path $PSScriptRoot 'Data.ps1'

    $years = @(Get-AvailableYears)
    if ($years.Count -eq 0) { $years = @($today.Year) }

    $pc = @{
        Ui          = $ui
        Form        = $ui.Surface
        Scale       = $Scale
        Year        = $today.Year
        Quarter     = [int][Math]::Floor(($today.Month - 1) / 3) + 1
        Expanded    = $false
        Years       = $years
        CodesByYear = @{}
        SrcByYear   = @{}
        Loading     = @{}
        LoadPs      = @{}
        LoadHnd     = @{}
        HoverKey    = ''
        LastShow    = [DateTime]::MinValue
        Tray        = $null
        Timer       = $null
        Ctx         = New-Object System.Windows.Forms.ApplicationContext
        DataLib     = $dataLib
    }
    $script:pc = $pc

    # --- показ / скрытие -----------------------------------------------------
    $pc.ResizeSurface = {
        $pc = $script:pc
        $rows = if ($pc.Expanded) { 2 } else { 1 }
        $size = Get-CalendarSurfaceSize -Ui $pc.Ui -Rows $rows
        $pc.Ui.Surface.ClientSize = New-Object System.Drawing.Size($size[0], $size[1])
    }
    $pc.Reposition = {
        $pc = $script:pc; $form = $pc.Form
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $mx = [int](8 * $pc.Scale)
        $form.Location = New-Object System.Drawing.Point(($wa.Right - $form.Width - $mx), ($wa.Bottom - $form.Height - $mx))
    }
    $pc.Show = {
        $pc = $script:pc
        & $pc.Reposition
        $pc.Form.Show(); $pc.Form.Activate(); $pc.Form.BringToFront()
        $pc.LastShow = [DateTime]::Now
    }
    $pc.Hide = { $script:pc.Form.Hide() }
    $pc.Toggle = { $pc = $script:pc; if ($pc.Form.Visible) { & $pc.Hide } else { & $pc.Show } }

    # --- данные --------------------------------------------------------------
    $pc.EnsureYear = {
        param([int]$Year)
        $pc = $script:pc
        if ($pc.CodesByYear.ContainsKey($Year) -or $pc.Loading.ContainsKey($Year)) { return }
        $pc.Loading[$Year] = $true
        $ps = [System.Management.Automation.PowerShell]::Create()
        [void]$ps.AddScript({
            param($libPath, $yr)
            . $libPath
            $d = Get-DayCodes -Year $yr
            [pscustomobject]@{ Year = $d.Year; Codes = $d.Codes; Source = $d.Source }
        }).AddArgument($pc.DataLib).AddArgument($Year)
        $pc.LoadHnd[$Year] = $ps.BeginInvoke()
        $pc.LoadPs[$Year] = $ps
    }
    $pc.CompleteLoads = {
        $pc = $script:pc
        $changed = $false
        foreach ($y in @($pc.LoadPs.Keys)) {
            $ps = $pc.LoadPs[$y]
            $hnd = $pc.LoadHnd[$y]
            if (-not $hnd.IsCompleted) { continue }
            try {
                $out = $ps.EndInvoke($hnd)
                $errs = $ps.Streams.Error.ReadAll()
                if ($errs.Count -eq 0 -and $out.Count -gt 0) {
                    $pc.CodesByYear[$y] = [string]$out[0].Codes
                    $pc.SrcByYear[$y] = [string]$out[0].Source
                } else {
                    $pc.SrcByYear[$y] = 'none'
                }
            } catch {
                $pc.SrcByYear[$y] = 'none'
            } finally {
                $ps.Dispose(); $pc.LoadPs.Remove($y); $pc.LoadHnd.Remove($y); $pc.Loading.Remove($y)
            }
            $changed = $true
        }
        if ($changed) { & $pc.UpdateView }
    }

    # --- отрисовка -----------------------------------------------------------
    $pc.UpdateView = {
        $pc = $script:pc; $ui = $pc.Ui
        $q = $pc.Quarter
        $prevQ = $q - 1; $prevY = $pc.Year
        if ($prevQ -lt 1) { $prevQ = 4; $prevY -= 1 }

        $months = New-Object System.Collections.ArrayList
        $needYears = New-Object System.Collections.Generic.HashSet[int]
        $addMonth = {
            param([int]$y, [int]$m)
            $codes = $null
            if ($pc.CodesByYear.ContainsKey($y)) { $codes = $pc.CodesByYear[$y] }
            $stat = $null
            if ($null -ne $codes) { $stat = Get-MonthStat -Codes $codes -Month $m -Year $y }
            [void]$months.Add([pscustomobject]@{ Year = $y; Month = $m; Codes = $codes; Stat = $stat })
            [void]$needYears.Add($y)
        }

        if ($pc.Expanded) {
            foreach ($m in (Get-QuarterMonths -Year $prevY -Quarter $prevQ)) { & $addMonth $prevY $m }
        }
        foreach ($m in (Get-QuarterMonths -Year $pc.Year -Quarter $q)) { & $addMonth $pc.Year $m }

        $header = "{0} · {1} квартал" -f $pc.Year, $script:CcQuarter[$q - 1]
        if ($pc.Expanded) {
            if ($prevY -eq $pc.Year) {
                $header = "{0} · {1}–{2} квартал" -f $pc.Year, $script:CcQuarter[$prevQ - 1], $script:CcQuarter[$q - 1]
            } else {
                $header = "{0} {1} кв. → {2} {3} кв." -f $prevY, $script:CcQuarter[$prevQ - 1], $pc.Year, $script:CcQuarter[$q - 1]
            }
        }

        $curCodes = $null
        if ($pc.CodesByYear.ContainsKey($pc.Year)) { $curCodes = $pc.CodesByYear[$pc.Year] }
        $src = 'нет'
        if ($pc.SrcByYear.ContainsKey($pc.Year)) { $src = $pc.SrcByYear[$pc.Year] }
        if ($null -eq $curCodes) {
            $footerLeft = 'данные: загрузка...'
        } else {
            $footerLeft = 'данные: ' + (Get-SourceDisplayName -Source $src -Year $pc.Year)
        }
        $footerRight = if ($pc.Expanded) { "$([char]0x25BC) свернуть · ‹ › кварталы" } else { "$([char]0x25B2) развернуть · ‹ › кварталы" }

        Draw-CalendarSurface -Ui $ui -Months $months.ToArray() -HeaderTitle $header `
            -FooterLeft $footerLeft -FooterRight $footerRight -HoverKey $pc.HoverKey

        foreach ($y in $needYears) { & $pc.EnsureYear $y }
    }

    # --- навигация -----------------------------------------------------------
    $pc.Step = {
        param([int]$Delta)
        $pc = $script:pc
        $q = $pc.Quarter + $Delta
        $y = $pc.Year
        while ($q -lt 1) { $q += 4; $y -= 1 }
        while ($q -gt 4) { $q -= 4; $y += 1 }
        if ($pc.Years.Count -gt 0) {
            $minY = ($pc.Years | Measure-Object -Minimum).Minimum
            $maxY = ($pc.Years | Measure-Object -Maximum).Maximum
            if ($y -lt $minY -or $y -gt $maxY) { return }
        }
        $pc.Year = $y; $pc.Quarter = $q
        & $pc.UpdateView
    }
    $pc.ToggleExpand = {
        $pc = $script:pc
        $pc.Expanded = -not $pc.Expanded
        & $pc.ResizeSurface
        & $pc.Reposition
        & $pc.UpdateView
    }

    # --- мышь / клавиатура ---------------------------------------------------
    $pc.HitTest = {
        param([int]$X, [int]$Y)
        $pc = $script:pc
        $hits = $pc.Ui.Hits
        for ($i = $hits.Count - 1; $i -ge 0; $i--) {
            $r = $hits[$i].Rect
            if ($X -ge $r.X -and $X -lt ($r.X + $r.Width) -and $Y -ge $r.Y -and $Y -lt ($r.Y + $r.Height)) { return $hits[$i] }
        }
        return $null
    }

    $pc.Form.Add_MouseMove({
        param($s, $e)
        $pc = $script:pc
        $hit = & $pc.HitTest $e.X $e.Y
        $key = ''
        if ($hit) {
            if ($hit.Kind -in @('prev','next','expand')) { $key = $hit.Kind }
            $pc.Ui.Tip.SetToolTip($pc.Form, $(if ($hit.Text) { $hit.Text } else { '' }))
        } else {
            $pc.Ui.Tip.SetToolTip($pc.Form, '')
        }
        if ($key -ne $pc.HoverKey) {
            $pc.HoverKey = $key
            & $pc.UpdateView
        }
    })
    $pc.Form.Add_MouseLeave({ $script:pc.Ui.Tip.SetToolTip($script:pc.Form, '') })
    $pc.Form.Add_MouseClick({
        param($s, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        $pc = $script:pc
        $hit = & $pc.HitTest $e.X $e.Y
        if (-not $hit) { return }
        switch ($hit.Kind) {
            'prev'   { & $pc.Step -Delta -1 }
            'next'   { & $pc.Step -Delta 1 }
            'expand' { & $pc.ToggleExpand }
        }
    })
    $pc.Form.Add_KeyDown({
        param($s, $e)
        $pc = $script:pc
        switch ($e.KeyCode) {
            ([System.Windows.Forms.Keys]::Escape)   { & $pc.Hide; $e.Handled = $true }
            ([System.Windows.Forms.Keys]::Left)     { & $pc.Step -Delta -1; $e.Handled = $true }
            ([System.Windows.Forms.Keys]::Right)    { & $pc.Step -Delta 1; $e.Handled = $true }
            ([System.Windows.Forms.Keys]::Up)       { if (-not $pc.Expanded) { & $pc.ToggleExpand }; $e.Handled = $true }
            ([System.Windows.Forms.Keys]::Down)     { if ($pc.Expanded) { & $pc.ToggleExpand }; $e.Handled = $true }
        }
    })
    $pc.Form.Add_Deactivate({
        if ((([DateTime]::Now) - $script:pc.LastShow).TotalMilliseconds -gt 250) { & $script:pc.Hide }
    })

    # --- трей ---------------------------------------------------------------
    $icon = $null
    $icoPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'assets\calendar.ico'
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
    $miUpd.Add_Click({
        $pc = $script:pc
        $pc.CodesByYear = @{}; $pc.SrcByYear = @{}
        & $pc.UpdateView
    })
    $miExit.Add_Click({ $script:pc.Tray.Visible = $false; $script:pc.Tray.Dispose(); $script:pc.Ctx.ExitThread() })
    $tray.ContextMenuStrip = $menu
    $tray.Add_MouseClick({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { & $script:pc.Toggle }
    })
    $tray.Add_DoubleClick({ & $script:pc.Show })

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 200
    $timer.Add_Tick({ & $script:pc.CompleteLoads })
    $timer.Start()

    $pc.Tray = $tray
    $pc.Timer = $timer

    & $pc.ResizeSurface
    return $pc
}

function Stop-ProdCalendarApp {
    param($App)
    if ($null -eq $App) { return }
    try { if ($App.Timer) { $App.Timer.Stop(); $App.Timer.Dispose() } } catch { }
    try { if ($App.Tray) { $App.Tray.Visible = $false; $App.Tray.Dispose() } } catch { }
    try { foreach ($ps in @($App.LoadPs.Values)) { $ps.Dispose() } } catch { }
    try { if ($App.Form) { $App.Form.Dispose() } } catch { }
}
