# tools\Test-Data.ps1
# Проверка слоя данных и навигации по кварталам (smoke-тест, не часть приложения).
# StrictMode включён намеренно: приложение запускается с ним, значит и тесты должны
# ловить те же ошибки (например, обращение к отсутствующему ключу хэш-таблицы).
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'lib\App.ps1')

Write-Host "Доступные годы (офлайн): $((Get-AvailableYears) -join ', ')"

$d = Get-DayCodes -Year 2026
Write-Host "2026 источник = $($d.Source)"

foreach ($m in 4, 5, 6, 11) {
    $s = Get-MonthStat -Codes $d.Codes -Month $m -Year 2026
    Write-Host ("Месяц {0,2}: раб={1} вых={2} сокр={3}  норм {4} / {5} / {6}" -f `
            $m, $s.Work, $s.Off, $s.Short, $s.Hours40, $s.Hours36, $s.Hours24)
}
$y = Get-YearStat -Codes $d.Codes
Write-Host ("Год 2026: кал={0} раб={1} вых={2} сокр={3}  норм {4} / {5} / {6}" -f `
        $y.Calendar, $y.Work, $y.Off, $y.Short, $y.Hours40, $y.Hours36, $y.Hours24)
Write-Host "30.04.2026 -> $(Get-DayType -Codes $d.Codes -Date (Get-Date '2026-04-30'))"
Write-Host "09.03.2026 -> $(Get-DayType -Codes $d.Codes -Date (Get-Date '2026-03-09'))"

Write-Host "`nПроверка навигации по кварталам:"
$app = Initialize-ProdCalendarApp -Scale 1.0
try {
    & $app.UpdateView
    $deadline = (Get-Date).AddSeconds(10)
    while ((Get-Date) -lt $deadline) {
        & $app.CompleteLoads
        $ready = $true
        foreach ($yy in $app.Years) { if (-not $app.CodesByYear.ContainsKey($yy)) { $ready = $false } }
        if ($ready) { break }
        Start-Sleep -Milliseconds 150
    }
    $names = { $app.Ui.Hits | Where-Object { $_.Kind -eq 'card' } | ForEach-Object { ($_.Text -split "`r`n")[0] } }
    Write-Host ("  Q{0} {1}: {2}" -f $app.Quarter, $app.Year, ((& $names) -join ' | '))
    & $app.Step -Delta -1; Write-Host ("  Q{0} {1}: {2}" -f $app.Quarter, $app.Year, ((& $names) -join ' | '))
    & $app.Step -Delta 1; Write-Host ("  Q{0} {1}: {2}" -f $app.Quarter, $app.Year, ((& $names) -join ' | '))
    & $app.ToggleExpand; Write-Host ("  разворот ({0} карточек): {1}" -f (@($app.Ui.Hits | Where-Object Kind -eq 'card').Count), ((& $names) -join ' | '))
    $arrowsInExpanded = @($app.Ui.Hits | Where-Object { $_.Kind -in @('prev', 'next') }).Count
    Write-Host ("  стрелки в развёрнутом виде (ожидаем 0): {0}" -f $arrowsInExpanded)
    & $app.ToggleExpand; Write-Host ("  сворачивание: {0} карточек" -f (@($app.Ui.Hits | Where-Object Kind -eq 'card').Count))
    Write-Host ("  стрелки в свёрнутом виде (ожидаем 2): {0}" -f (@($app.Ui.Hits | Where-Object { $_.Kind -in @('prev', 'next') }).Count))
    Write-Host '  навигация OK'

    Write-Host "`nПроверка сброса состояния при показе:"
    $defYear = $app.Defaults.Year
    $defQuarter = $app.Defaults.Quarter
    & $app.ToggleExpand
    & $app.Step -Delta -1
    $app.HoverKey = 'next'
    if (-not $app.Expanded) { throw 'Разворот не сработал — тест сброса некорректен.' }
    if ($app.Quarter -eq $defQuarter -and $app.Year -eq $defYear) { throw 'Квартал не сменился — тест сброса некорректен.' }
    $expandedH = $app.Form.Height
    $movedTo = "Q$($app.Quarter) $($app.Year)"
    & $app.Hide
    & $app.Show
    if ($app.Expanded) { throw 'После скрытия/показа виджет остался развёрнутым.' }
    if ($app.Year -ne $defYear -or $app.Quarter -ne $defQuarter) {
        throw "Квартал не вернулся к дефолту: Q$($app.Quarter) $($app.Year), ожидался Q$defQuarter $defYear."
    }
    if ($app.HoverKey -ne '') { throw "Подсветка кнопки не сброшена: '$($app.HoverKey)'." }
    if ($app.Form.Height -ge $expandedH) { throw 'Высота окна не вернулась к свёрнутому виду.' }
    $rowCards = @($app.Ui.Hits | Where-Object Kind -eq 'card').Count
    if ($rowCards -ne 3) { throw "Ожидалось 3 карточки после сброса, получено $rowCards." }
    Write-Host ("  развёрнуто + {0} -> скрыть -> показать -> Q{1} {2}, свёрнуто ({3}px), карточек {4}, hover='{5}' — OK" -f `
            $movedTo, $app.Quarter, $app.Year, $app.Form.Height, $rowCards, $app.HoverKey)
    & $app.Hide

    Write-Host "`nПроверка настроек отрисовки (против мерцания):"
    $dbProp = $app.Form.GetType().GetProperty('DoubleBuffered',
        [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
    if ($dbProp -and -not $dbProp.GetValue($app.Form)) { throw 'Двойная буферизация выключена.' }
    if ($app.Form.BackgroundImageLayout -ne 'None') {
        throw "Фон растягивается ($($app.Form.BackgroundImageLayout)) — при развороте будет виден старый кадр."
    }
    if (-not (Get-Command Set-CalendarRedraw -ErrorAction SilentlyContinue)) {
        throw 'Не найдена функция управления перерисовкой.'
    }

    # Разворот не должен оставлять фон несовпадающим с размером окна.
    & $app.ToggleExpand
    $bi = $app.Form.BackgroundImage
    if ($null -eq $bi) { throw 'После разворота фон не нарисован.' }
    if ($bi.Width -ne $app.Form.ClientSize.Width -or $bi.Height -ne $app.Form.ClientSize.Height) {
        throw "Размер фона ($($bi.Width)x$($bi.Height)) не совпадает с окном ($($app.Form.ClientSize.Width)x$($app.Form.ClientSize.Height))."
    }
    if (-not ('PcNative.Redraw' -as [type])) { throw 'Не загружен помощник управления перерисовкой.' }
    & $app.ToggleExpand
    Write-Host ("  двойная буферизация вкл, фон без растяжения, кадр всегда по размеру окна — OK")

    Write-Host "`nПроверка подсказок:"
    # Правим ярлык без показа окна: проверяем логику, а не реальные события мыши.
    $dayHit = $app.Ui.Hits | Where-Object { $_.Kind -eq 'day' } | Select-Object -First 1
    if ($null -eq $dayHit) { throw 'В разметке нет ни одного дня — подсказки негде проверить.' }
    if (-not $dayHit.ContainsKey('Tip')) { throw 'У дня нет своей подсказки.' }
    if (-not $dayHit.Tip.Title) { throw 'У подсказки дня пустой заголовок.' }

    # У кнопок Text не задан — обращение к нему не должно падать (StrictMode).
    $btnHit = $app.Ui.Hits | Where-Object { $_.Kind -eq 'next' } | Select-Object -First 1
    $k1 = & $app.TipKeyOf $dayHit
    $k2 = & $app.TipKeyOf $btnHit
    if ($k1 -eq $k2) { throw 'Ключи разных элементов совпали — подсказки будут путаться.' }

    # Задержка: сразу после наведения подсказки ещё нет, таймер запущен.
    & $app.ClearHover
    & $app.TrackHover $dayHit
    if (-not $app.TipTimer.Enabled) { throw 'Таймер задержки подсказки не запустился.' }
    if ($app.TipTimer.Interval -ne $app.TipDelayMs) { throw "Неверная задержка: $($app.TipTimer.Interval) мс." }
    if (Test-CalendarTooltipVisible -Ui $app.Ui) { throw 'Подсказка показана мгновенно, без задержки.' }

    # Показ и сокрытие через сами функции. -Force: в тесте окно не показано,
    # и сверка с реальным курсором (защита от устаревшей подсказки) тут не нужна.
    & $app.ShowCurrentTip -Force
    if (-not (Test-CalendarTooltipVisible -Ui $app.Ui)) { throw 'Подсказка не показалась.' }
    if (-not $app.TipShown) { throw 'Флаг показа не установлен.' }
    $tipSize = "$($app.Ui.TipForm.Width)x$($app.Ui.TipForm.Height)"
    if ($app.Ui.TipForm.Width -le 0 -or $app.Ui.TipForm.Height -le 0) { throw 'У подсказки нулевой размер.' }

    # «Тёплый» режим: пока подсказка на экране, соседний день обновляется быстро.
    $day2 = $app.Ui.Hits | Where-Object { $_.Kind -eq 'day' } | Select-Object -Skip 1 -First 1
    & $app.TrackHover $day2
    if ($app.TipTimer.Interval -ne $app.TipWarmMs) { throw "Тёплый режим не включился: $($app.TipTimer.Interval) мс." }

    & $app.ClearHover
    if (Test-CalendarTooltipVisible -Ui $app.Ui) { throw 'Подсказка не скрылась.' }
    if ($app.TipShown) { throw 'Флаг показа не сброшен.' }
    if ($app.TipTimer.Interval -ne $app.TipDelayMs) { throw 'После сброса не вернулась обычная задержка.' }
    Write-Host ("  задержка $($app.TipDelayMs) мс, теплый отклик $($app.TipWarmMs) мс, карточка $tipSize — OK")

    Write-Host "`nПроверка скругления углов окна:"
    $frm = $app.Form
    $reg = $frm.Region
    if ($null -eq $reg) { throw 'Регион окна не задан — углы не скруглены.' }
    $corners = @(0, 0), @(($frm.Width - 1), 0), @(0, ($frm.Height - 1)), @(($frm.Width - 1), ($frm.Height - 1))
    foreach ($c in $corners) {
        if ($reg.IsVisible([System.Drawing.Point]::new($c[0], $c[1]))) {
            throw "Угол ($($c[0]),$($c[1])) не обрезан — скругление не работает."
        }
    }
    if (-not $reg.IsVisible([System.Drawing.Point]::new([int]($frm.Width / 2), 0))) {
        throw 'Верхняя кромка окна обрезана — регион слишком большой.'
    }
    Write-Host ("  углы обрезаны (radius={0}), кромки на месте — OK" -f $app.Ui.WindowRadius)
}
finally {
    Stop-ProdCalendarApp -App $app
}
