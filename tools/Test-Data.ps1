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

foreach ($m in 4,5,6,11) {
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
    & $app.Step -Delta 1;  Write-Host ("  Q{0} {1}: {2}" -f $app.Quarter, $app.Year, ((& $names) -join ' | '))
    & $app.ToggleExpand;   Write-Host ("  разворот ({0} карточек): {1}" -f (@($app.Ui.Hits | Where-Object Kind -eq 'card').Count), ((& $names) -join ' | '))
    & $app.ToggleExpand;   Write-Host ("  сворачивание: {0} карточек" -f (@($app.Ui.Hits | Where-Object Kind -eq 'card').Count))
    Write-Host '  навигация OK'
} finally {
    Stop-ProdCalendarApp -App $app
}
