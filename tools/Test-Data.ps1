# tools\Test-Data.ps1
# Проверка слоя данных и навигации (быстрый smoke-тест, не часть приложения).
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
Write-Host "11.05.2026 -> $(Get-DayType -Codes $d.Codes -Date (Get-Date '2026-05-11'))"
Write-Host "01.04.2026 -> $(Get-DayType -Codes $d.Codes -Date (Get-Date '2026-04-01'))"

# --- навигация через модуль приложения (ловит регрессии вроде отсутствия Get-MonthStat) ---
Write-Host "`nПроверка UI-навигации:"
$app = Initialize-ProdCalendarApp -Scale 1.0
try {
    & $app.StartLoad $app.State.Year
    $deadline = (Get-Date).AddSeconds(10)
    while (-not $app.LoadHnd.IsCompleted -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 100 }
    & $app.CompleteLoad
    Write-Host ("  загрузка: год={0} источник={1} лет={2}" -f $app.State.Year, $app.State.Source, $app.State.Codes.Length)
    foreach ($delta in 1, 1, -1, -1, -1) {
        & $app.Step -Delta $delta
        Write-Host ("  шаг {0,2}: {1}" -f $delta, $app.Ui.Title.Text)
    }
    Write-Host ("  сводка: {0}" -f ($app.Ui.Summary.Text -replace "`r`n", ' | '))
    Write-Host ("  источник: {0}" -f $app.Ui.Source.Text)
    Write-Host '  навигация OK'
} finally {
    Stop-ProdCalendarApp -App $app
}
