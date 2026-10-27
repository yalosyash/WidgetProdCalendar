# tools\Test-Data.ps1
# Быстрая проверка слоя данных (шаг 2). Не часть приложения.
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'lib\Data.ps1')

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
