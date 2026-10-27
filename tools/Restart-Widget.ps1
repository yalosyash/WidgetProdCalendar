# tools\Restart-Widget.ps1
# Перезапуск виджета: тихо гасит текущий экземпляр и поднимает его заново
# (без окна консоли, через Run-Hidden.vbs).
#
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Restart-Widget.ps1
#
# Используется, чтобы применить изменения кода без ручного перезапуска.

[CmdletBinding()]
param(
    # Сколько ждать, пока процесс завершится и освободит мьютекс единственного экземпляра.
    [int]$WaitMs = 800
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$vbs = Join-Path $PSScriptRoot 'Run-Hidden.vbs'
$entry = Join-Path $projectRoot 'ProdCalendar.ps1'

if (-not (Test-Path $entry)) { throw "Не найден входной скрипт: $entry" }
if (-not (Test-Path $vbs)) { throw "Не найден $vbs" }

function Get-WidgetProcesses {
    # Экземпляры виджета — powershell.exe, запущенный на наш ProdCalendar.ps1.
    # Себя ($PID) исключаем, чтобы скрипт не убивал собственный процесс.
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -and $_.CommandLine -like "*$entry*" }
}

# --- 1) гасим текущий экземпляр --------------------------------------------
$running = @(Get-WidgetProcesses)
foreach ($p in $running) {
    try { Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop } catch { }
}
if ($running.Count -gt 0) { Write-Host "Остановлен виджет: PID $(($running | ForEach-Object ProcessId) -join ', ')" }

# Ждём завершения: мьютекс единственного экземпляра должен освободиться.
$deadline = (Get-Date).AddMilliseconds($WaitMs)
while ((Get-Date) -lt $deadline -and @(Get-WidgetProcesses).Count -gt 0) {
    Start-Sleep -Milliseconds 50
}

# --- 2) поднимаем заново ----------------------------------------------------
$sh = New-Object -ComObject WScript.Shell
[void]$sh.Run("wscript.exe `"$vbs`"", 0, $false)

# --- 3) убеждаемся, что новый экземпляр поднялся ---------------------------
$started = $null
$deadline = (Get-Date).AddSeconds(15)
while ((Get-Date) -lt $deadline) {
    $started = @(Get-WidgetProcesses)
    if ($started.Count -gt 0) { break }
    Start-Sleep -Milliseconds 200
}

if ($started.Count -eq 0) { throw 'Виджет не поднялся: процесс не найден за 15 секунд.' }
Write-Host "Виджет запущен: PID $(($started | ForEach-Object ProcessId) -join ', ')"
