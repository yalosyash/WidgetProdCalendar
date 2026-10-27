# tools\New-OfflineData.ps1
# Генерирует data\offline.json — вшитый офлайн-набор кодов дней по годам.
#
# Источник: API isdayoff.ru (запрос с &pre=1, иначе не приходят сокращённые дни).
# Зачем из API, а не из "ручных списков праздников": официальные переносы и
# предпраздничные дни меняются каждый год, и вручную их легко ошибиться.
# Особый контроль: для 2026 данные жёстко сверяются с эталоном (раздел 4.1 плана),
# поэтому ошибка в источнике не пройдёт молча.
#
# Запуск (из корня ProdCalendar):
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\New-OfflineData.ps1
#   powershell.exe ... -File .\tools\New-OfflineData.ps1 -Years 2026,2027

[CmdletBinding()]
param(
    [int[]]$Years = @(2026, 2027),
    [string]$OutFile = (Join-Path (Split-Path -Parent $PSScriptRoot) 'data\offline.json')
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Эталон 2026 для самопроверки (не подгоняем — сверяем).
$Expect2026 = @{
    Length  = 365
    Work    = 247   # рабочих дней (коды '0' + '2', сокращённые тоже рабочие)
    Off     = 118   # код '1'
    Short   = 4     # код '2'
    Hours40 = 1972.0
    Hours36 = 1774.4
    Hours24 = 1181.6
}

function Get-YearCodes {
    param([int]$Year)
    $uri = "https://isdayoff.ru/api/getdata?year=$Year&pre=1"
    $resp = Invoke-WebRequest -Uri $uri -UseBasicParsing -TimeoutSec 20
    if ($resp.StatusCode -ne 200) {
        throw "Год ${Year}: HTTP $($resp.StatusCode)"
    }
    $codes = ($resp.Content).Trim()
    $expected = if ((($Year % 4) -eq 0) -and (($Year % 100) -ne 0 -or ($Year % 400) -eq 0)) { 366 } else { 365 }
    if ($codes.Length -ne $expected) {
        throw "Год ${Year}: длина $($codes.Length), ожидалось $expected"
    }
    if ($codes -notmatch '^[012]+$') {
        throw "Год ${Year}: недопустимые символы в ответе ('$codes')"
    }
    return $codes
}

function Get-Stats {
    param([string]$Codes)
    $work = 0; $off = 0; $short = 0
    foreach ($ch in $Codes.ToCharArray()) {
        switch ($ch) {
            '0' { $work++ }
            '1' { $off++ }
            '2' { $short++ }
        }
    }
    $wr = $work + $short
    [pscustomobject]@{
        Work    = $wr
        Off     = $off
        Short   = $short
        Hours40 = [math]::Round($wr * 8.0 - $short, 1)
        Hours36 = [math]::Round($wr * 7.2 - $short, 1)
        Hours24 = [math]::Round($wr * 4.8 - $short, 1)
    }
}

$result = @{}
foreach ($year in $Years) {
    Write-Host "Год ${year}: запрашиваю isdayoff.ru ..." -NoNewline
    $codes = Get-YearCodes -Year $year
    $stats = Get-Stats -Codes $codes
    $result["$year"] = $codes
    Write-Host (" OK  раб={0} вых={1} сокр={2}  норм 40/36/24 = {3} / {4} / {5}" -f `
            $stats.Work, $stats.Off, $stats.Short, $stats.Hours40, $stats.Hours36, $stats.Hours24)

    if ($year -eq 2026) {
        $fail = @()
        if ($codes.Length -ne $Expect2026.Length) { $fail += "длина $($codes.Length) <> $($Expect2026.Length)" }
        if ($stats.Work -ne $Expect2026.Work) { $fail += "рабочих $($stats.Work) <> $($Expect2026.Work)" }
        if ($stats.Off -ne $Expect2026.Off) { $fail += "выходных $($stats.Off) <> $($Expect2026.Off)" }
        if ($stats.Short -ne $Expect2026.Short) { $fail += "сокращённых $($stats.Short) <> $($Expect2026.Short)" }
        if ($stats.Hours40 -ne $Expect2026.Hours40) { $fail += "40ч $($stats.Hours40) <> $($Expect2026.Hours40)" }
        if ($stats.Hours36 -ne $Expect2026.Hours36) { $fail += "36ч $($stats.Hours36) <> $($Expect2026.Hours36)" }
        if ($stats.Hours24 -ne $Expect2026.Hours24) { $fail += "24ч $($stats.Hours24) <> $($Expect2026.Hours24)" }
        if ($fail.Count -gt 0) {
            throw "2026 НЕ совпал с эталоном: $($fail -join '; ')"
        }
        Write-Host '          2026 сверен с эталоном раздела 4.1 — совпадает.'
    }
}

$json = $result | ConvertTo-Json -Compress
$dir = Split-Path -Parent $OutFile
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
# JSON — UTF-8 БЕЗ BOM (иначе ConvertFrom-Json может поперхнуться).
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Записано: $OutFile ($($json.Length) байт)"
