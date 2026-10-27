# lib\Data.ps1
# Слой данных: получение кодов дня и расчёт производных величин.
# Цепочка источников: API isdayoff.ru -> локальный кэш -> вшитый офлайн-набор.
# Подключается точкой:  . .\lib\Data.ps1

Set-StrictMode -Version Latest

# --- TLS: в PowerShell 5.1 без этого https может отвалиться даже при живой сети.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script:LibRoot   = Split-Path -Parent $PSScriptRoot          # ...\ProdCalendar
$script:OfflineFile = Join-Path $script:LibRoot 'data\offline.json'
$script:AppDataDir  = Join-Path $env:LOCALAPPDATA 'ProdCalendar'
$script:CacheFile   = Join-Path $script:AppDataDir 'cache.json'
$script:LogFile     = Join-Path $script:AppDataDir 'prodcalendar.log'

if (-not (Test-Path $script:AppDataDir)) { New-Item -ItemType Directory -Force -Path $script:AppDataDir | Out-Null }

function Write-Log {
    param([string]$Message)
    try {
        $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
        Add-Content -Path $script:LogFile -Value $line -Encoding UTF8
    } catch { }
}

function Test-LeapYear {
    param([int]$Year)
    return (($Year % 4 -eq 0) -and ($Year % 100 -ne 0 -or $Year % 400 -eq 0))
}

function Get-ExpectedLength {
    param([int]$Year)
    if (Test-LeapYear -Year $Year) { return 366 } else { return 365 }
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json)
    } catch {
        Write-Log "Не удалось прочитать JSON '$Path': $($_.Exception.Message)"
        return $null
    }
}

function Save-Cache {
    param([hashtable]$Table)
    try {
        $json = $Table | ConvertTo-Json -Compress
        [System.IO.File]::WriteAllText($script:CacheFile, $json, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        Write-Log "Не удалось записать кэш: $($_.Exception.Message)"
    }
}

function ConvertTo-CodesTable {
    # Приводит PSCustomObject/JSON к hashtable {год -> строка кодов}
    param($Obj)
    $t = @{}
    if ($null -eq $Obj) { return $t }
    foreach ($p in $Obj.PSObject.Properties) { $t[$p.Name] = [string]$p.Value }
    return $t
}

function Get-AvailableYears {
    # Годы, для которых есть вшитый офлайн-набор.
    $off = ConvertTo-CodesTable (Read-JsonFile $script:OfflineFile)
    if ($off.Count -eq 0) { return @() }
    return @($off.Keys | ForEach-Object { [int]$_ } | Sort-Object)
}

function Get-DayCodes {
    [CmdletBinding()]
    param([int]$Year = (Get-Date).Year)

    $expectedLen = Get-ExpectedLength -Year $Year
    $uri = "https://isdayoff.ru/api/getdata?year=$Year&pre=1"

    # 1) API. Используем HttpWebRequest с Proxy=$null: Invoke-WebRequest в PS 5.1
    #    тратит ~1 с на автодетект прокси (WPAD) и может подвиснуть в фоне.
    try {
        $req = [System.Net.HttpWebRequest]::Create($uri)
        $req.Proxy = $null
        $req.Timeout = 8000
        $req.ReadWriteTimeout = 8000
        $req.UserAgent = 'ProdCalendar/1.0'
        $resp = $req.GetResponse()
        try {
            $sr = New-Object System.IO.StreamReader($resp.GetResponseStream())
            $content = $sr.ReadToEnd()
            $sr.Close()
        } finally {
            $resp.Close()
        }
        $codes = $content.Trim()
        if ($codes.Length -eq $expectedLen -and $codes -match '^[012]+$') {
            # сохраняем/обновляем кэш
            $cache = ConvertTo-CodesTable (Read-JsonFile $script:CacheFile)
            $cache["$Year"] = $codes
            Save-Cache -Table $cache
            return @{ Year = $Year; Codes = $codes; Source = 'api' }
        }
        Write-Log "Год ${Year}: ответ API не прошёл проверку (длина $($codes.Length), ожидалось $expectedLen)"
    } catch {
        Write-Log "Год ${Year}: API недоступен ($($_.Exception.Message))"
    }

    # 2) Локальный кэш
    $cache = ConvertTo-CodesTable (Read-JsonFile $script:CacheFile)
    if ($cache.ContainsKey("$Year")) {
        $codes = $cache["$Year"]
        if ($codes.Length -eq $expectedLen -and $codes -match '^[012]+$') {
            Write-Log "Год ${Year}: используется кэш ($script:CacheFile)"
            return @{ Year = $Year; Codes = $codes; Source = 'cache' }
        }
    }

    # 3) Вшитый офлайн
    $off = ConvertTo-CodesTable (Read-JsonFile $script:OfflineFile)
    if ($off.ContainsKey("$Year")) {
        Write-Log "Год ${Year}: используется офлайн-набор"
        return @{ Year = $Year; Codes = $off["$Year"]; Source = 'offline' }
    }

    throw "Нет данных за $Year ни из API, ни из кэша, ни офлайн."
}

function Get-MonthStat {
    param(
        [Parameter(Mandatory)][string]$Codes,
        [Parameter(Mandatory)][int]$Month,
        [int]$Year = 0
    )
    if ($Month -lt 1 -or $Month -gt 12) { throw "Месяц вне диапазона: $Month" }
    $len = if ($Year -gt 0) { Get-ExpectedLength $Year } else { $Codes.Length }
    if ($Codes.Length -ne $len) { throw "Длина кодов ($($Codes.Length)) не совпадает с ожидаемой ($len)" }

    $offset = 0
    for ($m = 1; $m -lt $Month; $m++) { $offset += [DateTime]::DaysInMonth($Year, $m) }
    $days = [DateTime]::DaysInMonth($Year, $Month)
    $slice = $Codes.Substring($offset, $days)

    $work = 0; $off = 0; $short = 0
    foreach ($ch in $slice.ToCharArray()) {
        switch ($ch) {
            '0' { $work++ }
            '1' { $off++ }
            '2' { $short++ }
            default { $work++ }
        }
    }
    $wr = $work + $short
    [pscustomobject]@{
        Month       = $Month
        Calendar    = $days
        Work        = $wr
        Off         = $off
        Short       = $short
        Hours40     = [math]::Round($wr * 8.0 - $short, 1)
        Hours36     = [math]::Round($wr * 7.2 - $short, 1)
        Hours24     = [math]::Round($wr * 4.8 - $short, 1)
    }
}

function Get-YearStat {
    param([Parameter(Mandatory)][string]$Codes)
    $work = 0; $off = 0; $short = 0
    foreach ($ch in $Codes.ToCharArray()) {
        switch ($ch) {
            '0' { $work++ }
            '1' { $off++ }
            '2' { $short++ }
            default { $work++ }
        }
    }
    $wr = $work + $short
    [pscustomobject]@{
        Calendar = $Codes.Length
        Work     = $wr
        Off      = $off
        Short    = $short
        Hours40  = [math]::Round($wr * 8.0 - $short, 1)
        Hours36  = [math]::Round($wr * 7.2 - $short, 1)
        Hours24  = [math]::Round($wr * 4.8 - $short, 1)
    }
}

function Get-DayType {
    # Возвращает 'work' | 'off' | 'short' для конкретной даты.
    param(
        [Parameter(Mandatory)][string]$Codes,
        [Parameter(Mandatory)][datetime]$Date
    )
    $jan1 = Get-Date -Year $Date.Year -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
    $idx = [int](($Date.Date - $jan1.Date).TotalDays)
    if ($idx -lt 0 -or $idx -ge $Codes.Length) { throw "Дата вне диапазона кодов: $Date" }
    switch ($Codes[$idx]) {
        '0' { return 'work' }
        '1' { return 'off' }
        '2' { return 'short' }
        default { return 'work' }
    }
}

$script:MonthShortRu = @('Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек')
$script:MonthFullRu  = @('Январь','Февраль','Март','Апрель','Май','Июнь',
                        'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь')
