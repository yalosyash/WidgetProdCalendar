# tools\Install-Autostart.ps1
# Добавляет (или удаляет) ярлык виджета в папке автозагрузки текущего пользователя.
#
# Включить автозапуск:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Install-Autostart.ps1
# Выключить:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Install-Autostart.ps1 -Remove

[CmdletBinding()]
param([switch]$Remove)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$vbs = Join-Path $PSScriptRoot 'Run-Hidden.vbs'
$ico = Join-Path $projectRoot 'assets\calendar.ico'
$startup = [Environment]::GetFolderPath('Startup')
$lnkPath = Join-Path $startup 'ProdCalendar.lnk'

if ($Remove) {
    if (Test-Path $lnkPath) {
        Remove-Item -Force $lnkPath
        Write-Host "Автозапуск удалён: $lnkPath"
    }
    else {
        Write-Host "Ярлык автозапуска не найден: $lnkPath"
    }
    return
}

if (-not (Test-Path $vbs)) { throw "Не найден $vbs" }

$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut($lnkPath)
$lnk.TargetPath = $vbs
$lnk.WorkingDirectory = $projectRoot
$lnk.Description = 'Виджет производственного календаря'
if (Test-Path $ico) { $lnk.IconLocation = "$ico,0" }
$lnk.Save()

Write-Host "Автозапуск включён: $lnkPath"
Write-Host "Проверить после перезахода в систему: значок появится в трее автоматически."
