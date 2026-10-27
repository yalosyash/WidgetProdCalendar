# tools\New-Shortcut.ps1
# Создаёт ярлык «Производственный календарь» на рабочем столе (и/или в меню «Пуск»),
# по двойному клику на который виджет запускается без окна консоли.
#
# Создать ярлык на рабочем столе:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\New-Shortcut.ps1
# Ярлык в меню «Пуск» вместо рабочего стола:
#   ... -File .\tools\New-Shortcut.ps1 -StartMenu
# И там, и там:
#   ... -File .\tools\New-Shortcut.ps1 -Both
# Удалить ярлык(и):
#   ... -File .\tools\New-Shortcut.ps1 -Remove

[CmdletBinding()]
param(
    [switch]$StartMenu,
    [switch]$Both,
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$vbs = Join-Path $PSScriptRoot 'Run-Hidden.vbs'
$ico = Join-Path $projectRoot 'assets\calendar.ico'
$desktop = [Environment]::GetFolderPath('Desktop')
$startup = [Environment]::GetFolderPath('StartMenu')
$name = 'Производственный календарь'

if (-not (Test-Path $vbs)) { throw "Не найден $vbs" }

if ($Both) { $StartMenu = $true }

$targets = @(Join-Path $desktop ($name + '.lnk'))
if ($StartMenu) {
    $progDir = Join-Path $startup 'Programs'
    if (-not (Test-Path $progDir)) { $progDir = $startup }
    $targets += Join-Path $progDir ($name + '.lnk')
}

if ($Remove) {
    foreach ($t in $targets) {
        if (Test-Path $t) { Remove-Item -Force $t; Write-Host "Ярлык удалён: $t" }
        else { Write-Host "Ярлык не найден: $t" }
    }
    return
}

# Иконку генерируем при необходимости.
if (-not (Test-Path $ico)) {
    & (Join-Path $PSScriptRoot 'New-Icon.ps1') -OutPath $ico
}

$ws = New-Object -ComObject WScript.Shell
foreach ($t in $targets) {
    $lnk = $ws.CreateShortcut($t)
    $lnk.TargetPath = $vbs
    $lnk.WorkingDirectory = $projectRoot
    $lnk.IconLocation = "$ico,0"
    $lnk.Description = 'Виджет производственного календаря'
    $lnk.WindowStyle = 1
    $lnk.Save()
    Write-Host "Ярлык создан: $t"
}

Write-Host ''
Write-Host 'Готово. Двойной клик по ярлыку запускает виджет (значок появится в трее у часов).'
Write-Host 'Если ярлык на рабочем столе не обновился сразу — обновите вид (F5) на рабочем столе.'
