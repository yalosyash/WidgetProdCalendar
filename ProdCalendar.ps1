# ProdCalendar.ps1
# Точка входа виджета производственного календаря для Windows.
# Запуск:  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\ProdCalendar.ps1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Защита от второго экземпляра -------------------------------------------
$script:Mutex = New-Object System.Threading.Mutex($false, 'ProdCalendarWidgetSingleInstance')
$hasHandle = $false
try { $hasHandle = $script:Mutex.WaitOne(0, $false) } catch { }
if (-not $hasHandle) {
    # Уже запущен — тихо выходим, второй значок не создаём.
    exit 0
}

# --- DPI: убираем «мыло», затем считаем масштаб -----------------------------
Add-Type -Namespace PcNative -Name User32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
'@
try { [void][PcNative.User32]::SetProcessDPIAware() } catch { }

$probe = New-Object System.Windows.Forms.Form
$pg = $probe.CreateGraphics()
$scale = [double]$pg.DpiX / 96.0
$pg.Dispose(); $probe.Dispose()

# --- Приложение -------------------------------------------------------------
. (Join-Path $PSScriptRoot 'lib\App.ps1')

$app = Initialize-ProdCalendarApp -Scale $scale
& $app.UpdateView

try {
    [System.Windows.Forms.Application]::Run($app.Ctx)
}
finally {
    Stop-ProdCalendarApp -App $app
    try { $script:Mutex.ReleaseMutex() } catch { }
    $script:Mutex.Dispose()
}
