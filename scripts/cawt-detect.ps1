# cawt-detect.ps1 — CLI инструмент для поиска CAWT-совместимого Tcl
#
# Использование:
#   .\cawt-detect.ps1 [-Verbose] [-ForceRefresh]
#
# Возвращает:
#   Путь к интерпретатору (exit 0) или ошибку (exit 1)

param(
    [switch]$Verbose,
    [switch]$ForceRefresh
)

# Импортируем модуль
$modulePath = Join-Path $PSScriptRoot 'cawt-office.ps1'
Import-Module $modulePath -Force

# Применяем verbose
if ($Verbose) {
    $script:CawtVerbose = $true
}

# Запускаем детект
$tclPath = Get-CawtTclPath -ForceRefresh:$ForceRefresh

if ($tclPath) {
    # Выводим только путь (для использования в скриптах)
    Write-Output $tclPath
    
    if ($Verbose) {
        $info = Get-CawtInfo
        Write-Host ""
        Write-Host "Summary:" -ForegroundColor Cyan
        Write-Host "  Path:      $($info.Path)"
        Write-Host "  Tcl:       $($info.TclVersion)"
        Write-Host "  Twapi:     $($info.TwapiVersion)"
        Write-Host "  CAWT:      $($info.CawtVersion)"
    }
    
    exit 0
} else {
    $configPath = Join-Path (Split-Path $PSScriptRoot -Parent) '.cawt-office'
    
    Write-Host @"

CAWT-compatible Tcl interpreter not found.

Download portable Tcl 9 Kit with bundled CAWT:
  https://github.com/nektomk/tclwinkit/releases

Or create config file: $configPath
  tclsh_path=R:\tclwinkit-msvc.exe

"@ -ForegroundColor Red
    
    exit 1
}
