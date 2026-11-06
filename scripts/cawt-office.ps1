# cawt-office.ps1 — модуль для детекта Tcl-интерпретатора с CAWT
# Кэширование только в памяти сессии PowerShell
#
# Использование:
#   Import-Module .\cawt-office.ps1
#   $tclPath = Get-CawtTclPath
#   & $tclPath script.tcl args...

$script:CawtTclPath = $null
$script:CawtVerbose = $false
$script:CawtConfigPath = $null

# ============================================================================
# Внутренние функции
# ============================================================================

function Initialize-CawtConfig {
    # Инициализируем путь к конфигу при первом вызове
    if ($null -eq $script:CawtConfigPath) {
        $scriptDir = Split-Path $PSScriptRoot -Parent
        $script:CawtConfigPath = Join-Path $scriptDir '.cawt-office'
    }
}

function Read-CawtConfig {
    # Читаем конфиг .cawt-office (если существует)
    # Возвращает хэш с ключами: tclsh_path, verbose
    Initialize-CawtConfig
    
    $config = @{}
    
    if (Test-Path $script:CawtConfigPath) {
        try {
            Get-Content $script:CawtConfigPath | ForEach-Object {
                $line = $_.Trim()
                # Пропускаем комментарии и пустые строки
                if ($line -eq '' -or $line.StartsWith('#')) { return }
                
                # Парсим key=value
                if ($line -match '^([^=]+)=(.+)$') {
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    $config[$key] = $value
                }
            }
        } catch {
            # Игнорируем ошибки чтения
        }
    }
    
    return $config
}

function Test-CawtCompatible {
    param(
        [string]$TclPath
    )
    
    # Быстрая проверка: есть ли CAWT в данном интерпретаторе
    $checkScript = Join-Path $PSScriptRoot 'cawt-check.tcl'
    
    if (-not (Test-Path $TclPath)) {
        return $false
    }
    
    try {
        if (-not (Test-Path $checkScript)) {
            # Фоллбэк: проверяем через inline-скрипт
            $result = & $TclPath -c "package require cawt" 2>&1
            return ($LASTEXITCODE -eq 0)
        }
        
        # Запускаем проверку
        $output = & $TclPath $checkScript 2>&1
        $success = ($LASTEXITCODE -eq 0)
        
        return $success
    } catch {
        # Ошибка при запуске (не исполняемый файл, нет прав, и т.д.)
        if ($script:CawtVerbose) {
            Write-Host "[CAWT] Error running $TclPath : $_" -ForegroundColor Red
        }
        return $false
    }
}

function Get-CawtVersions {
    param(
        [string]$TclPath
    )
    
    # Получаем версии Tcl, Twapi, CAWT
    $checkScript = Join-Path $PSScriptRoot 'cawt-check.tcl'
    $output = & $TclPath $checkScript 2>&1
    
    $info = @{
        Path = $TclPath
        TclVersion = $null
        TwapiVersion = $null
        CawtVersion = $null
        Compatible = $false
    }
    
    foreach ($line in $output) {
        if ($line -match 'OK:.*CAWT=(\S+).*Twapi=(\S+).*Tcl=(\S+)') {
            $info.CawtVersion = $matches[1]
            $info.TwapiVersion = $matches[2]
            $info.TclVersion = $matches[3]
            $info.Compatible = $true
            break
        }
    }
    
    return $info
}

function Find-CawtCandidates {
    # Возвращает список всех возможных путей к Tcl-интерпретаторам
    $candidates = @()
    
    # 1. PATH: ищем tclsh*.exe через where.exe
    try {
        $whereOutput = where.exe tclsh*.exe 2>$null
        foreach ($path in $whereOutput) {
            if ($path -notmatch '^INFO:') {
                $candidates += $path.Trim()
            }
        }
    } catch {
        # where.exe не найден или нет результатов
    }
    
    # 2. Реестр: HKLM\SOFTWARE\Tcl\Tclsh
    $regPaths = @(
        'HKLM:\SOFTWARE\Tcl\Tclsh',
        'HKCU:\SOFTWARE\Tcl\Tclsh',
        'HKLM:\SOFTWARE\WOW6432Node\Tcl\Tclsh'
    )
    
    foreach ($regPath in $regPaths) {
        try {
            $key = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
            if ($key -and $key.'(Default)') {
                $candidates += $key.'(Default)'
            }
        } catch {
            # Ключ не найден
        }
    }
    
    # 3. Реестр: Uninstall keys
    try {
        $uninstallKeys = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        
        foreach ($keyPath in $uninstallKeys) {
            Get-ItemProperty -Path $keyPath -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DisplayName -and (
                    $_.DisplayName -like '*Tcl*' -or
                    $_.DisplayName -like '*TWAPI*' -or
                    $_.DisplayName -like '*CAWT*' -or
                    $_.DisplayName -like '*ActiveTcl*' -or
                    $_.DisplayName -like '*BAWT*'
                )
            } | ForEach-Object {
                if ($_.InstallLocation) {
                    $tclsh = Join-Path $_.InstallLocation 'bin\tclsh.exe'
                    if (Test-Path $tclsh) {
                        $candidates += $tclsh
                    }
                }
            }
        }
    } catch {
        # Игнорируем ошибки реестра
    }
    
    # 4. Известные пути установки
    $knownPaths = @(
        'C:\ATcl\bin\tclsh.exe',
        'C:\ATcl\bin\tclsh86.exe',
        'C:\Program Files\ActiveTcl\8.6\bin\tclsh.exe',
        'C:\Program Files (x86)\ActiveTcl\8.6\bin\tclsh.exe',
        'C:\Program Files\MagicSplat\Tcl\bin\tclsh.exe',
        "$env:LOCALAPPDATA\MagicSplat\Tcl\bin\tclsh.exe",
        'C:\msys64\ucrt64\bin\tclsh.exe',
        'C:\msys64\ucrt64\bin\tclsh86.exe',
        'C:\cygwin\bin\tclsh.exe',
        'C:\Program Files\BAWT\bin\tclsh.exe',
        'R:\tclwinkit-msvc.exe'
    )
    
    foreach ($path in $knownPaths) {
        if (Test-Path $path) {
            $candidates += $path
        }
    }
    
    # Удаляем дубликаты
    return $candidates | Select-Object -Unique
}

# ============================================================================
# Публичные функции
# ============================================================================

function Get-CawtTclPath {
    param(
        [switch]$ForceRefresh
    )
    
    <#
    .SYNOPSIS
        Находит совместимый с CAWT Tcl-интерпретатор.
    
    .DESCRIPTION
        Возвращает полный путь к Tcl-интерпретатору с установленным CAWT.
        Результат кэшируется в памяти сессии PowerShell.
    
    .PARAMETER ForceRefresh
        Игнорировать кэш и выполнить повторный детект.
    
    .EXAMPLE
        $tclPath = Get-CawtTclPath
        & $tclPath script.tcl arg1 arg2
    
    .EXAMPLE
        $tclPath = Get-CawtTclPath -ForceRefresh
    #>
    
    # Возвращаем закэшированный (если есть и не ForceRefresh)
    if (-not $ForceRefresh -and $script:CawtTclPath) {
        if ($script:CawtVerbose) {
            Write-Host "[CAWT] Using cached: $script:CawtTclPath" -ForegroundColor Cyan
        }
        return $script:CawtTclPath
    }
    
    if ($script:CawtVerbose) {
        Write-Host "[CAWT] Detecting Tcl interpreter..." -ForegroundColor Cyan
    }
    
    # 1. Читаем конфиг
    $config = Read-CawtConfig
    
    # Применяем verbose из конфига
    if ($config.verbose -eq 'true' -and -not $script:CawtVerbose) {
        $script:CawtVerbose = $true
    }
    
    # 2. Если в конфиге задан tclsh_path → проверяем только его
    if ($config.tclsh_path) {
        if ($script:CawtVerbose) {
            Write-Host "[CAWT] Config specifies: $($config.tclsh_path)" -ForegroundColor Yellow
        }
        
        # Проверяем существование файла
        if (-not (Test-Path $config.tclsh_path)) {
            Write-Warning "Config path does not exist: $($config.tclsh_path). Falling back to auto-detect."
        }
        # Проверяем наличие CAWT
        elseif (-not (Test-CawtCompatible -TclPath $config.tclsh_path)) {
            Write-Warning "Config path lacks CAWT: $($config.tclsh_path). Falling back to auto-detect."
        }
        # Всё ок — используем
        else {
            $info = Get-CawtVersions -TclPath $config.tclsh_path
            $script:CawtTclPath = $config.tclsh_path
            
            if ($script:CawtVerbose) {
                Write-Host "[CAWT] Found (config): Tcl=$($info.TclVersion) Twapi=$($info.TwapiVersion) CAWT=$($info.CawtVersion)" -ForegroundColor Green
            }
            
            return $script:CawtTclPath
        }
    }
    
    # 3. Сканируем кандидаты
    $candidates = Find-CawtCandidates
    
    if ($script:CawtVerbose) {
        Write-Host "[CAWT] Checking $($candidates.Count) candidates..." -ForegroundColor Cyan
    }
    
    foreach ($candidate in $candidates) {
        if ($script:CawtVerbose) {
            Write-Host "[CAWT] Testing: $candidate" -ForegroundColor Gray
        }
        
        if (Test-CawtCompatible -TclPath $candidate) {
            $info = Get-CawtVersions -TclPath $candidate
            $script:CawtTclPath = $candidate
            
            if ($script:CawtVerbose) {
                Write-Host "[CAWT] Found: Tcl=$($info.TclVersion) Twapi=$($info.TwapiVersion) CAWT=$($info.CawtVersion)" -ForegroundColor Green
            }
            
            return $script:CawtTclPath
        }
    }
    
    # 4. Ничего не найдено
    if ($script:CawtVerbose) {
        Write-Host "[CAWT] No CAWT-compatible Tcl found" -ForegroundColor Red
    }
    
    return $null
}

function Get-CawtInfo {
    param(
        [switch]$ForceRefresh
    )
    
    <#
    .SYNOPSIS
        Возвращает информацию о найденном Tcl-интерпретаторе.
    
    .DESCRIPTION
        Возвращает объект с информацией: Path, TclVersion, TwapiVersion, CawtVersion
    
    .EXAMPLE
        $info = Get-CawtInfo
        Write-Host "CAWT version: $($info.CawtVersion)"
    #>
    
    $path = Get-CawtTclPath -ForceRefresh:$ForceRefresh
    
    if (-not $path) {
        return $null
    }
    
    return Get-CawtVersions -TclPath $path
}

function Invoke-CawtScript {
    param(
        [string]$ScriptPath,
        [string[]]$Arguments,
        [switch]$ForceRefresh
    )
    
    <#
    .SYNOPSIS
        Запускает Tcl-скрипт через найденный CAWT-совместимый интерпретатор.
    
    .DESCRIPTION
        Автоматически находит Tcl-интерпретатор и запускает скрипт.
    
    .PARAMETER ScriptPath
        Путь к Tcl-скрипту.
    
    .PARAMETER Arguments
        Аргументы командной строки для скрипта.
    
    .EXAMPLE
        Invoke-CawtScript -ScriptPath "Excel-01_Create.tcl" -Arguments @("test.xlsx")
    #>
    
    $tclPath = Get-CawtTclPath -ForceRefresh:$ForceRefresh
    
    if (-not $tclPath) {
        throw "CAWT-compatible Tcl interpreter not found"
    }
    
    & $tclPath $ScriptPath @Arguments
}

# Примечание: функции доступны после импорта через Import-Module
# Export-ModuleMember работает только при загрузке как .psm1 модуль
