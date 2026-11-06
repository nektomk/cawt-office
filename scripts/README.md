# cawt-office — автоматический детект Tcl-интерпретатора

## Быстрый старт

Скилл `cawt-office` автоматически находит CAWT-совместимый Tcl-интерпретатор при первом запуске в сессии PowerShell.

### Использование в PowerShell

```powershell
# Импортировать модуль
Import-Module .\cawt-office\scripts\cawt-office.ps1

# Получить путь к Tcl (кэшируется в сессии)
$tcl = Get-CawtTclPath

# Запустить скрипт
& $tcl Excel-01_Create.tcl test.xlsx
```

### Прямой вызов детекта

```powershell
# Простой вывод пути
.\cawt-detect.ps1

# Подробный режим
.\cawt-detect.ps1 -Verbose

# Принудительное пересканирование
.\cawt-detect.ps1 -ForceRefresh
```

## Алгоритм детекта

1. **Кэш сессии** — если уже найден в этой сессии PowerShell
2. **Конфиг** — файл `.cawt-office` в корне скилла (если существует)
3. **PATH** — `where.exe tclsh*.exe`
4. **Реестр Windows** — ключи установки Tcl
5. **Известные пути** — ActiveTcl, BAWT, MagicSplat, MSYS2, Cygwin, portable kits

Для каждого кандидата проверяется наличие пакета CAWT.

## Конфигурация (опционально)

Создайте файл `.cawt-office` в корне скилла (`cawt-office\.cawt-office`):

```
# Путь к Tcl-интерпретатору
tclsh_path=R:\tclwinkit-msvc.exe

# Подробный режим
verbose=true
```

**Важно:** Файл создаётся вручную пользователем. Один файл на систему.

## Если интерпретатор не найден

```
CAWT-compatible Tcl interpreter not found.

Download portable Tcl 9 Kit with bundled CAWT:
  https://github.com/nektomk/tclwinkit/releases

Or create config file: cawt-office\.cawt-office
  tclsh_path=R:\tclwinkit-msvc.exe
```

## Файлы

| Файл | Назначение |
|------|------------|
| `cawt-office.ps1` | Модуль PowerShell: функции детекта + кэш в памяти |
| `cawt-detect.ps1` | CLI: прямой вызов детекта |
| `cawt-check.tcl` | Быстрая проверка CAWT (вызывается из PowerShell) |

## Функции PowerShell модуля

| Функция | Описание |
|---------|----------|
| `Get-CawtTclPath [-ForceRefresh]` | Вернуть путь к Tcl (кэш или детект) |
| `Get-CawtInfo [-ForceRefresh]` | Вернуть информацию: Path, TclVersion, TwapiVersion, CawtVersion |
| `Invoke-CawtScript -ScriptPath <path> [-Arguments <args>]` | Запустить Tcl-скрипт |

## Примеры

### Получить информацию об окружении

```powershell
Import-Module .\cawt-office\scripts\cawt-office.ps1
$info = Get-CawtInfo
Write-Host "Tcl: $($info.TclVersion), CAWT: $($info.CawtVersion)"
```

### Запустить скрипт с аргументами

```powershell
Invoke-CawtScript -ScriptPath "Excel-01_Create.tcl" -Arguments @("test.xlsx")
```

### Детект с отладочной информацией

```powershell
.\cawt-detect.ps1 -Verbose

# Вывод:
# [CAWT] Detecting Tcl interpreter...
# [CAWT] Checking 5 candidates...
# [CAWT] Testing: C:\ATcl\bin\tclsh.exe
# [CAWT] Found: Tcl=8.6 Twapi=5.2.0 CAWT=3.1.1
# C:\ATcl\bin\tclsh.exe
```
