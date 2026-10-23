# cawt-office

![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)
![MS Office](https://img.shields.io/badge/MS%20Office-Excel%20%C2%B7%20Word%20%C2%B7%20PowerPoint%20%C2%B7%20Outlook-D83B01)
![Tcl](https://img.shields.io/badge/Tcl-8.6%2B-506318)
![License](https://img.shields.io/badge/license-BSD--3--Clause-green)

> Скилл для [DeepSeek Harness](https://www.deepseek.com/) (DSH): управление **живыми,
> установленными** приложениями Microsoft Office на Windows через COM — с помощью
> Tcl-библиотеки [CAWT](https://www.tcl3d.org/cawt/).

## О скилле

`cawt-office` учит агента DSH автоматизировать настоящие приложения Office:
открывать окна Excel/Word/PowerPoint, читать и писать данные в уже запущенных
документах (включая несохранённые правки), строить таблицы с формулами и
оформлением, создавать документы и презентации, готовить письма Outlook.
Агент генерирует и запускает Tcl-скрипты через `tclsh`; готовые примеры из
каталога `scripts/` покрывают типовые сценарии.

### Чем это отличается от скиллов `office-docx` / `office-xlsx` / `office-pptx`

| | `office-*` (офлайн) | `cawt-office` (живое приложение) |
|---|---|---|
| Как работает | создаёт/читает **файлы** (Python, LibreOffice) | управляет **приложением** MS Office через COM |
| Требует MS Office | нет | да |
| Окно приложения | не показывается | можно открыть и передать пользователю |
| Живые данные | только то, что на диске | то, что сейчас в памяти приложения |
| Когда применять | массовая генерация файлов | интерактивная работа, макросы, печать, Outlook |

## Возможности

- **Excel** — создание и чтение книг; типизированная запись значений (текст /
  целые / вещественные); настоящие формулы; оформление (рамки, заливки, шрифты);
  диаграммы; сводные таблицы; экспорт в CSV; открытие видимого окна и передача
  файла пользователю; подключение к уже запущенному экземпляру и чтение живого
  (даже несохранённого) содержимого.
- **Word** — документы с заголовками, стилями и таблицами; конвертер
  Markdown → DOCX с полным оформлением (код-блоки, цитаты, списки); экспорт в PDF.
- **PowerPoint** — презентации, слайды, текстовые боксы, изображения, фигуры.
- **Outlook** — создание писем через рабочий профиль. **Без автоматической
  отправки:** любое письмо отправляется только после явного подтверждения
  пользователя.

## Требования

- Windows и установленный Microsoft Office (нужные приложения).
- Tcl/Tk 8.6+ — например, [ActiveTcl](https://www.activestate.com/products/tcl/)
  или [BAWT](https://www.tcl3d.org/bawt/); `tclsh` доступен в PATH.
- [Twapi](https://twapi.magicsplat.com/) 4.4+ (для Tcl 9 — версия 5+).
- [CAWT](https://www.tcl3d.org/cawt/download.html) 3.x; если `tclsh` не находит
  пакет, добавьте каталог CAWT в переменную окружения `TCLLIBPATH`.

Проверенная связка: ActiveTcl 8.6.13 + Twapi 5.2.0 + CAWT 3.1.1, локаль ru-RU —
все примеры из `scripts/` прошли end-to-end (Excel, Word, PowerPoint), включая
видимый режим, подключение к запущенному приложению и locale-тесты. Пример
Outlook не проверялся на живом почтовом профиле.

## Установка

1. Скопируйте каталог `cawt-office` в папку скиллов DSH:

   ```
   %USERPROFILE%\.agents\skills\cawt-office
   ```

2. Проверьте окружение:

   ```bat
   tclsh scripts\cawt_check.tcl
   ```

   Скрипт напечатает версии Tcl, Twapi, CAWT и подпакетов и завершится с кодом 0.

3. Готово: агент DSH распознаёт скилл и при задачах вроде «открой мой Excel и
   покажи, что изменилось» сам напишет и выполнит нужный Tcl-скрипт.

## Использование

### Через агента (основной режим)

Просто попросите — примеры формулировок:

- «Создай книгу Excel с таблицей продаж и totals-формулами, покажи результат на экране».
- «Подключись к открытому Excel и прочитай текущие значения листа».
- «Сделай из README.md красиво оформленный DOCX и открой в Word».
- «Собери презентацию из 5 слайдов с картинками».
- «Подготовь письмо в Outlook (не отправляй)».

Агент опирается на `SKILL.md` — в нём зафиксированы точные сигнатуры команд CAWT,
правила жизненного цикла приложений и особенности локалей (числа, даты, формулы,
CSV).

### Напрямую (готовые примеры)

```bat
tclsh scripts\Excel-01_Create.tcl      report.xlsx
tclsh scripts\Excel-02_Read.tcl        report.xlsx
tclsh scripts\Excel-03_OpenVisible.tcl report.xlsx
tclsh scripts\Excel-04_AttachRead.tcl  report.xlsx Sales
tclsh scripts\Word-01_Create.tcl       report.docx
tclsh scripts\md2docx.tcl              README.md out.docx
tclsh scripts\Ppt-01_Create.tcl        deck.ppt image.png
tclsh scripts\Outlook-01_Mail.tcl      user@example.com "Тема" "Тело письма"
tclsh scripts\locale_test.tcl
```

| Скрипт | Назначение |
|---|---|
| `cawt_check.tcl` | диагностика установки (tclsh, Twapi, CAWT, подпакеты) |
| `cawt_probe.tcl` | интроспекция: точные имена команд установленной CAWT |
| `Excel-01_Create.tcl` | создать книгу, заполнить таблицу, сохранить |
| `Excel-02_Read.tcl` | открыть книгу read-only и прочитать значения |
| `Excel-03_OpenVisible.tcl` | открыть книгу в **видимом** окне Excel и оставить пользователю |
| `Excel-04_AttachRead.tcl` | подключиться к запущенному Excel и прочитать живое содержимое |
| `locale_test.tcl` | тест локаль-чувствительности (числа, даты, формулы, CSV) |
| `Word-01_Create.tcl` | создать документ: текст + таблица, сохранить |
| `md2docx.tcl` | конвертер Markdown → DOCX с оформлением, в видимом Word |
| `Ppt-01_Create.tcl` | презентация: слайды, текстовые боксы, изображение |
| `Outlook-01_Mail.tcl` | создать письмо (без отправки) |

Примеры создают файлы в текущем каталоге (или по переданному пути). Невидимые
экземпляры Office скрипты закрывают сами; «видимые» (`Excel-03`, `Excel-04`,
`md2docx`) специально оставляют окно открытым — файл передаётся пользователю.

## Структура каталога

```
cawt-office/
├── SKILL.md                 # инструкции для агента (ядро скилла)
├── README.md                # этот файл
└── scripts/                 # готовые примеры и диагностика
    ├── cawt_check.tcl
    ├── cawt_probe.tcl
    ├── Excel-01_Create.tcl
    ├── Excel-02_Read.tcl
    ├── Excel-03_OpenVisible.tcl
    ├── Excel-04_AttachRead.tcl
    ├── locale_test.tcl
    ├── Word-01_Create.tcl
    ├── md2docx.tcl
    ├── Ppt-01_Create.tcl
    └── Outlook-01_Mail.tcl
```

## Важные замечания

- `Cawt Destroy $appId` освобождает COM-обработчик, но **не закрывает** окно
  приложения; закрывает — `<App> Quit $appId`. Видимое окно с открытым файлом
  переживает завершение скрипта — так файл «передаётся» пользователю.
- PowerPoint однопоточен/одноэкземплярен: `Ppt Open` подключается к уже
  запущенному приложению пользователя — закрывайте только то, что открыли сами.
- Если Excel занят (редактирование ячейки, модальный диалог), COM-вызовы
  отклоняются (`RPC_E_CALL_REJECTED`) — скрипты повторяют попытки; обычно
  достаточно нажать Esc в окне Excel.
- Региональные настройки влияют на форматы: числа пишите с явным типом
  (`fmt int|real`), коды форматов дат локальны, формулы через COM — в синтаксисе
  US. Полный перечень проверенных правил — раздел «Locale and data types» в
  `SKILL.md`.

## Лицензии

- **cawt-office** (скрипты и документация скилла) — [BSD 3-Clause](https://opensource.org/license/bsd-3-clause).
- **[Tcl/Tk](https://www.tcl-lang.org/)** — свободная BSD-подобная лицензия
  ([текст лицензии](https://www.tcl-lang.org/software/tcltk/license.html));
  интерпретатор `tclsh` — среда выполнения скриптов скилла.
- **[Twapi](https://twapi.magicsplat.com/)** — BSD-лицензия
  ([исходный код](https://github.com/apnadkarni/twapi)); прослойка между Tcl и
  Windows COM.
- **[CAWT](https://www.tcl3d.org/cawt/)** — BSD-лицензия, автор Paul Obermeier;
  библиотека высокоуровневых команд для MS Office, на которой построен скилл.

Скилл не модифицирует перечисленные компоненты и использует их как внешние
зависимости. Microsoft Office, Excel, Word, PowerPoint и Outlook — товарные
знаки Microsoft Corporation.

## Благодарности

Этот скилл создан в агентной среде **DeepSeek Harness**
([deepseek.com](https://www.deepseek.com/)) — от идеи и архитектуры до
end-to-end-тестирования на живых приложениях Office.

В разработке участвовали модели:

- **deepseek-v4-flash** — компания [DeepSeek](https://www.deepseek.com/);
- **Qwen3.8-Max** — команда [Qwen](https://qwen.ai/) (Alibaba Group);
- доступ к моделям обеспечивал провайдер **[PlusVibe](https://plusvibeapi.ru/)** —
  благодарим за стабильный и быстрый API.

Отдельное спасибо **Paul Obermeier** за библиотеку
[CAWT](https://www.tcl3d.org/cawt/) — фундамент, на котором держится весь скилл,
и сообществу [Tcl](https://www.tcl-lang.org/) за тридцать лет живого языка.
