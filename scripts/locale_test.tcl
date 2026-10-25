# locale_test.tcl - Locale-sensitivity test suite for CAWT+Twapi+Excel
# (written and verified on ru-RU; rerun after installing on another locale).
# Verifies: fmt real/int/display, locale-aware NumberFormat via GetNumberFormat,
# dates via Cawt IsoDateToOfficeDate + locale format codes, Formula vs
# FormulaLocal, SetRowValues real numbers, CSV export encoding (cp1251 on ru)
# and separators. Creates its own invisible instance; report in UTF-8.
# Source is ASCII-only (Cyrillic via \uXXXX): tclsh reads sources in the
# system encoding (cp1251 on ru-RU Windows), not UTF-8.
#
# Usage:
#   tclsh locale_test.tcl
#   -> writes test-output/locale_report.txt next to this script

set outDir     [file join [file dirname [info script]] "test-output"]
file mkdir $outDir
set reportPath [file join $outDir "locale_report.txt"]
set xlsxPath   [file nativename [file join $outDir "locale_test.xlsx"]]
set csvPath    [file nativename [file join $outDir "locale_test.csv"]]

# \u0414\u0414.\u041c\u041c.\u0413\u0413\u0413\u0413            = DD.MM.YYYY (ru codes)
# \u0414\u0414.\u041c\u041c.\u0413\u0413\u0413\u0413 \u0427\u0427:\u041c\u041c = + HH:MM (ru codes)
# \u041f\u0440\u0438\u0432\u0435\u0442 \u043c\u0438\u0440      = "Privet mir" (Hello world)
set fmtDateRU "\u0414\u0414.\u041c\u041c.\u0413\u0413\u0413\u0413"
set fmtDtRU   "\u0414\u0414.\u041c\u041c.\u0413\u0413\u0413\u0413 \u0427\u0427:\u041c\u041c"
set cyrText   "\u041f\u0440\u0438\u0432\u0435\u0442 \u043c\u0438\u0440"

set log {}
set appId ""

proc L { s } {
    global log
    lappend log $s
}

proc probe { label body } {
    global log
    if { [catch { uplevel 1 $body } res] } {
        lappend log [format "%-34s ERROR: %s" $label $res]
    } else {
        lappend log [format "%-34s %s" $label $res]
    }
}

proc getCell { ws r c } {
    return [Excel GetCellIdByIndex $ws $r $c]
}

proc dumpCell { ws r c label } {
    global log
    if { [catch {
        set cellId [getCell $ws $r $c]
        set parts {}
        catch { lappend parts "Value=[$cellId Value]" }
        catch { lappend parts "Text=[$cellId Text]" }
        catch { lappend parts "NumFmt=[$cellId NumberFormat]" }
        catch { set f [$cellId Formula]; if { $f ne "" } { lappend parts "Formula=$f" } }
        catch { set fl [$cellId FormulaLocal]; if { $fl ne "" } { lappend parts "FormulaLocal=$fl" } }
        catch { Cawt Destroy $cellId }
        set line [join $parts "  |  "]
    } err] } {
        set line "DUMP-ERROR: $err"
    }
    lappend log [format "%-34s %s" $label $line]
}

set rc [catch {
    package require cawt

    L "=== 0. Locale info via CAWT ==="
    set appId [Excel OpenNew false]
    probe "GetDecimalSeparator"     { Excel GetDecimalSeparator $appId }
    probe "GetThousandsSeparator"   { Excel GetThousandsSeparator $appId }
    probe "GetListItemSeparator"    { Excel GetListItemSeparator }
    probe "GetNumberFormat 0/00"    { Excel GetNumberFormat $appId "0" "00" }
    probe "IsoDateToOfficeDate"     { Cawt IsoDateToOfficeDate "2026-10-23 00:00:00" }
    probe "OfficeDateToIsoDate"     { Cawt OfficeDateToIsoDate 46318 }
    probe "Seconds->Office->Iso rt" { Cawt OfficeDateToIsoDate [Cawt SecondsToOfficeDate [Cawt IsoDateToSeconds "2026-10-23 00:00:00"]] }

    probe "DisplayAlerts=false"     { $appId DisplayAlerts false; format ok }
    set wb [Excel AddWorkbook $appId]
    set ws [Excel GetWorksheetIdByIndex $wb 1]

    L ""
    L "=== 1. Typed writes: fmt real / int / custom subFmt ==="
    probe "A1 real 3.14159 (dflt subFmt)" { Excel SetCellValue $ws 1 1 3.14159 real; format ok }
    probe "A2 real, subFmt 0/0000"        { Excel SetCellValue $ws 2 1 3.14159 real [Excel GetNumberFormat $appId "0" "0000"]; format ok }
    set serial [Cawt IsoDateToOfficeDate "2026-10-23 00:00:00"]
    probe "A3 int serial (General)"       { Excel SetCellValue $ws 3 1 $serial int; format ok }
    probe "A4 real serial fmt RU date"    { Excel SetCellValue $ws 4 1 $serial real $::fmtDateRU; format ok }
    set dtSerial [expr { $serial + 14.5 / 24.0 }]
    probe "A5 real datetime fmt RU"       { Excel SetCellValue $ws 5 1 $dtSerial real $::fmtDtRU; format ok }
    probe "A6 real serial fmt US codes"   { Excel SetCellValue $ws 6 1 $serial real "DD.MM.YYYY"; format ok }
    probe "A7 raw Value2 date string"     { set c [getCell $ws 7 1]; $c Value2 "23.10.2026"; Cawt Destroy $c; format ok }
    probe "A8 SetRowValues {100 200 300}" { Excel SetRowValues $ws 8 {100 200 300}; format ok }
    probe "D8 formula =SUM(A8:C8)"        { set c [getCell $ws 8 4]; $c Formula "=SUM(A8:C8)"; Cawt Destroy $c; format ok }
    probe "A9 cyrillic text"              { Excel SetCellValue $ws 9 1 $::cyrText; format ok }

    dumpCell $ws 1 1 "A1 real dflt"
    dumpCell $ws 2 1 "A2 real 4 digits"
    dumpCell $ws 3 1 "A3 int serial"
    dumpCell $ws 4 1 "A4 date RU fmt"
    dumpCell $ws 5 1 "A5 datetime RU fmt"
    dumpCell $ws 6 1 "A6 date US codes(!)"
    dumpCell $ws 7 1 "A7 raw Value2 string"
    dumpCell $ws 8 1 "A8 SetRowValues"
    dumpCell $ws 8 4 "D8 SUM formula"
    dumpCell $ws 9 1 "A9 cyrillic"

    L ""
    L "=== 2. Typed reads of A1 (stored real 3.14159, fmt 0,00) ==="
    probe "GetCellValue text"    { Excel GetCellValue $ws 1 1 }
    probe "GetCellValue display" { Excel GetCellValue $ws 1 1 display }
    probe "GetCellValue int"     { Excel GetCellValue $ws 1 1 int }
    probe "GetCellValue real"    { Excel GetCellValue $ws 1 1 real }
    probe "A5 read display"      { Excel GetCellValue $ws 5 1 display }
    probe "row8 GetMatrixValues" { Excel GetMatrixValues $ws 8 1 8 4 }

    L ""
    L "=== 3. Save xlsx + CSV (xlCSV=6), then quit ==="
    probe "SaveAs xlsx" { Excel SaveAs $wb $::xlsxPath; format ok }
    probe "SaveAs csv"  { $wb SaveAs $::csvPath 6; format ok }
    probe "Close+Quit"  { Excel Close $wb; Excel Quit $appId; format ok }
} fatalErr]

if { $rc } {
    L ""
    L "FATAL: $fatalErr"
    catch { Excel Quit $appId }
}
catch { Cawt Destroy }

L ""
L "=== 4. CSV bytes (expect cp1251 + ';' separators) ==="
if { [catch {
    set fh [open $csvPath rb]
    set raw [read $fh]
    close $fh
    L "hex   : [binary encode hex $raw]"
    L "cp1251 decoded:"
    foreach line [split [encoding convertfrom cp1251 $raw] "\r\n"] {
        if { $line ne "" } { L "  csv: $line" }
    }
    L "utf8 would be  : [string map {\n {} \r {}} [encoding convertfrom utf-8 $raw]]"
} err] } {
    L "csv read ERROR: $err"
}

set report [join $log "\n"]
set fh [open $reportPath w]
fconfigure $fh -encoding utf-8
puts $fh $report
close $fh
puts "locale test done; report: $reportPath"
if { $rc } {
    puts "NOTE: fatal error recorded in report"
    exit 1
}
