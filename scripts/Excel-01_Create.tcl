# Excel-01_Create.tcl - Create an Excel workbook, fill cells, save it.
#
# Usage:
#   tclsh Excel-01_Create.tcl <output.xlsx> [overwrite]
#   overwrite = 1/true/yes/force permits replacing an existing file;
#   without it an existing <output.xlsx> is refused with exit code 2 (S10).
#
# Requires: CAWT + Twapi, MS Excel installed.

set outFile [lindex $argv 0]
if { $outFile eq "" } {
    set outFile [file join [pwd] "report.xlsx"]
}

# S10 overwrite consent: never silently replace an existing file.
set owRaw [lindex $argv 1]
set overwrite [expr {[string tolower $owRaw] in {1 true yes overwrite force}}]
if { $owRaw ne "" && ! $overwrite } {
    puts "ERROR: bad overwrite value '$owRaw' (use 1)"
    exit 1
}
if { ! $overwrite && [file exists $outFile] } {
    puts "ERROR: $outFile already exists - rerun with overwrite arg 1 or choose another name"
    exit 2
}

package require cawt

# Prefer an already running Excel if one is open; "false" overrides that.
set appId [Excel OpenNew false]

# A new workbook with one worksheet.
set workbookId [Excel AddWorkbook $appId]
set worksheetId [Excel GetWorksheetIdByIndex $workbookId 1]
Excel SetWorksheetName $worksheetId "Sales"

# Fill a small table.
set rows { {"Product" "Q1" "Q2" "Q3"}
          {"Apples" 100 120 95}
          {"Pears"  80 90 110} }
for { set r 1 } { $r <= [llength $rows] } { incr r } {
    set rowValues [lindex $rows [expr {$r - 1}]]
    Excel SetRowValues $worksheetId $r $rowValues
}

# Totals column: header plus per-row sums computed in Tcl.
# Note the explicit fmt "int": SetCellValue defaults to fmt "text", which
# stores a TEXT cell (NumberFormat "@") that cannot be used in formulas.
set col [expr { [llength [lindex $rows 0]] + 1 }]
Excel SetCellValue $worksheetId 1 $col "Total"
for { set r 2 } { $r <= [llength $rows] } { incr r } {
    set sum 0
    foreach v [lrange [lindex $rows [expr {$r - 1}]] 1 end] {
        incr sum $v
    }
    Excel SetCellValue $worksheetId $r $col $sum int
}

# Save using the application's default extension if none given.
if { [string tolower [file extension $outFile]] eq "" } {
    append outFile [Excel GetExtString $appId]
}

Excel SaveAs $workbookId $outFile
puts "Saved: $outFile"

# Close the workbook and quit the invisible instance we created.
Excel Close $workbookId
Excel Quit $appId

Cawt Destroy
puts "Done."
