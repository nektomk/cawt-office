# Excel-02_Read.tcl - Open an existing Excel workbook read-only and print
# the used table values.
#
# Usage:
#   tclsh Excel-02_Read.tcl <input.xlsx> [sheetName]
#
# Requires: CAWT + Twapi, MS Excel installed.

set inFile [lindex $argv 0]
if { $inFile eq "" } {
    puts "Usage: tclsh Excel-02_Read.tcl <input.xlsx> [sheetName]"
    exit 1
}
set sheetName [lindex $argv 1]

package require cawt

set appId [Excel OpenNew false]
set workbookId [Excel OpenWorkbook $appId $inFile -readonly true]

# Select worksheet by name (defaults to the first one).
if { $sheetName ne "" } {
    set worksheetId [Excel GetWorksheetIdByName $workbookId $sheetName]
} else {
    set worksheetId [Excel GetWorksheetIdByIndex $workbookId 1]
}

puts "Worksheet: [Excel GetWorksheetName $worksheetId]"
set numRows [Excel GetLastUsedRow $worksheetId]
set numCols [Excel GetLastUsedColumn $worksheetId]
puts "Used range: $numRows row(s) x $numCols column(s)"
puts ""

# Print the used matrix.
foreach row [Excel GetMatrixValues $worksheetId 1 1 $numRows $numCols] {
    puts [join $row "\t"]
}

Excel Close $workbookId
Excel Quit $appId

Cawt Destroy
puts "Done."
