# Excel-02_Read.tcl - Open an existing Excel workbook read-only and print
# the used table values.
#
# Usage:
#   tclsh Excel-02_Read.tcl <input.xlsx> [sheetName|""|-] [maxRows]
#   maxRows caps printed rows (default 200; 0 or "all" = no cap).
#   Truncation is reported with the true row count; only the shown rows
#   are transferred over COM.
#   "-" may be passed instead of an empty sheetName (some shells, e.g.
#   PowerShell, drop empty arguments).
#
# Requires: CAWT + Twapi, MS Excel installed.

set inFile [lindex $argv 0]
if { $inFile eq "" } {
    puts "Usage: tclsh Excel-02_Read.tcl <input.xlsx> \[sheetName|\"\"|-\] \[maxRows]"
    exit 1
}
set sheetName [lindex $argv 1]
if { $sheetName eq "-" } { set sheetName "" }

# Output row cap: default 200 protects the agent context on large sheets;
# 0 or "all" disables the cap. Pass "" (or "-") as sheetName to reach maxRows.
set maxRows [lindex $argv 2]
if { $maxRows eq "" } { set maxRows 200 }
if { [string equal -nocase $maxRows "all"] } { set maxRows 0 }
if { ! [string is integer -strict $maxRows] || $maxRows < 0 } {
    puts "ERROR: maxRows must be a non-negative integer or 0/\"all\", got: $maxRows"
    exit 1
}

package require cawt

set appId [Excel OpenNew false]
catch { $appId DisplayAlerts false }

# Everything after OpenNew runs guarded: on ANY error the hidden instance is
# still quit - a failed run must never leave an invisible Excel behind.
# (No bare "return" inside the catch body: its last command value is the
# result; a bare return would surface as rc=2/TCL_RETURN - see SKILL traps.)
if { [catch {
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

    # Print the used matrix (capped at maxRows; fetch only what will be shown).
    set readRows $numRows
    if { $maxRows > 0 && $numRows > $maxRows } { set readRows $maxRows }
    foreach row [Excel GetMatrixValues $worksheetId 1 1 $readRows $numCols] {
        puts [join $row "\t"]
    }
    if { $readRows < $numRows } {
        puts "...truncated: shown $readRows of $numRows rows (pass maxRows=N or 0/all for more)"
    }

    Excel Close $workbookId
} err] } {
    catch { Excel Quit $appId }
    catch { Cawt Destroy }
    puts "ERROR: $err"
    exit 1
}

# Quit the invisible instance we created.
Excel Quit $appId

Cawt Destroy
puts "Done."
