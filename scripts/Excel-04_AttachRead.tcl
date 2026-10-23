# Excel-04_AttachRead.tcl - Attach to the RUNNING Excel instance and read the
# live (possibly unsaved) content of an open workbook.
#
# Usage:
#   tclsh Excel-04_AttachRead.tcl <open-workbook-name> [sheetName]
# Example:
#   tclsh Excel-04_AttachRead.tcl report.xlsx Sales
#
# Read-only: never closes, saves or quits anything; the user's Excel is left
# exactly as it was. While Excel is "busy" (cursor in cell-edit mode, modal
# dialog open), COM calls are rejected with RPC_E_CALL_REJECTED ("Call was
# rejected") - the script retries automatically.

set targetFile [lindex $argv 0]
set sheetName  [lindex $argv 1]
if { $targetFile eq "" } {
    puts "Usage: tclsh Excel-04_AttachRead.tcl <open-workbook-name> \[sheetName\]"
    exit 1
}

# Retry a body while COM calls are rejected by a busy Excel.
# Errors prefixed with "NOBOOK " are not retryable.
proc tryBusy { body {attempts 8} {delayMs 2000} } {
    for { set i 1 } { $i <= $attempts } { incr i } {
        if { [catch { uplevel 1 $body } res] } {
            if { [string match "NOBOOK *" $res] } {
                return -code error [string range $res 7 end]
            }
            if { $i >= $attempts } {
                return -code error $res
            }
            puts "Excel busy (attempt $i/$attempts): $res"
            puts "  -> retrying in [expr { $delayMs / 1000 }] s (press Esc in Excel if the cursor is in a cell)..."
            after $delayMs
        } else {
            return $res
        }
    }
}

proc readLive {} {
    global targetFile sheetName

    # "Excel Open true" attaches to the already running instance if one exists.
    set appId [Excel Open true]

    # Locate the workbook among the open ones (match by tail or full path).
    set workbookId ""
    set foundNames {}
    set wbs [$appId Workbooks]
    set n [$wbs Count]
    for { set i 1 } { $i <= $n } { incr i } {
        set wb [$wbs Item $i]
        set full [$wb FullName]
        lappend foundNames [file tail $full]
        if { [string equal -nocase [file tail $full] $targetFile]
             || [string equal -nocase $full $targetFile] } {
            set workbookId $wb
            break
        }
        Cawt Destroy $wb
    }
    Cawt Destroy $wbs
    if { $workbookId eq "" } {
        error "NOBOOK \"$targetFile\" is not open in Excel. Open workbooks: $foundNames"
    }

    # Requested sheet, or the active one.
    if { $sheetName ne "" } {
        if { [catch { Excel GetWorksheetIdByName $workbookId $sheetName } worksheetId] } {
            error "NOBOOK no worksheet \"$sheetName\" in $targetFile"
        }
    } else {
        set worksheetId [$workbookId ActiveSheet]
    }

    set numRows [Excel GetLastUsedRow $worksheetId]
    set numCols [Excel GetLastUsedColumn $worksheetId]
    set matrix  [Excel GetMatrixValues $worksheetId 1 1 $numRows $numCols]

    set out {}
    lappend out "Workbook  : [$workbookId FullName]"
    lappend out "Worksheet : [$worksheetId Name]"
    lappend out "Used range: $numRows row(s) x $numCols column(s)"
    lappend out "Saved     : [Cawt TclBool [$workbookId Saved]]  (false = unsaved in-memory edits)"
    lappend out ""
    foreach row $matrix {
        lappend out [join $row "\t"]
    }
    return [join $out "\n"]
}

if { [catch {
    package require cawt
    puts [tryBusy { readLive }]
} err] } {
    puts "ERROR: $err"
    puts "(If Excel rejected the calls: press Esc in Excel to leave cell-edit mode, close any dialog, then rerun.)"
    catch { Cawt Destroy }
    exit 1
}

# Release COM handlers only - the user's Excel stays open and untouched.
Cawt Destroy
