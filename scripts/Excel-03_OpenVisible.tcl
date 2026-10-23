# Excel-03_OpenVisible.tcl - Open a workbook in a VISIBLE Excel window and
# leave it open for the user after the script exits.
#
# Usage:
#   tclsh Excel-03_OpenVisible.tcl <file.xlsx>
#
# Requires: CAWT + Twapi, MS Excel installed.
#
# Note: unlike the other examples, this script intentionally does NOT quit
# Excel - a visible instance with an open workbook keeps running after the
# Tcl interpreter exits (same as the interactive mode of the official CAWT
# tests). Close the window normally in Excel when done.

set inFile [lindex $argv 0]
if { $inFile eq "" } {
    puts "Usage: tclsh Excel-03_OpenVisible.tcl <file.xlsx>"
    exit 1
}
if { ! [file exists $inFile] } {
    puts "ERROR: file not found: $inFile"
    exit 1
}
set inFile [file nativename [file normalize $inFile]]

package require cawt

# "Excel Open true" is the high-level call that shows the application window
# (see official test Excel-16_SetGet.tcl). Set Visible explicitly as well,
# so the window is shown regardless of the Open argument semantics.
set appId [Excel Open true]
catch { $appId Visible true }

set workbookId [Excel OpenWorkbook $appId $inFile]

# Maximize the window for convenience (ignore failure if unsupported).
catch { $appId -with { ActiveWindow } WindowState $::Excel::xlMaximized }

puts "Opened in visible Excel: $inFile"
puts "The Excel window stays open after this script exits."

# Release COM handlers; the visible Excel keeps running on its own.
Cawt Destroy
