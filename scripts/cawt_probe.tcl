# cawt_probe.tcl - List CAWT commands matching namespace patterns.
#
# Usage:
#   tclsh cawt_probe.tcl [pattern ...]
#
# Without arguments, probes the core lifecycle commands of each Office
# package. Useful to verify exact command names against the installed
# CAWT version before writing automation scripts.

package require cawt

set patterns $argv
if { [llength $patterns] == 0 } {
    set patterns {
        Excel::Open* Excel::Quit Excel::Save* Excel::Set* Excel::Get*Value* Excel::GetLastUsed*
        Word::Open* Word::Quit Word::Save* Word::Add* Word::SetCellValue Word::GetEndRange
        Ppt::Open* Ppt::Quit Ppt::Save* Ppt::Add* Ppt::Insert* Ppt::CloseAll
        Outlook::Open* Outlook::Quit Outlook::Create* Outlook::Send*
    }
}

foreach pat $patterns {
    puts "== $pat =="
    set cmds [lsort [info commands $pat]]
    if { [llength $cmds] == 0 } {
        puts "  (none)"
    }
    foreach cmd $cmds {
        puts "  $cmd"
    }
}
