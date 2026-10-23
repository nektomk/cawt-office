# Word-01_Create.tcl - Create a Word document, add text and a table, save it.
#
# Usage:
#   tclsh Word-01_Create.tcl <output.docx>
#
# Requires: CAWT + Twapi, MS Word installed.

set outFile [lindex $argv 0]
if { $outFile eq "" } {
    set outFile [file join [pwd] "report.docx"]
}

package require cawt

set appId [Word OpenNew false]
set docId [Word AddDocument $appId]

# Title paragraph.
# Note: Twapi 5 is strict about property types - Font.Bold is VT_I4 and needs
# an integer (-1 = wdToggleTrue / msoTrue, 0 = off). "Cawt TclBool" converts
# COM booleans TO Tcl, not the other way around, so do not use it here.
set titleRange [Word AppendText $docId "Quarterly report" true]
$titleRange -with { Font } Name "Arial"
$titleRange -with { Font } Size 18
$titleRange -with { Font } Bold -1
Cawt Destroy $titleRange

# Body text.
set bodyRange [Word AppendText $docId "Sales grew by 12% compared to last quarter." true]

# Add a small table.
set tableRange [Word GetEndRange $docId]
set tableId [Word AddTable $tableRange 3 2]
Cawt Destroy $tableRange
Word SetCellValue $tableId 1 1 "Region"
Word SetCellValue $tableId 1 2 "Revenue"
Word SetCellValue $tableId 2 1 "West"
Word SetCellValue $tableId 2 2 "120k"
Word SetCellValue $tableId 3 1 "East"
Word SetCellValue $tableId 3 2 "180k"

# Save using the application's default extension if none given.
if { [string tolower [file extension $outFile]] eq "" } {
    append outFile [Word GetExtString $appId]
}

Word SaveAs $docId $outFile
puts "Saved: $outFile"

# Close the document and quit the invisible instance we created.
Word Close $docId
Word Quit $appId

Cawt Destroy
puts "Done."
