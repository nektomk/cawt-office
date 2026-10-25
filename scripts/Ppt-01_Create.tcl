# Ppt-01_Create.tcl - Create a PowerPoint presentation with slides, text
# boxes, shapes and an image, save it.
#
# Usage:
#   tclsh Ppt-01_Create.tcl <output.pptx> [<imageFile>] [overwrite]
#   overwrite = 1/true/yes/force permits replacing an existing file;
#   without it an existing <output.pptx> is refused with exit code 2 (S10).
#
# Requires: CAWT + Twapi, MS PowerPoint installed.

set outFile [lindex $argv 0]
if { $outFile eq "" } {
    set outFile [file join [pwd] "presentation.pptx"]
}
set imgFile [lindex $argv 1]

# S10 overwrite consent: never silently replace an existing file.
# Checked BEFORE any COM startup: refusal costs no Office instance.
set owRaw [lindex $argv 2]
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

# Note: unlike Excel/Word, "Ppt OpenNew" does NOT take a useExistingFirst flag
# (its first argument is the slide width). Use plain "Ppt Open" as in the
# official CAWT tests. PowerPoint is a single-instance application: if the
# user already has PowerPoint open, this attaches to it - and the "Ppt Quit"
# at the end would close it. Remove the Quit line if that is not acceptable.
set appId [Ppt Open]

# A new, empty presentation.
set presId [Ppt AddPres $appId]

# Slide 1: title text box.
set slideId [Ppt AddSlide $presId]
set tbox [Ppt AddTextbox $slideId 1c 1c 20c 5c]
Ppt AddTextboxText $tbox "Quarterly review"

# Slide 2: bullets.
set slideId [Ppt AddSlide $presId]
set tbox [Ppt AddTextbox $slideId 1c 1c 20c 10c]
foreach line { "Revenue is up 12%" "Three new accounts closed" "Next steps: scale support" } {
    Ppt AddTextboxText $tbox $line true
}
Cawt Destroy $tbox

# Optional image on slide 2.
if { $imgFile ne "" } {
    Ppt InsertImage $slideId $imgFile 1c 12c 8c 8c
}
Cawt Destroy $slideId

# Save using the application's default extension if none given.
if { [string tolower [file extension $outFile]] eq "" } {
    append outFile [Ppt GetExtString $appId]
}

Ppt SaveAs $presId $outFile
puts "Saved: $outFile"

# Close all presentations and quit the invisible instance we created.
Ppt CloseAll $appId
Ppt Quit $appId

Cawt Destroy
puts "Done."
