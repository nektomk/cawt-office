# Ppt-01_Create.tcl - Create a PowerPoint presentation with slides, text
# boxes, shapes and an image, save it.
#
# Usage:
#   tclsh Ppt-01_Create.tcl <output.pptx> [<imageFile>|-] [overwrite]
#   Pass '-' (or omit) for no image; '-' keeps [overwrite] reachable because
#   shells like PowerShell drop empty args.
#   overwrite = 1/true/yes/force permits replacing an existing file;
#   without it an existing <output.pptx> is refused with exit code 2 (S10).
#
# Requires: CAWT + Twapi, MS PowerPoint installed.

set outFile [lindex $argv 0]
if { $outFile eq "" } {
    set outFile [file join [pwd] "presentation.pptx"]
}
set imgFile [lindex $argv 1]
# '-' is the alias for "no image" (shells drop empty args).
if { $imgFile eq "-" } { set imgFile "" }

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

# S5 hardening. Unlike Excel/Word, "Ppt OpenNew" does NOT take a
# useExistingFirst flag (its first argument is the slide width), and
# PowerPoint is single-instance at the COM level - "Ppt Open" attaches to a
# running PowerPoint if one exists. Ownership heuristic: count presentations
# BEFORE AddPres. Count > 0 => we attached to the user's app => guest mode:
# close only our own presentation, never CloseAll/Quit. Count == 0 => treat
# the instance as created by this script and quit it (PowerPoint normally
# exits together with its last presentation, so a running empty instance is
# a stray case; quitting it is reported below).
set appId [Ppt Open]
set numPresBefore [$appId -with { Presentations } Count]
set ownInstance [expr { $numPresBefore == 0 }]

set presId ""
set presClosed false
set rc [catch {
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

    # Close ONLY our own presentation (saved above - no prompt).
    Ppt Close $presId
    set presClosed true
} err]

# Error path: drop our half-built presentation, still never CloseAll.
if { ! $presClosed && $presId ne "" } { catch { Ppt Close $presId } }
# Quit only an instance this script created.
if { $ownInstance } { catch { Ppt Quit $appId } }
Cawt Destroy

if { $rc } {
    puts "ERROR: $err"
    exit 1
}
if { $ownInstance } {
    puts "Quit the PowerPoint instance created by this script."
} else {
    puts "Guest mode: attached to the running PowerPoint ($numPresBefore pre-existing presentation(s)); closed only our own, app left running."
}
puts "Done."
