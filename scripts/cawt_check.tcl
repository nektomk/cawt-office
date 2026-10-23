# cawt_check.tcl - Verify the CAWT installation.
#
# Usage:
#   tclsh cawt_check.tcl
#
# Prints Tcl, Twapi, CAWT and sub-package versions. Exits 0 on success,
# 1 on a missing mandatory dependency.
#
# Set TCLLIBPATH or lappend auto_path if CAWT is not on the Tcl library path:
#   set env(TCLLIBPATH) "C:/path/to/Cawt"
#   lappend auto_path "C:/path/to/Cawt"

set retVal [catch {package require cawt} cawtVersion]
if { $retVal != 0 } {
    puts "ERROR: CAWT package not found. Install CAWT or set TCLLIBPATH/auto_path."
    puts "       See https://www.tcl3d.org/cawt/ or BAWT https://www.tcl3d.org/bawt/"
    exit 1
}

puts [format "%-22s : %s" "Tcl version" [info patchlevel]]
puts [format "%-22s : %s" "Tcl platform" $::tcl_platform(platform)]

set twapiVersion [Cawt GetPkgVersion "twapi"]
puts [format "%-22s : %s" "Twapi version" $twapiVersion]
if { $twapiVersion eq "" } {
    puts "ERROR: Twapi not available. CAWT needs Twapi >= 4.4 (>= 5 for Tcl 9)."
    exit 1
}

puts ""
puts [format "%-22s : %s" "CAWT version" $cawtVersion]

foreach pkg { cawtcore cawtoffice cawtexcel cawtword cawtppt cawtoutlook } {
    puts [format "%-22s : %s" $pkg [Cawt GetPkgVersion $pkg]]
}

puts ""
if { [Cawt HavePkg "cawtexcel"] } {
    puts "OK: Excel, Word, Ppt and Outlook automation packages loaded."
} else {
    puts "WARNING: Some sub-packages could not be loaded; only the loaded ones are usable."
}
exit 0
