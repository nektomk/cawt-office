# Outlook-01_Mail.tcl - Create an Outlook mail item (does not send by default).
#
# Usage:
#   tclsh Outlook-01_Mail.tcl <recipient> <subject> <body> [attachment...]
#
# Requires: CAWT + Twapi, MS Outlook installed and configured with a profile.

set recipient [lindex $argv 0]
if { $recipient eq "" } {
    puts "Usage: tclsh Outlook-01_Mail.tcl <recipient> \[subject\] \[body\] \[attachment...\]"
    exit 1
}
set toList [list $recipient]
set subject [lindex $argv 1]
set body    [lindex $argv 2]

if { $subject eq "" } { set subject "Test mail from CAWT" }
if { $body eq "" }    { set body "This mail was generated with CAWT." }

# Attachments: all remaining arguments.
set attachmentList [lrange $argv 3 end]

package require cawt

set appId [Outlook Open]

set mailId [Outlook CreateMail $appId $toList $subject $body $attachmentList]
puts "Mail item created (not sent). It is visible in Outlook's outbox/drafts."
puts "To:      $toList"
puts "Subject: $subject"

# To actually send, uncomment the next line (or pass send=true as the last arg):
# Outlook SendMail $mailId

Cawt Destroy $mailId
Cawt Destroy
puts "Done."
