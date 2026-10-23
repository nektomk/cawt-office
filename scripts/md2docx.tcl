# md2docx.tcl - Convert a Markdown file to a styled Word document via CAWT,
# in a VISIBLE Word instance left open for the user.
#
# Usage:
#   C:\ATcl\bin\tclsh.exe md2docx.tcl <input.md> <output.docx>
#
# Supported Markdown subset: #/##/### headings, fenced code blocks (```),
# blockquotes (>), bullet/numbered lists with continuation lines, paragraphs
# with inline **bold**, `code`, [links](url) and <raw urls>.
# Source is ASCII-only; all Unicode comes from the input file (read as UTF-8).

if { [llength $argv] < 2 } {
    puts "Usage: tclsh md2docx.tcl <input.md> <output.docx>"
    exit 1
}
set mdPath  [file normalize [lindex $argv 0]]
set outPath [file nativename [file normalize [lindex $argv 1]]]

# ---------------- Read markdown ----------------
set fh [open $mdPath r]
fconfigure $fh -encoding utf-8
set content [read $fh]
close $fh
set mdLines [split $content "\n"]

# ---------------- Preprocess: join soft wraps into logical blocks ----------------
set blocks {}
set buf ""
set bufTag ""

proc flushBuf {} {
    global buf bufTag blocks
    if { $bufTag ne "" && [string trim $buf] ne "" } {
        lappend blocks [list $bufTag [string trim $buf]]
    }
    set buf ""
    set bufTag ""
}

set inCode false
foreach raw $mdLines {
    set line [string trimright $raw "\r"]

    if { [string match "```*" $line] } {
        flushBuf
        set inCode [expr { ! $inCode }]
        continue
    }
    if { $inCode } {
        lappend blocks [list code $line]
        continue
    }
    if { [string trim $line] eq "" } {
        flushBuf
        continue
    }
    if { [regexp {^(#{1,6})\s+(.*)$} $line -> hashes htxt] } {
        flushBuf
        set lvl [string length $hashes]
        if { $lvl > 3 } { set lvl 3 }
        lappend blocks [list h$lvl $htxt]
        continue
    }
    if { [string match ">*" $line] } {
        set qtxt [string trim [string range $line 1 end]]
        if { $qtxt eq "" } {
            flushBuf
            lappend blocks {qempty {}}
            continue
        }
        if { $bufTag eq "q" } {
            append buf " " $qtxt
        } else {
            flushBuf
            set bufTag "q"
            set buf $qtxt
        }
        continue
    }
    if { [regexp {^-\s+(.*)$} $line -> itxt] } {
        flushBuf
        lappend blocks [list ul $itxt]
        continue
    }
    if { [regexp {^\s*(\d+)\.\s+(.*)$} $line -> num itxt] } {
        flushBuf
        lappend blocks [list ol [list $num $itxt]]
        continue
    }

    # Plain text: soft-wrap continuation merges into the current buffer,
    # indented lines after a list item merge into that item.
    set trimmed [string trim $line]
    if { [regexp {^\s\s+} $line] && [llength $blocks] > 0 } {
        set lastTag  [lindex [lindex $blocks end] 0]
        set lastData [lindex [lindex $blocks end] 1]
        if { $lastTag eq "ul" } {
            lset blocks end 1 "$lastData $trimmed"
            continue
        }
        if { $lastTag eq "ol" } {
            lset blocks end 1 [list [lindex $lastData 0] "[lindex $lastData 1] $trimmed"]
            continue
        }
    }
    if { $bufTag eq "p" } {
        append buf " " $trimmed
    } else {
        flushBuf
        set bufTag "p"
        set buf $trimmed
    }
}
flushBuf

puts "Parsed [llength $mdLines] lines into [llength $blocks] blocks"

# ---------------- Inline markdown parser ----------------
# Returns a list of {text kind} segments; kind: n=normal, b=bold, c=code.
proc parseInline { text } {
    regsub -all {\[([^\]]+)\]\(([^)]+)\)} $text {\1 (\2)} text
    regsub -all {<(https?://[^>]+)>} $text {\1} text

    set segs {}
    set ci 0
    foreach codePart [split $text `] {
        incr ci
        if { $ci % 2 == 0 } {
            if { $codePart ne "" } { lappend segs [list $codePart c] }
            continue
        }
        set mapped [string map [list "**" \u0001] $codePart]
        set bi 0
        foreach bPart [split $mapped \u0001] {
            incr bi
            if { $bPart eq "" } { continue }
            if { $bi % 2 == 0 } {
                lappend segs [list $bPart b]
            } else {
                lappend segs [list $bPart n]
            }
        }
    }
    if { [llength $segs] == 0 } { lappend segs [list "" n] }
    return $segs
}

# ---------------- Rendering helpers ----------------
proc applyRunStyle { r kind } {
    switch -- $kind {
        b { catch { $r -with { Font } Bold -1 } }
        c {
            catch { $r -with { Font } Name "Consolas" }
            catch { $r -with { Font } Size 9.5 }
            catch { $r -with { Font } Color [Cawt GetColor "#9C2440"] }
            catch { $r -with { Font Shading } BackgroundPatternColor [Cawt GetColor "#F2F2F2"] }
        }
    }
}

# Append inline-parsed text as runs; paragraph break after the last run.
# $decor (optional) is eval'd per run with $r and $kind in scope.
proc appendSegs { docId text { decor "" } } {
    set segs [parseInline $text]
    set n [llength $segs]
    set i 0
    foreach seg $segs {
        incr i
        lassign $seg txt kind
        set r [Word AppendText $docId $txt [expr { $i == $n }]]
        applyRunStyle $r $kind
        if { $decor ne "" } { eval $decor }
        Cawt Destroy $r
    }
}

proc lastParaRange { docId } {
    set paras [$docId Paragraphs]
    set cnt [$paras Count]
    # Word's Paragraphs collection includes the trailing empty paragraph after
    # the final paragraph mark, so the paragraph just appended is Item Count-1
    # (verified live with word_paras_probe.tcl).
    set idx [expr { $cnt > 1 ? $cnt - 1 : $cnt }]
    set p [$paras Item $idx]
    set r [$p Range]
    Cawt Destroy $p
    Cawt Destroy $paras
    return $r
}

proc renderBlock { docId block } {
    lassign $block tag data
    switch -- $tag {
        h1 - h2 - h3 {
            set lvl   [string index $tag 1]
            set style [lindex {wdStyleHeading1 wdStyleHeading2 wdStyleHeading3} [expr { $lvl - 1 }]]
            set col   [lindex {#1F4E79 #2E74B5 #1F4E79} [expr { $lvl - 1 }]]
            set r [Word AppendText $docId $data true $style]
            catch { $r -with { Font } Color [Cawt GetColor $col] }
            if { $lvl == 1 } {
                catch { $r -with { Font } Size 20 }
                catch { $r -with { ParagraphFormat } SpaceAfter 10 }
            } elseif { $lvl == 2 } {
                catch { $r -with { Font } Size 15 }
                catch { $r -with { ParagraphFormat } SpaceBefore 14 SpaceAfter 5 }
            } else {
                catch { $r -with { Font } Size 12.5 }
            }
            Cawt Destroy $r
        }
        p {
            appendSegs $docId $data
            set pr [lastParaRange $docId]
            catch { $pr -with { ParagraphFormat } SpaceAfter 6 }
            Cawt Destroy $pr
        }
        q {
            appendSegs $docId $data {
                $r -with { Font } Italic -1
                if { $kind ne "c" } {
                    $r -with { Font } Color [Cawt GetColor "#33587A"]
                }
            }
            set pr [lastParaRange $docId]
            catch { $pr -with { ParagraphFormat } LeftIndent 14 RightIndent 14 SpaceBefore 2 SpaceAfter 3 }
            catch { $pr -with { Shading } BackgroundPatternColor [Cawt GetColor "#EEF3FA"] }
            Cawt Destroy $pr
        }
        qempty {
            set r [Word AppendText $docId "" true]
            Cawt Destroy $r
            set pr [lastParaRange $docId]
            catch { $pr -with { Font } Size 4 }
            catch { $pr -with { ParagraphFormat } LeftIndent 14 RightIndent 14 SpaceBefore 0 SpaceAfter 0 }
            catch { $pr -with { Shading } BackgroundPatternColor [Cawt GetColor "#EEF3FA"] }
            Cawt Destroy $pr
        }
        ul - ol {
            if { $tag eq "ul" } {
                set prefix "\u2022  "
                set body   $data
            } else {
                set prefix "[lindex $data 0].  "
                set body   [lindex $data 1]
            }
            set r [Word AppendText $docId $prefix false]
            catch { $r -with { Font } Bold -1 }
            catch { $r -with { Font } Color [Cawt GetColor "#1F4E79"] }
            Cawt Destroy $r
            appendSegs $docId $body
            set pr [lastParaRange $docId]
            catch { $pr -with { ParagraphFormat } LeftIndent 24 FirstLineIndent -13 SpaceAfter 3 }
            Cawt Destroy $pr
        }
        code {
            set r [Word AppendText $docId $data true]
            catch { $r -with { Font } Name "Consolas" }
            catch { $r -with { Font } Size 9 }
            catch { $r -with { Font } Color [Cawt GetColor "#1F1F1F"] }
            Cawt Destroy $r
            set pr [lastParaRange $docId]
            catch { $pr -with { ParagraphFormat } SpaceBefore 0 SpaceAfter 0 LeftIndent 10 }
            catch { $pr -with { Shading } BackgroundPatternColor [Cawt GetColor "#F5F5F5"] }
            Cawt Destroy $pr
        }
        default {
            puts "  (skip unknown block tag: $tag)"
        }
    }
}

# ---------------- Build the document in a visible Word ----------------
set docId ""
set appId ""
if { [catch {
    package require cawt

    # Own VISIBLE instance (first arg = visible). Never touches the user's
    # other Word windows; we do NOT quit it - the document stays on screen.
    set appId [Word OpenNew true]
    catch { $appId DisplayAlerts 0 }
    set docId [Word AddDocument $appId]

    set cnt 0
    foreach block $blocks {
        renderBlock $docId $block
        incr cnt
        if { $cnt % 20 == 0 } { puts "  rendered $cnt/[llength $blocks] blocks..." }
    }

    set nChars [Word GetNumCharacters $docId]
    puts "Rendered $cnt blocks, $nChars characters"

    Word SaveAs $docId $outPath
    puts "Saved: $outPath"
    catch { $docId Activate }
} err] } {
    puts "FATAL: $err"
    # Close only OUR document (no prompts); leave any Word window alone -
    # we cannot be 100% sure the instance is exclusively ours.
    if { $docId ne "" } { catch { $docId Close 0 } }
    catch { Cawt Destroy }
    exit 1
}

# Release COM handlers; the visible Word keeps the document open.
Cawt Destroy
