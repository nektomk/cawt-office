# md2docx.tcl - Convert a Markdown file to a styled Word document via CAWT,
# in a VISIBLE Word instance left open for the user.
#
# Usage:
#   tclsh md2docx.tcl <input.md> <output.docx> [overwrite]
#   overwrite = 1/true/yes/force permits replacing an existing file;
#   without it an existing <output.docx> is refused with exit code 2 (S10).
#
# Pipeline:  Markdown::convert (tcllib) -> XHTML fragment -> tdom DOM tree
#            -> recursive render into Word via CAWT.
#
# Supported Markdown: headings (# .. ######), fenced code blocks (```),
# blockquotes (>), bullet/numbered lists (incl. nested), tables (GFM:
# |---|---|), paragraphs with inline **bold**, *italic*, `code` and
# [links](url). Source is read as UTF-8 (may contain any Unicode).
#
# The Markdown::convert (tcllib, part of the Markdown 1.x package) and tdom
# (XML/DOM parser) packages must be installed; both ship with BAWT/ActiveTcl.

if { [llength $argv] < 2 } {
    puts "Usage: tclsh md2docx.tcl <input.md> <output.docx> [overwrite]"
    exit 1
}
set mdPath  [file normalize [lindex $argv 0]]
set outPath [file nativename [file normalize [lindex $argv 1]]]

# S10 overwrite consent: never silently replace an existing file.
set owRaw [lindex $argv 2]
set overwrite [expr {[string tolower $owRaw] in {1 true yes overwrite force}}]
if { $owRaw ne "" && ! $overwrite } {
    puts "ERROR: bad overwrite value '$owRaw' (use 1)"
    exit 1
}
if { ! $overwrite && [file exists $outPath] } {
    puts "ERROR: $outPath already exists - rerun with overwrite arg 1 or choose another name"
    exit 2
}

# ---------------- Read markdown + convert to XHTML ----------------
set fh [open $mdPath r]
fconfigure $fh -encoding utf-8
set content [read $fh]
close $fh

package require Markdown
package require tdom

# Returns an XHTML fragment (multiple top-level elements).
set frag [Markdown::convert $content]
set doc [dom parse "<root>$frag</root>"]
set root [$doc documentElement]

# ---------------- Rendering helpers ----------------
# Collect the plain text of a node (concatenated descendant text nodes).
proc nodeText { node } {
    set out ""
    foreach child [$node childNodes] {
        if { [$child nodeName] eq "#text" } {
            append out [$child data]
        } else {
            append out [nodeText $child]
        }
    }
    return $out
}

# Collect inline segments {text kind} from a node's children.
# kind: n=normal, b=bold, i=italic, c=code.
proc collectInline { node } {
    set segs {}
    foreach child [$node childNodes] {
        set name [$child nodeName]
        switch -- $name {
            "#text"     { set t [$child data]; if { $t ne "" } { lappend segs [list $t n] } }
            "strong" - "b" {
                foreach s [collectInline $child] { lset s 1 b; lappend segs $s }
            }
            "em" - "i"  {
                foreach s [collectInline $child] { lset s 1 i; lappend segs $s }
            }
            "code"      { lappend segs [list [nodeText $child] c] }
            # Links: rendered as plain text (no hyperlink creation here).
            "a"         { foreach s [collectInline $child] { lappend segs $s } }
            default     { foreach s [collectInline $child] { lappend segs $s } }
        }
    }
    if { [llength $segs] == 0 } { lappend segs [list "" n] }
    return $segs
}

proc applyRunStyle { r kind } {
    switch -- $kind {
        n { catch { $r -with { Font } Reset } }
        b { catch { $r -with { Font } Reset }; catch { $r -with { Font } Bold -1 } }
        i { catch { $r -with { Font } Reset }; catch { $r -with { Font } Italic -1 } }
        c {
            catch { $r -with { Font } Name "Consolas" }
            catch { $r -with { Font } Size 9.5 }
            catch { $r -with { Font } Color [Cawt GetColor "#9C2440"] }
            catch { $r -with { Font Shading } BackgroundPatternColor [Cawt GetColor "#F2F2F2"] }
        }
    }
}

# Append inline segments as runs in the current (last) paragraph.
# Manual GetEndRange+InsertAfter instead of Word AppendText: AppendText applies
# paragraph style wdStyleNormal via SetRangeStyle, which erases the direct
# formatting (fonts) of *all* runs in that paragraph, so the first run's style
# leaks to the end of the paragraph. InsertAfter only affects the new text.
proc appendSegs { docId segs } {
    foreach seg $segs {
        lassign $seg txt kind
        set r [Word GetEndRange $docId]
        $r InsertAfter $txt
        applyRunStyle $r $kind
        Cawt Destroy $r
    }
    # Terminate the paragraph after the last run.
    set end [Word GetEndRange $docId]
    $end InsertParagraphAfter
    Cawt Destroy $end
}

proc lastParaRange { docId } {
    set paras [$docId Paragraphs]
    set cnt [$paras Count]
    # Word's Paragraphs collection includes the trailing empty paragraph after
    # the final paragraph mark, so the paragraph just appended is Item Count-1.
    set idx [expr { $cnt > 1 ? $cnt - 1 : $cnt }]
    set p [$paras Item $idx]
    set r [$p Range]
    Cawt Destroy $p
    Cawt Destroy $paras
    return $r
}

# ---------------- Block renderers ----------------
proc renderHeading { docId node } {
    set lvl   [string index [$node nodeName] 1]
    if { $lvl > 3 } { set lvl 3 }
    set style [lindex {wdStyleHeading1 wdStyleHeading2 wdStyleHeading3} [expr { $lvl - 1 }]]
    set col   [lindex {#1F4E79 #2E74B5 #1F4E79} [expr { $lvl - 1 }]]
    set r [Word AppendText $docId [nodeText $node] true $style]
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

proc renderParagraph { docId node } {
    appendSegs $docId [collectInline $node]
    set pr [lastParaRange $docId]
    catch { $pr -with { ParagraphFormat } SpaceAfter 6 }
    Cawt Destroy $pr
}

proc renderQuote { docId node } {
    # A blockquote wraps its paragraphs in <p>; render the text inline with
    # quote styling (italic, tinted background, indents).
    appendSegs $docId [collectInline $node]
    set pr [lastParaRange $docId]
    catch { $pr -with { Font } Italic -1 }
    catch { $pr -with { ParagraphFormat } LeftIndent 14 RightIndent 14 SpaceBefore 2 SpaceAfter 3 }
    catch { $pr -with { Shading } BackgroundPatternColor [Cawt GetColor "#EEF3FA"] }
    Cawt Destroy $pr
}

proc renderCodeBlock { docId node } {
    # The pre/code text is multi-line: Word splits one AppendText call into
    # several paragraphs. Style the returned range $r (covers ALL those lines)
    # instead of lastParaRange (which would style only the trailing line).
    set r [Word AppendText $docId [nodeText $node] true]
    catch { $r -with { Font } Name "Consolas" }
    catch { $r -with { Font } Size 9 }
    catch { $r -with { Font } Color [Cawt GetColor "#1F1F1F"] }
    catch { $r -with { ParagraphFormat } SpaceBefore 0 SpaceAfter 0 LeftIndent 10 }
    catch { $r -with { Shading } BackgroundPatternColor [Cawt GetColor "#F5F5F5"] }
    Cawt Destroy $r
}

proc renderList { docId node } {
    set ordered [expr { [$node nodeName] eq "ol" }]
    set counter 0
    foreach li [$node childNodes] {
        if { [$li nodeName] ne "li" } { continue }
        incr counter
        # Marker: bullet or number. Nested lists inside <li> are rendered
        # inline (the Markdown package flattens them into the same <li>).
        if { $ordered } {
            set marker "$counter.  "
        } else {
            set marker "\u2022  "
        }
        set r [Word AppendText $docId $marker false]
        catch { $r -with { Font } Bold -1 }
        catch { $r -with { Font } Color [Cawt GetColor "#1F4E79"] }
        Cawt Destroy $r
        appendSegs $docId [collectInline $li]
        set pr [lastParaRange $docId]
        catch { $pr -with { ParagraphFormat } LeftIndent 24 FirstLineIndent -13 SpaceAfter 3 }
        Cawt Destroy $pr
    }
}

proc renderTable { docId node } {
    # Collect rows: list of {header|body rowCells}. The Markdown package emits
    # <table><thead><tr><th>..</th></tr></thead><tbody><tr><td>..</td></tr></tbody></table>.
    set rows {}       ;# each: list of cell texts
    set headerRow {}  ;# first row (thead) or first body row if no thead
    set hasHeader false

    foreach child [$node childNodes] {
        set name [$child nodeName]
        if { $name eq "thead" } {
            set hasHeader true
            # thead's direct children are <tr> rows; take the first one as header.
            foreach tr [$child childNodes] {
                if { [$tr nodeName] eq "tr" } {
                    set headerRow [renderTrCells $tr]
                    break
                }
            }
        } elseif { $name eq "tbody" } {
            foreach tr [$child childNodes] {
                if { [$tr nodeName] eq "tr" } {
                    lappend rows [renderTrCells $tr]
                }
            }
        } elseif { $name eq "tr" } {
            lappend rows [renderTrCells $child]
        }
    }

    if { $hasHeader } {
        # Prepend the header row as the first table row.
        set allRows [concat [list $headerRow] $rows]
        set headerIndex 0
    } else {
        set allRows $rows
        set headerIndex -1
    }

    if { [llength $allRows] == 0 } { return }

    # Determine column count (max across rows).
    set nCols 0
    foreach r $allRows {
        set c [llength $r]
        if { $c > $nCols } { set nCols $c }
    }
    if { $nCols == 0 } { return }

    set nRows [llength $allRows]
    set tableRange [Word GetEndRange $docId]
    set tableId [Word AddTable $tableRange $nRows $nCols]
    Cawt Destroy $tableRange

    set r 0
    foreach row $allRows {
        incr r
        set c 0
        foreach cell $row {
            incr c
            if { $c <= $nCols } {
                Word SetCellValue $tableId $r $c $cell
            }
        }
    }

    # Style the header row (or first row if no header): bold + background.
    if { $nRows > 0 } {
        set styleRow 1
        for { set c 1 } { $c <= $nCols } { incr c } {
            set cr [Word GetCellRange $tableId $styleRow $c]
            catch { $cr -with { Font } Bold -1 }
            catch { $cr -with { Font } Color [Cawt GetColor "#FFFFFF"] }
            catch { $cr -with { Shading } BackgroundPatternColor [Cawt GetColor "#1F4E79"] }
            Cawt Destroy $cr
        }
    }

    # Keep column widths reasonable: autofit content.
    catch { $tableId -with { Columns } Autofit 1 }
    Cawt Destroy $tableId

    # Word inserts a paragraph after a table; ensure the following content
    # starts on a new paragraph.
    set blank [Word AppendText $docId "" true]
    Cawt Destroy $blank
}

proc renderTrCells { trNode } {
    set cells {}
    foreach cell [$trNode childNodes] {
        if { [$cell nodeName] eq "th" || [$cell nodeName] eq "td" } {
            lappend cells [nodeText $cell]
        }
    }
    return $cells
}

proc renderNode { docId node } {
    set name [$node nodeName]
    # Skip whitespace-only text nodes at block level.
    if { $name eq "#text" } {
        if { [string trim [$node data]] eq "" } { return }
        # Stray text outside a paragraph: render as paragraph.
        appendSegs $docId [collectInline $node]
        set pr [lastParaRange $docId]
        catch { $pr -with { ParagraphFormat } SpaceAfter 6 }
        Cawt Destroy $pr
        return
    }
    switch -- $name {
        h1 - h2 - h3 - h4 - h5 - h6 { renderHeading $docId $node }
        p                           { renderParagraph $docId $node }
        blockquote                  { renderQuote $docId $node }
        pre                         { renderCodeBlock $docId $node }
        ul - ol                     { renderList $docId $node }
        table                       { renderTable $docId $node }
        # Skip other structural wrappers; render their children.
        div - span { foreach child [$node childNodes] { renderNode $docId $child } }
        hr        { set r [Word AppendText $docId "" true]; Cawt Destroy $r }
        default   { foreach child [$node childNodes] { renderNode $docId $child } }
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
    set total [llength [$root childNodes]]
    foreach child [$root childNodes] {
        renderNode $docId $child
        incr cnt
        if { $cnt % 20 == 0 } { puts "  rendered $cnt/$total nodes..." }
    }

    set nChars [Word GetNumCharacters $docId]
    set nTables [Word GetNumTables $docId]
    puts "Rendered $cnt nodes, $nChars characters, $nTables tables"

    Word SaveAs $docId $outPath
    puts "Saved: $outPath"
    catch { $docId Activate }
} err] } {
    puts "FATAL: $err"
    # Locked-target hint; the COM message is locale-dependent, so match the
    # English and the Russian wording ("...another process" / "...другим процессом").
    if { [string match -nocase "*another process*" $err] || [string match -nocase "*\u0434\u0440\u0443\u0433\u0438\u043c \u043f\u0440\u043e\u0446\u0435\u0441\u0441\u043e\u043c*" $err] } {
        puts "HINT: target file is locked (probably open in Word) - close it and retry"
    }
    # Close only OUR document (no prompts), then quit OUR instance.
    if { $docId ne "" } { catch { $docId Close 0 } }
    if { $appId ne "" } { catch { Word Quit $appId } }
    catch { Cawt Destroy }
    exit 1
}

# Release COM handlers; the visible Word keeps the document open.
Cawt Destroy
