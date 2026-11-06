#!/usr/bin/env tclsh
# cawt-check.tcl — быстрая проверка наличия CAWT
# Вызывается из PowerShell для проверки кандидата на интерпретатор
#
# Возвращает:
#   exit 0 + "OK <cawt_version>" если CAWT доступен
#   exit 1 + "FAIL <reason>" если CAWT отсутствует

if {[catch {package require cawt} msg]} {
    puts "FAIL: $msg"
    exit 1
}

set cawt_ver [package present cawt]
set twapi_ver [package present twapi]
set tcl_ver $tcl_version

puts "OK: CAWT=$cawt_ver Twapi=$twapi_ver Tcl=$tcl_ver"
exit 0
