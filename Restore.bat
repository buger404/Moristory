@echo off
move ".emr" "..\..\Moristory\.emr"
move ".git" "..\..\Moristory\.git"
move "MSSEditor" "..\..\Moristory\MSSEditor"
move "code" "..\..\Moristory\code"
move "core" "..\..\Moristory\core"
copy "*.*" "..\..\Moristory\*.*"
del "*.*"
pause