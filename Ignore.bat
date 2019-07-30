@echo off
move ".emr" "..\Moristory Src\before rs\.emr"
move ".git" "..\Moristory Src\before rs\.git"
move "MSSEditor" "..\Moristory Src\before rs\MSSEditor"
move "code" "..\Moristory Src\before rs\code"
move "core" "..\Moristory Src\before rs\core"
copy "*.*" "..\Moristory Src\before rs\*.*"
del "*.*"
copy "..\Moristory Src\before rs\*.exe" "*.exe"
copy "..\Moristory Src\before rs\*.dll" "*.dll"
pause