@echo off
"C:\Program Files\TortoiseSVN\bin\SubWCRev.exe" %1  %2 %3 > NUL
fc /B %3 %4 > NUL 2>&1
if not errorlevel 1 goto exit
copy /y %3 %4 > NUL
:exit