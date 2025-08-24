@echo off
title frmPerl.frm
echo %SCRIPT%
echo.

c:\perl\bin\perl.exe -c -w  "%SCRIPT%" %*

echo.
pause
