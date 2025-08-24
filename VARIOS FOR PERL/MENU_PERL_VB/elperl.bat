echo off
title %1 -- %2
if "%2"=="/DROP" goto DROP2
:INICIO
set path=%PATH%;c:\perl\bin
cls

ECHO ===%1====
echo 1  NORMAL (EN PANTALLA)
echo 2. VERBOSE (EN PANTALLA)
echo 3. CHECKEAR SINTAXIS (OK)
echo 4. SALIDA TEXTO 
echo 5. SALIDA HTML
echo 6. -DEBUG
echo 7. MSDOS
echo 8. SALIR   
echo 9. cONVERTIR A BAT
echo A. DRAG AND DROP
ECHO ============
choice.com /C:123456789a /S ELIGE OPCION
REM pause
if errorlevel 10 goto DROP
if errorlevel 9 goto BAT
if errorlevel 8 goto END
if errorlevel 7 goto MDOS
IF ERRORLEVEL 6 GOTO DEB
if errorlevel 5 goto EXP
if errorlevel 4 GOTo NOTE
if ERRORLEVEL 3 GOTO CHECK
IF ERRORLEVEL 2 GOTO VERBO
IF ERRORLEVEL 1 GOTO NORMAL

GOTO COMAN

pause
goto INICIO

:NORMAL
set CHECK=
set DEB=
goto COMAN

:CHECK
set CHECK=-c
set DEB=
goto COMAN

:NOTE
set ET=Perl_Out.txt
perl %1 >%ET%
start NOTEPAD.EXE %ET%
goto INICIO

:EXP
set ET=Perl_Out.html
perl %1>%ET%
start %ET%
goto INICIO

:DEB
perl -d %1
goto INICIO

:MDOS
command.com
goto INICIO

:VERBO
cls
perl -w %1
pause
goto INICIO

:COMAN
cls
perl %CHECK% %DEB% %1
pause
goto INICIO

:BAT
pl2bat %1
goto END

:DROP
prEnvironDos.exe XD
start /SHARED elperl.bat %1 /DROP
goto END

:DROP2
echo DROPPING TO PERL
echo.
perl %1 %XD%
echo.
set XD=
set %2=
pause
goto INICIO

:END