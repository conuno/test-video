@echo off
setlocal EnableDelayedExpansion
rem Converts all reports in txtResults\ to Excel files in excelResults\.
rem   txt2excel.bat
rem   txt2excel.bat -SourceDir "D:\reports" -TargetDir "D:\excel"

rem Re-quote every argument and double a trailing backslash (see run.bat).
rem Save the script path first: shift below also shifts %0.
set "SCRIPT=%~dp0txt2excel.ps1"
set "ARGS="
:next
if "%~1"=="" goto run
set "A=%~1"
if "!A:~-1!"=="\" set "A=!A!\"
set ARGS=!ARGS! "!A!"
shift
goto next

:run
powershell -NoProfile -ExecutionPolicy Bypass -File "!SCRIPT!" !ARGS!
exit /b %ERRORLEVEL%
