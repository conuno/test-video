@echo off
setlocal EnableDelayedExpansion
rem Analyses the given videos (files and/or folders) one after another.
rem   run.bat "D:\videos"                     all .mp4 in that folder
rem   run.bat "D:\a.mp4" "D:\b.mp4"            these files, in this order
rem   run.bat -TestClaude                      check Claude Code login only
rem   run.bat "D:\videos" -DryRun              frames and task files only, no Claude calls
rem   run.bat "D:\videos" -TestOneWindow       analyse only the first 15 s of the first video
rem   run.bat "D:\videos" -KeepWork            keep work\ for inspection

rem Re-quote every argument and double a trailing backslash: powershell.exe -File reads
rem "D:\videos\" as an escaped quote and would glue all following arguments into one.
rem Save the script path first: shift below also shifts %0.
set "SCRIPT=%~dp0videoqa.ps1"
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
