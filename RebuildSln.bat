@echo off

setlocal ENABLEDELAYEDEXPANSION
set LOGFILE=bin\x86\Release\build.log

:: rebuild solution in release ::
call "c:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\ide\devenv.exe" ReadMxd.sln /rebuild "Release|x86" /out !LOGFILE!

:: print last 3 lines of build log ::
for /f "tokens=*" %%a in ('type !LOGFILE!') do (
    set line3=!line2!
    set line2=!line1!
    set line1=%%a
)
echo.
echo !line3!
echo !line2!
echo !line1!
endlocal