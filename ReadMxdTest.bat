
@echo off

set READMXD=d:\Projects\ReadMxd\dist\ReadMxdW.exe
set TESTDIR=c:\temp\ReadMxdTest
set LOGFILE=%TESTDIR%\_results.txt
if exist %TESTDIR% (rmdir %TESTDIR% /s)
mkdir %TESTDIR%
pushd %TESTDIR%
echo Starting ReadMxd test at %DATE% %TIME%> %LOGFILE%

call :searchFolder d:\maps
call :searchFolder C:\Users\jpm\Documents\ArcGIS

:: check for errors etc ::
set TESTDIR=/%TESTDIR:\=/%
set TESTDIR=%TESTDIR::=%
call grep -Hi error %TESTDIR%/*.log >> %LOGFILE%
call grep -Hi warning %TESTDIR%/*.log >> %LOGFILE%
call grep -Hi unknown %TESTDIR%/*.log >> %LOGFILE%
call grep -Hi -B1 todo %TESTDIR%/*.log >> %LOGFILE%
call grep -Hi System.__ComObject %TESTDIR%/*.log >> %LOGFILE%
echo Done at %DATE% %TIME%>> %LOGFILE%
popd

echo OK

set TESTDIR=
set READMXD=
set LOGFILE=

goto :eof

:searchFolder
setlocal
	echo Searching %1.
	call :searchForType %1 mxd
	call :searchForType %1 lyr
	for /f "tokens=*" %%a in ('dir %1 /ad/b') do (
		call :searchFolder "%~1\%%~a"
	)
endlocal
goto :eof

:searchForType
setlocal
	if not exist %1\*.%2 goto :eof
	for /f "tokens=*" %%a in ('dir %1\*.%2 /b') do (
		echo %%a
		set LOGNAME=%%a
		if /i "%2" equ "mxd" (
			set LOGNAME=!LOGNAME:.mxd=_props.log!
		) else (
			set LOGNAME=!LOGNAME:.=_!_props.log
		)
		if exist %TESTDIR%\!LOGNAME! (
			for /f %%t in ('GetDateStamp.bat') do (
				set DATESTAMP=%%t
			)
			echo !LOGNAME! already exists, renaming
			ren "%TESTDIR%\!LOGNAME!" "!LOGNAME:.log=!_!DATESTAMP!.log"
		)
		call %READMXD% %1\%%a -s -b -y -e
	)
endlocal
goto :eof