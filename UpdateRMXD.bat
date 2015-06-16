
@echo off
set GETVER=getver.py

for /f %%a in ('!GETVER! "C:\Program Files (x86)\ArcGIS\Desktop10.2\bin\afcore.dll"') do (
  for /f "tokens=1,2 delims=." %%b in ('echo %%a') do (
    set HOMEBUILD=%%b.%%c
  )
)
if "!HOMEBUILD!" equ "" goto :end
if "!HOMEBUILD!" equ "?." goto :end

if exist %~dp0bin\x64\Release (
	set SRCDIR=%~dp0bin\x64\Release
) else if exist %~dp0bin\x86\Release (
	set SRCDIR=%~dp0bin\x86\Release
) else if exist %~dp0bin\Release (
	set SRCDIR=%~dp0bin\Release
) else (
	echo source dirs %~dp0bin\...\Release not found
	goto :end
)

echo copying Latest ReadMxd to:
call :copyRMXD D:\Projects\ReadMxd\dist

:end
echo.
set GETVER=
set HOMEBUILD=
set BUILD=
set SRCDIR=
pause
goto :eof

:copyRMXD
echo.
echo %1
if not exist %1 (
	echo %1 not found
  goto :eof
)
echo Copying ReadMxd from %SRCDIR% to %1
if exist "%SRCDIR%\ReadMxdW.exe" xcopy "%SRCDIR%\ReadMxdW.exe" %1 /y /q
if exist "%SRCDIR%\ReadMxdW.pdb" xcopy "%SRCDIR%\ReadMxdW.pdb" %1 /y /q
if exist "%SRCDIR%\ReadMxdXI.exe" xcopy "%SRCDIR%\ReadMxdXI.exe" %1 /y /q
if exist "%SRCDIR%\ReadMxdXI.pdb" xcopy "%SRCDIR%\ReadMxdXI.pdb" %1 /y /q
if exist "%SRCDIR%\Ionic.Utils.Zip.dll" xcopy "%SRCDIR%\Ionic.Utils.Zip.dll" %1 /y /q
if exist "%SRCDIR%\Interop.ArcGISVersionLib.dll" xcopy "%SRCDIR%\Interop.ArcGISVersionLib.dll" %1 /y /q
goto :eof
