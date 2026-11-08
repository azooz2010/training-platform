@echo off
setlocal
pushd "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-platform.ps1"
set "RESULT=%ERRORLEVEL%"
popd
if not "%RESULT%"=="0" (
  echo.
  echo Installation preparation failed. Send the file BUILD_LOG.txt to support.
  pause
)
exit /b %RESULT%
