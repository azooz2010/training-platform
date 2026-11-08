@echo off
setlocal
pushd "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-platform.ps1"
set "RESULT=%ERRORLEVEL%"
popd
if not "%RESULT%"=="0" (
  echo.
  echo تعذر تجهيز المثبت. أرسل ملف BUILD_LOG.txt للمراجعة.
  pause
)
exit /b %RESULT%
