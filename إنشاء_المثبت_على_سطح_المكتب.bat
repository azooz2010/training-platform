@echo off
setlocal EnableExtensions
chcp 65001 >nul
title تجهيز مثبت منصة اختبار إدخال البيانات
cd /d "%~dp0"

echo ==================================================
echo   منصة اختبار إدخال البيانات ومعالجة النصوص
echo   إنشاء ملف التثبيت لويندوز 64 بت
echo ==================================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo لم يتم العثور على Node.js.
  where winget >nul 2>nul
  if not errorlevel 1 (
    echo جاري تثبيت Node.js LTS باستخدام winget...
    winget install --id OpenJS.NodeJS.LTS -e --accept-package-agreements --accept-source-agreements
    echo.
    echo أغلق هذه النافذة وافتح هذا الملف مرة أخرى بعد اكتمال التثبيت.
    pause
    exit /b 1
  )
  echo ثبّت Node.js LTS من https://nodejs.org ثم أعد تشغيل هذا الملف.
  start "" "https://nodejs.org/"
  pause
  exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
  echo لم يتم العثور على npm. أعد تثبيت Node.js LTS ثم أعد المحاولة.
  pause
  exit /b 1
)

echo [1/3] تثبيت مكتبات التطبيق (يتطلب الإنترنت في أول مرة)...
call npm install --no-audit --no-fund
if errorlevel 1 goto fail

echo.
echo [2/3] إنشاء مثبت Windows x64...
call npx electron-builder --win nsis --x64
if errorlevel 1 goto fail

echo.
echo [3/3] التحقق من ملف التثبيت...
set "SETUP="
for %%F in ("dist\*.exe") do set "SETUP=%%~fF"
if not defined SETUP (
  echo لم يتم العثور على ملف EXE داخل dist.
  goto fail
)

echo تم إنشاء المثبت بنجاح:
echo "%SETUP%"
set "DESKTOP=%USERPROFILE%\Desktop"
if exist "%DESKTOP%" (
  copy /Y "%SETUP%" "%DESKTOP%\منصة_اختبار_إدخال_البيانات_Setup.exe" >nul
  if not errorlevel 1 echo تم نسخ ملف التثبيت أيضاً إلى سطح المكتب.
)

echo.
echo اكتملت العملية. شغّل ملف Setup.exe لتثبيت المنصة.
if exist dist start "" explorer "%CD%\dist"
pause
exit /b 0

:fail
echo.
echo فشل إنشاء المثبت. احتفظ برسالة الخطأ الظاهرة وأرسل صورة لها.
pause
exit /b 1
