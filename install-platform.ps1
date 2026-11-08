$ErrorActionPreference = 'Stop'
$Project = $PSScriptRoot
$Log = Join-Path $Project 'BUILD_LOG.txt'
Start-Transcript -Path $Log -Append | Out-Null
try {
    Write-Host ''
    Write-Host 'منصة اختبار إدخال البيانات - تجهيز مثبت Windows مستقل' -ForegroundColor Cyan
    Write-Host 'التحقق من Node.js و npm...'
    $node = Get-Command node.exe -ErrorAction SilentlyContinue
    $npm = Get-Command npm.cmd -ErrorAction SilentlyContinue
    if (-not $node -or -not $npm) {
        throw 'يلزم تثبيت Node.js LTS مرة واحدة على جهاز تجهيز المثبت. بعد ذلك أعد تشغيل هذا الملف.'
    }

    Set-Location $Project
    # Retry package downloads and use the Electron mirror to avoid the default GitHub binary timeout.
    $env:ELECTRON_MIRROR = 'https://github.com/electron/electron/releases/download/'
    & npm.cmd config set fetch-retries 5
    & npm.cmd config set fetch-timeout 300000
    & npm.cmd config set registry https://registry.npmjs.org/
    Write-Host 'تنزيل مكونات البناء مرة واحدة. قد يحتاج هذا الجزء إلى الإنترنت...' -ForegroundColor Cyan
    & npm.cmd install --no-audit --no-fund
    if ($LASTEXITCODE -ne 0) {
        throw 'فشل تنزيل مكونات البناء. راجع BUILD_LOG.txt واتصال الإنترنت أو جرّب لاحقاً.'
    }

    Write-Host 'إنشاء ملف Setup.exe الذي يعمل دون إنترنت عند تثبيته...' -ForegroundColor Cyan
    & npm.cmd run dist
    if ($LASTEXITCODE -ne 0) { throw "فشل إنشاء المثبت. رمز الخطأ: $LASTEXITCODE" }

    $dist = Join-Path $Project 'dist'
    $setup = Get-ChildItem -Path $dist -Filter '*.exe' -File | Where-Object { $_.Name -notlike '*uninstall*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $setup) { throw 'لم يتم العثور على Setup.exe داخل مجلد dist.' }

    $desktop = [Environment]::GetFolderPath('Desktop')
    if (-not (Test-Path $desktop)) { $desktop = Join-Path $env:USERPROFILE 'OneDrive\Desktop' }
    if (Test-Path $desktop) {
        $desktopSetup = Join-Path $desktop 'منصة_اختبار_إدخال_البيانات_Setup.exe'
        Copy-Item -LiteralPath $setup.FullName -Destination $desktopSetup -Force
        Write-Host "تم نسخ ملف التثبيت إلى سطح المكتب: $desktopSetup" -ForegroundColor Green
    }
    Write-Host ''
    Write-Host 'اكتمل إنشاء المثبت. يمكنك نسخ Setup.exe لأي جهاز Windows x64 وتثبيته دون إنترنت.' -ForegroundColor Green
    Write-Host 'لن يحتاج جهاز التثبيت إلى Node.js أو npm.' -ForegroundColor Green
    $open = Read-Host 'هل تريد تشغيل المثبت الآن؟ اكتب Y للتشغيل أو N للإغلاق'
    if ($open -match '^(Y|y)$') { Start-Process -FilePath $setup.FullName -Wait }
}
catch {
    Write-Host ''
    Write-Host ("ERROR: " + $_.Exception.Message) -ForegroundColor Red
    Write-Host "التفاصيل محفوظة في: $Log" -ForegroundColor Yellow
    Stop-Transcript | Out-Null
    exit 1
}
Stop-Transcript | Out-Null
exit 0
