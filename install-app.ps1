$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$html = Join-Path $root "index.html"
if (!(Test-Path $html)) { throw "لم يتم العثور على index.html داخل مجلد البرنامج." }

$edgeCandidates = @(
  "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
  "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
)
$edge = $edgeCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if (!$edge) {
  $chromeCandidates = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
  )
  $edge = $chromeCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}
if (!$edge) { throw "لم يتم العثور على Microsoft Edge أو Google Chrome. ثبّت أحدهما ثم أعد المحاولة." }

$uri = ([System.Uri]$html).AbsoluteUri
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "منصة اختبار إدخال البيانات.lnk"
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $edge
$shortcut.Arguments = "--app=`"$uri`""
$shortcut.WorkingDirectory = $root
$shortcut.Description = "منصة اختبار إدخال البيانات ومعالجة النصوص - تعمل دون إنترنت"
$logoIcon = Join-Path $root "assets\institute-logo.ico"
if (Test-Path $logoIcon) { $shortcut.IconLocation = "$logoIcon,0" } else { $shortcut.IconLocation = "$edge,0" }
$shortcut.Save()

Start-Process -FilePath $edge -ArgumentList "--app=`"$uri`""
Write-Host "تم إنشاء اختصار المنصة على سطح المكتب وتشغيلها." -ForegroundColor Green
Start-Sleep -Seconds 2
