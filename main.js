const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const { spawn } = require('child_process');
const { execFile } = require('child_process');

const OFFICE_BAR_HEIGHT = 112;
let officeTaskbarHidden = false;

function runWindowsOfficeLayout(action, kind = '') {
  if (process.platform !== 'win32') return;
  const script = `
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class OfficeWindowLayout {
 [DllImport("user32.dll", SetLastError=true)] public static extern IntPtr FindWindow(string cls, string title);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
 [DllImport("user32.dll", SetLastError=true)] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
}
"@
$taskbar = [OfficeWindowLayout]::FindWindow("Shell_TrayWnd", $null)
if ("${action}" -eq "hide") {
 if ($taskbar -ne [IntPtr]::Zero) { [OfficeWindowLayout]::ShowWindow($taskbar, 0) | Out-Null }
 Start-Sleep -Milliseconds 250
 if ("${kind}" -eq "word") { $procName = "WINWORD" } else { $procName = "EXCEL" }
 $p = Get-Process -Name $procName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
 if ($p) {
   Add-Type -AssemblyName System.Windows.Forms
   $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
   $bar = ${OFFICE_BAR_HEIGHT}
   [OfficeWindowLayout]::SetWindowPos($p.MainWindowHandle, [IntPtr]::Zero, $bounds.X, $bounds.Y, $bounds.Width, ($bounds.Height - $bar), 0x0040) | Out-Null
 }
} else {
 if ($taskbar -ne [IntPtr]::Zero) { [OfficeWindowLayout]::ShowWindow($taskbar, 5) | Out-Null }
}
`;
  execFile('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script], { windowsHide: true }, () => {});
}
function restoreWindowsTaskbar() {
  if (officeTaskbarHidden) {
    runWindowsOfficeLayout('show');
    officeTaskbarHidden = false;
  }
}


function findOfficeExe(kind) {
  const exe = kind === 'word' ? 'WINWORD.EXE' : 'EXCEL.EXE';
  const roots = [
    process.env['ProgramFiles'],
    process.env['ProgramFiles(x86)'],
    process.env['ProgramW6432'],
    process.env.LOCALAPPDATA
  ].filter(Boolean);
  const candidates = [];
  for (const root of roots) {
    candidates.push(
      path.join(root, 'Microsoft Office', 'root', 'Office16', exe),
      path.join(root, 'Microsoft Office', 'Office16', exe),
      path.join(root, 'Microsoft Office', 'Office15', exe),
      path.join(root, 'Microsoft Office', 'Office14', exe)
    );
  }
  for (const candidate of candidates) if (fs.existsSync(candidate)) return candidate;

  // Microsoft 365 / Office Click-to-Run commonly registers the executable path here.
  if (process.platform === 'win32') {
    const { execFileSync } = require('child_process');
    for (const hive of ['HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\' + exe,
                        'HKLM\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\' + exe,
                        'HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\' + exe]) {
      try {
        const output = execFileSync('reg', ['query', hive, '/ve'], { encoding: 'utf8', windowsHide: true });
        const match = output.match(/REG_SZ\s+(.+)/i);
        if (match && fs.existsSync(match[1].trim())) return match[1].trim();
      } catch (_) {}
    }
  }
  return null;
}

function rtfEscape(text) {
  let out = '';
  for (const ch of String(text)) {
    const code = ch.codePointAt(0);
    if (ch === '\\') out += '\\\\';
    else if (ch === '{') out += '\\{';
    else if (ch === '}') out += '\\}';
    else if (ch === '\n') out += '\\par\n';
    else if (code < 128) out += ch;
    else {
      const units = ch.length === 2 ? [ch.charCodeAt(0), ch.charCodeAt(1)] : [code];
      for (const unit of units) out += '\\u' + (unit > 32767 ? unit - 65536 : unit) + '?';
    }
  }
  return out;
}


let officeBarWindow = null;
let mainPlatformWindow = null;
let officeExamState = null;
let officeBarTicker = null;

function updateOfficeBar() {
  if (!officeBarWindow || officeBarWindow.isDestroyed() || !officeExamState) return;
  const task = officeExamState.tasks[officeExamState.index] || null;
  officeBarWindow.webContents.send('exam-bar:update', {
    kind: officeExamState.kind,
    index: officeExamState.index,
    count: officeExamState.tasks.length,
    prompt: task ? (task.prompt || task.title || 'نفّذ المهمة في البرنامج') : 'نفّذ المطلوب في المستند ثم احفظ عملك',
    title: task ? (task.title || '') : '',
    points: task ? Number(task.points || 0) : 0,
    seconds: officeExamState.seconds,
    finished: officeExamState.seconds <= 0
  });
}

function closeOfficeBar() {
  if (officeBarTicker) { clearInterval(officeBarTicker); officeBarTicker = null; }
  officeExamState = null;
  if (officeBarWindow && !officeBarWindow.isDestroyed()) officeBarWindow.close();
  officeBarWindow = null;
}

function createOfficeBar(kind, tasks, durationMinutes) {
  closeOfficeBar();
  const { screen } = require('electron');
  const area = screen.getPrimaryDisplay().bounds;
  const height = OFFICE_BAR_HEIGHT;
  officeExamState = {
    kind, tasks: Array.isArray(tasks) ? tasks : [], index: 0,
    seconds: Math.max(1, Number(durationMinutes) || 20) * 60
  };
  officeBarWindow = new BrowserWindow({
    x: area.x, y: area.y + area.height - height,
    width: area.width, height,
    frame: false, resizable: false, movable: false,
    alwaysOnTop: true, skipTaskbar: true,
    backgroundColor: '#f7f7f2',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true, nodeIntegration: false, sandbox: false
    }
  });
  officeBarWindow.setAlwaysOnTop(true, 'screen-saver');
  officeBarWindow.loadFile(path.join(__dirname, 'office-bar.html'));
  officeBarWindow.webContents.on('did-finish-load', updateOfficeBar);
  officeBarWindow.on('closed', () => {
    restoreWindowsTaskbar();
    officeBarWindow = null;
    if (officeBarTicker) { clearInterval(officeBarTicker); officeBarTicker = null; }
    officeExamState = null;
  });
  officeBarTicker = setInterval(() => {
    if (!officeExamState) return;
    officeExamState.seconds = Math.max(0, officeExamState.seconds - 1);
    updateOfficeBar();
    if (officeExamState.seconds === 0) {
      clearInterval(officeBarTicker); officeBarTicker = null;
      if (officeBarWindow && !officeBarWindow.isDestroyed()) {
        officeBarWindow.webContents.send('exam-bar:time-up');
      }
    }
  }, 1000);
}

ipcMain.on('exam-bar:previous', () => {
  if (!officeExamState) return;
  officeExamState.index = Math.max(0, officeExamState.index - 1);
  updateOfficeBar();
});
ipcMain.on('exam-bar:next', () => {
  if (!officeExamState) return;
  officeExamState.index = Math.min(Math.max(0, officeExamState.tasks.length - 1), officeExamState.index + 1);
  updateOfficeBar();
});
ipcMain.on('exam-bar:finish', () => {
  const completedKind = officeExamState?.kind || 'excel';
  closeOfficeBar();
  if (mainPlatformWindow && !mainPlatformWindow.isDestroyed()) {
    mainPlatformWindow.show();
    mainPlatformWindow.restore();
    mainPlatformWindow.focus();
    mainPlatformWindow.webContents.send('office-exam:returned', { kind: completedKind });
  }
});



function extractDocxText(filePath) {
  const zlib = require('zlib');
  const data = fs.readFileSync(filePath);
  let eocd = -1;
  for (let i = data.length - 22; i >= Math.max(0, data.length - 65558); i--) {
    if (data.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('ملف Word غير صالح أو تالف.');
  const entries = data.readUInt16LE(eocd + 10);
  let cursor = data.readUInt32LE(eocd + 16);
  for (let n = 0; n < entries; n++) {
    if (data.readUInt32LE(cursor) !== 0x02014b50) throw new Error('تعذر قراءة بنية ملف Word.');
    const method = data.readUInt16LE(cursor + 10);
    const compressedSize = data.readUInt32LE(cursor + 20);
    const nameLen = data.readUInt16LE(cursor + 28);
    const extraLen = data.readUInt16LE(cursor + 30);
    const commentLen = data.readUInt16LE(cursor + 32);
    const localOffset = data.readUInt32LE(cursor + 42);
    const name = data.subarray(cursor + 46, cursor + 46 + nameLen).toString('utf8');
    if (name === 'word/document.xml') {
      const localNameLen = data.readUInt16LE(localOffset + 26);
      const localExtraLen = data.readUInt16LE(localOffset + 28);
      const start = localOffset + 30 + localNameLen + localExtraLen;
      const compressed = data.subarray(start, start + compressedSize);
      const xml = (method === 0 ? compressed : method === 8 ? zlib.inflateRawSync(compressed) : null);
      if (!xml) throw new Error('نوع ضغط ملف Word غير مدعوم.');
      return xml.toString('utf8')
        .replace(/<w:tab\b[^>]*\/>/g, ' ')
        .replace(/<w:(?:br|cr)\b[^>]*\/>/g, '\n')
        .replace(/<\/w:p>/g, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
        .replace(/&#(\d+);/g, (_, x) => String.fromCodePoint(Number(x)))
        .replace(/&#x([0-9a-f]+);/gi, (_, x) => String.fromCodePoint(parseInt(x, 16)))
        .replace(/[\t ]+\n/g, '\n').trim();
    }
    cursor += 46 + nameLen + extraLen + commentLen;
  }
  throw new Error('لم يتم العثور على نص المستند داخل ملف Word.');
}

ipcMain.handle('typing:import-word', async () => {
  const result = await dialog.showOpenDialog(mainPlatformWindow || undefined, {
    title: 'استيراد نص الكتابة من Microsoft Word',
    properties: ['openFile'],
    filters: [{ name: 'مستندات Word', extensions: ['docx'] }]
  });
  if (result.canceled || !result.filePaths.length) return { ok: true, canceled: true };
  try {
    const content = extractDocxText(result.filePaths[0]);
    return { ok: true, content, fileName: path.basename(result.filePaths[0]) };
  } catch (error) {
    return { ok: false, message: error.message || 'تعذر استيراد الملف.' };
  }
});

ipcMain.handle('excel:import-files', async () => {
  const result = await dialog.showOpenDialog(mainPlatformWindow || undefined, {
    title: 'اختر ملفات Excel',
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Excel files', extensions: ['xlsx', 'xls', 'xlsm', 'csv'] }]
  });
  if (result.canceled || !result.filePaths.length) return { ok: true, files: [] };
  const folder = path.join(app.getPath('userData'), 'trainer-excel-files');
  fs.mkdirSync(folder, { recursive: true });
  const imported = [];
  for (const source of result.filePaths) {
    const ext = path.extname(source).toLowerCase();
    const base = path.basename(source, ext).replace(/[<>:"/\\|?*]/g, '_').slice(0, 80) || 'Excel';
    const safeName = base + '-' + Date.now() + '-' + Math.random().toString(16).slice(2, 8) + ext;
    const destination = path.join(folder, safeName);
    fs.copyFileSync(source, destination);
    imported.push({ id: safeName, name: path.basename(source), path: destination });
  }
  return { ok: true, files: imported };
});

ipcMain.handle('excel:preview-file', async (_event, suppliedFilePath = '') => {
  if (!suppliedFilePath || !fs.existsSync(suppliedFilePath)) return { ok: false, message: 'ملف Excel غير موجود. أعد إضافته من لوحة المدرب.' };
  const exe = findOfficeExe('excel');
  if (!exe) return { ok: false, message: 'لم يتم العثور على Microsoft Excel مثبتاً على هذا الكمبيوتر.' };
  try {
    const child = spawn(exe, [suppliedFilePath], { detached: true, stdio: 'ignore', windowsHide: false });
    child.unref();
    return { ok: true };
  } catch (error) { return { ok: false, message: 'تعذرت معاينة الملف: ' + error.message }; }
});

ipcMain.handle('office:open', async (_event, kind, suppliedTasks = [], trainee = '', suppliedContent = '', durationMinutes = 20, suppliedFilePath = '') => {
  if (kind !== 'word' && kind !== 'excel') return { ok: false, message: 'نوع البرنامج غير معروف.' };
  const exe = findOfficeExe(kind);
  if (!exe) {
    return { ok: false, message: kind === 'word'
      ? 'لم يتم العثور على Microsoft Word مثبتاً على هذا الكمبيوتر. تأكد من تثبيت Word ثم أعد المحاولة.'
      : 'لم يتم العثور على Microsoft Excel مثبتاً على هذا الكمبيوتر. تأكد من تثبيت Excel ثم أعد المحاولة.' };
  }
  try {
    const folder = path.join(app.getPath('documents'), 'منصة اختبار إدخال البيانات - ملفات التدريب');
    fs.mkdirSync(folder, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const tasks = Array.isArray(suppliedTasks) ? suppliedTasks.slice(0, 100).map((t, i) => `${i + 1}. ${String(t.title || 'مهمة').slice(0, 120)}: ${String(t.prompt || '').slice(0, 1500)} | معيار التنفيذ: ${String(t.expected || '').slice(0, 500)} | الدرجة: ${Number(t.points) || 0}`) : [];
    const content = String(suppliedContent || '').slice(0, 20000);
    let file;
    if (kind === 'word') {
      // Word document shows only the trainer's text; the questions stay in the exam panel below.
      const wordDoc = String(content || '').trim();
      file = path.join(folder, 'اختبار Word - ' + stamp + '.rtf');
      fs.writeFileSync(file, '{\\rtf1\\ansi\\deff0\\uc1{\\fonttbl{\\f0 Arial;}}\\f0\\rtlpar\\fs28 ' + rtfEscape(wordDoc) + '}', 'ascii');
    } else if (suppliedFilePath && fs.existsSync(suppliedFilePath)) {
      file = suppliedFilePath;
    } else {
      const contentRows = content.split(/\r?\n/).map(line => line.split('|').map(cell => String(cell ?? '')));
      const rows = contentRows.length ? contentRows : [['']];
      const xmlEscape = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
      const xmlRows = rows.map(row => '<Row>' + row.map(value => {
        const numeric = /^-?(?:\\d+\\.?\\d*|\\.\\d+)$/.test(value.trim());
        return '<Cell><Data ss:Type="' + (numeric ? 'Number' : 'String') + '">' + xmlEscape(value) + '</Data></Cell>';
      }).join('') + '</Row>').join('');
      const workbook = '<?xml version="1.0" encoding="UTF-8"?>' +
        '<?mso-application progid="Excel.Sheet"?>' +
        '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
        '<Worksheet ss:Name="بيانات الاختبار"><Table ss:ExpandedColumnCount="18" ss:ExpandedRowCount="' + rows.length + '">' + xmlRows + '</Table>' +
        '<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><DisplayRightToLeft/><FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>' +
        '</Worksheet></Workbook>';
      file = path.join(folder, 'تدريب Excel - ' + stamp + '.xml');
      fs.writeFileSync(file, '\uFEFF' + workbook, 'utf8');
    }
    const child = spawn(exe, [file], { detached: true, stdio: 'ignore', windowsHide: false });
    child.unref();
    if (process.platform === 'win32') {
      officeTaskbarHidden = true;
      setTimeout(() => runWindowsOfficeLayout('hide', kind), 900);
      setTimeout(() => runWindowsOfficeLayout('hide', kind), 1800);
    }
    createOfficeBar(kind, suppliedTasks, durationMinutes);
    return { ok: true, message: 'تم فتح ' + (kind === 'word' ? 'Microsoft Word' : 'Microsoft Excel') + ' الأصلي مع شريط الاختبار أسفل الشاشة. حُفظ ملف التدريب في مجلد المستندات.', file };
  } catch (error) {
    return { ok: false, message: 'تعذر تشغيل البرنامج: ' + error.message };
  }
});


// Trainer-only Office assets are copied into Electron userData, never exposed by path to the renderer.
ipcMain.handle('office-assets:import', async (_event, kind, role) => {
  if (!['word','excel'].includes(kind) || !['template','solution'].includes(role)) return { ok:false, message:'نوع الملف غير صالح.' };
  const extensions = kind === 'word' ? ['docx'] : ['xlsx','xlsm','xls','csv'];
  const title = `${kind === 'word' ? 'Word' : 'Excel'} - ${role === 'solution' ? 'الحل النموذجي' : 'نموذج الاختبار'}`;
  const result = await dialog.showOpenDialog(mainPlatformWindow || undefined, {
    title, properties:['openFile'],
    filters:[{ name: kind === 'word' ? 'Word DOCX' : 'Excel', extensions }]
  });
  if (result.canceled || !result.filePaths.length) return { ok:true, canceled:true };
  try {
    const source=result.filePaths[0], ext=path.extname(source).toLowerCase();
    const id=`${kind}_${role}_${Date.now()}_${Math.random().toString(16).slice(2,8)}${ext}`;
    const vault=path.join(app.getPath('userData'),'trainer-office-vault',kind,role);
    fs.mkdirSync(vault,{recursive:true});
    fs.copyFileSync(source,path.join(vault,id));
    return {ok:true,asset:{id,name:path.basename(source),kind,role}};
  } catch(error) { return {ok:false,message:'تعذر حفظ الملف: '+error.message}; }
});
ipcMain.handle('office-assets:delete', async (_event, kind, role, id) => {
  if (!['word','excel'].includes(kind) || !['template','solution'].includes(role) || !/^[a-z0-9_.-]+$/i.test(String(id||''))) return {ok:false};
  const target=path.join(app.getPath('userData'),'trainer-office-vault',kind,role,String(id));
  try { if(fs.existsSync(target)) fs.unlinkSync(target); return {ok:true}; }
  catch(error) { return {ok:false,message:error.message}; }
});

function createWindow() {
  const win = new BrowserWindow({
    width: 1200, height: 850, minWidth: 850, minHeight: 650,
    autoHideMenuBar: true, backgroundColor: '#f3f7f5',
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: false }
  });
  mainPlatformWindow = win;
  win.on('closed', () => { if (mainPlatformWindow === win) mainPlatformWindow = null; });
  win.webContents.on('did-finish-load', () => {
    win.webContents.executeJavaScript("Boolean(window.desktopApp && window.desktopApp.isDesktop)").then(ready => {
      if (!ready) console.error('Desktop bridge is unavailable: check preload.js packaging and launch the installed Windows app, not index.html in a browser.');
    }).catch(error => console.error('Desktop bridge check failed:', error));
  });
  win.loadFile(path.join(__dirname, 'index.html'));
}
app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('before-quit', restoreWindowsTaskbar);
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
