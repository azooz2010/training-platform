const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('desktopApp', {
  isDesktop: true,
  importExcelFiles: () => ipcRenderer.invoke('excel:import-files'),
  importOfficeAsset: (kind, role) => ipcRenderer.invoke('office-assets:import', kind, role),
  deleteOfficeAsset: (kind, role, id) => ipcRenderer.invoke('office-assets:delete', kind, role, id),
  importTypingWord: () => ipcRenderer.invoke('typing:import-word'),
  previewExcelFile: (filePath) => ipcRenderer.invoke('excel:preview-file', filePath),
  openOffice: (kind, tasks = [], trainee = '', content = '', duration = 20, filePath = '') => ipcRenderer.invoke('office:open', kind, tasks, trainee, content, duration, filePath),
  onOfficeExamReturned: (callback) => ipcRenderer.on('office-exam:returned', (_event, data) => callback(data)),
  examBar: {
    onUpdate: (callback) => ipcRenderer.on('exam-bar:update', (_event, data) => callback(data)),
    onTimeUp: (callback) => ipcRenderer.on('exam-bar:time-up', () => callback()),
    previous: () => ipcRenderer.send('exam-bar:previous'),
    next: () => ipcRenderer.send('exam-bar:next'),
    finish: () => ipcRenderer.send('exam-bar:finish')
  }
});
