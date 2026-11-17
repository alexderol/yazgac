const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const { readExcel, writeExcel, readWord, writeWord } = require('./office');

let win;

function createWindow() {
  win = new BrowserWindow({
    width: 1200,
    height: 800,
    title: 'Yazgaç',
    icon: path.join(__dirname, 'build', 'icon.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.loadFile(path.join(__dirname, 'index.html'));
  buildMenu();

  // Kaydedilmemiş değişiklik varsa kapatmadan önce renderer'a sor
  win.on('close', (e) => {
    if (forceClose) return;
    e.preventDefault();
    win.webContents.send('request-close');
  });
}

let forceClose = false;
ipcMain.on('close-now', () => {
  forceClose = true;
  if (win) win.close();
});

function send(channel, ...args) {
  if (win) win.webContents.send(channel, ...args);
}

function buildMenu() {
  const template = [
    {
      label: 'Dosya',
      submenu: [
        { label: 'Yeni Word Belgesi', accelerator: 'CmdOrCtrl+N', click: () => send('menu', 'new-word') },
        { label: 'Yeni Excel Tablosu', accelerator: 'CmdOrCtrl+Shift+N', click: () => send('menu', 'new-excel') },
        { type: 'separator' },
        { label: 'Aç...', accelerator: 'CmdOrCtrl+O', click: () => send('menu', 'open') },
        { label: 'Kaydet', accelerator: 'CmdOrCtrl+S', click: () => send('menu', 'save') },
        { label: 'Farklı Kaydet...', accelerator: 'CmdOrCtrl+Shift+S', click: () => send('menu', 'save-as') },
        { type: 'separator' },
        { label: 'Çıkış', role: 'quit' },
      ],
    },
    {
      label: 'Düzen',
      submenu: [
        { label: 'Geri Al', role: 'undo' },
        { label: 'Yinele', role: 'redo' },
        { type: 'separator' },
        { label: 'Kes', role: 'cut' },
        { label: 'Kopyala', role: 'copy' },
        { label: 'Yapıştır', role: 'paste' },
        { label: 'Tümünü Seç', role: 'selectAll' },
      ],
    },
    {
      label: 'Görünüm',
      submenu: [
        { label: 'Yakınlaştır', role: 'zoomIn' },
        { label: 'Uzaklaştır', role: 'zoomOut' },
        { label: 'Sıfırla', role: 'resetZoom' },
        { type: 'separator' },
        { label: 'Geliştirici Araçları', role: 'toggleDevTools' },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// ---------- IPC ----------

const FILTERS = {
  all: [
    { name: 'Ofis Belgeleri', extensions: ['docx', 'xlsx', 'xls', 'csv'] },
    { name: 'Word Belgesi', extensions: ['docx'] },
    { name: 'Excel Tablosu', extensions: ['xlsx', 'xls', 'csv'] },
  ],
  word: [{ name: 'Word Belgesi', extensions: ['docx'] }],
  excel: [
    { name: 'Excel Tablosu', extensions: ['xlsx'] },
    { name: 'CSV', extensions: ['csv'] },
  ],
};

ipcMain.handle('open-file', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    properties: ['openFile'],
    filters: FILTERS.all,
  });
  if (canceled || !filePaths.length) return null;
  const filePath = filePaths[0];
  const ext = path.extname(filePath).toLowerCase();
  try {
    if (ext === '.docx') {
      const html = await readWord(filePath);
      return { type: 'word', path: filePath, html };
    }
    if (['.xlsx', '.xls', '.csv', '.xlsm'].includes(ext)) {
      const sheets = readExcel(filePath);
      return { type: 'excel', path: filePath, sheets };
    }
    return { error: 'Desteklenmeyen dosya türü: ' + ext };
  } catch (e) {
    return { error: 'Dosya açılamadı: ' + e.message };
  }
});

ipcMain.handle('save-file', async (_evt, payload) => {
  // payload: { type, path, saveAs, blocks | sheets }
  let filePath = payload.path;
  if (!filePath || payload.saveAs) {
    const { canceled, filePath: chosen } = await dialog.showSaveDialog(win, {
      defaultPath: filePath || (payload.type === 'word' ? 'Belge.docx' : 'Tablo.xlsx'),
      filters: payload.type === 'word' ? FILTERS.word : FILTERS.excel,
    });
    if (canceled || !chosen) return null;
    filePath = chosen;
  }
  try {
    if (payload.type === 'word') {
      if (!filePath.toLowerCase().endsWith('.docx')) filePath += '.docx';
      await writeWord(filePath, payload.blocks);
    } else {
      const lower = filePath.toLowerCase();
      if (!/\.(xlsx|xls|csv)$/.test(lower)) filePath += '.xlsx';
      writeExcel(filePath, payload.sheets);
    }
    return { path: filePath };
  } catch (e) {
    return { error: 'Kaydedilemedi: ' + e.message };
  }
});

ipcMain.handle('confirm', async (_evt, message) => {
  const { response } = await dialog.showMessageBox(win, {
    type: 'question',
    buttons: ['Kaydet', 'Kaydetme', 'İptal'],
    defaultId: 0,
    cancelId: 2,
    message,
  });
  return response; // 0 kaydet, 1 kaydetme, 2 iptal
});

ipcMain.handle('set-title', (_evt, title) => {
  if (win) win.setTitle(title);
});

app.whenReady().then(createWindow);
app.on('window-all-closed', () => app.quit());
