// SVG'yi Electron ile render edip PNG ve çok boyutlu ICO üretir.
const { app, BrowserWindow, nativeImage } = require('electron');
const fs = require('fs');
const path = require('path');

const SIZES = [256, 128, 64, 48, 32, 16];

app.whenReady().then(async () => {
  const win = new BrowserWindow({ show: false, width: 256, height: 256, transparent: true, frame: false, webPreferences: { offscreen: true } });
  const svg = fs.readFileSync(path.join(__dirname, 'icon.svg'), 'utf8');
  const html = `<!DOCTYPE html><html><body style="margin:0;background:transparent;overflow:hidden">${svg}</body></html>`;
  await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
  await new Promise((r) => setTimeout(r, 500));
  const img = await win.webContents.capturePage({ x: 0, y: 0, width: 256, height: 256 });
  fs.writeFileSync(path.join(__dirname, 'icon.png'), img.toPNG());

  // ICO: başlık + dizin + PNG verileri
  const pngs = SIZES.map((s) => (s === 256 ? img : img.resize({ width: s, height: s, quality: 'best' })).toPNG());
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0); header.writeUInt16LE(1, 2); header.writeUInt16LE(SIZES.length, 4);
  const dir = Buffer.alloc(16 * SIZES.length);
  let offset = 6 + dir.length;
  SIZES.forEach((s, i) => {
    const e = i * 16;
    dir.writeUInt8(s === 256 ? 0 : s, e); dir.writeUInt8(s === 256 ? 0 : s, e + 1);
    dir.writeUInt8(0, e + 2); dir.writeUInt8(0, e + 3);
    dir.writeUInt16LE(1, e + 4); dir.writeUInt16LE(32, e + 6);
    dir.writeUInt32LE(pngs[i].length, e + 8); dir.writeUInt32LE(offset, e + 12);
    offset += pngs[i].length;
  });
  fs.writeFileSync(path.join(__dirname, 'icon.ico'), Buffer.concat([header, dir, ...pngs]));
  console.log('icon.png ve icon.ico üretildi');
  app.exit(0);
});
