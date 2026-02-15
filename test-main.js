
const electron = require('electron');
console.log('electron type:', typeof electron);
console.log('electron keys:', Object.keys(electron));
console.log('app:', typeof electron.app);
console.log('ipcMain:', typeof electron.ipcMain);
if (electron.app) {
  electron.app.whenReady().then(() => { console.log('ready!'); electron.app.quit(); });
} else {
  console.log('app is undefined, electron value:', electron);
  process.exit(1);
}
