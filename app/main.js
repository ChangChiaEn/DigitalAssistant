const { app, BrowserWindow, ipcMain, shell, clipboard } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { exec } = require('child_process');

// Simple JSON store
let SETTINGS_PATH;

function getSettingsPath() {
  if (!SETTINGS_PATH) {
    SETTINGS_PATH = path.join(app.getPath('userData'), 'settings.json');
  }
  return SETTINGS_PATH;
}

function loadSettings() {
  try { return JSON.parse(fs.readFileSync(getSettingsPath(), 'utf8')); }
  catch { return {}; }
}

function saveSettings(data) {
  fs.writeFileSync(getSettingsPath(), JSON.stringify(data, null, 2));
}

let mainWindow = null;
let settings = {};

function createWindow() {
  settings = loadSettings();

  mainWindow = new BrowserWindow({
    width: 420,
    height: 680,
    frame: false,
    transparent: true,
    resizable: true,
    alwaysOnTop: settings.alwaysOnTop !== false,
    skipTaskbar: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      // Enable media permissions for speech recognition
      permissions: ['microphone', 'camera']
    }
  });

  // Request microphone permission for speech recognition
  const { session } = require('electron');
  session.defaultSession.setPermissionRequestHandler((webContents, permission, callback) => {
    // Allow all media-related permissions (microphone, camera)
    if (permission === 'media' || permission === 'microphone' || permission === 'camera') {
      console.log('Granting permission:', permission);
      callback(true);
    } else {
      callback(false);
    }
  });

  // Grant media permissions automatically
  session.defaultSession.setPermissionCheckHandler((webContents, permission, requestingOrigin) => {
    if (permission === 'media' || permission === 'microphone' || permission === 'camera') {
      return true;
    }
    return false;
  });

  mainWindow.loadFile(path.join(__dirname, '..', 'public', 'index.html'));

  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  }

  mainWindow.on('closed', () => { mainWindow = null; });
}

// ===== System Skills (IPC Handlers) =====

ipcMain.handle('skill:open-path', async (event, filePath) => {
  try {
    await shell.openPath(filePath);
    return { success: true, message: `已開啟: ${filePath}` };
  } catch (err) {
    return { success: false, message: err.message };
  }
});

ipcMain.handle('skill:open-url', async (event, url) => {
  try {
    await shell.openExternal(url);
    return { success: true, message: `已開啟瀏覽器: ${url}` };
  } catch (err) {
    return { success: false, message: err.message };
  }
});

ipcMain.handle('skill:run-command', async (event, command) => {
  return new Promise((resolve) => {
    exec(command, { timeout: 15000, encoding: 'utf8' }, (err, stdout, stderr) => {
      if (err) {
        resolve({ success: false, message: stderr || err.message });
      } else {
        resolve({ success: true, message: stdout.trim() });
      }
    });
  });
});

ipcMain.handle('skill:launch-app', async (event, appName) => {
  const appMap = {
    'notepad': 'notepad.exe',
    '記事本': 'notepad.exe',
    'calculator': 'calc.exe',
    '計算機': 'calc.exe',
    'browser': 'start https://www.google.com',
    '瀏覽器': 'start https://www.google.com',
    'explorer': 'explorer.exe',
    '檔案總管': 'explorer.exe',
    'cmd': 'cmd.exe',
    'terminal': 'wt.exe',
    '終端機': 'wt.exe',
    'vscode': 'code',
    'spotify': 'start spotify:',
    'discord': 'start discord:',
  };

  const cmd = appMap[appName.toLowerCase()] || appName;
  return new Promise((resolve) => {
    exec(`start "" "${cmd}"`, { shell: true }, (err) => {
      if (err) {
        exec(cmd, { shell: true }, (err2) => {
          if (err2) resolve({ success: false, message: `無法啟動: ${appName}` });
          else resolve({ success: true, message: `已啟動: ${appName}` });
        });
      } else {
        resolve({ success: true, message: `已啟動: ${appName}` });
      }
    });
  });
});

ipcMain.handle('skill:system-info', async () => {
  return {
    success: true,
    message: JSON.stringify({
      platform: os.platform(),
      hostname: os.hostname(),
      cpus: os.cpus().length,
      totalMemory: `${(os.totalmem() / 1e9).toFixed(1)} GB`,
      freeMemory: `${(os.freemem() / 1e9).toFixed(1)} GB`,
      uptime: `${(os.uptime() / 3600).toFixed(1)} hours`
    })
  };
});

ipcMain.handle('skill:clipboard-read', async () => {
  return { success: true, message: clipboard.readText() };
});

ipcMain.handle('skill:clipboard-write', async (event, text) => {
  clipboard.writeText(text);
  return { success: true, message: '已複製到剪貼簿' };
});

// File operations
const WORKSPACE = path.join(os.homedir(), 'Desktop', 'AssistantOutput');
if (!fs.existsSync(WORKSPACE)) fs.mkdirSync(WORKSPACE, { recursive: true });

ipcMain.handle('skill:create-docx', async (event, title, content) => {
  try {
    const { execSync } = require('child_process');
    const filename = `${title.replace(/[^\w\s]/g, '_')}_${Date.now()}.docx`;
    const filepath = path.join(WORKSPACE, filename);
    
    // Use PowerShell to create Word document
    const psScript = `
      $word = New-Object -ComObject Word.Application
      $word.Visible = $false
      $doc = $word.Documents.Add()
      $doc.Content.Text = "${title}\n\n${content.replace(/"/g, '\\"').replace(/\n/g, '\\n')}"
      $doc.SaveAs([ref]"${filepath.replace(/\\/g, '\\\\')}")
      $doc.Close()
      $word.Quit()
    `;
    execSync(`powershell -Command "${psScript}"`, { timeout: 10000 });
    shell.openPath(filepath);
    return { success: true, message: `已建立並開啟: ${filepath}` };
  } catch (err) {
    // Fallback: create as text file
    try {
      const filename = `${title.replace(/[^\w\s]/g, '_')}_${Date.now()}.txt`;
      const filepath = path.join(WORKSPACE, filename);
      fs.writeFileSync(filepath, `${title}\n\n${content}`, 'utf8');
      shell.openPath(filepath);
      return { success: true, message: `已建立並開啟: ${filepath}` };
    } catch (e) {
      return { success: false, message: `建立文件失敗: ${err.message}` };
    }
  }
});

ipcMain.handle('skill:create-ppt', async (event, title, slidesJson) => {
  try {
    const slides = typeof slidesJson === 'string' ? JSON.parse(slidesJson) : slidesJson;
    const filename = `${title.replace(/[^\w\s]/g, '_')}_${Date.now()}.txt`;
    const filepath = path.join(WORKSPACE, filename);
    let content = `${title}\n\n`;
    slides.forEach((s, i) => {
      content += `第 ${i + 1} 頁: ${s.title || ''}\n${s.content || ''}\n\n`;
    });
    fs.writeFileSync(filepath, content, 'utf8');
    shell.openPath(filepath);
    return { success: true, message: `已建立並開啟: ${filepath}` };
  } catch (err) {
    return { success: false, message: `建立簡報失敗: ${err.message}` };
  }
});

ipcMain.handle('skill:create-xlsx', async (event, title, dataJson) => {
  try {
    const data = typeof dataJson === 'string' ? JSON.parse(dataJson) : dataJson;
    const filename = `${title.replace(/[^\w\s]/g, '_')}_${Date.now()}.csv`;
    const filepath = path.join(WORKSPACE, filename);
    const csv = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');
    fs.writeFileSync(filepath, '\ufeff' + csv, 'utf8'); // BOM for Excel
    shell.openPath(filepath);
    return { success: true, message: `已建立並開啟: ${filepath}` };
  } catch (err) {
    return { success: false, message: `建立試算表失敗: ${err.message}` };
  }
});

ipcMain.handle('skill:write-file', async (event, filename, content) => {
  try {
    const filepath = path.join(WORKSPACE, filename);
    fs.writeFileSync(filepath, content, 'utf8');
    return { success: true, message: `已儲存: ${filepath}` };
  } catch (err) {
    return { success: false, message: `寫入檔案失敗: ${err.message}` };
  }
});

ipcMain.handle('skill:read-file', async (event, filepath) => {
  try {
    const content = fs.readFileSync(filepath, 'utf8').substring(0, 10000);
    return { success: true, message: content };
  } catch (err) {
    return { success: false, message: `讀取檔案失敗: ${err.message}` };
  }
});

ipcMain.handle('skill:list-files', async (event, directory) => {
  try {
    const dir = directory || WORKSPACE;
    const items = fs.readdirSync(dir).slice(0, 50).map(name => {
      const full = path.join(dir, name);
      const stat = fs.statSync(full);
      const size = stat.isDirectory() ? '[DIR]' : `${stat.size}`.padStart(8);
      return `${size} ${name}`;
    });
    return { success: true, message: items.join('\n') || '(empty)' };
  } catch (err) {
    return { success: false, message: `列出檔案失敗: ${err.message}` };
  }
});

ipcMain.handle('skill:take-screenshot', async () => {
  try {
    const { execSync } = require('child_process');
    const filename = `screenshot_${Date.now()}.png`;
    const filepath = path.join(WORKSPACE, filename);
    // Use PowerShell to take screenshot
    execSync(`powershell -Command "Add-Type -AssemblyName System.Windows.Forms,System.Drawing; $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds; $bmp = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height; $graphics = [System.Drawing.Graphics]::FromImage($bmp); $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size); $bmp.Save('${filepath.replace(/\\/g, '\\\\')}'); $graphics.Dispose(); $bmp.Dispose()"`, { timeout: 5000 });
    shell.openPath(filepath);
    return { success: true, message: `已截圖: ${filepath}` };
  } catch (err) {
    return { success: false, message: `截圖失敗: ${err.message}` };
  }
});

ipcMain.handle('skill:get-datetime', async () => {
  const now = new Date();
  const weekdays = ['日', '一', '二', '三', '四', '五', '六'];
  const data = {
    date: now.toISOString().split('T')[0],
    time: now.toTimeString().split(' ')[0],
    weekday: weekdays[now.getDay()],
    full: `${now.getFullYear()}年${String(now.getMonth() + 1).padStart(2, '0')}月${String(now.getDate()).padStart(2, '0')}日 ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')} 星期${weekdays[now.getDay()]}`
  };
  return { success: true, message: JSON.stringify(data) };
});

ipcMain.handle('skill:search-web', async (event, query) => {
  try {
    const url = `https://www.google.com/search?q=${encodeURIComponent(query)}`;
    shell.openExternal(url);
    return { success: true, message: `已搜尋: ${query}` };
  } catch (err) {
    return { success: false, message: `搜尋失敗: ${err.message}` };
  }
});

ipcMain.handle('skill:fetch-news', async (event, query, maxResults = '5') => {
  try {
    // Use Python script to fetch news (since Electron doesn't have duckduckgo-search)
    const { execSync } = require('child_process');
    const scriptPath = path.join(__dirname, '..', 'main.py');
    const pythonCmd = `python -c "from main import AssistantAPI; api = AssistantAPI(); print(api.fetch_news('${query.replace(/'/g, "\\'")}', '${maxResults}'))"`;
    
    const result = execSync(pythonCmd, { 
      encoding: 'utf8', 
      timeout: 30000,
      cwd: path.join(__dirname, '..')
    });
    return JSON.parse(result.trim());
  } catch (err) {
    return { success: false, message: `搜尋失敗: ${err.message}` };
  }
});

ipcMain.handle('skill:kill-process', async (event, processName) => {
  return new Promise((resolve) => {
    exec(`taskkill /f /im ${processName}`, { timeout: 10000 }, (err, stdout, stderr) => {
      if (err) {
        resolve({ success: false, message: stderr || err.message });
      } else {
        resolve({ success: true, message: stdout.trim() || `已終止: ${processName}` });
      }
    });
  });
});

ipcMain.handle('skill:set-volume', async (event, level) => {
  try {
    const { execSync } = require('child_process');
    const psCmd = `
      $wshShell = New-Object -ComObject WScript.Shell;
      1..50 | ForEach-Object { $wshShell.SendKeys([char]174) };
      1..${Math.floor(Number(level) / 2)} | ForEach-Object { $wshShell.SendKeys([char]175) }
    `;
    execSync(`powershell -Command "${psCmd}"`, { timeout: 10000 });
    return { success: true, message: `音量已設定為 ${level}%` };
  } catch (err) {
    return { success: false, message: `設定音量失敗: ${err.message}` };
  }
});

ipcMain.handle('skill:notify', async (event, title, message) => {
  try {
    const { execSync } = require('child_process');
    const psCmd = `
      [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
      $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
      $textNodes = $template.GetElementsByTagName("text")
      $textNodes.Item(0).AppendChild($template.CreateTextNode("${title.replace(/"/g, '\\"')}")) | Out-Null
      $textNodes.Item(1).AppendChild($template.CreateTextNode("${message.replace(/"/g, '\\"')}")) | Out-Null
      $toast = [Windows.UI.Notifications.ToastNotification]::new($template)
      [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Digital Assistant").Show($toast)
    `;
    execSync(`powershell -Command "${psCmd}"`, { timeout: 10000 });
    return { success: true, message: `已發送通知: ${title}` };
  } catch (err) {
    return { success: false, message: `發送通知失敗: ${err.message}` };
  }
});

// AI Chat
ipcMain.handle('skill:chat-with-ai', async (event, messagesJson, skillsJson) => {
  const fetch = require('node-fetch');
  settings = loadSettings();
  const apiUrl = settings.apiUrl || '';
  
  if (!apiUrl) {
    return JSON.stringify({ text: '[Error] 尚未設定 Colab API URL，請點擊齒輪設定。' });
  }

  try {
    const messages = typeof messagesJson === 'string' ? JSON.parse(messagesJson) : messagesJson;
    const skills = typeof skillsJson === 'string' ? JSON.parse(skillsJson) : skillsJson;
    
    const response = await fetch(`${apiUrl}/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages, skills })
    });

    if (response.ok) {
      return JSON.stringify(await response.json());
    } else {
      return JSON.stringify({ text: `[Error] API: ${response.status}` });
    }
  } catch (err) {
    if (err.code === 'ECONNREFUSED' || err.name === 'FetchError' || err.code === 'ENOTFOUND') {
      return JSON.stringify({ text: '[Error] 無法連線到 AI 後端，請確認 Colab 是否已啟動。' });
    }
    return JSON.stringify({ text: `[Error] ${err.message}` });
  }
});

ipcMain.handle('skill:check-health', async () => {
  const fetch = require('node-fetch');
  settings = loadSettings();
  const apiUrl = settings.apiUrl || '';
  
  if (!apiUrl) {
    return JSON.stringify({ connected: false, message: '未設定' });
  }

  try {
    const response = await fetch(`${apiUrl}/health`);
    if (response.ok) {
      const data = await response.json();
      return JSON.stringify({ connected: true, message: `已連線 - ${data.model || 'AI Ready'}` });
    }
  } catch (err) {
    // Ignore connection errors
  }
  
  return JSON.stringify({ connected: false, message: '未連線 - 請啟動 Colab' });
});

// Window controls
ipcMain.handle('window:minimize', () => { mainWindow?.minimize(); });
ipcMain.handle('window:close', () => { mainWindow?.close(); });
ipcMain.handle('window:toggle-top', () => {
  const current = mainWindow?.isAlwaysOnTop();
  mainWindow?.setAlwaysOnTop(!current);
  settings.alwaysOnTop = !current;
  saveSettings(settings);
  return !current;
});

// Settings
ipcMain.handle('settings:get', (event, key) => {
  settings = loadSettings();
  return settings[key];
});

ipcMain.handle('settings:set', (event, key, value) => {
  settings = loadSettings();
  settings[key] = value;
  saveSettings(settings);
});

// ===== App Lifecycle =====
app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
