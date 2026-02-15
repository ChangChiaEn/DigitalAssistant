const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  // Window controls
  minimize: () => ipcRenderer.invoke('window:minimize'),
  close: () => ipcRenderer.invoke('window:close'),
  toggleAlwaysOnTop: () => ipcRenderer.invoke('window:toggle-top'),

  // Skills - system actions
  openPath: (p) => ipcRenderer.invoke('skill:open-path', p),
  openUrl: (url) => ipcRenderer.invoke('skill:open-url', url),
  runCommand: (cmd) => ipcRenderer.invoke('skill:run-command', cmd),
  launchApp: (name) => ipcRenderer.invoke('skill:launch-app', name),
  systemInfo: () => ipcRenderer.invoke('skill:system-info'),
  clipboardRead: () => ipcRenderer.invoke('skill:clipboard-read'),
  clipboardWrite: (text) => ipcRenderer.invoke('skill:clipboard-write', text),

  // File operations
  create_docx: (title, content) => ipcRenderer.invoke('skill:create-docx', title, content),
  create_ppt: (title, slidesJson) => ipcRenderer.invoke('skill:create-ppt', title, slidesJson),
  create_xlsx: (title, dataJson) => ipcRenderer.invoke('skill:create-xlsx', title, dataJson),
  write_file: (filename, content) => ipcRenderer.invoke('skill:write-file', filename, content),
  read_file: (filepath) => ipcRenderer.invoke('skill:read-file', filepath),
  list_files: (directory) => ipcRenderer.invoke('skill:list-files', directory),

  // Utility
  take_screenshot: () => ipcRenderer.invoke('skill:take-screenshot'),
  get_datetime: () => ipcRenderer.invoke('skill:get-datetime'),
  search_web: (query) => ipcRenderer.invoke('skill:search-web', query),
  fetch_news: (query, maxResults) => ipcRenderer.invoke('skill:fetch-news', query, maxResults),
  kill_process: (processName) => ipcRenderer.invoke('skill:kill-process', processName),
  set_volume: (level) => ipcRenderer.invoke('skill:set-volume', level),
  notify: (title, message) => ipcRenderer.invoke('skill:notify', title, message),

  // AI Chat
  chat_with_ai: (messagesJson, skillsJson) => ipcRenderer.invoke('skill:chat-with-ai', messagesJson, skillsJson),
  check_health: () => ipcRenderer.invoke('skill:check-health'),

  // Settings
  getSetting: (key) => ipcRenderer.invoke('settings:get', key),
  setSetting: (key, val) => ipcRenderer.invoke('settings:set', key, val),
});
