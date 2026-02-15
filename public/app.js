// ===== Digital Assistant - Frontend Logic (Supports Electron & PyWebView) =====

// Detect environment and create unified API
function createUnifiedAPI() {
  const isElectron = typeof window.electronAPI !== 'undefined';
  const isPyWebView = typeof window.pywebview !== 'undefined';

  if (isElectron) {
    // Electron API wrapper
    return {
      get_setting: async (key) => await window.electronAPI.getSetting(key) || '',
      set_setting: async (key, val) => await window.electronAPI.setSetting(key, val),
      launch_app: async (name) => JSON.stringify(await window.electronAPI.launchApp(name)),
      open_url: async (url) => { await window.electronAPI.openUrl(url); return JSON.stringify({ success: true, message: `已開啟: ${url}` }); },
      open_path: async (path) => { await window.electronAPI.openPath(path); return JSON.stringify({ success: true, message: `已開啟: ${path}` }); },
      run_command: async (cmd) => JSON.stringify(await window.electronAPI.runCommand(cmd)),
      clipboard_read: async () => JSON.stringify(await window.electronAPI.clipboardRead()),
      clipboard_write: async (text) => JSON.stringify(await window.electronAPI.clipboardWrite(text)),
      system_info: async () => JSON.stringify(await window.electronAPI.systemInfo()),
      create_docx: async (title, content) => JSON.stringify(await window.electronAPI.create_docx(title, content)),
      create_ppt: async (title, slidesJson) => JSON.stringify(await window.electronAPI.create_ppt(title, slidesJson)),
      create_xlsx: async (title, dataJson) => JSON.stringify(await window.electronAPI.create_xlsx(title, dataJson)),
      write_file: async (filename, content) => JSON.stringify(await window.electronAPI.write_file(filename, content)),
      read_file: async (filepath) => JSON.stringify(await window.electronAPI.read_file(filepath)),
      list_files: async (directory) => JSON.stringify(await window.electronAPI.list_files(directory)),
      take_screenshot: async () => JSON.stringify(await window.electronAPI.take_screenshot()),
      get_datetime: async () => JSON.stringify(await window.electronAPI.get_datetime()),
      search_web: async (query) => JSON.stringify(await window.electronAPI.search_web(query)),
      fetch_news: async (query, maxResults) => JSON.stringify(await window.electronAPI.fetch_news(query, maxResults)),
      kill_process: async (name) => JSON.stringify(await window.electronAPI.kill_process(name)),
      set_volume: async (level) => JSON.stringify(await window.electronAPI.set_volume(level)),
      notify: async (title, message) => JSON.stringify(await window.electronAPI.notify(title, message)),
      chat_with_ai: async (messagesJson, skillsJson) => await window.electronAPI.chat_with_ai(messagesJson, skillsJson),
      check_health: async () => await window.electronAPI.check_health(),
    };
  } else if (isPyWebView) {
    // PyWebView API (direct access)
    return window.pywebview.api;
  } else {
    console.error('Neither Electron nor PyWebView detected!');
    return null;
  }
}

// Wait for API to be ready
async function waitForAPI() {
  if (typeof window.electronAPI !== 'undefined') {
    return; // Electron is ready immediately
  }
  if (window.pywebview && window.pywebview.api) {
    return; // PyWebView is ready
  }
  return new Promise((resolve) => {
    window.addEventListener('pywebviewready', resolve);
  });
}

class DigitalAssistant {
  constructor() {
    this.api = null; // Unified API reference
    this.apiUrl = '';
    this.isConnected = false;
    this.isRecording = false;
    this.recognition = null;
    this.ttsEnabled = true;
    this.lang = 'zh-TW';
    this.conversationHistory = [];
    this.isElectron = typeof window.electronAPI !== 'undefined';

    this.start();
  }

  async start() {
    await waitForAPI();
    this.api = createUnifiedAPI();
    if (!this.api) {
      console.error('Failed to initialize API');
      return;
    }
    await this.init();
  }

  async init() {
    this.apiUrl = await this.api.get_setting('apiUrl') || '';
    this.ttsEnabled = (await this.api.get_setting('tts')) !== 'off';
    this.lang = await this.api.get_setting('lang') || 'zh-TW';

    this.bindEvents();
    this.initSpeechRecognition();
    if (this.apiUrl) this.checkConnection();
  }

  // ===== Event Binding =====
  bindEvents() {
    // Window controls
    document.getElementById('btnMin').addEventListener('click', () => {
      this.api.minimize_window();
    });
    document.getElementById('btnClose').addEventListener('click', () => {
      this.api.close_window();
    });
    document.getElementById('btnPin').addEventListener('click', () => {
      document.getElementById('btnPin').classList.toggle('active');
    });

    // Text input
    const textInput = document.getElementById('textInput');
    const sendBtn = document.getElementById('sendBtn');
    textInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        this.sendMessage(textInput.value.trim());
      }
    });
    sendBtn.addEventListener('click', () => this.sendMessage(textInput.value.trim()));

    // Mic button - click to start (Python-side auto-stops after speech)
    const micBtn = document.getElementById('micBtn');
    micBtn.addEventListener('click', () => this.startRecording());

    // Skill buttons
    document.getElementById('skillsGrid').addEventListener('click', (e) => {
      const btn = e.target.closest('.skill-btn');
      if (!btn) return;
      this.executeSkillButton(btn.dataset.skill, btn.dataset.arg);
    });

    // Skill hints in chat
    document.getElementById('chatContainer').addEventListener('click', (e) => {
      if (e.target.classList.contains('hint')) {
        this.sendMessage(e.target.dataset.cmd);
      }
    });

    // Skills panel toggle
    document.getElementById('skillsToggle').addEventListener('click', () => {
      document.getElementById('skillsPanel').classList.toggle('collapsed');
    });

    // Settings
    document.getElementById('btnSettings').addEventListener('click', () => this.openSettings());
    document.getElementById('settingSave').addEventListener('click', () => this.saveSettings());
    document.getElementById('settingCancel').addEventListener('click', () => this.closeSettings());
  }

  // ===== Speech Recognition (Python-side) =====
  initSpeechRecognition() {
    // Using Python speech_recognition library instead of browser Web Speech API
    // No initialization needed - Python handles mic access directly
  }

  async startRecording() {
    if (this.isRecording) return;
    this.isRecording = true;
    const micBtn = document.getElementById('micBtn');
    micBtn.classList.add('recording');
    document.getElementById('textInput').value = '聆聽中...';

    try {
      const raw = await this.api.start_listening(this.lang);
      const result = JSON.parse(raw);
      this.isRecording = false;
      micBtn.classList.remove('recording');

      if (result.success && result.text) {
        document.getElementById('textInput').value = result.text;
        this.sendMessage(result.text);
      } else {
        document.getElementById('textInput').value = '';
        if (result.error) {
          this.addMessage('assistant', `語音辨識: ${result.error}`, true);
        }
      }
    } catch (err) {
      this.isRecording = false;
      micBtn.classList.remove('recording');
      document.getElementById('textInput').value = '';
      this.addMessage('assistant', `語音辨識失敗: ${err.message || err}`, true);
    }
  }

  stopRecording() {
    // Python-side recording auto-stops after speech ends or timeout
  }

  // ===== TTS =====
  speak(text) {
    if (!this.ttsEnabled || !window.speechSynthesis) return;
    window.speechSynthesis.cancel();
    const utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = this.lang;
    utterance.rate = 1.1;
    window.speechSynthesis.speak(utterance);
  }

  // ===== Chat =====
  async sendMessage(text) {
    if (!text) return;
    document.getElementById('textInput').value = '';

    this.addMessage('user', text);

    // Check if it's a local skill command first
    const skillResult = await this.tryLocalSkill(text);
    if (skillResult) {
      this.addMessage('assistant', skillResult.message, skillResult.isSkill);
      this.speak(skillResult.speakText || skillResult.message);
      return;
    }

    // Send to AI backend via Python
    this.showTyping();
    try {
      const reply = await this.callAI(text);
      this.hideTyping();

      if (reply.skill) {
        const result = await this.executeAISkill(reply.skill, reply.args);
        this.addMessage('assistant', reply.text + (result ? `\n<div class="skill-result">${result}</div>` : ''));
        this.speak(reply.text);
        
        // If fetch_news was executed, add result to conversation history
        // Let AI decide what to do next based on the original user request
        if (reply.skill === 'fetch_news' && result) {
          this.conversationHistory.push({
            role: 'assistant',
            content: `已搜尋到以下資料：\n${result}`
          });
        }
      } else {
        this.addMessage('assistant', reply.text || reply);
        this.speak(reply.text || reply);
      }
    } catch (err) {
      this.hideTyping();
      this.addMessage('assistant', `[Error]無法連線到 AI 後端。\n${err.message || err}\n\n請確認 Colab 是否已啟動。`);
    }
  }

  // ===== Local Skill Detection =====
  async tryLocalSkill(text) {
    const lower = text.toLowerCase();

    const launchPatterns = [
      { patterns: ['開啟記事本', '打開記事本', 'open notepad'], app: 'notepad', name: '記事本' },
      { patterns: ['開啟計算機', '打開計算機', 'open calculator'], app: 'calculator', name: '計算機' },
      { patterns: ['開啟檔案總管', '打開檔案總管', 'open explorer'], app: 'explorer', name: '檔案總管' },
      { patterns: ['開啟瀏覽器', '打開瀏覽器', 'open browser'], app: 'browser', name: '瀏覽器' },
      { patterns: ['開啟vscode', '打開vscode', 'open vscode', '開啟編輯器'], app: 'vscode', name: 'VSCode' },
      { patterns: ['開啟終端', '打開終端', 'open terminal'], app: 'terminal', name: '終端機' },
      { patterns: ['開啟discord', '打開discord'], app: 'discord', name: 'Discord' },
      { patterns: ['開啟spotify', '打開spotify'], app: 'spotify', name: 'Spotify' },
    ];

    for (const item of launchPatterns) {
      if (item.patterns.some(p => lower.includes(p))) {
        await this.api.launch_app(item.app);
        return { message: `已為你開啟 ${item.name}`, speakText: `已開啟${item.name}`, isSkill: true };
      }
    }

    if (['系統狀態', '系統資訊', 'system info', '電腦狀態'].some(p => lower.includes(p))) {
      const raw = await this.api.system_info();
      const result = JSON.parse(raw);
      const info = JSON.parse(result.message);
      const msg = `系統資訊\n主機: ${info.hostname}\nCPU: ${info.cpus} 核心\n記憶體: ${info.freeMemory || '?'} / ${info.totalMemory || '?'}`;
      return { message: msg, speakText: `你的電腦有${info.cpus}個CPU核心`, isSkill: true };
    }

    if (['讀取剪貼簿', '剪貼簿內容', 'clipboard', '貼上內容'].some(p => lower.includes(p))) {
      const raw = await this.api.clipboard_read();
      const result = JSON.parse(raw);
      return { message: `剪貼簿內容:\n${result.message || '(空)'}`, speakText: '已讀取剪貼簿', isSkill: true };
    }

    // Screenshot
    if (['截圖', '螢幕截圖', 'screenshot', '截取畫面'].some(p => lower.includes(p))) {
      const raw = await this.api.take_screenshot();
      const result = JSON.parse(raw);
      return { message: result.message, speakText: '已截圖', isSkill: true };
    }

    // Date/Time
    if (['現在幾點', '目前時間', '今天日期', '幾號', 'what time', 'what date', '現在時間'].some(p => lower.includes(p))) {
      const raw = await this.api.get_datetime();
      const result = JSON.parse(raw);
      const info = JSON.parse(result.message);
      return { message: info.full, speakText: info.full, isSkill: true };
    }

    // Search web (but NOT if user wants to create presentation/document)
    // If user says "搜尋XX然後做成簡報/文件"， let AI handle it with fetch_news
    if (!text.match(/(?:然後|再|接著|並|且).*(?:做成|製作|建立|產生).*(?:簡報|文件|報告|PPT|Word|Excel)/i)) {
      const searchMatch = text.match(/(?:搜尋|搜索|查詢|search|幫我查|幫我搜)\s*(.+)/i);
      if (searchMatch) {
        await this.api.search_web(searchMatch[1]);
        return { message: `已搜尋: ${searchMatch[1]}`, speakText: `已搜尋${searchMatch[1]}`, isSkill: true };
      }
    }

    // List files
    if (['列出檔案', '檔案列表', 'list files', '查看檔案'].some(p => lower.includes(p))) {
      const raw = await this.api.list_files('');
      const result = JSON.parse(raw);
      return { message: `工作目錄檔案:\n${result.message}`, speakText: '已列出檔案', isSkill: true };
    }

    // URL
    const urlMatch = text.match(/(?:開啟|打開|open)\s*(https?:\/\/\S+)/i);
    if (urlMatch) {
      await this.api.open_url(urlMatch[1]);
      return { message: `已開啟: ${urlMatch[1]}`, speakText: '已開啟連結', isSkill: true };
    }

    return null;
  }

  // ===== AI Skill Execution =====
  async executeAISkill(skillName, args) {
    try {
      let raw;
      switch (skillName) {
        case 'launch_app':
          raw = await this.api.launch_app(args.name);
          return JSON.parse(raw).message;
        case 'open_url':
          await this.api.open_url(args.url);
          return `已開啟 ${args.url}`;
        case 'open_path':
          await this.api.open_path(args.path);
          return `已開啟 ${args.path}`;
        case 'run_command':
          raw = await this.api.run_command(args.command);
          return JSON.parse(raw).message;
        case 'clipboard_write':
          await this.api.clipboard_write(args.text);
          return '已複製到剪貼簿';
        case 'clipboard_read':
          raw = await this.api.clipboard_read();
          return JSON.parse(raw).message;
        case 'system_info':
          raw = await this.api.system_info();
          return JSON.parse(raw).message;
        case 'create_ppt': {
          const slidesStr = typeof args.slides_json === 'string' ? args.slides_json : JSON.stringify(args.slides_json);
          raw = await this.api.create_ppt(args.title, slidesStr, args.theme || 'dark');
          return JSON.parse(raw).message;
        }
        case 'create_docx':
          raw = await this.api.create_docx(args.title, args.content);
          return JSON.parse(raw).message;
        case 'create_xlsx': {
          const dataStr = typeof args.data_json === 'string' ? args.data_json : JSON.stringify(args.data_json);
          raw = await this.api.create_xlsx(args.title, dataStr);
          return JSON.parse(raw).message;
        }
        case 'write_file':
          raw = await this.api.write_file(args.filename, args.content);
          return JSON.parse(raw).message;
        case 'read_file':
          raw = await this.api.read_file(args.filepath);
          return JSON.parse(raw).message;
        case 'list_files':
          raw = await this.api.list_files(args.directory || '');
          return JSON.parse(raw).message;
        case 'take_screenshot':
          raw = await this.api.take_screenshot();
          return JSON.parse(raw).message;
        case 'get_datetime':
          raw = await this.api.get_datetime();
          return JSON.parse(JSON.parse(raw).message).full;
        case 'search_web':
          raw = await this.api.search_web(args.query);
          return JSON.parse(raw).message;
        case 'fetch_news':
          raw = await this.api.fetch_news(args.query, String(args.max_results || 5));
          return JSON.parse(raw).message;
        case 'kill_process':
          raw = await this.api.kill_process(args.process_name);
          return JSON.parse(raw).message;
        case 'set_volume':
          raw = await this.api.set_volume(String(args.level));
          return JSON.parse(raw).message;
        case 'notify':
          raw = await this.api.notify(args.title, args.message);
          return JSON.parse(raw).message;
        default:
          return `未知技能: ${skillName}`;
      }
    } catch (err) {
      return `技能執行失敗: ${err.message || err}`;
    }
  }

  // ===== Skill Button Execution =====
  async executeSkillButton(skill, arg) {
    switch (skill) {
      case 'launch-app':
        await this.api.launch_app(arg);
        this.addMessage('assistant', `已啟動: ${arg}`, true);
        break;
      case 'system-info': {
        const raw = await this.api.system_info();
        const result = JSON.parse(raw);
        const info = JSON.parse(result.message);
        this.addMessage('assistant',
          `系統資訊\n主機: ${info.hostname}\nCPU: ${info.cpus} 核心\n記憶體: ${info.freeMemory || '?'} / ${info.totalMemory || '?'}`,
          true
        );
        break;
      }
      case 'clipboard-read': {
        const raw = await this.api.clipboard_read();
        const result = JSON.parse(raw);
        this.addMessage('assistant', `剪貼簿: ${result.message || '(空)'}`, true);
        break;
      }
      case 'screenshot': {
        const raw = await this.api.take_screenshot();
        const result = JSON.parse(raw);
        this.addMessage('assistant', result.message, true);
        break;
      }
      case 'datetime': {
        const raw = await this.api.get_datetime();
        const result = JSON.parse(raw);
        const info = JSON.parse(result.message);
        this.addMessage('assistant', info.full, true);
        break;
      }
      case 'list-files': {
        const raw = await this.api.list_files('');
        const result = JSON.parse(raw);
        this.addMessage('assistant', `工作目錄:\n${result.message}`, true);
        break;
      }
      case 'search-web': {
        const query = arg || prompt('搜尋關鍵字:');
        if (query) {
          await this.api.search_web(query);
          this.addMessage('assistant', `已搜尋: ${query}`, true);
        }
        break;
      }
    }
  }

  // ===== AI Backend Communication =====
  async callAI(userMessage) {
    this.conversationHistory.push({ role: 'user', content: userMessage });

    if (this.conversationHistory.length > 40) {
      this.conversationHistory = this.conversationHistory.slice(-40);
    }

    const skills = [
      { name: 'launch_app', description: '啟動應用程式', params: { name: 'string' } },
      { name: 'open_url', description: '開啟網址', params: { url: 'string' } },
      { name: 'open_path', description: '開啟檔案或資料夾', params: { path: 'string' } },
      { name: 'run_command', description: '執行系統指令', params: { command: 'string' } },
      { name: 'clipboard_write', description: '寫入剪貼簿', params: { text: 'string' } },
      { name: 'clipboard_read', description: '讀取剪貼簿', params: {} },
      { name: 'system_info', description: '取得系統資訊', params: {} },
      { name: 'create_ppt', description: '建立PowerPoint簡報。theme可選: dark(科技), corporate(商務白底), nature(自然綠), warm(暖色), ocean(海洋藍), minimal(極簡黑白)。根據主題自動選擇最適合的theme。', params: { title: 'string', slides_json: '[{"title":"...","content":"..."}]', theme: 'string' } },
      { name: 'create_docx', description: '建立Word文件', params: { title: 'string', content: 'string' } },
      { name: 'create_xlsx', description: '建立Excel試算表', params: { title: 'string', data_json: '[["col1","col2"],["val1","val2"]]' } },
      { name: 'write_file', description: '寫入文字檔案', params: { filename: 'string', content: 'string' } },
      { name: 'read_file', description: '讀取檔案內容', params: { filepath: 'string' } },
      { name: 'list_files', description: '列出目錄中的檔案', params: { directory: 'string' } },
      { name: 'take_screenshot', description: '截取螢幕截圖', params: {} },
      { name: 'get_datetime', description: '取得目前日期時間', params: {} },
      { name: 'search_web', description: '用瀏覽器搜尋網頁', params: { query: 'string' } },
      { name: 'fetch_news', description: '搜尋網路資訊並回傳內容摘要(可用於蒐集資料再製作簡報/文件)', params: { query: 'string', max_results: 'number(預設5)' } },
      { name: 'kill_process', description: '終止指定程序', params: { process_name: 'string' } },
      { name: 'set_volume', description: '設定系統音量(0-100)', params: { level: 'number' } },
      { name: 'notify', description: '發送Windows桌面通知', params: { title: 'string', message: 'string' } },
    ];

    // Call backend which handles the HTTP request to Colab
    const rawResponse = await this.api.chat_with_ai(
      JSON.stringify(this.conversationHistory),
      JSON.stringify(skills)
    );

    let data = typeof rawResponse === 'string' ? JSON.parse(rawResponse) : rawResponse;

    // Colab double-wraps: data.text may itself be a JSON string - unwrap it
    if (data.text && typeof data.text === 'string' && data.text.trimStart().startsWith('{')) {
      try {
        const inner = JSON.parse(data.text);
        if (inner && typeof inner === 'object' && ('text' in inner || 'skill' in inner)) {
          data = inner;
        }
      } catch (e) {}
    }

    // Extract embedded JSON with skill from text
    if (data.text && typeof data.text === 'string') {
      try {
        const jsonMatch = data.text.match(/\{[\s\S]*"skill"\s*:[\s\S]*"args"\s*:[\s\S]*\}/);
        if (jsonMatch) {
          const embedded = JSON.parse(jsonMatch[0]);
          if (embedded.skill) {
            data.skill = embedded.skill;
            data.args = embedded.args || {};
            data.text = (embedded.text || data.text.replace(jsonMatch[0], '')).trim() || `正在執行: ${data.skill}`;
          }
        }
      } catch (e) {}
    }

    // Empty string skill = no skill
    if (!data.skill) data.skill = null;

    this.conversationHistory.push({ role: 'assistant', content: data.text || '' });
    return data;
  }

  // ===== Connection Check =====
  async checkConnection() {
    const dot = document.getElementById('statusDot');
    const textEl = document.getElementById('statusText');
    dot.className = 'status-dot connecting';
    textEl.textContent = '連線中...';

    const raw = await this.api.check_health();
    const result = typeof raw === 'string' ? JSON.parse(raw) : raw;

    if (result.connected) {
      dot.className = 'status-dot connected';
      textEl.textContent = result.message;
      this.isConnected = true;
    } else {
      dot.className = 'status-dot';
      textEl.textContent = result.message;
      this.isConnected = false;
    }
  }

  // ===== UI Helpers =====
  addMessage(role, text, isSkillResult = false) {
    const container = document.getElementById('chatContainer');
    const div = document.createElement('div');
    div.className = `message ${role}`;

    const content = document.createElement('div');
    content.className = 'message-content';

    if (isSkillResult) {
      content.innerHTML = `<div class="skill-result">${this.escapeHtml(text)}</div>`;
    } else if (text.includes('<div')) {
      content.innerHTML = text;
    } else {
      content.textContent = text;
    }

    div.appendChild(content);
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
  }

  showTyping() {
    const container = document.getElementById('chatContainer');
    const div = document.createElement('div');
    div.className = 'message assistant';
    div.id = 'typingMsg';
    div.innerHTML = `<div class="message-content"><div class="typing-indicator"><span></span><span></span><span></span></div></div>`;
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
  }

  hideTyping() {
    document.getElementById('typingMsg')?.remove();
  }

  escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // ===== Settings =====
  openSettings() {
    document.getElementById('settingApiUrl').value = this.apiUrl;
    document.getElementById('settingTts').value = this.ttsEnabled ? 'browser' : 'off';
    document.getElementById('settingLang').value = this.lang;
    document.getElementById('settingsModal').classList.add('active');
  }

  async saveSettings() {
    this.apiUrl = document.getElementById('settingApiUrl').value.trim().replace(/\/$/, '');
    this.ttsEnabled = document.getElementById('settingTts').value !== 'off';
    this.lang = document.getElementById('settingLang').value;

    await this.api.set_setting('apiUrl', this.apiUrl);
    await this.api.set_setting('tts', this.ttsEnabled ? 'browser' : 'off');
    await this.api.set_setting('lang', this.lang);

    if (this.recognition) this.recognition.lang = this.lang;

    this.closeSettings();
    if (this.apiUrl) this.checkConnection();
  }

  closeSettings() {
    document.getElementById('settingsModal').classList.remove('active');
  }
}

// ===== Start =====
const assistant = new DigitalAssistant();
