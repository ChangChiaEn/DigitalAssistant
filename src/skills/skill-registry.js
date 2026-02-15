/**
 * Skill Registry - Manages all available skills for the AI assistant.
 * Skills are actions the AI can invoke on the local system.
 * New skills can be registered dynamically.
 */

class SkillRegistry {
  constructor() {
    this.skills = new Map();
    this.registerBuiltinSkills();
  }

  register(name, definition) {
    this.skills.set(name, {
      name,
      description: definition.description,
      parameters: definition.parameters || {},
      handler: definition.handler,
      category: definition.category || 'general',
      confirmRequired: definition.confirmRequired || false,
    });
  }

  async execute(name, args = {}) {
    const skill = this.skills.get(name);
    if (!skill) {
      return { success: false, error: `Skill not found: ${name}` };
    }
    try {
      const result = await skill.handler(args);
      return { success: true, result };
    } catch (err) {
      return { success: false, error: err.message };
    }
  }

  list() {
    const entries = [];
    for (const [name, skill] of this.skills) {
      entries.push({
        name,
        description: skill.description,
        parameters: skill.parameters,
        category: skill.category,
      });
    }
    return entries;
  }

  toPromptDescription() {
    return this.list().map(s =>
      `- ${s.name}: ${s.description} (params: ${JSON.stringify(s.parameters)})`
    ).join('\n');
  }

  registerBuiltinSkills() {
    this.register('launch_app', {
      description: '啟動本機應用程式',
      parameters: { name: { type: 'string', description: '應用程式名稱' } },
      category: 'system',
      handler: async (args) => window.electronAPI.launchApp(args.name),
    });

    this.register('open_url', {
      description: '在瀏覽器開啟網址',
      parameters: { url: { type: 'string', description: 'URL' } },
      category: 'web',
      handler: async (args) => window.electronAPI.openUrl(args.url),
    });

    this.register('open_path', {
      description: '開啟本機檔案或資料夾',
      parameters: { path: { type: 'string', description: '檔案路徑' } },
      category: 'file',
      handler: async (args) => window.electronAPI.openPath(args.path),
    });

    this.register('run_command', {
      description: '執行系統指令（Shell command）',
      parameters: { command: { type: 'string', description: '指令' } },
      category: 'system',
      confirmRequired: true,
      handler: async (args) => window.electronAPI.runCommand(args.command),
    });

    this.register('clipboard_read', {
      description: '讀取剪貼簿內容',
      parameters: {},
      category: 'system',
      handler: async () => window.electronAPI.clipboardRead(),
    });

    this.register('clipboard_write', {
      description: '寫入文字到剪貼簿',
      parameters: { text: { type: 'string', description: '要複製的文字' } },
      category: 'system',
      handler: async (args) => window.electronAPI.clipboardWrite(args.text),
    });

    this.register('system_info', {
      description: '取得電腦系統資訊（CPU、記憶體、運行時間）',
      parameters: {},
      category: 'system',
      handler: async () => window.electronAPI.systemInfo(),
    });

    this.register('search_web', {
      description: '搜尋網路（開啟 Google 搜尋結果）',
      parameters: { query: { type: 'string', description: '搜尋關鍵字' } },
      category: 'web',
      handler: async (args) => {
        const url = `https://www.google.com/search?q=${encodeURIComponent(args.query)}`;
        return window.electronAPI.openUrl(url);
      },
    });

    this.register('set_timer', {
      description: '設定提醒計時器（秒數）',
      parameters: {
        seconds: { type: 'number', description: '秒數' },
        message: { type: 'string', description: '提醒訊息' }
      },
      category: 'utility',
      handler: async (args) => {
        return new Promise((resolve) => {
          setTimeout(() => {
            new Notification('Digital Assistant 提醒', { body: args.message || '時間到！' });
            resolve({ message: `計時器響了: ${args.message || '時間到'}` });
          }, (args.seconds || 60) * 1000);
          resolve({ message: `已設定 ${args.seconds} 秒後提醒: ${args.message || '時間到'}` });
        });
      },
    });
  }
}

// Export for use in app.js
if (typeof module !== 'undefined') module.exports = SkillRegistry;
