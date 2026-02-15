# Digital Assistant - 桌面 AI 數位助理

一個功能完整的桌面 AI 助理，支援語音控制、技能系統，並使用 Google Colab 作為 LLM 後端。

## 功能特色

### 核心功能
- **AI 對話**: 使用 Qwen2.5-14B-Instruct 模型，支援繁體中文與英文
- **語音輸入**: 支援語音轉文字（使用 Google Speech Recognition）
- **語音輸出**: 文字轉語音播放
- **技能系統**: 19+ 個內建技能，可擴展

### 主要技能

#### 系統控制
- 啟動應用程式（記事本、計算機、瀏覽器、VSCode 等）
- 執行系統指令
- 取得系統資訊（CPU、記憶體、運行時間）
- 終止程序
- 設定系統音量
- 發送 Windows 桌面通知

#### 檔案操作
- 開啟檔案或資料夾
- 讀取/寫入文字檔案
- 列出目錄檔案
- 截取螢幕截圖

#### 文件建立
- 建立 PowerPoint 簡報（支援多種主題）
- 建立 Word 文件
- 建立 Excel 試算表

#### 剪貼簿
- 讀取剪貼簿內容
- 寫入文字到剪貼簿

#### 網路功能
- 搜尋網頁（開啟瀏覽器）
- 搜尋新聞並返回內容摘要（用於製作簡報/文件）

#### 工具功能
- 取得目前日期時間
- 設定提醒計時器

## 技術架構

### 前端
- **Electron 版本**: Node.js + Electron（功能完整）
- **PyWebView 版本**: Python + PyWebView（輕量）

### 後端
- **AI 伺服器**: Google Colab + FastAPI
- **模型**: Qwen2.5-14B-Instruct（14B 參數）
- **網路隧道**: Cloudflare Tunnel（免費，無需帳號）

### 技術棧
- Frontend: HTML/CSS/JavaScript
- Backend: Python (PyWebView) / Node.js (Electron)
- AI: Transformers, Qwen2.5-14B-Instruct
- Web Framework: FastAPI
- Speech: Web Speech API, Google Speech Recognition

## 安裝步驟

### 前置需求
- Windows 10/11
- Python 3.8+（PyWebView 版本）
- Node.js 16+（Electron 版本）
- Google Colab Pro with GPU（A100 或 T4）

### 方式一：Electron 版本（推薦）

```bash
# 1. 安裝依賴
npm install

# 2. 啟動應用
npm start

# 開發模式（會開啟開發者工具）
npm run dev
```

### 方式二：PyWebView 版本

```bash
# 1. 建立虛擬環境
python -m venv .venv

# 2. 啟動虛擬環境
.venv\Scripts\activate

# 3. 安裝依賴
pip install pywebview requests python-docx python-pptx openpyxl pillow speechrecognition duckduckgo-search beautifulsoup4

# 4. 啟動應用
python main.py
```

### 啟動 AI 伺服器（Colab）

1. 開啟 `colab/DigitalAssistant_Server.ipynb`
2. 在 Google Colab 中執行所有 cell（Runtime → Run all）
3. 複製輸出的 `trycloudflare.com` URL
4. 在桌面應用中設定此 URL（點擊設定按鈕）

## 使用說明

### 基本使用

1. **啟動應用**: 執行 `npm start` 或 `python main.py`
2. **設定 AI 伺服器**: 在設定中貼上 Colab 的 URL
3. **開始對話**: 使用文字輸入或語音輸入

### 語音輸入

- 點擊麥克風按鈕（或按住）
- 說話後放開按鈕
- 系統會自動轉錄並發送

### 技能使用

#### 文字指令範例
- "幫我開啟記事本"
- "現在幾點？"
- "我的電腦資訊"
- "幫我建立一份關於 AI 的簡報"
- "搜尋最新的 AI 新聞然後做成簡報"

#### 技能按鈕
- 點擊介面下方的技能按鈕快速執行

### 進階功能

#### 搜尋並製作簡報
```
使用者: "搜尋 AI 最新新聞然後做成簡報"
AI 會:
1. 執行 fetch_news 搜尋資料
2. 分析搜尋結果
3. 自動執行 create_ppt 製作簡報
```

#### 文件建立
- PowerPoint: 支援 6 種主題（dark, corporate, nature, warm, ocean, minimal）
- Word: 支援 Markdown 格式標題和列表
- Excel: 支援 CSV 格式資料

## 專案結構

```
DigitalAssistant/
├── app/                    # Electron 版本
│   ├── main.js            # 主程序
│   └── preload.js         # 預載腳本
├── colab/                 # Colab AI 伺服器
│   └── DigitalAssistant_Server.ipynb
├── public/                # 前端介面
│   ├── index.html
│   ├── app.js
│   └── styles.css
├── src/                   # 原始碼
│   ├── mcp/              # MCP 客戶端
│   └── skills/            # 技能註冊
├── main.py               # PyWebView 版本主程序
├── package.json          # Node.js 依賴
└── README.md            # 本文件
```

## 技能列表

| 技能名稱 | 描述 | 參數 |
|---------|------|------|
| `launch_app` | 啟動應用程式 | `name: string` |
| `open_url` | 開啟網址 | `url: string` |
| `open_path` | 開啟檔案或資料夾 | `path: string` |
| `run_command` | 執行系統指令 | `command: string` |
| `clipboard_read` | 讀取剪貼簿 | - |
| `clipboard_write` | 寫入剪貼簿 | `text: string` |
| `system_info` | 取得系統資訊 | - |
| `create_ppt` | 建立 PowerPoint 簡報 | `title: string, slides_json: array, theme: string` |
| `create_docx` | 建立 Word 文件 | `title: string, content: string` |
| `create_xlsx` | 建立 Excel 試算表 | `title: string, data_json: array` |
| `write_file` | 寫入文字檔案 | `filename: string, content: string` |
| `read_file` | 讀取檔案內容 | `filepath: string` |
| `list_files` | 列出目錄檔案 | `directory: string` |
| `take_screenshot` | 截取螢幕截圖 | - |
| `get_datetime` | 取得日期時間 | - |
| `search_web` | 用瀏覽器搜尋網頁 | `query: string` |
| `fetch_news` | 搜尋網路資訊並返回摘要 | `query: string, max_results: number` |
| `kill_process` | 終止程序 | `process_name: string` |
| `set_volume` | 設定系統音量 | `level: number` |
| `notify` | 發送桌面通知 | `title: string, message: string` |

## 開發

### 環境設定

```bash
# Electron 版本
npm install

# PyWebView 版本
pip install -r requirements.txt
```

### 除錯模式

```bash
# Electron
npm run dev

# PyWebView
python main.py --dev
```

## 注意事項

1. **Colab 連線**: Colab notebook 必須保持運行，如果斷線需要重新設定 URL
2. **麥克風權限**: 首次使用語音功能需要允許麥克風權限
3. **網路需求**: 語音識別需要網路連線（使用 Google 服務）
4. **GPU 需求**: Colab 需要 GPU（建議 A100 或 T4），需要 Colab Pro

## 授權

本專案採用 MIT 授權。

## 貢獻

歡迎提交 Issue 和 Pull Request。

## 作者

ChangChiaEn

## 相關連結

- GitHub: https://github.com/ChangChiaEn/DigitalAssistant
- 模型: [Qwen2.5-14B-Instruct](https://huggingface.co/Qwen/Qwen2.5-14B-Instruct)

