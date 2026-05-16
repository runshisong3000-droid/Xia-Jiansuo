# 瞎检索 - 文献核查与格式化工具 (Desktop Edition)
# Xia-Jiansuo - Citation Verification &amp; Formatting Tool (Desktop Edition)

&gt; 📚 **CiteFix 家族成员** | CiteFix Family Member  
&gt; 基于 CrossRef API 的 DOI 获取工具 | DOI Retrieval Tool Based on CrossRef API

---

## 项目背景 | Project Background

**"能不能把这些参考文献整理成 APA 格式？"**  
— Every professor, everywhere, forever.

这是 [CiteFix](https://github.com/runshisong3000-droid/Citefix) 的**桌面应用分支**，采用纯 Python + Tkinter 开发，无需安装 Python 环境，直接运行 EXE 文件即可使用。

This is the **desktop application branch** of [CiteFix](https://github.com/runshisong3000-droid/Citefix), developed with pure Python + Tkinter. No Python environment installation required - just run the EXE file directly!

| 版本 Version | 技术栈 Tech Stack | 适用场景 Use Case |
|--------------|-------------------|-------------------|
| **[CiteFix](https://github.com/runshisong3000-droid/Citefix)** | FastAPI + React | Web 在线服务、团队协作 Online Service, Team Collaboration |
| **瞎检索 Xia-Jiansuo (本项目 this)** | Python + Tkinter | 桌面单机使用、无需网络服务 Desktop Use, No Service Required |

---

## 功能特点 | Features

- **DOI 自动获取**：通过 CrossRef API 批量获取文献 DOI | Automatic DOI Retrieval via CrossRef API
- **BibTeX 导出**：生成的 BibTeX 可直接拖入 Zotero 导入 | BibTeX Export for Direct Zotero Import
- **多格式支持**：支持 APA 7、BibTeX、CSV 报告导出 | Multi-format Support (APA 7, BibTeX, CSV)
- **中文文献支持**：本地解析 GB/T 7714 格式文献 | Chinese GB/T 7714 Format Support
- **离线可用**：打包为 EXE 后可离线使用 | Offline Capable After Packaging
- **漫画风格界面**：简洁美观的桌面 GUI 设计 | Comic-style Desktop GUI

---

## 安装与运行 | Installation &amp; Usage

### 方式一：直接使用（推荐）| Method 1: Direct Use (Recommended)

下载打包好的可执行文件：
```
dist/瞎检索.exe
```

双击即可运行，无需安装任何依赖。  
Double-click to run, no dependencies required.

### 方式二：从源码运行 | Method 2: Run from Source

```bash
# 安装依赖 Install dependencies
pip install habanero python-docx

# 运行程序 Run program
python xia_jiansuo_fixed_v5.py
```

### 打包为 EXE | Package as EXE

```bash
pip install pyinstaller
python -m PyInstaller --clean --onefile --windowed --icon xia_jiansuo_avatar.ico --name 瞎检索 xia_jiansuo_fixed_v5.py
```

---

## 使用方法 | How to Use

1. **粘贴文献**：在左侧输入框粘贴文献列表 | Paste references in left text box
2. **导入文件**：支持 `.docx` / `.txt` / `.md` 文件导入 | Import `.docx` / `.txt` / `.md` files
3. **开始核查**：点击「开始核查」按钮获取 DOI | Click "Start Processing" to retrieve DOIs
4. **导出 BibTeX**：保存为 `.bib` 文件，拖入 Zotero 导入 | Export as `.bib` file, drag into Zotero

### 核查状态说明 | Verification Status

| 图标 Icon | 状态 Status | 含义 Meaning |
|-----------|-------------|--------------|
| ✓ | Verified | DOI 已验证（高置信度）DOI Verified (High Confidence) |
| ⚠ | Uncertain | 需人工确认（中等置信度）Need Human Verification (Medium Confidence) |
| ✗ | Not Found | 未找到（需手动处理）Not Found (Manual Handling Required) |

---

## 技术栈 | Tech Stack

- **语言 Language**: Python 3.11+
- **GUI**: Tkinter
- **API**: habanero (CrossRef)
- **文档处理 Document Handling**: python-docx
- **打包 Packaging**: PyInstaller

---

## 项目结构 | Project Structure

```
d:\CiteFix/
├── xia_jiansuo_fixed_v5.py    # 主程序 Main program (最新版本 Latest)
├── xia_jiansuo_fixed_v4.py    # v4 版本
├── xia_jiansuo_fixed_v3.py    # v3 版本
├── xia_jiansuo_avatar.ico     # 应用图标 App icon
├── dist/                      # 打包后的可执行文件 Packaged EXE
│   └── 瞎检索.exe
└── build/                     # PyInstaller 临时文件 Temp files
```

---

## 相关项目 | Related Projects

- **[CiteFix (Web Edition)](https://github.com/runshisong3000-droid/Citefix)** - FastAPI + React Web 应用 Web Application

---

## License

MIT License

---

*Made with ☕ and determination by a student who should probably be studying.*
