# 瞎检索 - 文献核查与格式化工具 (Desktop Edition)

> 📚 **CiteFix 家族成员** | 基于 CrossRef API 的 DOI 获取工具

## 项目背景

**"能不能把这些参考文献整理成 APA 格式？"**
— Every professor, everywhere, forever.

这是 [CiteFix](https://github.com/runshisong3000-droid/Citefix) 的**桌面应用分支**，采用纯 Python + Tkinter 开发，无需安装 Python 环境，直接运行 EXE 文件即可使用。

| 版本 | 技术栈 | 适用场景 |
|------|--------|----------|
| **[CiteFix](https://github.com/runshisong3000-droid/Citefix)** | FastAPI + React | Web 在线服务、团队协作 |
| **瞎检索 (本项目)** | Python + Tkinter | 桌面单机使用、无需网络服务 |

## 功能特点

- **DOI 自动获取**：通过 CrossRef API 批量获取文献 DOI
- **BibTeX 导出**：生成的 BibTeX 可直接拖入 Zotero 导入
- **多格式支持**：支持 APA 7、BibTeX、CSV 报告导出
- **中文文献支持**：本地解析 GB/T 7714 格式文献
- **离线可用**：打包为 EXE 后可离线使用
- **漫画风格界面**：简洁美观的桌面 GUI 设计

## 安装与运行

### 方式一：直接使用（推荐）

下载打包好的可执行文件：
```
dist/瞎检索.exe
```

双击即可运行，无需安装任何依赖。

### 方式二：从源码运行

```bash
# 安装依赖
pip install habanero python-docx

# 运行程序
python xia_jiansuo_fixed_v5.py
```

### 打包为 EXE

```bash
pip install pyinstaller
python -m PyInstaller --clean --onefile --windowed --icon xia_jiansuo_avatar.ico --name 瞎检索 xia_jiansuo_fixed_v5.py
```

## 使用方法

1. **粘贴文献**：在左侧输入框粘贴文献列表
2. **导入文件**：支持 `.docx` / `.txt` / `.md` 文件导入
3. **开始核查**：点击「开始核查」按钮获取 DOI
4. **导出 BibTeX**：保存为 `.bib` 文件，拖入 Zotero 导入

### 核查状态说明

| 图标 | 状态 | 含义 |
|------|------|------|
| ✓ | Verified | DOI 已验证（高置信度） |
| ⚠ | Uncertain | 需人工确认（中等置信度） |
| ✗ | Not Found | 未找到（需手动处理） |

## 技术栈

- **语言**: Python 3.11+
- **GUI**: Tkinter
- **API**: habanero (CrossRef)
- **文档处理**: python-docx
- **打包**: PyInstaller

## 项目结构

```
d:\CiteFix/
├── xia_jiansuo_fixed_v5.py    # 主程序 (最新版本)
├── xia_jiansuo_fixed_v4.py    # v4 版本
├── xia_jiansuo_fixed_v3.py    # v3 版本
├── xia_jiansuo_avatar.ico     # 应用图标
├── dist/                      # 打包后的可执行文件
│   └── 瞎检索.exe
└── build/                     # PyInstaller 临时文件
```

## 相关项目

- **[CiteFix (Web Edition)](https://github.com/runshisong3000-droid/Citefix)** - FastAPI + React Web 应用

## License

MIT License

---

*Made with ☕ and determination by a student who should probably be studying.*
