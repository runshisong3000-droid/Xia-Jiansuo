# 瞎检索 - 文献核查与格式化工具

一款基于 CrossRef API 的文献检索与 DOI 获取工具，支持批量处理文献列表并导出为 APA 7、BibTeX 格式，可直接导入 Zotero 等文献管理软件。

## 功能特点

- **智能文献解析**：支持多种参考文献格式（APA、MLA、Vancouver、GB/T 7714 等）
- **DOI 自动获取**：通过 CrossRef API 批量获取文献 DOI
- **多格式导出**：支持 APA 7、BibTeX、CSV 报告
- **Zotero 兼容**：导出的 BibTeX 可直接拖入 Zotero
- **中文文献支持**：本地解析 GB/T 7714 格式文献
- **漫画风格界面**：简洁美观的 GUI 设计

## 安装

### 依赖

```bash
pip install habanero python-docx
```

### 运行

```bash
# 直接运行源码
python xia_jiansuo_fixed_v5.py

# 或使用打包好的可执行文件
./dist/瞎检索.exe
```

### 打包为 EXE

```bash
pip install pyinstaller
python -m PyInstaller --clean --onefile --windowed --icon xia_jiansuo_avatar.ico --name 瞎检索 xia_jiansuo_fixed_v5.py
```

## 使用方法

1. **粘贴文献**：在左侧输入框粘贴文献列表
2. **导入文件**：支持 `.docx` / `.txt` / `.md` 文件
3. **开始核查**：点击「开始核查」按钮
4. **导出结果**：可保存为 APA/BibTeX 或复制到剪贴板

## 项目结构

```
d:\CiteFix/
├── xia_jiansuo_fixed_v5.py    # 主程序
├── xia_jiansuo_avatar.ico     # 应用图标
├── dist/                      # 打包后的可执行文件
└── build/                    # 打包临时文件
```

## 技术栈

- Python 3.11+
- Tkinter (GUI)
- habanero (CrossRef API)
- python-docx (Word 文件读取)
- PyInstaller (打包)

## License

MIT License
