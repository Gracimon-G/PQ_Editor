# PQ Editor — Unchained M Engine

> **Slogan**: *Unchained M Engine · Zero Locks, Zero Freezes*  
> **Supported Workbooks**: `.xlsx`, `.xlsm`, `.xlsb`  
> **Zero External Dependencies**: Built 100% on the Python standard library (`tkinter`, `zipfile`, `xml`, `struct`, `re`).

---

## Overview / 项目概述

**PQ Editor** is a standalone, lightweight desktop application designed to inspect, edit, validate, and visualize Power Query (M) formulas directly inside Excel workbooks—**without opening Microsoft Excel**.

When editing M code in Excel's native Power Query editor, Excel locks the interface, blocks other workbooks, and frequently hangs or freezes during heavy query refreshes. PQ Editor bypasses Excel entirely by reverse-engineering the underlying `customXml` OpenXML parts and binary mashup envelopes (`DataMashup`), delivering an instant, freeze-free development experience.

---

## ✨ Key Features / 核心功能

- **⚡ Zero Locks, Zero Freezes / 彻底杜绝 Excel 假死**: Edits M code directly at the package and binary container level. Never spawns Excel processes or COM objects.
- **🌲 Query Groups & Destination Resolution / 查询组与加载方式识别**: Automatically unpacks serialized `.NET BinaryWriter` structures to reconstruct user query folders (e.g. `📁 Helper Query`) and detects whether queries load into a worksheet `Table`, `Connection Only`, or `Data Model`.
- **🎨 Modern Code Editor / 现代化代码编辑体验**:
  - Full syntax highlighting for M keywords, identifiers, strings, numbers, and comments.
  - Synchronized line-number margin canvas with wrap/toggle support.
  - Live cursor position tracking `(Ln xx, Col xx)`.
  - Find & Replace bar with live match counting and smooth `Enter` / `Shift+Enter` navigation.
- **🛡️ Pre-Save Syntax Validation / 保存前语法防护**: Tokenizes and validates bracket balancing (`()`, `[]`, `{}`), unclosed string literals, quoted identifiers (`#"..."`), and matched `let ... in` blocks before writing back to prevent workbook corruption.
- **📊 Lineage & Join Graph Visualizer / 依赖血缘与连接关系图**:
  - **In-App Canvas**: Interactive pan-and-drag and scroll-zoom graph. Double-click any query node to jump straight to its code definition.
  - **🌐 Web Graph**: Full-screen browser diagram rendered with Mermaid.js, detailing `JoinKind` (e.g. `LeftOuter`, `Inner`) and key matching columns.
  - **Mermaid Export**: 1-click clipboard copy or `.mmd` export.
- **💾 Safe Atomic Saves & `.bak` Protection / 安全原子保存与备份**: Writes back via atomic replacement (`tempfile` + `os.replace`). Always generates a `.bak` backup before modifying the workbook, with 1-click restore.
- **📤 Multi-Extension Code Export / 多格式脚本导出**: Export queries combined or split per query into `.pq`, `.m`, or `.txt` for version control.
- **🌙 Dark & Light Themes / 深浅双色主题**: Toggle seamlessly between VS Code Dark (`#1E1E1E`) and Office Light palettes.

---

## ⌨️ Keyboard Shortcuts / 快捷键

| Shortcut | Description (English) | 说明 (中文) |
| :--- | :--- | :--- |
| <kbd>Ctrl</kbd> + <kbd>S</kbd> | Save workbook in place (with automatic `.bak`) | 保存工作簿（自动建立 `.bak` 备份） |
| <kbd>Ctrl</kbd> + <kbd>F</kbd> | Open Find & Replace bar | 打开查找替换栏（聚焦查找框） |
| <kbd>Ctrl</kbd> + <kbd>H</kbd> | Open Find & Replace bar (focuses Replace field) | 打开查找替换栏（聚焦替换框） |
| <kbd>Ctrl</kbd> + <kbd>/</kbd> | Toggle comment (`//`) on active line or selection | 快速注释 / 取消注释当前行或选中行 |
| <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>/</kbd> | Uncomment selected lines (strips leading `//`) | 批量取消注释所选行 |
| <kbd>Enter</kbd> / <kbd>Shift+Enter</kbd> | Jump to next / previous search match in Find mode | 查找模式下跳至下一个 / 上一个匹配项 |
| <kbd>Esc</kbd> | Close Find & Replace bar | 关闭查找替换栏 |
| <kbd>Ctrl</kbd> + <kbd>Z</kbd> / <kbd>Ctrl</kbd> + <kbd>Y</kbd> | Undo / Redo edits | 撤销 / 重做代码编辑 |

---

## 🚀 Getting Started / 快速启动

### Running from Source
Requires Python 3.10+ (standard library only):
```bash
python PQ_editor_gui_mermaid.py
```

### Packaging Standalone Executable (`dist/PQ_Editor.exe`)
To build the single-file, non-console Windows `.exe`:
```bash
python -m PyInstaller --onefile --noconsole --name "PQ_Editor" --clean PQ_editor_gui_mermaid.py
```
Output: `dist/PQ_Editor.exe` (~14 MB, zero runtime dependencies, no black terminal window).

---

## 📚 Documentation / 详细文档

- [📖 User Guide (English & 中文双语用户指南)](user_guide.html)

---

## License
MIT License.