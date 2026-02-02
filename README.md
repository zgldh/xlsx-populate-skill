# xlsx-populate Skill

A Skill for OpenCode/Claude Code to edit Excel files while preserving original formatting.

一个用于 OpenCode/Claude Code 的 Skill，用于在保留原有格式的前提下编辑 Excel 文件。

[![GitHub](https://img.shields.io/badge/GitHub-zgldh%2Fxlsx--populate--skill-blue)](https://github.com/zgldh/xlsx-populate-skill)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![AgentSkillsRepo](https://img.shields.io/badge/AgentSkillsRepo-Submit-green)](https://agentskillsrepo.com/)

---

## ✨ Features | 特点

| English | 中文 |
|---------|------|
| ✅ **Preserve original formatting** - Keep styles and merged cells intact | ✅ **保留原有格式** - 不破坏原始文件的样式、合并单元格 |
| ✅ **Formula support** - Add Excel formulas for automatic calculation | ✅ **支持公式** - 添加 Excel 公式自动计算 |
| ✅ **Flexible editing** - Modify, add, or delete worksheets | ✅ **灵活编辑** - 修改、添加、删除工作表 |
| ✅ **Style control** - Fonts, colors, alignment, borders | ✅ **样式控制** - 字体、颜色、对齐、边框 |
| ✅ **Merge cells** - Create and preserve merged cells | ✅ **合并单元格** - 支持创建和保留合并单元格 |

---

## 📦 Installation | 安装

### Method 1: via npx (Recommended) | 方式 1：通过 npx（推荐）

```bash
npx skills add zgldh/xlsx-populate-skill
```

### Method 2: Clone to project directory | 方式 2：克隆到项目目录

```bash
# Clone to .opencode/skills/ directory | 克隆到 .opencode/skills/ 目录
git clone https://github.com/zgldh/xlsx-populate-skill.git .opencode/skills/xlsx-populate
```

### Method 3: Global installation | 方式 3：全局安装

```bash
# Clone to user config directory | 克隆到用户配置目录
git clone https://github.com/zgldh/xlsx-populate-skill.git ~/.config/opencode/skills/xlsx-populate
```

### Dependencies | 依赖安装

```bash
npm install xlsx-populate
```

---

## 🤖 Compatible AI Coding Assistants | 支持的 AI 编码助手

| Tool | Install Command | 工具 | 安装命令 |
|------|----------------|------|---------|
| **OpenCode** | `npx skills add zgldh/xlsx-populate-skill` | **OpenCode** | `npx skills add zgldh/xlsx-populate-skill` |
| **Claude Code** | `npx skills add zgldh/xlsx-populate-skill` | **Claude Code** | `npx skills add zgldh/xlsx-populate-skill` |
| **Cursor** | Install to `.cursor/skills/` | **Cursor** | 安装到 `.cursor/skills/` 目录 |
| **Goose** | Install to `.goose/skills/` | **Goose** | 安装到 `.goose/skills/` 目录 |
| **Roo Code** | Install to `.roo/skills/` | **Roo Code** | 安装到 `.roo/skills/` 目录 |
| **Windsurf** | Install to `.codeium/windsurf/skills/` | **Windsurf** | 安装到 `.codeium/windsurf/skills/` 目录 |

---

## 🚀 Quick Start | 快速开始

### English

```javascript
const XlsxPopulate = require('xlsx-populate');

async function editExcel() {
  // Load from file (preserve all formatting)
  const workbook = await XlsxPopulate.fromFileAsync('input.xlsx');
  
  // Get worksheet
  const sheet = workbook.sheet(0);
  
  // Modify cell
  sheet.cell('A1').value('New Title');
  sheet.cell('A1').style({
    bold: true,
    fontColor: 'FF0000',
    fontSize: 14
  });
  
  // Add new worksheet
  const newSheet = workbook.addSheet('New Sheet');
  newSheet.cell('A1').value('Content');
  
  // Save (preserve all original formatting)
  await workbook.toFileAsync('output.xlsx');
}
```

### 中文示例

```javascript
const XlsxPopulate = require('xlsx-populate');

async function editExcel() {
  // 从文件加载（保留所有格式）
  const workbook = await XlsxPopulate.fromFileAsync('input.xlsx');
  
  // 获取工作表
  const sheet = workbook.sheet(0);
  
  // 修改单元格
  sheet.cell('A1').value('新标题');
  sheet.cell('A1').style({
    bold: true,
    fontColor: 'FF0000',
    fontSize: 14
  });
  
  // 添加新工作表
  const newSheet = workbook.addSheet('新工作表');
  newSheet.cell('A1').value('内容');
  
  // 保存（保留所有原有格式）
  await workbook.toFileAsync('output.xlsx');
}
```

---

## 📚 Examples | 示例代码

Check the `examples/` directory for:
| File | Description | 文件 | 说明 |
|------|-------------|------|------|
| `basic-usage.js` | Basic usage examples | 基础用法示例 |
| `quotation-editor.js` | Quotation editor (real-world scenario) | 报价单编辑器（实际应用场景） |
| `excel-processor.js` | Encapsulated class for reuse | 封装类，便于复用 |

---

## 📋 Feature List | 功能列表

### Read & Write | 读取与写入
- ✅ Load from file (preserve formatting) | 从文件加载（保留格式）
- ✅ Create from blank | 从空白创建
- ✅ Save to file | 保存到文件

### Worksheet Operations | 工作表操作
- ✅ Add worksheet | 添加工作表
- ✅ Delete worksheet | 删除工作表
- ✅ Rename worksheet | 重命名工作表
- ✅ Move worksheet order | 移动工作表顺序
- ✅ Iterate all worksheets | 遍历所有工作表

### Cell Operations | 单元格操作
- ✅ Set value | 设置值
- ✅ Set formula | 设置公式
- ✅ Set style | 设置样式
- ✅ Batch write data | 批量写入数据

### Styling | 样式设置
- ✅ Font (size, color, bold, italic) | 字体（大小、颜色、粗体、斜体）
- ✅ Fill (background color) | 填充（背景色）
- ✅ Alignment (horizontal, vertical) | 对齐（水平、垂直）
- ✅ Border | 边框
- ✅ Number format | 数字格式

### Advanced Features | 高级功能
- ✅ Merge cells | 合并单元格
- ✅ Set column width / row height | 设置列宽/行高
- ✅ Conditional formatting (via code) | 条件格式（通过代码控制）

---

## ⚖️ Comparison with xlsx Library | 与 xlsx 库的对比

| Feature | xlsx-populate | xlsx |
|---------|---------------|------|
| Preserve original formatting | ✅ Perfect preservation | ❌ Destroys formatting |
| Merged cells | ✅ Supported | ⚠️ Limited support |
| Style editing | ✅ Full support | ⚠️ Limited support |
| File size | Larger | Smaller |
| Performance | Slower | Faster |

**Recommendation** | **建议**
- Use `xlsx-populate` if you need to preserve original formatting | 如果需要保留原有格式，使用 `xlsx-populate`
- Use `xlsx` if you only need to quickly read data | 如果只需要快速读取数据，使用 `xlsx`

---

## 📁 Project Setup | 项目设置

```bash
# Clone repository | 克隆仓库
git clone https://github.com/zgldh/xlsx-populate-skill.git
cd xlsx-populate-skill

# Install dependencies | 安装依赖
npm install

# Run examples | 运行示例
node examples/basic-usage.js
```

---

## 🔗 Dependencies | 依赖

- [xlsx-populate](https://github.com/dtjohnson/xlsx-populate) - Core library | 核心库

---

## 📄 License | 许可证

MIT License - see [LICENSE](LICENSE) file for details.

MIT 许可证 - 详情见 [LICENSE](LICENSE) 文件。

---

## 🤝 Contributing | 贡献

**English**: Issues and Pull Requests are welcome!

**中文**: 欢迎提交 Issue 和 Pull Request！

---

## 🙏 Acknowledgments | 致谢

**English**: Thanks to [xlsx-populate](https://github.com/dtjohnson/xlsx-populate) for the excellent Excel processing library.

**中文**: 感谢 [xlsx-populate](https://github.com/dtjohnson/xlsx-populate) 提供优秀的 Excel 处理库。

---

<div align="center">

**⭐ Star this repo if you find it helpful! | 如果觉得有用，请给个星星！⭐**

[Report Bug](https://github.com/zgldh/xlsx-populate-skill/issues) · [Request Feature](https://github.com/zgldh/xlsx-populate-skill/issues)

</div>
