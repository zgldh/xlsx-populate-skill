

# Processing Excel Files Skill

[![GitHub](https://img.shields.io/badge/GitHub-zgldh%2Fxlsx--populate--skill-blue)](https://github.com/zgldh/xlsx-populate-skill)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-2.0.0-green.svg)]()

A professional Skill for OpenCode, Claude Code, and other AI coding assistants to edit and create Excel (.xlsx) files while perfectly preserving original formatting, merged cells, and styles.

一个用于 OpenCode、Claude Code 和其他 AI 编码助手的专业 Skill，用于在完美保留原有格式、合并单元格和样式的前提下编辑和创建 Excel (.xlsx) 文件。

---

## ✨ Features | 特性

- ✅ **Perfect Format Preservation** - Keep all styles, merged cells, and layouts intact | **完美保留格式** - 保留所有样式、合并单元格和布局
- ✅ **Formula Support** - Add Excel formulas for automatic calculations | **公式支持** - 添加 Excel 公式进行自动计算
- ✅ **Flexible Worksheet Management** - Add, delete, rename, and reorder worksheets | **灵活工作表管理** - 添加、删除、重命名和重新排序工作表
- ✅ **Professional Styling** - Apply fonts, colors, alignment, borders, and number formats | **专业样式** - 应用字体、颜色、对齐、边框和数字格式
- ✅ **Progressive Disclosure** - Skill uses reference files for advanced topics (keeps SKILL.md concise) | **渐进式披露** - Skill 使用参考文件处理高级主题（保持 SKILL.md 简洁）

---

## 📦 Installation | 安装

### Method 1: Via npx (Recommended) | 方式 1：通过 npx（推荐）

```bash
npx skills add zgldh/xlsx-populate-skill
```

### Method 2: Clone to Project | 方式 2：克隆到项目

```bash
git clone https://github.com/zgldh/xlsx-populate-skill.git .opencode/skills/processing-excel-files
```

### Method 3: Global Installation | 方式 3：全局安装

```bash
git clone https://github.com/zgldh/xlsx-populate-skill.git ~/.config/opencode/skills/processing-excel-files
```

### Dependency | 依赖

```bash
npm install xlsx-populate
```

---

## 🚀 Quick Start | 快速开始

```javascript
const XlsxPopulate = require('xlsx-populate');

// Load and edit while preserving formatting
const workbook = await XlsxPopulate.fromFileAsync('input.xlsx');
workbook.sheet(0).cell('A1').value('Updated Value');
await workbook.toFileAsync('output.xlsx');
```

---

## 📚 Skill Structure | Skill 结构

This skill follows **skill-creator** best practices with progressive disclosure:

```
xlsx-populate-skill/
├── SKILL.md                    # Core instructions (concise)
├── BATCH-OPERATIONS.md         # Large dataset handling
├── FORMULAS.md                 # Financial modeling patterns
├── STYLES.md                   # Color schemes and formatting
├── EXAMPLES.md                 # Real-world scenarios
├── examples/                   # Executable code examples
│   ├── basic-usage.js
│   ├── quotation-editor.js
│   └── excel-processor.js
├── README.md                   # This file
├── package.json
└── LICENSE
```

---

## 🤖 Compatible AI Assistants | 兼容的 AI 助手

| Assistant | Installation | 助手 | 安装方式 |
|-----------|-------------|------|---------|
| **OpenCode** | `npx skills add zgldh/xlsx-populate-skill` | **OpenCode** | `npx skills add zgldh/xlsx-populate-skill` |
| **Claude Code** | `npx skills add zgldh/xlsx-populate-skill` | **Claude Code** | `npx skills add zgldh/xlsx-populate-skill` |
| **Cursor** | Clone to `.cursor/skills/` | **Cursor** | 克隆到 `.cursor/skills/` |
| **Goose** | Clone to `.goose/skills/` | **Goose** | 克隆到 `.goose/skills/` |
| **Roo Code** | Clone to `.roo/skills/` | **Roo Code** | 克隆到 `.roo/skills/` |

---

## 📝 Usage Examples | 使用示例

### Edit Existing File | 编辑现有文件

```javascript
const XlsxPopulate = require('xlsx-populate');

const workbook = await XlsxPopulate.fromFileAsync('report.xlsx');
const sheet = workbook.sheet(0);

// Modify cell
sheet.cell('A1').value('Updated Title');
sheet.cell('A1').style({
  bold: true,
  fontSize: 16,
  fill: '4472C4',
  fontColor: 'FFFFFF'
});

// Add formula
sheet.cell('D10').formula('=SUM(D2:D9)');

await workbook.toFileAsync('report-updated.xlsx');
```

### Create New Worksheet | 创建新工作表

```javascript
const newSheet = workbook.addSheet('Summary');
newSheet.cell('A1').value('Summary Report');
newSheet.range('A1:D1').merged(true).style({
  horizontalAlignment: 'center'
});
```

To run the bundled examples locally, install dependencies and use the provided scripts:

```bash
npm install
npm run example:basic
npm run example:quotation
npm run example:processor
```

---

## 📖 Reference Materials | 参考材料

The skill includes detailed reference files:

- **[BATCH-OPERATIONS.md](BATCH-OPERATIONS.md)** - Handling large datasets efficiently | 高效处理大数据集
- **[FORMULAS.md](FORMULAS.md)** - Financial modeling standards and formula patterns | 财务建模标准和公式模式
- **[STYLES.md](STYLES.md)** - Professional color schemes and formatting | 专业配色方案和格式
- **[EXAMPLES.md](EXAMPLES.md)** - Complete real-world examples | 完整真实场景示例

---

## 🎯 When to Use | 何时使用

**Use this skill when:**
- User wants to edit existing Excel files without destroying formatting
- Working with .xlsx files that have complex layouts or merged cells
- Need to add formulas, styling, or new worksheets to existing files
- Creating Excel reports from templates

**Do NOT use when:**
- Only need to read data from Excel (use `xlsx` library for better performance)
- Creating simple Excel files from scratch without formatting concerns

---

## 📄 License | 许可证

MIT License - See [LICENSE](LICENSE) file for details.

MIT 许可证 - 详见 [LICENSE](LICENSE) 文件。

---

## 🤝 Contributing | 贡献

**English**: Issues and Pull Requests are welcome! Follow [skill-creator best practices](https://github.com/anthropics/skills/tree/main/skill-creator).

**中文**: 欢迎提交 Issue 和 Pull Request！请遵循 [skill-creator 最佳实践](https://github.com/anthropics/skills/tree/main/skill-creator).

---

## 🙏 Acknowledgments | 致谢

- [xlsx-populate](https://github.com/dtjohnson/xlsx-populate) - The excellent Excel processing library
- [skill-creator](https://github.com/anthropics/skills/tree/main/skill-creator) - Best practices for creating effective skills

---

<div align="center">

**⭐ Star this repo if you find it helpful! | 如果觉得有用，请给个星星！⭐**

[Report Bug](https://github.com/zgldh/xlsx-populate-skill/issues) · [Request Feature](https://github.com/zgldh/xlsx-populate-skill/issues)

</div>
