# xlsx-populate Skill

一个用于 OpenCode/Claude Code 的 Skill，用于在保留原有格式的前提下编辑 Excel 文件。

[![GitHub](https://img.shields.io/badge/GitHub-zgldh%2Fxlsx--populate--skill-blue)](https://github.com/zgldh/xlsx-populate-skill)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![AgentSkillsRepo](https://img.shields.io/badge/AgentSkillsRepo-Submit-green)](https://agentskillsrepo.com/)

## 特点

- ✅ **保留原有格式** - 不破坏原始文件的样式、合并单元格
- ✅ **支持公式** - 添加 Excel 公式自动计算
- ✅ **灵活编辑** - 修改、添加、删除工作表
- ✅ **样式控制** - 字体、颜色、对齐、边框
- ✅ **合并单元格** - 支持创建和保留合并单元格

## 安装

### 方式 1：通过 npx 安装（推荐）

```bash
npx skills add zgldh/xlsx-populate-skill
```

### 方式 2：克隆到项目目录

```bash
# 克隆到项目的 .opencode/skills/ 目录
git clone https://github.com/zgldh/xlsx-populate-skill.git .opencode/skills/xlsx-populate
```

### 方式 3：全局安装

```bash
# 克隆到用户配置目录
git clone https://github.com/zgldh/xlsx-populate-skill.git ~/.config/opencode/skills/xlsx-populate
```

### 依赖安装

```bash
npm install xlsx-populate
```

## 支持的 AI 编码助手

本 skill 兼容以下 AI 编码助手：

| 工具 | 安装命令 |
|------|---------|
| **OpenCode** | `npx skills add zgldh/xlsx-populate-skill` |
| **Claude Code** | `npx skills add zgldh/xlsx-populate-skill` |
| **Cursor** | `安装到 .cursor/skills/` 目录 |
| **Goose** | `安装到 .goose/skills/` 目录 |
| **Goose** | `安装到 .roo/skills/` 目录 |
| **Windsurf** | `安装到 .codeium/windsurf/skills/` 目录 |

## 使用方法

### 直接使用

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

## 示例代码

查看 `examples/` 目录：

- `basic-usage.js` - 基础用法示例
- `quotation-editor.js` - 报价单编辑器（实际应用场景）
- `excel-processor.js` - 封装类，便于复用

## 快速开始

```bash
# 克隆仓库
git clone https://github.com/zgldh/xlsx-populate-skill.git
cd xlsx-populate-skill

# 安装依赖
npm install

# 运行示例
node examples/basic-usage.js
```

## 功能列表

### 读取与写入
- ✅ 从文件加载（保留格式）
- ✅ 从空白创建
- ✅ 保存到文件

### 工作表操作
- ✅ 添加工作表
- ✅ 删除工作表
- ✅ 重命名工作表
- ✅ 移动工作表顺序
- ✅ 遍历所有工作表

### 单元格操作
- ✅ 设置值
- ✅ 设置公式
- ✅ 设置样式
- ✅ 批量写入数据

### 样式设置
- ✅ 字体（大小、颜色、粗体、斜体）
- ✅ 填充（背景色）
- ✅ 对齐（水平、垂直）
- ✅ 边框
- ✅ 数字格式

### 高级功能
- ✅ 合并单元格
- ✅ 设置列宽/行高
- ✅ 条件格式（通过代码控制）

## 与 xlsx 库的对比

| 特性 | xlsx-populate | xlsx |
|------|---------------|------|
| 保留原有格式 | ✅ 完美保留 | ❌ 会破坏格式 |
| 合并单元格 | ✅ 支持 | ⚠️ 有限支持 |
| 样式编辑 | ✅ 完整支持 | ⚠️ 有限支持 |
| 文件大小 | 较大 | 较小 |
| 性能 | 较慢 | 较快 |

**建议**：如果需要保留原有格式，使用 `xlsx-populate`；如果只需要快速读取数据，使用 `xlsx`。

## 依赖

- [xlsx-populate](https://github.com/dtjohnson/xlsx-populate) - 核心库

## 许可证

MIT

## 贡献

欢迎提交 Issue 和 Pull Request！

## 致谢

感谢 [xlsx-populate](https://github.com/dtjohnson/xlsx-populate) 提供优秀的 Excel 处理库。
