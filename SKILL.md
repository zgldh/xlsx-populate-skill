---
name: xlsx-populate
description: Edit and manipulate Excel files while preserving original formatting, merged cells, and styles. Use xlsx-populate library for Node.js to read, modify, and create .xlsx files without destroying existing layouts.
source: local
category: data
license: MIT
tags: [excel, xlsx, spreadsheet, office, data-processing]
---

# xlsx-populate Skill

使用 `xlsx-populate` 库在保留原有格式的前提下编辑 Excel 文件。

## 特点

- ✅ **保留原有格式** - 不破坏原始文件的样式、合并单元格
- ✅ **支持公式** - 添加 Excel 公式自动计算
- ✅ **灵活编辑** - 修改、添加、删除工作表
- ✅ **样式控制** - 字体、颜色、对齐、边框
- ✅ **合并单元格** - 支持创建和保留合并单元格

## 安装依赖

```bash
npm install xlsx-populate
```

## 快速开始

### 1. 读取并保留格式编辑

```javascript
const XlsxPopulate = require('xlsx-populate');

async function editExcel() {
  // 从文件加载（保留所有格式）
  const workbook = await XlsxPopulate.fromFileAsync('input.xlsx');
  
  // 获取第一个工作表
  const sheet = workbook.sheet(0);
  
  // 修改单元格（保留其他格式）
  sheet.cell('A1').value('新标题');
  sheet.cell('A1').style({
    bold: true,
    fontColor: 'FF0000',
    fontSize: 14
  });
  
  // 保存（保留所有原有格式）
  await workbook.toFileAsync('output.xlsx');
}
```

### 2. 创建新工作表

```javascript
const XlsxPopulate = require('xlsx-populate');

async function createSheet() {
  const workbook = await XlsxPopulate.fromFileAsync('input.xlsx');
  
  // 添加新工作表
  const newSheet = workbook.addSheet('新工作表');
  
  // 添加内容
  newSheet.cell('A1').value('标题');
  newSheet.cell('A1').style({
    bold: true,
    fontSize: 16,
    fill: '4472C4',
    fontColor: 'FFFFFF'
  });
  
  // 合并单元格 A1:D1
  newSheet.range('A1:D1').merged(true);
  newSheet.range('A1:D1').style({
    horizontalAlignment: 'center'
  });
  
  // 设置列宽
  newSheet.column('A').width(20);
  newSheet.column('B').width(30);
  
  await workbook.toFileAsync('output.xlsx');
}
```

### 3. 使用公式

```javascript
// 设置公式
sheet.cell('D2').formula('=B2*C2');
sheet.cell('D10').formula('=SUM(D2:D9)');

// 设置公式样式
sheet.cell('D2').style({
  fill: 'E7E6E6',
  bold: true
});
```

### 4. 批量写入数据

```javascript
const data = [
  ['项目', '单价', '数量', '小计'],
  ['项目A', 1000, 5],
  ['项目B', 2000, 3],
  ['项目C', 1500, 4],
  ['', '', '总计']
];

data.forEach((row, rowIndex) => {
  const rowNum = rowIndex + 1;
  row.forEach((value, colIndex) => {
    sheet.cell(rowNum, colIndex + 1).value(value);
  });
});

// 添加公式行
sheet.cell(5, 4).formula('=SUM(D2:D4)');
sheet.cell(5, 4).style({ fill: 'FFC000', bold: true });
```

### 5. 调整工作表顺序

```javascript
const sheets = workbook.sheets();
// 将最后一个工作表移到最前面
sheets[sheets.length - 1].move(0);
```

### 6. 处理多个工作表

```javascript
// 遍历所有工作表
workbook.sheets().forEach((sheet, index) => {
  console.log(`${index + 1}. ${sheet.name()}`);
  
  // 读取单元格值
  const value = sheet.cell('A1').value();
  console.log(`  A1: ${value}`);
});

// 通过名称获取工作表
const sheet = workbook.sheet('Sheet1');

// 重命名工作表
sheet.name('新名称');
```

## 样式参考

### 字体样式
```javascript
.cell('A1').style({
  bold: true,              // 粗体
  italic: true,            // 斜体
  underline: true,         // 下划线
  fontSize: 14,            // 字号
  fontColor: 'FF0000',     // 字体颜色（RGB）
  fontFamily: 'Arial'      // 字体
});
```

### 填充和背景
```javascript
.cell('A1').style({
  fill: '4472C4'           // 背景色（RGB）
});
```

### 对齐方式
```javascript
.range('A1:D1').style({
  horizontalAlignment: 'center',  // 水平：left, center, right
  verticalAlignment: 'center'     // 垂直：top, center, bottom
});
```

### 边框
```javascript
.cell('A1').style({
  border: true,            // 显示边框
  borderColor: '000000',   // 边框颜色
  borderStyle: 'thin'      // 边框样式：thin, medium, thick
});
```

### 数字格式
```javascript
.cell('B2').style({
  numberFormat: '0.00'     // 保留2位小数
});

.cell('C3').style({
  numberFormat: '0%'       // 百分比
});

.cell('D4').style({
  numberFormat: 'yyyy-mm-dd'  // 日期格式
});
```

## 合并单元格

```javascript
// 合并 A1:C3
sheet.range('A1:C3').merged(true);

// 合并后设置值和样式
sheet.range('A1:C3')
  .value('合并后的内容')
  .style({
    horizontalAlignment: 'center',
    verticalAlignment: 'center',
    bold: true
  });
```

## 列宽和行高

```javascript
// 设置列宽
sheet.column('A').width(20);
sheet.column('B').width(30);

// 设置行高
sheet.row(1).height(30);

// 自动调整列宽
sheet.column('A').hidden(false);
```

## 条件格式（高级）

```javascript
// 根据值设置不同样式
if (sheet.cell('B2').value() > 1000) {
  sheet.cell('B2').style({
    fill: '90EE90',        // 浅绿色
    fontColor: '006400'    // 深绿色
  });
} else {
  sheet.cell('B2').style({
    fill: 'FFB6C1',        // 浅红色
    fontColor: '8B0000'    // 深红色
  });
}
```

## 完整示例：编辑报价单

```javascript
const XlsxPopulate = require('xlsx-populate');

async function editQuotation() {
  const workbook = await XlsxPopulate.fromFileAsync('原始报价单.xlsx');
  
  // 1. 在首页添加编辑标注
  const firstSheet = workbook.sheet(0);
  firstSheet.cell('H1').value('【AI编辑版】生成时间: ' + new Date().toLocaleString('zh-CN'));
  firstSheet.cell('H1').style({
    fontColor: 'FF0000',
    bold: true,
    italic: true
  });
  
  // 2. 添加数据统计工作表
  const summarySheet = workbook.addSheet('数据统计');
  
  // 标题
  summarySheet.cell('A1').value('📊 报价单统计');
  summarySheet.cell('A1').style({
    fontSize: 16,
    bold: true,
    fontColor: 'FFFFFF',
    fill: '70AD47'
  });
  summarySheet.range('A1:D1').merged(true).style({
    horizontalAlignment: 'center'
  });
  
  // 统计表格
  const stats = [
    ['统计项', '数值'],
    ['工作表数量', workbook.sheets().length],
    ['编辑日期', new Date().toLocaleDateString('zh-CN')]
  ];
  
  stats.forEach((row, i) => {
    const rowNum = i + 3;
    row.forEach((val, j) => {
      const cell = summarySheet.cell(rowNum, j + 1);
      cell.value(val);
      if (i === 0) {
        cell.style({ bold: true, fill: '4472C4', fontColor: 'FFFFFF' });
      }
    });
  });
  
  // 设置列宽
  summarySheet.column('A').width(20);
  summarySheet.column('B').width(30);
  
  // 3. 调整工作表顺序
  const sheets = workbook.sheets();
  sheets[sheets.length - 1].move(0); // 新sheet移到最前
  
  // 保存
  await workbook.toFileAsync('编辑后报价单.xlsx');
  console.log('✅ 编辑完成！');
}

editQuotation().catch(console.error);
```

## 最佳实践

### 1. 始终保留原文件
```javascript
// 不要直接覆盖原文件
await workbook.toFileAsync('新文件名.xlsx');
```

### 2. 使用公式而非硬编码值
```javascript
// ✅ 正确
sheet.cell('D10').formula('=SUM(D2:D9)');

// ❌ 错误
const sum = calculateSum();
sheet.cell('D10').value(sum);
```

### 3. 批量操作提高效率
```javascript
// 使用数组批量写入，比逐个cell快
const data = [...];
data.forEach((row, i) => {
  row.forEach((val, j) => {
    sheet.cell(i + 1, j + 1).value(val);
  });
});
```

### 4. 错误处理
```javascript
async function safeEdit() {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('file.xlsx');
    // ... 编辑操作
    await workbook.toFileAsync('output.xlsx');
    console.log('✅ 成功');
  } catch (error) {
    console.error('❌ 错误:', error.message);
    process.exit(1);
  }
}
```

## 常见问题

### Q: 如何检查合并单元格？
```javascript
const merges = sheet._mergeCells;
console.log('合并单元格数量:', Object.keys(merges).length);
```

### Q: 如何复制工作表？
```javascript
const original = workbook.sheet(0);
const clone = original.clone('副本');
```

### Q: 如何删除工作表？
```javascript
workbook.deleteSheet('Sheet2');
```

### Q: 如何设置打印区域？
```javascript
sheet.printArea('A1:D20');
```

## 参考链接

- [xlsx-populate GitHub](https://github.com/dtjohnson/xlsx-populate)
- [xlsx-populate 文档](https://github.com/dtjohnson/xlsx-populate/blob/master/docs/tutorial.md)
