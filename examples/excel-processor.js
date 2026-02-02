const XlsxPopulate = require('xlsx-populate');

/**
 * Excel自动化处理类
 * 
 * 封装常用的Excel操作，便于复用
 */
class ExcelProcessor {
  constructor() {
    this.workbook = null;
  }

  /**
   * 加载Excel文件
   * @param {string} filePath - 文件路径
   */
  async load(filePath) {
    this.workbook = await XlsxPopulate.fromFileAsync(filePath);
    return this;
  }

  /**
   * 从空白创建
   */
  async createBlank() {
    this.workbook = await XlsxPopulate.fromBlankAsync();
    return this;
  }

  /**
   * 获取工作表
   * @param {number|string} indexOrName - 索引或名称
   */
  sheet(indexOrName) {
    return this.workbook.sheet(indexOrName);
  }

  /**
   * 获取所有工作表
   */
  sheets() {
    return this.workbook.sheets();
  }

  /**
   * 添加工作表
   * @param {string} name - 工作表名称
   */
  addSheet(name) {
    return this.workbook.addSheet(name);
  }

  /**
   * 删除工作表
   * @param {string} name - 工作表名称
   */
  deleteSheet(name) {
    this.workbook.deleteSheet(name);
    return this;
  }

  /**
   * 重命名工作表
   * @param {number|string} indexOrName - 索引或名称
   * @param {string} newName - 新名称
   */
  renameSheet(indexOrName, newName) {
    this.sheet(indexOrName).name(newName);
    return this;
  }

  /**
   * 移动工作表
   * @param {number|string} indexOrName - 索引或名称
   * @param {number} toIndex - 目标位置
   */
  moveSheet(indexOrName, toIndex) {
    this.sheet(indexOrName).move(toIndex);
    return this;
  }

  /**
   * 设置单元格值
   * @param {string} cell - 单元格（如 'A1'）
   * @param {any} value - 值
   * @param {number} sheetIndex - 工作表索引
   */
  setValue(cell, value, sheetIndex = 0) {
    this.sheet(sheetIndex).cell(cell).value(value);
    return this;
  }

  /**
   * 设置单元格样式
   * @param {string} cell - 单元格
   * @param {object} style - 样式对象
   * @param {number} sheetIndex - 工作表索引
   */
  setStyle(cell, style, sheetIndex = 0) {
    this.sheet(sheetIndex).cell(cell).style(style);
    return this;
  }

  /**
   * 设置公式
   * @param {string} cell - 单元格
   * @param {string} formula - 公式
   * @param {number} sheetIndex - 工作表索引
   */
  setFormula(cell, formula, sheetIndex = 0) {
    this.sheet(sheetIndex).cell(cell).formula(formula);
    return this;
  }

  /**
   * 批量写入数据
   * @param {array} data - 二维数组数据
   * @param {number} startRow - 起始行
   * @param {number} sheetIndex - 工作表索引
   */
  writeData(data, startRow = 1, sheetIndex = 0) {
    const sheet = this.sheet(sheetIndex);
    data.forEach((row, rowIndex) => {
      const rowNum = startRow + rowIndex;
      row.forEach((value, colIndex) => {
        sheet.cell(rowNum, colIndex + 1).value(value);
      });
    });
    return this;
  }

  /**
   * 合并单元格
   * @param {string} range - 范围（如 'A1:C3'）
   * @param {number} sheetIndex - 工作表索引
   */
  mergeCells(range, sheetIndex = 0) {
    this.sheet(sheetIndex).range(range).merged(true);
    return this;
  }

  /**
   * 设置列宽
   * @param {string} column - 列（如 'A'）
   * @param {number} width - 宽度
   * @param {number} sheetIndex - 工作表索引
   */
  setColumnWidth(column, width, sheetIndex = 0) {
    this.sheet(sheetIndex).column(column).width(width);
    return this;
  }

  /**
   * 设置行高
   * @param {number} row - 行号
   * @param {number} height - 高度
   * @param {number} sheetIndex - 工作表索引
   */
  setRowHeight(row, height, sheetIndex = 0) {
    this.sheet(sheetIndex).row(row).height(height);
    return this;
  }

  /**
   * 设置表头样式（蓝色背景，白色字体）
   * @param {string} range - 范围
   * @param {number} sheetIndex - 工作表索引
   */
  setHeaderStyle(range, sheetIndex = 0) {
    this.sheet(sheetIndex).range(range).style({
      bold: true,
      fontColor: 'FFFFFF',
      fill: '4472C4',
      horizontalAlignment: 'center'
    });
    return this;
  }

  /**
   * 保存文件
   * @param {string} filePath - 文件路径
   */
  async save(filePath) {
    await this.workbook.toFileAsync(filePath);
    return this;
  }

  /**
   * 获取工作表信息
   */
  getSheetInfo() {
    return this.workbook.sheets().map((sheet, index) => ({
      index,
      name: sheet.name(),
      usedRange: sheet.usedRange().address()
    }));
  }
}

// 导出
module.exports = ExcelProcessor;

// 使用示例
async function example() {
  const processor = new ExcelProcessor();
  
  // 加载文件
  await processor.load('./data/input.xlsx');
  
  // 链式调用进行编辑
  await processor
    .setValue('A1', '新标题', 0)
    .setStyle('A1', { bold: true, fontSize: 16 }, 0)
    .addSheet('新工作表')
    .writeData([
      ['列1', '列2', '列3'],
      ['数据1', '数据2', '数据3'],
      ['数据4', '数据5', '数据6']
    ], 1, 1)
    .setHeaderStyle('A1:C1', 1)
    .setColumnWidth('A', 20, 1)
    .setColumnWidth('B', 30, 1)
    .save('./output/output.xlsx');
  
  console.log('✅ 编辑完成！');
}

// 如果直接运行此文件，执行示例
if (require.main === module) {
  example().catch(console.error);
}
