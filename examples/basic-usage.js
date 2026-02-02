const XlsxPopulate = require('xlsx-populate');

/**
 * 示例1：读取Excel文件并保留格式编辑
 */
async function example1_ReadAndEdit() {
  console.log('📖 示例1：读取并编辑Excel（保留格式）');
  
  // 从文件加载（保留所有格式）
  const workbook = await XlsxPopulate.fromFileAsync('./data/input.xlsx');
  
  // 获取第一个工作表
  const sheet = workbook.sheet(0);
  
  // 修改单元格值
  sheet.cell('A1').value('编辑后的标题');
  
  // 应用样式
  sheet.cell('A1').style({
    bold: true,
    fontColor: 'FF0000',  // 红色
    fontSize: 16,
    fill: 'FFFF00'        // 黄色背景
  });
  
  // 保存到新文件（原文件格式完全保留）
  await workbook.toFileAsync('./output/example1_output.xlsx');
  console.log('✅ 已保存到 output/example1_output.xlsx\n');
}

/**
 * 示例2：创建新工作表并添加数据
 */
async function example2_CreateNewSheet() {
  console.log('📊 示例2：创建新工作表');
  
  const workbook = await XlsxPopulate.fromFileAsync('./data/input.xlsx');
  
  // 添加新工作表
  const newSheet = workbook.addSheet('销售报表');
  
  // 添加标题（合并单元格）
  newSheet.cell('A1').value('2024年销售报表');
  newSheet.cell('A1').style({
    fontSize: 18,
    bold: true,
    fontColor: 'FFFFFF',
    fill: '4472C4'  // 蓝色
  });
  newSheet.range('A1:D1').merged(true);
  newSheet.range('A1:D1').style({
    horizontalAlignment: 'center',
    verticalAlignment: 'center'
  });
  
  // 添加表头
  const headers = ['产品', '单价', '销量', '销售额'];
  headers.forEach((header, index) => {
    const cell = newSheet.cell(3, index + 1);
    cell.value(header);
    cell.style({
      bold: true,
      fontColor: 'FFFFFF',
      fill: '4472C4'
    });
  });
  
  // 添加数据
  const data = [
    ['产品A', 100, 50],
    ['产品B', 200, 30],
    ['产品C', 150, 40],
    ['产品D', 300, 20]
  ];
  
  data.forEach((row, rowIndex) => {
    const rowNum = rowIndex + 4;
    row.forEach((value, colIndex) => {
      newSheet.cell(rowNum, colIndex + 1).value(value);
    });
    // 添加公式计算销售额
    newSheet.cell(rowNum, 4).formula(`=B${rowNum}*C${rowNum}`);
    newSheet.cell(rowNum, 4).style({ fill: 'E7E6E6' });
  });
  
  // 添加总计行
  newSheet.cell(8, 3).value('总计');
  newSheet.cell(8, 3).style({ bold: true });
  newSheet.cell(8, 4).formula('=SUM(D4:D7)');
  newSheet.cell(8, 4).style({ 
    bold: true, 
    fill: 'FFC000'  // 金色
  });
  
  // 设置列宽
  newSheet.column('A').width(15);
  newSheet.column('B').width(12);
  newSheet.column('C').width(12);
  newSheet.column('D').width(12);
  
  await workbook.toFileAsync('./output/example2_output.xlsx');
  console.log('✅ 已保存到 output/example2_output.xlsx\n');
}

/**
 * 示例3：批量数据处理
 */
async function example3_BatchProcessing() {
  console.log('📈 示例3：批量数据处理');
  
  const workbook = await XlsxPopulate.fromFileAsync('./data/input.xlsx');
  const sheet = workbook.sheet(0);
  
  // 准备批量数据
  const batchData = [];
  for (let i = 1; i <= 100; i++) {
    batchData.push([
      `项目${i}`,
      Math.floor(Math.random() * 1000) + 100,
      Math.floor(Math.random() * 50) + 1
    ]);
  }
  
  // 批量写入
  console.log(`正在写入 ${batchData.length} 行数据...`);
  batchData.forEach((row, index) => {
    const rowNum = index + 2;
    row.forEach((value, colIndex) => {
      sheet.cell(rowNum, colIndex + 1).value(value);
    });
    // 添加公式
    sheet.cell(rowNum, 4).formula(`=B${rowNum}*C${rowNum}`);
  });
  
  // 添加总计
  const lastRow = batchData.length + 1;
  sheet.cell(lastRow + 1, 3).value('总计');
  sheet.cell(lastRow + 1, 3).style({ bold: true });
  sheet.cell(lastRow + 1, 4).formula(`=SUM(D2:D${lastRow})`);
  sheet.cell(lastRow + 1, 4).style({ 
    bold: true, 
    fill: 'FFC000',
    fontSize: 14
  });
  
  await workbook.toFileAsync('./output/example3_output.xlsx');
  console.log('✅ 已保存到 output/example3_output.xlsx\n');
}

/**
 * 示例4：样式和格式化
 */
async function example4_Styles() {
  console.log('🎨 示例4：样式和格式化');
  
  const workbook = await XlsxPopulate.fromFileAsync('./data/input.xlsx');
  const sheet = workbook.sheet(0);
  
  // 字体样式
  sheet.cell('A1').style({
    bold: true,
    italic: true,
    underline: true,
    fontSize: 20,
    fontColor: 'FF0000',
    fontFamily: 'Microsoft YaHei'
  });
  
  // 填充颜色
  sheet.cell('A2').value('背景色示例');
  sheet.cell('A2').style({
    fill: '90EE90'  // 浅绿色
  });
  
  // 对齐方式
  sheet.cell('A3').value('居中对齐');
  sheet.cell('A3').style({
    horizontalAlignment: 'center',
    verticalAlignment: 'center'
  });
  
  // 边框
  sheet.cell('A4').value('带边框');
  sheet.cell('A4').style({
    border: true,
    borderColor: '000000',
    borderStyle: 'thick'
  });
  
  // 数字格式
  sheet.cell('B1').value(1234.5678);
  sheet.cell('B1').style({
    numberFormat: '0.00'  // 保留2位小数
  });
  
  sheet.cell('B2').value(0.85);
  sheet.cell('B2').style({
    numberFormat: '0%'  // 百分比
  });
  
  sheet.cell('B3').value(new Date());
  sheet.cell('B3').style({
    numberFormat: 'yyyy-mm-dd'  // 日期格式
  });
  
  // 合并单元格样式
  sheet.range('C1:E3').merged(true);
  sheet.range('C1:E3').value('合并单元格');
  sheet.range('C1:E3').style({
    horizontalAlignment: 'center',
    verticalAlignment: 'center',
    fill: 'FFB6C1',
    fontSize: 14,
    bold: true
  });
  
  await workbook.toFileAsync('./output/example4_output.xlsx');
  console.log('✅ 已保存到 output/example4_output.xlsx\n');
}

/**
 * 示例5：处理多个工作表
 */
async function example5_MultipleSheets() {
  console.log('📑 示例5：处理多个工作表');
  
  const workbook = await XlsxPopulate.fromFileAsync('./data/input.xlsx');
  
  // 遍历所有工作表
  console.log('工作表列表：');
  workbook.sheets().forEach((sheet, index) => {
    console.log(`  ${index + 1}. ${sheet.name()}`);
  });
  
  // 在每个工作表添加页脚
  workbook.sheets().forEach((sheet, index) => {
    const lastRow = sheet.usedRange().endCell().rowNumber();
    const footerCell = sheet.cell(lastRow + 2, 1);
    footerCell.value(`工作表 ${index + 1}: ${sheet.name()} - 编辑时间: ${new Date().toLocaleString('zh-CN')}`);
    footerCell.style({
      italic: true,
      fontColor: '666666',
      fontSize: 10
    });
  });
  
  // 调整工作表顺序
  const sheets = workbook.sheets();
  if (sheets.length > 1) {
    sheets[sheets.length - 1].move(0); // 最后一个移到第一个
    console.log('已调整工作表顺序');
  }
  
  await workbook.toFileAsync('./output/example5_output.xlsx');
  console.log('✅ 已保存到 output/example5_output.xlsx\n');
}

/**
 * 主函数：运行所有示例
 */
async function main() {
  console.log('═══════════════════════════════════════');
  console.log('  xlsx-populate Skill 示例程序');
  console.log('═══════════════════════════════════════\n');
  
  try {
    await example1_ReadAndEdit();
    await example2_CreateNewSheet();
    await example3_BatchProcessing();
    await example4_Styles();
    await example5_MultipleSheets();
    
    console.log('═══════════════════════════════════════');
    console.log('  ✅ 所有示例运行完成！');
    console.log('═══════════════════════════════════════');
  } catch (error) {
    console.error('❌ 错误:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

// 运行
main();
