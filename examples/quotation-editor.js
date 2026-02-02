const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');
const path = require('path');

/**
 * 报价单编辑器 - 实际应用场景示例
 * 
 * 功能：
 * 1. 读取原始报价单
 * 2. 保留所有格式
 * 3. 添加编辑标记
 * 4. 创建统计工作表
 * 5. 生成编辑版本
 */
async function editQuotation(inputFile, outputFile) {
  console.log(`📖 正在读取: ${inputFile}`);
  
  // 从文件加载（保留所有格式）
  const workbook = await XlsxPopulate.fromFileAsync(inputFile);
  
  console.log('\n📊 原始工作表:');
  workbook.sheets().forEach((sheet, index) => {
    console.log(`  ${index + 1}. ${sheet.name()}`);
  });
  
  // 获取第一个工作表
  const firstSheet = workbook.sheet(0);
  
  // 检查合并单元格
  const mergedCells = firstSheet._mergeCells || {};
  console.log(`\n📋 第一个工作表 "${firstSheet.name()}" 信息:`);
  console.log(`  - 合并单元格数量: ${Object.keys(mergedCells).length}`);
  
  // ==================== 编辑操作 ====================
  console.log('\n✏️  开始编辑（保留原有格式）...');
  
  // 1. 在第一个sheet的空白处添加标注
  firstSheet.cell('H1').value('【AI编辑版】生成时间: ' + new Date().toLocaleString('zh-CN'));
  firstSheet.cell('H1').style({
    fontColor: 'FF0000',
    bold: true,
    italic: true
  });
  
  // 2. 添加"数据统计"sheet
  console.log('📈 创建"数据统计"工作表...');
  const summarySheet = workbook.addSheet('数据统计');
  
  // 标题行
  summarySheet.cell('A1').value('📊 项目报价统计分析');
  summarySheet.cell('A1').style({
    fontSize: 16,
    bold: true,
    fontColor: 'FFFFFF',
    fill: '70AD47'
  });
  summarySheet.range('A1:D1').merged(true);
  summarySheet.range('A1:D1').style({
    horizontalAlignment: 'center',
    verticalAlignment: 'center'
  });
  
  // 统计信息表格
  const stats = [
    ['统计项', '数值', '说明'],
    ['原工作表数量', workbook.sheets().length - 1, '个'],
    ['编辑日期', new Date().toLocaleDateString('zh-CN'), ''],
    ['编辑时间', new Date().toLocaleTimeString('zh-CN'), ''],
    ['编辑人员', 'AI Assistant', 'OpenCode'],
    ['版本', 'V1.0-Edit', '编辑版'],
    ['', '', ''],
    ['💡 编辑说明', '', ''],
    ['1. 保留了所有原始工作表和格式', '', ''],
    ['2. 添加了此统计分析页', '', ''],
    ['3. 在首页面添加了编辑标注', '', ''],
    ['4. 合并单元格和样式完整保留', '', '']
  ];
  
  stats.forEach((row, index) => {
    const rowNum = index + 3;
    row.forEach((value, colIndex) => {
      const cell = summarySheet.cell(rowNum, colIndex + 1);
      cell.value(value);
      
      // 表头行样式
      if (index === 0) {
        cell.style({
          bold: true,
          fontColor: 'FFFFFF',
          fill: '4472C4'
        });
      }
      
      // 说明行样式
      if (index >= 7) {
        cell.style({
          italic: true,
          fontColor: '666666'
        });
      }
    });
  });
  
  // 设置列宽
  summarySheet.column('A').width(20);
  summarySheet.column('B').width(30);
  summarySheet.column('C').width(20);
  
  // 3. 添加"公式示例"sheet
  console.log('🔢 创建"公式示例"工作表...');
  const formulaSheet = workbook.addSheet('公式示例');
  
  // 标题
  formulaSheet.cell('A1').value('Excel公式演示');
  formulaSheet.cell('A1').style({
    fontSize: 14,
    bold: true,
    fontColor: 'FFFFFF',
    fill: '4472C4'
  });
  formulaSheet.range('A1:D1').merged(true);
  formulaSheet.range('A1:D1').style({
    horizontalAlignment: 'center'
  });
  
  // 表头
  const headers = ['项目', '单价', '数量', '小计'];
  headers.forEach((header, index) => {
    const cell = formulaSheet.cell(3, index + 1);
    cell.value(header);
    cell.style({
      bold: true,
      fontColor: 'FFFFFF',
      fill: '4472C4'
    });
  });
  
  // 数据行
  const data = [
    ['示例项目A', 1000, 5],
    ['示例项目B', 2000, 3],
    ['示例项目C', 1500, 4]
  ];
  
  data.forEach((row, rowIndex) => {
    const rowNum = rowIndex + 4;
    row.forEach((value, colIndex) => {
      formulaSheet.cell(rowNum, colIndex + 1).value(value);
    });
    // 添加公式计算小计
    formulaSheet.cell(rowNum, 4).formula(`=B${rowNum}*C${rowNum}`);
    formulaSheet.cell(rowNum, 4).style({
      fill: 'E7E6E6'
    });
  });
  
  // 总计行
  formulaSheet.cell(7, 3).value('总计');
  formulaSheet.cell(7, 3).style({ bold: true });
  formulaSheet.cell(7, 4).formula('=SUM(D4:D6)');
  formulaSheet.cell(7, 4).style({
    bold: true,
    fill: 'FFC000'
  });
  
  // 设置列宽
  formulaSheet.column('A').width(15);
  formulaSheet.column('B').width(12);
  formulaSheet.column('C').width(12);
  formulaSheet.column('D').width(12);
  
  // 4. 调整工作表顺序
  console.log('🔄 调整工作表顺序...');
  const sheets = workbook.sheets();
  const newOrder = [
    sheets[sheets.length - 2], // 数据统计
    sheets[sheets.length - 1], // 公式示例
    ...sheets.slice(0, -2)       // 其他原始sheet
  ];
  
  // 重新排序
  newOrder.forEach((sheet, index) => {
    sheet.move(index);
  });
  
  // 确保输出目录存在
  const outputDir = path.dirname(outputFile);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  
  // 保存文件
  console.log('\n💾 保存编辑后的文件...');
  await workbook.toFileAsync(outputFile);
  
  console.log('\n✅ 编辑完成!');
  console.log(`📁 输出文件: ${outputFile}`);
  console.log('\n📊 最终工作表列表:');
  workbook.sheets().forEach((sheet, index) => {
    console.log(`  ${index + 1}. ${sheet.name()}`);
  });
  
  console.log('\n🎉 成功保留的内容:');
  console.log('  ✓ 所有原始工作表');
  console.log('  ✓ 原有格式和样式');
  console.log('  ✓ 合并单元格');
  console.log('  ✓ 列宽设置');
  console.log('\n📝 新增内容:');
  console.log('  ✓ 首页面标注（H1单元格）');
  console.log('  ✓ "数据统计"工作表');
  console.log('  ✓ "公式示例"工作表（含公式）');
}

// 主函数
async function main() {
  const inputFile = process.argv[2] || './data/quotation.xlsx';
  const outputFile = process.argv[3] || './output/quotation-edited.xlsx';
  
  console.log('═══════════════════════════════════════════');
  console.log('  报价单编辑器 - xlsx-populate Skill');
  console.log('═══════════════════════════════════════════\n');
  
  try {
    // 检查输入文件是否存在
    if (!fs.existsSync(inputFile)) {
      console.error(`❌ 错误: 输入文件不存在: ${inputFile}`);
      console.log('\n使用方法:');
      console.log('  node quotation-editor.js <输入文件> <输出文件>');
      console.log('\n示例:');
      console.log('  node quotation-editor.js ./data/input.xlsx ./output/output.xlsx');
      process.exit(1);
    }
    
    await editQuotation(inputFile, outputFile);
    
    console.log('\n═══════════════════════════════════════════');
    console.log('  ✅ 处理完成！');
    console.log('═══════════════════════════════════════════');
  } catch (error) {
    console.error('\n❌ 错误:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

// 运行
main();
