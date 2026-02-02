# Real-World Examples

Complete, production-ready examples for common scenarios.

## Table of Contents

- [Example 1: Sales Report Generator](#example-1-sales-report-generator)
- [Example 2: Invoice Processor](#example-2-invoice-processor)
- [Example 3: Data Migration](#example-3-data-migration)
- [Example 4: Template-Based Report](#example-4-template-based-report)

## Example 1: Sales Report Generator

Generate monthly sales report with charts-ready data.

```javascript
const XlsxPopulate = require('xlsx-populate');

async function generateSalesReport(salesData, month) {
  const workbook = await XlsxPopulate.fromBlankAsync();
  const sheet = workbook.sheet(0);
  sheet.name(`${month} Sales Report`);
  
  // Title
  sheet.cell('A1').value(`${month} Sales Report`);
  sheet.cell('A1').style({
    fontSize: 18,
    bold: true,
    fontColor: 'FFFFFF',
    fill: '4472C4'
  });
  sheet.range('A1:E1').merged(true);
  
  // Headers
  const headers = ['Product', 'Units Sold', 'Unit Price', 'Revenue', 'Growth %'];
  headers.forEach((h, i) => {
    sheet.cell(3, i + 1).value(h).style({
      bold: true,
      fontColor: 'FFFFFF',
      fill: '4472C4'
    });
  });
  
  // Data rows
  salesData.forEach((row, index) => {
    const rowNum = index + 4;
    sheet.cell(rowNum, 1).value(row.product);
    sheet.cell(rowNum, 2).value(row.units);
    sheet.cell(rowNum, 3).value(row.price).style({ numberFormat: '$#,##0.00' });
    sheet.cell(rowNum, 4).formula(`=B${rowNum}*C${rowNum}`).style({ numberFormat: '$#,##0.00' });
    sheet.cell(rowNum, 5).value(row.growth).style({ numberFormat: '0.0%' });
  });
  
  // Totals
  const totalRow = salesData.length + 4;
  sheet.cell(totalRow, 1).value('TOTAL');
  sheet.cell(totalRow, 1).style({ bold: true });
  sheet.cell(totalRow, 2).formula(`=SUM(B4:B${totalRow - 1})`);
  sheet.cell(totalRow, 4).formula(`=SUM(D4:D${totalRow - 1})`).style({
    bold: true,
    fill: 'FFC000',
    numberFormat: '$#,##0.00'
  });
  
  // Column widths
  sheet.column('A').width(20);
  sheet.column('B').width(12);
  sheet.column('C').width(12);
  sheet.column('D').width(12);
  sheet.column('E').width(12);
  
  await workbook.toFileAsync(`sales-report-${month}.xlsx`);
}

// Usage
const data = [
  { product: 'Widget A', units: 150, price: 29.99, growth: 0.15 },
  { product: 'Widget B', units: 230, price: 49.99, growth: 0.23 },
  // ... more data
];
generateSalesReport(data, 'January 2024');
```

## Example 2: Invoice Processor

Process template invoice and fill in customer data.

```javascript
const XlsxPopulate = require('xlsx-populate');

async function generateInvoice(templatePath, invoiceData) {
  // Load template
  const workbook = await XlsxPopulate.fromFileAsync(templatePath);
  const sheet = workbook.sheet('Invoice');
  
  // Fill customer info
  sheet.cell('B3').value(invoiceData.customerName);
  sheet.cell('B4').value(invoiceData.customerAddress);
  sheet.cell('B5').value(invoiceData.customerEmail);
  
  // Invoice details
  sheet.cell('E3').value(invoiceData.invoiceNumber);
  sheet.cell('E4').value(invoiceData.invoiceDate);
  sheet.cell('E5').value(invoiceData.dueDate);
  
  // Line items
  invoiceData.items.forEach((item, index) => {
    const rowNum = 9 + index;
    sheet.cell(rowNum, 1).value(index + 1);
    sheet.cell(rowNum, 2).value(item.description);
    sheet.cell(rowNum, 3).value(item.quantity);
    sheet.cell(rowNum, 4).value(item.unitPrice).style({ numberFormat: '$#,##0.00' });
    sheet.cell(rowNum, 5).formula(`=C${rowNum}*D${rowNum}`).style({ numberFormat: '$#,##0.00' });
  });
  
  // Totals (assuming template has formula cells)
  sheet.cell('E20').formula(`=SUM(E9:E${8 + invoiceData.items.length})`);
  
  await workbook.toFileAsync(`invoice-${invoiceData.invoiceNumber}.xlsx`);
}
```

## Example 3: Data Migration

Migrate data from old format to new format while preserving audit trail.

```javascript
const XlsxPopulate = require('xlsx-populate');

async function migrateData(oldFilePath, newFilePath) {
  // Load old file
  const oldWorkbook = await XlsxPopulate.fromFileAsync(oldFilePath);
  const oldSheet = oldWorkbook.sheet(0);
  
  // Create new workbook
  const newWorkbook = await XlsxPopulate.fromBlankAsync();
  const newSheet = newWorkbook.sheet(0);
  newSheet.name('Migrated Data');
  
  // Migration mapping
  const mapping = {
    'A': 'A',  // ID stays same
    'B': 'C',  // Old Name -> New CustomerName
    'C': 'D',  // Old Email -> New Contact
    'D': 'B',  // Old Date -> New OrderDate
  };
  
  // Copy headers
  newSheet.cell('A1').value('ID');
  newSheet.cell('B1').value('OrderDate');
  newSheet.cell('C1').value('CustomerName');
  newSheet.cell('D1').value('Contact');
  newSheet.range('A1:D1').style({
    bold: true,
    fill: '4472C4',
    fontColor: 'FFFFFF'
  });
  
  // Migrate data
  let row = 2;
  while (oldSheet.cell(row, 1).value()) {
    Object.entries(mapping).forEach(([oldCol, newCol]) => {
      const value = oldSheet.cell(row, oldCol.charCodeAt(0) - 64).value();
      newSheet.cell(row, newCol.charCodeAt(0) - 64).value(value);
    });
    row++;
  }
  
  // Add migration metadata
  const metaSheet = newWorkbook.addSheet('Migration Info');
  metaSheet.cell('A1').value('Migration Date:');
  metaSheet.cell('B1').value(new Date().toISOString());
  metaSheet.cell('A2').value('Source File:');
  metaSheet.cell('B2').value(oldFilePath);
  metaSheet.cell('A3').value('Records Migrated:');
  metaSheet.cell('B3').value(row - 2);
  
  await newWorkbook.toFileAsync(newFilePath);
}
```

## Example 4: Template-Based Report

Fill template with dynamic data and multiple sections.

```javascript
const XlsxPopulate = require('xlsx-populate');

async function generateProjectReport(templatePath, projectData) {
  const workbook = await XlsxPopulate.fromFileAsync(templatePath);
  
  // Fill cover page
  const coverSheet = workbook.sheet('Cover');
  coverSheet.cell('B3').value(projectData.projectName);
  coverSheet.cell('B4').value(projectData.clientName);
  coverSheet.cell('B5').value(projectData.projectManager);
  coverSheet.cell('E3').value(new Date().toLocaleDateString());
  
  // Fill summary sheet
  const summarySheet = workbook.sheet('Summary');
  summarySheet.cell('B2').value(projectData.budget);
  summarySheet.cell('B3').value(projectData.spent);
  summarySheet.cell('B4').formula('=B2-B3');  // Remaining
  summarySheet.cell('B5').value(projectData.progress).style({ numberFormat: '0%' });
  
  // Fill milestones
  const milestoneSheet = workbook.sheet('Milestones');
  projectData.milestones.forEach((milestone, index) => {
    const rowNum = index + 2;
    milestoneSheet.cell(rowNum, 1).value(milestone.name);
    milestoneSheet.cell(rowNum, 2).value(milestone.date);
    milestoneSheet.cell(rowNum, 3).value(milestone.status);
    
    // Color code status
    const statusColors = {
      'Complete': { fill: '90EE90', fontColor: '006400' },
      'In Progress': { fill: 'FFD966', fontColor: 'B7950B' },
      'Pending': { fill: 'D9D9D9', fontColor: '666666' }
    };
    
    if (statusColors[milestone.status]) {
      milestoneSheet.cell(rowNum, 3).style(statusColors[milestone.status]);
    }
  });
  
  await workbook.toFileAsync(`project-report-${projectData.projectName}.xlsx`);
}
```

## Error Handling Pattern

```javascript
async function safeExcelOperation(operation, inputFile, outputFile) {
  try {
    console.log(`📖 Loading: ${inputFile}`);
    const workbook = await XlsxPopulate.fromFileAsync(inputFile);
    
    console.log('✏️  Processing...');
    await operation(workbook);
    
    console.log(`💾 Saving: ${outputFile}`);
    await workbook.toFileAsync(outputFile);
    
    console.log('✅ Complete!');
    return true;
  } catch (error) {
    console.error('❌ Error:', error.message);
    if (error.code === 'ENOENT') {
      console.error('   File not found');
    } else if (error.message.includes('password')) {
      console.error('   File is password protected');
    }
    return false;
  }
}

// Usage
safeExcelOperation(
  (workbook) => {
    workbook.sheet(0).cell('A1').value('Updated');
  },
  'input.xlsx',
  'output.xlsx'
);
```
