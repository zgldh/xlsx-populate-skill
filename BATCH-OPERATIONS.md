# Batch Operations for Large Datasets

Efficiently handle large Excel datasets with batch operations.

## Table of Contents

- [Writing Large Datasets](#writing-large-datasets)
- [Reading Large Files](#reading-large-files)
- [Streaming Approach](#streaming-approach)
- [Performance Tips](#performance-tips)

## Writing Large Datasets

```javascript
const XlsxPopulate = require('xlsx-populate');

async function writeLargeDataset() {
  const workbook = await XlsxPopulate.fromBlankAsync();
  const sheet = workbook.sheet(0);
  sheet.name('Data');
  
  // Write headers
  const headers = ['ID', 'Name', 'Value', 'Total'];
  headers.forEach((h, i) => {
    sheet.cell(1, i + 1).value(h).style({
      bold: true,
      fill: '4472C4',
      fontColor: 'FFFFFF'
    });
  });
  
  // Batch write data (much faster than individual cell writes)
  const data = generateData(10000); // 10k rows
  data.forEach((row, index) => {
    const rowNum = index + 2;
    row.forEach((value, colIndex) => {
      sheet.cell(rowNum, colIndex + 1).value(value);
    });
  });
  
  await workbook.toFileAsync('large-dataset.xlsx');
}
```

## Reading Large Files

```javascript
// For read-only operations on large files
const workbook = await XlsxPopulate.fromFileAsync('large.xlsx', {
  readOnly: true
});

// Access specific cells without loading entire file
const value = workbook.sheet(0).cell('A1').value();
```

## Streaming Approach

For extremely large datasets, consider CSV intermediate:

```javascript
const fs = require('fs');
const csv = require('csv-parser');

async function streamToExcel() {
  const workbook = await XlsxPopulate.fromBlankAsync();
  const sheet = workbook.sheet(0);
  
  let rowNum = 1;
  fs.createReadStream('data.csv')
    .pipe(csv())
    .on('data', (row) => {
      Object.values(row).forEach((val, colIndex) => {
        sheet.cell(rowNum, colIndex + 1).value(val);
      });
      rowNum++;
    })
    .on('end', async () => {
      await workbook.toFileAsync('output.xlsx');
    });
}
```

## Performance Tips

1. **Batch writes**: Always use array iteration, not individual cell access in loops
2. **Write-only mode**: Use `writeOnly: true` for new files
3. **Avoid excessive styling**: Apply styles to ranges, not individual cells
4. **Lazy loading**: Read cells only when needed

```javascript
// ❌ Slow - individual cell access
for (let i = 0; i < 1000; i++) {
  for (let j = 0; j < 10; j++) {
    sheet.cell(i + 1, j + 1).value(data[i][j]);
    sheet.cell(i + 1, j + 1).style({ ... });  // Don't style in loop!
  }
}

// ✅ Fast - batch write
const data = [...]; // Your data array
data.forEach((row, i) => {
  row.forEach((val, j) => {
    sheet.cell(i + 1, j + 1).value(val);
  });
});
// Style entire range at once
sheet.range('A1:J1000').style({ fontSize: 10 });
```
