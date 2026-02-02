# Excel Styling Guide

Color schemes, formatting patterns, and professional styling standards.

## Table of Contents

- [Color Schemes](#color-schemes)
- [Professional Formatting](#professional-formatting)
- [Cell Styles by Purpose](#cell-styles-by-purpose)

## Color Schemes

### Corporate Blue (Default)

```javascript
// Header row
sheet.range('A1:F1').style({
  fill: '4472C4',        // Blue
  fontColor: 'FFFFFF',   // White
  bold: true
});

// Alternate rows
sheet.range('A2:F2').style({ fill: 'D9E2F3' });  // Light blue
sheet.range('A3:F3').style({ fill: 'FFFFFF' });  // White
```

### Financial Green

```javascript
// Positive values
sheet.cell('B5').style({
  fontColor: '006400',   // Dark green
  fill: '90EE90'         // Light green
});

// Negative values
sheet.cell('B6').style({
  fontColor: '8B0000',   // Dark red
  fill: 'FFB6C1'         // Light red
});
```

### Executive Summary

```javascript
// Title
sheet.cell('A1').style({
  fontSize: 18,
  bold: true,
  fontColor: '1F4E79',
  fill: 'D9E2F3'
});

// Total row
sheet.range('A10:D10').style({
  bold: true,
  fill: 'FFC000',        // Gold
  border: true
});
```

## Professional Formatting

### Headers

```javascript
function applyHeaderStyle(cell) {
  cell.style({
    bold: true,
    fontSize: 11,
    fontColor: 'FFFFFF',
    fill: '4472C4',
    horizontalAlignment: 'center',
    verticalAlignment: 'center'
  });
}
```

### Data Cells

```javascript
function applyDataStyle(cell, format = 'general') {
  const formats = {
    general: {},
    currency: { numberFormat: '$#,##0.00' },
    percentage: { numberFormat: '0.0%' },
    date: { numberFormat: 'yyyy-mm-dd' },
    number: { numberFormat: '#,##0' }
  };
  
  cell.style({
    fontSize: 10,
    verticalAlignment: 'center',
    ...formats[format]
  });
}
```

### Borders

```javascript
// Add borders to range
sheet.range('A1:D10').style({
  border: true,
  borderStyle: 'thin',
  borderColor: '000000'
});

// Outline only
sheet.range('A1:D10').style({
  border: true,
  borderStyle: 'medium'
});
```

## Cell Styles by Purpose

### Input Cells (Assumptions)

```javascript
sheet.cell('B2').style({
  fontColor: '0070C0',   // Blue text
  fill: 'F2F2F2',        // Light gray fill
  italic: true
});
```

### Calculated Cells

```javascript
sheet.cell('C2').style({
  fontColor: '000000',   // Black text
  fill: 'FFFFFF'
});
// Add formula
sheet.cell('C2').formula('=B2*1.1');
```

### Linked Cells (From Other Sheets)

```javascript
sheet.cell('D2').style({
  fontColor: '00B050',   // Green text
  fill: 'FFFFFF'
});
sheet.cell('D2').formula('=Summary!A1');
```

### Warning/Highlight

```javascript
sheet.cell('E2').style({
  fill: 'FFFF00',        // Yellow
  fontColor: 'FF0000',   // Red text
  bold: true
});
```

## Number Formats

```javascript
// Currency
sheet.cell('A1').style({ numberFormat: '$#,##0.00' });

// Percentage
sheet.cell('B1').style({ numberFormat: '0.0%' });

// Date
sheet.cell('C1').style({ numberFormat: 'yyyy-mm-dd' });

// Thousands separator
sheet.cell('D1').style({ numberFormat: '#,##0' });

// Accounting
sheet.cell('E1').style({ numberFormat: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' });
```
