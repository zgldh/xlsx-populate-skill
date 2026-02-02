# Excel Formula Patterns

Standard formula patterns for financial modeling and data analysis.

## Table of Contents

- [Financial Modeling Standards](#financial-modeling-standards)
- [Common Formulas](#common-formulas)
- [Formula Best Practices](#formula-best-practices)

## Financial Modeling Standards

Follow these color conventions:

- **Blue text**: Hardcoded inputs (assumptions)
- **Black text**: Formulas only
- **Green text**: Links from other worksheets
- **Yellow background**: Key assumptions or totals

## Common Formulas

### Summation

```javascript
// Simple sum
sheet.cell('B20').formula('=SUM(B2:B19)');

// Sum with condition
sheet.cell('B21').formula('=SUMIF(A2:A19,">1000",B2:B19)');

// Running total
sheet.cell('C2').formula('=B2');
sheet.cell('C3').formula('=C2+B3');  // Copy down
```

### Calculations

```javascript
// Growth rate
sheet.cell('D5').formula('=(C5-C4)/C4');
sheet.cell('D5').style({ numberFormat: '0.0%' });

// Average
sheet.cell('E10').formula('=AVERAGE(E2:E9)');

// Count
sheet.cell('F10').formula('=COUNTA(F2:F9)');

// Lookup
sheet.cell('G5').formula('=VLOOKUP(A5,Data!A:B,2,FALSE)');
```

### Financial Formulas

```javascript
// Net Present Value
sheet.cell('H15').formula('=NPV(0.1,H2:H14)');

// Payment calculation
sheet.cell('I10').formula('=PMT(0.05/12,360,-300000)');

// Future value
sheet.cell('J20').formula('=FV(0.08,10,0,-10000)');
```

## Formula Best Practices

### 1. Never Hardcode Calculated Values

```javascript
// ❌ Wrong
const total = calculateTotal();
sheet.cell('B10').value(total);

// ✅ Correct
sheet.cell('B10').formula('=SUM(B2:B9)');
```

### 2. Use Named Ranges for Clarity

```javascript
// Reference with clear cell labeling
// B2: Revenue
// B3: Cost
// B4: Profit (formula)
sheet.cell('B4').formula('=B2-B3');
```

### 3. Chain Formulas for Auditability

```javascript
// Step 1: Calculate subtotals
sheet.cell('D2').formula('=B2*C2');  // Line total
sheet.cell('D3').formula('=B3*C3');
// ...

// Step 2: Sum subtotals
sheet.cell('D20').formula('=SUM(D2:D19)');  // Subtotal

// Step 3: Add tax
sheet.cell('D21').formula('=D20*0.1');  // Tax

// Step 4: Final total
sheet.cell('D22').formula('=D20+D21');  // Grand total
```

### 4. Handle Errors Gracefully

```javascript
// Wrap formulas with IFERROR
sheet.cell('E5').formula('=IFERROR(D5/C5,"N/A")');

// Avoid division by zero
sheet.cell('F5').formula('=IF(C5=0,0,D5/C5)');
```

### 5. Cross-Worksheet References

```javascript
// Reference another sheet
sheet.cell('A1').formula('=Summary!A1');

// Sum from another sheet
sheet.cell('B10').formula('=SUM(Data!B2:B100)');
```
