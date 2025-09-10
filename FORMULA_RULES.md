# Excel MCP Formula Rules

## Core Principle: Transparency Through Formulas

This MCP server follows a strict rule: **ALL calculated values MUST use Excel formulas**, never hardcoded values. This ensures:

1. **Transparency**: Users can see exactly how values are calculated
2. **Auditability**: Every calculation can be traced and verified
3. **Professional Standards**: Mimics how a human would create the spreadsheet
4. **Auto-calculation**: Values update automatically when source data changes

## How It Works

### When You Open an Excel File
- **Cell View**: Shows the calculated value (e.g., "4")
- **Formula Bar**: Shows the formula when you click the cell (e.g., "=SUM(A1,A2)")
- **Auto-calculation**: Formulas are automatically calculated when the file opens

### Example: Adding Two Numbers

❌ **WRONG** (Hardcoded value):
```javascript
// This violates our rules - never do this!
worksheet.getCell('A3').value = 4; // Just the result
```

✅ **CORRECT** (Formula):
```javascript
// This is the right way - transparent calculation
worksheet.getCell('A3').formula = '=SUM(A1,A2)';
// or
worksheet.getCell('A3').formula = '=A1+A2';
```

## Available Tools for Formula-Based Calculations

### 1. `excel_write_calculation`
Write any calculation using a formula:

```json
{
  "worksheetName": "Sheet1",
  "cell": "A3",
  "operation": "sum",
  "references": ["A1", "A2"],
  "description": "Total of values above"
}
```

Supported operations:
- `sum`: =SUM(references)
- `average`: =AVERAGE(references)
- `count`: =COUNT(references)
- `min`: =MIN(references)
- `max`: =MAX(references)
- `product`: =PRODUCT(references)
- `subtract`: =A1-A2
- `divide`: =A1/A2
- `multiply`: =A1*A2
- `custom`: Any custom formula

### 2. `excel_validate_formulas`
Verify that cells contain formulas, not hardcoded values:

```json
{
  "worksheetName": "Sheet1",
  "cells": ["A3", "B5", "C10"]
}
```

### 3. `excel_ensure_formula_calculation`
Ensure all formulas in a worksheet are calculated when opened:

```json
{
  "worksheetName": "Sheet1"
}
```

## Common Formula Patterns

### Basic Arithmetic
```excel
=A1+B1          // Addition
=A1-B1          // Subtraction
=A1*B1          // Multiplication
=A1/B1          // Division
=A1^2           // Power
```

### Aggregation
```excel
=SUM(A1:A10)    // Sum a range
=AVERAGE(B1:B5) // Average
=COUNT(C1:C20)  // Count numbers
=MAX(D1:D10)    // Maximum value
=MIN(E1:E10)    // Minimum value
```

### Financial
```excel
=PMT(rate, periods, -principal)  // Loan payment
=NPV(rate, cashflows)            // Net present value
=FV(rate, periods, payment)      // Future value
=PV(rate, periods, payment)      // Present value
```

### Conditional
```excel
=IF(A1>10, "High", "Low")        // Simple if
=SUMIF(A1:A10, ">5")             // Conditional sum
=COUNTIF(B1:B10, "Yes")          // Conditional count
```

## Best Practices

1. **Always Use Formulas for Calculations**
   - Even simple additions should use =A1+A2, not hardcoded values
   - This ensures transparency and auditability

2. **Document Complex Formulas**
   - Use the description parameter to explain what the formula does
   - Add cell notes for complex calculations

3. **Use Named Ranges**
   - For frequently referenced cells, create named ranges
   - Makes formulas more readable: =SUM(Revenue) vs =SUM(B2:B100)

4. **Validate Formula Presence**
   - Use `excel_validate_formulas` to ensure calculations use formulas
   - Run validation after creating spreadsheets

5. **Enable Auto-calculation**
   - Use `excel_ensure_formula_calculation` to guarantee formulas calculate
   - Essential for professional spreadsheets

## Example: Creating a Professional Invoice

```javascript
// Header
['Item', 'Quantity', 'Unit Price', 'Total']

// Line items with formula for total
['Widget', 10, 25.00, '=B2*C2']  // Total = Quantity * Unit Price
['Gadget', 5, 50.00, '=B3*C3']
['Tool', 3, 100.00, '=B4*C4']

// Summary with formulas
['', '', 'Subtotal:', '=SUM(D2:D4)']     // Sum all totals
['', '', 'Tax (10%):', '=D5*0.1']        // 10% of subtotal
['', '', 'Grand Total:', '=D5+D6']       // Subtotal + Tax
```

## Compliance Standards

This approach aligns with:
- **GAAP** (Generally Accepted Accounting Principles)
- **IFRS** (International Financial Reporting Standards)
- **SOX** (Sarbanes-Oxley) audit requirements
- **Best practices** in financial modeling and analysis

## Remember

> "If it's a calculation, it MUST be a formula. No exceptions."

This ensures every spreadsheet created by the MCP server looks and behaves exactly as if a professional created it manually in Excel.