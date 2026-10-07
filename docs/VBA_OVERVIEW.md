# VBA Automation Overview

## Main automation routine

The workbook contains a VBA procedure named `PasteTableInWord`.

Its workflow is approximately:

1. Read the destination folder from the worksheet.
2. Read the document/quotation name from the worksheet.
3. Build the payment statement using calculated Excel values.
4. Create a Microsoft Word application instance.
5. Create a new Word document.
6. Copy the configured text range from Excel.
7. Paste the text into Word.
8. Copy the main quotation table from Excel.
9. Paste the table into Word.
10. Add spacing before the second table.
11. Copy the payment-summary table.
12. Paste the second table into Word.
13. Adjust table widths.
14. Apply bold formatting to the first three lines.
15. Align the first two lines to the right.
16. Justify the fourth paragraph.
17. Set the Word proofing language to Spanish.
18. Save the document as a `.docx`.
19. Close the document and quit Word.

## Current Excel-to-Word ranges

The embedded VBA currently uses these ranges:

- Text/header section: `E5:E8`
- Main quotation table: `E9:G15`
- Payment summary table: `J1:L3`

If the worksheet layout changes, these ranges must also be updated in the VBA code.

## Amount calculation

The workbook calculates:

```text
Activity amount = Hours × Hourly fee
Subtotal = Sum of activity amounts
VAT = Subtotal × 16%
Total = Subtotal + VAT
```

The current workbook uses a 16% VAT rate.

## Number-to-words functionality

The VBA project includes a `NumLetras` function and supporting routines for converting numeric values into Spanish words.

The implementation contains Spanish number words for units, tens, hundreds, thousands, millions, and larger values. It also formats the decimal part as `00/100`.

## Microsoft Word dependency

The automation controls Word from Excel. The project uses Word automation rather than requiring a permanently open Word template.

Therefore:

- Excel and Word must be installed on the Windows computer.
- The user must allow VBA macros to run.
- Word automation may be affected by Microsoft Office security settings or local policies.

## Adaptation recommendations

When adapting the project for a different workflow:

1. Move user-editable values to clearly identified Excel cells.
2. Keep calculation cells separate from presentation cells.
3. Update the VBA ranges when the worksheet layout changes.
4. Validate the output filename and destination folder.
5. Add error handling for missing folders and unavailable Word installations.
6. Validate all financial calculations before delivery.
7. Remove confidential information before publishing the workbook.

## Source-code note

The VBA project remains embedded inside the `.xlsm` workbook. This repository package documents the embedded automation but does not claim that the VBA modules have been exported as standalone `.bas` files.
