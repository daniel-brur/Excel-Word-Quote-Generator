# Perito Quote Generator for Excel and Word

A Microsoft Excel VBA automation tool that generates a ready-to-deliver Microsoft Word quotation from information entered in an Excel workbook.

The project is designed for professional-service quotations, including expert-witness / forensic-engineering work, but the workbook can be adapted to other quotation workflows.

## What it does

The Excel workbook contains VBA macros that automate the preparation of a Word document:

- Reads quotation data from Excel cells.
- Calculates the subtotal, 16% VAT, and total amount.
- Calculates the total number of working hours from the activity table.
- Converts monetary amounts to Spanish words.
- Builds a payment statement using the calculated amounts.
- Creates a new Microsoft Word document.
- Copies selected text blocks from Excel to Word.
- Copies quotation tables from Excel to Word.
- Adds spacing between the copied sections.
- Adjusts table widths to fit the document.
- Applies paragraph alignment and bold formatting.
- Sets Spanish as the Word proofing language.
- Saves the generated document as a `.docx` file.
- Closes Word after saving.

## Project structure

```text
.
├── PeritoQuoteGenerator.xlsm
├── README.md
├── LICENSE
├── .gitignore
├── GITHUB_DESCRIPTION.txt
├── GITHUB_TOPICS.txt
└── docs/
    └── VBA_OVERVIEW.md
```

## Requirements

- Microsoft Excel for Windows with VBA support.
- Microsoft Word for Windows.
- Macros must be enabled in Excel.
- The VBA code uses Word automation through `CreateObject`, so a local Microsoft Word installation is required.
- The workbook was created for a Spanish-language quotation workflow. The generated payment text and number-to-words function are in Spanish.

## How to use

1. Download `PeritoQuoteGenerator.xlsm`.
2. Open it in Microsoft Excel for Windows.
3. Enable macros when prompted.
4. Enter or replace the quotation information in the worksheet.
5. Review the hourly rate, activities, hours, and calculated amounts.
6. Set the destination folder and quotation/document name according to the VBA configuration.
7. Run the `PasteTableInWord` macro.
8. Review the generated Word document before delivery.

## Current worksheet layout

The sample workbook uses the following main areas:

| Area | Purpose |
|---|---|
| `E3` | Quotation/document name |
| `E4` | Destination folder |
| `E5:E8` | Header and case information copied to Word |
| `E9:G15` | Main quotation table |
| `J1:L3` | Amount/payment summary |
| `E2` | Hourly professional fee |
| `F11:F17` | Hours by activity |
| `G11:G17` | Amount by activity |
| `F18:G18` | Total hours and subtotal |
| `L1` | Subtotal |
| `L2` | 16% VAT |
| `L3` | Total amount |

These ranges are part of the current VBA implementation and can be changed in the code when adapting the workbook.

## Number-to-words conversion

The workbook includes a VBA function named `NumLetras`. It converts numeric amounts into Spanish written amounts and handles the decimal/cents portion using the `00/100` format.

For example, an amount such as:

```text
30,209.08
```

can be incorporated into a Spanish payment statement as a written amount.

The wording is intended for quotation documents and should be reviewed before using the output in a legally binding or financial document.

## Customization

The main areas to customize are:

- Excel cell locations.
- Word document formatting.
- Output folder.
- Output file naming.
- Number-to-words wording.
- VAT rate.
- Currency wording.
- Quotation table ranges.
- Case-specific text.
- Professional fee calculation.

The VBA project currently contains the main Word-generation routine `PasteTableInWord` and the number-to-words functionality.

See `docs/VBA_OVERVIEW.md` for a higher-level description of the automation.

## Important security and privacy note

The GitHub version of the workbook is a sanitized sample. Personal names, banking information, case information, and local Windows paths have been replaced with placeholders.

Do **not** publish a production workbook containing:

- Bank account numbers.
- CLABE numbers.
- Personal addresses or telephone numbers.
- Client information.
- Confidential case information.
- Private file-system paths.
- Other personal or confidential data.

Always review the workbook before making a repository public.

## License

This project is released under the MIT License. See `LICENSE`.

## Disclaimer

This project is an Excel/Word automation example. It does not provide legal advice, accounting advice, or professional fee guidance. Users are responsible for reviewing the generated documents, calculations, legal references, and applicable tax or professional-fee requirements before delivery.
