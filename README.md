# Google Sheets Data and Formatting Copier

This Google Apps Script copies data and all cell formatting (including number formats) from a source sheet in one Google Spreadsheet to a target sheet in another Google Spreadsheet. The script ensures that all cell contents, backgrounds, fonts, alignments, text styles, and number formats are preserved in the target sheet.

## Features

- Copies data from the source sheet to the target sheet.
- Preserves background colors, font colors, styles, sizes, and alignments.
- Copies text styles such as bold, italic, and underline.
- Copies number formats (e.g., currency, date, percentage, etc.).
- Clears the target sheet before copying to avoid any data conflicts.
- Sets text wrapping strategy in the target sheet.

## Prerequisites

- Two Google Sheets: one as the source and another as the target.
- Access to Google Apps Script for adding and executing the script.

## Usage

1. Open the Google Spreadsheet that you want to use as the source.
2. Click on **Extensions > Apps Script** to open the Google Apps Script editor.
3. Copy and paste the code provided below into the editor.
4. Update the `sourceSheetId` and `targetSheetId` with the unique IDs of your source and target Google Sheets, respectively.
5. Customize the `sourceSheetName` and `targetSheetName` variables with the names of the source and target sheets.
6. Save the script and run the `copySheetDataAndFormatting` function.

## Code Explanation

The `copySheetDataAndFormatting` function performs the following steps:

1. **Initialization**:
   - Define the IDs and names of the source and target sheets.
   - Open both the source and target spreadsheets using `SpreadsheetApp.openById`.

2. **Clear Target Sheet**:
   - Clear the target sheet to remove any existing data and formatting using `targetSheet.clear()`.

3. **Copy Data**:
   - Retrieve all data from the source sheet using `getDataRange().getValues()` and set it in the target sheet with `setValues()`.

4. **Copy Formatting**:
   - Copy various cell formats from the source sheet to the target sheet:
     - **Background Colors**: `getBackgrounds()` and `setBackgrounds()`
     - **Font Colors**: `getFontColors()` and `setFontColors()`
     - **Font Styles**: `getFontFamilies()` and `setFontFamilies()`
     - **Font Sizes**: `getFontSizes()` and `setFontSizes()`
     - **Horizontal and Vertical Alignments**: `getHorizontalAlignments()`, `getVerticalAlignments()`, `setHorizontalAlignments()`, and `setVerticalAlignments()`
     - **Text Styles**: `getTextStyles()` and `setTextStyles()`
     - **Number Formats**: `getNumberFormats()` and `setNumberFormats()`

5. **Set Text Wrapping Strategy**:
   - Apply a text wrapping strategy (e.g., CLIP) to the target range using `setWrapStrategy()`.
