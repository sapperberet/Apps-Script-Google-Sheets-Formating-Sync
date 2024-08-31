function copySheetDataAndFormatting() {
  var sourceSheetId = "copyAPI";
  var targetSheetId = "pasteAPI";
  
  var sourceSheetName = "Sheet1";
  var targetSheetName = "Sheet1";
  
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSheetId);
  var targetSpreadsheet = SpreadsheetApp.openById(targetSheetId);
  
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
  
  targetSheet.clear();
  
  var dataRange = sourceSheet.getDataRange();
  var data = dataRange.getValues(); 
  targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
  var backgrounds = dataRange.getBackgrounds(); 
  targetSheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds); 

  var fontColors = dataRange.getFontColors();
  targetSheet.getRange(1, 1, fontColors.length, fontColors[0].length).setFontColors(fontColors);

  var fontStyles = dataRange.getFontFamilies();
  targetSheet.getRange(1, 1, fontStyles.length, fontStyles[0].length).setFontFamilies(fontStyles);

  var fontSizes = dataRange.getFontSizes();
  targetSheet.getRange(1, 1, fontSizes.length, fontSizes[0].length).setFontSizes(fontSizes);

  var horizontalAlignments = dataRange.getHorizontalAlignments();
  targetSheet.getRange(1, 1, horizontalAlignments.length, horizontalAlignments[0].length).setHorizontalAlignments(horizontalAlignments);
  
  var verticalAlignments = dataRange.getVerticalAlignments();
  targetSheet.getRange(1, 1, verticalAlignments.length, verticalAlignments[0].length).setVerticalAlignments(verticalAlignments);

  var textStyles = dataRange.getTextStyles();
  targetSheet.getRange(1, 1, textStyles.length, textStyles[0].length).setTextStyles(textStyles);
  
  var numRows = data.length;
  var numCols = data[0].length;
  var range = targetSheet.getRange(1, 1, numRows, numCols);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // var mergedRanges = getMergedRanges(sourceSheet);
  // for (var i = 0; i < mergedRanges.length; i++) {
  //   var mergedRange = mergedRanges[i];
  //   var targetRange = targetSheet.getRange(mergedRange[0], mergedRange[1], mergedRange[2], mergedRange[3]);
  //   targetRange.merge();
  // }
}

// function getMergedRanges(sheet) {
//   var mergedRanges = [];
//   var range = sheet.getDataRange();
//   var numRows = range.getNumRows();
//   var numCols = range.getNumColumns();
  
//   for (var row = 1; row <= numRows; row++) {
//     for (var col = 1; col <= numCols; col++) {
//       var cell = sheet.getRange(row, col);
//       if (cell.isPartOfMerge()) {
//         var mergedArea = cell.getMergedRanges()[0];
//         var mergedRangeA1 = mergedArea.getA1Notation();
//         var mergedRange = sheet.getRange(mergedRangeA1);
//         var mergedRangeData = [
//           mergedRange.getRow(),
//           mergedRange.getColumn(),
//           mergedRange.getNumRows(),
//           mergedRange.getNumColumns()
//         ];
//         if (!mergedRanges.some(function(r) {
//           return r.join() === mergedRangeData.join();
//         })) {
//           mergedRanges.push(mergedRangeData);
//         }
//       }
//     }
//   }
  
//   return mergedRanges;
// }
