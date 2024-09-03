function copySheetDataAndFormatting() {

  //The Main Sheet identifier (Quality Indicators Sheet)
  //  https://docs.google.com/spreadsheets/d/abc (notice the ID )

  var targetSheetId = "abc"; 

  //===================================================

  // ALL the 15 input form sheets links

  var sourceSheetIdDAR = ""; // دار اسماعيل
  var sourceSheetIdAWA = ""; // صلاح العوضي
  var sourceSheetIdANF = ""; // اطفال الانفوشي
  var sourceSheetIdFAW = ""; // فوزي معاذ
  var sourceSheetIdKOM = ""; // صدر كوم الشقافة
  var sourceSheetIdRAM = ""; // أطفال الرمل
  var sourceSheetIdFEV = ""; // الحميات
  var sourceSheetIdOPT = ""; // الرمد
  var sourceSheetIdGAM = ""; // جمال حمادة 
  var sourceSheetIdBOR = ""; // برج العرب
  var sourceSheetIdAMR = ""; // العامرية العام
  var sourceSheetIdRAS = ""; // رأس التين العام
  var sourceSheetIdGOM = ""; // الجمهورية
  var sourceSheetIdABO = ""; // أبو قير العام
  var sourceSheetIdMAM = ""; // صدر المعمورة  

  

  //===================================================

  // Tabs names in the Main Sheet (if you changed the tab name remember to change it here)

  var targetSheetNameDAR = "";
  var targetSheetNameAWA = "";
  var targetSheetNameANF = "";
  var targetSheetNameFAW = "";
  var targetSheetNameKOM = "";
  var targetSheetNameRAM = "";
  var targetSheetNameFEV = "";
  var targetSheetNameOPT = "";
  var targetSheetNameGAM = "";
  var targetSheetNameBOR = "";
  var targetSheetNameAMR = "";
  var targetSheetNameRAS = "";
  var targetSheetNameGOM = "";
  var targetSheetNameABO = "";
  var targetSheetNameMAM = "";


  //===================================================

  // the tab name  in all form sheets (set by default to Sheet1)

  var sourceSheetName = "Sheet1";

  //===================================================

  // uses the id to access the Spreadsheet App for the Main Sheet 

  var targetSpreadsheet = SpreadsheetApp.openById(targetSheetId);



  //===================================================

  //  uses the id to access the Spreadsheet App for the Other sheets 

  var sourceSpreadsheetDAR = SpreadsheetApp.openById(sourceSheetIdDAR);
  var sourceSpreadsheetAWA = SpreadsheetApp.openById(sourceSheetIdAWA);
  var sourceSpreadsheetANF = SpreadsheetApp.openById(sourceSheetIdANF);
  var sourceSpreadsheetFAW = SpreadsheetApp.openById(sourceSheetIdFAW);
  var sourceSpreadsheetKOM = SpreadsheetApp.openById(sourceSheetIdKOM);
  var sourceSpreadsheetRAM = SpreadsheetApp.openById(sourceSheetIdRAM);
  var sourceSpreadsheetFEV = SpreadsheetApp.openById(sourceSheetIdFEV);
  var sourceSpreadsheetOPT = SpreadsheetApp.openById(sourceSheetIdOPT);
  var sourceSpreadsheetGAM = SpreadsheetApp.openById(sourceSheetIdGAM);
  var sourceSpreadsheetBOR = SpreadsheetApp.openById(sourceSheetIdBOR);
  var sourceSpreadsheetAMR = SpreadsheetApp.openById(sourceSheetIdAMR);
  var sourceSpreadsheetRAS = SpreadsheetApp.openById(sourceSheetIdRAS);
  var sourceSpreadsheetGOM = SpreadsheetApp.openById(sourceSheetIdGOM);
  var sourceSpreadsheetABO = SpreadsheetApp.openById(sourceSheetIdABO);
  var sourceSpreadsheetMAM = SpreadsheetApp.openById(sourceSheetIdMAM);

  //===================================================
  
  // gets tab name within that id

  var targetSheetDAR = targetSpreadsheet.getSheetByName(targetSheetNameDAR);
  var targetSheetAWA = targetSpreadsheet.getSheetByName(targetSheetNameAWA);
  var targetSheetANF = targetSpreadsheet.getSheetByName(targetSheetNameANF);
  var targetSheetFAW = targetSpreadsheet.getSheetByName(targetSheetNameFAW);
  var targetSheetKOM = targetSpreadsheet.getSheetByName(targetSheetNameKOM);
  var targetSheetRAM = targetSpreadsheet.getSheetByName(targetSheetNameRAM);
  var targetSheetFEV = targetSpreadsheet.getSheetByName(targetSheetNameFEV);
  var targetSheetOPT = targetSpreadsheet.getSheetByName(targetSheetNameOPT);
  var targetSheetGAM = targetSpreadsheet.getSheetByName(targetSheetNameGAM);
  var targetSheetBOR = targetSpreadsheet.getSheetByName(targetSheetNameBOR);
  var targetSheetAMR = targetSpreadsheet.getSheetByName(targetSheetNameAMR);
  var targetSheetRAS = targetSpreadsheet.getSheetByName(targetSheetNameRAS);
  var targetSheetGOM = targetSpreadsheet.getSheetByName(targetSheetNameGOM);
  var targetSheetABO = targetSpreadsheet.getSheetByName(targetSheetNameABO);
  var targetSheetMAM = targetSpreadsheet.getSheetByName(targetSheetNameMAM);

  //===================================================

  // gets tab name within that id

  var sourceSheetDAR = sourceSpreadsheetDAR.getSheetByName(sourceSheetName);
  var sourceSheetAWA = sourceSpreadsheetAWA.getSheetByName(sourceSheetName);
  var sourceSheetANF = sourceSpreadsheetANF.getSheetByName(sourceSheetName);
  var sourceSheetFAW = sourceSpreadsheetFAW.getSheetByName(sourceSheetName);
  var sourceSheetKOM = sourceSpreadsheetKOM.getSheetByName(sourceSheetName);
  var sourceSheetRAM = sourceSpreadsheetRAM.getSheetByName(sourceSheetName);
  var sourceSheetFEV = sourceSpreadsheetFEV.getSheetByName(sourceSheetName);
  var sourceSheetOPT = sourceSpreadsheetOPT.getSheetByName(sourceSheetName);
  var sourceSheetGAM = sourceSpreadsheetGAM.getSheetByName(sourceSheetName);
  var sourceSheetBOR = sourceSpreadsheetBOR.getSheetByName(sourceSheetName);
  var sourceSheetAMR = sourceSpreadsheetAMR.getSheetByName(sourceSheetName);
  var sourceSheetRAS = sourceSpreadsheetRAS.getSheetByName(sourceSheetName);
  var sourceSheetGOM = sourceSpreadsheetGOM.getSheetByName(sourceSheetName);
  var sourceSheetABO = sourceSpreadsheetABO.getSheetByName(sourceSheetName);
  var sourceSheetMAM = sourceSpreadsheetMAM.getSheetByName(sourceSheetName);

  //===================================================
  
  //refreshes the the main sheet to apply the changes 

  targetSheetDAR.clear();
  targetSheetAWA.clear();
  targetSheetANF.clear();
  targetSheetFAW.clear();
  targetSheetKOM.clear();
  targetSheetRAM.clear();
  targetSheetFEV.clear();
  targetSheetOPT.clear();
  targetSheetGAM.clear();
  targetSheetBOR.clear();
  targetSheetAMR.clear();
  targetSheetRAS.clear();
  targetSheetGOM.clear();
  targetSheetABO.clear();
  targetSheetMAM.clear();

  //=================================================== 
  
  // The rest of the code is what becoming copied :
     //Background , Text format , Text Style , Number format , Font color ... etc


  //===================================================

  var dataRangeDAR = sourceSheetDAR.getDataRange();
  var dataRangeAWA = sourceSheetAWA.getDataRange();
  var dataRangeANF = sourceSheetANF.getDataRange();
  var dataRangeFAW = sourceSheetFAW.getDataRange();
  var dataRangeKOM = sourceSheetKOM.getDataRange();
  var dataRangeRAM = sourceSheetRAM.getDataRange();
  var dataRangeFEV = sourceSheetFEV.getDataRange();
  var dataRangeOPT = sourceSheetOPT.getDataRange();
  var dataRangeGAM = sourceSheetGAM.getDataRange();
  var dataRangeBOR = sourceSheetBOR.getDataRange();
  var dataRangeAMR = sourceSheetAMR.getDataRange();
  var dataRangeRAS = sourceSheetRAS.getDataRange();
  var dataRangeGOM = sourceSheetGOM.getDataRange();
  var dataRangeABO = sourceSheetABO.getDataRange();
  var dataRangeMAM = sourceSheetMAM.getDataRange();

  //===================================================
  var dataDAR = dataRangeDAR.getValues();
  var dataAWA = dataRangeAWA.getValues();
  var dataANF = dataRangeANF.getValues();
  var dataFAW = dataRangeFAW.getValues();
  var dataKOM = dataRangeKOM.getValues();
  var dataRAM = dataRangeRAM.getValues();
  var dataFEV = dataRangeFEV.getValues();
  var dataOPT = dataRangeOPT.getValues();
  var dataGAM = dataRangeGAM.getValues();
  var dataBOR = dataRangeBOR.getValues();
  var dataAMR = dataRangeAMR.getValues();
  var dataRAS = dataRangeRAS.getValues();
  var dataGOM = dataRangeGOM.getValues();
  var dataABO = dataRangeABO.getValues();
  var dataMAM = dataRangeMAM.getValues();
  //===================================================

  targetSheetDAR.getRange(1, 1, dataDAR.length, dataDAR[0].length).setValues(dataDAR);
  targetSheetAWA.getRange(1, 1, dataAWA.length, dataAWA[0].length).setValues(dataAWA);
  targetSheetANF.getRange(1, 1, dataANF.length, dataANF[0].length).setValues(dataANF);
  targetSheetFAW.getRange(1, 1, dataFAW.length, dataFAW[0].length).setValues(dataFAW);
  targetSheetKOM.getRange(1, 1, dataKOM.length, dataKOM[0].length).setValues(dataKOM);
  targetSheetRAM.getRange(1, 1, dataRAM.length, dataRAM[0].length).setValues(dataRAM);
  targetSheetFEV.getRange(1, 1, dataFEV.length, dataFEV[0].length).setValues(dataFEV);
  targetSheetOPT.getRange(1, 1, dataOPT.length, dataOPT[0].length).setValues(dataOPT);
  targetSheetGAM.getRange(1, 1, dataGAM.length, dataGAM[0].length).setValues(dataGAM);
  targetSheetBOR.getRange(1, 1, dataBOR.length, dataBOR[0].length).setValues(dataBOR);
  targetSheetAMR.getRange(1, 1, dataAMR.length, dataAMR[0].length).setValues(dataAMR);
  targetSheetRAS.getRange(1, 1, dataRAS.length, dataRAS[0].length).setValues(dataRAS);
  targetSheetGOM.getRange(1, 1, dataGOM.length, dataGOM[0].length).setValues(dataGOM);
  targetSheetABO.getRange(1, 1, dataABO.length, dataABO[0].length).setValues(dataABO);
  targetSheetMAM.getRange(1, 1, dataMAM.length, dataMAM[0].length).setValues(dataMAM);

  //===================================================

  var backgroundsDAR = dataRangeDAR.getBackgrounds();
  var backgroundsAWA = dataRangeAWA.getBackgrounds();
  var backgroundsANF = dataRangeANF.getBackgrounds();
  var backgroundsFAW = dataRangeFAW.getBackgrounds();
  var backgroundsKOM = dataRangeKOM.getBackgrounds();
  var backgroundsRAM = dataRangeRAM.getBackgrounds();
  var backgroundsFEV = dataRangeFEV.getBackgrounds();
  var backgroundsOPT = dataRangeOPT.getBackgrounds();
  var backgroundsGAM = dataRangeGAM.getBackgrounds();
  var backgroundsBOR = dataRangeBOR.getBackgrounds();
  var backgroundsAMR = dataRangeAMR.getBackgrounds();
  var backgroundsRAS = dataRangeRAS.getBackgrounds();
  var backgroundsGOM = dataRangeGOM.getBackgrounds();
  var backgroundsABO = dataRangeABO.getBackgrounds();
  var backgroundsMAM = dataRangeMAM.getBackgrounds();


  //===================================================


  targetSheetDAR.getRange(1, 1, backgroundsDAR.length, backgroundsDAR[0].length).setBackgrounds(backgroundsDAR);
  targetSheetAWA.getRange(1, 1, backgroundsAWA.length, backgroundsAWA[0].length).setBackgrounds(backgroundsAWA);
  targetSheetANF.getRange(1, 1, backgroundsANF.length, backgroundsANF[0].length).setBackgrounds(backgroundsANF);
  targetSheetFAW.getRange(1, 1, backgroundsFAW.length, backgroundsFAW[0].length).setBackgrounds(backgroundsFAW);
  targetSheetKOM.getRange(1, 1, backgroundsKOM.length, backgroundsKOM[0].length).setBackgrounds(backgroundsKOM);
  targetSheetRAM.getRange(1, 1, backgroundsRAM.length, backgroundsRAM[0].length).setBackgrounds(backgroundsRAM);
  targetSheetFEV.getRange(1, 1, backgroundsFEV.length, backgroundsFEV[0].length).setBackgrounds(backgroundsFEV);
  targetSheetOPT.getRange(1, 1, backgroundsOPT.length, backgroundsOPT[0].length).setBackgrounds(backgroundsOPT);
  targetSheetGAM.getRange(1, 1, backgroundsGAM.length, backgroundsGAM[0].length).setBackgrounds(backgroundsGAM);
  targetSheetBOR.getRange(1, 1, backgroundsBOR.length, backgroundsBOR[0].length).setBackgrounds(backgroundsBOR);
  targetSheetAMR.getRange(1, 1, backgroundsAMR.length, backgroundsAMR[0].length).setBackgrounds(backgroundsAMR);
  targetSheetRAS.getRange(1, 1, backgroundsRAS.length, backgroundsRAS[0].length).setBackgrounds(backgroundsRAS);
  targetSheetGOM.getRange(1, 1, backgroundsGOM.length, backgroundsGOM[0].length).setBackgrounds(backgroundsGOM);
  targetSheetABO.getRange(1, 1, backgroundsABO.length, backgroundsABO[0].length).setBackgrounds(backgroundsABO);
  targetSheetMAM.getRange(1, 1, backgroundsMAM.length, backgroundsMAM[0].length).setBackgrounds(backgroundsMAM);


  //===================================================

  var fontColorsDAR = dataRangeDAR.getFontColors();
  var fontColorsAWA = dataRangeAWA.getFontColors();
  var fontColorsANF = dataRangeANF.getFontColors();
  var fontColorsFAW = dataRangeFAW.getFontColors();
  var fontColorsKOM = dataRangeKOM.getFontColors();
  var fontColorsRAM = dataRangeRAM.getFontColors();
  var fontColorsFEV = dataRangeFEV.getFontColors();
  var fontColorsOPT = dataRangeOPT.getFontColors();
  var fontColorsGAM = dataRangeGAM.getFontColors();
  var fontColorsBOR = dataRangeBOR.getFontColors();
  var fontColorsAMR = dataRangeAMR.getFontColors();
  var fontColorsRAS = dataRangeRAS.getFontColors();
  var fontColorsGOM = dataRangeGOM.getFontColors();
  var fontColorsABO = dataRangeABO.getFontColors();
  var fontColorsMAM = dataRangeMAM.getFontColors();

  //===================================================
  targetSheetDAR.getRange(1, 1, fontColorsDAR.length, fontColorsDAR[0].length).setFontColors(fontColorsDAR);
  targetSheetAWA.getRange(1, 1, fontColorsAWA.length, fontColorsAWA[0].length).setFontColors(fontColorsAWA);
  targetSheetANF.getRange(1, 1, fontColorsANF.length, fontColorsANF[0].length).setFontColors(fontColorsANF);
  targetSheetFAW.getRange(1, 1, fontColorsFAW.length, fontColorsFAW[0].length).setFontColors(fontColorsFAW);
  targetSheetKOM.getRange(1, 1, fontColorsKOM.length, fontColorsKOM[0].length).setFontColors(fontColorsKOM);
  targetSheetRAM.getRange(1, 1, fontColorsRAM.length, fontColorsRAM[0].length).setFontColors(fontColorsRAM);
  targetSheetFEV.getRange(1, 1, fontColorsFEV.length, fontColorsFEV[0].length).setFontColors(fontColorsFEV);
  targetSheetOPT.getRange(1, 1, fontColorsOPT.length, fontColorsOPT[0].length).setFontColors(fontColorsOPT);
  targetSheetGAM.getRange(1, 1, fontColorsGAM.length, fontColorsGAM[0].length).setFontColors(fontColorsGAM);
  targetSheetBOR.getRange(1, 1, fontColorsBOR.length, fontColorsBOR[0].length).setFontColors(fontColorsBOR);
  targetSheetAMR.getRange(1, 1, fontColorsAMR.length, fontColorsAMR[0].length).setFontColors(fontColorsAMR);
  targetSheetRAS.getRange(1, 1, fontColorsRAS.length, fontColorsRAS[0].length).setFontColors(fontColorsRAS);
  targetSheetGOM.getRange(1, 1, fontColorsGOM.length, fontColorsGOM[0].length).setFontColors(fontColorsGOM);
  targetSheetABO.getRange(1, 1, fontColorsABO.length, fontColorsABO[0].length).setFontColors(fontColorsABO);
  targetSheetMAM.getRange(1, 1, fontColorsMAM.length, fontColorsMAM[0].length).setFontColors(fontColorsMAM);

  //===================================================

  var fontFamiliesDAR = dataRangeDAR.getFontFamilies();
  var fontFamiliesAWA = dataRangeAWA.getFontFamilies();
  var fontFamiliesANF = dataRangeANF.getFontFamilies();
  var fontFamiliesFAW = dataRangeFAW.getFontFamilies();
  var fontFamiliesKOM = dataRangeKOM.getFontFamilies();
  var fontFamiliesRAM = dataRangeRAM.getFontFamilies();
  var fontFamiliesFEV = dataRangeFEV.getFontFamilies();
  var fontFamiliesOPT = dataRangeOPT.getFontFamilies();
  var fontFamiliesGAM = dataRangeGAM.getFontFamilies();
  var fontFamiliesBOR = dataRangeBOR.getFontFamilies();
  var fontFamiliesAMR = dataRangeAMR.getFontFamilies();
  var fontFamiliesRAS = dataRangeRAS.getFontFamilies();
  var fontFamiliesGOM = dataRangeGOM.getFontFamilies();
  var fontFamiliesABO = dataRangeABO.getFontFamilies();
  var fontFamiliesMAM = dataRangeMAM.getFontFamilies();

  //===================================================
  targetSheetDAR.getRange(1, 1, fontFamiliesDAR.length, fontFamiliesDAR[0].length).setFontFamilies(fontFamiliesDAR);
  targetSheetAWA.getRange(1, 1, fontFamiliesAWA.length, fontFamiliesAWA[0].length).setFontFamilies(fontFamiliesAWA);
  targetSheetANF.getRange(1, 1, fontFamiliesANF.length, fontFamiliesANF[0].length).setFontFamilies(fontFamiliesANF);
  targetSheetFAW.getRange(1, 1, fontFamiliesFAW.length, fontFamiliesFAW[0].length).setFontFamilies(fontFamiliesFAW);
  targetSheetKOM.getRange(1, 1, fontFamiliesKOM.length, fontFamiliesKOM[0].length).setFontFamilies(fontFamiliesKOM);
  targetSheetRAM.getRange(1, 1, fontFamiliesRAM.length, fontFamiliesRAM[0].length).setFontFamilies(fontFamiliesRAM);
  targetSheetFEV.getRange(1, 1, fontFamiliesFEV.length, fontFamiliesFEV[0].length).setFontFamilies(fontFamiliesFEV);
  targetSheetOPT.getRange(1, 1, fontFamiliesOPT.length, fontFamiliesOPT[0].length).setFontFamilies(fontFamiliesOPT);
  targetSheetGAM.getRange(1, 1, fontFamiliesGAM.length, fontFamiliesGAM[0].length).setFontFamilies(fontFamiliesGAM);
  targetSheetBOR.getRange(1, 1, fontFamiliesBOR.length, fontFamiliesBOR[0].length).setFontFamilies(fontFamiliesBOR);
  targetSheetAMR.getRange(1, 1, fontFamiliesAMR.length, fontFamiliesAMR[0].length).setFontFamilies(fontFamiliesAMR);
  targetSheetRAS.getRange(1, 1, fontFamiliesRAS.length, fontFamiliesRAS[0].length).setFontFamilies(fontFamiliesRAS);
  targetSheetGOM.getRange(1, 1, fontFamiliesGOM.length, fontFamiliesGOM[0].length).setFontFamilies(fontFamiliesGOM);
  targetSheetABO.getRange(1, 1, fontFamiliesABO.length, fontFamiliesABO[0].length).setFontFamilies(fontFamiliesABO);
  targetSheetMAM.getRange(1, 1, fontFamiliesMAM.length, fontFamiliesMAM[0].length).setFontFamilies(fontFamiliesMAM);

  //===================================================

  var fontSizesDAR = dataRangeDAR.getFontSizes();
  var fontSizesAWA = dataRangeAWA.getFontSizes();
  var fontSizesANF = dataRangeANF.getFontSizes();
  var fontSizesFAW = dataRangeFAW.getFontSizes();
  var fontSizesKOM = dataRangeKOM.getFontSizes();
  var fontSizesRAM = dataRangeRAM.getFontSizes();
  var fontSizesFEV = dataRangeFEV.getFontSizes();
  var fontSizesOPT = dataRangeOPT.getFontSizes();
  var fontSizesGAM = dataRangeGAM.getFontSizes();
  var fontSizesBOR = dataRangeBOR.getFontSizes();
  var fontSizesAMR = dataRangeAMR.getFontSizes();
  var fontSizesRAS = dataRangeRAS.getFontSizes();
  var fontSizesGOM = dataRangeGOM.getFontSizes();
  var fontSizesABO = dataRangeABO.getFontSizes();
  var fontSizesMAM = dataRangeMAM.getFontSizes();

  //===================================================
  targetSheetDAR.getRange(1, 1, fontSizesDAR.length, fontSizesDAR[0].length).setFontSizes(fontSizesDAR);
  targetSheetAWA.getRange(1, 1, fontSizesAWA.length, fontSizesAWA[0].length).setFontSizes(fontSizesAWA);
  targetSheetANF.getRange(1, 1, fontSizesANF.length, fontSizesANF[0].length).setFontSizes(fontSizesANF);
  targetSheetFAW.getRange(1, 1, fontSizesFAW.length, fontSizesFAW[0].length).setFontSizes(fontSizesFAW);
  targetSheetKOM.getRange(1, 1, fontSizesKOM.length, fontSizesKOM[0].length).setFontSizes(fontSizesKOM);
  targetSheetRAM.getRange(1, 1, fontSizesRAM.length, fontSizesRAM[0].length).setFontSizes(fontSizesRAM);
  targetSheetFEV.getRange(1, 1, fontSizesFEV.length, fontSizesFEV[0].length).setFontSizes(fontSizesFEV);
  targetSheetOPT.getRange(1, 1, fontSizesOPT.length, fontSizesOPT[0].length).setFontSizes(fontSizesOPT);
  targetSheetGAM.getRange(1, 1, fontSizesGAM.length, fontSizesGAM[0].length).setFontSizes(fontSizesGAM);
  targetSheetBOR.getRange(1, 1, fontSizesBOR.length, fontSizesBOR[0].length).setFontSizes(fontSizesBOR);
  targetSheetAMR.getRange(1, 1, fontSizesAMR.length, fontSizesAMR[0].length).setFontSizes(fontSizesAMR);
  targetSheetRAS.getRange(1, 1, fontSizesRAS.length, fontSizesRAS[0].length).setFontSizes(fontSizesRAS);
  targetSheetGOM.getRange(1, 1, fontSizesGOM.length, fontSizesGOM[0].length).setFontSizes(fontSizesGOM);
  targetSheetABO.getRange(1, 1, fontSizesABO.length, fontSizesABO[0].length).setFontSizes(fontSizesABO);
  targetSheetMAM.getRange(1, 1, fontSizesMAM.length, fontSizesMAM[0].length).setFontSizes(fontSizesMAM);


  //===================================================

  var horizontalAlignmentsDAR = dataRangeDAR.getHorizontalAlignments();
  var horizontalAlignmentsAWA = dataRangeAWA.getHorizontalAlignments();
  var horizontalAlignmentsANF = dataRangeANF.getHorizontalAlignments();
  var horizontalAlignmentsFAW = dataRangeFAW.getHorizontalAlignments();
  var horizontalAlignmentsKOM = dataRangeKOM.getHorizontalAlignments();
  var horizontalAlignmentsRAM = dataRangeRAM.getHorizontalAlignments();
  var horizontalAlignmentsFEV = dataRangeFEV.getHorizontalAlignments();
  var horizontalAlignmentsOPT = dataRangeOPT.getHorizontalAlignments();
  var horizontalAlignmentsGAM = dataRangeGAM.getHorizontalAlignments();
  var horizontalAlignmentsBOR = dataRangeBOR.getHorizontalAlignments();
  var horizontalAlignmentsAMR = dataRangeAMR.getHorizontalAlignments();
  var horizontalAlignmentsRAS = dataRangeRAS.getHorizontalAlignments();
  var horizontalAlignmentsGOM = dataRangeGOM.getHorizontalAlignments();
  var horizontalAlignmentsABO = dataRangeABO.getHorizontalAlignments();
  var horizontalAlignmentsMAM = dataRangeMAM.getHorizontalAlignments();

  //===================================================
  targetSheetDAR.getRange(1, 1, horizontalAlignmentsDAR.length, horizontalAlignmentsDAR[0].length).setHorizontalAlignments(horizontalAlignmentsDAR);
  targetSheetAWA.getRange(1, 1, horizontalAlignmentsAWA.length, horizontalAlignmentsAWA[0].length).setHorizontalAlignments(horizontalAlignmentsAWA);
  targetSheetANF.getRange(1, 1, horizontalAlignmentsANF.length, horizontalAlignmentsANF[0].length).setHorizontalAlignments(horizontalAlignmentsANF);
  targetSheetFAW.getRange(1, 1, horizontalAlignmentsFAW.length, horizontalAlignmentsFAW[0].length).setHorizontalAlignments(horizontalAlignmentsFAW);
  targetSheetKOM.getRange(1, 1, horizontalAlignmentsKOM.length, horizontalAlignmentsKOM[0].length).setHorizontalAlignments(horizontalAlignmentsKOM);
  targetSheetRAM.getRange(1, 1, horizontalAlignmentsRAM.length, horizontalAlignmentsRAM[0].length).setHorizontalAlignments(horizontalAlignmentsRAM);
  targetSheetFEV.getRange(1, 1, horizontalAlignmentsFEV.length, horizontalAlignmentsFEV[0].length).setHorizontalAlignments(horizontalAlignmentsFEV);
  targetSheetOPT.getRange(1, 1, horizontalAlignmentsOPT.length, horizontalAlignmentsOPT[0].length).setHorizontalAlignments(horizontalAlignmentsOPT);
  targetSheetGAM.getRange(1, 1, horizontalAlignmentsGAM.length, horizontalAlignmentsGAM[0].length).setHorizontalAlignments(horizontalAlignmentsGAM);
  targetSheetBOR.getRange(1, 1, horizontalAlignmentsBOR.length, horizontalAlignmentsBOR[0].length).setHorizontalAlignments(horizontalAlignmentsBOR);
  targetSheetAMR.getRange(1, 1, horizontalAlignmentsAMR.length, horizontalAlignmentsAMR[0].length).setHorizontalAlignments(horizontalAlignmentsAMR);
  targetSheetRAS.getRange(1, 1, horizontalAlignmentsRAS.length, horizontalAlignmentsRAS[0].length).setHorizontalAlignments(horizontalAlignmentsRAS);
  targetSheetGOM.getRange(1, 1, horizontalAlignmentsGOM.length, horizontalAlignmentsGOM[0].length).setHorizontalAlignments(horizontalAlignmentsGOM);
  targetSheetABO.getRange(1, 1, horizontalAlignmentsABO.length, horizontalAlignmentsABO[0].length).setHorizontalAlignments(horizontalAlignmentsABO);
  targetSheetMAM.getRange(1, 1, horizontalAlignmentsMAM.length, horizontalAlignmentsMAM[0].length).setHorizontalAlignments(horizontalAlignmentsMAM);


  //===================================================

  var verticalAlignmentsDAR = dataRangeDAR.getVerticalAlignments();
  var verticalAlignmentsAWA = dataRangeAWA.getVerticalAlignments();
  var verticalAlignmentsANF = dataRangeANF.getVerticalAlignments();
  var verticalAlignmentsFAW = dataRangeFAW.getVerticalAlignments();
  var verticalAlignmentsKOM = dataRangeKOM.getVerticalAlignments();
  var verticalAlignmentsRAM = dataRangeRAM.getVerticalAlignments();
  var verticalAlignmentsFEV = dataRangeFEV.getVerticalAlignments();
  var verticalAlignmentsOPT = dataRangeOPT.getVerticalAlignments();
  var verticalAlignmentsGAM = dataRangeGAM.getVerticalAlignments();
  var verticalAlignmentsBOR = dataRangeBOR.getVerticalAlignments();
  var verticalAlignmentsAMR = dataRangeAMR.getVerticalAlignments();
  var verticalAlignmentsRAS = dataRangeRAS.getVerticalAlignments();
  var verticalAlignmentsGOM = dataRangeGOM.getVerticalAlignments();
  var verticalAlignmentsABO = dataRangeABO.getVerticalAlignments();
  var verticalAlignmentsMAM = dataRangeMAM.getVerticalAlignments();

  //===================================================
  targetSheetDAR.getRange(1, 1, verticalAlignmentsDAR.length, verticalAlignmentsDAR[0].length).setVerticalAlignments(verticalAlignmentsDAR);
  targetSheetAWA.getRange(1, 1, verticalAlignmentsAWA.length, verticalAlignmentsAWA[0].length).setVerticalAlignments(verticalAlignmentsAWA);
  targetSheetANF.getRange(1, 1, verticalAlignmentsANF.length, verticalAlignmentsANF[0].length).setVerticalAlignments(verticalAlignmentsANF);
  targetSheetFAW.getRange(1, 1, verticalAlignmentsFAW.length, verticalAlignmentsFAW[0].length).setVerticalAlignments(verticalAlignmentsFAW);
  targetSheetKOM.getRange(1, 1, verticalAlignmentsKOM.length, verticalAlignmentsKOM[0].length).setVerticalAlignments(verticalAlignmentsKOM);
  targetSheetRAM.getRange(1, 1, verticalAlignmentsRAM.length, verticalAlignmentsRAM[0].length).setVerticalAlignments(verticalAlignmentsRAM);
  targetSheetFEV.getRange(1, 1, verticalAlignmentsFEV.length, verticalAlignmentsFEV[0].length).setVerticalAlignments(verticalAlignmentsFEV);
  targetSheetOPT.getRange(1, 1, verticalAlignmentsOPT.length, verticalAlignmentsOPT[0].length).setVerticalAlignments(verticalAlignmentsOPT);
  targetSheetGAM.getRange(1, 1, verticalAlignmentsGAM.length, verticalAlignmentsGAM[0].length).setVerticalAlignments(verticalAlignmentsGAM);
  targetSheetBOR.getRange(1, 1, verticalAlignmentsBOR.length, verticalAlignmentsBOR[0].length).setVerticalAlignments(verticalAlignmentsBOR);
  targetSheetAMR.getRange(1, 1, verticalAlignmentsAMR.length, verticalAlignmentsAMR[0].length).setVerticalAlignments(verticalAlignmentsAMR);
  targetSheetRAS.getRange(1, 1, verticalAlignmentsRAS.length, verticalAlignmentsRAS[0].length).setVerticalAlignments(verticalAlignmentsRAS);
  targetSheetGOM.getRange(1, 1, verticalAlignmentsGOM.length, verticalAlignmentsGOM[0].length).setVerticalAlignments(verticalAlignmentsGOM);
  targetSheetABO.getRange(1, 1, verticalAlignmentsABO.length, verticalAlignmentsABO[0].length).setVerticalAlignments(verticalAlignmentsABO);
  targetSheetMAM.getRange(1, 1, verticalAlignmentsMAM.length, verticalAlignmentsMAM[0].length).setVerticalAlignments(verticalAlignmentsMAM);



  //===================================================

  var textStylesDAR = dataRangeDAR.getTextStyles();
  var textStylesAWA = dataRangeAWA.getTextStyles();
  var textStylesANF = dataRangeANF.getTextStyles();
  var textStylesFAW = dataRangeFAW.getTextStyles();
  var textStylesKOM = dataRangeKOM.getTextStyles();
  var textStylesRAM = dataRangeRAM.getTextStyles();
  var textStylesFEV = dataRangeFEV.getTextStyles();
  var textStylesOPT = dataRangeOPT.getTextStyles();
  var textStylesGAM = dataRangeGAM.getTextStyles();
  var textStylesBOR = dataRangeBOR.getTextStyles();
  var textStylesAMR = dataRangeAMR.getTextStyles();
  var textStylesRAS = dataRangeRAS.getTextStyles();
  var textStylesGOM = dataRangeGOM.getTextStyles();
  var textStylesABO = dataRangeABO.getTextStyles();
  var textStylesMAM = dataRangeMAM.getTextStyles();

  //===================================================

  targetSheetDAR.getRange(1, 1, textStylesDAR.length, textStylesDAR[0].length).setTextStyles(textStylesDAR);
  targetSheetAWA.getRange(1, 1, textStylesAWA.length, textStylesAWA[0].length).setTextStyles(textStylesAWA);
  targetSheetANF.getRange(1, 1, textStylesANF.length, textStylesANF[0].length).setTextStyles(textStylesANF);
  targetSheetFAW.getRange(1, 1, textStylesFAW.length, textStylesFAW[0].length).setTextStyles(textStylesFAW);
  targetSheetKOM.getRange(1, 1, textStylesKOM.length, textStylesKOM[0].length).setTextStyles(textStylesKOM);
  targetSheetRAM.getRange(1, 1, textStylesRAM.length, textStylesRAM[0].length).setTextStyles(textStylesRAM);
  targetSheetFEV.getRange(1, 1, textStylesFEV.length, textStylesFEV[0].length).setTextStyles(textStylesFEV);
  targetSheetOPT.getRange(1, 1, textStylesOPT.length, textStylesOPT[0].length).setTextStyles(textStylesOPT);
  targetSheetGAM.getRange(1, 1, textStylesGAM.length, textStylesGAM[0].length).setTextStyles(textStylesGAM);
  targetSheetBOR.getRange(1, 1, textStylesBOR.length, textStylesBOR[0].length).setTextStyles(textStylesBOR);
  targetSheetAMR.getRange(1, 1, textStylesAMR.length, textStylesAMR[0].length).setTextStyles(textStylesAMR);
  targetSheetRAS.getRange(1, 1, textStylesRAS.length, textStylesRAS[0].length).setTextStyles(textStylesRAS);
  targetSheetGOM.getRange(1, 1, textStylesGOM.length, textStylesGOM[0].length).setTextStyles(textStylesGOM);
  targetSheetABO.getRange(1, 1, textStylesABO.length, textStylesABO[0].length).setTextStyles(textStylesABO);
  targetSheetMAM.getRange(1, 1, textStylesMAM.length, textStylesMAM[0].length).setTextStyles(textStylesMAM);
  //===================================================


  var numberFormatsDAR = dataRangeDAR.getNumberFormats();
  var numberFormatsAWA = dataRangeAWA.getNumberFormats();
  var numberFormatsANF = dataRangeANF.getNumberFormats();
  var numberFormatsFAW = dataRangeFAW.getNumberFormats();
  var numberFormatsKOM = dataRangeKOM.getNumberFormats();
  var numberFormatsRAM = dataRangeRAM.getNumberFormats();
  var numberFormatsFEV = dataRangeFEV.getNumberFormats();
  var numberFormatsOPT = dataRangeOPT.getNumberFormats();
  var numberFormatsGAM = dataRangeGAM.getNumberFormats();
  var numberFormatsBOR = dataRangeBOR.getNumberFormats();
  var numberFormatsAMR = dataRangeAMR.getNumberFormats();
  var numberFormatsRAS = dataRangeRAS.getNumberFormats();
  var numberFormatsGOM = dataRangeGOM.getNumberFormats();
  var numberFormatsABO = dataRangeABO.getNumberFormats();
  var numberFormatsMAM = dataRangeMAM.getNumberFormats();



  //===================================================


  targetSheetDAR.getRange(1, 1, numberFormatsDAR.length, numberFormatsDAR[0].length).setNumberFormats(numberFormatsDAR);
  targetSheetAWA.getRange(1, 1, numberFormatsAWA.length, numberFormatsAWA[0].length).setNumberFormats(numberFormatsAWA);
  targetSheetANF.getRange(1, 1, numberFormatsANF.length, numberFormatsANF[0].length).setNumberFormats(numberFormatsANF);
  targetSheetFAW.getRange(1, 1, numberFormatsFAW.length, numberFormatsFAW[0].length).setNumberFormats(numberFormatsFAW);
  targetSheetKOM.getRange(1, 1, numberFormatsKOM.length, numberFormatsKOM[0].length).setNumberFormats(numberFormatsKOM);
  targetSheetRAM.getRange(1, 1, numberFormatsRAM.length, numberFormatsRAM[0].length).setNumberFormats(numberFormatsRAM);
  targetSheetFEV.getRange(1, 1, numberFormatsFEV.length, numberFormatsFEV[0].length).setNumberFormats(numberFormatsFEV);
  targetSheetOPT.getRange(1, 1, numberFormatsOPT.length, numberFormatsOPT[0].length).setNumberFormats(numberFormatsOPT);
  targetSheetGAM.getRange(1, 1, numberFormatsGAM.length, numberFormatsGAM[0].length).setNumberFormats(numberFormatsGAM);
  targetSheetBOR.getRange(1, 1, numberFormatsBOR.length, numberFormatsBOR[0].length).setNumberFormats(numberFormatsBOR);
  targetSheetAMR.getRange(1, 1, numberFormatsAMR.length, numberFormatsAMR[0].length).setNumberFormats(numberFormatsAMR);
  targetSheetRAS.getRange(1, 1, numberFormatsRAS.length, numberFormatsRAS[0].length).setNumberFormats(numberFormatsRAS);
  targetSheetGOM.getRange(1, 1, numberFormatsGOM.length, numberFormatsGOM[0].length).setNumberFormats(numberFormatsGOM);
  targetSheetABO.getRange(1, 1, numberFormatsABO.length, numberFormatsABO[0].length).setNumberFormats(numberFormatsABO);
  targetSheetMAM.getRange(1, 1, numberFormatsMAM.length, numberFormatsMAM[0].length).setNumberFormats(numberFormatsMAM);



  //===================================================

  var numRowsDAR = dataDAR.length;
  var numRowsAWA = dataAWA.length;
  var numRowsANF = dataANF.length;
  var numRowsFAW = dataFAW.length;
  var numRowsKOM = dataKOM.length;
  var numRowsRAM = dataRAM.length;
  var numRowsFEV = dataFEV.length;
  var numRowsOPT = dataOPT.length;
  var numRowsGAM = dataGAM.length;
  var numRowsBOR = dataBOR.length;
  var numRowsAMR = dataAMR.length;
  var numRowsRAS = dataRAS.length;
  var numRowsGOM = dataGOM.length;
  var numRowsABO = dataABO.length;
  var numRowsMAM = dataMAM.length;

  //===================================================

  var numColsDAR = dataDAR[0].length;
  var numColsAWA = dataAWA[0].length;
  var numColsANF = dataANF[0].length;
  var numColsFAW = dataFAW[0].length;
  var numColsKOM = dataKOM[0].length;
  var numColsRAM = dataRAM[0].length;
  var numColsFEV = dataFEV[0].length;
  var numColsOPT = dataOPT[0].length;
  var numColsGAM = dataGAM[0].length;
  var numColsBOR = dataBOR[0].length;
  var numColsAMR = dataAMR[0].length;
  var numColsRAS = dataRAS[0].length;
  var numColsGOM = dataGOM[0].length;
  var numColsABO = dataABO[0].length;
  var numColsMAM = dataMAM[0].length;

  //===================================================

  var rangeDAR = targetSheetDAR.getRange(1, 1, numRowsDAR, numColsDAR);
  var rangeAWA = targetSheetAWA.getRange(1, 1, numRowsAWA, numColsAWA);
  var rangeANF = targetSheetANF.getRange(1, 1, numRowsANF, numColsANF);
  var rangeFAW = targetSheetFAW.getRange(1, 1, numRowsFAW, numColsFAW);
  var rangeKOM = targetSheetKOM.getRange(1, 1, numRowsKOM, numColsKOM);
  var rangeRAM = targetSheetRAM.getRange(1, 1, numRowsRAM, numColsRAM);
  var rangeFEV = targetSheetFEV.getRange(1, 1, numRowsFEV, numColsFEV);
  var rangeOPT = targetSheetOPT.getRange(1, 1, numRowsOPT, numColsOPT);
  var rangeGAM = targetSheetGAM.getRange(1, 1, numRowsGAM, numColsGAM);
  var rangeBOR = targetSheetBOR.getRange(1, 1, numRowsBOR, numColsBOR);
  var rangeAMR = targetSheetAMR.getRange(1, 1, numRowsAMR, numColsAMR);
  var rangeRAS = targetSheetRAS.getRange(1, 1, numRowsRAS, numColsRAS);
  var rangeGOM = targetSheetGOM.getRange(1, 1, numRowsGOM, numColsGOM);
  var rangeABO = targetSheetABO.getRange(1, 1, numRowsABO, numColsABO);
  var rangeMAM = targetSheetMAM.getRange(1, 1, numRowsMAM, numColsMAM);



  //===================================================

  rangeDAR.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeAWA.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeANF.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeFAW.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeKOM.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeRAM.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeFEV.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeOPT.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeGAM.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeBOR.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeAMR.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeRAS.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeGOM.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeABO.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangeMAM.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);



}


