function myFunction() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");

    for (var i = 1; i <= sheet.getLastRow(); i++) {

        var range = sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (range[24] == "") {
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").copyTo(SpreadsheetApp.getActiveSpreadsheet());
            SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of Template"));
            SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet(range[0]);
            var finalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(range[0]);
            finalSheet.getRange("B1").setValue("Extension: " + range[0]);
            finalSheet.getRange("B3").setValue(range[4]);
            finalSheet.getRange("B4").setValue(range[5]);
            finalSheet.getRange("B5").setValue(range[6]);
            finalSheet.getRange("B6").setValue(range[7]);
            finalSheet.getRange("B7").setValue(range[8]);
            finalSheet.getRange("B8").setValue(range[9]);
            finalSheet.getRange("B9").setValue(range[10]);
            finalSheet.getRange("B10").setValue(range[11]);
            finalSheet.getRange("B11").setValue(range[12]);
            finalSheet.getRange("B12").setValue(range[13]);
            finalSheet.getRange("D3").setValue(range[14]);
            finalSheet.getRange("D4").setValue(range[15]);
            finalSheet.getRange("D5").setValue(range[16]);
            finalSheet.getRange("D6").setValue(range[17]);
            finalSheet.getRange("D7").setValue(range[18]);
            finalSheet.getRange("D8").setValue(range[19]);
            finalSheet.getRange("D9").setValue(range[20]);
            finalSheet.getRange("D10").setValue(range[21]);
            finalSheet.getRange("D11").setValue(range[22]);
            finalSheet.getRange("D12").setValue(range[23]);

            finalSheet.getRange("A2:E2").setBackgroundRGB(169, 193, 240);
            finalSheet.getRange("A4:E4").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A6:E6").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A8:E8").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A10:E10").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A12:E12").setBackgroundRGB(203, 218, 245);
        } else {
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template2").copyTo(SpreadsheetApp.getActiveSpreadsheet());
            SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of Template2"));
            SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet(range[0]);
            var finalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(range[0]);
            finalSheet.getRange("B1").setValue("Extension: " + range[0]);

            finalSheet.getRange("B2").setValue(range[24]);
            finalSheet.getRange("B3").setValue(range[25]);
            finalSheet.getRange("B4").setValue(range[26]);
            finalSheet.getRange("B5").setValue(range[27]);
            finalSheet.getRange("B6").setValue(range[28]);
            finalSheet.getRange("B7").setValue(range[29]);
            finalSheet.getRange("B8").setValue(range[30]);
            finalSheet.getRange("B9").setValue(range[31]);
            finalSheet.getRange("B10").setValue(range[32]);
            finalSheet.getRange("B11").setValue(range[33]);
            finalSheet.getRange("B12").setValue(range[34]);
            finalSheet.getRange("B13").setValue(range[35]);
            finalSheet.getRange("B14").setValue(range[36]);
            finalSheet.getRange("B15").setValue(range[37]);
            finalSheet.getRange("B16").setValue(range[38]);
            finalSheet.getRange("B17").setValue(range[39]);
            finalSheet.getRange("B18").setValue(range[40]);

            finalSheet.getRange("F3").setValue(range[4]);
            finalSheet.getRange("F4").setValue(range[5]);
            finalSheet.getRange("F5").setValue(range[6]);
            finalSheet.getRange("F6").setValue(range[7]);
            finalSheet.getRange("F7").setValue(range[8]);
            finalSheet.getRange("F8").setValue(range[9]);
            finalSheet.getRange("F9").setValue(range[10]);
            finalSheet.getRange("F10").setValue(range[11]);
            finalSheet.getRange("F11").setValue(range[12]);
            finalSheet.getRange("F12").setValue(range[13]);
            finalSheet.getRange("H3").setValue(range[14]);
            finalSheet.getRange("H4").setValue(range[15]);
            finalSheet.getRange("H5").setValue(range[16]);
            finalSheet.getRange("H6").setValue(range[17]);
            finalSheet.getRange("H7").setValue(range[18]);
            finalSheet.getRange("H8").setValue(range[19]);
            finalSheet.getRange("H9").setValue(range[20]);
            finalSheet.getRange("H10").setValue(range[21]);
            finalSheet.getRange("H11").setValue(range[22]);
            finalSheet.getRange("H12").setValue(range[23]);

            finalSheet.getRange("A2:C2").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A4:C4").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A6:C6").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A8:C8").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A10:C10").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A12:C12").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A14:C14").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A16:C16").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("A18:C18").setBackgroundRGB(203, 218, 245);

            finalSheet.getRange("E2:I2").setBackgroundRGB(169, 193, 240);
            finalSheet.getRange("E4:I4").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("E6:I6").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("E8:I8").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("E10:I10").setBackgroundRGB(203, 218, 245);
            finalSheet.getRange("E12:I12").setBackgroundRGB(203, 218, 245);
        }
    }

}
