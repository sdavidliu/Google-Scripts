//Global sheet variable
var sheet = SpreadsheetApp.getActiveSheet()

function updateSpreadsheet() {

  //Below is the folder ID of all the files you want to update
  var spreadsheetIDs = getFileIDs("0B6yPbMopq0nJcDQ4Z00xQkRmQ2c")
  for (i = 0; i < spreadsheetIDs.length; i += 1) {
    sheet = SpreadsheetApp.openById(spreadsheetIDs[i]).getSheetByName('Leveling');
    Logger.log(sheet.getRange("C8").getValue());

    //Add code here
    //setRowBold(8, true)
  }
}

/*
Functions:
- addRow(row, array)
- deleteRow(row)
- setTextColor(cell, color)
- setRowTextColor(row, color)
- setBackgroundColor(cell, color)
- setRowBackgroundColor(row, color)
- setBackgroundRGB(cell, red, green, blue)
- setRowBackgroundRGB(row, red, green, blue)
- addCol(col)
- deleteCol(col)
- setBorders(cell, top, left, bottom, right, vertical, horizontal)
- setCellFont(cell, font)
- setRowFont(row, font)
- setCellFontSize(cell, size)
- setRowFontSize(row, size)
- setCellFormula(cell, formula)
- setRowFormula(row, formula)
- setCellHorizontalAlignment(cell, alignment)
- setRowHorizontalAlignment(row, alignment)
- setText(cell, text)
- setCellVerticalAlignment(cell, alignment)
- setRowVerticalAlignment(row, alignment)
- setCellWrap(cell, wrap)
- setRowWrap(row, wrap)
- setCellBold(cell, bold)
- setRowBold(row, bold)
- setCellItalic(cell, italic)
- setRowItalic(row, italic)
- setCellUnderline(cell, underline)
- setRowUnderline(row, underline)
- setCellStrikethrough(cell, strikethrough)
- setRowStrikethrough(row, strikethrough)
- copyCell(cellToCopy, cellToPaste)
- copyRow(rowToCopy, rowToPaste)

Documentation:
addRow(row, array)
  row - row number as an integer
  array - array of text you want to put in the row
  ex. addRow(1, ["","Name:","Peter","Employee #","1234",""])

deleteRow(row)
  row - row number as an integer
  ex. deleteRow(1)

setTextColor(cell, color)
  cell - cell in letter number notation
  color - color as a string
  ex. setTextColor("A1", "BLUE")

setRowTextColor(row, color)
  row - row number as an integer
  color - color as a string
  ex. setRowTextColor(1, "BLUE")

setBackgroundColor(cell, color)
  cell - cell in letter number notation
  color - color as a string
  ex. setBackgroundColor("A1", "BLUE")

setRowBackgroundColor(row, color)
  row - row number as an integer
  color - color as a string
  ex. setRowBackgroundColor(1, "BLUE")

setBackgroundRGB(cell, red, green, blue)
  cell - cell in letter number notation
  red, green, blue - integer between 0 and 255 inclusive
  ex. setBackgroundRGB("A1", 255, 255, 255)

setRowBackgroundRGB(row, red, green, blue)
  row - row number as an integer
  red, green, blue - integer between 0 and 255 inclusive
  ex. setRowBackgroundRGB(1, 255, 255, 255)

addCol(col)
  col - column number as an integer starting at 1
  ex. addCol(1)

deleteCol(col)
  col - column number as an integer starting at 1
  ex. deleteCol(1)

setBorders(cell, top, left, bottom, right, vertical, horizontal)
  cell - cell in letter number notation
  top, left, bottom, right, vertical, horizontal - true for border, false for no border, and null for no change
  ex. setBorders("A1", true, true, true, true, null, null)

setCellFont(cell, font)
  cell - cell in letter number notation
  font - font name as a string
  ex. setCellFont("A1", "Arial")

setRowFont(row, font)
  row - row number as an integer
  font - font name as a string
  ex. setRowFont(1, "Arial")

setCellFontSize(cell, size)
  cell - cell in letter number notation
  size - font size as an integer
  ex. setCellFontSize("A1", 12)

setRowFontSize(row, size)
  row - row number as an integer
  size - fotn size as an integer
  ex. setRowFontSize(1, 12)

setCellFormula(cell, formula)
  cell - cell in letter number notation
  formula - formula as a string
  ex. setCellFormula("A1", "=ADD(2,3)")

setRowFormula(row, formula)
  row - row number as an integer
  formula - formula as a string
  ex. setRowFormula(1, "=ADD(2,3)")

setCellHorizontalAlignment(cell, alignment)
  cell - cell in letter number notation
  alignment - "right", "left", "center"
  ex. setCellHorizontalAlignment("A1", "center")

setRowHorizontalAlignment(row, alignment)
  row - row number as an integer
  alignment - "right", "left", "center"
  ex. setRowHorizontalAlignment(1, "center")

setText(cell, text)
  cell - cell in letter number notation
  text - text as a string
  ex. setText("A1", "hello world!")

setCellVerticalAlignment(cell, alignment)
  cell - cell in letter number notation
  alignment - "top", "middle", "bottom"
  ex. setCellVerticalAlignment("A1", "middle")

setRowVerticalAlignment(row, alignment)
  row - row number as an integer
  alignment - "top", "middle", "bottom"
  ex. setRowVerticalAlignment(1, "middle")

setCellWrap(cell, wrap)
  cell - cell in letter number notation
  wrap - true for wrap, false for no wrap
  ex. setCellWrap("A1", true)

setRowWrap(row, wrap)
  row - row number as an integer
  wrap - true for wrap, false for no wrap
  ex. setRowWrap(1, true)

setCellBold(scell, bold)
  cell - cell in letter number notation
  bold - true for bold, false for not bold
  ex. setCellBold("A1", true)

setRowBold(row, bold)
  row - row number as an integer
  bold - true for bold, false for not bold
  ex. setRowBold(1, true)

setCellItalic(cell, italic)
  cell - cell in letter number notation
  italic - true for italic, false for not italic
  ex. setCellItalic("A1", true)

setRowItalic(row, italic)
  row - row number as an integer
  italic - true for italic, false for not italic
  ex. setRowItalic(1, true)

setCellUnderline(cell, underline)
  cell - cell in letter number notation
  underline - true for underline, false for not underline
  ex. setCellUnderline("A1", true)

setRowUnderline(row, underline)
  row - row number as an integer
  underline - true for underline, false for not underline
  ex. setRowUnderline(1, true)

setCellStrikethrough(cell, strikethrough)
  cell - cell in letter number notation
  strikethrough - true for strikethrough, false for not strikethrough
  ex. setCellStrikethrough("A1", true)

setRowStrikethrough(row, strikethrough)
  row - row number as an integer
  strikethrough - true for strikethrough, false for not strikethrough
  ex. setRowStrikethrough(1, true)

copyCell(cellToCopy, cellToPaste)
  cellToCopy - cell in letter number notation
  cellToPast - cell in letter number notation
  ex. copyCell("C8", "C9")

copyRow(rowToCopy, rowToPaste)
  rowToCopy - row number as an integer
  rowToPaste - row number as an integer
  ex. copyRow(8,10)
*/

function addRow(row, array) {
  sheet.insertRowBefore(row);
  for (c = 0; c < array.length; c += 1) {
    sheet.getRange(row, c+1).setValue(array[c]);
  }
}

function deleteRow(row) {
  sheet.deleteRow(row);
}

function setTextColor(cell, color) {
  sheet.getRange(cell).setFontColor(color);
}

function setRowTextColor(row, color) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    sheet.getRange(row, c+1).setFontColor(color);
  }
}

function setBackgroundColor(cell, color) {
  sheet.getRange(cell).setBackground(color);
}

function setRowBackgroundColor(row, color) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    sheet.getRange(row, c+1).setBackGround(color);
  }
}

function setBackgroundRGB(cell, red, green, blue) {
  sheet.getRange(cell).setBackgroundRGB(red, green, blue);
}

function setRowBackgroundRGB(row, red, green, blue) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    sheet.getRange(row, c+1).setBackgroundRGB(red, green, blue);
  }
}

function addCol(col) {
  sheet.insertColumnBefore(col);
}

function deleteCol(col) {
  sheet.deleteColumn(col);
}

function setBorders(cell, top, left, bottom, right, vertical, horizontal) {
  sheet.getRange(cell).setBorder(top, left, bottom, right, vertical, horizontal);
}

function setCellFont(cell, font) {
  sheet.getRange(cell).setFontFamily(font);
}

function setRowFont(row, font) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    sheet.getRange(row, c+1).setFontFamily(font);
  }
}

function setCellFontSize(cell, size) {
  sheet.getRange(cell).setFontSize(size);
}

function setRowFontSize(row, size) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    sheet.getRange(row, c+1).setFontSize(size);
  }
}

function setCellFormula(cell, formula) {
  sheet.getRange(cell).setFormula(formula);
}

function setRowFormula(row, formula) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    sheet.getRange(row, c+1).setFormula(formula);
  }
}

function setCellHorizontalAlignment(cell, alignment) {
  sheet.getRange(cell).setHorizontalAlignment(alignment);
}

function setRowHorizontalAlignment(row, alignment) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    sheet.getRange(row, c+1).setHorizontalAlignment(alignment);
  }
}

function setText(cell, text) {
  sheet.getRange(cell).setValue(text);
}

function setCellVerticalAlignment(cell, alignment) {
  sheet.getRange(cell).setVerticalAlignment(alignment);
}

function setRowVerticalAlignment(row, alignment) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    sheet.getRange(row, c+1).setVerticalAlignment(alignment);
  }
}

function setCellWrap(cell, wrap) {
  sheet.getRange(cell).setWrap(wrap);
}

function setRowWrap(row, wrap) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    sheet.getRange(row, c+1).setWrap(wrap);
  }
}

function setCellBold(cell, bold) {
  if (bold == true){
    sheet.getRange(cell).setFontWeight("bold");
  }else{
    sheet.getRange(cell).setFontWeight("normal");
  }
}

function setRowBold(row, bold) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    if (bold == true){
      sheet.getRange(row, c+1).setFontWeight("bold");
    }else{
      sheet.getRange(row, c+1).setFontWeight("normal");
    }
  }
}

function setCellItalic(cell, italic) {
  if (italic == true){
    sheet.getRange(cell).setFontStyle("italic");
  }else{
    sheet.getRange(cell).setFontStyle("normal");
  }
}

function setRowItalic(row, italic) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    if (italic == true){
      sheet.getRange(row, c+1).setFontStyle("italic");
    }else{
      sheet.getRange(row, c+1).setFontStyle("normal");
    }
  }
}

function setCellUnderline(cell, underline) {
  if (underline == true){
    sheet.getRange(cell).setFontLine("underline");
  }else{
    sheet.getRange(cell).setFontLine("normal");
  }
}

function setRowUnderline(row, underline) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    if (underline == true){
      sheet.getRange(row, c+1).setFontStyle("underline");
    }else{
      sheet.getRange(row, c+1).setFontStyle("normal");
    }
  }
}

function setCellStrikethrough(cell, strikethrough) {
  if (strikethrough == true){
    sheet.getRange(cell).setFontLine("line-through");
  }else{
    sheet.getRange(cell).setFontLine("normal");
  }
}

function setRowStrikethrough(row, strikethrough) {
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    if (strikethrough == true){
      sheet.getRange(row, c+1).setFontStyle("line-through");
    }else{
      sheet.getRange(row, c+1).setFontStyle("normal");
    }
  }
}

function copyCell(cellToCopy, cellToPaste) {
  sheet.getRange(cellToPaste).setValue(sheet.getRange(cellToCopy).getValue());
}

function copyRow(rowToCopy, rowToPaste) {
  var array = [];
  for (c = 0; c < sheet.getLastColumn(); c += 1) {
    array.push(sheet.getRange(rowToCopy, c+1).getValue());
  }
  addRow(rowToPaste, array);
}

function getFileIDs(id) {
  var folder = DriveApp.getFolderById(id);
  var contents = folder.getFiles();

  var cnt = 0;
  var file;

  var answer = []

  while (contents.hasNext()) {
    var file = contents.next();
    cnt++;

    answer.push(file.getId())

  };
  return answer
}
