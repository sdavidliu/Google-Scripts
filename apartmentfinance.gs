//Code for Stanford 474 Spreadsheet. Go to Tools -> Script Editor, and copy and paste this into Code.gs and press the Run button (traingle shaped).
function updateSpreadsheet() {
  //Get the active sheet, make sure it's the right one!
  var sheet = SpreadsheetApp.getActiveSheet();

  //Just to calculate which row is the first totals
  var firstTotalRow = getFirstTotalRow(sheet)

  //Reset all the totals to $0.00
  setTotalsToZero(sheet, firstTotalRow)

  var startRow = 28;  // First row of data to process
  var dataRange = sheet.getRange(startRow, 2, sheet.getLastRow()-startRow, sheet.getLastColumn()-1)
  var data = dataRange.getValues();

  //Loop through every row from row 28 until the break statement
  for (i in data) {
    var row = data[i];

    //Get each value from the row
    var name = row[3];
    var david = row[4];
    var dillon = row[5];
    var michael = row[6];
    var nick = row[7];
    var simeon = row[8];
    var perPerson = row[9];
    if (name == ""){
      //If reach end of items list, break
      break
    }

    //Figure out who paid for item and put in calculations in Totals table
    if (name == "David") {
      if (dillon == "✔") {
        sheet.getRange(Number(firstTotalRow) + 0, 5).setValue(Number(sheet.getRange(Number(firstTotalRow) + 0, 5).getValue()) + Number(perPerson))
      }
      if (michael == "✔") {
        sheet.getRange(Number(firstTotalRow) + 0, 6).setValue(Number(sheet.getRange(Number(firstTotalRow) + 0, 6).getValue()) + Number(perPerson))
      }
      if (nick == "✔") {
        sheet.getRange(Number(firstTotalRow) + 0, 7).setValue(Number(sheet.getRange(Number(firstTotalRow) + 0, 7).getValue()) + Number(perPerson))
      }
      if (simeon == "✔") {
        sheet.getRange(Number(firstTotalRow) + 0, 8).setValue(Number(sheet.getRange(Number(firstTotalRow) + 0, 8).getValue()) + Number(perPerson))
      }
    }else if (name == "Dillon") {
      if (david == "✔") {
        sheet.getRange(Number(firstTotalRow) + 1, 4).setValue(Number(sheet.getRange(Number(firstTotalRow) + 1, 4).getValue()) + Number(perPerson))
      }
      if (michael == "✔") {
        sheet.getRange(Number(firstTotalRow) + 1, 6).setValue(Number(sheet.getRange(Number(firstTotalRow) + 1, 6).getValue()) + Number(perPerson))
      }
      if (nick == "✔") {
        sheet.getRange(Number(firstTotalRow) + 1, 7).setValue(Number(sheet.getRange(Number(firstTotalRow) + 1, 7).getValue()) + Number(perPerson))
      }
      if (simeon == "✔") {
        sheet.getRange(Number(firstTotalRow) + 1, 8).setValue(Number(sheet.getRange(Number(firstTotalRow) + 1, 8).getValue()) + Number(perPerson))
      }
    }else if (name == "Michael") {
      if (david == "✔") {
        sheet.getRange(Number(firstTotalRow) + 2, 4).setValue(Number(sheet.getRange(Number(firstTotalRow) + 2, 4).getValue()) + Number(perPerson))
      }
      if (dillon == "✔") {
        sheet.getRange(Number(firstTotalRow) + 2, 5).setValue(Number(sheet.getRange(Number(firstTotalRow) + 2, 5).getValue()) + Number(perPerson))
      }
      if (nick == "✔") {
        sheet.getRange(Number(firstTotalRow) + 2, 7).setValue(Number(sheet.getRange(Number(firstTotalRow) + 2, 7).getValue()) + Number(perPerson))
      }
      if (simeon == "✔") {
        sheet.getRange(Number(firstTotalRow) + 2, 8).setValue(Number(sheet.getRange(Number(firstTotalRow) + 2, 8).getValue()) + Number(perPerson))
      }
    }else if (name == "Nick") {
      if (david == "✔") {
        sheet.getRange(Number(firstTotalRow) + 3, 4).setValue(Number(sheet.getRange(Number(firstTotalRow) + 3, 4).getValue()) + Number(perPerson))
      }
      if (dillon == "✔") {
        sheet.getRange(Number(firstTotalRow) + 3, 5).setValue(Number(sheet.getRange(Number(firstTotalRow) + 3, 5).getValue()) + Number(perPerson))
      }
      if (michael == "✔") {
        sheet.getRange(Number(firstTotalRow) + 3, 6).setValue(Number(sheet.getRange(Number(firstTotalRow) + 3, 6).getValue()) + Number(perPerson))
      }
      if (simeon == "✔") {
        sheet.getRange(Number(firstTotalRow) + 3, 8).setValue(Number(sheet.getRange(Number(firstTotalRow) + 3, 8).getValue()) + Number(perPerson))
      }
    }else if (name == "Simeon") {
      if (david == "✔") {
        sheet.getRange(Number(firstTotalRow) + 4, 4).setValue(Number(sheet.getRange(Number(firstTotalRow) + 4, 4).getValue()) + Number(perPerson))
      }
      if (dillon == "✔") {
        sheet.getRange(Number(firstTotalRow) + 4, 5).setValue(Number(sheet.getRange(Number(firstTotalRow) + 4, 5).getValue()) + Number(perPerson))
      }
      if (michael == "✔") {
        sheet.getRange(Number(firstTotalRow) + 4, 6).setValue(Number(sheet.getRange(Number(firstTotalRow) + 4, 6).getValue()) + Number(perPerson))
      }
      if (nick == "✔") {
        sheet.getRange(Number(firstTotalRow) + 4, 7).setValue(Number(sheet.getRange(Number(firstTotalRow) + 4, 7).getValue()) + Number(perPerson))
      }
    }
  }

  //Calculate summary at the end
  calculateSummary(sheet, firstTotalRow)
}

function getFirstTotalRow(sheet){
  var dataRange = sheet.getRange(30, 2, sheet.getLastRow()-30, 1)
  var data = dataRange.getValues();
  for (i in data) {
    if (data[i][0] == "Totals") {
      return Number(i) + 33
    }
  }
}

function setTotalsToZero(sheet, firstTotalRow) {
  var dataRange = sheet.getRange(firstTotalRow, 4, 5, 5);
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
    for (j in row) {
      if (row[j] != "✖") {
        sheet.getRange(Number(firstTotalRow) + Number(i), 4 + Number(j)).setValue(0);
      }
    }
  }
}

function calculateSummary(sheet, firstTotalRow) {
  var summaryRow = 0;
  var dataRange = sheet.getRange(Number(firstTotalRow) + 5, 2, sheet.getLastRow() - (Number(firstTotalRow) + 5), 1);
  var data = dataRange.getValues();
  for (i in data) {
    if (data[i][0] == "Summary") {
      summaryRow = Number(i) + Number(firstTotalRow) + 5 + 1;
      break
    }
  }
  if (summaryRow == 0){
    Logger.log("Can't find summary row");
    return
  }
  Logger.log("First total row: " + firstTotalRow);
  Logger.log("Summary row: " + summaryRow);
  var davidDillon = summaryString("David", sheet.getRange("E" + (Number(firstTotalRow) + 0)).getValue(), "Dillon", sheet.getRange("D" + (Number(firstTotalRow) + 1)).getValue());
  var davidMichael = summaryString("David", sheet.getRange("F" + (Number(firstTotalRow) + 0)).getValue(), "Michael", sheet.getRange("D" + (Number(firstTotalRow) + 2)).getValue());
  var davidNick = summaryString("David", sheet.getRange("G" + (Number(firstTotalRow) + 0)).getValue(), "Nick", sheet.getRange("D" + (Number(firstTotalRow) + 3)).getValue());
  var davidSimeon = summaryString("David", sheet.getRange("H" + (Number(firstTotalRow) + 0)).getValue(), "Simeon", sheet.getRange("D" + (Number(firstTotalRow) + 4)).getValue());
  var dillonMichael = summaryString("Dillon", sheet.getRange("F" + (Number(firstTotalRow) + 1)).getValue(), "Michael", sheet.getRange("E" + (Number(firstTotalRow) + 2)).getValue());
  var dillonNick = summaryString("Dillon", sheet.getRange("G" + (Number(firstTotalRow) + 1)).getValue(), "Nick", sheet.getRange("E" + (Number(firstTotalRow) + 3)).getValue());
  var dillonSimeon = summaryString("Dillon", sheet.getRange("H" + (Number(firstTotalRow) + 1)).getValue(), "Simeon", sheet.getRange("E" + (Number(firstTotalRow) + 4)).getValue());
  var michaelNick = summaryString("Michael", sheet.getRange("G" + (Number(firstTotalRow) + 2)).getValue(), "Nick", sheet.getRange("F" + (Number(firstTotalRow) + 3)).getValue());
  var michaelSimeon = summaryString("Michael", sheet.getRange("H" + (Number(firstTotalRow) + 2)).getValue(), "Simeon", sheet.getRange("F" + (Number(firstTotalRow) + 4)).getValue());
  var nickSimeon = summaryString("Nick", sheet.getRange("H" + (Number(firstTotalRow) + 3)).getValue(), "Simeon", sheet.getRange("G" + (Number(firstTotalRow) + 4)).getValue());
  var summary = davidDillon + "\n" + "\n" + davidMichael + "\n" + "\n" + davidNick + "\n" + "\n" + davidSimeon + "\n" + "\n" + dillonMichael + "\n" + "\n" + dillonNick + "\n" + "\n" + dillonSimeon + "\n" + "\n" + michaelNick + "\n" + "\n" + michaelSimeon + "\n" + "\n" + nickSimeon
  sheet.getRange("B" + summaryRow).setValue(summary)
  var d = new Date()
  sheet.getRange("B" + (Number(summaryRow) + 2)).setValue("Last updated: " + (d.getMonth()+1) + "/" + d.getDate() + "/" + d.getYear())
}

function summaryString(name1, value1, name2, value2) {
  if (value1 == value2) {
    return "Nothing between " + name1 + " and " + name2;
  }
  if (value1 > value2) {
    return name2 + " pays " + name1 + " $" + (Number(value1) - Number(value2)).toFixed(2);
  }
  return name1 + " pays " + name2 + " $" + (Number(value2) - Number(value1)).toFixed(2);
}
