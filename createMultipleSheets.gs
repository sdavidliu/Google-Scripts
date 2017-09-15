/*
Summary:
Run "createStudentFiles()" to create new student files based on a template.
Make sure to set the "file" variable to the right template.
Make sure to set column variables to which column each information is at.
This code loops through every student and creates a file for every row.

If error:
Go to View -> Logs to see error logs.
Contact David Liu (Created September 2017)
*/

function createStudentFiles() {
  //This is your current sheet with list of students.
  var studentSheet = SpreadsheetApp.getActiveSheet();

  //Make sure to set the ID to the template.
  var file = DriveApp.getFileById("1PDqmG1xSk-uqMFJSBJA495TxF6Tv0HgzZ7fQ7FYOc3o");

  //Set these variables to which column. Ex. The name column is in column 1.
  var nameCol = 1
  var departmentCol = 2
  var idCol = 3
  var hireCol = 4
  var gradCol = 5
  var statusCol = 6
  var checkCol1 = 8
  var checkCol2 = 9
  var checkCol3 = 10
  var checkCol4 = 11

  for (i = 1; i < studentSheet.getLastRow(); i += 1) {
    var name = studentSheet.getRange(i + 1, nameCol).getValue();
    var department = studentSheet.getRange(i + 1, departmentCol).getValue();
    var id = studentSheet.getRange(i + 1, idCol).getValue();
    var hire = studentSheet.getRange(i + 1, hireCol).getValue();
    var grad = studentSheet.getRange(i + 1, gradCol).getValue();
    var status = studentSheet.getRange(i + 1, statusCol).getValue();
    var check1 = studentSheet.getRange(i + 1, checkCol1).getValue();
    var check2 = studentSheet.getRange(i + 1, checkCol2).getValue();
    var check3 = studentSheet.getRange(i + 1, checkCol3).getValue();
    var check4 = studentSheet.getRange(i + 1, checkCol4).getValue();

    //Set which template to copy depending on department, make sure to get department name exactly the same!
    if (department == "AntMedia Videographer") {
      file = DriveApp.getFileById("1PDqmG1xSk-uqMFJSBJA495TxF6Tv0HgzZ7fQ7FYOc3o");
    }else if (department == "Marketing") {
      //file = DriveApp.getFileById("PUT ID HERE");
    }else if (department == "Marketing/Ops"){
      //file = DriveApp.getFileById("PUT ID HERE");
    }else if (department == "AntMedia Videographer & Photographer"){
      //file = DriveApp.getFileById("PUT ID HERE");
    }else if (department == "AntMedia Photographer"){
      //file = DriveApp.getFileById("PUT ID HERE");
    }else if (department == "AntMedia/Ops"){
      //file = DriveApp.getFileById("PUT ID HERE");
    }else if (department == "Operations"){
      //file = DriveApp.getFileById("PUT ID HERE");
    }else if (department == "Reservation Specialist"){
      //file = DriveApp.getFileById("PUT ID HERE");
    }else if (department == "Marketing Assisstant"){
      //file = DriveApp.getFileById("PUT ID HERE");
    }else{
      //file = DriveApp.getFileById("PUT ID HERE");
    }

    //Only creates new files for ACTIVE students
    if (status.indexOf("Active") != -1) {

      //Make sure to copy and paste the ID of the folder you want every sheet to be copied to!
      var newFile = file.makeCopy(name + " Videographer Leveling Requirements", DriveApp.getFolderById("0B6yPbMopq0nJb0hKWHlSeUVrTUE"));

      var sheet = SpreadsheetApp.open(newFile).getActiveSheet();
      //You may need to adjust the cell location depending on template. For ex. Videographer template has name at C8.
      sheet.getRange("C8").setValue(name.toUpperCase());
      sheet.getRange("F8").setValue(id);
      sheet.getRange("F8").setHorizontalAlignment("left")
      try {
        sheet.getRange("C9").setValue(Utilities.formatDate(hire, "PDT", "MM/dd/yyyy"))
        sheet.getRange("C9").setHorizontalAlignment("left")
      } catch(e) {
        Logger.log(e)
      }
      try {
        sheet.getRange("F9").setValue(Utilities.formatDate(grad, "PDT", "MM/yyyy"))
        sheet.getRange("F9").setHorizontalAlignment("left")
      } catch(e) {
        Logger.log(e)
      }
      try {
        sheet.getRange("C15").setValue(Utilities.formatDate(check1, "PDT", "MM/dd/yyyy") + ", " + Utilities.formatDate(check2, "PDT", "MM/dd/yyyy") + ", " + Utilities.formatDate(check3, "PDT", "MM/dd/yyyy") + ", " + Utilities.formatDate(check4, "PDT", "MM/dd/yyyy"))
        sheet.getRange("C15").setHorizontalAlignment("left")
      } catch(e) {
        Logger.log(e)
      }
    }

    //Everything that is not found is set to 'xxxx'
    setEverythingElseX(sheet)

    //Uncomment the break below to just test it on one student
    //break
  }
}

function setEverythingElseX(sheet) {
  if (sheet.getRange("C8").getValue() == "") {
    sheet.getRange("C8").setValue("xxxx")
  }
  if (sheet.getRange("F8").getValue() == "") {
    sheet.getRange("F8").setValue("xxxx")
  }
  if (sheet.getRange("C9").getValue() == "") {
    sheet.getRange("C9").setValue("xxxx")
  }
  if (sheet.getRange("F9").getValue() == "") {
    sheet.getRange("F9").setValue("xxxx")
  }
  if (sheet.getRange("C10").getValue() == "") {
    sheet.getRange("C10").setValue("xxxx")
  }
  if (sheet.getRange("F10").getValue() == "") {
    sheet.getRange("F10").setValue("xxxx")
  }
  if (sheet.getRange("C11").getValue() == "") {
    sheet.getRange("C11").setValue("xxxx")
  }
  if (sheet.getRange("F11").getValue() == "") {
    sheet.getRange("F11").setValue("xxxx")
  }
}
