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
    var file = DriveApp.getFileById("1JXRMAcHjONC-r-rDgAtmSp7U5lPIXKcdG67KVUOiOJg");
    
    //Set these variables to which column.
    var nameCol = 4
    var empidCol = 8
    var majorCol = 10
    var phoneCol = 9
    var hireCol = 14
    var gradCol = 11
    var statusCol = 12
    var emailCol = 6
    
    var array = studentSheet.getRange(1, 1, studentSheet.getLastRow(), studentSheet.getLastColumn()).getValues()
    
    //Set these two variables to create a certain amount at a time. Maybe rows 9 to 50? Then 51 to 90? etc.
    var firstRow = 9
    var lastRow = 50
    
    //for (i = 1; i < studentSheet.getLastRow(); i += 1) {
    for (i = firstRow; i < lastRow; i += 1) {
        var name = array[i][nameCol-1];
        var empid = array[i][empidCol-1];
        var hire = array[i][hireCol-1];
        var grad = array[i][gradCol-1];
        var status = array[i][statusCol-1];
        var major = array[i][majorCol-1];
        var phone = array[i][phoneCol-1];
        var email = array[i][emailCol-1];
        
        //Only creates new files for ACTIVE students
        if (status.indexOf("Active") != -1) {
            
            //Make sure to copy and paste the ID of the folder you want every sheet to be copied to!
            var newFile = file.makeCopy(name + " Operations Leveling Requirements", DriveApp.getFolderById("0B9SeY6Mq4-wMdFNsOS0wQmhGeTA"));
            
            var sheet = SpreadsheetApp.open(newFile).getActiveSheet();
            //You may need to adjust the cell location depending on template. For ex. Videographer template has name at C8.
            sheet.getRange("C8").setValue(name);
            sheet.getRange("F8").setValue(empid);
            sheet.getRange("C10").setValue(major);
            sheet.getRange("F11").setValue(phone);
            sheet.getRange("C11").setValue(email);
            sheet.getRange("C9").setValue(hire);
            sheet.getRange("F9").setValue(grad);
            sheet.getRange("C8:C11").setHorizontalAlignment("left")
            sheet.getRange("F8:G11").setHorizontalAlignment("left")
            
        }
    }
}

