//Need function to clear everything!

function main() {

  var opsProgressSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Progress - Ops");
  var emailData = opsProgressSheet.getRange(1, 2, opsProgressSheet.getLastRow(), 1).getValues();

  opsProgress(opsProgressSheet, emailData);
}

function opsProgress(opsProgressSheet, emailData) {
  var w2wSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ops W2W Shift Data");
  var data = w2wSheet.getRange(3,4,w2wSheet.getLastRow()-2,w2wSheet.getLastColumn()-3).getValues();

  //clearProgressOps(opsProgressSheet);

  //YL1OpeningShifts(opsProgressSheet, emailData, data);
  //ZL1ClosingShifts(opsProgressSheet, emailData, data);
  //AAL1HousekeepingShifts(opsProgressSheet, emailData, data);
  //ABL1BasicAVShifts(opsProgressSheet, emailData, data);
  //ADL1BTrainShifts(opsProgressSheet, emailData, data);
  //AEL1AdvancedAVShifts(opsProgressSheet, emailData, data);
  //AFL1AdvancedAVShadowShifts(opsProgressSheet, emailData, data);
  //AGL1ShiftHours(opsProgressSheet, emailData, data);
  //AIL1RingMallShift(opsProgressSheet, emailData, data);
  //AJL2ShiftHours(opsProgressSheet, emailData, data);
  //ATL2CLTrainingShift(opsProgressSheet, emailData, data);
  //AVL2VisitorCenterHours(opsProgressSheet, emailData, data);
  //AWL3ShiftHours(opsProgressSheet, emailData, data);
  //BBL3AVTrainerShift(opsProgressSheet, emailData, data);

}

function YL1OpeningShifts(opsProgressSheet, emailData, data) {
  var OPENING_COL = "Y";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L1 - Train Ops: Morning #1" || data[i][1] == "L1 - Train Ops: Morning #2" || data[i][1] == "L1 - Train Ops: Morning #3") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(OPENING_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(OPENING_COL + getStudentRow(emailData, data[i][4])).setValue(1.0/3.0);
        }else{
          opsProgressSheet.getRange(OPENING_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(1.0/3.0)));
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

function ZL1ClosingShifts(opsProgressSheet, emailData, data) {
  var CLOSING_COL = "Z";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L1 - Train Ops: Night #1" || data[i][1] == "L1 - Train Ops: Night #2" || data[i][1] == "L1 - Train Ops: Night #3") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(CLOSING_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(CLOSING_COL + getStudentRow(emailData, data[i][4])).setValue(1.0/3.0);
        }else{
          opsProgressSheet.getRange(CLOSING_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(1.0/3.0)));
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

function AAL1HousekeepingShifts(opsProgressSheet, emailData, data) {
  var HOUSEKEEPING_COL = "AA";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L1 - Train Housekeeping - Basic" || data[i][1] == "L1 - Train Housekeeping - Restroom") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(HOUSEKEEPING_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(HOUSEKEEPING_COL + getStudentRow(emailData, data[i][4])).setValue(1.0/2.0);
        }else{
          opsProgressSheet.getRange(HOUSEKEEPING_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(1.0/2.0)));
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

//2 or 3?
function ABL1BasicAVShifts(opsProgressSheet, emailData, data) {
  var BASIC_AV_COL = "AB";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L1 - Train AV \(*Basic - 1st Round\)" || data[i][1] == "L1 - Train AV \(*Basic - 2nd Round\)") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(BASIC_AV_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(BASIC_AV_COL + getStudentRow(emailData, data[i][4])).setValue(1.0/2.0);
        }else{
          opsProgressSheet.getRange(BASIC_AV_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(1.0/2.0)));
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

//Can't find on leveling whether hours or shifts
function ADL1BTrainShifts(opsProgressSheet, emailData, data) {
  var L1B_COL = "AD";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L1 - Train AV \(*Basic - 1st Round\)" || data[i][1] == "L1 - Train AV \(*Basic - 2nd Round\)") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(L1B_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(L1B_COL + getStudentRow(emailData, data[i][4])).setValue(1.0/2.0);
        }else{
          opsProgressSheet.getRange(L1B_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(1.0/2.0)));
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

function AEL1AdvancedAVShifts(opsProgressSheet, emailData, data) {
  var ADVANCED_AV_COL = "AE";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L1 - Train AV \(Advanced - 1st Round\)" || data[i][1] == "L1 - Train AV \(Advanced - 2nd Round\)" || data[i][1] == "L1 - Train AV \(Advanced - 3rd Round\)") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(ADVANCED_AV_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(ADVANCED_AV_COL + getStudentRow(emailData, data[i][4])).setValue(1.0/3.0);
        }else{
          opsProgressSheet.getRange(ADVANCED_AV_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(1.0/3.0)));
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

function AFL1AdvancedAVShadowShifts(opsProgressSheet, emailData, data) {
  var ADVANCED_AV_SHADOW_COL = "AF";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L1 - Train AV \(Advanced - Shadow Shift\)") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(ADVANCED_AV_SHADOW_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(ADVANCED_AV_SHADOW_COL + getStudentRow(emailData, data[i][4])).setValue(1.0/1.0);
        }else{
          opsProgressSheet.getRange(ADVANCED_AV_SHADOW_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(1.0/1.0)));
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

function AGL1ShiftHours(opsProgressSheet, emailData, data) {
  var L1_SHIFT_COL = "AG";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L1.a - Ops Crew" || data[i][1] == "L1.b - Ops Crew") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(L1_SHIFT_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(L1_SHIFT_COL + getStudentRow(emailData, data[i][4])).setValue(data[i][3]/150.0);
        }else{
          opsProgressSheet.getRange(L1_SHIFT_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(data[i][3]))/150.0);
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

//Shifts or hours
function AIL1RingMallShift(opsProgressSheet, emailData, data) {
  var RING_MALL_COL = "AI";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L1 - Train Ring Mall") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(RING_MALL_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(RING_MALL_COL + getStudentRow(emailData, data[i][4])).setValue(data[i][3]/150.0);
        }else{
          opsProgressSheet.getRange(RING_MALL_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(data[i][3]))/150.0);
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

function AJL2ShiftHours(opsProgressSheet, emailData, data) {
  var AV_TECH_COL = "AJ";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L2 - AV Tech" || data[i][1] == "L2 - AV Tech CCA" || data[i][1] == "L2 - Building Lead" || data[i][1] == "L2 - Crew Lead Training" || data[i][1] == "L2 - Event Lead" || data[i][1] == "L2 - Office Assistant" || data[i][1] == "L2 - Ops Trainer" || data[i][1] == "L2 - Ring Mall Lead" || data[i][1] == "L2 - Visitor Center") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(AV_TECH_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(AV_TECH_COL + getStudentRow(emailData, data[i][4])).setValue(data[i][3]/100.0);
        }else{
          opsProgressSheet.getRange(AV_TECH_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(data[i][3]))/100.0);
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

function ATL2CLTrainingShift(opsProgressSheet, emailData, data) {
  var CL_TRAINING_COL = "AT";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L2 - Crew Lead Training") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(CL_TRAINING_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(CL_TRAINING_COL + getStudentRow(emailData, data[i][4])).setValue(1.0/3.0);
        }else{
          opsProgressSheet.getRange(CL_TRAINING_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(1.0/3.0)));
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

function AVL2VisitorCenterHours(opsProgressSheet, emailData, data) {
  var VISITOR_CENTER_COL = "AV";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L2 - Visitor Center") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(VISITOR_CENTER_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(VISITOR_CENTER_COL + getStudentRow(emailData, data[i][4])).setValue(data[i][3]/2.0);
        }else{
          opsProgressSheet.getRange(VISITOR_CENTER_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(data[i][3]))/2.0);
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

//"L3 - AV Trainer (Basic/Adv.)"?
function AWL3ShiftHours(opsProgressSheet, emailData, data) {
  var L3_SHIFT_COL = "AW";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L3 - Ops Assistant \"Joe\'s Pros\"" || data[i][1] == "L3 - Ops Crew Leader" || data[i][1] == "L3 - Student Lead Training") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(L3_SHIFT_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(L3_SHIFT_COL + getStudentRow(emailData, data[i][4])).setValue(data[i][3]/100.0);
        }else{
          opsProgressSheet.getRange(L3_SHIFT_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(data[i][3]))/100.0);
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}

function BBL3AVTrainerShift(opsProgressSheet, emailData, data) {
  var AV_TRAINING_COL = "BB";

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == "") {
      break;
    }
    if (data[i][1] == "L3 - AV Trainer (Basic/Adv.)") {
      if (data[i][4] != "" && getStudentRow(emailData, data[i][4]) != -1){
        var currentValue = opsProgressSheet.getRange(CL_TRAINING_COL + getStudentRow(emailData, data[i][4])).getValue();
        if (currentValue == "") {
          opsProgressSheet.getRange(CL_TRAINING_COL + getStudentRow(emailData, data[i][4])).setValue(1.0/2.0);
        }else{
          opsProgressSheet.getRange(CL_TRAINING_COL + getStudentRow(emailData, data[i][4])).setValue(parseFloat(currentValue) + (parseFloat(1.0/2.0)));
        }
      }
    }else{
      Logger.log("Error " + data[i][4]);
    }
  }
}


//Function to return row of student based on their email
function getStudentRow(emailData, email) {
  if (email == "") {
    return -1;
  }
  for (var j = 0; j < emailData.length; j++){
    if (emailData[j] == email){
      return((j + 1));
    }
  }
  return -1;
}

function clearProgressOps(opsProgressSheet) {

}
